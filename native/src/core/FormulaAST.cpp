#include "FormulaAST.h"
#include <QRegularExpression>

// ============================================================================
// FormulaASTPool — Singleton + pool management
// ============================================================================

FormulaASTPool& FormulaASTPool::instance() {
    static FormulaASTPool pool;
    return pool;
}

uint32_t FormulaASTPool::allocNode() {
    // Caller must hold exclusive lock
    uint32_t idx = static_cast<uint32_t>(m_nodes.size());
    m_nodes.emplace_back();
    return idx;
}

uint32_t FormulaASTPool::storeLiteral(const QVariant& value) {
    uint32_t idx = static_cast<uint32_t>(m_literals.size());
    m_literals.push_back(value);
    return idx;
}

const QVariant& FormulaASTPool::getLiteral(uint32_t index) const {
    static const QVariant empty;
    if (index >= m_literals.size()) return empty;
    return m_literals[index];
}

const ASTNode& FormulaASTPool::getNode(uint32_t index) const {
    return m_nodes[index];
}

uint16_t FormulaASTPool::internFunction(const QString& name) {
    QString upper = name.toUpper();
    auto it = m_funcNameToId.find(upper);
    if (it != m_funcNameToId.end()) return it->second;
    uint16_t id = static_cast<uint16_t>(m_functionNames.size());
    m_functionNames.push_back(upper);
    m_funcNameToId[upper] = id;
    return id;
}

const QString& FormulaASTPool::getFunctionName(uint16_t id) const {
    static const QString empty;
    if (id >= m_functionNames.size()) return empty;
    return m_functionNames[id];
}

size_t FormulaASTPool::cachedFormulas() const {
    return m_cache.size();
}

size_t FormulaASTPool::totalNodes() const {
    return m_nodes.size();
}

void FormulaASTPool::clear() {
    m_nodes.clear();
    m_literals.clear();
    m_cache.clear();
    m_functionNames.clear();
    m_funcNameToId.clear();
}

// ============================================================================
// Recursive descent parser — formula string → AST node tree
// ============================================================================

namespace {

struct ParseContext {
    const QString& formula;
    int pos;
    FormulaASTPool& pool;

    ParseContext(const QString& f, FormulaASTPool& p)
        : formula(f), pos(0), pool(p) {}

    QChar peek() const {
        if (pos >= formula.length()) return QChar();
        return formula[pos];
    }

    QChar advance() {
        if (pos >= formula.length()) return QChar();
        return formula[pos++];
    }

    bool atEnd() const { return pos >= formula.length(); }

    void skipSpaces() {
        while (pos < formula.length() && formula[pos].isSpace())
            ++pos;
    }

    bool match(QChar c) {
        skipSpaces();
        if (pos < formula.length() && formula[pos] == c) {
            ++pos;
            return true;
        }
        return false;
    }
};

// Forward declarations
uint32_t parseExpr(ParseContext& ctx);
uint32_t parseComparison(ParseContext& ctx);
uint32_t parseTerm(ParseContext& ctx);
uint32_t parseMultiplicative(ParseContext& ctx);
uint32_t parsePower(ParseContext& ctx);
uint32_t parseUnary(ParseContext& ctx);
uint32_t parsePrimary(ParseContext& ctx);

// Parse cell reference like A1, $A$1, AA100
bool parseCellRef(ParseContext& ctx, int& row, int& col, bool& rowAbs, bool& colAbs) {
    int startPos = ctx.pos;
    colAbs = false;
    rowAbs = false;

    if (ctx.peek() == QLatin1Char('$')) {
        colAbs = true;
        ctx.advance();
    }

    // Column letters
    if (!ctx.peek().isLetter()) {
        ctx.pos = startPos;
        return false;
    }

    col = 0;
    while (ctx.pos < ctx.formula.length() && ctx.formula[ctx.pos].isLetter()) {
        col = col * 26 + (ctx.formula[ctx.pos].toUpper().unicode() - 'A' + 1);
        ctx.pos++;
    }
    col--; // 0-based

    if (ctx.peek() == QLatin1Char('$')) {
        rowAbs = true;
        ctx.advance();
    }

    // Row digits
    if (!ctx.peek().isDigit()) {
        ctx.pos = startPos;
        return false;
    }

    row = 0;
    while (ctx.pos < ctx.formula.length() && ctx.formula[ctx.pos].isDigit()) {
        row = row * 10 + (ctx.formula[ctx.pos].unicode() - '0');
        ctx.pos++;
    }
    row--; // 0-based

    return true;
}

// Parse number literal
bool parseNumber(ParseContext& ctx, double& value) {
    int startPos = ctx.pos;
    bool hasDigit = false;
    bool hasDot = false;

    while (ctx.pos < ctx.formula.length()) {
        QChar c = ctx.formula[ctx.pos];
        if (c.isDigit()) {
            hasDigit = true;
            ctx.pos++;
        } else if (c == QLatin1Char('.') && !hasDot) {
            hasDot = true;
            ctx.pos++;
        } else {
            break;
        }
    }

    // Scientific notation
    if (hasDigit && ctx.pos < ctx.formula.length() &&
        (ctx.formula[ctx.pos] == QLatin1Char('e') || ctx.formula[ctx.pos] == QLatin1Char('E'))) {
        ctx.pos++;
        if (ctx.pos < ctx.formula.length() &&
            (ctx.formula[ctx.pos] == QLatin1Char('+') || ctx.formula[ctx.pos] == QLatin1Char('-')))
            ctx.pos++;
        while (ctx.pos < ctx.formula.length() && ctx.formula[ctx.pos].isDigit())
            ctx.pos++;
    }

    if (!hasDigit) {
        ctx.pos = startPos;
        return false;
    }

    bool ok;
    value = QStringView(ctx.formula).mid(startPos, ctx.pos - startPos).toDouble(&ok);
    if (!ok) {
        ctx.pos = startPos;
        return false;
    }
    return true;
}

// Parse string literal "..."
bool parseString(ParseContext& ctx, QString& value) {
    if (ctx.peek() != QLatin1Char('"')) return false;
    ctx.advance(); // skip opening quote
    value.clear();
    while (!ctx.atEnd()) {
        QChar c = ctx.advance();
        if (c == QLatin1Char('"')) {
            // Escaped quote ""
            if (ctx.peek() == QLatin1Char('"')) {
                value += QLatin1Char('"');
                ctx.advance();
            } else {
                return true;
            }
        } else {
            value += c;
        }
    }
    return true; // unterminated string — accept anyway
}

// Parse function name (letters, digits, dots, underscores)
bool parseFunctionName(ParseContext& ctx, QString& name) {
    int startPos = ctx.pos;
    if (!ctx.peek().isLetter() && ctx.peek() != QLatin1Char('_')) return false;

    while (ctx.pos < ctx.formula.length()) {
        QChar c = ctx.formula[ctx.pos];
        if (c.isLetterOrNumber() || c == QLatin1Char('_') || c == QLatin1Char('.'))
            ctx.pos++;
        else
            break;
    }

    name = ctx.formula.mid(startPos, ctx.pos - startPos);
    return !name.isEmpty();
}

// Primary: number, string, bool, cell ref, range, function call, parenthesized expr
uint32_t parsePrimary(ParseContext& ctx) {
    ctx.skipSpaces();

    // String literal
    if (ctx.peek() == QLatin1Char('"')) {
        QString str;
        if (parseString(ctx, str)) {
            uint32_t node = ctx.pool.allocNode();
            ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
            n.type = ASTNodeType::Literal;
            n.literalIndex = ctx.pool.storeLiteral(QVariant(str));
            return node;
        }
    }

    // Number literal
    double numVal;
    if (ctx.peek().isDigit() || ctx.peek() == QLatin1Char('.')) {
        if (parseNumber(ctx, numVal)) {
            uint32_t node = ctx.pool.allocNode();
            ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
            n.type = ASTNodeType::Literal;
            n.literalIndex = ctx.pool.storeLiteral(QVariant(numVal));
            return node;
        }
    }

    // Parenthesized expression
    if (ctx.peek() == QLatin1Char('(')) {
        ctx.advance();
        uint32_t inner = parseExpr(ctx);
        ctx.match(QLatin1Char(')'));
        return inner;
    }

    // Try identifier: could be function, boolean, cell ref, or range
    int savedPos = ctx.pos;
    QString name;
    if (parseFunctionName(ctx, name)) {
        QString upper = name.toUpper();

        // Boolean TRUE/FALSE
        if (upper == QStringLiteral("TRUE") || upper == QStringLiteral("FALSE")) {
            uint32_t node = ctx.pool.allocNode();
            ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
            n.type = ASTNodeType::Boolean;
            n.boolValue = (upper == QStringLiteral("TRUE"));
            return node;
        }

        // Function call: NAME(args...)
        ctx.skipSpaces();
        if (ctx.peek() == QLatin1Char('(')) {
            ctx.advance();
            uint16_t funcId = ctx.pool.internFunction(upper);

            // Parse arguments
            std::vector<uint32_t> args;
            ctx.skipSpaces();
            if (ctx.peek() != QLatin1Char(')')) {
                args.push_back(parseExpr(ctx));
                while (ctx.match(QLatin1Char(','))) {
                    args.push_back(parseExpr(ctx));
                }
            }
            ctx.match(QLatin1Char(')'));

            uint32_t node = ctx.pool.allocNode();
            ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
            n.type = ASTNodeType::FunctionCall;
            n.func.funcId = funcId;
            n.func.argCount = static_cast<uint16_t>(args.size());
            // Store arg node indices in a contiguous block of literal values
            n.func.argStart = ctx.pool.storeLiteral(QVariant());
            // Overwrite: store actual arg indices as consecutive literals
            // We need a way to find arg nodes. Store them as literals.
            // Clear the placeholder and store properly:
            // Actually, let's use a simpler approach: store arg indices
            // as an array in the literal pool.
            QVariantList argList;
            for (uint32_t a : args) argList.append(a);
            // Replace the placeholder literal with the arg list
            const_cast<QVariant&>(ctx.pool.getLiteral(n.func.argStart)) = QVariant(argList);

            return node;
        }

        // Cross-sheet reference: SheetName!A1 or SheetName!A1:B10
        ctx.skipSpaces();
        if (ctx.peek() == QLatin1Char('!')) {
            ctx.advance();
            QString sheetName = name;
            uint32_t sheetNameIdx = ctx.pool.storeLiteral(QVariant(sheetName));

            // Parse cell/range ref after '!'
            int refRow, refCol;
            bool refRowAbs, refColAbs;
            if (parseCellRef(ctx, refRow, refCol, refRowAbs, refColAbs)) {
                ctx.skipSpaces();
                if (ctx.peek() == QLatin1Char(':')) {
                    ctx.advance();
                    ctx.skipSpaces();
                    int row2, col2;
                    bool rowAbs2, colAbs2;
                    if (parseCellRef(ctx, row2, col2, rowAbs2, colAbs2)) {
                        uint32_t node = ctx.pool.allocNode();
                        ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
                        n.type = ASTNodeType::CrossSheetRange;
                        n.crossSheetRange.sheetNameIndex = sheetNameIdx;
                        n.crossSheetRange.startRow = refRow;
                        n.crossSheetRange.startCol = refCol;
                        n.crossSheetRange.endRow = row2;
                        n.crossSheetRange.endCol = col2;
                        return node;
                    }
                }
                uint32_t node = ctx.pool.allocNode();
                ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
                n.type = ASTNodeType::CrossSheetCell;
                n.crossSheetCell.sheetNameIndex = sheetNameIdx;
                n.crossSheetCell.row = refRow;
                n.crossSheetCell.col = refCol;
                return node;
            }
        }

        // Not a function — backtrack and try as cell ref
        ctx.pos = savedPos;
    }

    // Quoted cross-sheet reference: 'Sheet Name'!A1
    if (ctx.peek() == QLatin1Char('\'')) {
        int savedPos = ctx.pos;
        ctx.advance(); // skip opening quote
        int nameStart = ctx.pos;
        while (!ctx.atEnd() && ctx.peek() != QLatin1Char('\'')) ctx.advance();
        QString sheetName = ctx.formula.mid(nameStart, ctx.pos - nameStart);
        if (!ctx.atEnd()) ctx.advance(); // skip closing quote

        if (ctx.peek() == QLatin1Char('!')) {
            ctx.advance();
            uint32_t sheetNameIdx = ctx.pool.storeLiteral(QVariant(sheetName));

            int refRow, refCol;
            bool refRowAbs, refColAbs;
            if (parseCellRef(ctx, refRow, refCol, refRowAbs, refColAbs)) {
                ctx.skipSpaces();
                if (ctx.peek() == QLatin1Char(':')) {
                    ctx.advance();
                    ctx.skipSpaces();
                    int row2, col2;
                    bool rowAbs2, colAbs2;
                    if (parseCellRef(ctx, row2, col2, rowAbs2, colAbs2)) {
                        uint32_t node = ctx.pool.allocNode();
                        ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
                        n.type = ASTNodeType::CrossSheetRange;
                        n.crossSheetRange.sheetNameIndex = sheetNameIdx;
                        n.crossSheetRange.startRow = refRow;
                        n.crossSheetRange.startCol = refCol;
                        n.crossSheetRange.endRow = row2;
                        n.crossSheetRange.endCol = col2;
                        return node;
                    }
                }
                uint32_t node = ctx.pool.allocNode();
                ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
                n.type = ASTNodeType::CrossSheetCell;
                n.crossSheetCell.sheetNameIndex = sheetNameIdx;
                n.crossSheetCell.row = refRow;
                n.crossSheetCell.col = refCol;
                return node;
            }
        }
        ctx.pos = savedPos; // backtrack
    }

    // Cell reference, range reference, or column reference (A1, $A$1, A1:B10, D:D)
    {
        int row, col;
        bool rowAbs, colAbs;
        int beforeRef = ctx.pos;
        if (parseCellRef(ctx, row, col, rowAbs, colAbs)) {
            ctx.skipSpaces();
            // Check for range operator ':'
            if (ctx.peek() == QLatin1Char(':')) {
                ctx.advance();
                ctx.skipSpaces();
                int row2, col2;
                bool rowAbs2, colAbs2;
                if (parseCellRef(ctx, row2, col2, rowAbs2, colAbs2)) {
                    uint32_t node = ctx.pool.allocNode();
                    ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
                    n.type = ASTNodeType::RangeRef;
                    n.rangeRef.startRow = row;
                    n.rangeRef.startCol = col;
                    n.rangeRef.endRow = row2;
                    n.rangeRef.endCol = col2;
                    return node;
                }
                // Check for column reference: A:B (second part has no digits)
                int colRefPos = ctx.pos;
                int col2Only = 0;
                bool col2Abs = false;
                if (ctx.peek() == QLatin1Char('$')) { col2Abs = true; ctx.advance(); }
                if (ctx.peek().isLetter()) {
                    while (ctx.pos < ctx.formula.length() && ctx.formula[ctx.pos].isLetter()) {
                        col2Only = col2Only * 26 + (ctx.formula[ctx.pos].toUpper().unicode() - 'A' + 1);
                        ctx.pos++;
                    }
                    col2Only--;
                    // Confirm no digits follow (pure column ref)
                    if (!ctx.peek().isDigit()) {
                        uint32_t node = ctx.pool.allocNode();
                        ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
                        n.type = ASTNodeType::ColumnRef;
                        n.colRef.startCol = col;
                        n.colRef.endCol = col2Only;
                        return node;
                    }
                }
                // Failed — backtrack to just cell ref
                ctx.pos = beforeRef;
                parseCellRef(ctx, row, col, rowAbs, colAbs);
            }

            uint32_t node = ctx.pool.allocNode();
            ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
            n.type = ASTNodeType::CellRef;
            n.cellRef.row = row;
            n.cellRef.col = col;
            n.cellRef.rowAbsolute = rowAbs;
            n.cellRef.colAbsolute = colAbs;
            return node;
        }

        // Check for column-only reference at start: D:D (no row digits before colon)
        ctx.pos = beforeRef;
        int colOnly = 0;
        bool colOnlyAbs = false;
        if (ctx.peek() == QLatin1Char('$')) { colOnlyAbs = true; ctx.advance(); }
        if (ctx.peek().isLetter()) {
            int letterStart = ctx.pos;
            while (ctx.pos < ctx.formula.length() && ctx.formula[ctx.pos].isLetter()) {
                colOnly = colOnly * 26 + (ctx.formula[ctx.pos].toUpper().unicode() - 'A' + 1);
                ctx.pos++;
            }
            colOnly--;
            if (ctx.peek() == QLatin1Char(':')) {
                ctx.advance();
                int col2 = 0;
                if (ctx.peek() == QLatin1Char('$')) ctx.advance();
                if (ctx.peek().isLetter()) {
                    while (ctx.pos < ctx.formula.length() && ctx.formula[ctx.pos].isLetter()) {
                        col2 = col2 * 26 + (ctx.formula[ctx.pos].toUpper().unicode() - 'A' + 1);
                        ctx.pos++;
                    }
                    col2--;
                    if (!ctx.peek().isDigit()) {
                        uint32_t node = ctx.pool.allocNode();
                        ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
                        n.type = ASTNodeType::ColumnRef;
                        n.colRef.startCol = colOnly;
                        n.colRef.endCol = col2;
                        return node;
                    }
                }
            }
            ctx.pos = beforeRef; // backtrack
        }
    }

    // Error node — unparseable token
    uint32_t node = ctx.pool.allocNode();
    ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
    n.type = ASTNodeType::Error;
    n.errorIndex = ctx.pool.storeLiteral(QVariant(QStringLiteral("#PARSE!")));
    // Skip one character to avoid infinite loop
    if (!ctx.atEnd()) ctx.advance();
    return node;
}

// Unary: -expr, +expr
uint32_t parseUnary(ParseContext& ctx) {
    ctx.skipSpaces();
    if (ctx.peek() == QLatin1Char('-')) {
        ctx.advance();
        uint32_t operand = parseUnary(ctx);
        uint32_t node = ctx.pool.allocNode();
        ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
        n.type = ASTNodeType::UnaryOp;
        n.unary.op = UnaryOp::Negate;
        n.unary.operand = operand;
        return node;
    }
    if (ctx.peek() == QLatin1Char('+')) {
        ctx.advance();
        return parseUnary(ctx);
    }
    return parsePrimary(ctx);
}

// Power: unary ^ unary (right-associative)
uint32_t parsePower(ParseContext& ctx) {
    uint32_t left = parseUnary(ctx);
    ctx.skipSpaces();
    if (ctx.peek() == QLatin1Char('^')) {
        ctx.advance();
        uint32_t right = parsePower(ctx); // right-associative
        uint32_t node = ctx.pool.allocNode();
        ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
        n.type = ASTNodeType::BinaryOp;
        n.binary.op = BinaryOp::Pow;
        n.binary.left = left;
        n.binary.right = right;
        return node;
    }
    return left;
}

// Multiplicative: power (* | /) power
uint32_t parseMultiplicative(ParseContext& ctx) {
    uint32_t left = parsePower(ctx);
    while (true) {
        ctx.skipSpaces();
        QChar c = ctx.peek();
        BinaryOp op;
        if (c == QLatin1Char('*')) op = BinaryOp::Mul;
        else if (c == QLatin1Char('/')) op = BinaryOp::Div;
        else break;

        ctx.advance();
        uint32_t right = parsePower(ctx);
        uint32_t node = ctx.pool.allocNode();
        ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
        n.type = ASTNodeType::BinaryOp;
        n.binary.op = op;
        n.binary.left = left;
        n.binary.right = right;
        left = node;
    }
    return left;
}

// Term: multiplicative (+ | - | &) multiplicative
uint32_t parseTerm(ParseContext& ctx) {
    uint32_t left = parseMultiplicative(ctx);
    while (true) {
        ctx.skipSpaces();
        QChar c = ctx.peek();
        BinaryOp op;
        if (c == QLatin1Char('+')) op = BinaryOp::Add;
        else if (c == QLatin1Char('-')) op = BinaryOp::Sub;
        else if (c == QLatin1Char('&')) op = BinaryOp::Concat;
        else break;

        ctx.advance();
        uint32_t right = parseMultiplicative(ctx);
        uint32_t node = ctx.pool.allocNode();
        ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
        n.type = ASTNodeType::BinaryOp;
        n.binary.op = op;
        n.binary.left = left;
        n.binary.right = right;
        left = node;
    }
    return left;
}

// Comparison: term (= | <> | < | > | <= | >=) term
uint32_t parseComparison(ParseContext& ctx) {
    uint32_t left = parseTerm(ctx);
    while (true) {
        ctx.skipSpaces();
        BinaryOp op;
        QChar c = ctx.peek();
        if (c == QLatin1Char('=')) {
            ctx.advance();
            op = BinaryOp::Eq;
        } else if (c == QLatin1Char('<')) {
            ctx.advance();
            if (ctx.peek() == QLatin1Char('>')) { ctx.advance(); op = BinaryOp::Neq; }
            else if (ctx.peek() == QLatin1Char('=')) { ctx.advance(); op = BinaryOp::Lte; }
            else op = BinaryOp::Lt;
        } else if (c == QLatin1Char('>')) {
            ctx.advance();
            if (ctx.peek() == QLatin1Char('=')) { ctx.advance(); op = BinaryOp::Gte; }
            else op = BinaryOp::Gt;
        } else {
            break;
        }

        uint32_t right = parseTerm(ctx);
        uint32_t node = ctx.pool.allocNode();
        ASTNode& n = const_cast<ASTNode&>(ctx.pool.getNode(node));
        n.type = ASTNodeType::BinaryOp;
        n.binary.op = op;
        n.binary.left = left;
        n.binary.right = right;
        left = node;
    }
    return left;
}

// Top-level expression
uint32_t parseExpr(ParseContext& ctx) {
    return parseComparison(ctx);
}

} // anonymous namespace


// ============================================================================
// FormulaASTPool::parse — Entry point
// ============================================================================
uint32_t FormulaASTPool::parse(const QString& formula) {
    // Strip leading '=' if present
    QString cleaned = formula;
    if (!cleaned.isEmpty() && cleaned[0] == QLatin1Char('='))
        cleaned = cleaned.mid(1);

    // Check cache
    auto it = m_cache.find(cleaned);
    if (it != m_cache.end()) return it->second;

    ParseContext ctx(cleaned, *this);
    uint32_t root = parseExpr(ctx);

    m_cache[cleaned] = root;
    return root;
}
