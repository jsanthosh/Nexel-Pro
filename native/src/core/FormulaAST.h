#ifndef FORMULAAST_H
#define FORMULAAST_H

#include <cstdint>
#include <vector>
#include <unordered_map>
#include <QString>
#include <QVariant>
#include "CellRange.h"

// ============================================================================
// FormulaAST — Parse-once, evaluate-many formula cache
// ============================================================================
// Instead of re-parsing formula strings on every recalculation,
// we parse once into an AST node pool and evaluate by walking the tree.
//
// Typical speedup: ~10x for recalculation of 2M formula cells.
//
// AST nodes are stored in a flat arena (FormulaASTPool::m_nodes).
// Each formula's root is cached by formula string hash.
//

// AST node types
enum class ASTNodeType : uint8_t {
    Literal,         // numeric or string constant
    CellRef,         // reference to a single cell (may have absolute markers)
    RangeRef,        // reference to a cell range
    ColumnRef,       // full column reference (D:D, A:Z)
    BinaryOp,        // +, -, *, /, ^, &, comparison operators
    UnaryOp,         // unary minus, unary plus
    FunctionCall,    // function name + argument list
    Boolean,         // TRUE/FALSE
    Error,           // #REF!, #VALUE!, etc.
    CrossSheetCell,  // Sheet1!A1 — cross-sheet cell reference
    CrossSheetRange, // Sheet1!A1:B10 — cross-sheet range reference
};

// Binary operators
enum class BinaryOp : uint8_t {
    Add, Sub, Mul, Div, Pow, Concat,
    Eq, Neq, Lt, Gt, Lte, Gte
};

// Unary operators
enum class UnaryOp : uint8_t {
    Negate, Plus
};

// AST node stored in flat arena
struct ASTNode {
    ASTNodeType type;

    union {
        // Literal: index into value pool
        uint32_t literalIndex;

        // CellRef: packed row|col with absolute flags
        struct {
            int row;
            int col;
            bool rowAbsolute;
            bool colAbsolute;
        } cellRef;

        // RangeRef: start and end packed
        struct {
            int startRow, startCol;
            int endRow, endCol;
        } rangeRef;

        // BinaryOp: operator + indices to left and right children
        struct {
            BinaryOp op;
            uint32_t left;   // index into node arena
            uint32_t right;  // index into node arena
        } binary;

        // UnaryOp: operator + child index
        struct {
            UnaryOp op;
            uint32_t operand; // index into node arena
        } unary;

        // FunctionCall: function ID + argument range in node arena
        struct {
            uint16_t funcId;     // index into function name table
            uint16_t argCount;
            uint32_t argStart;   // index of first argument node in arena
        } func;

        // Boolean
        bool boolValue;

        // Error: index into error string pool
        uint32_t errorIndex;

        // ColumnRef: start and end column indices
        struct {
            int startCol, endCol;
        } colRef;

        // CrossSheetCell: sheet name literal index + cell row/col
        struct {
            uint32_t sheetNameIndex;  // literal pool index for sheet name
            int row, col;
        } crossSheetCell;

        // CrossSheetRange: sheet name literal index + range
        struct {
            uint32_t sheetNameIndex;
            int startRow, startCol, endRow, endCol;
        } crossSheetRange;
    };

    ASTNode() : type(ASTNodeType::Literal) { literalIndex = 0; }
};


// ============================================================================
// FormulaASTPool — Global pool of parsed formula ASTs
// ============================================================================
class FormulaASTPool {
public:
    static FormulaASTPool& instance();

    // Parse a formula string and return the root node index.
    // If the same formula was parsed before, returns the cached root.
    // Returns UINT32_MAX on parse error.
    uint32_t parse(const QString& formula);

    // Get a node by index
    const ASTNode& getNode(uint32_t index) const;

    // Allocate a new node in the arena and return its index
    uint32_t allocNode();

    // Store a literal value and return its index
    uint32_t storeLiteral(const QVariant& value);

    // Get a stored literal value
    const QVariant& getLiteral(uint32_t index) const;

    // Register or look up a function name, returning its ID
    uint16_t internFunction(const QString& name);

    // Get function name by ID
    const QString& getFunctionName(uint16_t id) const;

    // Cache statistics
    size_t cachedFormulas() const;
    size_t totalNodes() const;

    // Clear all cached ASTs (e.g., when formula structure changes globally)
    void clear();

private:
    FormulaASTPool() = default;

    // No mutex needed: parse is called from main thread only.
    // Parallel recalc uses thread-local FormulaEngines that call
    // parse() sequentially (formulas are already cached by then).

    // Arena of all AST nodes (flat, cache-friendly)
    std::vector<ASTNode> m_nodes;

    // Literal value pool
    std::vector<QVariant> m_literals;

    // Formula string → root node index cache
    std::unordered_map<QString, uint32_t> m_cache;

    // Function name table (deduplication)
    std::vector<QString> m_functionNames;
    std::unordered_map<QString, uint16_t> m_funcNameToId;
};

#endif // FORMULAAST_H
