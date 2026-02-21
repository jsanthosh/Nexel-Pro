#include "ChatPanel.h"
#include "../core/Spreadsheet.h"
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QLabel>
#include <QScrollBar>
#include <QSettings>
#include <QInputDialog>
#include <QRegularExpression>
#include <QHBoxLayout>
#include <QPainter>
#include <QPainterPath>

ChatPanel::ChatPanel(QWidget* parent)
    : QWidget(parent) {

    m_networkManager = new QNetworkAccessManager(this);
    connect(m_networkManager, &QNetworkAccessManager::finished, this, &ChatPanel::onApiResponse);

    // Main layout
    m_mainLayout = new QVBoxLayout(this);
    m_mainLayout->setContentsMargins(0, 0, 0, 0);
    m_mainLayout->setSpacing(0);

    // ---- Header ----
    auto* header = new QWidget(this);
    header->setFixedHeight(48);
    header->setStyleSheet(
        "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, "
        "stop:0 #1B5E3B, stop:1 #2E8B57);"
        "border: none;");
    auto* headerLayout = new QHBoxLayout(header);
    headerLayout->setContentsMargins(14, 0, 10, 0);

    auto* headerLabel = new QLabel("Claude Assistant", header);
    headerLabel->setStyleSheet(
        "color: white; font-weight: 600; font-size: 14px; "
        "letter-spacing: 0.3px; background: transparent; border: none;");
    headerLayout->addWidget(headerLabel);
    headerLayout->addStretch();

    auto* apiKeyBtn = new QPushButton(header);
    apiKeyBtn->setText("\u2699");
    apiKeyBtn->setToolTip("Set API Key");
    apiKeyBtn->setFixedSize(30, 30);
    apiKeyBtn->setStyleSheet(
        "QPushButton { background: rgba(255,255,255,0.15); color: white; border: none; "
        "border-radius: 15px; font-size: 16px; }"
        "QPushButton:hover { background: rgba(255,255,255,0.3); }");
    headerLayout->addWidget(apiKeyBtn);
    connect(apiKeyBtn, &QPushButton::clicked, this, [this]() {
        QString key = QInputDialog::getText(this, "Claude API Key",
            "Enter your Anthropic API key:", QLineEdit::Password, m_apiKey);
        if (!key.isEmpty()) {
            setApiKey(key);
            QSettings settings("Nexel", "Nexel");
            settings.setValue("claude_api_key", key);
        }
    });

    m_mainLayout->addWidget(header);

    // ---- Scrollable message area ----
    m_scrollArea = new QScrollArea(this);
    m_scrollArea->setWidgetResizable(true);
    m_scrollArea->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_scrollArea->setStyleSheet(
        "QScrollArea { background: #ECE5DD; border: none; }"
        "QScrollArea > QWidget > QWidget { background: #ECE5DD; }"
        "QScrollBar:vertical { width: 6px; background: transparent; margin: 2px; }"
        "QScrollBar::handle:vertical { background: rgba(0,0,0,0.2); border-radius: 3px; min-height: 30px; }"
        "QScrollBar::handle:vertical:hover { background: rgba(0,0,0,0.35); }"
        "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }");

    m_messageContainer = new QWidget();
    m_messageContainer->setStyleSheet("background: #ECE5DD;");
    m_messageLayout = new QVBoxLayout(m_messageContainer);
    m_messageLayout->setContentsMargins(8, 8, 8, 8);
    m_messageLayout->setSpacing(6);
    m_messageLayout->addStretch(); // push messages to bottom initially

    m_scrollArea->setWidget(m_messageContainer);
    m_mainLayout->addWidget(m_scrollArea, 1);

    // ---- Thinking indicator (hidden by default) ----
    m_thinkingWidget = new QWidget(this);
    m_thinkingWidget->setFixedHeight(44);
    m_thinkingWidget->setStyleSheet("background: #ECE5DD; border: none;");
    auto* thinkingLayout = new QHBoxLayout(m_thinkingWidget);
    thinkingLayout->setContentsMargins(16, 4, 16, 4);

    // Three-dot typing indicator bubble (WhatsApp style)
    m_thinkingLabel = new QLabel(m_thinkingWidget);
    m_thinkingLabel->setFixedSize(64, 32);
    m_thinkingLabel->setAlignment(Qt::AlignCenter);
    m_thinkingLabel->setStyleSheet(
        "background: white; border-radius: 16px; color: #667085; "
        "font-size: 20px; font-weight: bold; letter-spacing: 3px;");
    m_thinkingLabel->setText("\u2022 \u2022 \u2022");
    thinkingLayout->addWidget(m_thinkingLabel);
    thinkingLayout->addStretch();

    m_thinkingWidget->hide();
    m_mainLayout->addWidget(m_thinkingWidget);

    // Thinking animation timer
    m_thinkingTimer = new QTimer(this);
    m_thinkingTimer->setInterval(400);
    connect(m_thinkingTimer, &QTimer::timeout, this, &ChatPanel::onThinkingTick);

    // ---- Input area ----
    auto* inputContainer = new QWidget(this);
    inputContainer->setFixedHeight(56);
    inputContainer->setStyleSheet("background: #F0F0F0; border-top: 1px solid #D9D9D9;");
    auto* inputLayout = new QHBoxLayout(inputContainer);
    inputLayout->setContentsMargins(8, 8, 8, 8);
    inputLayout->setSpacing(8);

    m_inputField = new QLineEdit(inputContainer);
    m_inputField->setPlaceholderText("Type a message...");
    m_inputField->setStyleSheet(
        "QLineEdit { background: white; border: 1px solid #D9D9D9; border-radius: 20px; "
        "padding: 8px 16px; font-size: 13px; color: #1E293B; "
        "font-family: -apple-system, 'SF Pro Text', 'Segoe UI', system-ui, sans-serif; }"
        "QLineEdit:focus { border-color: #25D366; }");
    inputLayout->addWidget(m_inputField);

    m_sendBtn = new QPushButton(inputContainer);
    m_sendBtn->setText("\u27A4");
    m_sendBtn->setFixedSize(38, 38);
    m_sendBtn->setStyleSheet(
        "QPushButton { background: #25D366; color: white; border: none; border-radius: 19px; "
        "font-size: 16px; font-weight: bold; }"
        "QPushButton:hover { background: #1DA851; }"
        "QPushButton:disabled { background: #C8C8C8; }");
    inputLayout->addWidget(m_sendBtn);

    m_mainLayout->addWidget(inputContainer);

    // Connections
    connect(m_sendBtn, &QPushButton::clicked, this, &ChatPanel::onSendMessage);
    connect(m_inputField, &QLineEdit::returnPressed, this, &ChatPanel::onSendMessage);

    // Load saved API key
    QSettings settings("Nexel", "Nexel");
    m_apiKey = settings.value("claude_api_key").toString();

    // Add welcome message
    addWelcomeMessage();
}

void ChatPanel::addWelcomeMessage() {
    auto* welcomeWidget = new QWidget();
    welcomeWidget->setStyleSheet("background: transparent;");
    auto* wLayout = new QVBoxLayout(welcomeWidget);
    wLayout->setContentsMargins(16, 20, 16, 12);
    wLayout->setSpacing(10);

    // App icon circle
    auto* iconLabel = new QLabel();
    iconLabel->setFixedSize(48, 48);
    iconLabel->setAlignment(Qt::AlignCenter);
    iconLabel->setText("\u2728");
    iconLabel->setStyleSheet(
        "background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #22C55E, stop:1 #15803D); "
        "border-radius: 24px; font-size: 22px; color: white; border: none;");
    auto* iconRow = new QHBoxLayout();
    iconRow->addStretch();
    iconRow->addWidget(iconLabel);
    iconRow->addStretch();
    wLayout->addLayout(iconRow);

    auto* title = new QLabel("Claude Assistant");
    title->setAlignment(Qt::AlignCenter);
    title->setStyleSheet(
        "font-size: 16px; font-weight: 700; color: #111B21; background: transparent; border: none;");
    wLayout->addWidget(title);

    auto* subtitle = new QLabel("Your AI spreadsheet assistant. Ask me to modify data, insert charts, or format cells.");
    subtitle->setAlignment(Qt::AlignCenter);
    subtitle->setWordWrap(true);
    subtitle->setStyleSheet(
        "font-size: 12px; color: #667085; background: transparent; border: none; padding: 0 8px;");
    wLayout->addWidget(subtitle);

    wLayout->addSpacing(4);

    // Suggestion chips as individual rounded pills
    QStringList tips = {
        "Create a monthly budget table",
        "Insert column chart for A1:D10",
        "Add sparklines for B2:G2",
        "Run a macro to fill cells",
        "Make row 1 bold and blue"
    };
    for (const auto& tip : tips) {
        auto* chipRow = new QHBoxLayout();
        chipRow->setContentsMargins(0, 0, 0, 0);

        auto* chip = new QLabel(tip);
        chip->setWordWrap(true);
        chip->setAlignment(Qt::AlignCenter);
        chip->setStyleSheet(
            "background: white; color: #1B5E3B; border-radius: 14px; "
            "padding: 7px 14px; font-size: 11px; border: none;");
        chip->setMaximumWidth(230);

        chipRow->addStretch();
        chipRow->addWidget(chip);
        chipRow->addStretch();
        wLayout->addLayout(chipRow);
    }

    wLayout->addSpacing(4);

    auto* keyHint = new QLabel("Click \u2699 to set your API key");
    keyHint->setAlignment(Qt::AlignCenter);
    keyHint->setStyleSheet(
        "font-size: 10px; color: #94A3B8; background: transparent; border: none;");
    wLayout->addWidget(keyHint);

    // Insert before the stretch
    m_messageLayout->insertWidget(m_messageLayout->count() - 1, welcomeWidget);
}

void ChatPanel::setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet) {
    m_spreadsheet = spreadsheet;
}

void ChatPanel::setApiKey(const QString& apiKey) {
    m_apiKey = apiKey;
}

void ChatPanel::scrollToBottom() {
    QTimer::singleShot(10, this, [this]() {
        auto* sb = m_scrollArea->verticalScrollBar();
        sb->setValue(sb->maximum());
    });
}

void ChatPanel::showThinkingIndicator() {
    m_thinkingDots = 0;
    m_thinkingWidget->show();
    m_thinkingTimer->start();
    m_sendBtn->setEnabled(false);
    m_inputField->setEnabled(false);
    scrollToBottom();
}

void ChatPanel::hideThinkingIndicator() {
    m_thinkingTimer->stop();
    m_thinkingWidget->hide();
    m_sendBtn->setEnabled(true);
    m_inputField->setEnabled(true);
    m_inputField->setFocus();
}

void ChatPanel::onThinkingTick() {
    m_thinkingDots = (m_thinkingDots + 1) % 4;

    // Animate dots with fading colors (WhatsApp-style typing animation)
    QString dots;
    for (int i = 0; i < 3; ++i) {
        int alpha = (i == (m_thinkingDots % 3)) ? 255 : 100;
        dots += QString("<span style='color: rgba(102,112,133,%1); font-size: 24px;'>\u2022</span> ").arg(alpha);
    }
    m_thinkingLabel->setText(dots);
}

void ChatPanel::onSendMessage() {
    QString text = m_inputField->text().trimmed();
    if (text.isEmpty()) return;

    m_inputField->clear();

    // Show user message immediately
    addMessage("You", text, true);

    if (m_apiKey.isEmpty()) {
        addMessage("Claude", "Please set your API key first using the \u2699 button above.", false);
        return;
    }

    showThinkingIndicator();
    sendToApi(text);
}

void ChatPanel::addMessage(const QString& sender, const QString& text, bool isUser) {
    Q_UNUSED(sender);

    auto* rowWidget = new QWidget();
    rowWidget->setStyleSheet("background: transparent;");
    auto* rowLayout = new QHBoxLayout(rowWidget);
    rowLayout->setContentsMargins(4, 2, 4, 2);
    rowLayout->setSpacing(0);

    auto* bubble = new QLabel();
    bubble->setWordWrap(true);
    bubble->setTextFormat(Qt::PlainText);
    bubble->setText(text);
    bubble->setMaximumWidth(240);
    bubble->setSizePolicy(QSizePolicy::Preferred, QSizePolicy::Minimum);

    if (isUser) {
        // User bubble: right-aligned, green (WhatsApp style)
        bubble->setStyleSheet(
            "background: #DCF8C6; color: #111B21; "
            "border-radius: 12px; border-top-right-radius: 2px; "
            "padding: 8px 12px; font-size: 13px; "
            "font-family: -apple-system, 'SF Pro Text', system-ui, sans-serif;");
        rowLayout->addStretch();
        rowLayout->addWidget(bubble);
    } else {
        // Assistant bubble: left-aligned, white
        bubble->setStyleSheet(
            "background: white; color: #111B21; "
            "border-radius: 12px; border-top-left-radius: 2px; "
            "padding: 8px 12px; font-size: 13px; "
            "font-family: -apple-system, 'SF Pro Text', system-ui, sans-serif;");
        rowLayout->addWidget(bubble);
        rowLayout->addStretch();
    }

    // Insert before the bottom stretch
    m_messageLayout->insertWidget(m_messageLayout->count() - 1, rowWidget);
    scrollToBottom();
}

void ChatPanel::addSystemMessage(const QString& text) {
    auto* rowWidget = new QWidget();
    rowWidget->setStyleSheet("background: transparent;");
    auto* rowLayout = new QHBoxLayout(rowWidget);
    rowLayout->setContentsMargins(20, 2, 20, 2);

    auto* label = new QLabel();
    label->setWordWrap(true);
    label->setText(QString("\u2713 %1").arg(text));
    label->setAlignment(Qt::AlignCenter);
    label->setStyleSheet(
        "background: rgba(255,255,255,0.6); color: #15803D; "
        "border-radius: 8px; padding: 4px 12px; font-size: 11px; font-weight: 500;");
    label->setMaximumWidth(260);

    rowLayout->addStretch();
    rowLayout->addWidget(label);
    rowLayout->addStretch();

    m_messageLayout->insertWidget(m_messageLayout->count() - 1, rowWidget);
    scrollToBottom();
}

QString ChatPanel::buildContext() const {
    if (!m_spreadsheet) return "";

    QString context;
    context += "Sheet: " + m_spreadsheet->getSheetName() + "\n";
    context += "Rows: " + QString::number(m_spreadsheet->getMaxRow() + 1) + "\n";
    context += "Cols: " + QString::number(m_spreadsheet->getMaxColumn() + 1) + "\n";

    // Tables
    const auto& tables = m_spreadsheet->getTables();
    if (!tables.empty()) {
        context += "Tables: ";
        for (const auto& t : tables) {
            context += t.name + " (" + t.theme.name + ") ";
        }
        context += "\n";
    }

    // Merged regions
    const auto& merged = m_spreadsheet->getMergedRegions();
    if (!merged.empty()) {
        context += "Merged regions: " + QString::number(merged.size()) + "\n";
    }

    // Sample data
    int maxRow = qMin(m_spreadsheet->getMaxRow(), 14);
    int maxCol = qMin(m_spreadsheet->getMaxColumn(), 9);

    if (maxRow >= 0 && maxCol >= 0) {
        context += "\nData (first rows/cols):\n";
        context += "\t";
        for (int c = 0; c <= maxCol; ++c) {
            context += QString(QChar('A' + c)) + "\t";
        }
        context += "\n";
        for (int r = 0; r <= maxRow; ++r) {
            context += QString::number(r + 1) + "\t";
            for (int c = 0; c <= maxCol; ++c) {
                CellAddress addr(r, c);
                auto val = m_spreadsheet->getCellValue(addr);
                QString cellStr = val.toString();
                if (cellStr.length() > 20) cellStr = cellStr.left(17) + "...";
                context += cellStr + "\t";
            }
            context += "\n";
        }
    }

    return context;
}

void ChatPanel::sendToApi(const QString& userMessage) {
    m_pendingUserMessage = userMessage;

    QNetworkRequest request(QUrl("https://api.anthropic.com/v1/messages"));
    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");
    request.setRawHeader("x-api-key", m_apiKey.toUtf8());
    request.setRawHeader("anthropic-version", "2023-06-01");

    QString context = buildContext();

    // Available table themes for reference
    QString themeList;
    auto themes = getBuiltinTableThemes();
    for (int i = 0; i < static_cast<int>(themes.size()); ++i) {
        themeList += QString::number(i) + "=" + themes[i].name + ", ";
    }

    QString systemPrompt =
        "You are Claude, an AI spreadsheet assistant inside Nexel. "
        "You can explain things AND directly modify the spreadsheet, insert charts, and add shapes by returning action blocks.\n\n"
        "Return actions using this EXACT format:\n"
        "[ACTIONS]\n"
        "[\n"
        "  {\"action\": \"set_cell\", \"cell\": \"A1\", \"value\": \"Hello\"},\n"
        "  {\"action\": \"format\", \"range\": \"A1:D1\", \"bold\": true, \"bg_color\": \"#4472C4\"},\n"
        "  {\"action\": \"insert_chart\", \"type\": \"column\", \"range\": \"A1:D10\", \"title\": \"Sales\"},\n"
        "  {\"action\": \"insert_shape\", \"type\": \"star\", \"fill_color\": \"#FFD700\", \"text\": \"Hello\"}\n"
        "]\n"
        "[/ACTIONS]\n\n"
        "Available actions:\n"
        "- set_cell: Set cell value. Fields: cell, value (string or number)\n"
        "- set_formula: Set formula. Fields: cell, formula (starts with =)\n"
        "- format: Apply formatting. Fields: range, and any of: bold, italic, underline, strikethrough (bool), "
        "bg_color, fg_color (hex like \"#4472C4\"), font_size (int), font_name (string), "
        "h_align (\"left\"/\"center\"/\"right\"), v_align (\"top\"/\"middle\"/\"bottom\")\n"
        "- merge/unmerge: Merge or unmerge cells. Fields: range\n"
        "- border: Apply borders. Fields: range, type (\"all\"/\"outside\"/\"none\"/\"bottom\"/\"top\"/\"left\"/\"right\"/\"thick_outside\")\n"
        "- table: Apply table theme. Fields: range, theme (index 0-11). Themes: " + themeList + "\n"
        "- number_format: Set number format. Fields: range, format (\"General\"/\"Number\"/\"Currency\"/\"Percentage\"/\"Date\"/\"Text\")\n"
        "- set_row_height: Set row height. Fields: row (1-based), height (pixels)\n"
        "- set_col_width: Set column width. Fields: col (letter), width (pixels)\n"
        "- clear: Clear cell values and formatting. Fields: range\n"
        "- insert_chart: Insert a chart. Fields: type (\"column\"/\"bar\"/\"line\"/\"area\"/\"scatter\"/\"pie\"/\"donut\"/\"histogram\"), "
        "range (data range like \"A1:D10\"), title (optional), x_axis (optional), y_axis (optional), theme (0-5, optional)\n"
        "- insert_shape: Insert a shape. Fields: type (\"rectangle\"/\"rounded_rect\"/\"circle\"/\"ellipse\"/\"triangle\"/\"star\"/\"arrow\"/\"diamond\"/\"pentagon\"/\"hexagon\"/\"callout\"/\"line\"), "
        "fill_color (hex, optional), stroke_color (hex, optional), text (optional), text_color (hex, optional), width (pixels, optional), height (pixels, optional)\n"
        "- insert_sparkline: Insert in-cell sparkline. Fields: cell (destination like \"A2\"), data_range (like \"B2:G2\"), "
        "type (\"line\"/\"column\"/\"winloss\", default \"line\"), color (hex, optional), show_high (bool, optional), show_low (bool, optional)\n"
        "- insert_image: Insert floating image. Fields: path (file path), width (pixels, optional), height (pixels, optional)\n"
        "- run_macro: Execute JavaScript macro. Fields: code (JS string using sheet.getCellValue/setCellValue/setBold etc.)\n"
        "- record_macro: Start/stop macro recording. Fields: action (\"start\"/\"stop\")\n\n"
        "Rules:\n"
        "- Always explain what you're doing in plain text BEFORE the [ACTIONS] block\n"
        "- Use cell references like A1, B2, AA1. Ranges use colon: A1:D10\n"
        "- For formulas, use standard Excel syntax starting with =\n"
        "- When user says \"insert chart for X\" or \"chart for X\", determine the data range from the spreadsheet data that matches column headers containing X\n"
        "- You can combine many actions in one response\n"
        "- Be concise but friendly\n";

    if (!context.isEmpty()) {
        systemPrompt += "\nCurrent spreadsheet state:\n" + context;
    }

    QJsonObject body;
    body["model"] = "claude-sonnet-4-5-20250929";
    body["max_tokens"] = 4096;
    body["system"] = systemPrompt;

    QJsonArray messages;
    QJsonObject msg;
    msg["role"] = "user";
    msg["content"] = userMessage;
    messages.append(msg);
    body["messages"] = messages;

    m_networkManager->post(request, QJsonDocument(body).toJson());
}

QString ChatPanel::extractAndProcessActions(const QString& responseText) {
    QRegularExpression re("\\[ACTIONS\\]\\s*(\\[.*?\\])\\s*\\[/ACTIONS\\]",
                          QRegularExpression::DotMatchesEverythingOption);
    QRegularExpressionMatch match = re.match(responseText);

    if (!match.hasMatch()) return responseText;

    QString jsonStr = match.captured(1);
    QJsonParseError parseError;
    QJsonDocument doc = QJsonDocument::fromJson(jsonStr.toUtf8(), &parseError);

    if (parseError.error != QJsonParseError::NoError || !doc.isArray()) {
        return responseText;
    }

    QJsonArray actions = doc.array();

    if (!actions.isEmpty()) {
        emit executeActions(actions);

        int setCells = 0, formulas = 0, formats = 0, merges = 0, borders = 0, tables = 0;
        int charts = 0, shapes = 0, sparklines = 0, images = 0, macros = 0, other = 0;
        for (const auto& item : actions) {
            QJsonObject obj = item.toObject();
            QString type = obj["action"].toString();
            if (type == "set_cell") setCells++;
            else if (type == "set_formula") formulas++;
            else if (type == "format") formats++;
            else if (type == "merge" || type == "unmerge") merges++;
            else if (type == "border") borders++;
            else if (type == "table") tables++;
            else if (type == "insert_chart") charts++;
            else if (type == "insert_shape") shapes++;
            else if (type == "insert_sparkline") sparklines++;
            else if (type == "insert_image") images++;
            else if (type == "run_macro" || type == "record_macro") macros++;
            else other++;
        }

        QStringList parts;
        if (setCells > 0) parts << QString::number(setCells) + " cell(s) set";
        if (formulas > 0) parts << QString::number(formulas) + " formula(s)";
        if (formats > 0) parts << QString::number(formats) + " format(s)";
        if (merges > 0) parts << QString::number(merges) + " merge(s)";
        if (borders > 0) parts << QString::number(borders) + " border(s)";
        if (tables > 0) parts << QString::number(tables) + " table(s)";
        if (charts > 0) parts << QString::number(charts) + " chart(s)";
        if (shapes > 0) parts << QString::number(shapes) + " shape(s)";
        if (sparklines > 0) parts << QString::number(sparklines) + " sparkline(s)";
        if (images > 0) parts << QString::number(images) + " image(s)";
        if (macros > 0) parts << QString::number(macros) + " macro(s)";
        if (other > 0) parts << QString::number(other) + " other";
        addSystemMessage("Applied: " + parts.join(", "));
    }

    QString cleanText = responseText;
    cleanText.replace(re, "");
    cleanText = cleanText.trimmed();
    return cleanText;
}

void ChatPanel::onApiResponse(QNetworkReply* reply) {
    hideThinkingIndicator();

    if (reply->error() != QNetworkReply::NoError) {
        QString errorMsg = reply->errorString();
        if (reply->error() == QNetworkReply::AuthenticationRequiredError) {
            errorMsg = "Invalid API key. Please check your key and try again.";
        }
        addMessage("Claude", "Error: " + errorMsg, false);
        reply->deleteLater();
        return;
    }

    QByteArray data = reply->readAll();
    QJsonDocument doc = QJsonDocument::fromJson(data);
    QJsonObject obj = doc.object();

    QJsonArray content = obj["content"].toArray();
    QString responseText;
    for (const auto& item : content) {
        QJsonObject block = item.toObject();
        if (block["type"].toString() == "text") {
            responseText += block["text"].toString();
        }
    }

    if (responseText.isEmpty()) {
        responseText = "Sorry, I couldn't process that request. Please try again.";
    }

    QString displayText = extractAndProcessActions(responseText);

    if (!displayText.isEmpty()) {
        addMessage("Claude", displayText, false);
    }

    reply->deleteLater();
}
