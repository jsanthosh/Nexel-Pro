#include "PivotTableDialog.h"
#include "../core/Spreadsheet.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>
#include <QLabel>
#include <QLineEdit>
#include <QComboBox>
#include <QCheckBox>
#include <QListWidget>
#include <QTableWidget>
#include <QHeaderView>
#include <QDialogButtonBox>
#include <QPushButton>
#include <QTimer>

PivotTableDialog::PivotTableDialog(std::shared_ptr<Spreadsheet> sheet,
                                   const CellRange& sourceRange,
                                   QWidget* parent)
    : QDialog(parent), m_sourceSheet(sheet), m_sourceRange(sourceRange) {
    setWindowTitle("Create Pivot Table");
    setMinimumSize(800, 560);
    resize(850, 600);

    m_previewTimer = new QTimer(this);
    m_previewTimer->setSingleShot(true);
    m_previewTimer->setInterval(300);
    connect(m_previewTimer, &QTimer::timeout, this, &PivotTableDialog::updatePreview);

    createLayout();
    populateSourceFields();

    setStyleSheet(
        "QDialog { background: #FAFBFC; }"
        "QGroupBox { font-weight: bold; border: 1px solid #D0D5DD; border-radius: 6px; "
        "margin-top: 8px; padding-top: 16px; background: white; }"
        "QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #344054; }"
        "QLineEdit { border: 1px solid #D0D5DD; border-radius: 4px; padding: 5px 8px; background: white; }"
        "QLineEdit:focus { border-color: #4A90D9; }"
        "QComboBox { border: 1px solid #D0D5DD; border-radius: 4px; padding: 5px 8px; "
        "background: white; min-height: 20px; }"
        "QComboBox::drop-down { border: none; width: 20px; }"
        "QComboBox::down-arrow { image: none; border-left: 4px solid transparent; "
        "border-right: 4px solid transparent; border-top: 5px solid #667085; margin-right: 6px; }"
        "QListWidget { border: 1px solid #D0D5DD; border-radius: 6px; background: white; outline: none; }"
        "QListWidget::item { padding: 4px 8px; border-radius: 3px; }"
        "QListWidget::item:selected { background-color: #E8F0FE; color: #1A1A1A; }"
        "QListWidget::item:hover:!selected { background-color: #F5F5F5; }"
        "QCheckBox { spacing: 6px; }"
        "QTableWidget { border: 1px solid #D0D5DD; border-radius: 4px; gridline-color: #E0E3E8; }"
        "QHeaderView::section { background: #F0F2F5; border: none; border-right: 1px solid #E0E3E8; "
        "border-bottom: 1px solid #E0E3E8; padding: 4px 6px; font-weight: bold; font-size: 11px; }"
    );
}

void PivotTableDialog::createLayout() {
    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(8);

    // Source range
    QHBoxLayout* rangeLayout = new QHBoxLayout();
    rangeLayout->addWidget(new QLabel("Source Range:"));
    m_rangeEdit = new QLineEdit(m_sourceRange.toString());
    m_rangeEdit->setReadOnly(true);
    m_rangeEdit->setMaximumWidth(200);
    rangeLayout->addWidget(m_rangeEdit);
    rangeLayout->addStretch();
    mainLayout->addLayout(rangeLayout);

    // Main content: 3 columns
    QHBoxLayout* contentLayout = new QHBoxLayout();
    contentLayout->setSpacing(8);

    // Left: Source fields
    QGroupBox* sourceGroup = new QGroupBox("Source Fields");
    QVBoxLayout* sourceLayout = new QVBoxLayout(sourceGroup);
    m_sourceFieldList = new QListWidget();
    m_sourceFieldList->setMaximumWidth(160);
    sourceLayout->addWidget(m_sourceFieldList);
    sourceGroup->setFixedWidth(180);
    contentLayout->addWidget(sourceGroup);

    // Center: Field zones
    QVBoxLayout* zonesLayout = new QVBoxLayout();
    zonesLayout->setSpacing(6);

    auto createZone = [&](const QString& title, QListWidget*& zone,
                          QPushButton*& addBtn, QPushButton*& removeBtn) {
        QGroupBox* group = new QGroupBox(title);
        QVBoxLayout* layout = new QVBoxLayout(group);
        layout->setSpacing(4);
        zone = new QListWidget();
        zone->setMaximumHeight(70);
        layout->addWidget(zone);
        QHBoxLayout* btnLayout = new QHBoxLayout();
        addBtn = new QPushButton("+");
        addBtn->setFixedSize(28, 24);
        addBtn->setStyleSheet("QPushButton { background: #E8F0FE; border: 1px solid #4A90D9; "
                               "border-radius: 4px; font-weight: bold; color: #4A90D9; }"
                               "QPushButton:hover { background: #D6E4F0; }");
        removeBtn = new QPushButton("-");
        removeBtn->setFixedSize(28, 24);
        removeBtn->setStyleSheet("QPushButton { background: #FEE8E8; border: 1px solid #D94A4A; "
                                  "border-radius: 4px; font-weight: bold; color: #D94A4A; }"
                                  "QPushButton:hover { background: #F0D6D6; }");
        btnLayout->addWidget(addBtn);
        btnLayout->addWidget(removeBtn);
        btnLayout->addStretch();
        layout->addLayout(btnLayout);
        return group;
    };

    QPushButton *addFilterBtn, *removeFilterBtn;
    QPushButton *addRowBtn, *removeRowBtn;
    QPushButton *addColBtn, *removeColBtn;
    QPushButton *addValBtn, *removeValBtn;

    zonesLayout->addWidget(createZone("Filters", m_filterZone, addFilterBtn, removeFilterBtn));
    zonesLayout->addWidget(createZone("Columns", m_columnZone, addColBtn, removeColBtn));
    zonesLayout->addWidget(createZone("Rows", m_rowZone, addRowBtn, removeRowBtn));

    // Value zone with aggregation combo
    QGroupBox* valGroup = new QGroupBox("Values");
    QVBoxLayout* valLayout = new QVBoxLayout(valGroup);
    valLayout->setSpacing(4);
    m_valueZone = new QListWidget();
    m_valueZone->setMaximumHeight(70);
    valLayout->addWidget(m_valueZone);
    QHBoxLayout* valBtnLayout = new QHBoxLayout();
    addValBtn = new QPushButton("+");
    addValBtn->setFixedSize(28, 24);
    addValBtn->setStyleSheet("QPushButton { background: #E8F0FE; border: 1px solid #4A90D9; "
                              "border-radius: 4px; font-weight: bold; color: #4A90D9; }"
                              "QPushButton:hover { background: #D6E4F0; }");
    removeValBtn = new QPushButton("-");
    removeValBtn->setFixedSize(28, 24);
    removeValBtn->setStyleSheet("QPushButton { background: #FEE8E8; border: 1px solid #D94A4A; "
                                 "border-radius: 4px; font-weight: bold; color: #D94A4A; }"
                                 "QPushButton:hover { background: #F0D6D6; }");
    valBtnLayout->addWidget(addValBtn);
    valBtnLayout->addWidget(removeValBtn);
    valBtnLayout->addWidget(new QLabel("Agg:"));
    m_aggregationCombo = new QComboBox();
    m_aggregationCombo->addItems({"SUM", "COUNT", "AVERAGE", "MIN", "MAX", "COUNT DISTINCT"});
    m_aggregationCombo->setFixedWidth(120);
    valBtnLayout->addWidget(m_aggregationCombo);
    valBtnLayout->addStretch();
    valLayout->addLayout(valBtnLayout);
    zonesLayout->addWidget(valGroup);

    contentLayout->addLayout(zonesLayout, 1);

    // Right: Preview
    QGroupBox* previewGroup = new QGroupBox("Preview");
    QVBoxLayout* previewLayout = new QVBoxLayout(previewGroup);
    m_previewTable = new QTableWidget(0, 0);
    m_previewTable->setMinimumWidth(200);
    m_previewTable->horizontalHeader()->setDefaultSectionSize(80);
    m_previewTable->verticalHeader()->setDefaultSectionSize(24);
    m_previewTable->verticalHeader()->hide();
    m_previewTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    m_previewTable->setSelectionMode(QAbstractItemView::NoSelection);
    previewLayout->addWidget(m_previewTable);
    previewGroup->setMinimumWidth(220);
    contentLayout->addWidget(previewGroup, 1);

    mainLayout->addLayout(contentLayout, 1);

    // Options row
    QHBoxLayout* optLayout = new QHBoxLayout();
    m_showGrandTotalRow = new QCheckBox("Grand Total Row");
    m_showGrandTotalRow->setChecked(true);
    m_showGrandTotalCol = new QCheckBox("Grand Total Column");
    m_showGrandTotalCol->setChecked(true);
    m_autoChart = new QCheckBox("Auto Chart");
    m_autoChart->setChecked(true);
    m_chartTypeCombo = new QComboBox();
    m_chartTypeCombo->addItems({"Column", "Bar", "Line", "Pie"});
    m_chartTypeCombo->setFixedWidth(100);
    optLayout->addWidget(m_showGrandTotalRow);
    optLayout->addWidget(m_showGrandTotalCol);
    optLayout->addSpacing(20);
    optLayout->addWidget(m_autoChart);
    optLayout->addWidget(m_chartTypeCombo);
    optLayout->addStretch();
    mainLayout->addLayout(optLayout);

    // Buttons
    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    buttons->button(QDialogButtonBox::Ok)->setText("Create Pivot Table");
    buttons->button(QDialogButtonBox::Ok)->setStyleSheet(
        "QPushButton { background: #217346; color: white; border: none; border-radius: 4px; "
        "padding: 8px 24px; font-weight: bold; }"
        "QPushButton:hover { background: #1B5E3B; }");
    buttons->button(QDialogButtonBox::Cancel)->setStyleSheet(
        "QPushButton { background: #F0F2F5; border: 1px solid #D0D5DD; border-radius: 4px; "
        "padding: 8px 20px; }"
        "QPushButton:hover { background: #E8ECF0; }");
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
    mainLayout->addWidget(buttons);

    // Connect buttons
    connect(addFilterBtn, &QPushButton::clicked, this, &PivotTableDialog::onAddFilterField);
    connect(removeFilterBtn, &QPushButton::clicked, this, &PivotTableDialog::onRemoveFilterField);
    connect(addRowBtn, &QPushButton::clicked, this, &PivotTableDialog::onAddRowField);
    connect(removeRowBtn, &QPushButton::clicked, this, &PivotTableDialog::onRemoveRowField);
    connect(addColBtn, &QPushButton::clicked, this, &PivotTableDialog::onAddColumnField);
    connect(removeColBtn, &QPushButton::clicked, this, &PivotTableDialog::onRemoveColumnField);
    connect(addValBtn, &QPushButton::clicked, this, &PivotTableDialog::onAddValueField);
    connect(removeValBtn, &QPushButton::clicked, this, &PivotTableDialog::onRemoveValueField);
    connect(m_aggregationCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &PivotTableDialog::onAggregationChanged);
    connect(m_showGrandTotalRow, &QCheckBox::toggled, this, &PivotTableDialog::schedulePreviewUpdate);
    connect(m_showGrandTotalCol, &QCheckBox::toggled, this, &PivotTableDialog::schedulePreviewUpdate);
}

void PivotTableDialog::populateSourceFields() {
    m_sourceColumns = m_engine.detectColumnHeaders(m_sourceSheet, m_sourceRange);
    for (int i = 0; i < m_sourceColumns.size(); ++i) {
        auto* item = new QListWidgetItem(m_sourceColumns[i], m_sourceFieldList);
        item->setData(Qt::UserRole, i); // column index
    }
}

void PivotTableDialog::moveFieldToZone(QListWidget* targetZone, const QString& prefix) {
    auto* sourceItem = m_sourceFieldList->currentItem();
    if (!sourceItem) return;

    int colIdx = sourceItem->data(Qt::UserRole).toInt();
    QString name = sourceItem->text();
    QString displayName = prefix.isEmpty() ? name : prefix + name;

    auto* newItem = new QListWidgetItem(displayName, targetZone);
    newItem->setData(Qt::UserRole, colIdx);

    delete m_sourceFieldList->takeItem(m_sourceFieldList->row(sourceItem));
    schedulePreviewUpdate();
}

void PivotTableDialog::moveFieldBack(QListWidget* sourceZone) {
    auto* item = sourceZone->currentItem();
    if (!item) return;

    int colIdx = item->data(Qt::UserRole).toInt();
    QString name = (colIdx >= 0 && colIdx < m_sourceColumns.size()) ? m_sourceColumns[colIdx] : "?";

    auto* newItem = new QListWidgetItem(name, m_sourceFieldList);
    newItem->setData(Qt::UserRole, colIdx);

    delete sourceZone->takeItem(sourceZone->row(item));
    schedulePreviewUpdate();
}

int PivotTableDialog::getSelectedFieldColumn(QListWidget* list) const {
    auto* item = list->currentItem();
    return item ? item->data(Qt::UserRole).toInt() : -1;
}

void PivotTableDialog::onAddRowField() { moveFieldToZone(m_rowZone); }
void PivotTableDialog::onAddColumnField() { moveFieldToZone(m_columnZone); }
void PivotTableDialog::onAddFilterField() { moveFieldToZone(m_filterZone); }

void PivotTableDialog::onAddValueField() {
    auto* sourceItem = m_sourceFieldList->currentItem();
    if (!sourceItem) return;

    int colIdx = sourceItem->data(Qt::UserRole).toInt();
    QString name = sourceItem->text();
    QString aggName = m_aggregationCombo->currentText();

    QString displayName = aggName + " of " + name;
    auto* newItem = new QListWidgetItem(displayName, m_valueZone);
    newItem->setData(Qt::UserRole, colIdx);
    newItem->setData(Qt::UserRole + 1, m_aggregationCombo->currentIndex());

    delete m_sourceFieldList->takeItem(m_sourceFieldList->row(sourceItem));
    schedulePreviewUpdate();
}

void PivotTableDialog::onRemoveRowField() { moveFieldBack(m_rowZone); }
void PivotTableDialog::onRemoveColumnField() { moveFieldBack(m_columnZone); }
void PivotTableDialog::onRemoveValueField() { moveFieldBack(m_valueZone); }
void PivotTableDialog::onRemoveFilterField() { moveFieldBack(m_filterZone); }

void PivotTableDialog::onAggregationChanged(int /*index*/) {
    // Update the display name of the currently selected value field
    auto* item = m_valueZone->currentItem();
    if (!item) return;

    int colIdx = item->data(Qt::UserRole).toInt();
    QString name = (colIdx >= 0 && colIdx < m_sourceColumns.size()) ? m_sourceColumns[colIdx] : "?";
    QString aggName = m_aggregationCombo->currentText();
    item->setText(aggName + " of " + name);
    item->setData(Qt::UserRole + 1, m_aggregationCombo->currentIndex());
    schedulePreviewUpdate();
}

void PivotTableDialog::schedulePreviewUpdate() {
    m_previewTimer->start();
}

void PivotTableDialog::updatePreview() {
    PivotConfig config = getConfig();
    if (config.valueFields.empty()) {
        m_previewTable->clear();
        m_previewTable->setRowCount(0);
        m_previewTable->setColumnCount(0);
        return;
    }

    m_engine.setSource(m_sourceSheet, config);
    PivotResult result = m_engine.compute();

    int totalCols = result.numRowHeaderColumns + static_cast<int>(result.columnLabels.size());
    int totalRows = static_cast<int>(result.rowLabels.size());
    if (config.showGrandTotalRow) totalRows++;

    // Limit preview size
    int maxRows = qMin(totalRows, 12);
    int maxCols = qMin(totalCols, 8);

    m_previewTable->clear();
    m_previewTable->setRowCount(maxRows);
    m_previewTable->setColumnCount(maxCols);

    // Column headers
    QStringList headerLabels;
    for (size_t rf = 0; rf < config.rowFields.size() && static_cast<int>(rf) < maxCols; ++rf) {
        headerLabels << config.rowFields[rf].name;
    }
    for (size_t c = 0; c < result.columnLabels.size() &&
         static_cast<int>(result.numRowHeaderColumns + c) < maxCols; ++c) {
        const auto& label = result.columnLabels[c];
        headerLabels << (label.empty() ? "" : label.back());
    }
    m_previewTable->setHorizontalHeaderLabels(headerLabels);

    // Data rows
    for (int r = 0; r < maxRows; ++r) {
        bool isGrandTotal = (r == static_cast<int>(result.rowLabels.size()));

        if (isGrandTotal) {
            auto* item = new QTableWidgetItem("Grand Total");
            QFont f = item->font();
            f.setBold(true);
            item->setFont(f);
            m_previewTable->setItem(r, 0, item);

            for (size_t c = 0; c < result.grandTotalRow.size() &&
                 static_cast<int>(result.numRowHeaderColumns + c) < maxCols; ++c) {
                int col = result.numRowHeaderColumns + static_cast<int>(c);
                auto* cell = new QTableWidgetItem(
                    QString::number(result.grandTotalRow[c].toDouble(), 'f', 0));
                QFont bf = cell->font();
                bf.setBold(true);
                cell->setFont(bf);
                m_previewTable->setItem(r, col, cell);
            }
        } else if (r < static_cast<int>(result.rowLabels.size())) {
            // Row labels
            for (size_t rf = 0; rf < result.rowLabels[r].size() && static_cast<int>(rf) < maxCols; ++rf) {
                auto* item = new QTableWidgetItem(result.rowLabels[r][rf]);
                QFont f = item->font();
                f.setBold(true);
                item->setFont(f);
                m_previewTable->setItem(r, static_cast<int>(rf), item);
            }

            // Data values
            for (size_t c = 0; c < result.data[r].size() &&
                 static_cast<int>(result.numRowHeaderColumns + c) < maxCols; ++c) {
                int col = result.numRowHeaderColumns + static_cast<int>(c);
                double val = result.data[r][c].toDouble();
                auto* cell = new QTableWidgetItem(QString::number(val, 'f', 0));
                cell->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
                if (r % 2 == 1) cell->setBackground(QColor("#E8F0FE"));
                m_previewTable->setItem(r, col, cell);
            }
        }
    }

    m_previewTable->resizeColumnsToContents();
}

PivotConfig PivotTableDialog::getConfig() const {
    PivotConfig config;
    config.sourceRange = m_sourceRange;
    config.sourceSheetIndex = 0;
    config.showGrandTotalRow = m_showGrandTotalRow ? m_showGrandTotalRow->isChecked() : true;
    config.showGrandTotalColumn = m_showGrandTotalCol ? m_showGrandTotalCol->isChecked() : true;
    config.autoChart = m_autoChart ? m_autoChart->isChecked() : false;

    // Map chart type combo to ChartType enum value
    if (m_chartTypeCombo) {
        static const int chartTypeMap[] = {6, 1, 0, 3}; // Column, Bar, Line, Pie
        int idx = m_chartTypeCombo->currentIndex();
        config.chartType = (idx >= 0 && idx < 4) ? chartTypeMap[idx] : 6;
    }

    // Row fields
    for (int i = 0; i < m_rowZone->count(); ++i) {
        auto* item = m_rowZone->item(i);
        PivotField f;
        f.sourceColumnIndex = item->data(Qt::UserRole).toInt();
        f.name = (f.sourceColumnIndex >= 0 && f.sourceColumnIndex < m_sourceColumns.size())
                 ? m_sourceColumns[f.sourceColumnIndex] : item->text();
        config.rowFields.push_back(f);
    }

    // Column fields
    for (int i = 0; i < m_columnZone->count(); ++i) {
        auto* item = m_columnZone->item(i);
        PivotField f;
        f.sourceColumnIndex = item->data(Qt::UserRole).toInt();
        f.name = (f.sourceColumnIndex >= 0 && f.sourceColumnIndex < m_sourceColumns.size())
                 ? m_sourceColumns[f.sourceColumnIndex] : item->text();
        config.columnFields.push_back(f);
    }

    // Value fields
    for (int i = 0; i < m_valueZone->count(); ++i) {
        auto* item = m_valueZone->item(i);
        PivotValueField f;
        f.sourceColumnIndex = item->data(Qt::UserRole).toInt();
        f.name = (f.sourceColumnIndex >= 0 && f.sourceColumnIndex < m_sourceColumns.size())
                 ? m_sourceColumns[f.sourceColumnIndex] : "";
        int aggIdx = item->data(Qt::UserRole + 1).toInt();
        f.aggregation = static_cast<AggregationFunction>(aggIdx);
        config.valueFields.push_back(f);
    }

    // Filter fields
    for (int i = 0; i < m_filterZone->count(); ++i) {
        auto* item = m_filterZone->item(i);
        PivotFilterField f;
        f.sourceColumnIndex = item->data(Qt::UserRole).toInt();
        f.name = (f.sourceColumnIndex >= 0 && f.sourceColumnIndex < m_sourceColumns.size())
                 ? m_sourceColumns[f.sourceColumnIndex] : item->text();
        // For now, no specific filter values (all pass)
        config.filterFields.push_back(f);
    }

    return config;
}
