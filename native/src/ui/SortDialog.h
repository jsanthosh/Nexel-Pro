#ifndef SORTDIALOG_H
#define SORTDIALOG_H

#include <QDialog>
#include <QComboBox>
#include <QCheckBox>
#include <QPushButton>
#include <QVBoxLayout>
#include <vector>

struct SortLevel {
    int column;      // 0-based column index
    bool ascending;  // true = A-Z
};

class SortDialog : public QDialog {
    Q_OBJECT
public:
    explicit SortDialog(int startCol, int endCol, bool hasHeaders, QWidget* parent = nullptr);

    std::vector<SortLevel> getSortLevels() const;
    bool hasHeaders() const;

private:
    struct LevelRow {
        QComboBox* columnCombo;
        QComboBox* orderCombo;
        QWidget* rowWidget;
    };

    void addLevel();
    void deleteLevel();
    QString columnLabel(int col) const;

    int m_startCol;
    int m_endCol;
    QCheckBox* m_headersCheckbox;
    QPushButton* m_addBtn;
    QPushButton* m_deleteBtn;
    QVBoxLayout* m_levelsLayout;
    std::vector<LevelRow> m_levels;

    static constexpr int MAX_LEVELS = 4;
};

#endif // SORTDIALOG_H
