#ifndef SPARKLINEDIALOG_H
#define SPARKLINEDIALOG_H

#include <QDialog>
#include "../core/SparklineConfig.h"

class QLineEdit;
class QComboBox;
class QCheckBox;
class QPushButton;
class QColorDialog;

class SparklineDialog : public QDialog {
    Q_OBJECT

public:
    explicit SparklineDialog(QWidget* parent = nullptr);

    SparklineConfig getConfig() const;
    QString getDestinationRange() const;
    void setDataRange(const QString& range);
    void setDestination(const QString& dest);

private:
    void createLayout();

    QLineEdit* m_dataRangeEdit = nullptr;
    QLineEdit* m_destinationEdit = nullptr;
    QComboBox* m_typeCombo = nullptr;
    QPushButton* m_lineColorBtn = nullptr;
    QPushButton* m_highColorBtn = nullptr;
    QPushButton* m_lowColorBtn = nullptr;
    QCheckBox* m_showHighCheck = nullptr;
    QCheckBox* m_showLowCheck = nullptr;

    QColor m_lineColor = QColor("#4472C4");
    QColor m_highColor = QColor("#22C55E");
    QColor m_lowColor = QColor("#EF4444");
};

#endif // SPARKLINEDIALOG_H
