#ifndef REMOVEDUPLICATESDIALOG_H
#define REMOVEDUPLICATESDIALOG_H

#include <QDialog>
#include <QCheckBox>
#include <QVector>
#include <QStringList>

class RemoveDuplicatesDialog : public QDialog {
    Q_OBJECT

public:
    explicit RemoveDuplicatesDialog(const QStringList& columnHeaders, QWidget* parent = nullptr);

    QVector<int> getSelectedColumns() const;
    bool hasHeaders() const;

private:
    QCheckBox* m_hasHeadersCheck;
    QVector<QCheckBox*> m_columnChecks;
};

#endif // REMOVEDUPLICATESDIALOG_H
