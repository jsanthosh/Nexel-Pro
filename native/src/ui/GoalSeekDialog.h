#ifndef GOALSEEEKDIALOG_H
#define GOALSEEEKDIALOG_H

#include <QDialog>
#include <QLineEdit>
#include <QLabel>

class Spreadsheet;

class GoalSeekDialog : public QDialog {
    Q_OBJECT
public:
    explicit GoalSeekDialog(QWidget* parent = nullptr);

    QString getSetCell() const;
    double getToValue() const;
    QString getByChangingCell() const;

private:
    QLineEdit* m_setCellEdit;
    QLineEdit* m_toValueEdit;
    QLineEdit* m_byChangingEdit;
};

#endif
