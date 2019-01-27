#ifndef WIDGET_H
#define WIDGET_H

#include <QWidget>
#include <QDateTime>
#include <QFileDialog>
#include <QSettings>
#include <ActiveQt/qaxobject.h>
#include <qmessagebox.h>
namespace Ui {
class Widget;
}

class Widget : public QWidget
{
    Q_OBJECT

public:
    explicit Widget(QWidget *parent = 0);
    ~Widget();

private slots:
    void on_Online_clicked();

    void on_Opne_clicked();

private:
    Ui::Widget *ui;
    QDate Today;
    //QAxObject *pApplication;
    //QAxObject *pWorkBooks;
    //QAxObject *pWorkBook;
    //QAxObject *pSheets;
    //QAxObject *pSheet;
};

#endif // WIDGET_H
