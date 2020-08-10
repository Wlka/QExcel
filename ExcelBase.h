#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QAxObject>
#include <objbase.h>
#include <QTreeWidget>
#include <QStackedWidget>

namespace Ui {
class ExcelBase;
}

class ExcelBase : public QMainWindow
{
    Q_OBJECT

public:
    explicit ExcelBase(QWidget *parent = nullptr);
    ~ExcelBase();

    bool openExcelFile(QString filePath);
    bool closeExcelFile();

    int getSheetsCount();
    QAxObject *addSheet(QString &sheetName);
    template<typename T> bool deleteSheet(T sheetName);
    template<typename T> QAxObject* getSheet(T sheetName);

    QAxObject *getRows(QAxObject *sheet);
    int getRowsCount(QAxObject *sheet);
    QAxObject *getColumns(QAxObject *sheet);
    int getColumnsCount(QAxObject *sheet);

    QString getCell(QAxObject* sheet,int row,int column);
    QString getCell(QAxObject* sheet,QString &number);
    bool setCell(QAxObject* sheet,int row,int column,QString value);
    bool setCell(QAxObject* sheet, QString &number,QString &value);

    bool setPrintArea(QAxObject *sheet, QString area);
    bool setPrintTitleRow(QAxObject *sheet, QString row);
    bool setPrintTitleColumn(QAxObject *sheet, QString column);
    bool setPrintMargin(QAxObject *sheet,
                        double topMargin=1.9, double rightMargin=1.8,
                        double bottomMargin=1.9, double leftMargin=1.8,
                        double headerMargin=0.8, double footerMargin=0.8);
    bool setPrintOrientation(QAxObject *sheet,bool isCenterHorizontally,bool isCenterVertically);

    bool setHeader(QAxObject *sheet,QString header,int position=ExcelBase::Left);
    bool setFooter(QAxObject *sheet,QString footer,int position=ExcelBase::Left);

    bool setWindowsView(QAxObject*,int);

    QVariant readAll(QAxObject* sheet);
    void castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant>> &res);

    enum Mode
    {
        XlNormalView=1,             //普通视图
        XlPageBreakPreview=2,       //分页预览
        XlPageLayoutView=3,         //页面布局
        CenterHorizontally=4,       //水平居中
        CenterVertically=5,         //垂直居中
        Left=6,                     //靠左
        Center=7,                   //居中
        Right=8                     //靠右
    };

private:
    Ui::ExcelBase *ui;



    QAxObject *excel;
    QAxObject *workBook;
    QAxObject *workSheets;
    QString filePath;
    QList<QList<QVariant>> res;


};

#endif // MAINWINDOW_H
