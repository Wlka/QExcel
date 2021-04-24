#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QAxObject>
#include <objbase.h>
#include <QTreeWidget>
#include <QStackedWidget>
#include <vector>

namespace Ui {
class ExcelBase;
}

class ExcelBase : public QMainWindow
{
    Q_OBJECT

public:
    explicit ExcelBase(QWidget *parent = nullptr);
    ~ExcelBase();

    void openExcelFile(const QString &filePath, bool isVisible,bool isDisplayAlerts);
    void closeExcelFile();

    int getSheetsCount();
    QAxObject *addSheet(const QString &sheetName);
    template<typename T> bool deleteSheet(T sheetName);
    template<typename T> QAxObject* getSheet(T sheetName);

    QAxObject *getRows(QAxObject *sheet);
    int getRowsCount(QAxObject *sheet);
    QAxObject *getColumns(QAxObject *sheet);
    int getColumnsCount(QAxObject *sheet);

    QVariant getCell(QAxObject* sheet,int row,int column);
    QVariant getCell(QAxObject* sheet,QString cellAddress);
    void setCell(QAxObject* sheet,int row,int column, const QString &value);
    void setCell(QAxObject* sheet, const QString &cellAddress, const QString &value);

	QVariant getRange(QAxObject* sheet, const QString &range);
	void setRange(QAxObject* sheet, const QString &range, const QString &value);
	void setRange(QAxObject* sheet, const QString &range, QList<QList<QVariant>> &value);

    void setPrintArea(QAxObject *sheet, const QString &area);
    void setPrintTitleRow(QAxObject *sheet, const QString &row);
    void setPrintTitleColumn(QAxObject *sheet, const QString &column);
    void setPrintMargin(QAxObject *sheet,
                        double topMargin=1.9, double rightMargin=1.8,
                        double bottomMargin=1.9, double leftMargin=1.8,
                        double headerMargin=0.8, double footerMargin=0.8);
    void setPrintOrientation(QAxObject *sheet,bool isCenterHorizontally,bool isCenterVertically);

	void setHeader(QAxObject *sheet, const QString &header,int position=ExcelBase::Left);
	void setFooter(QAxObject *sheet, const QString &footer,int position=ExcelBase::Left);

	void setWindowsView(QAxObject* excel,int viewMode);

	void addChart(QAxObject *sheet, const QString &chartArea, const QString &chartTitle, const QString &xValueDataSource, const QString &yValueDataSource, const QString &legendTitleSource,
                  double xMinScale,double xMaxScale,double xMajorUnit,double xMinorUnit,
                  double yMinScale,double yMaxScale,double yMajorUnit,double yMinorUnit,
                  int xTickLabelPosition,int yTickLabelPosition,int legendPosition);
	void clearChart(QAxObject*);

    QVariant readAll(QAxObject* sheet);
	QVariant castListListVariant2Variant(QList<QList<QVariant>> &res);
	QList<QList<QVariant>> castVariant2ListListVariant(const QVariant &var);
	int letterToNumber(const QString &letter);

    enum Mode
    {
        XlNormalView=1,                 //普通视图
        XlPageBreakPreview=2,           //分页预览
        XlPageLayoutView=3,             //页面布局

        xlTickLabelPositionHigh=-4127, 	//刻度线靠上或靠右
        xlTickLabelPositionLow=-4134,   //刻度线靠下或靠左
        xlTickLabelPositionNextToAxis=4, //刻度线在轴旁
        xlTickLabelPositionNone=-4142,  //无刻度线

        xlLegendPositionBottom=-4107,   //图例靠下
        xlLegendPositionCorner=2,       //图例靠右上角
        xlLegendPositionLeft=-4131,     //图例靠左
        xlLegendPositionRight=-4152,    //图例靠右
        xlLegendPositionTop=-4160,      //图例靠上

        CenterHorizontally=7,           //水平居中
        CenterVertically=8,             //垂直居中
        Left=9,                         //靠左
        Center=10,                      //居中
        Right=11                        //靠右
    };

private:
    Ui::ExcelBase *ui;



    QAxObject *excel;
    QAxObject *workBook;
	QAxObject *workBook_2;
    QAxObject *workSheets;
    QString filePath;
    //QList<QList<QVariant>> res;

    


};

#endif // MAINWINDOW_H
