#include "excelbase.h"
#include "ui_excelbase.h"
#include <QMessageBox>
#include <QDir>
#include <QRegExp>
#include <QDebug>

ExcelBase::ExcelBase(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::ExcelBase)
{
    ui->setupUi(this);

    qDebug()<<openExcelFile("C:/Users/Wlka/Desktop/data.xlsx");
    QAxObject *workSheet=getSheet(1);

    qDebug()<<addChart(workSheet,"$H$1:$J$47","","$A$3:$A$47","$B$3:$G$47","$B$1:$G$1",-20,40,10,10,-24,0,1,1,
                       ExcelBase::xlTickLabelPositionHigh,ExcelBase::xlTickLabelPositionLow,ExcelBase::xlLegendPositionBottom);
}

ExcelBase::~ExcelBase()
{
    closeExcelFile();   //关闭excel文件，防止excel程序留在后台
    delete ui;
}

//打开excel文件
bool ExcelBase::openExcelFile(QString filePath)
{
    this->filePath=filePath;

    //初始化excel对象
    CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    excel=new(std::nothrow) QAxObject();
    if(!excel)
    {
        QMessageBox::critical(this,"错误","创建Excel对象失败",QMessageBox::Ok,QMessageBox::Ok);
        return false;
    }
    try
    {
        //连接excel，并打开文件，获取所有的sheet
        excel->setControl("Excel.Application");
        excel->dynamicCall("SetVisible(bool Visible)","true");
        excel->setProperty("DisplayAlerts","false");
        workBook=excel->querySubObject("WorkBooks")->querySubObject("Open(QString&)",filePath);
        workSheets=workBook->querySubObject("WorkSheets");
    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","打开文件失败",QMessageBox::Ok,QMessageBox::Ok);
        return false;
    }
    return true;
}

//关闭指定的excel文件
bool ExcelBase::closeExcelFile()
{
    if(excel)
    {
//        workBook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filePath));
        workBook->dynamicCall("Save()");
        workBook->dynamicCall("Close()");
        workBook->dynamicCall("Quit()");
        delete excel;
        excel=nullptr;
    }
    return true;
}

//获取sheet的数量
int ExcelBase::getSheetsCount()
{
    return workSheets->property("Count").toInt();
}

//添加一个sheet
QAxObject* ExcelBase::addSheet(QString &sheetName)
{
    QAxObject *tmpWorkSheet=nullptr;
    try
    {
        int count=getSheetsCount();
        QAxObject *lastSheet=workSheets->querySubObject("Item(int)",count);
        tmpWorkSheet=workSheets->querySubObject("Add(QVariant)",lastSheet->asVariant());
        lastSheet->dynamicCall("Move(QVariant)",tmpWorkSheet->asVariant());
        tmpWorkSheet->setProperty("Name",sheetName);
    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","创建sheet失败",QMessageBox::Ok,QMessageBox::Ok);
    }
    return tmpWorkSheet;
}

//删除指定的sheet
template<typename T> bool ExcelBase::deleteSheet(T sheetMark)
{
    try
    {
        if(std::is_same<T,int>::value)
        {
            workSheets->querySubObject("Item(int)",sheetMark)->dynamicCall("delete");
        }
        else if(std::is_same<T,QString>::value)
        {
            workSheets->querySubObject("Item(QString)",sheetMark)->dynamicCall("delete");
        }
    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","删除sheet失败",QMessageBox::Ok,QMessageBox::Ok);
        return false;
    }
    return true;
}

//获取指定的sheet
template<typename T> QAxObject* ExcelBase::getSheet(T sheetMark)
{
    QAxObject *tmpWorkSheet=nullptr;
    try
    {
        if(std::is_same<T,int>::value)
        {
            tmpWorkSheet=workSheets->querySubObject("Item(int)",sheetMark);
        }
        else if(std::is_same<T,QString>::value)
        {
            tmpWorkSheet=workSheets->querySubObject("Item(QString)",sheetMark);
        }

    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","获取sheet失败",QMessageBox::Ok,QMessageBox::Ok);
    }
    return tmpWorkSheet;
}

//获取指定sheet的行
QAxObject* ExcelBase::getRows(QAxObject *sheet)
{
    QAxObject* rows=nullptr;
    try
    {
        rows=sheet->querySubObject("Rows");
    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","获取行失败",QMessageBox::Ok,QMessageBox::Ok);
    }
    return rows;
}

//获取指定sheet的行数
int ExcelBase::getRowsCount(QAxObject *sheet)
{
    int rows=0;
    try
    {
        rows=getRows(sheet)->property("Count").toInt();
    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","获取行数失败",QMessageBox::Ok,QMessageBox::Ok);
    }
    return rows;
}

//获取指定sheet的列
QAxObject* ExcelBase::getColumns(QAxObject *sheet)
{
    QAxObject* columns=nullptr;
    try
    {
        columns=sheet->querySubObject("Rows");
    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","获取列失败",QMessageBox::Ok,QMessageBox::Ok);
    }
    return columns;
}

//获取指定sheet的列数
int ExcelBase::getColumnsCount(QAxObject *sheet)
{
    int columns=0;
    try
    {
        columns=getRows(sheet)->property("Count").toInt();
    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","获取列数失败",QMessageBox::Ok,QMessageBox::Ok);
    }
    return columns;
}

//获取指定单元格内容
QVariant ExcelBase::getCell(QAxObject* sheet,int row,int column)
{
    return sheet->querySubObject("Cells(int,int)",row,column)->property("Value");
}

//获取指定单元格内容
QVariant ExcelBase::getCell(QAxObject* sheet, QString &number)
{
    return sheet->querySubObject("Range(QString)",number)->property("Value");
}

//设置指定单元格内容
bool ExcelBase::setCell(QAxObject* sheet,int row,int column,QString value)
{
    try
    {
        sheet->querySubObject("Cells(int,int)",row,column)->setProperty("Value",value);
    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","写入单元格信息失败",QMessageBox::Ok,QMessageBox::Ok);
        return false;
    }
    return true;
}

//设置指定单元格内容
bool ExcelBase::setCell(QAxObject* sheet, QString &number,QString &value)
{
    try
    {
        sheet->querySubObject("Range(QString)", number)->setProperty("Value", value);
    }
    catch(...)
    {
        QMessageBox::critical(this,"错误","写入单元格信息失败",QMessageBox::Ok,QMessageBox::Ok);
        return false;
    }
    return true;
}

//设置打印区域
bool ExcelBase::setPrintArea(QAxObject *sheet, QString area)
{
    try{
        sheet->querySubObject("PageSetup")->setProperty("PrintArea",area);
    }
    catch(...)
    {
        return false;
    }
    return true;
}

//设置打印重复标题行
bool ExcelBase::setPrintTitleRow(QAxObject *sheet, QString row)
{
    try
    {
        sheet->querySubObject("Pagesetup")->setProperty("PrintTitleRows",row);

    }
    catch(...)
    {
        return false;
    }
    return true;
}

//设置打印重复左侧列
bool ExcelBase::setPrintTitleColumn(QAxObject *sheet, QString column)
{
    try{
        sheet->querySubObject("Pagesetup")->setProperty("PrintTitleColumns",column);
    }
    catch(...)
    {
        return false;
    }
    return true;
}

//设置页边距
bool ExcelBase::setPrintMargin(QAxObject *sheet,
                               double topMargin, double rightMargin,double bottomMargin,
                               double leftMargin, double headerMargin, double footerMargin)
{
    try{
        sheet->querySubObject("Pagesetup")->setProperty("TopMargin",28.35*topMargin);
        sheet->querySubObject("Pagesetup")->setProperty("RightMargin",28.35*rightMargin);
        sheet->querySubObject("Pagesetup")->setProperty("BottomMargin",28.35*bottomMargin);
        sheet->querySubObject("Pagesetup")->setProperty("LeftMargin",28.35*leftMargin);
        sheet->querySubObject("Pagesetup")->setProperty("HeaderMargin",28.35*headerMargin);
        sheet->querySubObject("Pagesetup")->setProperty("FooterMargin",28.35*footerMargin);
    }
    catch(...)
    {
        return false;
    }
    return true;
}

//设置打印时是否居中
bool ExcelBase::setPrintOrientation(QAxObject *sheet, bool isCenterHorizontally, bool isCenterVertically)
{
    try{
        sheet->querySubObject("Pagesetup")->setProperty("CenterHorizontally",isCenterHorizontally);
        sheet->querySubObject("Pagesetup")->setProperty("CenterVertically",isCenterVertically);
    }
    catch(...)
    {
        return false;
    }
    return true;
}

//设置页眉
bool ExcelBase::setHeader(QAxObject *sheet, QString header, int position)
{
    try{
        switch (position)
        {
        case ExcelBase::Left:
            sheet->querySubObject("Pagesetup")->setProperty("LeftHeader",header);
            break;
        case ExcelBase::Center:
            sheet->querySubObject("Pagesetup")->setProperty("CenterHeader",header);
            break;
        case ExcelBase::Right:
            sheet->querySubObject("Pagesetup")->setProperty("RightHeader",header);
            break;
        }
    }
    catch(...)
    {
        return false;
    }
    return true;
}

//设置页脚
bool ExcelBase::setFooter(QAxObject *sheet, QString footer, int position)
{
    try{
        switch (position)
        {
        case ExcelBase::Left:
            sheet->querySubObject("Pagesetup")->setProperty("LeftFooter",footer);
            break;
        case ExcelBase::Center:
            sheet->querySubObject("Pagesetup")->setProperty("CenterFooter",footer);
            break;
        case ExcelBase::Right:
            sheet->querySubObject("Pagesetup")->setProperty("RightFooter",footer);
            break;
        }
    }
    catch(...)
    {
        return false;
    }
    return true;
}

//设置活动页面的布局
bool ExcelBase::setWindowsView(QAxObject *excel,int viewMode)
{
    try{
        excel->querySubObject("Windows")->querySubObject("Item(int)",1)->setProperty("View",viewMode);
    }
    catch(...)
    {
        return false;
    }
    return true;
}

//插入Chart
//TODO 修改成适合多种图形的函数
bool ExcelBase::addChart(QAxObject *sheet,QString chartArea,QString chartTitle,QString xValueDataSource,QString yValueDataSource,QString legendTitleSource,
                         double xMinScale,double xMaxScale,double xMajorUnit,double xMinorUnit,
                         double yMinScale,double yMaxScale,double yMajorUnit,double yMinorUnit,
                         int xTickLabelPosition,int yTickLabelPosition,int legendPosition)
{
    try
    {
        sheet->querySubObject("Shapes")->querySubObject("AddChart2(240, xlXYScatterLines)")->dynamicCall("Select(void)");
        QAxObject *chart=excel->querySubObject("ActiveChart");
        chart->setProperty("HasTitle","True");
        chart->querySubObject("ChartTitle")->setProperty("Text",chartTitle);
        chart->setProperty("HasLegend","True");
        chart->querySubObject("Legend")->setProperty("Position",legendPosition);

        QStringList xDataRange=xValueDataSource.split(QRegExp("[$:]"));
        QStringList yDataRange=yValueDataSource.split(QRegExp("[$:]"));
        QStringList legendTitleRange=legendTitleSource.split(QRegExp("[$:]"));

        QList<QVariant> listValues;
        for(int i=xDataRange[2].toInt();i<=xDataRange[5].toInt();++i)
        {
            listValues.push_back(getCell(sheet,i,letterToNumber(xDataRange[1])));
        }

        for(int i=letterToNumber(yDataRange[1]),seriesCollectionCnt=1;i<=letterToNumber(yDataRange[4]);++i)
        {
            QList<QVariant> listXValues;
            for(int j=yDataRange[2].toInt();j<=yDataRange[5].toInt();++j)
            {
                if(!getCell(sheet,j,i).isValid())
                    break;
                listXValues.push_back(getCell(sheet,j,i));
            }
            chart->querySubObject("SeriesCollection()")->dynamicCall("NewSeries(void)");
            QAxObject *seriesCollection=chart->querySubObject("SeriesCollection(int)",seriesCollectionCnt++);
            seriesCollection->setProperty("Name",getCell(sheet,legendTitleRange[2].toInt(),i));
            seriesCollection->setProperty("Values",listValues);
            seriesCollection->setProperty("XValues",listXValues);
        }

        chart->querySubObject("Axes(xlCategory)")->setProperty("MinimumScale",xMinScale);
        chart->querySubObject("Axes(xlCategory)")->setProperty("MaximumScale",xMaxScale);
        chart->querySubObject("Axes(xlCategory)")->setProperty("MajorUnit",xMajorUnit);
        chart->querySubObject("Axes(xlCategory)")->setProperty("MinorUnit",xMinorUnit);
        chart->querySubObject("Axes(xlCategory)")->setProperty("TickLabelPosition",xTickLabelPosition);

        chart->querySubObject("Axes(xlValue)")->setProperty("MinimumScale",yMinScale);
        chart->querySubObject("Axes(xlValue)")->setProperty("MaximumScale",yMaxScale);
        chart->querySubObject("Axes(xlValue)")->setProperty("MajorUnit",yMajorUnit);
        chart->querySubObject("Axes(xlValue)")->setProperty("MinorUnit",yMinorUnit);
        chart->querySubObject("Axes(xlValue)")->setProperty("TickLabelPosition",yTickLabelPosition);

        QAxObject *rangeCells=sheet->querySubObject(QString("Range("+chartArea+")").toLatin1());
        chart->querySubObject("ChartArea")->setProperty("Top",rangeCells->property("Top"));
        chart->querySubObject("ChartArea")->setProperty("Left",rangeCells->property("Left"));
        chart->querySubObject("ChartArea")->setProperty("Width",rangeCells->property("Width"));
        chart->querySubObject("ChartArea")->setProperty("Height",rangeCells->property("Height"));
    }
    catch(...)
    {
        return false;
    }
    return true;
}

//清空指定Chart的数据
bool ExcelBase::clearChart(QAxObject *chart)
{
    try
    {
        while(chart->querySubObject("SeriesCollection(int)", 1))
        {
            chart->querySubObject("SeriesCollection(int)", 1)->dynamicCall("Delete(void)");
        }
//        chart->querySubObject("ChartArea")->dynamicCall("Clear(void)");
    }
    catch(...)
    {
        return false;
    }
    return true;

}

//快速读取指定sheet所有的内容
QVariant ExcelBase::readAll(QAxObject* sheet)
{
    QVariant var;
    if (sheet && !sheet->isNull())
    {
        QAxObject *usedRange = sheet->querySubObject("UsedRange");
        if(!usedRange || usedRange->isNull())
        {
            return var;
        }
        var = usedRange->dynamicCall("Value");
        delete usedRange;
    }
    return var;
}

//将readAll获取到的内容转为listlist
void ExcelBase::castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant>> &res)
{
    QVariantList varRows = var.toList();
    if(varRows.isEmpty())
    {
        return;
    }
    const int rowCount = varRows.size();
    QVariantList rowData;
    for(int i=0;i<rowCount;++i)
    {
        rowData = varRows[i].toList();
        res.push_back(rowData);
    }
}


int ExcelBase::letterToNumber(QString letter)
{
    int res=0;
    for(auto l:letter)
    {
        res=res*26+(l.unicode()-'A'+1);
    }
    return res;
}

























