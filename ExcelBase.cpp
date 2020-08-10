#include "excelbase.h"
#include "ui_excelbase.h"
#include <QMessageBox>
#include <QDir>
#include <QDebug>

ExcelBase::ExcelBase(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::ExcelBase)
{
    ui->setupUi(this);

    qDebug()<<openExcelFile("C:/Users/Wlka/Desktop/data.xlsx");
    QAxObject *workSheet=getSheet(1);
//    getSheet(2)->dynamicCall("Select(void)");
    qDebug()<<setFooter(workSheet,"dfgdhog",ExcelBase::Right);

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
        excel->dynamicCall("SetVisible(bool Visible)","false");
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
        workBook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filePath));
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
QString ExcelBase::getCell(QAxObject* sheet,int row,int column)
{
    QString strCell="";
    try
    {
        strCell=sheet->querySubObject("Cells(int,int)",row,column)->property("Value").toString();
    }
    catch (...)
    {
        QMessageBox::critical(this,"错误","获取单元格信息失败",QMessageBox::Ok,QMessageBox::Ok);
    }
    return strCell;
}

//获取指定单元格内容
QString ExcelBase::getCell(QAxObject* sheet, QString &number)
{
    QString strCell="";
    try
    {
        strCell=sheet->querySubObject("Range(QString)",number)->property("Value").toString();
    }
    catch (...)
    {
        QMessageBox::critical(this,"错误","获取单元格信息失败",QMessageBox::Ok,QMessageBox::Ok);
    }
    return strCell;
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


























