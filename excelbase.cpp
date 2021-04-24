#include "excelbase.h"
#include "ui_excelbase.h"
#include <QMessageBox>
#include <QDir>
#include <QRegExp>
#include <QDebug>
#include <QFileInfo>
#include <qvector.h>

ExcelBase::ExcelBase(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::ExcelBase)
{
    ui->setupUi(this);
	openExcelFile("C:\\Users\\14641\\Desktop\\data.xlsx", true, false);
    QAxObject *workSheet=getSheet(1);

	QList<QList<QVariant>> ls;
	int cnt=0;
	for (int i = 0; i < 1000; ++i)
	{
		QList<QVariant> tmpLs;
		for (int j = 0; j < 2782; ++j)
		{
			tmpLs.append(cnt++);
		}
		ls.append(tmpLs);
	}
	setRange(workSheet, "A1:DBZ1000", ls);

    //addChart(workSheet,"$H$1:$J$47","","$A$3:$A$47","$B$3:$G$47","$B$1:$G$1",-20,40,10,10,-24,0,1,1,ExcelBase::xlTickLabelPositionHigh,ExcelBase::xlTickLabelPositionLow,ExcelBase::xlLegendPositionBottom);
}

ExcelBase::~ExcelBase()
{
	closeExcelFile();   //关闭excel文件，防止excel程序留在后台
	delete ui;
}

//打开excel文件
void ExcelBase::openExcelFile(const QString &filePath, bool isVisible, bool isDisplayAlerts)
{
	//初始化excel对象
	CoInitializeEx(nullptr, COINIT_MULTITHREADED);
	excel = new(std::nothrow) QAxObject();
	if (!excel)
	{
		QMessageBox::critical(this, QString::fromLocal8Bit("error"), QString::fromLocal8Bit("创建Excel对象失败"), QMessageBox::Ok, QMessageBox::Ok);
		delete excel;
		excel = nullptr;
		exit(1);
	}
	this->filePath = filePath;
	QFileInfo fileInfo(filePath);
	//连接excel，并打开文件，获取所有的sheet
	if (fileInfo.exists())
	{
		excel->setControl("Excel.Application");
		excel->dynamicCall("SetVisible(bool Visible)", isVisible);
		excel->setProperty("DisplayAlerts", isDisplayAlerts);
		workBook = excel->querySubObject("WorkBooks")->querySubObject("Open(QString&)", filePath);
		workSheets = workBook->querySubObject("WorkSheets");
	}
	else
	{
		QMessageBox::critical(this, QString::fromLocal8Bit("错误"), QString::fromLocal8Bit("打开文件失败,请检查路径下文件是否存在"), QMessageBox::Ok, QMessageBox::Ok);
		excel->dynamicCall("Quit(void)");
		delete excel;
		excel = nullptr;
		exit(1);
	}
}

//关闭指定的excel文件
void ExcelBase::closeExcelFile()
{
    if(!excel)
    {
		QMessageBox::critical(this, QString::fromLocal8Bit("错误"), QString::fromLocal8Bit("请先创建Excel对象"), QMessageBox::Ok, QMessageBox::Ok);
		exit(1);
    }
	//workBook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(filePath));
	workBook->dynamicCall("Save()");
	workBook->dynamicCall("Close()");
	workBook->dynamicCall("Quit()");
	excel->dynamicCall("Quit(void)");
	delete excel;
	excel = nullptr;
}

//获取sheet的数量
int ExcelBase::getSheetsCount()
{
    return workSheets->property("Count").toInt();
}

//添加一个sheet
QAxObject* ExcelBase::addSheet(const QString &sheetName)
{
    QAxObject *tmpWorkSheet=nullptr;
    int count=getSheetsCount();
    QAxObject *lastSheet=workSheets->querySubObject("Item(int)",count);
    tmpWorkSheet=workSheets->querySubObject("Add(QVariant)",lastSheet->asVariant());
    lastSheet->dynamicCall("Move(QVariant)",tmpWorkSheet->asVariant());
    tmpWorkSheet->setProperty("Name",sheetName);
    return tmpWorkSheet;
}

//删除指定的sheet
template<typename T> bool ExcelBase::deleteSheet(T sheetMark)
{
    if(std::is_same<T,int>::value)
    {
        workSheets->querySubObject("Item(int)",sheetMark)->dynamicCall("delete");
    }
    else if(std::is_same<T,QString>::value)
    {
        workSheets->querySubObject("Item(QString)",sheetMark)->dynamicCall("delete");
    }
    return true;
}

//获取指定的sheet
template<typename T> QAxObject* ExcelBase::getSheet(T sheetMark)
{
    QAxObject *tmpWorkSheet=nullptr;
    if(std::is_same<T,int>::value)
    {
        tmpWorkSheet=workSheets->querySubObject("Item(int)",sheetMark);
    }
    else if(std::is_same<T,QString>::value)
    {
        tmpWorkSheet=workSheets->querySubObject("Item(QString)",sheetMark);
    }
    if(!tmpWorkSheet)
    {
        QMessageBox::critical(this, QString::fromLocal8Bit("错误"), QString::fromLocal8Bit("获取指定Sheet失败"),QMessageBox::Ok,QMessageBox::Ok);
        exit(1);
    }
    return tmpWorkSheet;
}

//获取指定sheet的行
QAxObject* ExcelBase::getRows(QAxObject *sheet)
{
    QAxObject* rows=nullptr;
    rows=sheet->querySubObject("Rows");
    if(!rows)
    {
        QMessageBox::critical(this, QString::fromLocal8Bit("错误"), QString::fromLocal8Bit("获取行失败"),QMessageBox::Ok,QMessageBox::Ok);
        exit(1);
    }
    return rows;
}

//获取指定sheet的行数
int ExcelBase::getRowsCount(QAxObject *sheet)
{
    return getRows(sheet)->property("Count").toInt();
}

//获取指定sheet的列
QAxObject* ExcelBase::getColumns(QAxObject *sheet)
{
    QAxObject* columns=nullptr;
    columns=sheet->querySubObject("Columns");
    if(!columns)
    {
        QMessageBox::critical(this, QString::fromLocal8Bit("错误"), QString::fromLocal8Bit("获取列失败"),QMessageBox::Ok,QMessageBox::Ok);
        exit(1);
    }
    return columns;
}

//获取指定sheet的列数
int ExcelBase::getColumnsCount(QAxObject *sheet)
{
    return getRows(sheet)->property("Count").toInt();;
}

//获取指定单元格内容
QVariant ExcelBase::getCell(QAxObject* sheet,int row,int column)
{
    return sheet->querySubObject("Cells(int,int)",row,column)->property("Value");
}

//获取指定单元格内容
QVariant ExcelBase::getCell(QAxObject* sheet, QString cellAddress)
{
    return sheet->querySubObject("Range(QString)", cellAddress)->property("Value");
}

//设置指定单元格内容
void ExcelBase::setCell(QAxObject* sheet,int row,int column, const QString &value)
{
    sheet->querySubObject("Cells(int,int)",row,column)->setProperty("Value",value);
}

//设置指定单元格内容
void ExcelBase::setCell(QAxObject* sheet, const QString &cellAddress, const QString &value)
{
    sheet->querySubObject("Range(QString)", cellAddress)->setProperty("Value", value);
}

//获取指定单元格区域内容
QVariant ExcelBase::getRange(QAxObject* sheet, const QString &range)
{
	return sheet->querySubObject("Range(const QString&)", range)->property("Value");
}

//设置指定单元格区域内容
void ExcelBase::setRange(QAxObject* sheet, const QString &range, const QString &value)
{
	sheet->querySubObject("Range(const QString&)", range)->setProperty("Value", value);
}

//设置指定单元格区域内容
void ExcelBase::setRange(QAxObject* sheet, const QString &range, QList<QList<QVariant>> &value)
{
	long cnt = 0;
	for (auto v : value)
	{
		cnt += v.size();
	}
	if (sheet->querySubObject("Range(const QString&)", range)->property("Count") != cnt)
	{
		QMessageBox::critical(this, QString::fromLocal8Bit("错误"), QString::fromLocal8Bit("数据大小与单元格区域大小不相等"), QMessageBox::Ok, QMessageBox::Ok);
		exit(1);
	}
	sheet->querySubObject("Range(const QString&)", range)->setProperty("Value", castListListVariant2Variant(value));
}

//设置打印区域
void ExcelBase::setPrintArea(QAxObject *sheet, const QString &area)
{
    sheet->querySubObject("PageSetup")->setProperty("PrintArea",area);
}

//设置打印重复标题行
void ExcelBase::setPrintTitleRow(QAxObject *sheet, const QString &row)
{
    sheet->querySubObject("Pagesetup")->setProperty("PrintTitleRows",row);
}

//设置打印重复左侧列
void ExcelBase::setPrintTitleColumn(QAxObject *sheet, const QString &column)
{
    sheet->querySubObject("Pagesetup")->setProperty("PrintTitleColumns",column);
}

//设置页边距
void ExcelBase::setPrintMargin(QAxObject *sheet,
                               double topMargin, double rightMargin,double bottomMargin,
                               double leftMargin, double headerMargin, double footerMargin)
{
    sheet->querySubObject("Pagesetup")->setProperty("TopMargin",28.35*topMargin);
    sheet->querySubObject("Pagesetup")->setProperty("RightMargin",28.35*rightMargin);
    sheet->querySubObject("Pagesetup")->setProperty("BottomMargin",28.35*bottomMargin);
    sheet->querySubObject("Pagesetup")->setProperty("LeftMargin",28.35*leftMargin);
    sheet->querySubObject("Pagesetup")->setProperty("HeaderMargin",28.35*headerMargin);
    sheet->querySubObject("Pagesetup")->setProperty("FooterMargin",28.35*footerMargin);
}

//设置打印时是否居中
void ExcelBase::setPrintOrientation(QAxObject *sheet, bool isCenterHorizontally, bool isCenterVertically)
{
    sheet->querySubObject("Pagesetup")->setProperty("CenterHorizontally",isCenterHorizontally);
    sheet->querySubObject("Pagesetup")->setProperty("CenterVertically",isCenterVertically);
}

//设置页眉
void ExcelBase::setHeader(QAxObject *sheet, const QString &header, int position)
{
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

//设置页脚
void ExcelBase::setFooter(QAxObject *sheet, const QString &footer, int position)
{
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

//设置活动页面的布局
void ExcelBase::setWindowsView(QAxObject *excel,int viewMode)
{
    excel->querySubObject("Windows")->querySubObject("Item(int)",1)->setProperty("View",viewMode);
}

//插入Chart
//TODO 修改成适合多种图形的函数
void ExcelBase::addChart(QAxObject *sheet, const QString &chartArea, const QString &chartTitle, const QString &xValueDataSource, const QString &yValueDataSource, const QString &legendTitleSource,
                         double xMinScale,double xMaxScale,double xMajorUnit,double xMinorUnit,
                         double yMinScale,double yMaxScale,double yMajorUnit,double yMinorUnit,
                         int xTickLabelPosition,int yTickLabelPosition,int legendPosition)
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

//清空指定Chart的数据
void ExcelBase::clearChart(QAxObject *chart)
{
    while(chart->querySubObject("SeriesCollection(int)", 1))
    {
        chart->querySubObject("SeriesCollection(int)", 1)->dynamicCall("Delete(void)");
    }
    //chart->querySubObject("ChartArea")->dynamicCall("Clear(void)");

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

//将listlist转为variant，用于快速批量写入单元格
QVariant ExcelBase::castListListVariant2Variant(QList<QList<QVariant>> &value)
{
	QList<QVariant> valueLs;
	for (auto v : value)
	{
		valueLs.append(QVariant(v));
	}
	return QVariant(valueLs);
}

//将readAll获取到的内容转为listlist
QList<QList<QVariant>> ExcelBase::castVariant2ListListVariant(const QVariant &var)
{
	QList<QList<QVariant>> res;
    QVariantList varRows = var.toList();
    if(varRows.isEmpty())
    {
		QMessageBox::information(this, QString::fromLocal8Bit("提示"), QString::fromLocal8Bit("传入的QVariant没有数据"), QMessageBox::Ok, QMessageBox::Ok);
		exit(0);
    }
    QVariantList rowData;
    for(int i=0;i< varRows.size();++i)
    {
        rowData = varRows[i].toList();
        res.push_back(rowData);
    }
	return res;
}

//列号转数字
int ExcelBase::letterToNumber(const QString &letter)
{
    int res=0;
    for(auto l:letter)
    {
        res=res*26+(l.unicode()-'A'+1);
    }
    return res;
}

























