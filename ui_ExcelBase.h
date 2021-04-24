/********************************************************************************
** Form generated from reading UI file 'ExcelBase.ui'
**
** Created by: Qt User Interface Compiler version 5.12.2
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_EXCELBASE_H
#define UI_EXCELBASE_H

#include <QtCore/QVariant>
#include <QtWidgets/QApplication>
#include <QtWidgets/QGridLayout>
#include <QtWidgets/QMainWindow>
#include <QtWidgets/QMenuBar>
#include <QtWidgets/QStatusBar>
#include <QtWidgets/QToolBar>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_ExcelBase
{
public:
    QWidget *centralWidget;
    QGridLayout *gridLayout;
    QMenuBar *menuBar;
    QToolBar *mainToolBar;
    QStatusBar *statusBar;

    void setupUi(QMainWindow *ExcelBase)
    {
        if (ExcelBase->objectName().isEmpty())
            ExcelBase->setObjectName(QString::fromUtf8("ExcelBase"));
        ExcelBase->resize(400, 300);
        centralWidget = new QWidget(ExcelBase);
        centralWidget->setObjectName(QString::fromUtf8("centralWidget"));
        gridLayout = new QGridLayout(centralWidget);
        gridLayout->setSpacing(6);
        gridLayout->setContentsMargins(11, 11, 11, 11);
        gridLayout->setObjectName(QString::fromUtf8("gridLayout"));
        ExcelBase->setCentralWidget(centralWidget);
        menuBar = new QMenuBar(ExcelBase);
        menuBar->setObjectName(QString::fromUtf8("menuBar"));
        menuBar->setGeometry(QRect(0, 0, 400, 23));
        ExcelBase->setMenuBar(menuBar);
        mainToolBar = new QToolBar(ExcelBase);
        mainToolBar->setObjectName(QString::fromUtf8("mainToolBar"));
        ExcelBase->addToolBar(Qt::TopToolBarArea, mainToolBar);
        statusBar = new QStatusBar(ExcelBase);
        statusBar->setObjectName(QString::fromUtf8("statusBar"));
        ExcelBase->setStatusBar(statusBar);

        retranslateUi(ExcelBase);

        QMetaObject::connectSlotsByName(ExcelBase);
    } // setupUi

    void retranslateUi(QMainWindow *ExcelBase)
    {
        ExcelBase->setWindowTitle(QApplication::translate("ExcelBase", "MainWindow", nullptr));
    } // retranslateUi

};

namespace Ui {
    class ExcelBase: public Ui_ExcelBase {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_EXCELBASE_H
