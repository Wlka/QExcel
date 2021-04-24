#include "excelbase.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    ExcelBase w;
    w.show();

    return a.exec();
}
