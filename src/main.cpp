#include "openrgbinstaller.h"

#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    a.setStyle("fusion");
    OpenRGBInstaller w;
    w.show();
    return a.exec();
}
