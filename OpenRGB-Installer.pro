QT       += core gui network

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17

# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

INCLUDEPATH += \
    src/OpenRGBInstaller

SOURCES += \
    src/main.cpp \
    src/OpenRGBInstaller/openrgbinstaller.cpp

HEADERS += \
    src/OpenRGBInstaller/openrgbinstaller.h

FORMS += \
    src/OpenRGBInstaller/openrgbinstaller.ui

RESOURCES += \
    src/Resources/resources.qrc \

RC_FILE = src/Resources/appicon.rc

LIBS += -lole32 -lshell32

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target
