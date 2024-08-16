#ifndef OPENRGBINSTALLER_H
#define OPENRGBINSTALLER_H

#include <QMainWindow>
#include <QThread>
#include <QMessageBox>
#include <QIcon>
#include <QProgressBar>
#include <QLabel>
#include <QPushButton>
#include <QTranslator>
#include <QLocale>
#include <QProcess>
#include <QFile>
#include <QDir>
#include <QTemporaryFile>
#include <QNetworkAccessManager>
#include <QNetworkReply>

QT_BEGIN_NAMESPACE
namespace Ui {
class OpenRGBInstaller;
}
QT_END_NAMESPACE

class InstallWorker : public QThread {
    Q_OBJECT

public:
    InstallWorker(const QString &url, const QString &targetPath, QObject *parent = nullptr);
    void run() override;

signals:
    void progress(int percentage);
    void status(const QString &message);
    void finished(bool success, const QString &message);

private:
    QString url;
    QString targetPath;
    QString extractPath;

    void checkAndKillProcess(const QString &processName);
    void downloadFile(const QString &url, const QString &savePath);
    void extractZip(const QString &filePath);
    void moveExtractedFolder();
    void createShortcuts(const QString &executablePath);
};

class OpenRGBInstaller : public QMainWindow
{
    Q_OBJECT

public:
    OpenRGBInstaller(QWidget *parent = nullptr);
    ~OpenRGBInstaller();

private slots:
    void handleLaunchButtonClicked();
    void installProgram();
    void onInstallFinished(bool success, const QString &message);
    void showMessage(const QString &title, const QString &message);

private:
    Ui::OpenRGBInstaller *ui;
    QString url;
    QString targetPath;
    InstallWorker *installWorker;
};
#endif // OPENRGBINSTALLER_H
