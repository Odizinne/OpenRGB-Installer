#include "openrgbinstaller.h"
#include "ui_openrgbinstaller.h"
#include <QNetworkRequest>
#include <QNetworkReply>
#include <QProcess>
#include <QStandardPaths>
#include <QFile>
#include <QDir>
#include <QFileInfo>
#include <QDebug>
#include <QTextStream>
#include <windows.h>
#include <shlobj.h>
#include <shellapi.h>

InstallWorker::InstallWorker(const QString &url, const QString &targetPath, QObject *parent)
    : QThread(parent)
    , url(url)
    , targetPath(targetPath)
{

}

void InstallWorker::run()
{
    try {
        checkAndKillProcess("OpenRGB.exe");
        QString zipFilePath = QStandardPaths::writableLocation(QStandardPaths::TempLocation) + "/openrgb.zip";
        emit progress(0);
        downloadFile(url, zipFilePath);
        emit progress(25);
        extractZip(zipFilePath);
        emit progress(50);
        moveExtractedFolder();
        emit progress(75);


        createShortcuts(targetPath + "/OpenRGB Windows 64-bit");
        emit progress(100);
        emit finished(true, tr("OpenRGB installed successfully! Created start menu and desktop shortcuts."));
    } catch (const std::exception &e) {
        emit finished(false, tr("An error occurred: %1").arg(e.what()));
    }
}

void InstallWorker::checkAndKillProcess(const QString &processName)
{
    QProcess process;
    process.start("tasklist", QStringList() << "/FI" << QString("IMAGENAME eq %1").arg(processName));
    process.waitForFinished();

    QString output = process.readAll();

    if (output.contains(processName)) {
        process.start("taskkill", QStringList() << "/F" << "/IM" << processName);
        process.waitForFinished();

        if (!(process.exitStatus() == QProcess::NormalExit)) {
            throw std::runtime_error(tr("Failed to terminate the process: %1").arg(processName).toStdString());
        }
    }
}

void InstallWorker::downloadFile(const QString &url, const QString &savePath)
{
    emit status(tr("Downloading..."));

    QNetworkAccessManager manager;
    QNetworkRequest request(url);
    QNetworkReply *reply = manager.get(request);

    QEventLoop loop;
    connect(reply, &QNetworkReply::finished, &loop, &QEventLoop::quit);
    loop.exec();

    if (reply->error() != QNetworkReply::NoError) {
        throw std::runtime_error(reply->errorString().toStdString());
    }

    QFile file(savePath);
    if (!file.open(QIODevice::WriteOnly)) {
        throw std::runtime_error(tr("Failed to open file for writing.").toStdString());
    }

    file.write(reply->readAll());
    file.close();
}

void InstallWorker::extractZip(const QString &zipFilePath)
{
    emit status(tr("Extracting files..."));

    QString extractPath = QStandardPaths::writableLocation(QStandardPaths::TempLocation) + "/OpenRGBExtracted";
    if (QDir(extractPath).exists()) {
        if (!QDir(extractPath).removeRecursively()) {
            throw std::runtime_error(tr("Failed to remove existing extraction directory.").toStdString());
        }
    }

    QProcess process;
    QString command = "powershell";
    QStringList arguments;
    arguments << "-Command" << QString("Expand-Archive -Path '%1' -DestinationPath '%2'").arg(zipFilePath, extractPath);
    process.start(command, arguments);

    if (!process.waitForFinished()) {
        throw std::runtime_error(tr("An error occurred while extracting the ZIP file.").toStdString());
    }

    QDir extractedDir(extractPath);
    QStringList entries = extractedDir.entryList(QDir::AllEntries | QDir::NoDotAndDotDot);

    if (entries.isEmpty()) {
        throw std::runtime_error(tr("No files or folders found in the extracted contents.").toStdString());
    }
}

void InstallWorker::moveExtractedFolder()
{
    QString extractPath = QStandardPaths::writableLocation(QStandardPaths::TempLocation) + "/OpenRGBExtracted";
    QString extractedFolderName = "OpenRGB Windows 64-bit";
    QString sourceDir = extractPath + "/" + extractedFolderName;
    QString targetDir = targetPath + "/" + extractedFolderName;

    if (QDir(targetDir).exists()) {
        if (!QDir(targetDir).removeRecursively()) {
            throw std::runtime_error(tr("Failed to remove existing target directory.").toStdString());
        }
    }

    if (!QDir(targetDir).mkpath(targetPath)) {
        throw std::runtime_error(tr("Failed to create target directory.").toStdString());
    }

    if (!QDir(targetDir).rename(sourceDir, targetDir)) {
        throw std::runtime_error(tr("Failed to move extracted folder to target path.").toStdString());
    }
}

void InstallWorker::createShortcuts(const QString &executablePath)
{
    QString startMenuPath = QStandardPaths::writableLocation(QStandardPaths::ApplicationsLocation) + "/OpenRGB.lnk";
    QString desktopPath = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation) + "/OpenRGB.lnk";

    auto createShortcut = [](const QString &path, const QString &target, const QString &iconPath, const QString &startIn) {
        HRESULT hResult;
        IShellLink *pShellLink = nullptr;
        IPersistFile *pPersistFile = nullptr;

        hResult = CoInitialize(nullptr);
        if (FAILED(hResult)) {
            return;
        }

        hResult = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_IShellLink, (void **)&pShellLink);
        if (SUCCEEDED(hResult)) {
            pShellLink->SetPath(target.toStdWString().c_str());
            pShellLink->SetIconLocation(iconPath.toStdWString().c_str(), 0);
            pShellLink->SetWorkingDirectory(startIn.toStdWString().c_str());

            hResult = pShellLink->QueryInterface(IID_IPersistFile, (void **)&pPersistFile);
            if (SUCCEEDED(hResult)) {
                hResult = pPersistFile->Save(path.toStdWString().c_str(), TRUE);
                pPersistFile->Release();
            }
            pShellLink->Release();
        }
        CoUninitialize();
    };

    createShortcut(startMenuPath, executablePath + "/OpenRGB.exe", executablePath + "/OpenRGB.exe", targetPath + "/OpenRGB Windows 64-bit");
    createShortcut(desktopPath, executablePath + "/OpenRGB.exe", executablePath + "/OpenRGB.exe", targetPath + "/OpenRGB Windows 64-bit");
}


OpenRGBInstaller::OpenRGBInstaller(QWidget *parent) : QMainWindow(parent), ui(new Ui::OpenRGBInstaller)
{
    ui->setupUi(this);
    setWindowIcon(QIcon(":/icons/openrgb-installer.png"));
    setFixedSize(this->size());
    populateComboBox();

    url = "https://gitlab.com/CalcProgrammer1/OpenRGB/-/jobs/artifacts/master/download?job=Windows%2064";
    targetPath = QStandardPaths::writableLocation(QStandardPaths::GenericDataLocation) + "/Programs";
    currentlyInstalledVersionFilePath = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation) + "/installed_version.txt";

    installWorker = new InstallWorker(url, targetPath, this);
    connect(installWorker, &InstallWorker::progress, ui->statusProgressBar, &QProgressBar::setValue);
    connect(installWorker, &InstallWorker::status, ui->statusProgressBar, &QProgressBar::setFormat);
    connect(installWorker, &InstallWorker::finished, this, &OpenRGBInstaller::onInstallFinished);
    connect(ui->installButton, &QPushButton::pressed, this, &OpenRGBInstaller::installProgram);
    connect(ui->uninstallButton, &QPushButton::pressed, this, &OpenRGBInstaller::uninstallProgram);
    connect(ui->launchButton, &QPushButton::clicked, this, &OpenRGBInstaller::handleLaunchButtonClicked);

    ui->installedVersionLabel->setText(checkCurrentlyInstalledVersionFile());
    ui->uninstallButton->setEnabled(isOpenRGBInstalled());
    ui->launchButton->setEnabled(isOpenRGBInstalled());
    if (isOpenRGBInstalled())
        ui->installButton->setText("Reinstall");
    ui->statusProgressBar->setFormat("Ready");
}

OpenRGBInstaller::~OpenRGBInstaller()
{
    delete ui;
}

void OpenRGBInstaller::handleLaunchButtonClicked()
{
    QString executablePath = targetPath + "/OpenRGB Windows 64-bit/OpenRGB.exe";
    if (QFile::exists(executablePath)) {
        QProcess::startDetached(executablePath);
        close();
    } else {
        showMessage(tr("Error"), tr("OpenRGB executable not found."));
    }
}

void OpenRGBInstaller::installProgram()
{
    ui->installButton->setEnabled(false);
    ui->uninstallButton->setEnabled(false);
    ui->launchButton->setEnabled(false);
    ui->versionComboBox->setEnabled(false);
    QString selectedVersion = ui->versionComboBox->currentText();

    if (selectedVersion == "Master") {
        url = "https://gitlab.com/CalcProgrammer1/OpenRGB/-/jobs/artifacts/master/download?job=Windows%2064";
    } else if (selectedVersion == "0.9") {
        url = "https://openrgb.org/releases/release_0.9/OpenRGB_0.9_Windows_64_b5f46e3.zip";
    } else if (selectedVersion == "0.8") {
        url = "https://openrgb.org/releases/release_0.8/OpenRGB_0.8_Windows_64_fb88964.zip";
    } else if (selectedVersion == "0.7") {
        url = "https://openrgb.org/releases/release_0.7/OpenRGB_0.7_Windows_64_6128731.zip";
    } else if (selectedVersion == "0.6") {
        url = "https://openrgb.org/releases/release_0.6/OpenRGB_0.6_Windows_64_405ff7f.zip";
    } else {
        return;
    }

    ui->statusProgressBar->setValue(0);
    ui->statusProgressBar->setVisible(true);
    ui->statusProgressBar->setFormat(tr("Starting install..."));
    installWorker->url = url;
    installWorker->start();
}


void OpenRGBInstaller::uninstallProgram()
{
    ui->statusProgressBar->setFormat(tr("Uninstalling OpenRGB..."));
    ui->statusProgressBar->setValue(0);
    ui->installButton->setEnabled(false);
    ui->uninstallButton->setEnabled(false);
    ui->launchButton->setEnabled(false);

    try {
        installWorker->checkAndKillProcess("OpenRGB.exe");
        ui->statusProgressBar->setValue(25);

        QString installDir = targetPath + "/OpenRGB Windows 64-bit";
        QDir dir(installDir);
        if (dir.exists()) {
            if (!dir.removeRecursively()) {
                throw std::runtime_error(tr("Failed to remove OpenRGB installation directory.").toStdString());
            }
        }
        ui->statusProgressBar->setValue(50);

        QString startMenuShortcut = QStandardPaths::writableLocation(QStandardPaths::ApplicationsLocation) + "/OpenRGB.lnk";
        QString desktopShortcut = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation) + "/OpenRGB.lnk";

        QFile startMenuFile(startMenuShortcut);
        QFile desktopFile(desktopShortcut);

        if (startMenuFile.exists()) {
            if (!startMenuFile.remove()) {
                throw std::runtime_error(tr("Failed to remove Start Menu shortcut.").toStdString());
            }
        }

        if (desktopFile.exists()) {
            if (!desktopFile.remove()) {
                throw std::runtime_error(tr("Failed to remove Desktop shortcut.").toStdString());
            }
        }

        ui->statusProgressBar->setValue(75);

        ui->statusProgressBar->setValue(100);
        ui->statusProgressBar->setFormat(tr("Uninstall complete."));
        showMessage(tr("Success"), tr("OpenRGB uninstalled successfully."));
        removeCurrentlyInstalledVersionFile();
        ui->installedVersionLabel->setText(checkCurrentlyInstalledVersionFile());
        ui->installButton->setEnabled(true);
        ui->uninstallButton->setEnabled(false);
        ui->launchButton->setEnabled(false);
        ui->installButton->setText("Install");
    } catch (const std::exception &e) {
        ui->statusProgressBar->setFormat(tr("Uninstall failed."));
        showMessage(tr("Error"), tr("An error occurred during uninstallation: %1").arg(e.what()));
    }
    ui->statusProgressBar->setValue(0);
}

void OpenRGBInstaller::onInstallFinished(bool success, const QString &message)
{
    if (success) {
        ui->launchButton->setEnabled(true);
        ui->uninstallButton->setEnabled(true);
        ui->installButton->setEnabled(true);
        ui->versionComboBox->setEnabled(true);
        ui->installButton->setText("Reinstall");
        ui->statusProgressBar->setFormat(tr("Install finished."));
        showMessage(tr("Success"), message);
        createCurrentlyInstalledVersionFile();
        ui->installedVersionLabel->setText(checkCurrentlyInstalledVersionFile());
        ui->statusProgressBar->setValue(0);
    } else {
        ui->statusProgressBar->setFormat(tr("Install failed."));
        showMessage(tr("Error"), message);
    }
}

void OpenRGBInstaller::showMessage(const QString &title, const QString &message)
{
    QMessageBox msgBox(this);
    msgBox.setWindowTitle(title);
    msgBox.setText(message);
    msgBox.setIcon(title == tr("Success") ? QMessageBox::Information : QMessageBox::Critical);
    msgBox.exec();
}

bool OpenRGBInstaller::isOpenRGBInstalled()
{
    QString installDir = targetPath + "/OpenRGB Windows 64-bit";
    QDir dir(installDir);

    return dir.exists();
}

void OpenRGBInstaller::populateComboBox()
{
    ui->versionComboBox->addItem("Master");
    ui->versionComboBox->addItem("0.9");
    ui->versionComboBox->addItem("0.8");
    ui->versionComboBox->addItem("0.7");
    ui->versionComboBox->addItem("0.6");
}

QString OpenRGBInstaller::checkCurrentlyInstalledVersionFile()
{
    QFile file(currentlyInstalledVersionFilePath);

    if (!file.exists()) {
        return "N/A";
    }

    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        return "N/A";
    }

    QString fileContent = file.readAll();

    file.close();

    return fileContent;
}


void OpenRGBInstaller::createCurrentlyInstalledVersionFile()
{
    QString appDataLocation = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
    QDir dir(appDataLocation);

    if (!dir.exists()) {
        if (!dir.mkpath(appDataLocation)) {
            return;
        }
    }

    QFile file(currentlyInstalledVersionFilePath);

    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        return;
    }

    QTextStream out(&file);
    out << ui->versionComboBox->currentText();

    file.close();
}

void OpenRGBInstaller::removeCurrentlyInstalledVersionFile()
{
    QString appDataLocation = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
    QDir dir(appDataLocation);
    QFile file(currentlyInstalledVersionFilePath);

    if (file.exists()) {
        if (file.remove()) {
            dir.removeRecursively();
        }
    }
}
