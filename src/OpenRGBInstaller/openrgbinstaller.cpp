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
    : QThread(parent), url(url), targetPath(targetPath) {}

void InstallWorker::run() {
    try {
        checkAndKillProcess("OpenRGB.exe");
        QString zipFilePath = QStandardPaths::writableLocation(QStandardPaths::TempLocation) + "/openrgb.zip";
        emit progress(0);
        qDebug() << "Downloading OpenRGB ZIP file...";
        downloadFile(url, zipFilePath);
        qDebug() << "Downloaded ZIP file to:" << zipFilePath;
        emit progress(25);
        qDebug() << "Extracting ZIP file...";
        extractZip(zipFilePath);
        emit progress(50);
        qDebug() << "Moving extracted folder to target path...";
        moveExtractedFolder();
        emit progress(75);


        qDebug() << "Creating shortcuts...";
        createShortcuts(targetPath + "/OpenRGB Windows 64-bit");
        emit progress(100);
        emit finished(true, tr("OpenRGB installed successfully! Created start menu and desktop shortcuts."));
    } catch (const std::exception &e) {
        qDebug() << "Exception caught:" << e.what();
        emit finished(false, tr("An error occurred: %1").arg(e.what()));
    }
}

void InstallWorker::checkAndKillProcess(const QString &processName) {
    // Create a QProcess to execute the tasklist command
    QProcess process;
    process.start("tasklist", QStringList() << "/FI" << QString("IMAGENAME eq %1").arg(processName));
    process.waitForFinished();

    QString output = process.readAll();

    if (output.contains(processName)) {
        qDebug() << processName << "is running. Attempting to kill it...";

        // Execute the taskkill command to terminate the process
        process.start("taskkill", QStringList() << "/F" << "/IM" << processName);
        process.waitForFinished();

        if (process.exitStatus() == QProcess::NormalExit) {
            qDebug() << processName << "terminated successfully.";
        } else {
            qDebug() << "Failed to terminate" << processName;
            throw std::runtime_error(tr("Failed to terminate the process: %1").arg(processName).toStdString());
        }
    } else {
        qDebug() << processName << "is not running.";
    }
}

void InstallWorker::downloadFile(const QString &url, const QString &savePath) {
    emit status(tr("Downloading..."));

    QNetworkAccessManager manager;
    QNetworkRequest request(url);
    QNetworkReply *reply = manager.get(request);

    QEventLoop loop;
    connect(reply, &QNetworkReply::finished, &loop, &QEventLoop::quit);
    loop.exec();

    if (reply->error() != QNetworkReply::NoError) {
        qDebug() << "Download error:" << reply->errorString();
        throw std::runtime_error(reply->errorString().toStdString());
    }

    QFile file(savePath);
    if (!file.open(QIODevice::WriteOnly)) {
        qDebug() << "Failed to open file for writing:" << savePath;
        throw std::runtime_error(tr("Failed to open file for writing.").toStdString());
    }

    file.write(reply->readAll());
    file.close();
    qDebug() << "File downloaded and saved to:" << savePath;
}

void InstallWorker::extractZip(const QString &zipFilePath) {
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
        qDebug() << "Extracted folder is empty, throwing error.";
        throw std::runtime_error(tr("No files or folders found in the extracted contents.").toStdString());
    } else {
        qDebug() << "Extraction successful. Files and folders found:";
        for (const QString &entry : entries) {
            qDebug() << "  - " << entry;
        }
    }

    qDebug() << "Extraction completed successfully.";
}

void InstallWorker::moveExtractedFolder() {
    QString extractPath = QStandardPaths::writableLocation(QStandardPaths::TempLocation) + "/OpenRGBExtracted";
    QString extractedFolderName = "OpenRGB Windows 64-bit";
    QString sourceDir = extractPath + "/" + extractedFolderName;
    QString targetDir = targetPath + "/" + extractedFolderName;

    if (QDir(targetDir).exists()) {
        qDebug() << "Target directory already exists:" << targetDir;

        qDebug() << "Attempting to remove the existing directory...";
        if (!QDir(targetDir).removeRecursively()) {
            qDebug() << "Failed to remove existing target directory:" << targetDir;
            throw std::runtime_error(tr("Failed to remove existing target directory.").toStdString());
        }
        qDebug() << "Existing target directory removed successfully.";
    }

    if (!QDir(targetDir).mkpath(targetPath)) {
        qDebug() << "Failed to create target directory:" << targetPath;
        throw std::runtime_error(tr("Failed to create target directory.").toStdString());
    }

    if (!QDir(targetDir).rename(sourceDir, targetDir)) {
        qDebug() << "Failed to move extracted folder from:" << sourceDir << "to:" << targetDir;
        throw std::runtime_error(tr("Failed to move extracted folder to target path.").toStdString());
    }

    qDebug() << "Moved extracted folder to:" << targetDir;
}

void InstallWorker::createShortcuts(const QString &executablePath) {
    QString startMenuPath = QStandardPaths::writableLocation(QStandardPaths::ApplicationsLocation) + "/OpenRGB.lnk";
    QString desktopPath = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation) + "/OpenRGB.lnk";

    // Function to create a shortcut
    auto createShortcut = [](const QString &path, const QString &target, const QString &iconPath, const QString &startIn) {
        HRESULT hResult;
        IShellLink *pShellLink = nullptr;
        IPersistFile *pPersistFile = nullptr;

        hResult = CoInitialize(nullptr);
        if (FAILED(hResult)) {
            qDebug() << "Failed to initialize COM library";
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

    // Create Start Menu shortcut
    createShortcut(startMenuPath, executablePath + "/OpenRGB.exe", executablePath + "/OpenRGB.exe", targetPath + "/OpenRGB Windows 64-bit");

    // Create Desktop shortcut
    createShortcut(desktopPath, executablePath + "/OpenRGB.exe", executablePath + "/OpenRGB.exe", targetPath + "/OpenRGB Windows 64-bit");

    qDebug() << "Shortcuts created. Start Menu:" << startMenuPath << "Desktop:" << desktopPath;
}


OpenRGBInstaller::OpenRGBInstaller(QWidget *parent) : QMainWindow(parent), ui(new Ui::OpenRGBInstaller) {
    ui->setupUi(this);
    setWindowIcon(QIcon(":/icons/openrgb-installer.png"));
    setFixedSize(this->size());

    url = "https://gitlab.com/CalcProgrammer1/OpenRGB/-/jobs/artifacts/master/download?job=Windows%2064";
    targetPath = QStandardPaths::writableLocation(QStandardPaths::GenericDataLocation) + "/Programs";

    installWorker = new InstallWorker(url, targetPath, this);
    connect(installWorker, &InstallWorker::progress, ui->statusProgressBar, &QProgressBar::setValue);
    connect(installWorker, &InstallWorker::status, ui->statusLabel, &QLabel::setText);
    connect(installWorker, &InstallWorker::finished, this, &OpenRGBInstaller::onInstallFinished);

    ui->launchButton->setEnabled(false);
    connect(ui->launchButton, &QPushButton::clicked, this, &OpenRGBInstaller::handleLaunchButtonClicked);

    qDebug() << "Starting installation process...";
    installProgram();
}

OpenRGBInstaller::~OpenRGBInstaller() {
    delete ui;
}

void OpenRGBInstaller::handleLaunchButtonClicked() {
    QString executablePath = targetPath + "/OpenRGB Windows 64-bit/OpenRGB.exe";
    qDebug() << "Checking if executable exists at:" << executablePath;
    if (QFile::exists(executablePath)) {
        qDebug() << "Executable found, starting...";
        QProcess::startDetached(executablePath);
        close();
    } else {
        qDebug() << "Executable not found.";
        showMessage(tr("Error"), tr("OpenRGB executable not found."));
    }
}

void OpenRGBInstaller::installProgram() {
    // Kill OpenRGB processes if needed (not implemented in this C++ code)
    ui->statusProgressBar->setValue(0);
    ui->statusLabel->setText(tr("Starting install..."));
    qDebug() << "Initiating install worker thread.";
    installWorker->start();
}

void OpenRGBInstaller::onInstallFinished(bool success, const QString &message) {
    if (success) {
        ui->launchButton->setEnabled(true);
        ui->statusLabel->setText(tr("Install finished."));
        showMessage(tr("Success"), message);
    } else {
        ui->statusLabel->setText(tr("Install failed."));
        showMessage(tr("Error"), message);
    }
}

void OpenRGBInstaller::showMessage(const QString &title, const QString &message) {
    qDebug() << "Showing message box. Title:" << title << "Message:" << message;
    QMessageBox msgBox(this);
    msgBox.setWindowTitle(title);
    msgBox.setText(message);
    msgBox.setIcon(title == tr("Success") ? QMessageBox::Information : QMessageBox::Critical);
    msgBox.exec();
}
