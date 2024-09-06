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


OpenRGBInstaller::OpenRGBInstaller(QWidget *parent) : QMainWindow(parent), ui(new Ui::OpenRGBInstaller) {
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
        qDebug() << "Unknown version selected. Unable to determine URL.";
        return;
    }

    qDebug() << "Selected version:" << selectedVersion;
    qDebug() << "Corresponding download URL:" << url;

    ui->statusProgressBar->setValue(0);
    ui->statusProgressBar->setVisible(true);
    ui->statusProgressBar->setFormat(tr("Starting install..."));
    qDebug() << "Initiating install worker thread.";
    installWorker->url = url;
    installWorker->start();
}


void OpenRGBInstaller::uninstallProgram() {
    ui->statusProgressBar->setFormat(tr("Uninstalling OpenRGB..."));
    ui->statusProgressBar->setValue(0);
    ui->installButton->setEnabled(false);
    ui->uninstallButton->setEnabled(false);
    ui->launchButton->setEnabled(false);

    try {
        qDebug() << "Checking if OpenRGB process is running...";
        installWorker->checkAndKillProcess("OpenRGB.exe");
        ui->statusProgressBar->setValue(25);
        qDebug() << "OpenRGB process closed.";

        QString installDir = targetPath + "/OpenRGB Windows 64-bit";
        QDir dir(installDir);
        if (dir.exists()) {
            qDebug() << "Removing OpenRGB installation directory:" << installDir;
            if (!dir.removeRecursively()) {
                throw std::runtime_error(tr("Failed to remove OpenRGB installation directory.").toStdString());
            }
            qDebug() << "Installation directory removed.";
        } else {
            qDebug() << "Installation directory does not exist.";
        }
        ui->statusProgressBar->setValue(50);

        QString startMenuShortcut = QStandardPaths::writableLocation(QStandardPaths::ApplicationsLocation) + "/OpenRGB.lnk";
        QString desktopShortcut = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation) + "/OpenRGB.lnk";

        QFile startMenuFile(startMenuShortcut);
        QFile desktopFile(desktopShortcut);

        if (startMenuFile.exists()) {
            qDebug() << "Removing Start Menu shortcut:" << startMenuShortcut;
            if (!startMenuFile.remove()) {
                throw std::runtime_error(tr("Failed to remove Start Menu shortcut.").toStdString());
            }
            qDebug() << "Start Menu shortcut removed.";
        } else {
            qDebug() << "Start Menu shortcut does not exist.";
        }

        if (desktopFile.exists()) {
            qDebug() << "Removing Desktop shortcut:" << desktopShortcut;
            if (!desktopFile.remove()) {
                throw std::runtime_error(tr("Failed to remove Desktop shortcut.").toStdString());
            }
            qDebug() << "Desktop shortcut removed.";
        } else {
            qDebug() << "Desktop shortcut does not exist.";
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
        qDebug() << "Uninstall error:" << e.what();
        ui->statusProgressBar->setFormat(tr("Uninstall failed."));
        showMessage(tr("Error"), tr("An error occurred during uninstallation: %1").arg(e.what()));
    }
    ui->statusProgressBar->setValue(0);
}

void OpenRGBInstaller::onInstallFinished(bool success, const QString &message) {
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

void OpenRGBInstaller::showMessage(const QString &title, const QString &message) {
    qDebug() << "Showing message box. Title:" << title << "Message:" << message;
    QMessageBox msgBox(this);
    msgBox.setWindowTitle(title);
    msgBox.setText(message);
    msgBox.setIcon(title == tr("Success") ? QMessageBox::Information : QMessageBox::Critical);
    msgBox.exec();
}

bool OpenRGBInstaller::isOpenRGBInstalled() {
    QString installDir = targetPath + "/OpenRGB Windows 64-bit";

    QDir dir(installDir);
    if (dir.exists()) {
        qDebug() << "OpenRGB is installed. Directory exists at:" << installDir;
        return true;
    } else {
        qDebug() << "OpenRGB is not installed. Directory does not exist.";
        return false;
    }
}

void OpenRGBInstaller::populateComboBox() {
    ui->versionComboBox->addItem("Master");
    ui->versionComboBox->addItem("0.9");
    ui->versionComboBox->addItem("0.8");
    ui->versionComboBox->addItem("0.7");
    ui->versionComboBox->addItem("0.6");
}

QString OpenRGBInstaller::checkCurrentlyInstalledVersionFile() {
    QFile file(currentlyInstalledVersionFilePath);
    qDebug() << currentlyInstalledVersionFilePath;
    if (!file.exists()) {
        qDebug() << "here";
        return "N/A";
    }

    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        qDebug() << "there";

        return "N/A";
    }

    QString fileContent = file.readAll();

    file.close();
    return fileContent;
}


void OpenRGBInstaller::createCurrentlyInstalledVersionFile() {
    QString appDataLocation = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
    QDir dir(appDataLocation);

    if (!dir.exists()) {
        if (!dir.mkpath(appDataLocation)) {
            qDebug() << "Failed to create directory:" << appDataLocation;
            return;
        }
    }

    QFile file(currentlyInstalledVersionFilePath);

    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        qDebug() << "Failed to open file for writing";
        return;
    }

    QTextStream out(&file);
    out << ui->versionComboBox->currentText();

    file.close();
}

void OpenRGBInstaller::removeCurrentlyInstalledVersionFile() {
    QString appDataLocation = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
    QDir dir(appDataLocation);
    QFile file(currentlyInstalledVersionFilePath);

    if (file.exists()) {
        if (file.remove()) {
            dir.removeRecursively();
            qDebug() << "File deleted successfully!";
        } else {
            qDebug() << "Failed to delete the file.";
        }
    } else {
        qDebug() << "File does not exist.";
    }
}
