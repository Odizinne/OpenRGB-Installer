import os
import sys
import shutil
import subprocess
import zipfile
import tempfile
import requests
import psutil
import winshell
from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QIcon
from design import Ui_MainWindow


def kill_openrgb():
    for process in psutil.process_iter(attrs=["pid", "name"]):
        try:
            if process.info["name"].lower() == "openrgb.exe":
                process.terminate()
                process.wait()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass


class InstallWorker(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, url, target_path):
        super().__init__()
        self.url = url
        self.target_path = target_path

    def run(self):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as temp_file:
                self.download_file(self.url, temp_file.name)
                self.extract_zip(temp_file.name)
                self.move_folder()
                self.create_shortcuts(os.path.join(self.target_path, "OpenRGB Windows 64-bit"))

            self.finished.emit(True, "OpenRGB installed successfully!\nCreated start menu and desktop shortcuts.")
        except Exception as e:
            self.finished.emit(False, str(e))

    def download_file(self, url, save_path):
        try:
            self.status.emit("Downloading...")

            response = requests.get(url, stream=True)
            response.raise_for_status()

            total_size = int(response.headers.get("content-length", 0))
            downloaded_size = 0

            with open(save_path, "wb") as file:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        file.write(chunk)
                        downloaded_size += len(chunk)

                        progress = int((downloaded_size / total_size) * 100)
                        self.progress.emit(progress)
        except requests.exceptions.HTTPError as http_err:
            raise Exception(f"HTTP error occurred: {http_err}")
        except requests.exceptions.ConnectionError:
            raise Exception("A connection error occurred. Cannot reach the server.")
        except requests.exceptions.Timeout:
            raise Exception("The request timed out.")
        except requests.exceptions.RequestException as err:
            raise Exception(f"An error occurred while downloading the file: {err}")
        except Exception as err:
            raise Exception(f"An unexpected error occurred: {err}")

    def extract_zip(self, file_path):
        try:
            self.status.emit("Extracting files...")

            self.extract_path = tempfile.mkdtemp()
            with zipfile.ZipFile(file_path, "r") as zip_ref:
                zip_ref.extractall(self.extract_path)
        except zipfile.BadZipFile:
            raise Exception("Error: The downloaded file is not a valid zip file.")
        except Exception as err:
            raise Exception(f"An error occurred while extracting the zip file: {err}")

    def move_folder(self):
        try:
            self.status.emit("Installing to target path...")

            extracted_folders = [
                d for d in os.listdir(self.extract_path) if os.path.isdir(os.path.join(self.extract_path, d))
            ]
            if not extracted_folders:
                raise Exception("No folder found in the extracted contents.")

            extracted_folder_name = extracted_folders[0]

            source_path = os.path.join(self.extract_path, extracted_folder_name)
            destination_path = os.path.join(self.target_path, extracted_folder_name)

            if os.path.exists(destination_path):
                shutil.rmtree(destination_path)

            shutil.move(source_path, destination_path)
        except Exception as err:
            raise Exception(f"An error occurred while moving the folder: {err}")

    def create_shortcuts(self, executable_path):
        start_menu_folder = winshell.programs()
        desktop_folder = winshell.desktop()

        start_menu_shortcut_path = os.path.join(start_menu_folder, "OpenRGB.lnk")
        with winshell.shortcut(start_menu_shortcut_path) as start_menu_shortcut:
            start_menu_shortcut.path = os.path.join(executable_path, "OpenRGB.exe")
            start_menu_shortcut.description = "OpenRGB"
            start_menu_shortcut.working_directory = executable_path

        desktop_shortcut_path = os.path.join(desktop_folder, "OpenRGB.lnk")
        with winshell.shortcut(desktop_shortcut_path) as desktop_shortcut:
            desktop_shortcut.path = os.path.join(executable_path, "OpenRGB.exe")
            desktop_shortcut.description = "OpenRGB"
            desktop_shortcut.working_directory = executable_path


class OpenRGBInstaller(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle("OpenRGB installer")
        self.setWindowIcon(QIcon("icons/openrgb-installer.png"))
        self.setFixedSize(self.size())
        self.url = "https://gitlab.com/CalcProgrammer1/OpenRGB/-/jobs/artifacts/master/download?job=Windows%2064"
        self.target_path = os.path.join(os.getenv("LOCALAPPDATA"), "Programs")
        self.install_worker = InstallWorker(self.url, self.target_path)
        self.install_worker.progress.connect(self.ui.status_progress_bar.setValue)
        self.install_worker.status.connect(self.ui.status_label.setText)
        self.install_worker.finished.connect(self.on_install_finished)
        self.ui.launch_button.setEnabled(False)
        self.ui.launch_button.clicked.connect(self.handle_launch_button_clicked)
        self.install_program()

    def handle_launch_button_clicked(self):
        try:
            executable_path = os.path.join(self.target_path, "OpenRGB Windows 64-bit", "OpenRGB.exe")
            if os.path.exists(executable_path):
                subprocess.Popen(executable_path, creationflags=subprocess.CREATE_NEW_CONSOLE)
                self.close()
            else:
                raise FileNotFoundError("OpenRGB executable not found.")
        except Exception as e:
            self.show_message("Error", f"Failed to launch OpenRGB: {e}")

    def install_program(self):
        kill_openrgb()
        self.ui.status_progress_bar.setValue(0)
        self.ui.status_label.setText("Starting install...")
        self.install_worker.start()

    def on_install_finished(self, success, message):
        if success:
            self.ui.launch_button.setEnabled(True)
            self.ui.status_label.setText("install finished.")
            self.show_message("Success", message)
        else:
            self.ui.status_label.setText("install failed.")
            self.show_message("Error", message)

    def show_message(self, title, message):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(QMessageBox.Icon.Information if title == "Success" else QMessageBox.Icon.Critical)
        msg_box.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    installr = OpenRGBInstaller()
    installr.show()
    sys.exit(app.exec())
