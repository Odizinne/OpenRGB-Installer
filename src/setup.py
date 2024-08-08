import os
from cx_Freeze import setup, Executable

src_dir = os.path.dirname(os.path.abspath(__file__))
build_dir = "build/OpenRGB-Installer"

zip_include_packages = ["PyQt6"]

include_files = [
    os.path.join(src_dir, "icons/"),
]

build_exe_options = {
    "include_files": include_files,
    "build_exe": build_dir,
    "zip_include_packages": zip_include_packages,
    "excludes": ["tkinter"],
    "silent": True,
}

executables = [
    Executable(
        os.path.join(src_dir, "openrgb-installer.py"),
        base="Win32GUI",
        icon=os.path.join(src_dir, "icons/openrgb-installer.ico"),
        target_name="OpenRGB-Installer.exe",
    )
]

setup(
    name="OpenRGB-Installer",
    version="1.0",
    options={"build_exe": build_exe_options},
    executables=executables,
)
