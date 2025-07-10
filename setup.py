import sys
import os
from cx_Freeze import setup, Executable

# Tăng giới hạn đệ quy
sys.setrecursionlimit(5000)

# Đường dẫn tới các file cần kèm theo
icon_path = "F:\code\Python\deepTest\\icon.ico"
background_path = "F:\code\Python\deepTest\\image.jpg"
dlib_model_path = "F:\code\Python\deepTest\\dlib_face_recognition_resnet_model_v1.dat"
shape_predictor_path = "F:\code\Python\deepTest\\shape_predictor_68_face_landmarks.dat"

# Danh sách các file cần kèm theo
include_files = [
    (icon_path, os.path.basename(icon_path)),
    (background_path, os.path.basename(background_path)),
    (dlib_model_path, os.path.basename(dlib_model_path)),
    (shape_predictor_path, os.path.basename(shape_predictor_path))
]

# Các gói cần thiết
packages = [
    "os", "json", "cv2", "numpy", "PIL", "tkinter", "datetime",
    "pandas", "threading", "time", "shutil", "sqlite3", "functools",
    "face_recognition", "dlib", "pytz"
]

# Các module cụ thể cần bao gồm
includes = [
    "tkinter.ttk", "PIL.Image", "PIL.ImageTk", "sqlite3"
]

# Các module bị loại trừ
excludes = [
    "matplotlib", "PyQt5", "PyQt6", "PySide2", "PySide6", "IPython",
    "scipy", "wx", "cupy", "triton", "tornado", "qtpy", "sphinx", "docutils",
    "Cython", "sympy", "torch", "torchvision", "tensorboard",
    "setuptools", "wheel", "pytest"
]

# Tìm thư mục shape_predictor
import face_recognition
face_recognition_path = os.path.dirname(face_recognition.__file__)

build_exe_options = {
    "packages": packages,
    "includes": includes,
    "excludes": excludes,
    "include_files": include_files,
    "include_msvcr": True,
    "build_exe": "build_G",
    "optimize": 0,  # Tắt tối ưu hóa để dễ debug
    "zip_include_packages": [],  # Không nén bất cứ package nào
    "zip_exclude_packages": ["*"],  # Loại trừ tất cả package khỏi việc nén
}

# Thiết lập chương trình
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Sử dụng GUI cho Windows

setup(
    name="JinYeFace",
    version="1.0",
    description="Ứng dụng nhận diện khuôn mặt",
    author="MIS-IT",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "JinYeFace.py",
            base=base,
            icon=icon_path,
            target_name="JinYeFace.exe"
        )
    ]
)