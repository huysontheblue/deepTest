# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['JinYeFace.py'],
    pathex=[],
    binaries=[],
    datas=[('F:\\code\\Python\\deepTest\\image.jpg', '.'), ('F:\\code\\Python\\deepTest\\dlib_face_recognition_resnet_model_v1.dat', '.'), ('F:\\code\\Python\\deepTest\\shape_predictor_68_face_landmarks.dat', 'face_recognition_models/models'), ('F:\\code\\Python\\deepTest\\shape_predictor_5_face_landmarks.dat', 'face_recognition_models/models')],
    hiddenimports=['dlib', 'face_recognition', 'face_recognition_models', 'numpy', 'cv2', 'pandas'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='JinYeFace',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['F:\\code\\Python\\deepTest\\icon.ico'],
)
