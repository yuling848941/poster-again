# -*- mode: python ; coding: utf-8 -*-
"""
PosterAgain PyInstaller spec

体积优化要点：
1. excludes 排除所有不用的 PySide6 模块（3D/WebEngine/Multimedia/Qml 等）
2. JDK 路径过滤：PyInstaller 会误把系统 PATH 里的 JDK DLL 卷入，必须剔除
3. UPX 压缩未启用（PySide6/Python 3.13 的 CFG 保护 DLL 不兼容 UPX，会拖慢且无效）
"""

# 不需要的 PySide6 模块（QtCore/QtGui/QtWidgets/QtNetwork 是必需的，其余排除）
EXCLUDED_QT_MODULES = [
    'PySide6.Qt3DCore', 'PySide6.Qt3DRender', 'PySide6.Qt3DInput',
    'PySide6.Qt3DLogic', 'PySide6.Qt3DAnimation', 'PySide6.Qt3DExtras',
    'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngine',
    'PySide6.QtMultimedia', 'PySide6.QtMultimediaWidgets',
    'PySide6.QtPositioning', 'PySide6.QtLocation',
    'PySide6.QtSensors', 'PySide6.QtSerialPort',
    'PySide6.QtSql', 'PySide6.QtTest',
    'PySide6.QtBluetooth', 'PySide6.QtNfc',
    'PySide6.QtDesigner', 'PySide6.QtHelp',
    'PySide6.QtOpenGL', 'PySide6.QtOpenGLWidgets',
    'PySide6.QtPdf', 'PySide6.QtPdfWidgets',
    'PySide6.QtQuick', 'PySide6.QtQuickWidgets', 'PySide6.QtQuick3D', 'PySide6.QtQuickControls2',
    'PySide6.QtRemoteObjects', 'PySide6.QtScxml', 'PySide6.QtStateMachine',
    'PySide6.QtSvg', 'PySide6.QtSvgWidgets',
    'PySide6.QtTextToSpeech', 'PySide6.QtVirtualKeyboard',
    'PySide6.QtWaylandClient', 'PySide6.QtWaylandCompositor',
    'PySide6.QtWebChannel', 'PySide6.QtWebSockets', 'PySide6.QtXml',
    'PySide6.QtQml', 'PySide6.QtQuickShapes', 'PySide6.QtShaderTools',
    # 系统环境里装了但本项目完全用不到的巨型科学计算库（PyInstaller 误收集，合计 7GB+）
    # 本项目只用 pandas + numpy，以下全部排除
    'torch', 'torchvision', 'torchaudio',
    'scipy', 'numba', 'llvmlite', 'pyarrow',
    # 系统里其他无关第三方包
    'yt_dlp', 'websockets', 'curl_cffi', 'mutagen', 'secretstorage',
    'sympy', 'networkx', 'mpl_toolkits', 'matplotlib', 'PIL.ImageQt',
    # 其他不用的标准库（注意：distutils/setuptools/pip 不排除，
    # PyInstaller 内部 hook 依赖它们，排除会导致 ValueError）
    'tkinter', 'unittest', 'pydoc_data',
]

# PySide6 Qt DLL 黑名单（文件名前缀匹配）——这些是各 exclude 模块对应的 Qt6 原生 DLL
EXCLUDED_QT_DLLS = [
    'Qt6Quick', 'Qt6Qml', 'Qt6QmlModels', 'Qt6QmlMeta', 'Qt6QmlWorkerScript',
    'Qt63D', 'Qt6Pdf', 'Qt6Multimedia', 'Qt6Location', 'Qt6Sensors',
    'Qt6SerialPort', 'Qt6Sql', 'Qt6Test', 'Qt6Bluetooth', 'Qt6Nfc',
    'Qt6Designer', 'Qt6Help', 'Qt6OpenGL', 'Qt6Pdf',
    'Qt6VirtualKeyboard', 'Qt6WebChannel', 'Qt6WebSockets',
    'Qt6Quick3D', 'Qt6QuickControls2', 'Qt6QuickShapes', 'Qt6ShaderTools',
    'Qt6RemoteObjects', 'Qt6Scxml', 'Qt6StateMachine', 'Qt6Svg',
    'Qt6TextToSpeech', 'Qt6Wayland',
]


def _should_exclude_binary(path):
    """判断一个二进制文件是否应该从打包结果中剔除"""
    p = path.replace('\\', '/').lower()

    # 1. 剔除 JDK / JRE（系统 PATH 里误卷入的，Python 应用绝不需要）
    if 'jdk' in p or 'jre' in p or '/java/' in p:
        return True

    # 2. 剔除不用的 PySide6 Qt 原生 DLL
    import os
    filename = os.path.basename(p)
    for prefix in EXCLUDED_QT_DLLS:
        if filename.startswith(prefix.lower()):
            return True

    return False


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('resources', 'resources')],
    hiddenimports=[
        'PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtWidgets', 'PySide6.QtNetwork',
        'pandas', 'openpyxl', 'pptx', 'numpy', 'yaml', 'psutil',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDED_QT_MODULES,
    noarchive=False,
    optimize=0,
)

# 过滤掉不该打包的二进制（JDK DLL + 不用的 Qt DLL）
# 注意：a.binaries 的元素是 3 元组 (path, dest_dir, kind)
filtered_binaries = []
excluded_count = 0
for entry in a.binaries:
    # 兼容 2 元组和 3 元组两种格式
    binary_path = entry[0]
    if _should_exclude_binary(binary_path):
        excluded_count += 1
    else:
        filtered_binaries.append(entry)
a.binaries = filtered_binaries

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PosterAgain',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # PySide6/Python 3.13 的 CFG DLL 与 UPX 不兼容，禁用
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['resources\\icon.ico'],
)
