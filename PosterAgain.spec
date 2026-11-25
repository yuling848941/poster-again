# -*- mode: python ; coding: utf-8 -*-
"""
PosterAgain PyInstaller spec

体积优化要点：
1. excludes 排除所有不用的 PySide6 模块（3D/WebEngine/Multimedia/Qml 等）
2. JDK 路径过滤：PyInstaller 会误把系统 PATH 里的 JDK DLL 卷入，必须剔除
3. UPX 压缩未启用（PySide6/Python 3.13 的 CFG 保护 DLL 不兼容 UPX，会拖慢且无效）
4. 二进制黑名单剔除（详见 EXCLUDED_EXTRA_DLLS，源体积 ~38MB / exe ~19MB）：
   - opengl32sw：Qt 软件 OpenGL 渲染器（纯 Widgets 应用不加载，PyInstaller 无条件卷入）
   - mfc140u + win32ui：pywin32 的 Pythonwin IDE 组件（COM 只用 pythoncom/pywintypes）
   - qdirect2d：备用平台插件（仅 QT_QPA_PLATFORM=direct2d 时加载）
   - Qt6Network 链路：无任何网络代码（TLS/networkinformation 插件随删）
   注意：numpy 的 OpenBLAS（~19MB）不能删--numpy 2.x import 时即硬链接，
   删除直接崩溃（已实测）
5. 不捆绑思源字体（exe 少 33.5MB）：内置渲染引擎按名字使用系统已安装字体
   （font_registry 的系统字体索引），模板里的字体只要系统装了就能用，没装的回退
   系统微软雅黑。需要思源而系统没装时：把字体文件放到 exe 旁的 resources/fonts/
   即可被识别（bundled_font_dir 的候选目录之一），或把下面的 LIB_FONTS 加回 datas。
   安卓 APK 里仍自带思源（那边没有系统字体），与桌面版无关。
"""

import os

# 不需要的 PySide6 模块（QtCore/QtGui/QtWidgets 是必需的，其余排除）
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
    # 瘦身：无网络代码/不用 win32ui/不读 AVIF（PIL 插件按需惰性加载）
    'PySide6.QtNetwork', 'win32ui', 'PIL._avif',
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

# 瘦身黑名单（文件名前缀匹配，Qt 与非 Qt 混合）：
# - opengl32sw：Qt 软件 OpenGL（19.7MB，Widgets 应用永不加载）
# - qtnetwork/Qt6Network + tls/networkinformation 插件：无网络功能
# - qdirect2d：备用平台插件
# - win32ui + mfc140u：pywin32 的 Pythonwin IDE 组件（6.6MB）
# 注意 qtnetwork 前缀同时匹配 QtNetwork.pyd 与 Qt6Network.dll（前缀匹配不区分）
EXCLUDED_EXTRA_DLLS = [
    'opengl32sw', 'qt6network', 'qtnetwork',
    'qopensslbackend', 'qcertonlybackend', 'qschannelbackend', 'qnetworklistmanager',
    'qdirect2d', 'win32ui', 'mfc140u',
]


def _should_exclude_binary(path):
    """判断一个二进制文件是否应该从打包结果中剔除"""
    p = path.replace('\\', '/').lower()

    # 1. 剔除 JDK / JRE（系统 PATH 里误卷入的，Python 应用绝不需要）
    if 'jdk' in p or 'jre' in p or '/java/' in p:
        return True

    # 2. 剔除不用的 PySide6 Qt 原生 DLL + 瘦身黑名单
    import os
    filename = os.path.basename(p)
    for prefix in EXCLUDED_QT_DLLS + EXCLUDED_EXTRA_DLLS:
        if filename.startswith(prefix.lower()):
            return True

    return False


# 仓库内自带的两个库（libs/render_core、libs/fontfx，见 AGENTS.md）。
# 这里显式给出源码路径：即使 .venv 里没做可编辑安装，也能直接从源码打包。
LIB_SRC = [
    os.path.join(SPECPATH, 'libs', 'render_core', 'src'),
    os.path.join(SPECPATH, 'libs', 'fontfx', 'src'),
]
# 渲染库的捆绑字体（思源黑体/宋体，SIL OFL）。桌面版不再把它打进包（见文件头第 5 条）：
# 保留这个定义只为方便恢复——把 LIB_FONTS 加回下面 Analysis 的 datas 即可。
LIB_FONTS = (os.path.join(SPECPATH, 'libs', 'render_core', 'src', 'render_core', 'resources'),
             os.path.join('render_core', 'resources'))

a = Analysis(
    ['main.py'],
    pathex=LIB_SRC,
    binaries=[],
    datas=[('resources', 'resources')],
    hiddenimports=[
        'PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtWidgets',
        'pandas', 'openpyxl', 'pptx', 'numpy', 'yaml', 'psutil',
        'fontTools', 'fontTools.ttLib',
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
