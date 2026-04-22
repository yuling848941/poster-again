# WPS Office COM组件注册解决方案

## 问题描述
在GUI界面的办公套件下拉菜单中选择WPS时，显示"WPS演示 当前不可用"的错误。

## 根本原因
WPS Office虽然已安装，但COM组件未正确注册到Windows系统中。这是因为：
1. WPS安装在用户目录下（AppData\Local）
2. 安装时没有自动注册COM组件
3. COM注册需要管理员权限

## 解决方案

### 方案1：使用自动注册脚本（推荐）

1. **运行快速注册脚本**
   ```bash
   python quick_wps_register.py
   ```

2. **按照提示操作**
   - 启动WPS演示程序
   - 让程序完全启动
   - 正常关闭WPS演示程序
   - 重新启动PPT应用程序

### 方案2：手动注册COM组件

1. **找到WPS安装路径**
   ```
   C:\Users\Administrator\AppData\Local\kingsoft\WPS Office\12.1.0.23542\office6
   ```

2. **以管理员身份打开命令提示符**
   - 右键点击"开始"菜单
   - 选择"Windows PowerShell (管理员)"或"命令提示符 (管理员)"

3. **运行注册命令**
   ```cmd
   cd "C:\Users\Administrator\AppData\Local\kingsoft\WPS Office\12.1.0.23542\office6"
   wpp.exe /regserver
   ```

4. **注册其他组件（可选）**
   ```cmd
   et.exe /regserver    # WPS表格
   wps.exe /regserver   # WPS文字
   ```

### 方案3：使用批处理文件

1. **双击运行** `register_wps_com.bat`
2. **如果提示权限不足**，右键选择"以管理员身份运行"

### 方案4：重新安装WPS Office

如果以上方法都不成功，建议：
1. 卸载当前的WPS Office
2. 从官网下载完整版安装包
3. **以管理员身份运行安装程序**
4. 选择"完整安装"而不是"自定义安装"

## 验证是否成功

### 方法1：检查注册表
1. 按 `Win + R`，输入 `regedit`
2. 导航到 `HKEY_CLASSES_ROOT\WPP.Application`
3. 如果存在，说明注册成功

### 方法2：使用检测工具
```bash
python debug_wps_detection.py
```

查看是否有"WPP.Application"显示为"SUCCESS 已注册"。

### 方法3：在应用程序中测试
1. 启动PPT应用程序
2. 在办公套件下拉菜单中查看是否显示"WPS演示"（没有"不可用"标记）
3. 选择"WPS演示"
4. 不应该弹出"不可用"警告

## 常见问题

### Q: 为什么需要注册COM组件？
A: COM（Component Object Model）是Windows的组件技术，应用程序需要通过COM接口与WPS Office通信。

### Q: 注册后还是不可用怎么办？
A: 尝试以下步骤：
1. 重启计算机
2. 以管理员身份重新运行注册命令
3. 检查WPS Office版本是否支持COM接口

### Q: 是否每次使用都需要注册？
A: 不需要，只需要注册一次。除非重新安装WPS Office。

### Q: 为什么我的WPS没有自动注册？
A: 从应用商店安装的WPS通常不会自动注册COM组件。建议从官网下载完整版。

## 技术详情

WPS Office的主要ProgID：
- `WPP.Application` - WPS演示
- `ET.Application` - WPS表格
- `WPS.Application` - WPS文字

注册命令 `/regserver` 会：
1. 在Windows注册表中注册COM组件
2. 创建ProgID到CLSID的映射
3. 注册组件的类型库信息

注册成功后，其他应用程序可以通过`win32com.client.Dispatch("WPP.Application")`来使用WPS的功能。