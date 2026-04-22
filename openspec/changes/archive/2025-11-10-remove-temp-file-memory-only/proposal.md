## Why
完全移除临时文件生成机制，将所有数据处理转换为纯内存操作，消除文件系统依赖和临时文件管理复杂性，提升系统性能和可靠性。

## What Changes
- **移除临时文件创建逻辑**：删除 `create_formatted_temp_file()` 和 `create_temp_file_with_term_data()` 方法
- **移除临时文件管理类**：删除 `TempFileManager` 相关代码
- **移除临时文件清理机制**：删除 `cleanup_temp_files()` 和相关清理逻辑
- **转换为内存处理**：所有数据格式化、期趸数据添加等操作在内存中完成
- **更新PPT生成流程**：PPT生成直接使用内存中的数据，不依赖临时文件

## Impact
- **Affected specs**: 无（这是技术债务清理，不影响功能规格）
- **Affected code**:
  - `src/data_reader.py` - 移除临时文件相关方法
  - `src/ppt_generator.py` - 更新为直接使用内存数据
  - `src/gui/main_window.py` - 移除临时文件路径处理逻辑
  - 内存管理模块 - 需要处理数据格式化和期趸数据生成
- **Breaking changes**: 移除临时文件API，但不应影响用户可见功能
- **Performance improvement**: 减少文件I/O操作，提升响应速度
- **Reliability improvement**: 消除临时文件权限、清理失败等问题