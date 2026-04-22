## Why
用户需要在GUI界面实现"直接生成图片"功能，将生成的PPT直接转换为指定格式的图片文件，而不需要用户手动转换。

## What Changes
- 在PPTGenerator类中添加图片生成功能
- 修改批量生成流程以支持图片格式输出
- 更新GUI工作线程以处理图片生成进度
- 添加图片格式配置和错误处理

## Impact
- 受影响规格: ppt-generation, batch-processing
- 受影响代码: src/ppt_generator.py, src/gui/ppt_worker_thread.py, src/gui/main_window.py