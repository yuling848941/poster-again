## ADDED Requirements

### Requirement: 捆绑字体与字体注册表
渲染引擎 SHALL 通过字体注册表将模板 typeface 解析为字体文件，字体来源包含随应用分发的捆绑字体（思源黑体/思源宋体 SC 子集，SIL OFL）与系统字体（Windows 字体路径表）；解析策略 SHALL 可由配置 `image_generation.font_policy`（auto/bundled_first/system_first）或环境变量控制，默认 auto 在 Windows 上 MUST 系统优先以保持既有渲染输出像素级一致。

#### Scenario: Windows 默认策略零回归
- **WHEN** 在 Windows 上以默认 auto（系统优先）渲染现有模板
- **THEN** 字体解析结果与改造前一致（系统微软雅黑优先），既有像素级测试全部通过

#### Scenario: 捆绑字体兜底
- **WHEN** 模板引用的字体系统未安装但注册表捆绑族可替代（如无中文系统字体的非 Windows 平台）
- **THEN** 渲染使用捆绑字体文件，不抛错；预检不再对该字体报告 FONT_MISSING

#### Scenario: 无直接命中的最终回退
- **WHEN** typeface 在系统与捆绑来源均无直接命中
- **THEN** 按默认链（微软雅黑→捆绑思源黑体）回退并记录告警日志，渲染继续不中断

### Requirement: 排版校准常数按物理字体绑定
行距/首行基线/删除线位置校准常数 SHALL 由字体注册表按实际解析到的物理字体提供；首行基线采用仿射模型 baseline = (a + g×lnSpc)×字号px（纯比例模型为 a=0 特例；微软雅黑历史常数 0.72×行高等价改写为 a=0/g=0.9，渲染输出不变）。思源黑体/宋体常数已经 COM 对照校准（残差 ≤2px@q2）；未经校准的字体 MUST 沿用已校准字体值并在日志中告警。

#### Scenario: 雅黑渲染不变
- **WHEN** 解析到系统微软雅黑文件
- **THEN** 使用 COM 实测校准值（ratio=1.25, a=0, g=0.9 ≡ 历史 0.72×行高），渲染输出与改造前一致

#### Scenario: 思源校准常数生效
- **WHEN** 解析到捆绑思源黑体/宋体（bundled_first 策略或非 Windows 平台）
- **THEN** 使用 COM 对照校准的仿射常数（ratio=1.1942, a=0.1852, g=0.7803），多行文本行距与 PowerPoint 参考逐像素一致
