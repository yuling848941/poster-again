"""
验证图片质量参数从 UI 选择一路贯穿到 slide.Export 的 width/height

核心场景（每个 UI 档位都验证 Microsoft 和 WPS 两条路径）：
  - 原始(1.0x) → Export 收到 width = slide_width * 1
  - 增强(2.0x) → Export 收到 width = slide_width * 2
  - 高质量(3.0x) → Export 收到 width = slide_width * 3

通过 mock win32com.client + python-pptx，不依赖真实 Office。
"""
import sys
import os
import types
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))


class FakeSlide:
    """记录 Export 调用参数的假 slide"""
    def __init__(self, index):
        self.index = index
        self.export_calls = []  # 记录每次 Export 的参数

    def Export(self, *args):
        self.export_calls.append(args)


class FakePresentation:
    """假 presentation，含若干 slide"""
    def __init__(self, slide_count=1):
        self._slides = [FakeSlide(i) for i in range(1, slide_count + 1)]

    @property
    def Slides(self):
        class _SlidesAccessor:
            def __init__(self, slides):
                self._slides = slides
                self.Count = len(slides)

            def __call__(self, i):
                return self._slides[i - 1]
        return _SlidesAccessor(self._slides)

    def Close(self):
        pass


class FakePowerPoint:
    """假 PowerPoint.Application，Presentations.Open 返回 FakePresentation"""
    def __init__(self, slide_count=1):
        self._presentation = FakePresentation(slide_count)
        self.Presentations = MagicMock()
        self.Presentations.Open.return_value = self._presentation
        self.Presentations.Count = 0
        self._quit_called = False

    def Quit(self):
        self._quit_called = True

    @property
    def presentation(self):
        return self._presentation


def make_win32com(powerpoint_instance):
    """构造 fake win32com，GetActiveObject 返回给定实例，Dispatch 抛异常"""
    fake = types.SimpleNamespace()

    def get_active(progid):
        return powerpoint_instance  # 复用已有实例

    def dispatch(progid):
        raise Exception("不应走到 Dispatch")

    fake.client = types.SimpleNamespace(
        GetActiveObject=get_active,
        Dispatch=dispatch,
    )
    return fake


def run_quality_case(office_type, ui_text, expected_scale):
    """
    跑一个档位 + 一条路径，断言 slide.Export 收到的 width/height 是 slide_dim * expected_scale
    """
    from src.gui.settings_manager import SettingsManager
    from src.core.processors.com_image_exporter import COMImageExporter

    # 1. UI 文本 → float（这是 main_window 里会发生的第一步转换）
    quality_float = SettingsManager.quality_text_to_value(ui_text)
    assert quality_float == expected_scale, \
        f"UI 文本 '{ui_text}' → float 应为 {expected_scale}，实际 {quality_float}"

    # 2. 构造 mock 环境
    fake_pp = FakePowerPoint(slide_count=1)
    fake_win32com = make_win32com(fake_pp)

    # 3. 调 COMImageExporter._export_with_microsoft 或 _export_with_wps
    #    模拟 slide 基础尺寸 960x540
    slide_w, slide_h = 960, 540

    with patch.dict(sys.modules, {'win32com': fake_win32com, 'win32com.client': fake_win32com.client}):
        exporter = COMImageExporter()
        exporter.office_info = {'type': office_type, 'name': office_type, 'version': '16'}

        if office_type == 'microsoft':
            exporter._export_with_microsoft(
                "fake.pptx", "./output_test", "PNG", quality_float, slide_w, slide_h
            )
        else:
            exporter._export_with_wps(
                "fake.pptx", "./output_test", "PNG", quality_float, slide_w, slide_h
            )

    # 4. 检查 slide.Export 收到的参数
    slide = fake_pp.presentation._slides[0]
    assert len(slide.export_calls) == 1, f"Export 应被调用 1 次，实际 {len(slide.export_calls)} 次"
    call_args = slide.export_calls[0]

    # Export(output_path, format, width, height) —— 4 个参数
    assert len(call_args) == 4, \
        f"Export 应收到 4 个参数 (path, fmt, w, h)，实际 {len(call_args)} 个: {call_args}"

    actual_path, actual_fmt, actual_w, actual_h = call_args
    expected_w = int(slide_w * expected_scale)
    expected_h = int(slide_h * expected_scale)

    assert actual_w == expected_w, \
        f"[{office_type}/{ui_text}] Export width 应为 {expected_w}，实际 {actual_w}"
    assert actual_h == expected_h, \
        f"[{office_type}/{ui_text}] Export height 应为 {expected_h}，实际 {actual_h}"

    print(f"[PASS] {office_type:9s} | {ui_text} ({expected_scale}x) → Export 收到 {actual_w}x{actual_h}")
    return True


def main():
    print("=" * 64)
    print("验证图片质量参数贯穿（UI → float → slide.Export 的 width/height）")
    print("=" * 64)

    results = []
    cases = [
        ("原始", 1.0),
        ("增强", 2.0),
        ("高质量", 3.0),
    ]

    for office_type in ("microsoft", "wps"):
        print(f"\n--- {office_type} 路径 ---")
        for ui_text, scale in cases:
            try:
                results.append(run_quality_case(office_type, ui_text, scale))
            except AssertionError as e:
                print(f"[FAIL] {office_type:9s} | {ui_text}: {e}")
                results.append(False)
            except Exception as e:
                print(f"[FAIL] {office_type:9s} | {ui_text}: 异常 {type(e).__name__}: {e}")
                results.append(False)

    print()
    print("=" * 64)
    passed = sum(results)
    total = len(results)
    print(f"结果: {passed}/{total} 通过")
    if passed == total:
        print("✓ 所有档位 × 两条路径 都正确传递了 width/height 参数")
    print("=" * 64)
    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()
