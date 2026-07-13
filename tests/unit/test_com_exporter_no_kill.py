"""
验证 com_image_exporter._export_with_microsoft 的进程保护逻辑

核心场景：
  1. 用户已打开 PowerPoint（GetActiveObject 命中）→ we_started=False → 绝不 Quit
  2. 用户没打开（GetActiveObject 失败，Dispatch 新建）→ we_started=True →
     仅当 Presentations.Count==0 时才 Quit（清理自己启动的进程）
  3. 边界：我们启动的实例，但用户在导出期间又打开了别的 PPT
     → Presentations.Count>0 → 不 Quit（保护用户的文档）

通过 mock win32com.client 模拟，不依赖真实 Office 环境。
"""
import sys
import os
import types
from unittest.mock import MagicMock, patch

# 项目根目录加入路径
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))


class FakePresentations:
    """模拟 powerpoint.Presentations 集合"""
    def __init__(self, count=0):
        self._count = count

    @property
    def Count(self):
        return self._count

    def Open(self, *args, **kwargs):
        return MagicMock()  # 返回一个 fake presentation


class FakePowerPoint:
    """模拟 PowerPoint.Application COM 对象"""
    def __init__(self, presentations_count=0):
        self.Presentations = FakePresentations(presentations_count)
        self._quit_called = False

    def Quit(self):
        self._quit_called = True

    @property
    def quit_called(self):
        return self._quit_called


def make_win32com_module(getactive_returns=None, dispatch_returns=None):
    """
    构造一个 fake win32com.client 模块
      getactive_returns: GetActiveObject 的返回值（None 表示抛异常）
      dispatch_returns: Dispatch 的返回值（None 表示抛异常）
    """
    fake = types.SimpleNamespace()

    def get_active(progid):
        if getactive_returns is None:
            raise Exception("无活动实例")
        return getactive_returns

    def dispatch(progid):
        if dispatch_returns is None:
            raise Exception("Dispatch 失败")
        return dispatch_returns

    fake.client = types.SimpleNamespace(
        GetActiveObject=get_active,
        Dispatch=dispatch,
    )
    return fake


def run_case(case_name, getactive_returns, dispatch_returns, expected_quit):
    """跑一个场景，断言 Quit 是否被调用符合预期"""
    # fake 实例（用来事后检查 quit_called）
    fake_pp_user = getactive_returns  # 复用路径会用到这个
    fake_pp_ours = dispatch_returns   # Dispatch 路径会用到这个

    fake_win32com = make_win32com_module(
        getactive_returns=fake_pp_user,
        dispatch_returns=fake_pp_ours,
    )

    # patch win32com 到 sys.modules
    with patch.dict(sys.modules, {'win32com': fake_win32com, 'win32com.client': fake_win32com.client}):
        # 重新导入 exporter（确保用上我们的 mock）
        if 'src.core.processors.com_image_exporter' in sys.modules:
            del sys.modules['src.core.processors.com_image_exporter']
        from src.core.processors.com_image_exporter import COMImageExporter

        exporter = COMImageExporter()
        # 强制 office_info 为 microsoft，走 _export_with_microsoft 分支
        exporter.office_info = {'type': 'microsoft', 'name': 'MS', 'version': '16'}

        # 调用 export_to_images；它内部会调 _export_with_microsoft
        # 注意：export_to_images 会先用 python-pptx 读尺寸，这里用一个真实的最小 pptx 不可得，
        # 所以直接调 _export_with_microsoft 验证逻辑
        try:
            exporter._export_with_microsoft(
                pptx_path="fake.pptx",
                output_dir="./output_test",
                image_format="PNG",
                quality="原始大小",
                slide_width=960,
                slide_height=540,
            )
        except Exception as e:
            # 导出过程会因为 mock 的 presentation 不完整而失败，但 finally 块一定执行过
            pass

    # 检查 Quit 是否被调用
    # 复用路径：检查 fake_pp_user 的 quit_called
    # Dispatch 路径：检查 fake_pp_ours 的 quit_called
    actual_quit = False
    if fake_pp_user is not None and isinstance(fake_pp_user, FakePowerPoint) and fake_pp_user.quit_called:
        actual_quit = True
    if fake_pp_ours is not None and isinstance(fake_pp_ours, FakePowerPoint) and fake_pp_ours.quit_called:
        actual_quit = True

    status = "PASS" if actual_quit == expected_quit else "FAIL"
    print(f"[{status}] {case_name}")
    print(f"       期望 Quit={expected_quit}, 实际 Quit={actual_quit}")
    return status == "PASS"


def main():
    print("=" * 60)
    print("验证 _export_with_microsoft 进程保护逻辑")
    print("=" * 60)

    results = []

    # 场景1：用户已打开 PowerPoint（GetActiveObject 命中，0 个文档残留）
    # 期望：绝不 Quit（保护用户进程）
    fake_user_pp = FakePowerPoint(presentations_count=0)
    results.append(run_case(
        "场景1: 用户已开 PPT (GetActiveObject 命中) → 绝不 Quit",
        getactive_returns=fake_user_pp,
        dispatch_returns=None,
        expected_quit=False,
    ))

    # 场景2：用户没开（GetActiveObject 失败，Dispatch 新建），导出后无残留文档
    # 期望：Quit（清理我们启动的进程，防止残留）
    fake_our_pp = FakePowerPoint(presentations_count=0)
    results.append(run_case(
        "场景2: 我们新建 PPT 进程，无残留文档 → Quit 清理",
        getactive_returns=None,
        dispatch_returns=fake_our_pp,
        expected_quit=True,
    ))

    # 场景3：我们启动的进程，但 Presentations.Count > 0
    # （模拟用户在我们导出期间又打开了别的 PPT）
    # 期望：不 Quit（保护用户刚打开的文档）
    fake_our_pp_with_docs = FakePowerPoint(presentations_count=2)
    results.append(run_case(
        "场景3: 我们新建 PPT 进程，但有其他文档残留 → 不 Quit (保护用户文档)",
        getactive_returns=None,
        dispatch_returns=fake_our_pp_with_docs,
        expected_quit=False,
    ))

    print()
    print("=" * 60)
    passed = sum(results)
    total = len(results)
    print(f"结果: {passed}/{total} 通过")
    print("=" * 60)
    sys.exit(0 if passed == total else 1)


if __name__ == "__main__":
    main()
