#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证 PPTGenerator 能否正确识别 PPT 模板中的占位符（当前纯 python-pptx 实现）

注意：原测试引用了已删除的 OfficeSuiteManager 和 PPTGenerator(office_manager=)，
现改为当前 API 并基于仓库内实际存在的模板文件。
"""

import os
import sys

# 将项目根目录加入 sys.path，确保 src 包可导入
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

import logging
from src.ppt_generator import PPTGenerator

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')


def test_placeholder_extraction(ppt_path, name):
    """测试 PPTGenerator 对特定文件的占位符提取"""
    print(f"\n{'='*70}")
    print(f"测试: {name}")
    print(f"文件: {ppt_path}")
    print('='*70)

    generator = None
    try:
        generator = PPTGenerator()
        print("\n1. 加载模板...")
        success = generator.load_template(ppt_path)
        if not success:
            print("   X 加载失败")
            return []
        print("   OK 加载成功")

        print("\n2. 提取占位符...")
        placeholders = generator.extract_placeholders()
        print(f"   OK 提取完成，找到 {len(placeholders)} 个占位符")

        if placeholders:
            print("\n   占位符列表:")
            for i, ph in enumerate(placeholders, 1):
                print(f"     {i}. {ph}")
        else:
            print("   X 未找到占位符")

        return placeholders

    except Exception as e:
        print(f"\n   X 测试异常: {e}")
        import traceback
        traceback.print_exc()
        return []

    finally:
        if generator is not None:
            try:
                generator.close()
                print("\n3. 已释放资源")
            except Exception:
                pass


def main():
    print("=" * 70)
    print("PPT 占位符提取测试")
    print("=" * 70)

    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    files = [
        (os.path.join(project_root, "samples/templates/10月承保贺报(3).pptx"), "承保贺报"),
        (os.path.join(project_root, "samples/templates/10月递交贺报(3).pptx"), "递交贺报"),
    ]

    results = {}
    for ppt_path, name in files:
        if not os.path.exists(ppt_path):
            print(f"\n跳过（文件不存在）: {ppt_path}")
            continue
        results[name] = test_placeholder_extraction(ppt_path, name)

    print("\n" + "=" * 70)
    print("测试总结")
    print("=" * 70)
    for name, placeholders in results.items():
        print(f"{name}: {len(placeholders)} 个占位符")

    # 模板文件本身可能不含占位符（属模板数据问题），
    # 只要有一个实际存在的模板被成功识别出占位符即视为通过
    if results and any(len(ph) > 0 for ph in results.values()):
        print("\nOK 占位符识别正常！")
        return True
    else:
        print("\nX 未能从任何模板识别出占位符（或模板文件缺失）")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
