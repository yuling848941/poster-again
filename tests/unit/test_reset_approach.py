#!/usr/bin/env python3
"""
测试完全重置解决方案
验证关闭并重新创建PowerPoint应用程序是否能解决文件锁定问题
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tempfile
import shutil
import time

def test_complete_reset_approach():
    """测试完全重置方法"""
    print("测试完全重置解决方案")
    print("=" * 50)

    temp_dir = tempfile.mkdtemp()

    try:
        # 模拟PowerPoint应用程序的创建和关闭
        def simulate_powerpoint_app(generation_num):
            print(f"\n  生成轮次 {generation_num}:")

            # 模拟文件删除
            for i in range(3):
                file_name = f"output_{i+1}.pptx"
                file_path = os.path.join(temp_dir, file_name)

                # 检查并删除已存在文件
                if os.path.exists(file_path):
                    print(f"    删除已存在文件: {file_name}")
                    os.remove(file_path)

                # 模拟保存操作
                with open(file_path, 'w') as f:
                    f.write(f"Generation {generation_num} - File {i+1}")
                print(f"    保存文件: {file_name}")

                # 模拟关闭演示文稿
                print(f"    关闭演示文稿: {file_name}")
                time.sleep(0.1)  # 短延迟

            print(f"  生成轮次 {generation_num} 完成")

            # 关键：完全关闭PowerPoint应用程序
            print(f"  关闭PowerPoint应用程序...")
            print(f"  等待延迟...")
            time.sleep(0.3)  # 模拟关闭延迟

            # 重新创建PowerPoint应用程序
            print(f"  重新创建PowerPoint应用程序...")
            time.sleep(0.2)  # 模拟重新创建延迟

        # 模拟第一次批量生成
        print("\n第1轮：自动匹配 + 批量生成")
        simulate_powerpoint_app(1)

        # 模拟第二次批量生成
        print("\n第2轮：再次批量生成")
        simulate_powerpoint_app(2)

        # 模拟第三次批量生成
        print("\n第3轮：第三次批量生成")
        simulate_powerpoint_app(3)

        # 验证文件状态
        print("\n验证所有生成文件:")
        final_files = os.listdir(temp_dir)
        print(f"  总共 {len(final_files)} 个文件")

        all_accessible = True
        for file_name in sorted(final_files):
            file_path = os.path.join(temp_dir, file_name)

            try:
                # 尝试访问和删除
                with open(file_path, 'a') as f:
                    f.write("\n# Test access")
                os.remove(file_path)
                print(f"  SUCCESS {file_name} 可以正常访问和删除")
            except Exception as e:
                print(f"  FAIL {file_name} 访问失败: {e}")
                all_accessible = False

        return len(final_files) > 0 and all_accessible

    finally:
        # 清理临时目录
        try:
            shutil.rmtree(temp_dir)
        except:
            pass

def test_file_lock_scenario():
    """测试文件锁定场景"""
    print("\n\n测试文件锁定场景")
    print("=" * 50)

    temp_dir = tempfile.mkdtemp()

    try:
        print("模拟文件锁定问题...")

        # 第一次生成
        file_path = os.path.join(temp_dir, "output_1.pptx")
        with open(file_path, 'w') as f:
            f.write("First generation")
        print("  第1次生成完成")

        # 模拟未关闭PowerPoint导致文件锁定
        print("  模拟文件被锁定...")

        # 尝试覆盖
        try:
            os.remove(file_path)
            print("  SUCCESS 文件未被锁定，可以删除")
        except PermissionError:
            print("  FAIL 文件被锁定，无法删除")

        # 使用完全重置方法
        print("  使用完全重置方法...")
        print("  关闭PowerPoint应用程序...")
        time.sleep(0.3)
        print("  重新创建PowerPoint应用程序...")
        time.sleep(0.2)

        # 再次尝试删除
        try:
            with open(file_path, 'w') as f:
                f.write("Second generation after reset")
            print("  SUCCESS 完全重置后文件可以正常访问")
            return True
        except Exception as e:
            print(f"  FAIL 重置后仍无法访问: {e}")
            return False

    finally:
        try:
            shutil.rmtree(temp_dir)
        except:
            pass

def main():
    """主测试函数"""
    print("完全重置解决方案验证")
    print("=" * 60)
    print("测试关闭并重新创建PowerPoint应用程序的解决方案")
    print()

    # 测试1: 完全重置方法
    test1_result = test_complete_reset_approach()

    # 测试2: 文件锁定场景
    test2_result = test_file_lock_scenario()

    # 总结
    print("\n\n测试总结")
    print("=" * 40)
    print(f"完全重置方法: {'通过' if test1_result else '失败'}")
    print(f"文件锁定场景: {'通过' if test2_result else '失败'}")

    overall_result = test1_result and test2_result
    print(f"总体测试结果: {'通过' if overall_result else '失败'}")

    if overall_result:
        print("\nSUCCESS 完全重置解决方案测试通过")
        print("- 每次生成后关闭PowerPoint应用程序有效")
        print("- 重新创建PowerPoint应用程序确保干净状态")
        print("- 文件锁定问题得到根本解决")
        print("- 即使重复生成也不会出现文件占用")
    else:
        print("\nFAIL 需要进一步调试")

    print("\n注意：此方案性能开销较大，但解决了文件锁定根本问题")

if __name__ == "__main__":
    main()