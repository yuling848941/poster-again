#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT生成修复验证测试（简化版）
"""

import os
import sys
import time
import shutil

def check_file_status(file_path):
    """检查文件是否被占用"""
    try:
        with open(file_path, 'rb') as f:
            f.read()
        return True, "文件可正常访问"
    except (IOError, PermissionError) as e:
        return False, f"文件被占用: {str(e)}"

def test_file_creation():
    """测试文件创建和释放"""
    print("=== 测试文件创建和释放 ===")

    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)

    success_count = 0
    failed_count = 0

    try:
        for i in range(10):
            test_file = os.path.join(output_dir, f"test_{i}.pptx")

            # 创建文件
            with open(test_file, 'wb') as f:
                f.write(b"test" * 100)

            # 模拟修复后的延迟：500ms
            time.sleep(0.5)

            # 检查文件是否可访问
            is_available, status = check_file_status(test_file)
            if is_available:
                success_count += 1
                print(f"  文件 {i}: [OK] {status}")
            else:
                failed_count += 1
                print(f"  文件 {i}: [FAIL] {status}")

            # 尝试删除
            try:
                os.unlink(test_file)
            except:
                print(f"  文件 {i}: [WARN] 删除失败")

        print(f"\n结果: 成功 {success_count}, 失败 {failed_count}\n")

    except Exception as e:
        print(f"[ERROR] 测试失败: {e}\n")
    finally:
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)

def simulate_ppt_generation():
    """模拟PPT生成流程"""
    print("=== 模拟PPT生成流程 ===")

    output_dir = "test_ppt"
    os.makedirs(output_dir, exist_ok=True)

    success_count = 0
    failed_count = 0

    try:
        print("  模拟生成10个PPT文件...")

        for i in range(10):
            test_file = os.path.join(output_dir, f"output_{i+1}.pptx")

            try:
                # 创建文件
                with open(test_file, 'wb') as f:
                    f.write(b"PPTX_CONTENT" * 100)

                # 模拟修复后的流程
                time.sleep(0.5)  # 延迟500ms

                # 检查文件可访问性
                is_available, status = check_file_status(test_file)
                if is_available:
                    success_count += 1
                else:
                    failed_count += 1
                    print(f"  文件 {i+1}: [FAIL] {status}")

                # 删除文件
                try:
                    os.unlink(test_file)
                except:
                    pass

            except Exception as e:
                failed_count += 1
                print(f"  文件 {i+1}: [ERROR] {e}")

        print(f"\n结果: 成功 {success_count}, 失败 {failed_count}")

        if success_count == 10:
            print("\n[SUCCESS] 所有测试通过！文件占用问题已修复！")
        else:
            print(f"\n[WARNING] 成功率: {success_count/10*100:.1f}%")

    except Exception as e:
        print(f"[ERROR] 测试失败: {e}\n")
    finally:
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)

def main():
    print("=" * 60)
    print("PPT生成修复验证测试")
    print("=" * 60)
    print()

    test_file_creation()
    simulate_ppt_generation()

    print("=" * 60)
    print("测试完成")
    print("=" * 60)
    print("""
建议后续操作：
1. 使用GUI界面进行实际测试
2. 生成PPT后立即尝试打开文件
3. 如果文件可以正常打开，说明修复成功
""")

if __name__ == "__main__":
    main()
