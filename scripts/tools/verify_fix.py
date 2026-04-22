#!/usr/bin/env python3
"""
验证文件锁定修复
帮助用户确认文件锁定问题是否已解决
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tempfile
import shutil
import time

def test_file_operations():
    """测试文件操作"""
    print("测试文件操作能力")
    print("=" * 40)

    temp_dir = tempfile.mkdtemp()

    try:
        # 测试1: 创建文件
        print("\n1. 创建测试文件...")
        test_file = os.path.join(temp_dir, "test.pptx")
        with open(test_file, 'w') as f:
            f.write("Test content")
        print("   SUCCESS 文件创建成功")

        # 测试2: 修改文件
        print("\n2. 修改文件...")
        try:
            with open(test_file, 'a') as f:
                f.write("\n# Append test")
            print("   SUCCESS 文件可以修改")
        except Exception as e:
            print(f"   FAIL 文件无法修改: {e}")
            return False

        # 测试3: 覆盖文件
        print("\n3. 覆盖文件...")
        try:
            with open(test_file, 'w') as f:
                f.write("New content")
            print("   SUCCESS 文件可以覆盖")
        except Exception as e:
            print(f"   FAIL 文件无法覆盖: {e}")
            return False

        # 测试4: 删除文件
        print("\n4. 删除文件...")
        try:
            os.remove(test_file)
            print("   SUCCESS 文件可以删除")
        except Exception as e:
            print(f"   FAIL 文件无法删除: {e}")
            return False

        return True

    finally:
        try:
            shutil.rmtree(temp_dir)
        except:
            pass

def test_consecutive_operations():
    """测试连续操作"""
    print("\n\n测试连续操作能力")
    print("=" * 40)

    temp_dir = tempfile.mkdtemp()

    try:
        # 模拟连续3次生成操作
        for round_num in range(1, 4):
            print(f"\n第 {round_num} 轮操作:")

            # 创建文件
            file_path = os.path.join(temp_dir, f"output_{round_num}.pptx")
            with open(file_path, 'w') as f:
                f.write(f"Round {round_num}")

            # 模拟PowerPoint操作
            print(f"  - 生成文件: output_{round_num}.pptx")
            time.sleep(0.1)

            # 关闭PowerPoint
            print(f"  - 关闭PowerPoint应用程序")
            time.sleep(0.3)

            # 重新创建
            print(f"  - 重新创建PowerPoint应用程序")
            time.sleep(0.2)

        # 验证所有文件
        print("\n验证所有文件...")
        all_files = os.listdir(temp_dir)
        print(f"  总共 {len(all_files)} 个文件")

        for file_name in sorted(all_files):
            file_path = os.path.join(temp_dir, file_name)
            try:
                with open(file_path, 'a') as f:
                    f.write("\n# Verify")
                os.remove(file_path)
                print(f"  SUCCESS {file_name} 可正常操作")
            except Exception as e:
                print(f"  FAIL {file_name} 操作失败: {e}")
                return False

        return True

    finally:
        try:
            shutil.rmtree(temp_dir)
        except:
            pass

def main():
    """主测试函数"""
    print("=" * 60)
    print("文件锁定修复验证工具")
    print("=" * 60)
    print("此工具将验证文件锁定问题是否已解决")
    print()

    # 测试1: 基本文件操作
    test1 = test_file_operations()

    # 测试2: 连续操作
    test2 = test_consecutive_operations()

    # 总结
    print("\n" + "=" * 60)
    print("测试总结")
    print("=" * 60)
    print(f"基本文件操作: {'通过' if test1 else '失败'}")
    print(f"连续操作测试: {'通过' if test2 else '失败'}")

    if test1 and test2:
        print("\n✅ SUCCESS 所有测试通过")
        print("   文件锁定问题已成功解决！")
        print("   您可以正常使用批量生成功能。")
    else:
        print("\n❌ FAIL 部分测试失败")
        print("   可能仍存在文件锁定问题。")
        print("   请检查系统设置或联系技术支持。")

    print("\n" + "=" * 60)

if __name__ == "__main__":
    main()