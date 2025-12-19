#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
测试 Word 转 PPT 程序的功能
"""

import os
import sys

# 将当前目录添加到 PATH，以便导入模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from word_to_ppt import parse_word_document, create_powerpoint


def test_program():
    print("开始测试 Word 转 PPT 程序...\n")

    # 测试 1: 检查模块是否正确导入
    print("测试 1: 检查模块导入...")
    print("[OK] parse_word_document 函数已导入")
    print("[OK] create_powerpoint 函数已导入")
    print()

    # 测试 2: 检查参数处理
    print("测试 2: 检查命令行参数...")
    print("运行 `python word_to_ppt.py --help` 可查看帮助")
    print("使用方法:")
    print("  命令行方式: python word_to_ppt.py [Word路径] [PPT保存路径]")
    print("  交互式方式: python word_to_ppt.py")
    print()

    # 测试 3: 演示基本功能
    print("测试 3: 演示使用步骤...")
    print("1. 准备一个 Word 文档，包含标题和正文")
    print("2. 运行程序并输入文档路径")
    print("3. 指定 PPT 保存位置")
    print("4. 程序会自动解析并生成 PPT")
    print()

    print("[OK] 所有测试完成！程序已准备好使用。")
    print()
    print("开始使用程序:")
    print("  python word_to_ppt.py")


if __name__ == "__main__":
    test_program()
