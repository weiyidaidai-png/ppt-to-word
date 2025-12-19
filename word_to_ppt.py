#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Word 转 PPT 自动化程序
将 Word 文档内容解析为 PPT 幻灯片，标题生成新页面，正文作为页面内容
"""

import argparse
import os
from typing import List, Dict
import tkinter as tk
from tkinter import filedialog, messagebox

from docx import Document
from pptx import Presentation


def parse_word_document(word_path: str) -> List[Dict[str, any]]:
    """
    解析 Word 文档，提取标题和对应的正文内容

    Args:
        word_path: Word 文档路径

    Returns:
        list: 包含章节信息的列表，每个元素为字典 {"title": str, "content": list}
    """
    doc = Document(word_path)
    sections = []
    current_title = None
    current_content = []
    title_font_size_threshold = 14  # 默认标题字号阈值

    for para in doc.paragraphs:
        if not para.text.strip():  # 跳过空行
            continue

        # 检测是否为标题
        is_title = False
        para_font_size = None

        # 获取段落的主要字号
        if para.runs:
            para_font_size = para.runs[0].font.size
            if para_font_size:
                para_font_size = para_font_size.pt

        # 判断标题的条件：
        # 1. 使用了标题样式（Heading 1, Heading 2, Heading 3）
        # 2. 字号大于等于阈值
        if para.style.name.startswith("Heading") or (para_font_size and para_font_size >= title_font_size_threshold):
            is_title = True

        if is_title:
            # 保存上一个章节
            if current_title:
                sections.append({
                    "title": current_title,
                    "content": current_content
                })
            # 开始新章节
            current_title = para.text.strip()
            current_content = []
        else:
            # 添加到当前章节的正文
            if para.text.strip():
                current_content.append(para.text.strip())

    # 保存最后一个章节
    if current_title:
        sections.append({
            "title": current_title,
            "content": current_content
        })

    # 如果没有识别到标题，使用默认策略
    if not sections and len(doc.paragraphs) > 0:
        print("[WARN] 警告：未检测到明显的标题样式，将第一段落作为标题处理")
        all_content = []
        for para in doc.paragraphs:
            if para.text.strip():
                all_content.append(para.text.strip())
        if all_content:
            sections.append({
                "title": all_content[0],
                "content": all_content[1:]
            })

    return sections


def create_powerpoint(sections: List[Dict[str, any]], ppt_path: str) -> None:
    """
    根据解析的章节内容创建 PowerPoint 文件

    Args:
        sections: 章节信息列表
        ppt_path: 生成的 PPT 文件保存路径
    """
    prs = Presentation()

    for section in sections:
        # 使用标题和正文布局的幻灯片
        slide_layout = prs.slide_layouts[1]  # 布局 1 通常是标题 + 内容
        slide = prs.slides.add_slide(slide_layout)

        # 设置标题
        title = slide.shapes.title
        title.text = section["title"]

        # 设置正文内容
        content_placeholder = slide.placeholders[1]
        content_box = content_placeholder.text_frame
        content_box.clear()  # 清空默认内容

        for idx, content in enumerate(section["content"]):
            # 添加段落
            p = content_box.add_paragraph()
            p.text = content

            # 如果是多个正文，添加项目符号
            if len(section["content"]) > 1:
                p.level = 0  # 第一级项目符号

    # 保存 PPT
    prs.save(ppt_path)
    print(f"[OK] PPT 已成功保存到: {ppt_path}")


def select_word_file() -> str:
    """打开文件选择对话框选择 Word 文件"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择 Word 文档",
        filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
    )
    return file_path


def select_save_path() -> str:
    """打开文件保存对话框选择 PPT 保存路径"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    save_path = filedialog.asksaveasfilename(
        title="保存 PPT 文件",
        defaultextension=".pptx",
        filetypes=[("PowerPoint 文件", "*.pptx"), ("所有文件", "*.*")]
    )
    return save_path


def main_gui() -> None:
    """图形用户界面版本的主函数"""
    print("=" * 50)
    print("Word 转 PPT 自动化程序")
    print("=" * 50)
    print()

    # 选择 Word 文件
    print("[INFO] 请选择要转换的 Word 文档...")
    word_path = select_word_file()
    if not word_path:
        print("[ERROR] 未选择 Word 文件")
        messagebox.showerror("错误", "未选择要转换的 Word 文档")
        return

    # 选择保存路径
    print(f"[INFO] 已选择 Word 文件: {word_path}")
    print("[INFO] 请选择 PPT 保存路径...")
    ppt_path = select_save_path()
    if not ppt_path:
        print("[ERROR] 未选择保存路径")
        messagebox.showerror("错误", "未选择 PPT 保存路径")
        return

    # 验证输入
    if not os.path.exists(word_path):
        print(f"[ERROR] Word 文件不存在: {word_path}")
        messagebox.showerror("错误", f"Word 文件不存在: {word_path}")
        return

    # 验证输出路径
    ppt_dir = os.path.dirname(ppt_path)
    if ppt_dir and not os.path.exists(ppt_dir):
        os.makedirs(ppt_dir)

    # 解析 Word 文档
    print(f"[INFO] 正在解析 Word 文档: {word_path}")
    sections = parse_word_document(word_path)

    if not sections:
        print("[ERROR] 文档内容为空或无法解析")
        messagebox.showerror("错误", "文档内容为空或无法解析")
        return

    print(f"[INFO] 成功解析 {len(sections)} 个章节")

    # 创建 PPT
    print(f"[INFO] 正在生成 PPT: {ppt_path}")
    try:
        create_powerpoint(sections, ppt_path)
        messagebox.showinfo("成功", f"PPT 已成功保存到:\n{ppt_path}")
    except Exception as e:
        print(f"[ERROR] 生成 PPT 时出错: {str(e)}")
        messagebox.showerror("错误", f"生成 PPT 时出错: {str(e)}")


def main_cli() -> None:
    """命令行版本的主函数"""
    parser = argparse.ArgumentParser(description="Word 转 PPT 自动化程序")
    parser.add_argument("word_path", nargs="?", help="Word 文档路径")
    parser.add_argument("ppt_path", nargs="?", help="生成的 PPT 保存路径")
    args = parser.parse_args()

    # 如果命令行参数未提供，使用交互式输入
    if not args.word_path:
        args.word_path = input("请输入 Word 文档路径: ").strip()
    if not args.ppt_path:
        args.ppt_path = input("请输入 PPT 保存路径: ").strip()

    # 验证输入
    if not os.path.exists(args.word_path):
        print(f"[ERROR] 错误: Word 文件不存在: {args.word_path}")
        return

    # 验证输出路径
    ppt_dir = os.path.dirname(args.ppt_path)
    if ppt_dir and not os.path.exists(ppt_dir):
        os.makedirs(ppt_dir)

    # 解析 Word 文档
    print(f"[INFO] 正在解析 Word 文档: {args.word_path}")
    sections = parse_word_document(args.word_path)

    if not sections:
        print("[ERROR] 错误: 文档内容为空或无法解析")
        return

    print(f"[INFO] 成功解析 {len(sections)} 个章节")

    # 创建 PPT
    print(f"[INFO] 正在生成 PPT: {args.ppt_path}")
    create_powerpoint(sections, args.ppt_path)


def main() -> None:
    """主函数，根据参数选择运行模式"""
    import sys
    # 如果有命令行参数，使用命令行模式，否则使用 GUI 模式
    if len(sys.argv) > 1:
        main_cli()
    else:
        try:
            main_gui()
        except Exception as e:
            print(f"[ERROR] GUI 模式运行失败，切换到命令行模式: {str(e)}")
            main_cli()


if __name__ == "__main__":
    main()