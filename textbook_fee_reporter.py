#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
===============================================================================
项目名称: 北京建筑大学 教材购书费用计算与统计程序
文件名称: textbook_fee_reporter.py
作者:     电气241樊彧
创建日期: 2025-09-05
最后修改: 2025-09-05
===============================================================================

【程序说明】
本程序用于对学生购书记录与教材价格进行计算和统计，自动生成各类统计报表，
以便于对账和核对数据。可用于班级的教材费用统计和分析工作。

【使用说明】
1. 确保Excel文件格式正确，列名需与程序中定义的列名匹配。
2. 配置 `BOOK_SALES_FILE` 和 `STUDENT_RECORDS_FILE` 为实际文件路径。
3. 设置 `TARGET_CLASS` 为目标班级名称。
4. 运行程序后，请仔细核对输出结果。

【免责声明】
本程序计算结果仅供参考，请务必核对程序计算的总金额与文件记录的合计金额是否一致。
如程序输出结果与文件记录存在差异，最终以文件中记录的金额为准。
作者不对因使用本程序而导致的任何直接或间接损失承担责任。

===============================================================================
"""

import pandas as pd
from collections import defaultdict
import re

def process_book_excel(file_path, target_class):
    """处理教材销售Excel文件，提取指定班级的教材信息"""
    try:
        df = pd.read_excel(file_path, header=None)

        # 定位班级行
        class_mask = df[0].astype(str).str.contains(target_class)
        if not class_mask.any():
            raise ValueError(f"未找到班级 '{target_class}'")
        class_row = class_mask.idxmax()

        # 检查标题行
        if df.iloc[class_row + 1, 0] != "序号":
            raise ValueError("未找到'序号'标题")

        header_row = class_row + 1
        # 检测所有需要的列（必须全部存在）
        required_columns = {
            '序号': df.iloc[header_row].astype(str).str.contains("序号"),
            '教材名称': df.iloc[header_row].astype(str).str.contains("教材名称"),
            '折扣价': df.iloc[header_row].astype(str).str.contains("折扣价")
        }

        # 验证所有必要列都必须存在
        for col, mask in required_columns.items():
            if not mask.any():
                raise ValueError(f"缺少必要列：'{col}'")

        # 获取各列索引
        col_indices = {col: mask.idxmax() for col, mask in required_columns.items()}

        book_data = []
        current_row = header_row + 1
        expected_serial = 1

        while current_row < len(df):
            # 检查序号连续性
            serial_value = df.iloc[current_row, col_indices['序号']]
            if pd.isna(serial_value):
                break

            serial_number = int(serial_value)
            if serial_number != expected_serial:
                raise ValueError(f"序号不连续：应为 {expected_serial} 但找到 {serial_number}")

            # 收集教材信息
            book_info = {
                '序号': serial_number,
                '教材名称': df.iloc[current_row, col_indices['教材名称']],
                '折扣价': float(df.iloc[current_row, col_indices['折扣价']]) if pd.notna(
                    df.iloc[current_row, col_indices['折扣价']]) else 0.0
            }
            book_data.append(book_info)

            current_row += 1
            expected_serial += 1

        if not book_data:
            raise ValueError("未找到有效数据")

        return book_data

    except Exception as e:
        raise ValueError(f"处理教材Excel失败: {str(e)}")

def process_student_excel(file_path, target_class):
    """处理学生购书Excel文件，提取指定班级的购书信息"""
    try:
        df = pd.read_excel(file_path, header=None, sheet_name=0, engine='openpyxl', keep_default_na=False)

        # 自动检测标题行（查找包含"姓名"的行）
        header_row = None
        for i in range(len(df)):
            if df.iloc[i].astype(str).str.contains("姓名").any():
                header_row = i
                break

        if header_row is None:
            raise ValueError("未找到包含'姓名'的标题行")

        # 检测所有需要的列（必须全部存在）
        required_columns = {
            '姓名': df.iloc[header_row].astype(str).str.contains("姓名"),
            '班级': df.iloc[header_row].astype(str).str.contains("班级"),
            '教材名称': df.iloc[header_row].astype(str).str.contains("教材名称")
        }

        # 验证所有必要列都必须存在
        for col, mask in required_columns.items():
            if not mask.any():
                raise ValueError(f"缺少必要列：'{col}'")

        # 获取各列索引
        col_indices = {col: mask.idxmax() for col, mask in required_columns.items()}

        student_records = []
        for row in range(header_row + 1, len(df)):
            student_class = str(df.iloc[row, col_indices['班级']]).strip()
            if target_class in student_class:
                record = {
                    '姓名': str(df.iloc[row, col_indices['姓名']]).strip(),
                    '教材名称': str(df.iloc[row, col_indices['教材名称']]).strip()
                }

                # 确保必填字段都有值
                if all(record.values()):
                    student_records.append(record)

        if not student_records:
            raise ValueError(f"未找到班级 '{target_class}'的有效学生购书记录")

        return student_records

    except Exception as e:
        raise ValueError(f"处理学生Excel失败: {str(e)}")

def str_format_1(input_str):
    """去除字符串中的空格和括号（保留括号内的内容）"""
    if not isinstance(input_str, str):
        input_str = str(input_str)
    # 去除所有空格（包括全角空格）
    no_space = input_str.replace(' ', '').replace('　', '')
    # 去除括号字符（保留内容）
    return no_space.replace('(', '').replace(')', '').replace('（', '').replace('）', '')

def str_format_2(input_str):
    """
    1. 先去除括号及括号中的内容（包括中文括号和英文括号）
    2. 然后去除所有符号（只保留字母、数字和中文字符）
    3. 最后去除所有空格（包括全角空格）
    """
    if not isinstance(input_str, str):
        input_str = str(input_str)

    # 第一步：去除中文括号及内容
    while '（' in input_str and '）' in input_str:
        start = input_str.find('（')
        end = input_str.find('）')
        if start < end:
            input_str = input_str[:start] + input_str[end + 1:]
        else:
            break

    # 第二步：去除英文括号及内容
    while '(' in input_str and ')' in input_str:
        start = input_str.find('(')
        end = input_str.find(')')
        if start < end:
            input_str = input_str[:start] + input_str[end + 1:]
        else:
            break

    # 第三步：去除所有符号（只保留字母、数字和中文字符）
    cleaned_str = re.sub(r'[^\w\u4e00-\u9fa5]', '', input_str)

    # 第四步：去除所有空格
    return cleaned_str.replace(' ', '').replace('　', '')

def calculate_student_payments(book_data, student_records):
    """
    计算每个学生的购书总费用（仅使用教材名称匹配）
    参数:
        book_data: 教材清单 [{'教材名称':xx, '折扣价':xx}, ...]
        student_records: 学生购书记录 [{'姓名':xx, '教材名称':xx}]
    返回:
        [{'姓名':xx, '购书费用':xx}, ...]
    """
    try:
        # 建立价格映射关系（三种格式）
        price_maps = {
            'raw_name': {},    # 原始名称
            'fmt1_name': {},   # 去除空格和括号
            'fmt2_name': {}    # 去除括号内容和所有符号
        }

        for book in book_data:
            raw_name = str(book['教材名称']).strip()
            fmt1_name = str_format_1(raw_name)
            fmt2_name = str_format_2(raw_name)

            price_maps['raw_name'][raw_name] = book['折扣价']
            price_maps['fmt1_name'][fmt1_name] = book['折扣价']
            price_maps['fmt2_name'][fmt2_name] = book['折扣价']

        student_totals = defaultdict(float)
        unmatched_records = []

        for student in student_records:
            book_name = str(student['教材名称']).strip()
            fmt1_name = str_format_1(book_name)
            fmt2_name = str_format_2(book_name)

            matched_price = None
            # 尝试按优先级匹配
            if book_name in price_maps['raw_name']:
                matched_price = price_maps['raw_name'][book_name]
            elif fmt1_name in price_maps['fmt1_name']:
                matched_price = price_maps['fmt1_name'][fmt1_name]
            elif fmt2_name in price_maps['fmt2_name']:
                matched_price = price_maps['fmt2_name'][fmt2_name]

            if matched_price is not None:
                student_totals[student['姓名']] += matched_price
            else:
                unmatched_records.append(f"学生[{student['姓名']}] 教材[{book_name}]")

        if unmatched_records:
            error_msg = "以下教材无法匹配价格：\n"
            error_msg += "\n".join(unmatched_records)
            error_msg += "\n匹配顺序：1.原始名称 2.去除括号 3.去除括号及内容"
            raise ValueError(error_msg)

        return [{'姓名': name, '购书费用': round(total, 2)} for name, total in student_totals.items()]

    except Exception as e:
        raise ValueError(f"计算购书费用失败: {str(e)}")

def generate_report(book_data, student_records, student_payments):
    """生成并打印报表"""
    print("=" * 80)
    print("教材清单：")
    print("=" * 80)
    for book in book_data:
        print(f"序号: {book['序号']}, 教材: {book['教材名称']}, 价格: ￥{book['折扣价']:.2f}")

    # 输出学生购书记录（同一学生合并）
    print("\n" + "=" * 80)
    print("学生购书记录：")
    print("=" * 80)
    student_books_map = defaultdict(list)
    for record in student_records:
        student_books_map[record['姓名']].append(record['教材名称'])

    for student_name, books in student_books_map.items():
        print(f"姓名: {student_name}, 教材: {', '.join(books)}")

    # 输出学生购书费用
    print("\n" + "=" * 80)
    print("学生购书费用：")
    print("=" * 80)
    total_amount = 0
    for payment in student_payments:
        print(f"{payment['姓名']}: ￥{payment['购书费用']:.2f}")
        total_amount += payment['购书费用']
    print(f"总计: ￥{total_amount:.2f}")
    print(f"请仔细核对程序计算的总金额与文件记录的合计金额是否一致，最终金额应以文件记录为准。")

def main(book_sales_file, student_records_file, target_class):
    """
    主处理流程
    参数:
        book_sales_file: 书商售书单文件路径
        student_records_file: 教务系统导出文件路径
        target_class: 目标班级
    """
    try:
        print("开始处理教材Excel文件...")
        book_data = process_book_excel(book_sales_file, target_class)
        print(f"成功读取 {len(book_data)} 本教材信息")

        print("开始处理学生购书Excel文件...")
        student_records = process_student_excel(student_records_file, target_class)
        print(f"成功读取 {len(student_records)} 条学生购书记录")

        print("开始计算购书费用...")
        student_payments = calculate_student_payments(book_data, student_records)
        print("费用计算完成")

        # 生成报表
        generate_report(book_data, student_records, student_payments)

    except FileNotFoundError as e:
        print(f"文件错误: {str(e)}")
    except ValueError as e:
        print(f"数据错误: {str(e)}")
    except Exception as e:
        print(f"运行时错误: {str(e)}")

if __name__ == "__main__":
    # 通过传参调用 main
    # BOOK_SALES_FILE = "./2 书商提供（按班）-学生教材售书单-北京建筑大学2024-2025学年第2学期.xlsx"
    # STUDENT_RECORDS_FILE = "./3 教务系统（按人）-学生订单导出-北京建筑大学2024-2025学年第2学期.xlsx"

    BOOK_SALES_FILE = "./书商售书单-在校生2025-2026学年第1学期教材统计表.xlsx"
    STUDENT_RECORDS_FILE = "./教务系统导出-在校生及2025年专升本同学-2025-2026学年第1学期教材统计表.xlsx"

    TARGET_CLASS = "电气241"

    main(BOOK_SALES_FILE, STUDENT_RECORDS_FILE, TARGET_CLASS)
