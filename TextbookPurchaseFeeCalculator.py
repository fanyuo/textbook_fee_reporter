#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
===============================================================================
项目名称: 北京建筑大学 教材购书费用计算与统计程序
文件名称: TextbookPurchaseFeeCalculator.py
作者:    樊彧
创建日期: 2025-09-20
最后修改: 2025-09-20
===============================================================================

【程序说明】
本程序基于 PyQt5 图形界面，结合 pandas 处理 Excel 数据，
可自动读取书商售书单与学生购书记录，对指定班级的学生购书费用进行
计算与统计；结果以三栏 HTML 表格美化展示，并支持导出 Excel 与 TXT 格式，
方便对账和核对。

【主要功能】
1. 读取并解析书商售书单，提取教材名称及价格。
2. 读取并解析学生购书记录，按班级筛选对应学生记录。
3. 自动匹配教材价格并计算每位学生购书总额，支持模糊匹配检测。
4. 三栏美化显示：教材清单、学生购书记录、学生购书费用。
5. 支持结果导出为 Excel 或 TXT 文件。
6. 自动记忆上次选择的文件路径和班级名称。
7. 大窗口布局优化、列宽设置合理，提升可视化效果。

【使用说明】
1. 确保 Excel 文件列名与程序中的关键列匹配（例如："序号"、"教材名称"、"折扣价"、"姓名"、"班级"）。
2. 运行程序后，在界面中依次选择“售书单”、“学生记录”Excel文件，并输入目标班级（如“电气231”）。
3. 点击“开始计算”，计算结果将显示在界面右侧的三栏区域。
4. 点击“导出学生购书费用”按钮，可将统计结果保存为 Excel 或文本文件。
5. 请核对结果是否与原始文件金额一致。

【免责声明】
本程序计算结果仅供参考。由于文件格式差异、数据录入问题等原因造成的计算误差，
请使用者结合原始记录进行核对。作者不对因使用本程序而造成的任何直接或间接损失承担责任。

===============================================================================
"""

import sys
import pandas as pd
from collections import defaultdict
import re
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QTextEdit, QVBoxLayout, QWidget,
    QFileDialog, QLineEdit, QLabel, QHBoxLayout, QMessageBox, QSplitter, QGroupBox
)
from PyQt5.QtCore import Qt, QSettings


# ---------------- 数据处理函数 ----------------
def process_book_excel(file_path, target_class):
    df = pd.read_excel(file_path, header=None)
    class_mask = df[0].astype(str).str.contains(target_class)
    if not class_mask.any():
        raise ValueError(f"未找到班级 '{target_class}'")
    class_row = class_mask.idxmax()
    if df.iloc[class_row + 1, 0] != "序号":
        raise ValueError("未找到'序号'标题")
    header_row = class_row + 1
    required_columns = {
        '序号': df.iloc[header_row].astype(str).str.contains("序号"),
        '教材名称': df.iloc[header_row].astype(str).str.contains("教材名称"),
        '折扣价': df.iloc[header_row].astype(str).str.contains("折扣价")
    }
    for col, mask in required_columns.items():
        if not mask.any():
            raise ValueError(f"缺少必要列：'{col}'")
    col_indices = {col: mask.idxmax() for col, mask in required_columns.items()}
    book_data = []
    current_row = header_row + 1
    expected_serial = 1
    while current_row < len(df):
        serial_value = df.iloc[current_row, col_indices['序号']]
        if pd.isna(serial_value):
            break
        serial_number = int(serial_value)
        if serial_number != expected_serial:
            raise ValueError(f"序号不连续：应为 {expected_serial} 但找到 {serial_number}")
        book_info = {
            '序号': serial_number,
            '教材名称': str(df.iloc[current_row, col_indices['教材名称']]).strip(),
            '折扣价': float(df.iloc[current_row, col_indices['折扣价']]) if pd.notna(
                df.iloc[current_row, col_indices['折扣价']]) else 0.0
        }
        book_data.append(book_info)
        current_row += 1
        expected_serial += 1
    if not book_data:
        raise ValueError("未找到有效数据")
    return book_data


def process_student_excel(file_path, target_class):
    df = pd.read_excel(file_path, header=None, engine='openpyxl', keep_default_na=False)
    header_row = None
    for i in range(len(df)):
        if df.iloc[i].astype(str).str.contains("姓名").any():
            header_row = i
            break
    if header_row is None:
        raise ValueError("未找到包含'姓名'的标题行")
    required_columns = {
        '姓名': df.iloc[header_row].astype(str).str.contains("姓名"),
        '班级': df.iloc[header_row].astype(str).str.contains("班级"),
        '教材名称': df.iloc[header_row].astype(str).str.contains("教材名称")
    }
    for col, mask in required_columns.items():
        if not mask.any():
            raise ValueError(f"缺少必要列：'{col}'")
    col_indices = {col: mask.idxmax() for col, mask in required_columns.items()}
    student_records = []
    for row in range(header_row + 1, len(df)):
        student_class = str(df.iloc[row, col_indices['班级']]).strip()
        if target_class in student_class:
            record = {
                '姓名': str(df.iloc[row, col_indices['姓名']]).strip(),
                '教材名称': str(df.iloc[row, col_indices['教材名称']]).strip()
            }
            if all(record.values()):
                student_records.append(record)
    if not student_records:
        raise ValueError(f"未找到班级 '{target_class}'的有效学生购书记录")
    return student_records


def str_format_1(s):
    return str(s).replace(' ', '').replace('　', '').replace('(', '').replace(')', '').replace('（', '').replace('）', '')


def str_format_2(s):
    s = str(s)
    while '（' in s and '）' in s:
        s = s[:s.find('（')] + s[s.find('）') + 1:]
    while '(' in s and ')' in s:
        s = s[:s.find('(')] + s[s.find(')') + 1:]
    return re.sub(r'[^\w\u4e00-\u9fa5]', '', s).replace(' ', '').replace('　', '')


def calculate_student_payments(book_data, student_records):
    price_maps = {'raw_name': {}, 'fmt1_name': {}, 'fmt2_name': {}}
    for book in book_data:
        raw = str(book['教材名称']).strip()
        fmt1, fmt2 = str_format_1(raw), str_format_2(raw)
        price_maps['raw_name'][raw] = book['折扣价']
        price_maps['fmt1_name'][fmt1] = book['折扣价']
        price_maps['fmt2_name'][fmt2] = book['折扣价']
    totals = defaultdict(float)
    unmatched, flag_greedy = [], False
    for stu in student_records:
        name, bname = stu['姓名'], str(stu['教材名称']).strip()
        fmt1, fmt2 = str_format_1(bname), str_format_2(bname)
        price = price_maps['raw_name'].get(bname) \
                or price_maps['fmt1_name'].get(fmt1) \
                or price_maps['fmt2_name'].get(fmt2)
        if price is not None:
            if price_maps['fmt2_name'].get(fmt2) and not price_maps['raw_name'].get(bname) and not price_maps['fmt1_name'].get(fmt1):
                flag_greedy = True
            totals[name] += price
        else:
            unmatched.append(f"{name} - {bname}")
    if unmatched:
        raise ValueError("无法匹配价格的记录:\n" + "\n".join(unmatched))
    return [{'姓名': k, '购书费用': round(v, 2)} for k, v in totals.items()], flag_greedy


# -------- HTML 表格生成函数 --------
def make_html_table(headers, rows, col_widths=None):
    html = '<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse; border:1px solid #cccccc;">\n<tr>'
    for i, h in enumerate(headers):
        width_attr = f' width="{col_widths[i]}"' if col_widths and i < len(col_widths) else ''
        html += f'<th{width_attr} style="border:1px solid #cccccc; background-color:#f0f0f0;">{h}</th>'
    html += '</tr>\n'
    for row in rows:
        html += '<tr>'
        for i, cell in enumerate(row):
            width_attr = f' width="{col_widths[i]}"' if col_widths and i < len(col_widths) else ''
            html += f'<td{width_attr} style="border:1px solid #cccccc;">{cell}</td>'
        html += '</tr>\n'
    html += '</table>'
    return html


# ---------------- 界面 ----------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("教材购书费用计算与统计程序")
        self.resize(1300, 850)
        self.book_sales_file = ""
        self.student_records_file = ""
        self.result_data = None
        self.cls_name = ""
        self.settings = QSettings("BJU", "BookFeeCalc")

        layout = QVBoxLayout()

        row1 = QHBoxLayout()
        self.book_file_input = QLineEdit()
        btn1 = QPushButton("浏览")
        btn1.clicked.connect(self.select_book_file)
        row1.addWidget(QLabel("售书单:"))
        row1.addWidget(self.book_file_input)
        row1.addWidget(btn1)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.stu_file_input = QLineEdit()
        btn2 = QPushButton("浏览")
        btn2.clicked.connect(self.select_student_file)
        row2.addWidget(QLabel("学生记录:"))
        row2.addWidget(self.stu_file_input)
        row2.addWidget(btn2)
        layout.addLayout(row2)

        row3 = QHBoxLayout()
        self.class_input = QLineEdit()
        self.class_input.setPlaceholderText("如 电气231")
        row3.addWidget(QLabel("班级:"))
        row3.addWidget(self.class_input)
        layout.addLayout(row3)

        btnrow = QHBoxLayout()
        run_btn = QPushButton("开始计算")
        run_btn.clicked.connect(self.run_calculation)
        self.export_btn = QPushButton("导出学生购书费用")
        self.export_btn.clicked.connect(self.export_result)
        self.export_btn.setEnabled(False)
        btnrow.addWidget(run_btn)
        btnrow.addWidget(self.export_btn)
        layout.addLayout(btnrow)

        splitter = QSplitter(Qt.Horizontal)
        self.text_books = QTextEdit(); self.text_books.setReadOnly(True)
        self.text_students = QTextEdit(); self.text_students.setReadOnly(True)
        self.text_payments = QTextEdit(); self.text_payments.setReadOnly(True)
        gb1 = QGroupBox("教材清单"); vb1 = QVBoxLayout(); vb1.addWidget(self.text_books); gb1.setLayout(vb1)
        gb2 = QGroupBox("学生购书记录"); vb2 = QVBoxLayout(); vb2.addWidget(self.text_students); gb2.setLayout(vb2)
        gb3 = QGroupBox("学生购书费用"); vb3 = QVBoxLayout(); vb3.addWidget(self.text_payments); gb3.setLayout(vb3)
        splitter.addWidget(gb1); splitter.addWidget(gb2); splitter.addWidget(gb3)
        splitter.setSizes([400, 450, 450])
        layout.addWidget(splitter)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        self.load_settings()

    def load_settings(self):
        self.book_sales_file = self.settings.value("book_file", "")
        self.student_records_file = self.settings.value("stu_file", "")
        self.cls_name = self.settings.value("class_name", "")
        self.book_file_input.setText(self.book_sales_file)
        self.stu_file_input.setText(self.student_records_file)
        self.class_input.setText(self.cls_name)

    def save_settings(self):
        self.settings.setValue("book_file", self.book_sales_file)
        self.settings.setValue("stu_file", self.student_records_file)
        self.settings.setValue("class_name", self.cls_name)

    def select_book_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择书商售书单", "", "Excel 文件 (*.xlsx *.xls)")
        if file:
            self.book_sales_file = file
            self.book_file_input.setText(file)
            self.save_settings()

    def select_student_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择学生购书记录", "", "Excel 文件 (*.xlsx *.xls)")
        if file:
            self.student_records_file = file
            self.stu_file_input.setText(file)
            self.save_settings()

    def run_calculation(self):
        self.book_sales_file = self.book_file_input.text().strip()
        self.student_records_file = self.stu_file_input.text().strip()
        self.cls_name = self.class_input.text().strip()
        if not (self.book_sales_file and self.student_records_file and self.cls_name):
            QMessageBox.warning(self, "提示", "请完整填写")
            return
        try:
            books = process_book_excel(self.book_sales_file, self.cls_name)
            students = process_student_excel(self.student_records_file, self.cls_name)
            payments, flag_greedy = calculate_student_payments(books, students)

            books_html = make_html_table(
                headers=["序号", "教材名称", "价格"],
                rows=[[b['序号'], b['教材名称'], f"￥{b['折扣价']:.2f}"] for b in books],
                col_widths=["50", None, "100"]
            )
            self.text_books.setHtml(books_html)

            stu_map = defaultdict(list)
            for r in students:
                stu_map[r['姓名']].append(r['教材名称'])
            students_html = make_html_table(
                headers=["姓名", "所购教材"],
                rows=[[name, ", ".join(stu_map[name])] for name in sorted(stu_map.keys())],
                col_widths=["100", None]
            )
            self.text_students.setHtml(students_html)

            total = sum(p['购书费用'] for p in payments)
            pay_rows = [[p['姓名'], f"￥{p['购书费用']:.2f}"] for p in sorted(payments, key=lambda x: x['姓名'])]
            pay_rows.append(["总计", f"￥{total:.2f}"])
            payments_html = make_html_table(headers=["姓名", "金额"], rows=pay_rows, col_widths=["100", "120"])
            if flag_greedy:
                payments_html += "<p style='color:red;'>注意：部分匹配为模糊匹配，请核对</p>"
            self.text_payments.setHtml(payments_html)

            self.result_data = (payments, total)
            self.export_btn.setEnabled(True)
            self.save_settings()

        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def export_result(self):
        if not self.result_data: return
        payments, total = self.result_data
        path, _ = QFileDialog.getSaveFileName(self, "导出学生购书费用", "", "Excel 文件 (*.xlsx);;文本文件 (*.txt)")
        if not path: return
        try:
            if path.endswith(".txt"):
                with open(path, "w", encoding="utf-8") as f:
                    f.write(f"班级: {self.cls_name}\n\n姓名\t购书费用\n")
                    for p in payments: f.write(f"{p['姓名']}\t{p['购书费用']:.2f}\n")
                    f.write(f"\n总计: ￥{total:.2f}")
            else:
                df = pd.DataFrame(payments)
                df = pd.concat([df, pd.DataFrame([{"姓名": "总计", "购书费用": total}])], ignore_index=True)
                df.insert(0, "班级", [self.cls_name] + [""]*(len(df)-1))
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name=f"{self.cls_name}_购书费用")
            QMessageBox.information(self, "成功", f"结果已导出到:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "导出错误", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
