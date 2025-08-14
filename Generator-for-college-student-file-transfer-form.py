#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import re
from datetime import datetime
from pathlib import Path
import pandas as pd
import numpy as np
from docx import Document
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
                             QFileDialog, QMessageBox, QTabWidget, QLabel, QLineEdit,
                             QGridLayout, QGroupBox, QHeaderView, QAbstractItemView,
                             QDialog, QDialogButtonBox, QFormLayout, QComboBox,
                             QProgressDialog)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

class MissingFieldsDialog(QDialog):
    """缺失字段填写对话框"""
    def __init__(self, missing_fields, row_info, parent=None):
        super().__init__(parent)
        self.setWindowTitle("缺失字段填写")
        self.setModal(True)
        self.fields = {}
        
        # 过滤掉不需要用户填写的字段
        auto_fields = {'转档字号'}  # 自动生成的字段
        fields_to_fill = missing_fields - auto_fields
        
        layout = QVBoxLayout()
        
        # 提示信息
        info_label = QLabel(f"学号：{row_info.get('学号', 'N/A')} 姓名：{row_info.get('姓名', 'N/A')}\n以下字段在Excel中未找到，请填写（可留空）：")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # 字段输入
        form_layout = QFormLayout()
        for field in fields_to_fill:
            line_edit = QLineEdit()
            # 如果是年份字段，添加提示
            if field == '年':
                line_edit.setPlaceholderText('后两位，如：24、25')
                line_edit.setMaxLength(2)
            self.fields[field] = line_edit
            form_layout.addRow(f"{field}:", line_edit)
        
        layout.addLayout(form_layout)
        
        # 按钮
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)
        self.resize(400, min(300 + len(fields_to_fill) * 30, 600))
    
    def get_values(self):
        values = {}
        for field, edit in self.fields.items():
            value = edit.text()
            # 如果是年份字段且输入了4位，只取后两位
            if field == '年' and len(value) > 2:
                value = value[-2:]
            values[field] = value
        return values

class WordGeneratorThread(QThread):
    """Word文档生成线程"""
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(int)
    error = pyqtSignal(str)
    
    def __init__(self, data_rows, template_path, output_dir):
        super().__init__()
        self.data_rows = data_rows
        self.template_path = template_path
        self.output_dir = output_dir
        self.success_count = 0
        
    def run(self):
        try:
            total = len(self.data_rows)
            for i, row_data in enumerate(self.data_rows):
                self.status.emit(f"正在生成：{row_data.get('姓名', 'unknown')}")
                
                # 生成文档
                self.generate_single_doc(row_data)
                self.success_count += 1
                
                # 更新进度
                progress = int((i + 1) / total * 100)
                self.progress.emit(progress)
                
            self.finished.emit(self.success_count)
        except Exception as e:
            self.error.emit(str(e))
    
    def generate_single_doc(self, data):
        """生成单个文档"""
        doc = Document(self.template_path)
        
        # 替换文档中的所有占位符
        for paragraph in doc.paragraphs:
            self.replace_text_in_paragraph(paragraph, data)
        
        # 替换表格中的占位符
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self.replace_text_in_paragraph(paragraph, data)
        
        # 生成文件名
        filename = f"{data.get('学号', 'unknown')}_{data.get('姓名', 'unknown')}_{data.get('班级', 'unknown')}.docx"
        # 清理文件名中的非法字符
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        
        # 保存文档
        output_path = os.path.join(self.output_dir, filename)
        doc.save(output_path)
    
    def replace_text_in_paragraph(self, paragraph, data):
        """替换段落中的占位符，保持原有格式"""
        # 先检查整个段落文本中是否包含占位符
        full_text = ''.join(run.text for run in paragraph.runs)
        
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in full_text:
                # 占位符可能被分割在多个run中，需要特殊处理
                self.replace_placeholder_in_runs(paragraph.runs, placeholder, str(value) if value else '')
    
    def replace_placeholder_in_runs(self, runs, placeholder, replacement):
        """在runs中替换占位符，处理占位符被分割的情况"""
        text = ''
        run_info = []
        
        # 收集所有run的文本和信息
        for run in runs:
            run_start = len(text)
            text += run.text
            run_end = len(text)
            run_info.append({
                'run': run,
                'start': run_start,
                'end': run_end,
                'original_text': run.text
            })
        
        # 查找并替换占位符
        new_text = text.replace(placeholder, replacement)
        
        # 如果没有变化，直接返回
        if new_text == text:
            return
        
        # 计算需要调整的位置
        placeholder_pos = text.find(placeholder)
        if placeholder_pos == -1:
            return
        
        # 找到占位符涉及的runs
        affected_runs = []
        for info in run_info:
            if info['start'] <= placeholder_pos < info['end'] or \
               info['start'] < placeholder_pos + len(placeholder) <= info['end'] or \
               (info['start'] >= placeholder_pos and info['end'] <= placeholder_pos + len(placeholder)):
                affected_runs.append(info)
        
        if affected_runs:
            # 在第一个受影响的run中进行替换
            first_run = affected_runs[0]['run']
            
            # 构建新的文本
            before_placeholder = text[:placeholder_pos]
            after_placeholder = text[placeholder_pos + len(placeholder):]
            
            # 计算第一个run应该包含的文本
            first_run_start = affected_runs[0]['start']
            first_run_text_before = before_placeholder[first_run_start:] if first_run_start < len(before_placeholder) else ''
            
            # 设置第一个run的新文本
            first_run.text = first_run_text_before + replacement
            
            # 处理剩余的文本
            remaining_text_start = placeholder_pos + len(placeholder)
            if len(affected_runs) > 1:
                last_run = affected_runs[-1]
                if last_run['end'] > remaining_text_start:
                    # 如果最后一个受影响的run还有剩余文本
                    remaining_in_last = after_placeholder[:last_run['end'] - remaining_text_start]
                    first_run.text += remaining_in_last
                
                # 清空中间的runs
                for info in affected_runs[1:]:
                    info['run'].text = ''

class ArchiveTransferGenerator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_data = None
        self.template_variables = None
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('档案转递文档批量生成工具')
        self.setGeometry(100, 100, 1200, 700)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # 创建选项卡
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # Excel导入页面
        self.excel_tab = self.create_excel_tab()
        self.tab_widget.addTab(self.excel_tab, "Excel批量生成")
        
        # 手动填写页面
        self.manual_tab = self.create_manual_tab()
        self.tab_widget.addTab(self.manual_tab, "手动填写生成")
        
        # 状态栏
        self.statusBar().showMessage('就绪')
        
    def create_excel_tab(self):
        """创建Excel导入页面"""
        tab = QWidget()
        layout = QVBoxLayout()
        
        # 按钮组
        button_layout = QHBoxLayout()
        
        self.load_excel_btn = QPushButton('📁 读取Excel文件')
        self.load_excel_btn.clicked.connect(self.load_excel)
        button_layout.addWidget(self.load_excel_btn)
        
        self.select_all_btn = QPushButton('☑ 全选')
        self.select_all_btn.clicked.connect(self.select_all)
        self.select_all_btn.setEnabled(False)
        button_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton('☐ 取消全选')
        self.deselect_all_btn.clicked.connect(self.deselect_all)
        self.deselect_all_btn.setEnabled(False)
        button_layout.addWidget(self.deselect_all_btn)
        
        self.generate_btn = QPushButton('📄 批量生成Word文档')
        self.generate_btn.clicked.connect(self.batch_generate)
        self.generate_btn.setEnabled(False)
        button_layout.addWidget(self.generate_btn)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # 数据表格
        self.data_table = QTableWidget()
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.data_table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        layout.addWidget(self.data_table)
        
        tab.setLayout(layout)
        return tab
    
    def create_manual_tab(self):
        """创建手动填写页面"""
        tab = QWidget()
        layout = QVBoxLayout()
        
        # 创建滚动区域以容纳所有字段
        form_group = QGroupBox("档案信息填写")
        form_layout = QGridLayout()
        
        # 定义所有可能的字段（不包括提交时间）
        self.manual_fields = {}
        field_list = [
            ('姓名', '请输入姓名'),
            ('学号', '请输入学号'),
            ('班级', '请输入班级'),
            ('届', '如：2023'),
            ('年', '后两位，如：24、25'),
            ('月', '如：7'),
            ('日', '如：15'),
            ('身份证号', '请输入身份证号'),
            ('收档单位名称', '请输入收档单位名称'),
            ('转递编号', '请输入转递编号'),
            ('生源地', '请输入生源地'),
            ('手机号', '请输入手机号'),
            ('档案转递类型', '请输入档案转递类型'),
            ('就业单位名称', '请输入就业单位名称'),
            ('就业单位地址', '请输入就业单位地址'),
        ]
        
        row = 0
        col = 0
        for field_name, placeholder in field_list:
            label = QLabel(f"{field_name}:")
            line_edit = QLineEdit()
            line_edit.setPlaceholderText(placeholder)
            
            # 如果是年份字段，限制最大长度为2
            if field_name == '年':
                line_edit.setMaxLength(2)
            
            self.manual_fields[field_name] = line_edit
            
            form_layout.addWidget(label, row, col * 2)
            form_layout.addWidget(line_edit, row, col * 2 + 1)
            
            col += 1
            if col >= 3:  # 每行3个字段
                col = 0
                row += 1
        
        form_group.setLayout(form_layout)
        layout.addWidget(form_group)
        
        # 按钮
        button_layout = QHBoxLayout()
        
        self.clear_btn = QPushButton('🔄 清空所有字段')
        self.clear_btn.clicked.connect(self.clear_manual_fields)
        button_layout.addWidget(self.clear_btn)
        
        self.generate_single_btn = QPushButton('📄 生成Word文档')
        self.generate_single_btn.clicked.connect(self.generate_single)
        button_layout.addWidget(self.generate_single_btn)
        
        # 快速填充今天日期
        self.fill_today_btn = QPushButton('📅 填充今天日期')
        self.fill_today_btn.clicked.connect(self.fill_today_date)
        button_layout.addWidget(self.fill_today_btn)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        tab.setLayout(layout)
        return tab
    
    def load_excel(self):
        """加载Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "选择Excel文件", 
            "", 
            "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                # 读取Excel文件
                self.excel_data = pd.read_excel(file_path)
                self.excel_data = self.excel_data.fillna('')  # 将NaN替换为空字符串
                
                # 显示数据到表格
                self.display_data()
                
                # 启用按钮
                self.select_all_btn.setEnabled(True)
                self.deselect_all_btn.setEnabled(True)
                self.generate_btn.setEnabled(True)
                
                # 检查是否有提交时间列
                status_msg = f'已加载 {len(self.excel_data)} 条记录'
                if '提交时间' in self.excel_data.columns:
                    status_msg += ' (已自动从提交时间提取年月日)'
                
                self.statusBar().showMessage(status_msg)
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取Excel文件失败：\n{str(e)}")
    
    def display_data(self):
        """显示数据到表格"""
        if self.excel_data is None:
            return
        
        # 创建显示用的数据副本
        display_df = self.excel_data.copy()
        
        # 支持多个可能的日期字段名
        date_field_names = ['提交日期', '提交时间', '日期', '时间', '创建时间', '更新时间']
        date_field = None
        
        # 查找存在的日期字段
        for field_name in date_field_names:
            if field_name in display_df.columns:
                date_field = field_name
                print(f"[显示数据] 检测到日期字段: '{field_name}'，预先提取年月日用于显示")
                break
        
        # 如果找到日期字段，添加年月日列用于预览
        if date_field:
            for idx, row in display_df.iterrows():
                date_value = row[date_field]
                if date_value and not pd.isna(date_value):
                    try:
                        # 如果是datetime对象
                        if hasattr(date_value, 'year'):
                            year = str(date_value.year)[-2:]  # 只取后两位
                            display_df.at[idx, '年(预览)'] = year
                            display_df.at[idx, '月(预览)'] = str(date_value.month)
                            display_df.at[idx, '日(预览)'] = str(date_value.day)
                        else:
                            # 如果是字符串
                            date_str = str(date_value).strip()
                            if ' ' in date_str:
                                date_part = date_str.split(' ')[0]
                            else:
                                date_part = date_str
                            
                            if '/' in date_part:
                                parts = date_part.split('/')
                                if len(parts) >= 3:
                                    year = parts[0][-2:]  # 只取后两位
                                    display_df.at[idx, '年(预览)'] = year
                                    display_df.at[idx, '月(预览)'] = str(int(parts[1])) if parts[1].isdigit() else parts[1]
                                    display_df.at[idx, '日(预览)'] = str(int(parts[2])) if parts[2].isdigit() else parts[2]
                            elif '-' in date_part:
                                parts = date_part.split('-')
                                if len(parts) >= 3:
                                    year = parts[0][-2:]  # 只取后两位
                                    display_df.at[idx, '年(预览)'] = year
                                    display_df.at[idx, '月(预览)'] = str(int(parts[1])) if parts[1].isdigit() else parts[1]
                                    display_df.at[idx, '日(预览)'] = str(int(parts[2])) if parts[2].isdigit() else parts[2]
                        
                        # 生成转档字号预览
                        if '学号' in row and '班级' in row:
                            year = display_df.at[idx, '年(预览)'] if '年(预览)' in display_df.columns else ''
                            student_id = str(row['学号']) if pd.notna(row['学号']) else ''
                            class_name = str(row['班级']) if pd.notna(row['班级']) else ''
                            if year and student_id and class_name:
                                display_df.at[idx, '转档字号(预览)'] = f"{year}{student_id}_{class_name}"
                    except:
                        pass
        
        # 设置表格
        self.data_table.setRowCount(len(display_df))
        self.data_table.setColumnCount(len(display_df.columns) + 1)
        
        # 设置表头
        headers = ['选择'] + list(display_df.columns)
        self.data_table.setHorizontalHeaderLabels(headers)
        
        # 填充数据
        for row_idx, row_data in display_df.iterrows():
            # 添加复选框
            checkbox = QTableWidgetItem()
            checkbox.setCheckState(Qt.CheckState.Unchecked)
            self.data_table.setItem(row_idx, 0, checkbox)
            
            # 添加数据
            for col_idx, value in enumerate(row_data):
                # 处理各种数据类型
                if pd.isna(value):
                    item_text = ''
                elif isinstance(value, (pd.Timestamp, datetime)):
                    # 格式化日期时间显示
                    item_text = value.strftime('%Y/%m/%d %H:%M:%S') if hasattr(value, 'strftime') else str(value)
                else:
                    item_text = str(value)
                item = QTableWidgetItem(item_text)
                self.data_table.setItem(row_idx, col_idx + 1, item)
        
        # 调整列宽
        self.data_table.resizeColumnsToContents()
        self.data_table.horizontalHeader().setStretchLastSection(True)
    
    def select_all(self):
        """全选"""
        for row in range(self.data_table.rowCount()):
            item = self.data_table.item(row, 0)
            if item:
                item.setCheckState(Qt.CheckState.Checked)
    
    def deselect_all(self):
        """取消全选"""
        for row in range(self.data_table.rowCount()):
            item = self.data_table.item(row, 0)
            if item:
                item.setCheckState(Qt.CheckState.Unchecked)
    
    def get_template_variables(self):
        """获取模板中的变量"""
        template_dir = Path('./template')
        if not template_dir.exists():
            template_dir = Path('./模板')
            if not template_dir.exists():
                QMessageBox.warning(self, "警告", "未找到template文件夹，请确保模板文件夹存在")
                return None
        
        # 查找Word模板文件
        template_files = list(template_dir.glob('*.docx'))
        if not template_files:
            QMessageBox.warning(self, "警告", "template文件夹中未找到Word模板文件")
            return None
        
        template_path = template_files[0]
        
        try:
            doc = Document(template_path)
            variables = set()
            
            # 查找段落中的变量
            for paragraph in doc.paragraphs:
                # 获取段落的完整文本（合并所有runs）
                full_text = ''.join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text
                variables.update(re.findall(r'\{\{(\w+)\}\}', full_text))
            
            # 查找表格中的变量
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            # 获取段落的完整文本（合并所有runs）
                            full_text = ''.join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text
                            variables.update(re.findall(r'\{\{(\w+)\}\}', full_text))
            
            print(f"[模板解析] 找到的变量: {variables}")
            return template_path, variables
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取模板文件失败：\n{str(e)}")
            return None
    
    def extract_date_fields(self, row_data):
        """从提交时间字段提取年月日"""
        if '提交时间' in row_data and row_data['提交时间']:
            try:
                date_str = str(row_data['提交时间']).strip()
                print(f"正在处理日期: {date_str}")
                
                # 处理可能的日期格式，去除时间部分
                if ' ' in date_str:
                    date_part = date_str.split(' ')[0]
                else:
                    date_part = date_str
                
                # 解析日期 (格式: 年/月/日 或 年-月-日)
                if '/' in date_part:
                    parts = date_part.split('/')
                    if len(parts) >= 3:
                        year = parts[0].strip()
                        month = parts[1].strip()
                        day = parts[2].strip()
                        
                        # 去除前导零
                        month = str(int(month)) if month.isdigit() else month
                        day = str(int(day)) if day.isdigit() else day
                        
                        row_data['年'] = year
                        row_data['月'] = month
                        row_data['日'] = day
                        print(f"成功提取日期 (/分隔): 年={year}, 月={month}, 日={day}")
                elif '-' in date_part:
                    parts = date_part.split('-')
                    if len(parts) >= 3:
                        year = parts[0].strip()
                        month = parts[1].strip()
                        day = parts[2].strip()
                        
                        # 去除前导零
                        month = str(int(month)) if month.isdigit() else month
                        day = str(int(day)) if day.isdigit() else day
                        
                        row_data['年'] = year
                        row_data['月'] = month
                        row_data['日'] = day
                        print(f"成功提取日期 (-分隔): 年={year}, 月={month}, 日={day}")
                else:
                    print(f"无法识别的日期格式: {date_part}")
            except Exception as e:
                print(f"解析日期失败: {date_str}, 错误: {e}")
        else:
            print("没有找到提交时间字段或字段为空")
        
        return row_data
    
    def batch_generate(self):
        """批量生成Word文档"""
        # 获取选中的行
        selected_rows = []
        for row in range(self.data_table.rowCount()):
            if self.data_table.item(row, 0).checkState() == Qt.CheckState.Checked:
                selected_rows.append(row)
        
        if not selected_rows:
            QMessageBox.warning(self, "警告", "请至少选择一行数据")
            return
        
        # 获取模板变量
        template_info = self.get_template_variables()
        if not template_info:
            return
        
        template_path, template_variables = template_info
        print(f"[模板变量] 需要的字段: {template_variables}")
        
        # 选择输出目录
        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if not output_dir:
            return
        
        # 准备数据
        data_rows = []
        for row_idx in selected_rows:
            # 直接从原始Excel数据获取
            if row_idx < len(self.excel_data):
                row_data = self.excel_data.iloc[row_idx].to_dict()
                
                # 转换NaN为空字符串
                for key in row_data:
                    if pd.isna(row_data[key]):
                        row_data[key] = ''
                    elif not isinstance(row_data[key], str):
                        # 保持日期对象不变，其他转为字符串
                        if not hasattr(row_data[key], 'year'):
                            row_data[key] = str(row_data[key])
            else:
                continue
            
            print(f"\n[第{row_idx+1}行] 原始数据字段: {list(row_data.keys())}")
            
            # 检查是否有日期相关字段
            date_fields = ['提交日期', '提交时间', '日期', '时间', '创建时间', '更新时间']
            for field in date_fields:
                if field in row_data and row_data[field]:
                    print(f"[第{row_idx+1}行] 发现{field}: '{row_data[field]}' (类型: {type(row_data[field])})")
                    break
            
            # 从日期字段提取年月日
            original_keys = set(row_data.keys())
            row_data = self.extract_date_fields(row_data)
            
            # 生成转档字号
            row_data = self.generate_transfer_number(row_data)
            
            # 打印提取结果
            new_keys = set(row_data.keys()) - original_keys
            if new_keys:
                print(f"[第{row_idx+1}行] ✓ 新增字段: {new_keys}")
                for key in ['年', '月', '日', '转档字号']:
                    if key in row_data:
                        print(f"[第{row_idx+1}行]   {key} = '{row_data[key]}'")
            
            # 检查缺失字段（排除自动生成的字段）
            auto_generated_fields = {'转档字号'}  # 自动生成的字段列表
            missing_fields = set()
            for field in template_variables:
                if field in auto_generated_fields:
                    continue  # 跳过自动生成的字段
                if field not in row_data:
                    missing_fields.add(field)
                    print(f"[第{row_idx+1}行] ✗ 字段 '{field}' 不存在")
                elif not row_data[field]:
                    missing_fields.add(field)
                    print(f"[第{row_idx+1}行] ✗ 字段 '{field}' 为空")
                else:
                    print(f"[第{row_idx+1}行] ✓ 字段 '{field}' = '{row_data[field]}'")
            
            if missing_fields:
                print(f"[第{row_idx+1}行] 缺失字段汇总: {missing_fields}")
                # 显示对话框时，提取一些信息用于显示
                info_for_dialog = {
                    '姓名': row_data.get('姓名', 'N/A'),
                    '学号': row_data.get('学号', 'N/A')
                }
                dialog = MissingFieldsDialog(missing_fields, info_for_dialog, self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    filled_values = dialog.get_values()
                    row_data.update(filled_values)
                    print(f"[第{row_idx+1}行] 用户填写了: {filled_values}")
                    
                    # 如果用户填写了年、学号或班级，重新生成转档字号
                    if any(key in filled_values for key in ['年', '学号', '班级']):
                        row_data = self.generate_transfer_number(row_data)
                        print(f"[第{row_idx+1}行] 重新生成转档字号: {row_data.get('转档字号', '')}")
                else:
                    print(f"[第{row_idx+1}行] 用户取消，跳过此行")
                    continue  # 跳过这一行
            else:
                print(f"[第{row_idx+1}行] ✓ 所有字段完整")
            
            data_rows.append(row_data)
        
        if not data_rows:
            QMessageBox.warning(self, "警告", "没有要生成的数据")
            return
        
        print(f"\n[批量生成] 准备生成 {len(data_rows)} 个文档")
        
        # 创建进度对话框
        progress_dialog = QProgressDialog("正在生成Word文档...", "取消", 0, 100, self)
        progress_dialog.setWindowTitle("批量生成进度")
        progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        progress_dialog.show()
        
        # 创建并启动生成线程
        self.generator_thread = WordGeneratorThread(data_rows, template_path, output_dir)
        self.generator_thread.progress.connect(progress_dialog.setValue)
        self.generator_thread.status.connect(lambda msg: progress_dialog.setLabelText(msg))
        self.generator_thread.finished.connect(lambda count: self.on_generation_finished(count, progress_dialog))
        self.generator_thread.error.connect(lambda msg: self.on_generation_error(msg, progress_dialog))
        self.generator_thread.start()
    
    def on_generation_finished(self, count, progress_dialog):
        """生成完成处理"""
        progress_dialog.close()
        QMessageBox.information(self, "完成", f"成功生成 {count} 个Word文档")
        self.statusBar().showMessage(f'成功生成 {count} 个文档')
    
    def on_generation_error(self, error_msg, progress_dialog):
        """生成错误处理"""
        progress_dialog.close()
        QMessageBox.critical(self, "错误", f"生成文档时出错：\n{error_msg}")
        self.statusBar().showMessage('生成失败')
    
    def clear_manual_fields(self):
        """清空手动填写的字段"""
        for field_edit in self.manual_fields.values():
            field_edit.clear()
    
    def fill_today_date(self):
        """填充今天的日期"""
        today = datetime.now()
        if '年' in self.manual_fields:
            self.manual_fields['年'].setText(str(today.year)[-2:])  # 只取后两位
        if '月' in self.manual_fields:
            self.manual_fields['月'].setText(str(today.month))
        if '日' in self.manual_fields:
            self.manual_fields['日'].setText(str(today.day))
    
    def generate_single(self):
        """生成单个Word文档"""
        # 获取模板信息
        template_info = self.get_template_variables()
        if not template_info:
            return
        
        template_path, template_variables = template_info
        
        # 收集填写的数据
        data = {}
        for field_name, field_edit in self.manual_fields.items():
            value = field_edit.text().strip()
            if value:
                data[field_name] = value
        
        # 检查必填字段
        required_fields = ['姓名', '学号', '班级']
        missing_required = [f for f in required_fields if not data.get(f)]
        if missing_required:
            QMessageBox.warning(self, "警告", f"请填写必填字段：{', '.join(missing_required)}")
            return
        
        # 检查模板中的其他变量
        missing_fields = template_variables - set(data.keys())
        if missing_fields:
            reply = QMessageBox.question(
                self, 
                "缺失字段", 
                f"以下字段未填写：{', '.join(missing_fields)}\n是否继续生成？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
            
            # 将缺失字段设为空
            for field in missing_fields:
                data[field] = ''
        
        # 选择输出目录
        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if not output_dir:
            return
        
        try:
            # 生成文档
            doc = Document(template_path)
            
            # 替换占位符（保持格式）
            for paragraph in doc.paragraphs:
                for key, value in data.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in paragraph.text:
                        # 遍历所有runs以保持格式
                        for run in paragraph.runs:
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, str(value))
            
            # 替换表格中的占位符（保持格式）
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for key, value in data.items():
                                placeholder = f"{{{{{key}}}}}"
                                if placeholder in paragraph.text:
                                    for run in paragraph.runs:
                                        if placeholder in run.text:
                                            run.text = run.text.replace(placeholder, str(value))
            
            # 生成文件名
            filename = f"{data.get('学号', 'unknown')}_{data.get('姓名', 'unknown')}_{data.get('班级', 'unknown')}.docx"
            filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
            
            # 保存文档
            output_path = os.path.join(output_dir, filename)
            doc.save(output_path)
            
            QMessageBox.information(self, "成功", f"文档已生成：\n{filename}")
            self.statusBar().showMessage(f'文档已生成：{filename}')
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成文档失败：\n{str(e)}")

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # 设置应用程序字体
    font = QFont()
    font.setPointSize(10)
    app.setFont(font)
    
    window = ArchiveTransferGenerator()
    window.show()
    
    sys.exit(app.exec())

if __name__ == '__main__':
    main()