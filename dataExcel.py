import sys
import os
import time
import re
import chardet
import requests
import pandas as pd
from datetime import datetime
from urllib.parse import urlparse
from bs4 import BeautifulSoup

from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QTextEdit, QFileDialog, QMessageBox, QFrame)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QTextCursor

# =========================================================
# 后台工作线程（处理网络请求，防止UI假死）
# =========================================================
class CheckerThread(QThread):
    log_signal = pyqtSignal(str)          
    progress_signal = pyqtSignal(int)     
    finished_signal = pyqtSignal(list)    

    def __init__(self, sites_data):
        super().__init__()
        self.sites_data = sites_data
        self.is_running = True

    def detect_encoding(self, content):
        try:
            result = chardet.detect(content)
            encoding = result['encoding']
            confidence = result['confidence']
            if confidence < 0.7:
                for enc in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
                    try:
                        content.decode(enc)
                        return enc
                    except:
                        continue
                return 'utf-8'
            return encoding if encoding else 'utf-8'
        except:
            return 'utf-8'

    def extract_title_and_domain(self, html_content, original_url):
        if not original_url.startswith(('http://', 'https://')):
            original_url = 'http://' + original_url
        domain = urlparse(original_url).netloc

        title = "无法获取标题"
        try:
            encoding = self.detect_encoding(html_content)
            try:
                decoded_content = html_content.decode(encoding, errors='ignore')
            except:
                decoded_content = html_content.decode('utf-8', errors='ignore')

            soup = BeautifulSoup(decoded_content, 'html.parser')
            if soup.title and soup.title.string:
                title = soup.title.get_text(strip=True)
                if not title:
                    title = "标题为空"
        except Exception as e:
            title = f"解析异常"
            
        return title, domain

    def check_website(self, url):
        if not url.startswith(('http://', 'https://')):
            url = 'http://' + url
            
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Connection': 'keep-alive'
        }
        
        try:
            response = requests.get(url, headers=headers, timeout=10, allow_redirects=True)
            if response.status_code == 200:
                title, domain = self.extract_title_and_domain(response.content, response.url)
            else:
                title = ""
                domain = urlparse(response.url).netloc
                
            return {
                'accessible': response.status_code == 200,
                'title': title,
                'domain': domain
            }
        except requests.exceptions.RequestException as e:
            return {
                'accessible': False,
                'title': '',
                'domain': urlparse(url).netloc if urlparse(url).netloc else url
            }

    def run(self):
        results = []
        total = len(self.sites_data)
        
        self.log_signal.emit(f"=== 开始执行检查，共计 {total} 条任务 ===")
        
        for i, site in enumerate(self.sites_data, 1):
            if not self.is_running:
                self.log_signal.emit("已手动终止检查。")
                break
                
            name = site.get('name', '未知')
            link = site.get('link', '')
            
            self.log_signal.emit(f"[{i}/{total}] 正在检查: {name}...")
            
            result_data = self.check_website(link)
            
            record = {
                '标题': name,
                '地址': link,
                '是否能打开': '是' if result_data['accessible'] else '否',
                '网页标题': result_data['title'],
                '网页地址': result_data['domain']
            }
            results.append(record)
            
            if result_data['accessible']:
                self.log_signal.emit(f"  ✓ 可打开 | 域名: {result_data['domain']} | 标题: {result_data['title']}")
            else:
                self.log_signal.emit(f"  ✗ 无法打开 | {link}")
                
            self.progress_signal.emit(i)
            
            if i < total:
                time.sleep(1)
                
        self.log_signal.emit("=== 所有检查任务执行完毕 ===")
        self.finished_signal.emit(results)

    def stop(self):
        self.is_running = False


# =========================================================
# 主界面 UI 
# =========================================================
class WebsiteCheckerUI(QWidget):
    def __init__(self):
        super().__init__()
        self.sites_data = []  
        self.output_dir = ""  # 用于存储自动保存的目录路径
        self.checker_thread = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('网站可访问性批量检查工具 V1.2')
        self.resize(900, 600)
        
        # 全局字体和背景色
        self.setStyleSheet("""
            QWidget { 
                font-family: 'Microsoft YaHei'; 
                font-size: 10pt; 
                background-color: #f4f6f9;
            }
        """)

        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # ================= 左侧操作区 (卡片式美化) =================
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setAlignment(Qt.AlignTop)
        left_layout.setSpacing(20)

        # 1. 操作说明卡片
        instruction_frame = QFrame()
        instruction_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #e0e6ed;
                border-radius: 10px;
            }
            QLabel {
                color: #34495e;
                line-height: 1.6;
                padding: 5px;
            }
        """)
        instruction_layout = QVBoxLayout(instruction_frame)
        instruction_label = QLabel(
            "<span style='font-size:12pt; font-weight:bold; color:#2c3e50;'>✨ 操作说明</span><br><br>"
            "1. 点击 <b>选择表格</b> 导入本地文件。<br>"
            "   <span style='color:#7f8c8d; font-size:9pt;'>* 支持 .csv, .xls, .xlsx，可多选。</span><br>"
            "2. 表格要求：第一列为标题，第二列为链接。<br>"
            "   <span style='color:#7f8c8d; font-size:9pt;'>* 程序会自动忽略表头和空行。</span><br>"
            "3. 点击 <b>开始检查</b>，右侧将显示进度日志。<br>"
            "4. 检查完成后，结果会 <b>自动保存</b> 到导入文件<br>所在的同一目录下。"
        )
        instruction_label.setWordWrap(True)
        instruction_layout.addWidget(instruction_label)
        left_layout.addWidget(instruction_frame)

        # 2. 状态提示框
        self.status_label = QLabel("🎯 当前未选择任何文件。\n待处理网站数量：0")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                background-color: #e3f2fd;
                color: #1565c0;
                font-weight: bold;
                padding: 15px;
                border-radius: 8px;
                border: 1px dashed #90caf9;
            }
        """)
        left_layout.addWidget(self.status_label)

        # 3. 按钮组
        button_style = """
            QPushButton {
                border: none;
                border-radius: 8px;
                padding: 12px;
                font-size: 11pt;
                font-weight: bold;
            }
        """
        
        # 选择文件按钮 (科技蓝)
        self.btn_select = QPushButton("📁 1. 选择本地表格文件")
        self.btn_select.setStyleSheet(button_style + """
            QPushButton { background-color: #3498db; color: white; }
            QPushButton:hover { background-color: #2980b9; }
            QPushButton:pressed { background-color: #1f618d; }
            QPushButton:disabled { background-color: #bdc3c7; color: #ecf0f1; }
        """)
        self.btn_select.clicked.connect(self.select_files)
        left_layout.addWidget(self.btn_select)

        # 开始检查按钮 (活力绿)
        self.btn_start = QPushButton("🚀 2. 开始检查并导出")
        self.btn_start.setStyleSheet(button_style + """
            QPushButton { background-color: #2ecc71; color: white; }
            QPushButton:hover { background-color: #27ae60; }
            QPushButton:pressed { background-color: #1e8449; }
            QPushButton:disabled { background-color: #bdc3c7; color: #ecf0f1; }
        """)
        self.btn_start.clicked.connect(self.start_checking)
        self.btn_start.setEnabled(False)
        left_layout.addWidget(self.btn_start)
        
        left_layout.addStretch() # 把上面内容往上顶

        # ================= 右侧日志区 =================
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        log_label = QLabel("<b>📝 执行日志</b>")
        log_label.setStyleSheet("color: #2c3e50; font-size: 11pt;")
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e; 
                color: #00ff00; 
                font-family: Consolas, 'Microsoft YaHei'; 
                border-radius: 8px;
                padding: 10px;
                font-size: 10pt;
            }
        """)
        
        right_layout.addWidget(log_label)
        right_layout.addWidget(self.log_text)

        # 将左右布局加入主布局，设置比例 1:2
        main_layout.addWidget(left_widget, 1)
        main_layout.addWidget(right_widget, 2)
        
        self.setLayout(main_layout)

    def append_log(self, text):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {text}")
        self.log_text.moveCursor(QTextCursor.End)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "选择表格文件", 
            "", 
            "表格文件 (*.xlsx *.xls *.csv)"
        )
        
        if not files:
            return

        # 获取第一个文件的所在目录，作为后续的自动保存路径
        self.output_dir = os.path.dirname(files[0])

        self.sites_data.clear()
        self.append_log(f"选中了 {len(files)} 个文件，正在读取数据...")

        for file in files:
            try:
                if file.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, encoding='utf-8')
                    except UnicodeDecodeError:
                        df = pd.read_csv(file, encoding='gbk')
                else:
                    df = pd.read_excel(file)

                if df.shape[1] < 2:
                    self.append_log(f"⚠️ 文件 {os.path.basename(file)} 列数不足两列，跳过。")
                    continue

                count = 0
                for index, row in df.iterrows():
                    name = str(row.iloc[0]).strip()
                    link = str(row.iloc[1]).strip()
                    
                    if not link or link.lower() == 'nan' or name.lower() == 'nan':
                        continue
                        
                    self.sites_data.append({'name': name, 'link': link})
                    count += 1
                
                self.append_log(f"成功从 {os.path.basename(file)} 读取 {count} 条数据。")

            except Exception as e:
                self.append_log(f"❌ 读取文件 {os.path.basename(file)} 失败: {str(e)}")

        total_sites = len(self.sites_data)
        self.status_label.setText(f"🎯 提取成功！\n共提取网站：{total_sites} 条")
        self.status_label.setStyleSheet("""
            QLabel {
                background-color: #e8f5e9;
                color: #2e7d32;
                font-weight: bold;
                padding: 15px;
                border-radius: 8px;
                border: 1px dashed #81c784;
            }
        """)
        
        if total_sites > 0:
            self.btn_start.setEnabled(True)
            self.append_log(f"--- 准备就绪，点击【开始检查】启动任务 ---")
        else:
            self.btn_start.setEnabled(False)

    def start_checking(self):
        if not self.sites_data:
            return

        self.btn_select.setEnabled(False)
        self.btn_start.setEnabled(False)
        self.btn_start.setText("⏳ 检查进行中...")

        self.log_text.clear()

        self.checker_thread = CheckerThread(self.sites_data)
        self.checker_thread.log_signal.connect(self.append_log)
        self.checker_thread.finished_signal.connect(self.export_results)
        self.checker_thread.start()

    def export_results(self, results):
        self.append_log("正在自动生成 Excel 文件...")
        
        if not results:
            self.append_log("没有生成任何结果！")
            self.reset_ui()
            return

        # 构造自动保存的完整路径
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"网站检查结果_{timestamp}.xlsx"
        save_path = os.path.join(self.output_dir, file_name)

        try:
            # 转换成 DataFrame 并导出到 Excel
            df = pd.DataFrame(results)
            columns_order = ['标题', '地址', '是否能打开', '网页标题', '网页地址']
            df = df[columns_order]
            
            df.to_excel(save_path, index=False, engine='openpyxl')
            self.append_log(f"🎉 导出成功！\n文件已自动保存在:\n{save_path}")
            
            # 弹窗提示，并询问是否要打开该文件夹
            reply = QMessageBox.information(
                self, 
                "任务完成", 
                f"检查任务已完成！\n结果已自动保存至：\n{save_path}\n\n是否立即打开所在文件夹？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            
            # 如果用户点击“是”，自动打开对应的文件夹
            if reply == QMessageBox.Yes:
                # 兼容 Windows 系统的路径打开方式
                os.startfile(self.output_dir)
                
        except Exception as e:
            self.append_log(f"❌ 自动导出失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"保存文件时出错:\n{str(e)}\n\n请检查文件是否被其他程序占用。")

        self.reset_ui()

    def reset_ui(self):
        self.btn_select.setEnabled(True)
        self.btn_start.setEnabled(True)
        self.btn_start.setText("🚀 2. 开始检查并导出")

    def closeEvent(self, event):
        if self.checker_thread and self.checker_thread.isRunning():
            self.checker_thread.stop()
            self.checker_thread.wait()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = WebsiteCheckerUI()
    window.show()
    sys.exit(app.exec_())
