import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
import json
import requests
from bs4 import BeautifulSoup
import threading
import re
import urllib3
import os
import time
from urllib.parse import urlparse, urljoin

# 忽略不安全的SSL警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 设置现代UI的主题和颜色
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ExcelInspectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("数据源表质量巡检工具 v4.5 (含链接智联版阅读器)")
        self.root.geometry("1000x650")
        
        self.file_path = ""
        self.is_running = False
        
        self.setup_ui()

    def setup_ui(self):
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # ================= 左侧操作区 =================
        left_frame = ctk.CTkFrame(self.root, width=320, corner_radius=15)
        left_frame.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")
        left_frame.grid_propagate(False)

        title_label = ctk.CTkLabel(left_frame, text="⚙️ 巡检参数配置", font=("微软雅黑", 20, "bold"))
        title_label.pack(pady=(20, 20))

        # 1. 文件选择
        ctk.CTkLabel(left_frame, text="1. 选择源表 (.xlsx)", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=20, pady=(0, 5))
        self.btn_select = ctk.CTkButton(left_frame, text="上传源表文件", command=self.select_file, height=40)
        self.btn_select.pack(fill="x", padx=20)
        self.lbl_file = ctk.CTkLabel(left_frame, text="未选择文件", text_color="gray", font=("微软雅黑", 12))
        self.lbl_file.pack(anchor="w", padx=20, pady=(5, 15))

        # 2. 违禁词配置
        ctk.CTkLabel(left_frame, text="2. 屏蔽关键词配置", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=20, pady=(0, 5))
        self.txt_keywords = ctk.CTkTextbox(left_frame, height=180, font=("微软雅黑", 12))
        self.txt_keywords.pack(fill="x", padx=20, pady=(0, 15))
        
        default_keywords = "初审,一审,二审,三审,复审,终审,上一页,下一页,上一篇,下一篇,上页,下页,上篇,下篇,打印按钮,无障碍阅读,扫一扫,下载DOC,关闭窗口"
        self.txt_keywords.insert("0.0", default_keywords.replace(",", "\n"))

        # 3. 最小字数配置
        ctk.CTkLabel(left_frame, text="3. 正文最小字数", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=20, pady=(0, 5))
        self.ent_min_words = ctk.CTkEntry(left_frame, height=40)
        self.ent_min_words.insert(0, "50")
        self.ent_min_words.pack(fill="x", padx=20, pady=(0, 20))

        # 4. 执行按钮
        self.btn_start = ctk.CTkButton(left_frame, text="▶ 开始执行巡检", font=("微软雅黑", 15, "bold"), 
                                       command=self.start_inspection, height=45, fg_color="#2FA572", hover_color="#1F7A54")
        self.btn_start.pack(fill="x", side="bottom", padx=20, pady=20)

        # ================= 右侧日志区 =================
        right_frame = ctk.CTkFrame(self.root, corner_radius=15, fg_color="transparent")
        right_frame.grid(row=0, column=1, padx=(0, 15), pady=15, sticky="nsew")

        ctk.CTkLabel(right_frame, text="🖥️ 运行日志", font=("微软雅黑", 18, "bold")).pack(anchor="w", pady=(0, 10))
        
        self.log_area = ctk.CTkTextbox(right_frame, font=("Consolas", 13), fg_color="#1E1E1E", text_color="#00FF00")
        self.log_area.pack(fill="both", expand=True)
        self.log_area.configure(state="disabled")

    def log(self, message):
        def append():
            self.log_area.configure(state="normal")
            time_str = time.strftime("[%H:%M:%S] ")
            self.log_area.insert("end", time_str + message + "\n")
            self.log_area.see("end")
            self.log_area.configure(state="disabled")
        self.root.after(0, append)

    def select_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file:
            self.file_path = file
            filename = os.path.basename(file)
            self.lbl_file.configure(text=filename, text_color="#2FA572")
            self.log(f"✅ 已成功导入源表: {filename}")

    def start_inspection(self):
        if not self.file_path:
            messagebox.showwarning("提示", "请先在左侧上传源表文件！")
            return
        if self.is_running:
            return
            
        try:
            min_words = int(self.ent_min_words.get().strip())
        except ValueError:
            messagebox.showerror("错误", "最小字数必须是整数！")
            return

        kw_text = self.txt_keywords.get("0.0", "end").replace(",", "\n")
        keywords = [k.strip() for k in kw_text.split('\n') if k.strip()]

        self.is_running = True
        self.btn_start.configure(state="disabled", text="⏳ 正在执行中...", fg_color="gray")
        self.log_area.configure(state="normal")
        self.log_area.delete("0.0", "end")
        self.log_area.configure(state="disabled")
        
        self.log("="*40)
        self.log("🚀 开始巡检任务")
        self.log("="*40)
        
        threading.Thread(target=self.process_excel, args=(self.file_path, keywords, min_words), daemon=True).start()

    def generate_html_reader(self, save_path, html_data_list):
        """生成带超链接追踪的沉浸式 HTML 阅读器"""
        json_data = json.dumps(html_data_list, ensure_ascii=False)
        
        # 包含了用户修改的样式 br{display:none} 等
        html_template = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>沉浸式正文阅读器</title>
    <style>
        body {{ margin: 0; padding: 0; background-color: #F4F6F8; font-family: 'Microsoft YaHei', sans-serif; color: #333; }}
        .header {{ position: fixed; top: 0; width: 100%; background: #fff; box-shadow: 0 2px 10px rgba(0,0,0,0.05); display: flex; justify-content: space-between; align-items: center; padding: 10px 30px; box-sizing: border-box; z-index: 1000; }}
        .progress-box {{ font-size: 16px; font-weight: bold; color: #2FA572; }}
        .jump-box input {{ width: 50px; text-align: center; padding: 4px; border: 1px solid #ddd; border-radius: 4px; outline: none; margin: 0 5px; }}
        .jump-box button {{ padding: 5px 10px; background: #2FA572; color: #fff; border: none; border-radius: 4px; cursor: pointer; }}
        .jump-box button:hover {{ background: #1F7A54; }}
        .result-badge {{ padding: 6px 12px; border-radius: 20px; font-size: 14px; font-weight: bold; color: #fff; max-width: 400px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        
        .container {{ max-width: 90%; margin: 80px auto 40px auto; background: #fff; padding: 40px 50px; box-shadow: 0 4px 20px rgba(0,0,0,0.05); border-radius: 8px; min-height: 70vh; }}
        
        /* 来源、列表、正文跳转信息栏 */
        .article-info {{ font-size: 13px; color: #666; margin-bottom: 20px; text-align: center; background: #f9f9f9; padding: 10px; border-radius: 6px; border: 1px dashed #ccc; }}
        .article-info a {{ color: #2FA572; text-decoration: none; margin: 0 8px; font-weight: bold; }}
        .article-info a:hover {{ text-decoration: underline; color: #1F7A54; }}
        
        .article-title {{ font-size: 22px; font-weight: bold; text-align: center; margin-bottom: 25px; color: #222; line-height: 1.4; }}
        
        /* 用户自定义排版样式 */
        .article-content {{ font-size: 14px; line-height: 1.4; color: #444; white-space: pre-wrap; word-wrap: break-word; }}
        br {{ display: none; }}
        
        .tips {{ text-align: center; color: #999; font-size: 13px; margin-top: 20px; }}
        kbd {{ background-color: #eee; border-radius: 3px; border: 1px solid #b4b4b4; box-shadow: 0 1px 1px rgba(0,0,0,.2); color: #333; display: inline-block; font-size: .85em; font-weight: 700; line-height: 1; padding: 2px 4px; white-space: nowrap; }}
    </style>
</head>
<body>
    <div class="header">
        <div class="progress-box">当前: <span id="current-idx">1</span> / <span id="total-idx">?</span> 篇</div>
        <div class="jump-box">
            跳转至 <input type="number" id="jump-input" min="1" value="1"> 
            <button onclick="jump()">GO</button>
        </div>
        <div id="result-badge" class="result-badge">状态加载中...</div>
    </div>

    <div class="container">
        <!-- 包含三个跳转链接的信息栏 -->
        <div class="article-info" id="article-info">链接加载中...</div>
        
        <div class="article-title" id="article-title">标题加载中...</div>
        <div class="article-content" id="article-content">正文加载中...</div>
    </div>
    
    <div class="tips">操作提示：按键盘 <kbd>←</kbd> 上一篇 ， 按 <kbd>→</kbd> 下一篇</div>

    <script>
        const articles = {json_data};
        let currentIndex = 0;
        
        document.getElementById('total-idx').innerText = articles.length;

        function render(index) {{
            if (index < 0) index = 0;
            if (index >= articles.length) index = articles.length - 1;
            currentIndex = index;

            const item = articles[currentIndex];

            document.getElementById('current-idx').innerText = currentIndex + 1;
            document.getElementById('jump-input').value = currentIndex + 1;
            
            // 组装上方的信息栏链接
            let main_link = item.main_url ? `<a href="${{item.main_url}}" target="_blank">首页: ${{item.site || "未知"}}</a>` : `<span>首页: ${{item.site || "未知"}}(无链接)</span>`;
            let list_link = item.list_url ? `<a href="${{item.list_url}}" target="_blank">📄 公告列表页</a>` : `<span style="color:#aaa;">📄 公告列表页(无)</span>`;
            let detail_link = item.detail_url ? `<a href="${{item.detail_url}}" target="_blank">🔗 当前文章页</a>` : `<span style="color:#aaa;">🔗 当前文章页(无)</span>`;
            
            document.getElementById('article-info').innerHTML = main_link + " | " + list_link + " | " + detail_link;

            document.getElementById('article-title').innerText = item.title || "无标题";
            document.getElementById('article-content').innerText = item.content || "【未提取到正文内容】";

            const result = item.result;
            const badge = document.getElementById('result-badge');
            badge.innerText = result;
            badge.title = result; 
            if (result === "正常") {{
                badge.style.backgroundColor = "#2FA572";
            }} else {{
                badge.style.backgroundColor = "#E74C3C";
            }}
            window.scrollTo(0, 0);
        }}

        document.addEventListener('keydown', function(event) {{
            if (event.key === "ArrowRight") render(currentIndex + 1);
            else if (event.key === "ArrowLeft") render(currentIndex - 1);
        }});

        function jump() {{
            let val = parseInt(document.getElementById('jump-input').value);
            if (!isNaN(val) && val > 0 && val <= articles.length) render(val - 1);
            else alert("请输入有效的页码！");
        }}

        document.getElementById('jump-input').addEventListener('keypress', function(e) {{
            if (e.key === 'Enter') jump();
        }});

        if(articles.length > 0) render(0);
    </script>
</body>
</html>
"""
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(html_template)

    def process_excel(self, path, keywords, min_words):
        try:
            self.log("读取 Excel 数据中，请稍候...")
            df = pd.read_excel(path)
            
            if len(df.columns) < 15:
                self.log("❌ [严重错误] Excel 列数不足！")
                self.finish_inspection()
                return

            out_site_names = []
            out_main_titles = []  
            out_results = []
            out_m_contents = []
            out_l_titles = []
            
            html_data_list = []
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

            for index, row in df.iterrows():
                row_num = index + 2 
                self.log(f"🔎 正在处理第 {row_num} 行数据...")
                
                val_A = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else "未知网站"
                val_C = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
                val_G = str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else ""
                val_L = str(row.iloc[11]).strip() if pd.notna(row.iloc[11]) else ""
                val_M = str(row.iloc[12]).strip() if pd.notna(row.iloc[12]) else ""
                val_N = str(row.iloc[13]).strip() if pd.notna(row.iloc[13]) else ""
                val_O = str(row.iloc[14]).strip() if pd.notna(row.iloc[14]) else ""

                row_errors = []
                l_title_extracted = ""
                detail_url_extracted = ""
                main_url = ""
                main_title = ""

                out_site_names.append(val_A)
                out_m_contents.append(val_M)

                # ================= 获取主站URL及首页Title =================
                if val_C and val_C.startswith("http"):
                    try:
                        parsed_url = urlparse(val_C)
                        main_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
                        
                        self.log(f"   > 正在抓取首页: {main_url}")
                        resp_main = requests.get(main_url, headers=headers, timeout=10, verify=False)
                        resp_main.encoding = resp_main.apparent_encoding
                        
                        if resp_main.status_code == 200:
                            soup_main = BeautifulSoup(resp_main.text, "html.parser")
                            title_tag = soup_main.find("title")
                            if title_tag and title_tag.string:
                                main_title = title_tag.string.strip()
                            else:
                                main_title = "[无Title标签]"
                        else:
                            main_title = "[获取失败]"
                    except Exception:
                        main_title = "[获取失败]"
                else:
                    main_title = "[C列URL不合法]"
                
                out_main_titles.append(main_title)

                # ================= L列解析 (含详情页URL绝对路径补全) =================
                if val_L:
                    try:
                        json_data = json.loads(val_L)
                        l_title_extracted = json_data.get("title", "").strip()
                        raw_detail_url = json_data.get("url", "").strip()
                        
                        # 智能补全相对路径为绝对路径
                        if raw_detail_url and val_C and not raw_detail_url.startswith("http"):
                            detail_url_extracted = urljoin(val_C, raw_detail_url)
                        else:
                            detail_url_extracted = raw_detail_url

                        if val_M and val_M != l_title_extracted:
                            row_errors.append(f"M列与L列标题不一致")
                    except json.JSONDecodeError:
                        l_title_extracted = "[JSON解析失败]"
                        row_errors.append("L列JSON格式无法解析")
                
                out_l_titles.append(l_title_extracted)

                # 空值校验
                empty_cols = []
                if not val_M: empty_cols.append("M列")
                if not val_N: empty_cols.append("N列")
                if not val_O: empty_cols.append("O列")
                if empty_cols:
                    row_errors.append(f"{','.join(empty_cols)}为空")

                # 违禁词校验
                if val_O:
                    found_kws = [kw for kw in keywords if kw in val_O]
                    if found_kws:
                        row_errors.append(f"包含屏蔽词:[{','.join(found_kws)}]")

                # 字数统计
                if val_O:
                    clean_text = re.sub(r'\s+', '', val_O)
                    if len(clean_text) < min_words:
                        row_errors.append(f"正文不足{min_words}字(当前{len(clean_text)}字)")

                # 探测列表页
                if not val_C or not val_C.startswith("http"):
                    row_errors.append("C列URL为空或不合法")
                elif not val_G:
                    row_errors.append("G列选择器为空")
                else:
                    try:
                        self.log(f"   > 探测列表选择器...")
                        resp = requests.get(val_C, headers=headers, timeout=10, verify=False)
                        resp.encoding = resp.apparent_encoding
                        if resp.status_code != 200:
                            row_errors.append(f"网页异常(HTTP {resp.status_code})")
                        else:
                            soup = BeautifulSoup(resp.text, "html.parser")
                            try:
                                elements = soup.select(val_G)
                                if len(elements) == 0:
                                    row_errors.append("未找到选择器(可能为动态页面)")
                            except Exception:
                                row_errors.append("G列非合法选择器")
                    except Exception:
                        row_errors.append("网页连接失败/超时")

                # 总结单行结果
                if len(row_errors) == 0:
                    final_res = "正常"
                else:
                    final_res = "[异常] " + " | ".join(row_errors)
                out_results.append(final_res)

                # ====== 将组装好的URL全部推入前端容器 ======
                html_data_list.append({
                    "site": val_A,
                    "title": val_M,
                    "content": val_O,
                    "result": final_res,
                    "main_url": main_url if main_title != "[C列URL不合法]" else "",
                    "list_url": val_C,
                    "detail_url": detail_url_extracted
                })

            # ================= 导出流程 =================
            self.log("正在生成精简版 Excel 结果表...")
            final_df = pd.DataFrame({
                "网站名称": out_site_names,
                "首页title": out_main_titles,
                "逻辑检查结果": out_results,
                "M列的title(标题预览)": out_m_contents,
                "L列的title(源数据)": out_l_titles
            })
            
            dir_name = os.path.dirname(path)
            base_name = os.path.basename(path)
            
            excel_new_name = base_name.replace(".xlsx", "_精简巡检结果.xlsx")
            excel_save_path = os.path.join(dir_name, excel_new_name)
            final_df.to_excel(excel_save_path, index=False)

            self.log("正在打包生成沉浸式 HTML 阅读器...")
            html_new_name = base_name.replace(".xlsx", "_沉浸阅读器.html")
            html_save_path = os.path.join(dir_name, html_new_name)
            self.generate_html_reader(html_save_path, html_data_list)
            
            self.log("="*40)
            self.log(f"🎉 全部任务圆满完成！")
            
            self.root.after(0, lambda: messagebox.showinfo("任务完成", 
                f"巡检已全部完成！\n\n1. 结果表：{excel_new_name}\n2. 阅读器：{html_new_name}\n\n建议直接双击打开 .html 文件进行沉浸式审查！"))

        except Exception as e:
            self.log(f"❌ [代码异常] {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("程序异常", f"发生致命错误：\n{str(e)}"))
        finally:
            self.finish_inspection()

    def finish_inspection(self):
        self.is_running = False
        def reset_btn():
            self.btn_start.configure(state="normal", text="▶ 开始执行巡检", fg_color="#2FA572")
        self.root.after(0, reset_btn)

if __name__ == "__main__":
    root = ctk.CTk()
    app = ExcelInspectorApp(root)
    root.mainloop()
