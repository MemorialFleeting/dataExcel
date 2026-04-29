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

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ExcelInspectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("数据源表质量巡检工具 v6.0 (团队协同分发版)")
        self.root.geometry("1000x700") # 稍微拉长窗口适应新配置框
        
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
        title_label.pack(pady=(15, 15))

        # 1. 文件选择
        ctk.CTkLabel(left_frame, text="1. 选择源表 (.xlsx)", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=20, pady=(0, 2))
        self.btn_select = ctk.CTkButton(left_frame, text="上传源表文件", command=self.select_file, height=35)
        self.btn_select.pack(fill="x", padx=20)
        self.lbl_file = ctk.CTkLabel(left_frame, text="未选择文件", text_color="gray", font=("微软雅黑", 12))
        self.lbl_file.pack(anchor="w", padx=20, pady=(2, 10))

        # 2. 人员分发名单 (新增)
        ctk.CTkLabel(left_frame, text="2. 人员名单 (对应R列, 换行分隔)", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=20, pady=(0, 2))
        self.txt_personnel = ctk.CTkTextbox(left_frame, height=90, font=("微软雅黑", 12))
        self.txt_personnel.pack(fill="x", padx=20, pady=(0, 10))
        default_personnel = "王明晓,史乃友,傅建文,姚凯莉,张伟,张欣,王蕊平,仲朋鹏,董海洋,李艳亮"
        self.txt_personnel.insert("0.0", default_personnel.replace(",", "\n"))

        # 3. 违禁词配置
        ctk.CTkLabel(left_frame, text="3. 屏蔽关键词配置", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=20, pady=(0, 2))
        self.txt_keywords = ctk.CTkTextbox(left_frame, height=110, font=("微软雅黑", 12))
        self.txt_keywords.pack(fill="x", padx=20, pady=(0, 10))
        default_keywords = "初审,一审,二审,三审,复审,终审,上一页,下一页,上一篇,下一篇,上页,下页,上篇,下篇,打印按钮,无障碍阅读,扫一扫,下载DOC,关闭窗口"
        self.txt_keywords.insert("0.0", default_keywords.replace(",", "\n"))

        # 4. 最小字数配置
        ctk.CTkLabel(left_frame, text="4. 正文最小字数", font=("微软雅黑", 14, "bold")).pack(anchor="w", padx=20, pady=(0, 2))
        self.ent_min_words = ctk.CTkEntry(left_frame, height=35)
        self.ent_min_words.insert(0, "50")
        self.ent_min_words.pack(fill="x", padx=20, pady=(0, 10))

        # 5. 执行按钮
        self.btn_start = ctk.CTkButton(left_frame, text="▶ 开始生成审核工作站", font=("微软雅黑", 15, "bold"), 
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
            self.log(f"✅ 已导入源表: {filename}")

    def start_inspection(self):
        if not self.file_path:
            messagebox.showwarning("提示", "请先在左侧上传源表文件！")
            return
        if self.is_running: return
            
        try:
            min_words = int(self.ent_min_words.get().strip())
        except ValueError:
            messagebox.showerror("错误", "最小字数必须是整数！")
            return

        kw_text = self.txt_keywords.get("0.0", "end").replace(",", "\n")
        keywords = [k.strip() for k in kw_text.split('\n') if k.strip()]

        p_text = self.txt_personnel.get("0.0", "end").replace(",", "\n")
        personnel_list = [p.strip() for p in p_text.split('\n') if p.strip()]

        if not personnel_list:
            messagebox.showwarning("提示", "审核人员名单不能为空！")
            return

        self.is_running = True
        self.btn_start.configure(state="disabled", text="⏳ 正在执行中...", fg_color="gray")
        self.log_area.configure(state="normal")
        self.log_area.delete("0.0", "end")
        self.log_area.configure(state="disabled")
        
        self.log("="*40)
        self.log("🚀 开始团队协同巡检与分发")
        self.log("="*40)
        
        threading.Thread(target=self.process_excel, args=(self.file_path, keywords, personnel_list, min_words), daemon=True).start()

    def generate_html_reader(self, save_path, html_data_list, personnel_list):
        """生成带团队门禁和隔离系统的工作站"""
        json_data = json.dumps(html_data_list, ensure_ascii=False)
        json_personnel = json.dumps(personnel_list, ensure_ascii=False)
        
        html_template = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>协同审核工作站</title>
    <style>
        body {{ margin: 0; padding: 0; background-color: #F0F2F5; font-family: 'Microsoft YaHei', sans-serif; color: #333; overflow: hidden; }}
        
        .header {{ position: fixed; top: 0; width: 100%; height: 60px; background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.05); display: flex; justify-content: space-between; align-items: center; padding: 0 30px; box-sizing: border-box; z-index: 1000; }}
        
        /* 审核人下拉框特化样式 */
        .reviewer-box {{ font-size: 16px; font-weight: bold; color: #2C3E50; display: flex; align-items: center; }}
        .reviewer-box select {{ margin-left: 10px; padding: 6px 12px; font-size: 15px; border-radius: 4px; border: 2px solid #2FA572; outline: none; font-weight: bold; color: #2FA572; cursor: pointer; }}
        
        .progress-box {{ font-size: 16px; font-weight: bold; color: #2FA572; display: none; }}
        .jump-box {{ display: none; align-items: center; }}
        .jump-box input {{ width: 50px; text-align: center; padding: 4px; border: 1px solid #ddd; border-radius: 4px; outline: none; margin: 0 5px; }}
        .jump-box button {{ padding: 5px 10px; background: #2FA572; color: #fff; border: none; border-radius: 4px; cursor: pointer; }}
        
        .export-btn {{ padding: 8px 20px; background: #E67E22; color: #fff; font-weight: bold; border: none; border-radius: 4px; cursor: pointer; font-size: 15px; opacity: 0.5; cursor: not-allowed; transition: all 0.3s; }}
        .export-btn.active {{ opacity: 1; cursor: pointer; }}
        .export-btn.active:hover {{ background: #D35400; }}

        /* 欢迎遮罩层 */
        .welcome-screen {{ display: flex; flex-direction: column; justify-content: center; align-items: center; height: calc(100vh - 60px); margin-top: 60px; background-color: #fff; }}
        .welcome-screen h1 {{ font-size: 28px; color: #333; margin-bottom: 10px; }}
        .welcome-screen p {{ font-size: 16px; color: #7f8c8d; }}

        .main-container {{ display: none; margin-top: 60px; height: calc(100vh - 60px); }}

        /* 左侧表单 */
        .sidebar {{ width: 350px; background: #fff; box-shadow: 2px 0 8px rgba(0,0,0,0.05); padding: 20px; overflow-y: auto; display: flex; flex-direction: column; }}
        .sidebar h3 {{ margin-top: 0; margin-bottom: 20px; font-size: 18px; color: #2C3E50; border-bottom: 2px solid #2FA572; padding-bottom: 10px; display: flex; justify-content: space-between; }}
        
        .form-group {{ margin-bottom: 18px; }}
        .form-group label {{ display: block; font-weight: bold; font-size: 13px; color: #555; margin-bottom: 5px; }}
        .form-group input, .form-group textarea, .form-group select {{ width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-family: inherit; font-size: 14px; outline: none; }}
        .form-group textarea {{ resize: vertical; height: 70px; }}
        
        .hint-text {{ font-size: 12px; color: #E74C3C; margin-top: 5px; display: none; }}
        .error-border {{ border: 2px solid #E74C3C !important; background-color: #FDEDEC !important; }}

        /* 右侧阅读区 */
        .content-area {{ flex: 1; padding: 30px; overflow-y: auto; background-color: #F4F6F8; position: relative; }}
        .article-card {{ max-width: 100%; margin: 0 auto; background: #fff; padding: 40px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); border-radius: 8px; min-height: 100%; box-sizing: border-box; }}
        
        .article-info {{ font-size: 13px; color: #666; margin-bottom: 20px; text-align: center; background: #f9f9f9; padding: 10px; border-radius: 6px; border: 1px dashed #ccc; }}
        .article-info a {{ color: #2FA572; text-decoration: none; margin: 0 8px; font-weight: bold; }}
        
        .article-title {{ font-size: 22px; font-weight: bold; text-align: center; margin-bottom: 25px; color: #222; line-height: 1.4; }}
        .article-content {{ font-size: 14px; line-height: 1.4; color: #444; white-space: pre-wrap; word-wrap: break-word; }}
        br {{ display: none; }}
        .tips {{ text-align: center; color: #999; font-size: 13px; margin-top: 30px; }}
        kbd {{ background-color: #eee; border-radius: 3px; border: 1px solid #b4b4b4; padding: 2px 4px; font-weight: bold; }}
    </style>
</head>
<body>
    <div class="header">
        <!-- 身份门禁下拉框 -->
        <div class="reviewer-box">
            身份认证 
            <select id="reviewer-select" onchange="onReviewerChange()">
                <option value="">-- 请选择您的姓名 --</option>
            </select>
        </div>
        
        <div class="progress-box" id="progress-box">当前阅读: <span id="current-idx">1</span> / <span id="total-idx">?</span> 篇</div>
        <div class="jump-box" id="jump-box">
            跳转至 <input type="number" id="jump-input" min="1" value="1"> 
            <button onclick="jump()">GO</button>
        </div>
        
        <!-- 导出按钮默认禁用 -->
        <button class="export-btn" id="btn-export" onclick="exportToCSV()" disabled>📥 导出我的数据</button>
    </div>

    <!-- 门禁遮罩层 -->
    <div class="welcome-screen" id="welcome-screen">
        <h1>👋 欢迎来到协同审核工作站</h1>
        <p>请先在左上角下拉菜单中选择您的姓名，系统将自动下发您的专属审核任务。</p>
    </div>

    <div class="main-container" id="main-container">
        <!-- 左侧表单 -->
        <div class="sidebar">
            <h3>📝 人工审核 <span style="font-size:12px; color:#999; font-weight:normal;" id="my-name-tag"></span></h3>
            <div class="form-group">
                <label>1. 网站名称</label>
                <input type="text" id="f_site" oninput="updateData()">
            </div>
            <div class="form-group">
                <label>2. 网站类型</label>
                <input type="text" id="f_type" oninput="updateData()">
                <div class="hint-text" id="h_type">⚠️ 网站名称包含特殊信息</div>
            </div>
            <div class="form-group">
                <label>3. 名称是否正确</label>
                <select id="f_name_ok" onchange="handleNameOkChange()">
                    <option value="是">是</option>
                    <option value="否">否</option>
                </select>
            </div>
            <div class="form-group">
                <label>4. 名称修正</label>
                <input type="text" id="f_name_fix" oninput="updateData()">
            </div>
            <div class="form-group">
                <label>5. 选择器是否正确</label>
                <select id="f_selector" onchange="handleSelectorOkChange()">
                    <option value="是">是</option>
                    <option value="否">否</option>
                </select>
            </div>
            <div class="form-group">
                <label>6. 备注</label>
                <textarea id="f_remark" oninput="updateData()"></textarea>
                <div class="hint-text" id="h_remark" style="color: #E67E22;">🔍 需要人为查验填写</div>
            </div>
        </div>

        <!-- 右侧阅读 -->
        <div class="content-area" id="scroll-area">
            <div class="article-card">
                <div class="article-info" id="article-info">链接加载中...</div>
                <div class="article-title" id="article-title">标题加载中...</div>
                <div class="article-content" id="article-content">正文加载中...</div>
                <div class="tips">操作提示：编辑完左侧表单后，按键盘 <kbd>←</kbd> 上一篇 ， 按 <kbd>→</kbd> 下一篇</div>
            </div>
        </div>
    </div>

    <script>
        const allArticles = {json_data};
        const personnelList = {json_personnel};
        
        let filteredArticles = [];
        let currentIndex = 0;
        let currentReviewer = "";

        // 1. 初始化下拉菜单
        const selectEl = document.getElementById('reviewer-select');
        personnelList.forEach(p => {{
            let opt = document.createElement('option');
            opt.value = p;
            opt.innerText = p;
            selectEl.appendChild(opt);
        }});

        // 2. 身份切换逻辑 (核心拦截)
        function onReviewerChange() {{
            currentReviewer = selectEl.value;
            const welcome = document.getElementById('welcome-screen');
            const mainBox = document.getElementById('main-container');
            const progBox = document.getElementById('progress-box');
            const jumpBox = document.getElementById('jump-box');
            const expBtn  = document.getElementById('btn-export');

            if (!currentReviewer) {{
                welcome.style.display = 'flex';
                mainBox.style.display = 'none';
                progBox.style.display = 'none';
                jumpBox.style.display = 'none';
                expBtn.disabled = true;
                expBtn.classList.remove('active');
                return;
            }}

            // 数据过滤，提取当前审核人的专属数据 (对象引用传递，修改联动生效)
            filteredArticles = allArticles.filter(item => item.assignee === currentReviewer);

            if (filteredArticles.length === 0) {{
                welcome.innerHTML = `<h1>🎉 恭喜！</h1><p>当前系统内没有分配给 [${{currentReviewer}}] 的审核任务，您可以休息了。</p>`;
                welcome.style.display = 'flex';
                mainBox.style.display = 'none';
                progBox.style.display = 'none';
                jumpBox.style.display = 'none';
                expBtn.disabled = true;
                expBtn.classList.remove('active');
            }} else {{
                welcome.style.display = 'none';
                mainBox.style.display = 'flex';
                progBox.style.display = 'block';
                jumpBox.style.display = 'flex';
                expBtn.disabled = false;
                expBtn.classList.add('active');
                
                document.getElementById('my-name-tag').innerText = `(操作人: ${{currentReviewer}})`;
                document.getElementById('total-idx').innerText = filteredArticles.length;
                render(0);
            }}
        }}

        // 下列所有的渲染和保存，全部依赖于 filteredArticles 而不是 allArticles
        function updateData() {{
            if(filteredArticles.length === 0) return;
            let item = filteredArticles[currentIndex].audit;
            item.site = document.getElementById('f_site').value;
            item.type = document.getElementById('f_type').value;
            item.name_ok = document.getElementById('f_name_ok').value;
            item.name_fix = document.getElementById('f_name_fix').value;
            item.selector = document.getElementById('f_selector').value;
            item.remark = document.getElementById('f_remark').value;
            updateUIStyles();
        }}

        function handleNameOkChange() {{
            if (document.getElementById('f_name_ok').value === '是') document.getElementById('f_name_fix').value = '';
            updateData();
        }}

        function handleSelectorOkChange() {{
            if (document.getElementById('f_selector').value === '是') document.getElementById('f_remark').value = '';
            updateData();
        }}

        function updateUIStyles() {{
            const elNameOk = document.getElementById('f_name_ok');
            const elNameFix = document.getElementById('f_name_fix');
            const elSelectorOk = document.getElementById('f_selector');
            const elRemark = document.getElementById('f_remark');
            
            elNameOk.value === '否' ? elNameOk.classList.add('error-border') : elNameOk.classList.remove('error-border');
            elSelectorOk.value === '否' ? elSelectorOk.classList.add('error-border') : elSelectorOk.classList.remove('error-border');
            elNameFix.value.trim() !== '' ? elNameFix.classList.add('error-border') : elNameFix.classList.remove('error-border');
            elRemark.value.trim() !== '' ? elRemark.classList.add('error-border') : elRemark.classList.remove('error-border');

            const item = filteredArticles[currentIndex];
            let displaySite = document.getElementById('f_site').value || "未知";
            let main_link = item.main_url ? `<a href="${{item.main_url}}" target="_blank">首页: ${{displaySite}}</a>` : `<span>首页: ${{displaySite}}(无)</span>`;
            let list_link = item.list_url ? `<a href="${{item.list_url}}" target="_blank">📄 列表页</a>` : `<span style="color:#aaa;">📄 列表页(无)</span>`;
            let detail_link = item.detail_url ? `<a href="${{item.detail_url}}" target="_blank">🔗 文章页</a>` : `<span style="color:#aaa;">🔗 文章页(无)</span>`;
            document.getElementById('article-info').innerHTML = main_link + " | " + list_link + " | " + detail_link;
        }}

        function render(index) {{
            if (index < 0) index = 0;
            if (index >= filteredArticles.length) index = filteredArticles.length - 1;
            currentIndex = index;
            const item = filteredArticles[currentIndex];

            document.getElementById('current-idx').innerText = currentIndex + 1;
            document.getElementById('jump-input').value = currentIndex + 1;
            
            document.getElementById('f_site').value = item.audit.site;
            document.getElementById('f_type').value = item.audit.type;
            document.getElementById('f_name_ok').value = item.audit.name_ok;
            document.getElementById('f_name_fix').value = item.audit.name_fix;
            document.getElementById('f_selector').value = item.audit.selector;
            document.getElementById('f_remark').value = item.audit.remark;
            
            let f_type = document.getElementById('f_type');
            if (item.flags.is_special) {{
                f_type.classList.add("error-border");
                document.getElementById('h_type').style.display = "block";
            }} else {{
                f_type.classList.remove("error-border");
                document.getElementById('h_type').style.display = "none";
            }}

            if (item.flags.has_error) document.getElementById('h_remark').style.display = "block";
            else document.getElementById('h_remark').style.display = "none";

            document.getElementById('article-title').innerText = item.title || "无标题";
            document.getElementById('article-content').innerText = item.content || "【未提取到正文】";

            updateUIStyles();
            document.getElementById('scroll-area').scrollTop = 0;
        }}

        document.addEventListener('keydown', function(event) {{
            if (event.target.tagName === 'INPUT' || event.target.tagName === 'TEXTAREA' || event.target.tagName === 'SELECT') return;
            if (event.key === "ArrowRight") render(currentIndex + 1);
            else if (event.key === "ArrowLeft") render(currentIndex - 1);
        }});

        function jump() {{
            let val = parseInt(document.getElementById('jump-input').value);
            if (!isNaN(val) && val > 0 && val <= filteredArticles.length) render(val - 1);
            else alert("请输入有效的页码！");
        }}

        document.getElementById('jump-input').addEventListener('keypress', function(e) {{
            if (e.key === 'Enter') jump();
        }});

        // === 隔离导出机制：只导出这个人名下的数据 ===
        function exportToCSV() {{
            let headers = ["网站名称", "网站类型", "名称是否正确", "名称修正", "选择器是否正确", "备注", "审核归属人"];
            let csvContent = "\\uFEFF" + headers.join(",") + "\\n";

            // 循环遍历 filteredArticles 而非全部数据
            filteredArticles.forEach(item => {{
                let row = [
                    item.audit.site, item.audit.type, item.audit.name_ok,
                    item.audit.name_fix, item.audit.selector, item.audit.remark,
                    item.assignee
                ];
                let escapedRow = row.map(field => `"${{String(field || "").replace(/"/g, '""')}}"`);
                csvContent += escapedRow.join(",") + "\\n";
            }});

            let blob = new Blob([csvContent], {{ type: 'text/csv;charset=utf-8;' }});
            let link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            // 导出文件名直接打上这个人的人名标签
            let dateStr = new Date().toISOString().slice(0,10).replace(/-/g,"");
            link.download = `人工审核结果_${{currentReviewer}}_${{dateStr}}.csv`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}
    </script>
</body>
</html>
"""
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(html_template)

    def process_excel(self, path, keywords, personnel_list, min_words):
        try:
            self.log("读取 Excel 数据中，请稍候...")
            df = pd.read_excel(path)
            
            # 判断是否包含18列 (下标0~17)
            if len(df.columns) < 18:
                self.log("❌ [严重错误] Excel 列数不足18列，无法读取R列(人员分工)！")
                self.finish_inspection()
                return

            html_data_list = []
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

            for index, row in df.iterrows():
                row_num = index + 2 
                
                # ================= R列人员拦截系统 =================
                val_R = str(row.iloc[17]).strip() if pd.notna(row.iloc[17]) else ""
                if not val_R or val_R not in personnel_list:
                    # 如果R列为空或者不在我们配置的人员名单里，直接丢弃这条数据
                    continue
                
                self.log(f"🔎 分配处理 [行{row_num} -> {val_R}] ...")
                
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

                if val_C and val_C.startswith("http"):
                    try:
                        parsed_url = urlparse(val_C)
                        main_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
                        
                        resp_main = requests.get(main_url, headers=headers, timeout=10, verify=False)
                        resp_main.encoding = resp_main.apparent_encoding
                        if resp_main.status_code == 200:
                            soup_main = BeautifulSoup(resp_main.text, "html.parser")
                            title_tag = soup_main.find("title")
                            main_title = title_tag.string.strip() if title_tag and title_tag.string else "[无Title标签]"
                        else:
                            main_title = "[获取失败]"
                    except Exception:
                        main_title = "[获取失败]"
                else:
                    main_title = "[C列URL不合法]"

                if val_L:
                    try:
                        json_data = json.loads(val_L)
                        l_title_extracted = json_data.get("title", "").strip()
                        raw_detail_url = json_data.get("url", "").strip()
                        if raw_detail_url and val_C and not raw_detail_url.startswith("http"):
                            detail_url_extracted = urljoin(val_C, raw_detail_url)
                        else:
                            detail_url_extracted = raw_detail_url

                        if val_M and val_M != l_title_extracted:
                            row_errors.append(f"M列与L列标题不一致")
                    except json.JSONDecodeError:
                        row_errors.append("L列JSON格式无法解析")

                empty_cols = []
                if not val_M: empty_cols.append("M列")
                if not val_N: empty_cols.append("N列")
                if not val_O: empty_cols.append("O列")
                if empty_cols: row_errors.append(f"{','.join(empty_cols)}为空")

                if val_O:
                    found_kws = [kw for kw in keywords if kw in val_O]
                    if found_kws: row_errors.append(f"包含屏蔽词:[{','.join(found_kws)}]")

                if val_O:
                    clean_text = re.sub(r'\s+', '', val_O)
                    if len(clean_text) < min_words:
                        row_errors.append(f"正文不足{min_words}字(当前{len(clean_text)}字)")

                if not val_C or not val_C.startswith("http"):
                    row_errors.append("C列URL为空")
                elif not val_G:
                    row_errors.append("G列选择器为空")
                else:
                    try:
                        resp = requests.get(val_C, headers=headers, timeout=10, verify=False)
                        resp.encoding = resp.apparent_encoding
                        if resp.status_code != 200:
                            row_errors.append(f"网页异常({resp.status_code})")
                        else:
                            soup = BeautifulSoup(resp.text, "html.parser")
                            try:
                                elements = soup.select(val_G)
                                if len(elements) == 0: row_errors.append("未找到选择器")
                            except Exception:
                                row_errors.append("G列非合法选择器")
                    except Exception:
                        row_errors.append("网页连接超时")

                final_res = "正常" if len(row_errors) == 0 else " | ".join(row_errors)

                is_special = "人才" in val_A or "聚合" in val_A
                name_is_ok = "是" if val_A == main_title else "否"
                name_fix_val = "" if val_A == main_title else main_title
                selector_is_ok = "是" if final_res == "正常" else "否"
                remark_val = "" if final_res == "正常" else final_res

                html_data_list.append({
                    "assignee": val_R,  # ====== 新增字段：分配的审核人 ======
                    "title": val_M,
                    "content": val_O,
                    "main_url": main_url if main_title != "[C列URL不合法]" else "",
                    "list_url": val_C,
                    "detail_url": detail_url_extracted,
                    "flags": {
                        "is_special": is_special,
                        "has_error": final_res != "正常"
                    },
                    "audit": {
                        "site": val_A,
                        "type": "源站",
                        "name_ok": name_is_ok,
                        "name_fix": name_fix_val,
                        "selector": selector_is_ok,
                        "remark": remark_val
                    }
                })

            dir_name = os.path.dirname(path)
            base_name = os.path.basename(path)
            
            self.log("正在打包协同审核工作站 (HTML)...")
            html_new_name = base_name.replace(".xlsx", "_团队协同工作站.html")
            html_save_path = os.path.join(dir_name, html_new_name)
            
            # 将名单传入 HTML 生成函数
            self.generate_html_reader(html_save_path, html_data_list, personnel_list)
            
            self.log("="*40)
            self.log(f"🎉 全部任务圆满完成！有效分发数据共 {len(html_data_list)} 条。")
            
            self.root.after(0, lambda: messagebox.showinfo("任务完成", 
                f"团队协同工作站已生成！\n\n请将该 HTML 文件发给审核人员：\n{html_new_name}\n\n每个人选中自己名字后，即可开展隔离审核。"))

        except Exception as e:
            self.log(f"❌ [代码异常] {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("程序异常", f"发生致命错误：\n{str(e)}"))
        finally:
            self.finish_inspection()

    def finish_inspection(self):
        self.is_running = False
        def reset_btn():
            self.btn_start.configure(state="normal", text="▶ 开始生成审核工作站", fg_color="#2FA572")
        self.root.after(0, reset_btn)

if __name__ == "__main__":
    root = ctk.CTk()
    app = ExcelInspectorApp(root)
    root.mainloop()
