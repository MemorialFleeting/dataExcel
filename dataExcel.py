import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import threading
import os
import datetime

class CSVSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("超大 CSV 分校拆分工具")
        self.root.geometry("850x600")
        self.root.configure(padx=10, pady=10)
        
        # 默认分校数据 (使用英文逗号)
        self.default_groups = [
            "山东分校,广东分校,河南分校,河北分校,湖北分校",
            "吉林分校,山西分校,陕西分校,安徽分校,辽宁分校,云南分校",
            "江苏分校,湖南分校,四川分校,黑龙江分校,广西分校,新疆分校,浙江分校,江西分校,北京分校,内蒙古分校",
            "贵州分校,福建分校,甘肃分校,海南分校,宁夏分校,青海分校,厦门分校,上海分校,天津分校,西藏分校,重庆分校"
        ]
        
        self.text_inputs = []
        self.file_path = tk.StringVar()
        
        self.create_widgets()

    def create_widgets(self):
        # 使用 PanedWindow 实现左右分栏
        paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)

        # ================= 左侧：操作区 =================
        left_frame = ttk.Frame(paned_window, padding=(10, 0, 10, 0))
        paned_window.add(left_frame, weight=3)

        # 1. 规则输入区
        ttk.Label(left_frame, text="分校配置 (请务必使用 英文逗号 隔开):", font=("微软雅黑", 10, "bold")).pack(pady=(0, 10), anchor="w")
        
        for i in range(4):
            frame = ttk.Frame(left_frame)
            frame.pack(fill="x", pady=5)
            ttk.Label(frame, text=f"表 {i+1}:", width=5).pack(side="left", anchor="n", pady=2)
            
            # 使用 tk.Text 以支持多行，但应用更清爽的边框
            text_box = tk.Text(frame, height=3, width=40, font=("微软雅黑", 9), relief="solid", borderwidth=1)
            text_box.insert("1.0", self.default_groups[i])
            text_box.pack(side="left", fill="x", expand=True, padx=(5, 0))
            self.text_inputs.append(text_box)

        # 2. 文件选择区
        file_frame = ttk.Frame(left_frame)
        file_frame.pack(fill="x", pady=20)
        
        self.select_btn = ttk.Button(file_frame, text="1. 选择 CSV 源文件", command=self.select_file)
        self.select_btn.pack(side="left")
        
        ttk.Entry(file_frame, textvariable=self.file_path, state="readonly").pack(side="left", fill="x", expand=True, padx=(10, 0))

        # 3. 操作区
        self.process_btn = ttk.Button(left_frame, text="2. 开始自动拆分", command=self.start_processing)
        self.process_btn.pack(fill="x", pady=10, ipady=5)

        # ================= 右侧：日志区 =================
        right_frame = ttk.Frame(paned_window, padding=(10, 0, 0, 0))
        paned_window.add(right_frame, weight=2)

        ttk.Label(right_frame, text="运行记录:", font=("微软雅黑", 10, "bold")).pack(anchor="w", pady=(0, 5))
        
        # 日志文本框与滚动条
        log_scroll = ttk.Scrollbar(right_frame)
        log_scroll.pack(side="right", fill="y")
        
        self.log_text = tk.Text(right_frame, font=("Consolas", 9), bg="#f4f4f4", yscrollcommand=log_scroll.set, state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.config(command=self.log_text.yview)
        
        self.log_message("系统初始化完成。等待选择文件...")

    def log_message(self, message):
        """向右侧日志区添加带时间戳的记录"""
        now = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"[{now}] {message}\n")
        self.log_text.see(tk.END) # 自动滚动到最新
        self.log_text.config(state="disabled")

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV 数据文件", "*.csv")])
        if path:
            self.file_path.set(path)
            self.log_message(f"已选择源文件: {os.path.basename(path)}")
            self.log_message(f"保存目录默认设为: {os.path.dirname(path)}")

    def parse_input(self, text):
        # 强制将可能误输入的中文逗号替换为英文逗号，增加容错率
        text = text.replace('，', ',').replace('\n', ',')
        return [s.strip() for s in text.split(',') if s.strip()]

    def start_processing(self):
        if not self.file_path.get():
            messagebox.showwarning("提示", "请先选择 CSV 源文件！")
            return
        
        self.process_btn.config(state="disabled", text="处理中...")
        self.select_btn.config(state="disabled")
        
        # 开启新线程处理
        threading.Thread(target=self.process_csv).start()

    def process_csv(self):
        start_time = datetime.datetime.now()
        self.log_message("-" * 30)
        self.log_message("任务开始！正在读取巨型 CSV 文件...")
        
        try:
            file_path = self.file_path.get()
            output_dir = os.path.dirname(file_path) # 默认保存在源文件同目录
            
            # 读取 CSV
            df = pd.read_csv(file_path, low_memory=False)

            if len(df.columns) < 3:
                raise ValueError("表格列数不足 3 列，无法按照第三列拆分！")
            
            target_col = df.columns[2]
            self.log_message(f"成功加载数据，共 {len(df)} 行。")
            self.log_message(f"依据第 3 列 [{target_col}] 进行拆分...")

            # 获取用户输入的分组规则
            groups = [self.parse_input(box.get("1.0", tk.END)) for box in self.text_inputs]

            # 开始切分并保存
            for i, group_schools in enumerate(groups):
                filtered_df = df[df[target_col].isin(group_schools)]
                
                if not filtered_df.empty:
                    out_name = f"表{i+1}_切分结果.csv"
                    out_path = os.path.join(output_dir, out_name)
                    # 使用 utf-8-sig 编码，确保用 Excel 打开 CSV 时不会乱码
                    filtered_df.to_csv(out_path, index=False, encoding='utf-8-sig')
                    self.log_message(f"✔ 生成: {out_name} (匹配到 {len(filtered_df)} 条数据)")
                else:
                    self.log_message(f"⚠ 表{i+1}: 未匹配到任何对应分校的数据，跳过生成。")

            end_time = datetime.datetime.now()
            cost_time = (end_time - start_time).seconds
            self.log_message(f"任务结束！总耗时: {cost_time} 秒。")
            self.log_message("-" * 30)
            
            self.root.after(0, self.processing_complete, "拆分成功！文件已保存在源文件同级目录下。")
            
        except Exception as e:
            self.log_message(f"❌ 发生错误: {str(e)}")
            self.root.after(0, self.processing_complete, f"处理失败，请查看右侧日志。", error=True)

    def processing_complete(self, message, error=False):
        self.process_btn.config(state="normal", text="2. 开始自动拆分")
        self.select_btn.config(state="normal")
        if error:
            messagebox.showerror("错误", message)
        else:
            messagebox.showinfo("完成", message)

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVSplitterApp(root)
    root.mainloop()
