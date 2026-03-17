import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
import os
import threading
import time
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# ==============================================================================
# 配置区域
# ==============================================================================
FIXED_ORDER_CITIES = [
    "新乡", "沧州", "松原", "临汾", "汉中", "朝阳", "徐州",
    "海淀", "怀柔", "顺义", "房山", "大兴", "昌平", "平谷", "通州", "延庆", "密云",
    "包头", "通辽", "赤峰",
    "天津", "武清", "静海", "滨海"
]

DEFAULT_BRANCH_MAP = {
    "新乡": "河南分校", "沧州": "河北分校", "松原": "吉林分校", "临汾": "山西分校",
    "汉中": "陕西分校", "朝阳": "辽宁分校", "徐州": "江苏分校",
    "海淀": "北京分校", "怀柔": "北京分校", "顺义": "北京分校", "房山": "北京分校",
    "大兴": "北京分校", "昌平": "北京分校", "平谷": "北京分校", "通州": "北京分校",
    "延庆": "北京分校", "密云": "北京分校",
    "包头": "内蒙古分校", "通辽": "内蒙古分校", "赤峰": "内蒙古分校",
    "天津": "天津分校", "武清": "天津分校", "静海": "天津分校", "滨海": "天津分校"
}

# ==============================================================================
# 核心处理类
# ==============================================================================
class UniversalProcessor:
    def __init__(self, log_func):
        self.log = log_func
        self.use_special_city = False

    def excel_lookup_find(self, text, keywords, results, default_value):
        if pd.isna(text): return default_value
        text = str(text)
        for k, r in zip(reversed(keywords), reversed(results)):
            if k in text: return r
        return default_value

    def check_keyword_flag(self, text, keywords):
        if pd.isna(text): return ""
        text = str(text)
        for k in keywords:
            if k in text: return 1
        return ""

    def extract_city_smart(self, dept_path):
        if pd.isna(dept_path): return "未知"
        text_str = str(dept_path).replace('\r', '\n')
        lines = [line.strip() for line in text_str.split('\n') if line.strip()]
        if not lines: return "其他"
        target_line = ""
        for line in lines:
            if "地市分校" in line: target_line = line; break
        if not target_line:
            for line in lines:
                if "分校" in line: target_line = line; break
        if not target_line: target_line = lines[0]
        clean_line = target_line.replace('"', '').replace("'", "").replace("华图教育", "").strip()
        parts = [p.strip() for p in clean_line.split('/') if p.strip()]
        if not parts: return "其他"
        final_city = ""
        for part in parts:
            if "地市分校" in part:
                if part == "地市分校": continue 
                final_city = part.replace("地市分校", "")
                break
        if not final_city:
            for i, part in enumerate(parts):
                if "分校" in part and len(part) > 2:
                    if i + 1 < len(parts):
                        target = parts[i+1]
                        if target in ["地市分校", "各校区"]:
                            if i + 2 < len(parts): target = parts[i+2]
                            else: continue
                        final_city = target
                        break
        if not final_city: final_city = parts[-1]
        suffixes = ["区属学习中心", "高校学习中心", "学习中心", "办事处", "地市分校", "分校", "分部"]
        for suffix in suffixes:
            if final_city.endswith(suffix): final_city = final_city[:-len(suffix)]
        return final_city

    def standardize_city_name(self, city_name):
        if pd.isna(city_name): return "其他"
        s = str(city_name).strip()
        s = s.replace("_x000D_", "").replace("_x000d_", "")
        s = re.sub(r'[\s\r\n\t\u200b\ufeff\xa0]+', '', s)
        if len(s) > 2 and s.endswith("市"): s = s[:-1]
        return s

    def process_step1(self, df, date_keyword, custom_total_kws_str):
        self.log("【数据清洗】正在提取地市与计算标签...")
        df.columns = [str(c).replace('\ufeff', '').strip() for c in df.columns]
        
        branch_keywords = ["通辽分校", "陕西分校", "湖北分校", "辽宁分校", "河北分校", "甘肃分校", "厦门分校", "福建分校", "山东分校", "北京分校", "安徽分校", "黑龙江分校", "吉林分校", "江苏分校", "重庆分校", "广东分校", "天津分校", "河南分校", "云南分校", "江西分校", "湖南分校", "贵州分校", "广西分校", "山西分校", "宁夏分校", "内蒙古分校", "浙江分校", "新疆分校", "青海分校", "南疆分校", "四川分校", "上海分校", "海南分校", "西藏分校", "赤峰"]
        branch_results = ["内蒙古分校", "陕西分校", "湖北分校", "辽宁分校", "河北分校", "甘肃分校", "厦门分校", "福建分校", "山东分校", "北京分校", "安徽分校", "黑龙江分校", "吉林分校", "江苏分校", "重庆分校", "广东分校", "天津分校", "河南分校", "云南分校", "江西分校", "湖南分校", "贵州分校", "广西分校", "山西分校", "宁夏分校", "内蒙古分校", "浙江分校", "新疆分校", "青海分校", "新疆分校", "四川分校", "上海分校", "海南分校", "西藏分校", "内蒙古分校"]
        df['分校'] = df['负责人所属部门'].apply(lambda x: self.excel_lookup_find(x, branch_keywords, branch_results, "总部"))
        df['地市'] = df['负责人所属部门'].apply(self.extract_city_smart)
        df['地市'] = df['地市'].apply(self.standardize_city_name)

        def determine_category(row):
            dept_raw = str(row['负责人所属部门']) if pd.notna(row['负责人所属部门']) else ""
            branch = str(row['分校'])
            if "地市分校" in dept_raw: return "地市"
            if branch in ["北京分校", "天津分校"] and "各校区" in dept_raw: return "地市"
            return "其它"
        df['所属类别'] = df.apply(determine_category, axis=1)

        df['公职'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["26公职","27公职","28公职","29公职"]))
        kw_shiye = ["26事业","26三支","26社区","26辅警","26书记员","26国企","30事业","30三支","30社区","30辅警","30书记员","30国企", "27事业","27三支","27社区","27辅警","27书记员","27国企","28事业","28三支","28社区","28辅警","28书记员","28国企", "29事业","29三支","29社区","29辅警","29书记员","29国企"]
        df['事业辅助列'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, kw_shiye))
        kw_jiaoshi = ["26教师","26特岗","26教资","30教师","30特岗","30教资","27教师","27特岗","27教资","28教师","28特岗","28教资","29教师","29特岗","29教资"]
        df['教师'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, kw_jiaoshi))
        df['文职'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30文职","26文职","27文职","28文职","29文职"]))
        df['医疗'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30医疗","26医疗","27医疗","28医疗","29医疗"]))
        df['银行'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30银行","26银行","27银行","28银行","29银行"]))
        df['考研'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30考研","26考研","27考研","28考研","29考研"]))
        df['学历'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30学历","26学历","27学历","28学历","29学历"]))
        df['其他'] = ""
        
        custom_kws = [k.strip() for k in custom_total_kws_str.replace('，', ',').split(',') if k.strip()]
        if custom_kws:
            def check_custom_total(remark):
                if pd.isna(remark): return 0
                return 1 if any(kw in str(remark) for kw in custom_kws) else 0
            df['标备总数'] = df['备注名'].apply(check_custom_total)
        else:
            all_cols_check = ['公职', '事业辅助列', '教师', '文职', '医疗', '银行', '考研', '学历', '其他']
            df['标备总数'] = df.apply(lambda row: 1 if any(row[c] in [1, "1"] for c in all_cols_check) else 0, axis=1)

        date_kws = [k.strip() for k in date_keyword.replace('，', ',').split(',') if k.strip()]
        if not date_kws: date_kws = ["FAIL_SAFE"]
        
        df['！！！是否新增'] = df['创建时间'].apply(lambda x: "是" if any(k in str(x) for k in date_kws) else "否")

        kw_ch = ["网络","社会","高校","线上","线上平台","线上活动","现场","线下活动"]
        res_ch = ["线上平台","线下活动","高校","线上平台","线上平台","线上平台","考试现场","线下活动"]
        df['渠道'] = df['渠道活码分组'].apply(lambda x: self.excel_lookup_find(x, kw_ch, res_ch, "其他"))
        
        kw_on = ["网站","公众号","小红书","视频号","抖音","文章页","网页","专题","附件","小程序"]
        res_on = ["网站","公众号","小红书","视频号","抖音","网站","网站","网站","网站","其他"]
        df['线上渠道'] = df['渠道活码分组'].apply(lambda x: self.excel_lookup_find(x, kw_on, res_on, "其他"))

        df['事业2'] = df.apply(lambda row: 1 if row['公职'] not in [1, "1"] and row['事业辅助列'] in [1, "1"] else 0, axis=1)

        def check_reply(row):
            has_reply = pd.notna(row['客户上次回复时间']) and str(row['客户上次回复时间']).strip() != ""
            return "是" if has_reply and str(row['标备总数']) == "1" else "否"
        df['客户回话'] = df.apply(check_reply, axis=1)

        def fmt_date(val):
            try: return f"{pd.to_datetime(val).month}月{pd.to_datetime(val).day}日"
            except: return str(val)
        df['日期'] = df['好友添加时间'].apply(fmt_date)

        has_status_col = '添加好友状态' in df.columns
        col_z_name = df.columns[25] if len(df.columns) > 25 else None
        
        ext_col = 'ExternalUserId'
        if ext_col not in df.columns and len(df.columns) > 1:
            ext_col = df.columns[1] 
        if ext_col in df.columns and ext_col != 'ExternalUserId':
            df.rename(columns={ext_col: 'ExternalUserId'}, inplace=True)
            if ext_col == col_z_name: col_z_name = 'ExternalUserId'
        
        # 提取 5 个质检专属字段
        qc_required_cols = ['好友添加来源', '添加渠道码', '员工发送消息数', '客户回复消息数', '是否同意会话存档']
        qc_missing = [c for c in qc_required_cols if c not in df.columns]
        if qc_missing:
            for m in qc_missing:
                if '数' in m: df[m] = 0
                else: df[m] = "未知"

        final_cols = ['分校', '地市', '所属类别', '公职', '事业2', '教师', '文职', '医疗', '银行', '考研', '学历', '其他', '标备总数', '！！！是否新增', '渠道', '线上渠道', '客户回话', '日期']
        if '备注名' in df.columns: final_cols.insert(0, '备注名')
        if col_z_name and col_z_name not in final_cols: final_cols.insert(1, col_z_name)
        if has_status_col: final_cols.append('添加好友状态')
        
        if 'ExternalUserId' not in final_cols: final_cols.append('ExternalUserId')
        for qc_col in qc_required_cols:
            if qc_col not in final_cols: final_cols.append(qc_col)

        df_clean = df[[c for c in final_cols if c in df.columns]].copy()
        
        for col in ['公职', '事业2', '标备总数', '员工发送消息数', '客户回复消息数']:
            if col in df_clean.columns:
                df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

        df_dedup_prov = df_clean.drop_duplicates(subset=[col_z_name, '分校'], keep='first') if col_z_name else df_clean.copy()
        df_dedup_city = df_clean.drop_duplicates(subset=[col_z_name, '分校', '地市'], keep='first') if col_z_name else df_clean.copy()
        
        if has_status_col:
            df_retention = df_clean[df_clean['添加好友状态'].astype(str).str.contains("已添加", na=False)].copy()
        else:
            df_retention = pd.DataFrame()

        return df_clean, df_dedup_prov, df_dedup_city, df_retention

    def calc_stats(self, df_raw, df_dedup):
        raw_cnt = len(df_raw)
        dedup_cnt = len(df_dedup)
        for c in ['公职', '事业2', '标备总数']:
            if c in df_dedup.columns: df_dedup[c] = pd.to_numeric(df_dedup[c], errors='coerce').fillna(0)
        
        new_df = df_dedup[df_dedup['！！！是否新增'].astype(str).str.contains("是", na=False)]
        new_add = len(new_df)
        gz = new_df['公职'].sum()
        sy = new_df['事业2'].sum()
        tb = new_df['标备总数'].sum()
        rep = len(new_df[new_df['客户回话'].astype(str).str.contains("是", na=False)])
        return {"raw_cnt": raw_cnt, "dedup_cnt": dedup_cnt, "new_add": new_add, "gz": gz, "sy": sy, "tb": tb, "rep": rep}

    def fmt_stats(self, s, is_province_mode=False):
        def safe_div(a, b): return a / b if b != 0 else 0.0
        def to_pct(v): return f"{v:.2%}" if is_province_mode else f"{int(round(v*100))}%"
        dup_rate = safe_div((s["dedup_cnt"] - s["new_add"]), s["dedup_cnt"])
        other = s["tb"] - s["gz"] - s["sy"]
        tb_rate = safe_div(s["tb"], s["new_add"])
        rep_rate = safe_div(s["rep"], s["new_add"])
        return [int(s["raw_cnt"]), int(s["dedup_cnt"]), int(s["new_add"]), to_pct(dup_rate), int(s["gz"]), int(s["sy"]), int(other), int(s["tb"]), to_pct(tb_rate), int(s["rep"]), to_pct(rep_rate)]

    def calc_stats_retention(self, df_subset):
        retention_count = len(df_subset)
        if retention_count == 0:
            return {"gz":0, "sy":0, "tb":0, "ret":0}
        
        gz = df_subset['公职'].sum()
        sy = df_subset['事业2'].sum()
        tb = df_subset['标备总数'].sum()
        
        return {"gz": gz, "sy": sy, "tb": tb, "ret": retention_count}

    def fmt_stats_retention(self, s, is_province_mode=False):
        def safe_div(a, b): return a / b if b != 0 else 0.0
        def to_pct(v): return f"{v:.2%}" if is_province_mode else f"{int(round(v*100))}%"
        
        other = s["tb"] - s["gz"] - s["sy"]
        tb_rate = safe_div(s["tb"], s["ret"])
        
        return [
            int(s["ret"]), 
            int(s["gz"]), 
            int(s["sy"]), 
            int(other), 
            int(s["tb"]), 
            to_pct(tb_rate)
        ]

    # --- 质检报表生成 ---
    def gen_quality_inspection_report(self, df_clean, df_ret, output_file, remark_exclude, channel_exclude, qc_dates_str):
        self.log(f"生成 [质检分析综合报表]: {os.path.basename(output_file)}")
        
        # 独立过滤日期
        df_clean_qc = df_clean.copy()
        df_ret_qc = df_ret.copy()
        if qc_dates_str.strip():
            target_dates = [d.strip() for d in qc_dates_str.replace('，', ',').split(',') if d.strip()]
            if target_dates:
                df_clean_qc = df_clean_qc[df_clean_qc['日期'].isin(target_dates)]
                df_ret_qc = df_ret_qc[df_ret_qc['日期'].isin(target_dates)]
                self.log(f"🔍 质检模块已应用独立日期过滤: {target_dates}")

        if df_ret_qc.empty:
            self.log("❌ 警告：经过日期过滤后，质检数据为空，停止生成本报表！")
            return

        groups = {
            "A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"],
            "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"],
            "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]
        }

        def safe_div(a, b): return a / b if b else 0.0
        def to_pct_int(v): return f"{int(round(v * 100))}%"

        base_path = os.path.splitext(output_file)[0]
        csv_file1 = f"{base_path}_表1_详情明细底稿.csv"
        csv_file3 = f"{base_path}_表3_加几个分校明细底稿.csv"
        
        # ---------------------------------------------------------
        # 表1：详情数据表 (不以 会话存档 作为过滤条件)
        # ---------------------------------------------------------
        df1 = df_ret_qc.copy()
        df1 = df1[df1['标备总数'] == 1]
        df1 = df1[df1['客户回复消息数'] == 0]
        df1 = df1[df1['好友添加来源'].astype(str).str.contains('扫描渠道二维码', na=False)]
        
        if remark_exclude.strip():
            rem_kws = [k.strip() for k in remark_exclude.replace('，', ',').split(',') if k.strip()]
            if rem_kws: df1 = df1[~df1['备注名'].astype(str).apply(lambda x: any(k in x for k in rem_kws))]
        if channel_exclude.strip():
            ch_kws = [k.strip() for k in channel_exclude.replace('，', ',').split(',') if k.strip()]
            if ch_kws: df1 = df1[~df1['添加渠道码'].astype(str).apply(lambda x: any(k in x for k in ch_kws))]
        
        df1.to_csv(csv_file1, index=False, encoding='utf-8-sig')
        self.log(f"✅ [表1] 已提取至 CSV: {os.path.basename(csv_file1)}")
        
        # ---------------------------------------------------------
        # 表3：加几个分校
        # ---------------------------------------------------------
        df3_base = df_ret_qc.drop_duplicates(subset=['ExternalUserId', '分校'])
        df3 = df3_base.groupby('ExternalUserId').size().reset_index(name='跨分校')
        
        df3.to_csv(csv_file3, index=False, encoding='utf-8-sig')
        self.log(f"✅ [表3] 已提取至 CSV: {os.path.basename(csv_file3)}")

        # ---------------------------------------------------------
        # 汇总看板 Excel 
        # ---------------------------------------------------------
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            
            # 表2：虚假备注筛选 (★ 核心过滤：基于表1进一步筛出会话存档=是)
            df1_for_t2 = df1[df1['是否同意会话存档'].astype(str) == '是']
            rows_t2 = []
            grand_t2 = 0
            for g_name, branches in groups.items():
                g_total = 0
                for branch in branches:
                    cnt = len(df1_for_t2[df1_for_t2['分校'] == branch])
                    g_total += cnt
                    rows_t2.append([branch, cnt])
                rows_t2.append([f"{g_name}类总计", g_total])
                grand_t2 += g_total
            rows_t2.append(["全国总计", grand_t2])
            df2 = pd.DataFrame(rows_t2, columns=['分校', '虚假备注'])
            df2.to_excel(writer, sheet_name='2_虚假备注筛选', index=False)
            
            # 表4：跨10个以上分校
            over10_ids = df3[df3['跨分校'] >= 10]['ExternalUserId']
            df4_filtered = df3_base[df3_base['ExternalUserId'].isin(over10_ids)]
            rows_t4 = []
            grand_t4 = 0
            for g_name, branches in groups.items():
                g_total = 0
                for branch in branches:
                    cnt = len(df4_filtered[df4_filtered['分校'] == branch])
                    g_total += cnt
                    rows_t4.append([branch, cnt])
                rows_t4.append([f"{g_name}类总计", g_total])
                grand_t4 += g_total
            rows_t4.append(["全国总计", grand_t4])
            df4 = pd.DataFrame(rows_t4, columns=['分校', '添加客服量'])
            df4.to_excel(writer, sheet_name='4_跨10个以上分校', index=False)
            
            # 表5：会话情况
            rows_t5 = []
            grand_t5 = {'ret':0, 'has_rep':0, 'no_rep':0, 'db_zero':0}
            for g_name, branches in groups.items():
                g_t5 = {'ret':0, 'has_rep':0, 'no_rep':0, 'db_zero':0}
                for branch in branches:
                    grp = df_ret_qc[df_ret_qc['分校'] == branch]
                    ret_cnt = len(grp)
                    has_reply = (grp['客户回复消息数'] > 0).sum()
                    no_reply = (grp['客户回复消息数'] == 0).sum()
                    double_zero = ((grp['员工发送消息数'] == 0) & (grp['客户回复消息数'] == 0)).sum()
                    
                    g_t5['ret'] += ret_cnt; g_t5['has_rep'] += has_reply; g_t5['no_rep'] += no_reply; g_t5['db_zero'] += double_zero
                    grand_t5['ret'] += ret_cnt; grand_t5['has_rep'] += has_reply; grand_t5['no_rep'] += no_reply; grand_t5['db_zero'] += double_zero
                    
                    r1 = safe_div(no_reply, ret_cnt)
                    r2 = safe_div(double_zero, ret_cnt)
                    rows_t5.append([branch, ret_cnt, has_reply, no_reply, to_pct_int(r1), double_zero, to_pct_int(r2)])
                
                r1_g = safe_div(g_t5['no_rep'], g_t5['ret'])
                r2_g = safe_div(g_t5['db_zero'], g_t5['ret'])
                rows_t5.append([f"{g_name}类总计", g_t5['ret'], g_t5['has_rep'], g_t5['no_rep'], to_pct_int(r1_g), g_t5['db_zero'], to_pct_int(r2_g)])

            r1_all = safe_div(grand_t5['no_rep'], grand_t5['ret'])
            r2_all = safe_div(grand_t5['db_zero'], grand_t5['ret'])
            rows_t5.append(["全国总计", grand_t5['ret'], grand_t5['has_rep'], grand_t5['no_rep'], to_pct_int(r1_all), grand_t5['db_zero'], to_pct_int(r2_all)])

            df5 = pd.DataFrame(rows_t5, columns=['分校', '留存量', '客户有会话', '客户0会话', '客户0会话率', '双向无会话', '双向无会话率'])
            df5.to_excel(writer, sheet_name='5_会话情况', index=False)

            # 表6：好友添加来源 
            std_sources = ["获客助手", "名片分享", "其他", "群聊", "扫描名片二维码", "扫描渠道二维码", "搜索手机号", "微信联系人", "未知"]
            
            temp_clean_t6 = df_clean_qc.copy()
            temp_clean_t6['好友添加来源'] = temp_clean_t6['好友添加来源'].fillna('未知').apply(lambda x: x if x in std_sources else '其他')
            
            rows_t6 = []
            grand_t6 = {s: 0 for s in std_sources}
            
            for g_name, branches in groups.items():
                g_t6 = {s: 0 for s in std_sources}
                for branch in branches:
                    grp = temp_clean_t6[temp_clean_t6['分校'] == branch]
                    counts_dict = grp['好友添加来源'].value_counts().to_dict()
                    
                    b_counts = [counts_dict.get(s, 0) for s in std_sources]
                    b_total = sum(b_counts)
                    
                    for i, s in enumerate(std_sources):
                        g_t6[s] += b_counts[i]
                        grand_t6[s] += b_counts[i]
                        
                    b_pcts = [to_pct_int(safe_div(c, b_total)) for c in b_counts]
                    rows_t6.append([branch] + b_counts + [b_total] + [branch] + b_pcts + [b_total])
                    
                g_counts = [g_t6[s] for s in std_sources]
                g_total = sum(g_counts)
                g_pcts = [to_pct_int(safe_div(c, g_total)) for c in g_counts]
                rows_t6.append([f"{g_name}类总计"] + g_counts + [g_total] + [f"{g_name}类总计"] + g_pcts + [g_total])

            grand_counts = [grand_t6[s] for s in std_sources]
            grand_tot = sum(grand_counts)
            grand_pcts = [to_pct_int(safe_div(c, grand_tot)) for c in grand_counts]
            rows_t6.append(["全国总计"] + grand_counts + [grand_tot] + ["全国总计"] + grand_pcts + [grand_tot])

            headers = ["分校"] + std_sources + ["总计"] + ["分校"] + std_sources + ["总计"]
            df6_final = pd.DataFrame(rows_t6, columns=headers)
            df6_final.to_excel(writer, sheet_name='6_好友添加来源', index=False)

        self._style_excel(output_file)

    def gen_prov_long_merged(self, df_raw, df_dedup, output_file):
        self.log(f"生成 [省份-一维]: {os.path.basename(output_file)}")
        groups = {"A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"], "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"], "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]}
        sub_headers = ["新增好友", "本月分校内去重", "净新增", "重复率", "公职标备", "事业标备", "其他标备", "标备总量", "标备率", "回话备注", "回话备注率"]
        iter_list = [
            ("线上平台", "网站", "网站"), ("线上平台", "小红书", "小红书"), ("线上平台", "公众号", "公众号"),
            ("线上平台", "抖音", "抖音"), ("线上平台", "视频号", "视频号"), ("线上平台", "其他", "其他"),
            ("线下活动", "线下活动", "线下活动"), ("考试现场", "考试现场", "考试现场"), ("高校", "高校", "高校"), ("其他", "其他", "其他")
        ]
        final_rows = []
        grand_acc = {key: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for key in [x[2] for x in iter_list]}
        group_acc = {g: {key: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for key in [x[2] for x in iter_list]} for g in groups}

        for g_name, branches in groups.items():
            for branch in branches:
                b_raw = df_raw[df_raw['分校'] == branch]
                b_dedup = df_dedup[df_dedup['分校'] == branch]
                
                for platform_name, filter_kw, specific_name in iter_list:
                    if platform_name == "线上平台":
                        c_raw = b_raw[(b_raw['渠道'] == '线上平台') & (b_raw['线上渠道'] == filter_kw)]
                        c_dedup = b_dedup[(b_dedup['渠道'] == '线上平台') & (b_dedup['线上渠道'] == filter_kw)]
                    else:
                        c_raw = b_raw[b_raw['渠道'] == filter_kw]
                        c_dedup = b_dedup[b_dedup['渠道'] == filter_kw]

                    stats = self.calc_stats(c_raw, c_dedup)
                    for k in stats:
                        group_acc[g_name][specific_name][k] += stats[k]
                        grand_acc[specific_name][k] += stats[k]
                    final_rows.append([branch, platform_name, specific_name] + self.fmt_stats(stats, True))

            for platform_name, filter_kw, specific_name in iter_list:
                final_rows.append([f"{g_name}类总计", platform_name, specific_name] + self.fmt_stats(group_acc[g_name][specific_name], True))

        for platform_name, filter_kw, specific_name in iter_list:
            final_rows.append(["全国", platform_name, specific_name] + self.fmt_stats(grand_acc[specific_name], True))

        pd.DataFrame(final_rows, columns=["分校", "所属平台", "所属渠道"]+sub_headers).to_excel(output_file, index=False)
        self._style_excel(output_file)

    def gen_prov_wide(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成: {os.path.basename(output_file)}")
        groups = {"A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"], "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"], "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]}
        final_rows = []
        grand_acc = {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map}
        group_acc = {g: {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map} for g in groups}
        for g_name, branches in groups.items():
            for branch in branches:
                b_raw = df_raw[df_raw['分校'] == branch]
                b_dedup = df_dedup[df_dedup['分校'] == branch]
                row = [branch]
                for ch_title, ch_filter in channel_map:
                    if ch_filter:
                        if strict_online:
                            c_raw = b_raw[(b_raw['渠道'] == '线上平台') & (b_raw['线上渠道'] == ch_filter)]
                            c_dedup = b_dedup[(b_dedup['渠道'] == '线上平台') & (b_dedup['线上渠道'] == ch_filter)]
                        else:
                            c_raw = b_raw[b_raw['渠道'] == ch_filter]
                            c_dedup = b_dedup[b_dedup['渠道'] == ch_filter]
                            if ch_title == "线上平台":
                                c_raw = b_raw[(b_raw['渠道'] == '线上平台')]
                                c_dedup = b_dedup[(b_dedup['渠道'] == '线上平台')]
                    else:
                        c_raw = b_raw; c_dedup = b_dedup
                    
                    stats = self.calc_stats(c_raw, c_dedup)
                    for k in stats:
                        grand_acc[ch_title][k] += stats[k]
                        group_acc[g_name][ch_title][k] += stats[k]
                    row.extend(self.fmt_stats(stats, True))
                final_rows.append(row)
            g_row = [f"{g_name}类总计"]
            for ch_title, _ in channel_map: g_row.extend(self.fmt_stats(group_acc[g_name][ch_title], True))
            final_rows.append(g_row)
        n_row = ["全国"]
        for ch_title, _ in channel_map: n_row.extend(self.fmt_stats(grand_acc[ch_title], True))
        final_rows.append(n_row)
        self._write_wide_excel(output_file, final_rows, channel_map, 2, False)

    def gen_city_long_merged(self, df_raw, df_dedup, df_ret, output_file, do_retention):
        self.log(f"生成 [地市-一维]: {os.path.basename(output_file)}")
        base_headers = ["新增好友", "本月分校内去重", "净新增", "重复率", "净-公职标备", "净-事业标备", "净-其他标备", "净-标备总量", "净-标备率", "回话备注", "回话备注率"]
        ext_headers = []
        if do_retention: ext_headers = ["总-留存量", "总-公职标备", "总-事业标备", "总-其他标备", "总-标备总量"]

        iter_list = [
            ("线上平台", "网站", "网站"), ("线上平台", "小红书", "小红书"), ("线上平台", "公众号", "公众号"),
            ("线上平台", "抖音", "抖音"), ("线上平台", "视频号", "视频号"), ("线上平台", "其他", "其他"),
            ("线下活动", "线下活动", "线下活动"), ("考试现场", "考试现场", "考试现场"), ("高校", "高校", "高校"), ("其他", "其他", "其他")
        ]
        
        final_rows = []
        grand_acc = {key: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for key in [x[2] for x in iter_list]}
        grand_retention_acc = {key: {"gz":0,"sy":0,"tb":0,"ret":0} for key in [x[2] for x in iter_list]}

        branches = sorted(df_raw['分校'].dropna().unique())
        for branch in branches:
            b_raw = df_raw[df_raw['分校'] == branch]
            b_dedup = df_dedup[df_dedup['分校'] == branch]
            b_ret = df_ret[df_ret['分校'] == branch] if do_retention else None
            
            cities = sorted(b_raw['地市'].dropna().unique())
            for city in cities:
                c_raw_base = b_raw[b_raw['地市'] == city]
                c_dedup_base = b_dedup[b_dedup['地市'] == city]
                cat = c_raw_base['所属类别'].iloc[0] if not c_raw_base.empty else "其它"
                
                for platform_name, filter_kw, specific_name in iter_list:
                    if platform_name == "线上平台":
                        c_raw = c_raw_base[(c_raw_base['渠道'] == '线上平台') & (c_raw_base['线上渠道'] == filter_kw)]
                        c_dedup = c_dedup_base[(c_dedup_base['渠道'] == '线上平台') & (c_dedup_base['线上渠道'] == filter_kw)]
                    else:
                        c_raw = c_raw_base[c_raw_base['渠道'] == filter_kw]
                        c_dedup = c_dedup_base[c_dedup_base['渠道'] == filter_kw]
                    
                    stats = self.calc_stats(c_raw, c_dedup)
                    for k in stats: grand_acc[specific_name][k] += stats[k]
                    
                    ext_cols = []
                    if do_retention:
                        c_ret_base = b_ret[b_ret['地市'] == city]
                        if platform_name == "线上平台":
                            c_ret = c_ret_base[(c_ret_base['渠道'] == '线上平台') & (c_ret_base['线上渠道'] == filter_kw)]
                        else:
                            c_ret = c_ret_base[c_ret_base['渠道'] == filter_kw]
                        
                        ret_stats = self.calc_stats_retention(c_ret)
                        for k in ret_stats: grand_retention_acc[specific_name][k] += ret_stats[k]
                        other_ret = ret_stats["tb"] - ret_stats["gz"] - ret_stats["sy"]
                        ext_cols = [int(ret_stats["ret"]), int(ret_stats["gz"]), int(ret_stats["sy"]), int(other_ret), int(ret_stats["tb"])]

                    final_rows.append([branch, city, cat, platform_name, specific_name] + self.fmt_stats(stats, False) + ext_cols)

        for platform_name, filter_kw, specific_name in iter_list:
            row = ["全国", "总计", "-", platform_name, specific_name] + self.fmt_stats(grand_acc[specific_name], False)
            if do_retention:
                g_ret = grand_retention_acc[specific_name]
                g_other_ret = g_ret["tb"] - g_ret["gz"] - g_ret["sy"]
                row += [int(g_ret["ret"]), int(g_ret["gz"]), int(g_ret["sy"]), int(g_other_ret), int(g_ret["tb"])]
            final_rows.append(row)

        pd.DataFrame(final_rows, columns=["分校", "地市", "所属类别", "所属平台", "所属渠道"] + base_headers + ext_headers).to_excel(output_file, index=False)
        self._style_excel(output_file)

    def gen_city_wide(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成: {os.path.basename(output_file)}")
        final_rows = []
        grand_acc = {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map}
        branches = sorted(df_raw['分校'].dropna().unique())
        for branch in branches:
            b_raw = df_raw[df_raw['分校'] == branch]
            b_dedup = df_dedup[df_dedup['分校'] == branch]
            cities = sorted(b_raw['地市'].dropna().unique())
            for city in cities:
                row = [branch, city]
                c_raw_base = b_raw[b_raw['地市'] == city]
                c_dedup_base = b_dedup[b_dedup['地市'] == city]
                for ch_title, ch_filter in channel_map:
                    if ch_filter:
                        if strict_online:
                            c_raw = c_raw_base[(c_raw_base['渠道'] == '线上平台') & (c_raw_base['线上渠道'] == ch_filter)]
                            c_dedup = c_dedup_base[(c_dedup_base['渠道'] == '线上平台') & (c_dedup_base['线上渠道'] == ch_filter)]
                        else:
                            c_raw = c_raw_base[c_raw_base['渠道'] == ch_filter]
                            c_dedup = c_dedup_base[c_dedup_base['渠道'] == ch_filter]
                            if ch_title == "线上平台":
                                c_raw = c_raw_base[(c_raw_base['渠道'] == '线上平台')]
                                c_dedup = c_dedup_base[(c_dedup_base['渠道'] == '线上平台')]
                    else:
                        c_raw = c_raw_base; c_dedup = c_dedup_base
                    stats = self.calc_stats(c_raw, c_dedup)
                    for k in stats: grand_acc[ch_title][k] += stats[k]
                    row.extend(self.fmt_stats(stats, False))
                final_rows.append(row)
        n_row = ["全国", "总计"]
        for ch_title, _ in channel_map: n_row.extend(self.fmt_stats(grand_acc[ch_title], False))
        final_rows.append(n_row)
        self._write_wide_excel(output_file, final_rows, channel_map, 3, True)

    def gen_special_city_report(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成 [单独地市]: {os.path.basename(output_file)}")
        final_rows = []
        grand_acc = {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map}
        for city in FIXED_ORDER_CITIES:
            row_raw = df_raw[df_raw['地市'] == city]
            row_dedup = df_dedup[df_dedup['地市'] == city]
            branch_name = row_raw['分校'].iloc[0] if not row_raw.empty else DEFAULT_BRANCH_MAP.get(city, "未知分校")
            row_data = [branch_name, city]
            for ch_title, ch_filter in channel_map:
                if ch_filter:
                    if strict_online:
                        c_raw = row_raw[(row_raw['渠道'] == '线上平台') & (row_raw['线上渠道'] == ch_filter)]
                        c_dedup = row_dedup[(row_dedup['渠道'] == '线上平台') & (row_dedup['线上渠道'] == ch_filter)]
                    else:
                        c_raw = row_raw[row_raw['渠道'] == ch_filter]
                        c_dedup = row_dedup[row_dedup['渠道'] == ch_filter]
                        if ch_title == "线上平台":
                            c_raw = row_raw[(row_raw['渠道'] == '线上平台')]
                            c_dedup = row_dedup[(row_dedup['渠道'] == '线上平台')]
                else:
                    c_raw = row_raw; c_dedup = row_dedup
                stats = self.calc_stats(c_raw, c_dedup)
                for k in stats: grand_acc[ch_title][k] += stats[k]
                row_data.extend(self.fmt_stats(stats, False))
            final_rows.append(row_data)
        n_row = ["全国", "总计"]
        for ch_title, _ in channel_map: n_row.extend(self.fmt_stats(grand_acc[ch_title], False))
        final_rows.append(n_row)
        self._write_wide_excel(output_file, final_rows, channel_map, 3, True)

    def gen_wide_retention(self, df_ret, output_file, channel_map, strict_online, is_province_mode, is_special_mode):
        self.log(f"生成 [留存版-宽表]: {os.path.basename(output_file)}")
        groups = {"A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"], "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"], "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]}
        final_rows = []
        grand_acc = {ch[0]: {"gz":0,"sy":0,"tb":0,"ret":0} for ch in channel_map}
        group_acc = {g: {ch[0]: {"gz":0,"sy":0,"tb":0,"ret":0} for ch in channel_map} for g in groups}

        if is_special_mode:
            for city in FIXED_ORDER_CITIES:
                row_ret = df_ret[df_ret['地市'] == city]
                branch_name = row_ret['分校'].iloc[0] if not row_ret.empty else DEFAULT_BRANCH_MAP.get(city, "未知分校")
                row_data = [branch_name, city]
                for ch_title, ch_filter in channel_map:
                    if ch_filter:
                        if strict_online:
                            c_ret = row_ret[(row_ret['渠道'] == '线上平台') & (row_ret['线上渠道'] == ch_filter)]
                        else:
                            c_ret = row_ret[row_ret['渠道'] == ch_filter]
                            if ch_title == "线上平台": c_ret = row_ret[row_ret['渠道'] == '线上平台']
                    else:
                        c_ret = row_ret
                    stats = self.calc_stats_retention(c_ret)
                    for k in stats: grand_acc[ch_title][k] += stats[k]
                    row_data.extend(self.fmt_stats_retention(stats, False))
                final_rows.append(row_data)
        
        elif is_province_mode:
            for g_name, branches in groups.items():
                for branch in branches:
                    b_ret = df_ret[df_ret['分校'] == branch]
                    row_data = [branch]
                    for ch_title, ch_filter in channel_map:
                        if ch_filter:
                            if strict_online:
                                c_ret = b_ret[(b_ret['渠道'] == '线上平台') & (b_ret['线上渠道'] == ch_filter)]
                            else:
                                c_ret = b_ret[b_ret['渠道'] == ch_filter]
                                if ch_title == "线上平台": c_ret = b_ret[b_ret['渠道'] == '线上平台']
                        else:
                            c_ret = b_ret
                        stats = self.calc_stats_retention(c_ret)
                        for k in stats:
                            grand_acc[ch_title][k] += stats[k]
                            group_acc[g_name][ch_title][k] += stats[k]
                        row_data.extend(self.fmt_stats_retention(stats, True))
                    final_rows.append(row_data)
                g_row = [f"{g_name}类总计"]
                for ch_title, _ in channel_map: g_row.extend(self.fmt_stats_retention(group_acc[g_name][ch_title], True))
                final_rows.append(g_row)

        else:
            branches = sorted(df_ret['分校'].dropna().unique())
            for branch in branches:
                b_ret = df_ret[df_ret['分校'] == branch]
                cities = sorted(b_ret['地市'].dropna().unique())
                for city in cities:
                    c_ret_base = b_ret[b_ret['地市'] == city]
                    row_data = [branch, city]
                    for ch_title, ch_filter in channel_map:
                        if ch_filter:
                            if strict_online:
                                c_ret = c_ret_base[(c_ret_base['渠道'] == '线上平台') & (c_ret_base['线上渠道'] == ch_filter)]
                            else:
                                c_ret = c_ret_base[c_ret_base['渠道'] == ch_filter]
                                if ch_title == "线上平台": c_ret = c_ret_base[c_ret_base['渠道'] == '线上平台']
                        else:
                            c_ret = c_ret_base
                        stats = self.calc_stats_retention(c_ret)
                        for k in stats: grand_acc[ch_title][k] += stats[k]
                        row_data.extend(self.fmt_stats_retention(stats, False))
                    final_rows.append(row_data)

        n_row = ["全国", "总计"] if not is_province_mode else ["全国"]
        for ch_title, _ in channel_map: n_row.extend(self.fmt_stats_retention(grand_acc[ch_title], is_province_mode))
        final_rows.append(n_row)

        self._write_wide_excel_retention(output_file, final_rows, channel_map, 2 if is_province_mode else 3, not is_province_mode)

    def gen_date_summary(self, df_dedup, target_dates_str, output_file, is_city_level=True):
        if not target_dates_str.strip(): return
        self.log(f"生成日期汇总: {os.path.basename(output_file)}")
        dates = [d.strip() for d in target_dates_str.replace('，', ',').split(',') if d.strip()]
        mask = df_dedup['日期'].isin(dates)
        filtered = df_dedup[mask]
        group_cols = ['分校', '地市'] if is_city_level else ['分校']
        cnt_all = filtered.groupby(group_cols).size().reset_index(name='total')
        cnt_web = filtered[filtered['线上渠道'] == '网站'].groupby(group_cols).size().reset_index(name='web')
        res = pd.merge(cnt_all, cnt_web, on=group_cols, how='left').fillna(0)
        res['web'] = res['web'].astype(int)
        
        if hasattr(self, 'use_special_city') and self.use_special_city:
             res = res[res['地市'].isin(FIXED_ORDER_CITIES)]
             res['地市'] = pd.Categorical(res['地市'], categories=FIXED_ORDER_CITIES, ordered=True)
             res = res.sort_values('地市')
             res['地市'] = res['地市'].astype(str)

        res.loc['Sum'] = pd.Series(res[['total', 'web']].sum(), index=['total', 'web'])
        res.at['Sum', '分校'] = '全国'
        if is_city_level: res.at['Sum', '地市'] = '总计'
        res.rename(columns={'total': f"汇总({','.join(dates)})", 'web': "其中:网站"}, inplace=True)
        res.to_excel(output_file, index=False)
        self._style_excel(output_file)

    def _write_wide_excel(self, output_file, rows, channel_map, start_col_idx, has_city_col):
        sub_headers = ["新增好友", "本月分校内去重", "净新增", "重复率", "公职标备", "事业标备", "其他标备", "标备总量", "标备率", "回话备注", "回话备注率"]
        df_out = pd.DataFrame(rows)
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_out.to_excel(writer, index=False, header=False, startrow=2)
            ws = writer.sheets['Sheet1']
            ws['A1'] = "分校"; ws.merge_cells('A1:A2')
            if has_city_col: ws['B1'] = "地市"; ws.merge_cells('B1:B2')
            curr_col = start_col_idx
            for ch_title, _ in channel_map:
                ws.cell(row=1, column=curr_col).value = ch_title
                for sub in sub_headers:
                    ws.cell(row=2, column=curr_col).value = sub
                    curr_col += 1
                ws.merge_cells(start_row=1, start_column=curr_col-len(sub_headers), end_row=1, end_column=curr_col-1)
            if has_city_col:
                max_r = ws.max_row
                curr_b = ws['A3'].value
                start_r = 3
                for r in range(4, max_r + 2):
                    val = ws[f'A{r}'].value if r <= max_r else None
                    if val != curr_b:
                        if start_r < r - 1: ws.merge_cells(f'A{start_r}:A{r-1}')
                        curr_b = val
                        start_r = r
            self._style_excel_ws(ws)

    def _write_wide_excel_retention(self, output_file, rows, channel_map, start_col_idx, has_city_col):
        sub_headers = ["留存量", "公职标备", "事业标备", "其他标备", "标备总量", "标备率"]
        df_out = pd.DataFrame(rows)
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_out.to_excel(writer, index=False, header=False, startrow=2)
            ws = writer.sheets['Sheet1']
            ws['A1'] = "分校"; ws.merge_cells('A1:A2')
            if has_city_col: ws['B1'] = "地市"; ws.merge_cells('B1:B2')
            curr_col = start_col_idx
            for ch_title, _ in channel_map:
                ws.cell(row=1, column=curr_col).value = ch_title
                for sub in sub_headers:
                    ws.cell(row=2, column=curr_col).value = sub
                    curr_col += 1
                ws.merge_cells(start_row=1, start_column=curr_col-len(sub_headers), end_row=1, end_column=curr_col-1)
            
            if has_city_col:
                max_r = ws.max_row
                curr_b = ws['A3'].value
                start_r = 3
                for r in range(4, max_r + 2):
                    val = ws[f'A{r}'].value if r <= max_r else None
                    if val != curr_b:
                        if start_r < r - 1: ws.merge_cells(f'A{start_r}:A{r-1}')
                        curr_b = val
                        start_r = r
            self._style_excel_ws(ws)

    def _style_excel(self, file_path):
        try:
            wb = load_workbook(file_path)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                self._style_excel_ws(ws)
            wb.save(file_path)
        except Exception as e:
            self.log(f"样式渲染异常(跳过): {e}")

    def _style_excel_ws(self, ws):
        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        align = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = align
        for row in ws.iter_rows(min_row=1, max_row=2):
            for cell in row:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
        ws.column_dimensions['A'].width = 18
        if ws['B1'].value == '地市': ws.column_dimensions['B'].width = 15

# ==============================================================================
# GUI 界面 (4.4 业务解耦版)
# ==============================================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("数据自动化统计工具 4.4 (宽屏解耦版)")
        self.root.geometry("980x520")
        self.root.configure(bg="#f8f9fa")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background="#f8f9fa")
        style.configure('TLabelframe', background="#f8f9fa")
        style.configure('TLabelframe.Label', font=("Microsoft YaHei", 10, "bold"), background="#f8f9fa", foreground="#333")
        style.configure('TLabel', background="#f8f9fa", font=("Microsoft YaHei", 9))
        style.configure('TCheckbutton', background="#f8f9fa", font=("Microsoft YaHei", 9))
        style.configure('TButton', font=("Microsoft YaHei", 10), background="#007bff", foreground="white", borderwidth=0)
        style.map('TButton', background=[('active', '#0056b3')], foreground=[('active', 'white')])
        style.configure('TNotebook', background="#f8f9fa")
        style.configure('TNotebook.Tab', font=("Microsoft YaHei", 10, "bold"), padding=[15, 5])

        title_frame = tk.Frame(root, bg="#007bff", height=50)
        title_frame.pack(fill="x")
        tk.Label(title_frame, text="🚀 数据自动化统计工具 4.4", font=("Microsoft YaHei", 15, "bold"), bg="#007bff", fg="white").pack(pady=10)

        main_frame = tk.Frame(root, bg="#f8f9fa")
        main_frame.pack(fill="both", expand=True, padx=15, pady=10)

        left_frame = tk.Frame(main_frame, bg="#f8f9fa")
        left_frame.pack(side="left", fill="both", expand=True)

        right_frame = tk.Frame(main_frame, bg="#f8f9fa", width=350)
        right_frame.pack(side="right", fill="y", padx=(15, 0))
        right_frame.pack_propagate(False)

        self.notebook = ttk.Notebook(left_frame)
        self.notebook.pack(fill="both", expand=True)

        tab_import = ttk.Frame(self.notebook)
        tab_regular = ttk.Frame(self.notebook)
        tab_qc = ttk.Frame(self.notebook)
        
        self.notebook.add(tab_import, text=" 📂 数据源导入 ")
        self.notebook.add(tab_regular, text=" 📈 常规报表 ")
        self.notebook.add(tab_qc, text=" 📊 质检数据分析 ")

        # ==================== 标签 1：导入 ====================
        step1_frame = ttk.LabelFrame(tab_import, text=" 步骤 1：导入数据 ")
        step1_frame.pack(fill="x", pady=20, padx=15, ipady=15)
        
        tk.Label(step1_frame, text="请选择包含数据的 Excel 或 CSV 文件（支持多选）", bg="#f8f9fa", fg="#666").pack(pady=(5, 10))
        f_file = ttk.Frame(step1_frame)
        f_file.pack(fill="x", padx=15)
        self.ent_f = ttk.Entry(f_file)
        self.ent_f.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(f_file, text="📂 浏览文件", command=self.browse, width=12).pack(side="right")

        # ==================== 标签 2：常规报表 ====================
        step2_frame = ttk.LabelFrame(tab_regular, text=" 步骤 2：参数设置 ")
        step2_frame.pack(fill="x", pady=(10, 5), padx=10, ipady=3)
        
        f_p1 = ttk.Frame(step2_frame)
        f_p1.pack(fill="x", padx=10, pady=5)
        ttk.Label(f_p1, text="📅 判断月份 (如 2026-01):").pack(side="left")
        self.ent_m = ttk.Entry(f_p1, width=20)
        self.ent_m.insert(0, "2026-01")
        self.ent_m.pack(side="left", padx=10)

        f_p2 = ttk.Frame(step2_frame)
        f_p2.pack(fill="x", padx=10, pady=5)
        ttk.Label(f_p2, text="📆 特定日期 (如 1月1日):").pack(side="left")
        self.ent_d = ttk.Entry(f_p2)
        self.ent_d.pack(side="left", fill="x", expand=True, padx=10)

        f_p3 = ttk.Frame(step2_frame)
        f_p3.pack(fill="x", padx=10, pady=5)
        ttk.Label(f_p3, text="🔍 自定义标备词:").pack(side="left")
        self.ent_ckw = ttk.Entry(f_p3)
        self.ent_ckw.pack(side="left", fill="x", expand=True, padx=10)

        step3_frame = ttk.LabelFrame(tab_regular, text=" 步骤 3：常规报表任务 ")
        step3_frame.pack(fill="x", pady=5, padx=10, ipady=2)
        
        self.v_retention = tk.BooleanVar(value=False)
        self.v_pl = tk.BooleanVar(value=False)
        self.v_pw = tk.BooleanVar(value=False)
        self.v_cl = tk.BooleanVar(value=False)
        self.v_cw = tk.BooleanVar(value=False)
        self.v_special = tk.BooleanVar(value=False)

        ttk.Checkbutton(step3_frame, text="🔥 同时生成“留存版”报表 (部分表依赖)", variable=self.v_retention, style='TCheckbutton').pack(anchor="w", padx=15, pady=5)
        grid_frame = ttk.Frame(step3_frame)
        grid_frame.pack(fill="x", padx=15, pady=2)
        ttk.Checkbutton(grid_frame, text="省份 - 一维数据", variable=self.v_pl).grid(row=0, column=0, sticky="w", pady=5, padx=10)
        ttk.Checkbutton(grid_frame, text="省份 - 宽表数据", variable=self.v_pw).grid(row=0, column=1, sticky="w", pady=5, padx=10)
        ttk.Checkbutton(grid_frame, text="地市 - 一维数据", variable=self.v_cl).grid(row=1, column=0, sticky="w", pady=5, padx=10)
        ttk.Checkbutton(grid_frame, text="地市 - 宽表数据", variable=self.v_cw).grid(row=1, column=1, sticky="w", pady=5, padx=10)
        ttk.Checkbutton(grid_frame, text="★ 单独地市报表", variable=self.v_special).grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=5)

        # ★ 常规页签专属按钮
        self.btn_regular = ttk.Button(tab_regular, text="📈 仅生成常规报表", command=lambda: self.run('regular'))
        self.btn_regular.pack(side="bottom", fill="x", pady=10, padx=15, ipady=8)

        # ==================== 标签 3：质检分析 ====================
        step4_frame = ttk.LabelFrame(tab_qc, text=" 质检专属配置 (不干扰常规报表) ")
        step4_frame.pack(fill="both", expand=True, pady=10, padx=10, ipady=5)
        
        info_lbl = tk.Label(step4_frame, text="💡 提示：点击下方按钮将只执行质检逻辑。表1(明细)不会受到会话存档限制，表2(汇总)会自动过滤出会话存档=是的数据。", 
                            bg="#e9ecef", fg="#495057", font=("Microsoft YaHei", 9), justify="left", wraplength=500)
        info_lbl.pack(fill="x", padx=15, pady=10)

        f_q0 = ttk.Frame(step4_frame)
        f_q0.pack(fill="x", padx=15, pady=6)
        ttk.Label(f_q0, text="📆 独立过滤日期:").pack(side="left")
        self.ent_q_date = ttk.Entry(f_q0)
        self.ent_q_date.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Label(f_q0, text="(如 3月3日,3月7日)", foreground="gray").pack(side="right")

        f_q1 = ttk.Frame(step4_frame)
        f_q1.pack(fill="x", padx=15, pady=6)
        ttk.Label(f_q1, text="🚫 备注名排除包含:").pack(side="left")
        self.ent_q_rem = ttk.Entry(f_q1)
        self.ent_q_rem.insert(0, "学26,学27,报26,课26,前台")
        self.ent_q_rem.pack(side="left", fill="x", expand=True, padx=5)

        f_q2 = ttk.Frame(step4_frame)
        f_q2.pack(fill="x", padx=15, pady=6)
        ttk.Label(f_q2, text="🚫 渠道码排除包含:").pack(side="left")
        self.ent_q_ch = ttk.Entry(f_q2)
        self.ent_q_ch.insert(0, "前台")
        self.ent_q_ch.pack(side="left", fill="x", expand=True, padx=5)

        # ★ 质检页签专属按钮
        self.btn_qc = ttk.Button(tab_qc, text="🎯 仅生成质检报表", command=lambda: self.run('qc'))
        self.btn_qc.pack(side="bottom", fill="x", pady=10, padx=15, ipady=8)

        # ==================== 右侧：常驻日志 ====================
        log_label_frame = ttk.LabelFrame(right_frame, text=" 实时运行日志 ")
        log_label_frame.pack(fill="both", expand=True)
        self.log_txt = scrolledtext.ScrolledText(log_label_frame, font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4", relief="flat")
        self.log_txt.pack(fill="both", expand=True, padx=2, pady=2)


    def log(self, msg):
        self.log_txt.config(state='normal')
        self.log_txt.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_txt.see(tk.END)
        self.log_txt.config(state='disabled')

    def browse(self):
        fs = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx *.xls *.csv")])
        if fs: self.ent_f.delete(0, tk.END); self.ent_f.insert(0, ";".join(fs))

    # ★ 运行方法改造：接收 task_type 区分任务
    def run(self, task_type):
        threading.Thread(target=self.process, args=(task_type,)).start()

    def process(self, task_type):
        files = self.ent_f.get().split(";")
        if not files or files[0] == "": return messagebox.showerror("错误", "请先选择数据文件！")
        
        start_time = time.time()
        # 锁定两个按钮
        self.btn_regular.config(state="disabled", text="⏳ 处理中...")
        self.btn_qc.config(state="disabled", text="⏳ 处理中...")
        self.log_txt.config(state='normal')
        self.log_txt.delete(1.0, tk.END)
        self.log_txt.config(state='disabled')
        
        try:
            task_name = "【质检业务流】" if task_type == 'qc' else "【常规业务流】"
            self.log(f"🚀 开始执行 {task_name}")
            
            proc = UniversalProcessor(self.log)
            dfs = []
            for f in files:
                if f.strip(): dfs.append(pd.read_excel(f))
            if not dfs: raise ValueError("未读取到有效数据")
            full_df = pd.concat(dfs, ignore_index=True)
            
            # 第一步：通用的全量提取与清洗（底层复用）
            df_clean, df_dp, df_dc, df_ret = proc.process_step1(full_df, self.ent_m.get(), self.ent_ckw.get())
            
            base_dir = os.path.dirname(files[0])
            base_name = os.path.splitext(os.path.basename(files[0]))[0]

            # ==============================
            # 分支 A：常规报表逻辑
            # ==============================
            if task_type == 'regular':
                # 常规底稿保存
                df_dc.to_csv(os.path.join(base_dir, f"{base_name}_00_清洗后源数据.csv"), index=False, encoding='utf-8-sig')
                if self.v_retention.get():
                    df_ret.to_csv(os.path.join(base_dir, f"{base_name}_00_留存版源数据(未去重).csv"), index=False, encoding='utf-8-sig')
                self.log("✅ 底稿保存完毕")

                m_ch = [("微信好友总量", None), ("线上平台", "线上平台"), ("线下平台", "线下活动"), ("考试现场", "考试现场"), ("高校", "高校"), ("其他", "其他")]
                m_on = [("微信好友总量", None), ("网站", "网站"), ("小红书", "小红书"), ("公众号", "公众号"), ("抖音", "抖音"), ("视频号", "视频号"), ("其他", "其他")]
                m_ch_ret = m_ch.copy(); m_on_ret = m_on.copy()

                do_retention = self.v_retention.get() and not df_ret.empty

                if self.v_pl.get(): proc.gen_prov_long_merged(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_01_省份一维_合并报表.xlsx"))
                if self.v_pw.get():
                    proc.gen_prov_wide(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_02_省份合并_渠道.xlsx"), m_ch, False)
                    proc.gen_prov_wide(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_03_省份合并_线上.xlsx"), m_on, True)
                    if do_retention:
                        proc.gen_wide_retention(df_ret, os.path.join(base_dir, f"{base_name}_02_省份合并_渠道_留存版.xlsx"), m_ch_ret, False, True, False)
                        proc.gen_wide_retention(df_ret, os.path.join(base_dir, f"{base_name}_03_省份合并_线上_留存版.xlsx"), m_on_ret, True, True, False)
                if self.v_cl.get(): proc.gen_city_long_merged(df_clean, df_dc, df_ret, os.path.join(base_dir, f"{base_name}_04_地市一维_合并报表.xlsx"), do_retention)
                if self.v_cw.get():
                    proc.gen_city_wide(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_05_地市合并_渠道.xlsx"), m_ch, False)
                    proc.gen_city_wide(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_06_地市合并_线上.xlsx"), m_on, True)
                    if do_retention:
                        proc.gen_wide_retention(df_ret, os.path.join(base_dir, f"{base_name}_05_地市合并_渠道_留存版.xlsx"), m_ch_ret, False, False, False)
                        proc.gen_wide_retention(df_ret, os.path.join(base_dir, f"{base_name}_06_地市合并_线上_留存版.xlsx"), m_on_ret, True, False, False)
                if self.v_special.get():
                    proc.gen_special_city_report(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_07_单独地市_渠道.xlsx"), m_ch, False)
                    proc.gen_special_city_report(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_08_单独地市_线上.xlsx"), m_on, True)
                    if do_retention:
                        proc.gen_wide_retention(df_ret, os.path.join(base_dir, f"{base_name}_07_单独地市_渠道_留存版.xlsx"), m_ch_ret, False, False, True)
                        proc.gen_wide_retention(df_ret, os.path.join(base_dir, f"{base_name}_08_单独地市_线上_留存版.xlsx"), m_on_ret, True, False, True)

                if self.ent_d.get().strip():
                    proc.use_special_city = self.v_special.get()
                    use_city = self.v_cl.get() or self.v_cw.get() or self.v_special.get()
                    target_df = df_dc if use_city else df_dp
                    proc.gen_date_summary(target_df, self.ent_d.get(), os.path.join(base_dir, f"{base_name}_09_日期汇总.xlsx"), use_city)

            # ==============================
            # 分支 B：质检报表逻辑
            # ==============================
            elif task_type == 'qc':
                if df_ret.empty:
                    self.log("❌ 无法生成质检报表：未获取到'已添加'的留存数据，请检查源表。")
                else:
                    qc_file = os.path.join(base_dir, f"{base_name}_10_质检分析汇总看板.xlsx")
                    proc.gen_quality_inspection_report(df_clean, df_ret, qc_file, self.ent_q_rem.get(), self.ent_q_ch.get(), self.ent_q_date.get())

            duration = time.time() - start_time
            self.log(f"🎉 任务完成！耗时: {duration:.2f}秒")
            messagebox.showinfo("成功", f"执行完毕\n耗时: {duration:.2f} 秒")

        except Exception as e:
            self.log(f"❌ 错误: {e}")
            messagebox.showerror("错误", str(e))
        finally:
            self.btn_regular.config(state="normal", text="📈 仅生成常规报表")
            self.btn_qc.config(state="normal", text="🎯 仅生成质检报表")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
