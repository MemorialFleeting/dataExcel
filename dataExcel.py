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
        self.log("【步骤1】清洗数据、提取地市、计算标签...")
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

        # 标签
        df['公职'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["26公职","27公职","28公职","29公职"]))
        kw_shiye = ["26事业","26三支","26社区","26辅警","26书记员","26国企","25事业","25三支","25社区","25辅警","25书记员","25国企", "27事业","27三支","27社区","27辅警","27书记员","27国企","28事业","28三支","28社区","28辅警","28书记员","28国企", "29事业","29三支","29社区","29辅警","29书记员","29国企"]
        df['事业辅助列'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, kw_shiye))
        kw_jiaoshi = ["26教师","26特岗","26教资","25教师","25特岗","25教资","27教师","27特岗","27教资","28教师","28特岗","28教资","29教师","29特岗","29教资"]
        df['教师'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, kw_jiaoshi))
        df['文职'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["25文职","26文职","27文职","28文职","29文职"]))
        df['医疗'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["25医疗","26医疗","27医疗","28医疗","29医疗"]))
        df['银行'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["25银行","26银行","27银行","28银行","29银行"]))
        df['考研'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["25考研","26考研","27考研","28考研","29考研"]))
        df['学历'] = df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["25学历","26学历","27学历","28学历","29学历"]))
        df['其他'] = ""
        
        # 自定义标备
        custom_kws = [k.strip() for k in custom_total_kws_str.replace('，', ',').split(',') if k.strip()]
        if custom_kws:
            self.log(f"⚡ 启用自定义标备逻辑，关键词: {custom_kws}")
            def check_custom_total(remark):
                if pd.isna(remark): return 0
                remark_str = str(remark)
                return 1 if any(kw in remark_str for kw in custom_kws) else 0
            df['标备总数'] = df['备注名'].apply(check_custom_total)
        else:
            all_cols_check = ['公职', '事业辅助列', '教师', '文职', '医疗', '银行', '考研', '学历', '其他']
            df['标备总数'] = df.apply(lambda row: 1 if any(row[c] in [1, "1"] for c in all_cols_check) else 0, axis=1)

        # 多月份
        date_kws = [k.strip() for k in date_keyword.replace('，', ',').split(',') if k.strip()]
        if not date_kws: date_kws = ["FAIL_SAFE"]
        self.log(f"📅 判断新增的日期关键词: {date_kws}")
        
        def check_is_new(create_time):
            s_time = str(create_time)
            if any(k in s_time for k in date_kws): return "是"
            return "否"
        df['！！！是否新增'] = df['创建时间'].apply(check_is_new)

        kw_ch = ["网络","社会","高校","线上","线上平台","线上活动","现场","线下活动"]
        res_ch = ["线上平台","线下活动","高校","线上平台","线上平台","线上平台","考试现场","线下活动"]
        df['渠道'] = df['渠道活码分组'].apply(lambda x: self.excel_lookup_find(x, kw_ch, res_ch, "其他"))
        
        kw_on = ["网站","公众号","小红书","视频号","抖音","文章页","网页","专题","附件","小程序"]
        res_on = ["网站","公众号","小红书","视频号","抖音","网站","网站","网站","网站","其他"]
        df['线上渠道'] = df['渠道活码分组'].apply(lambda x: self.excel_lookup_find(x, kw_on, res_on, "其他"))

        def check_shiye2(row):
            if row['公职'] in [1, "1"]: return 0
            if row['事业辅助列'] in [1, "1"]: return 1
            return 0
        df['事业2'] = df.apply(check_shiye2, axis=1)

        def check_reply(row):
            has_reply = pd.notna(row['客户上次回复时间']) and str(row['客户上次回复时间']).strip() != ""
            return "是" if has_reply and str(row['标备总数']) == "1" else "否"
        df['客户回话'] = df.apply(check_reply, axis=1)

        def fmt_date(val):
            try: return f"{pd.to_datetime(val).month}月{pd.to_datetime(val).day}日"
            except: return str(val)
        df['日期'] = df['好友添加时间'].apply(fmt_date)

        col_z_name = df.columns[25] if len(df.columns) > 25 else None
        final_cols = []
        if '备注名' in df.columns: final_cols.append('备注名')
        if col_z_name: final_cols.append(col_z_name)
        final_cols.extend(['分校', '地市', '所属类别', '公职', '事业2', '教师', '文职', '医疗', '银行', '考研', '学历', '其他', '标备总数', '！！！是否新增', '渠道', '线上渠道', '客户回话', '日期'])
        df_clean = df[[c for c in final_cols if c in df.columns]].copy()
        
        for col in ['公职', '事业2', '标备总数']:
            if col in df_clean.columns:
                df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

        df_dedup_prov = df_clean.drop_duplicates(subset=[col_z_name, '分校'], keep='first') if col_z_name else df_clean.copy()
        df_dedup_city = df_clean.drop_duplicates(subset=[col_z_name, '分校', '地市'], keep='first') if col_z_name else df_clean.copy()
        
        return df_clean, df_dedup_prov, df_dedup_city

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

    # --- 生成报表 ---
    def gen_prov_long(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成: {os.path.basename(output_file)}")
        groups = {"A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"], "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"], "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]}
        sub_headers = ["新增好友", "本月分校内去重", "净新增", "重复率", "公职标备", "事业标备", "其他标备", "标备总量", "标备率", "回话备注", "回话备注率"]
        final_rows = []
        grand_acc = {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map}
        group_acc = {g: {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map} for g in groups}
        for g_name, branches in groups.items():
            for branch in branches:
                b_raw = df_raw[df_raw['分校'] == branch]
                b_dedup = df_dedup[df_dedup['分校'] == branch]
                for ch_title, ch_filter in channel_map:
                    c_raw = b_raw[(b_raw['渠道'] == '线上平台') & (b_raw['线上渠道'] == ch_filter)] if (ch_filter and strict_online) else (b_raw[b_raw['渠道'] == ch_filter] if ch_filter else b_raw)
                    c_dedup = b_dedup[(b_dedup['渠道'] == '线上平台') & (b_dedup['线上渠道'] == ch_filter)] if (ch_filter and strict_online) else (b_dedup[b_dedup['渠道'] == ch_filter] if ch_filter else b_dedup)
                    stats = self.calc_stats(c_raw, c_dedup)
                    for k in stats:
                        group_acc[g_name][ch_title][k] += stats[k]
                        grand_acc[ch_title][k] += stats[k]
                    final_rows.append([branch, ch_title] + self.fmt_stats(stats, True))
            for ch_title, _ in channel_map:
                final_rows.append([f"{g_name}类总计", ch_title] + self.fmt_stats(group_acc[g_name][ch_title], True))
        for ch_title, _ in channel_map:
            final_rows.append(["全国", ch_title] + self.fmt_stats(grand_acc[ch_title], True))
        pd.DataFrame(final_rows, columns=["分校", "所属平台"]+sub_headers).to_excel(output_file, index=False)
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
                        # ★★★ 修复点：移除了 strict_online 为 False 时的子渠道判断
                        if strict_online: # 线上细分
                            c_raw = b_raw[(b_raw['渠道'] == '线上平台') & (b_raw['线上渠道'] == ch_filter)]
                            c_dedup = b_dedup[(b_dedup['渠道'] == '线上平台') & (b_dedup['线上渠道'] == ch_filter)]
                        else: # 渠道维度
                            c_raw = b_raw[b_raw['渠道'] == ch_filter]
                            c_dedup = b_dedup[b_dedup['渠道'] == ch_filter]
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

    def gen_city_long(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成: {os.path.basename(output_file)}")
        sub_headers = ["新增好友", "本月分校内去重", "净新增", "重复率", "公职标备", "事业标备", "其他标备", "标备总量", "标备率", "回话备注", "回话备注率"]
        final_rows = []
        grand_acc = {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map}
        branches = sorted(df_raw['分校'].dropna().unique())
        for branch in branches:
            b_raw = df_raw[df_raw['分校'] == branch]
            b_dedup = df_dedup[df_dedup['分校'] == branch]
            cities = sorted(b_raw['地市'].dropna().unique())
            for city in cities:
                c_raw_base = b_raw[b_raw['地市'] == city]
                c_dedup_base = b_dedup[b_dedup['地市'] == city]
                cat = c_raw_base['所属类别'].iloc[0] if not c_raw_base.empty else "其它"
                for ch_title, ch_filter in channel_map:
                    if ch_filter:
                        c_raw = c_raw_base[(c_raw_base['渠道'] == '线上平台') & (c_raw_base['线上渠道'] == ch_filter)] if strict_online else c_raw_base[c_raw_base['渠道'] == ch_filter]
                        c_dedup = c_dedup_base[(c_dedup_base['渠道'] == '线上平台') & (c_dedup_base['线上渠道'] == ch_filter)] if strict_online else c_dedup_base[c_dedup_base['渠道'] == ch_filter]
                    else:
                        c_raw = c_raw_base; c_dedup = c_dedup_base
                    stats = self.calc_stats(c_raw, c_dedup)
                    for k in stats: grand_acc[ch_title][k] += stats[k]
                    final_rows.append([branch, city, cat, ch_title] + self.fmt_stats(stats, False))
        for ch_title, _ in channel_map:
            final_rows.append(["全国", "总计", "-", ch_title] + self.fmt_stats(grand_acc[ch_title], False))
        pd.DataFrame(final_rows, columns=["分校", "地市", "所属类别", "所属平台"]+sub_headers).to_excel(output_file, index=False)
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
                        # ★★★ 修复点：移除了 strict_online 为 False 时的子渠道判断
                        if strict_online:
                            c_raw = c_raw_base[(c_raw_base['渠道'] == '线上平台') & (c_raw_base['线上渠道'] == ch_filter)]
                            c_dedup = c_dedup_base[(c_dedup_base['渠道'] == '线上平台') & (c_dedup_base['线上渠道'] == ch_filter)]
                        else:
                            c_raw = c_raw_base[c_raw_base['渠道'] == ch_filter]
                            c_dedup = c_dedup_base[c_dedup_base['渠道'] == ch_filter]
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
                    # ★★★ 修复点：移除了 strict_online 为 False 时的子渠道判断
                    if strict_online:
                        c_raw = row_raw[(row_raw['渠道'] == '线上平台') & (row_raw['线上渠道'] == ch_filter)]
                        c_dedup = row_dedup[(row_dedup['渠道'] == '线上平台') & (row_dedup['线上渠道'] == ch_filter)]
                    else:
                        c_raw = row_raw[row_raw['渠道'] == ch_filter]
                        c_dedup = row_dedup[row_dedup['渠道'] == ch_filter]
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

    def _style_excel(self, file_path):
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            self._style_excel_ws(ws)
            wb.save(file_path)
        except: pass

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
        ws.column_dimensions['A'].width = 15
        if ws['B1'].value == '地市': ws.column_dimensions['B'].width = 15

# ==============================================================================
# GUI 界面
# ==============================================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("数据统计工具2.0")
        self.root.geometry("680x750")
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

        title_frame = tk.Frame(root, bg="#007bff", height=50)
        title_frame.pack(fill="x")
        tk.Label(title_frame, text="🚀 数据自动化统计工具", font=("Microsoft YaHei", 14, "bold"), bg="#007bff", fg="white").pack(pady=10)

        main_frame = tk.Frame(root, bg="#f8f9fa")
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        step1_frame = ttk.LabelFrame(main_frame, text=" 步骤 1：数据源 ")
        step1_frame.pack(fill="x", pady=5, ipady=5)
        self.ent_f = ttk.Entry(step1_frame)
        self.ent_f.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=5)
        ttk.Button(step1_frame, text="📂 浏览文件", command=self.browse, width=12).pack(side="right", padx=10, pady=5)

        step2_frame = ttk.LabelFrame(main_frame, text=" 步骤 2：参数设置 ")
        step2_frame.pack(fill="x", pady=10, ipady=5)
        f_p1 = ttk.Frame(step2_frame)
        f_p1.pack(fill="x", padx=10, pady=5)
        ttk.Label(f_p1, text="📅 判断月份 (如 2026-01,2026-02):").pack(side="left")
        self.ent_m = ttk.Entry(f_p1, width=25)
        self.ent_m.insert(0, "2026-01")
        self.ent_m.pack(side="left", padx=10)

        f_p2 = ttk.Frame(step2_frame)
        f_p2.pack(fill="x", padx=10, pady=5)
        ttk.Label(f_p2, text="📆 特定日期 (如 1月1日,1月2日):").pack(side="left")
        self.ent_d = ttk.Entry(f_p2)
        self.ent_d.pack(side="left", fill="x", expand=True, padx=10)

        f_p3 = ttk.Frame(step2_frame)
        f_p3.pack(fill="x", padx=10, pady=5)
        ttk.Label(f_p3, text="🔍 自定义标备词 (选填,逗号隔开):").pack(side="left")
        self.ent_ckw = ttk.Entry(f_p3)
        self.ent_ckw.pack(side="left", fill="x", expand=True, padx=10)
        ttk.Label(f_p3, text="(填了则覆盖默认逻辑)", foreground="gray").pack(side="right")

        step3_frame = ttk.LabelFrame(main_frame, text=" 步骤 3：选择任务 ")
        step3_frame.pack(fill="x", pady=5, ipady=5)
        self.v_pl = tk.BooleanVar(value=False)
        self.v_pw = tk.BooleanVar(value=False)
        self.v_cl = tk.BooleanVar(value=False)
        self.v_cw = tk.BooleanVar(value=False)
        self.v_special = tk.BooleanVar(value=False)

        grid_frame = ttk.Frame(step3_frame)
        grid_frame.pack(fill="x", padx=15, pady=5)
        ttk.Checkbutton(grid_frame, text="省份 - 一维数据 (2个表)", variable=self.v_pl).grid(row=0, column=0, sticky="w", pady=2, padx=10)
        ttk.Checkbutton(grid_frame, text="省份 - 合并数据 (2个表)", variable=self.v_pw).grid(row=0, column=1, sticky="w", pady=2, padx=10)
        ttk.Checkbutton(grid_frame, text="地市 - 一维数据 (2个表)", variable=self.v_cl).grid(row=1, column=0, sticky="w", pady=2, padx=10)
        ttk.Checkbutton(grid_frame, text="地市 - 合并数据 (2个表)", variable=self.v_cw).grid(row=1, column=1, sticky="w", pady=2, padx=10)
        sep = ttk.Separator(grid_frame, orient='horizontal')
        sep.grid(row=2, column=0, columnspan=2, sticky="ew", pady=8)
        ttk.Checkbutton(grid_frame, text="★ 单独地市 (固定24地市列表)", variable=self.v_special, style='TCheckbutton').grid(row=3, column=0, columnspan=2, sticky="w", padx=10)

        self.btn = ttk.Button(main_frame, text="▶ 开始处理", command=self.run)
        self.btn.pack(pady=15, ipady=5, ipadx=20)

        log_frame = ttk.LabelFrame(main_frame, text=" 运行日志 ")
        log_frame.pack(fill="both", expand=True, pady=5)
        self.log_txt = scrolledtext.ScrolledText(log_frame, height=8, font=("Consolas", 9), bg="#f4f4f4", relief="flat")
        self.log_txt.pack(fill="both", expand=True, padx=5, pady=5)

    def log(self, msg):
        self.log_txt.config(state='normal')
        self.log_txt.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_txt.see(tk.END)
        self.log_txt.config(state='disabled')

    def browse(self):
        fs = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx *.xls *.csv")])
        if fs: self.ent_f.delete(0, tk.END); self.ent_f.insert(0, ";".join(fs))

    def run(self):
        threading.Thread(target=self.process).start()

    def process(self):
        files = self.ent_f.get().split(";")
        if not files or files[0] == "": return messagebox.showerror("错误", "请先选择数据文件！")
        
        start_time = time.time()
        self.btn.config(state="disabled", text="⏳ 处理中...")
        self.log_txt.config(state='normal')
        self.log_txt.delete(1.0, tk.END)
        self.log_txt.config(state='disabled')
        
        try:
            proc = UniversalProcessor(self.log)
            self.log("开始读取文件...")
            dfs = []
            for f in files:
                if f.strip(): dfs.append(pd.read_excel(f))
            if not dfs: raise ValueError("未读取到有效数据")
            full_df = pd.concat(dfs, ignore_index=True)
            
            df_clean, df_dp, df_dc = proc.process_step1(full_df, self.ent_m.get(), self.ent_ckw.get())
            
            base_dir = os.path.dirname(files[0])
            base_name = os.path.splitext(os.path.basename(files[0]))[0]
            
            df_dc.to_excel(os.path.join(base_dir, f"{base_name}_00_清洗后源数据.xlsx"), index=False)
            self.log("✅ 源数据已保存")

            m_ch = [("微信好友总量", None), ("线上平台", "线上平台"), ("线下平台", "线下活动"), ("考试现场", "考试现场"), ("高校", "高校"), ("其他", "其他")]
            m_on = [("微信好友总量", None), ("网站", "网站"), ("小红书", "小红书"), ("公众号", "公众号"), ("抖音", "抖音"), ("视频号", "视频号"), ("其他", "其他")]

            if self.v_pl.get():
                proc.gen_prov_long(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_01_省份一维_渠道.xlsx"), m_ch, False)
                proc.gen_prov_long(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_02_省份一维_线上.xlsx"), m_on, True)
                self.log("✅ 省份一维表已生成")

            if self.v_pw.get():
                proc.gen_prov_wide(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_03_省份合并_渠道.xlsx"), m_ch, False)
                proc.gen_prov_wide(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_04_省份合并_线上.xlsx"), m_on, True)
                self.log("✅ 省份宽表已生成")

            if self.v_cl.get():
                proc.gen_city_long(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_05_地市一维_渠道.xlsx"), m_ch, False)
                proc.gen_city_long(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_06_地市一维_线上.xlsx"), m_on, True)
                self.log("✅ 地市一维表已生成")

            if self.v_cw.get():
                proc.gen_city_wide(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_07_地市合并_渠道.xlsx"), m_ch, False)
                proc.gen_city_wide(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_08_地市合并_线上.xlsx"), m_on, True)
                self.log("✅ 地市宽表已生成")

            if self.v_special.get():
                proc.gen_special_city_report(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_09_单独地市_渠道.xlsx"), m_ch, False)
                proc.gen_special_city_report(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_10_单独地市_线上.xlsx"), m_on, True)
                self.log("✅ 单独地市表已生成")

            if self.ent_d.get().strip():
                proc.use_special_city = self.v_special.get()
                use_city = self.v_cl.get() or self.v_cw.get() or self.v_special.get()
                target_df = df_dc if use_city else df_dp
                proc.gen_date_summary(target_df, self.ent_d.get(), os.path.join(base_dir, f"{base_name}_11_日期汇总.xlsx"), use_city)

            duration = time.time() - start_time
            self.log(f"🎉 全部完成！总耗时: {duration:.2f}秒")
            messagebox.showinfo("成功", f"处理完成\n耗时: {duration:.2f} 秒")

        except Exception as e:
            self.log(f"❌ 错误: {e}")
            messagebox.showerror("错误", str(e))
        finally:
            self.btn.config(state="normal", text="▶ 开始处理")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
