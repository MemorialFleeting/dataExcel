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
# 配置区域：分校分类与固定顺序
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

BRANCH_GROUPS = {
    "A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"],
    "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"],
    "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]
}

# ==============================================================================
# 核心处理类
# ==============================================================================
class UniversalProcessor:
    def __init__(self, log_func):
        self.log_callback = log_func
        self.use_special_city_logic = False

    def log(self, message):
        """带有 UI 强制刷新的日志输出"""
        self.log_callback(message)

    def excel_lookup_find(self, text, keywords, results, default_value):
        """模拟 Excel 的 LOOKUP FIND 逻辑，用于识别分校"""
        if pd.isna(text):
            return default_value
        
        text_string = str(text)
        for keyword, result in zip(reversed(keywords), reversed(results)):
            if keyword in text_string:
                return result
        return default_value

    def check_keyword_flag(self, text, keywords):
        """检查文本是否包含关键词列表中的任何一个"""
        if pd.isna(text):
            return ""
        
        text_string = str(text)
        for keyword in keywords:
            if keyword in text_string:
                return 1
        return ""

    def extract_city_smart(self, dept_path):
        """从负责人所属部门路径中智能提取地市名称"""
        if pd.isna(dept_path):
            return "未知"
        
        text_string = str(dept_path).replace('\r', '\n')
        lines = [line.strip() for line in text_string.split('\n') if line.strip()]
        if not lines:
            return "其他"
            
        target_line = ""
        for line in lines:
            if "地市分校" in line:
                target_line = line
                break
        if not target_line:
            for line in lines:
                if "分校" in line:
                    target_line = line
                    break
        if not target_line:
            target_line = lines[0]
            
        clean_line = target_line.replace('"', '').replace("'", "").replace("华图教育", "").strip()
        parts = [part.strip() for part in clean_line.split('/') if part.strip()]
        if not parts:
            return "其他"
            
        final_city = ""
        for part in parts:
            if "地市分校" in part and part != "地市分校":
                final_city = part.replace("地市分校", "")
                break
        
        if not final_city:
            final_city = parts[-1]
            
        suffixes = ["区属学习中心", "高校学习中心", "学习中心", "办事处", "地市分校", "分校", "分部"]
        for suffix in suffixes:
            if final_city.endswith(suffix):
                final_city = final_city[:-len(suffix)]
        return final_city

    def standardize_city_name(self, city_name):
        """标准化地市名称，去除市字和空白符"""
        if pd.isna(city_name):
            return "其他"
        
        standard_name = str(city_name).strip()
        standard_name = re.sub(r'[\s\r\n\t\u200b\ufeff\xa0]+', '', standard_name)
        if len(standard_name) > 2 and standard_name.endswith("市"):
            standard_name = standard_name[:-1]
        return standard_name

    def process_step1(self, source_df, date_keyword, custom_total_kws_str):
        """步骤1：全量数据清洗与业务标签计算"""
        self.log("【数据清洗】正在提取地市与计算业务标签...")
        
        # 统一表头
        source_df.columns = [str(column).replace('\ufeff', '').strip() for column in source_df.columns]
        
        # 识别分校
        branch_keywords = ["通辽分校", "陕西分校", "湖北分校", "辽宁分校", "河北分校", "甘肃分校", "厦门分校", "福建分校", "山东分校", "北京分校", "安徽分校", "黑龙江分校", "吉林分校", "江苏分校", "重庆分校", "广东分校", "天津分校", "河南分校", "云南分校", "江西分校", "湖南分校", "贵州分校", "广西分校", "山西分校", "宁夏分校", "内蒙古分校", "浙江分校", "新疆分校", "青海分校", "南疆分校", "四川分校", "上海分校", "海南分校", "西藏分校", "赤峰"]
        branch_results = ["内蒙古分校", "陕西分校", "湖北分校", "辽宁分校", "河北分校", "甘肃分校", "厦门分校", "福建分校", "山东分校", "北京分校", "安徽分校", "黑龙江分校", "吉林分校", "江苏分校", "重庆分校", "广东分校", "天津分校", "河南分校", "云南分校", "江西分校", "湖南分校", "贵州分校", "广西分校", "山西分校", "宁夏分校", "内蒙古分校", "浙江分校", "新疆分校", "青海分校", "新疆分校", "四川分校", "上海分校", "海南分校", "西藏分校", "内蒙古分校"]
        
        source_df['分校'] = source_df['负责人所属部门'].apply(lambda x: self.excel_lookup_find(x, branch_keywords, branch_results, "总部"))
        source_df['地市'] = source_df['负责人所属部门'].apply(self.extract_city_smart).apply(self.standardize_city_name)

        def determine_category(row):
            dept_raw = str(row['负责人所属部门'])
            if "地市分校" in dept_raw:
                return "地市"
            if str(row['分校']) in ["北京分校", "天津分校"] and "各校区" in dept_raw:
                return "地市"
            return "其它"
        source_df['所属类别'] = source_df.apply(determine_category, axis=1)

        # 核心业务标签
        source_df['公职'] = source_df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["26公职","27公职","28公职","29公职"]))
        kw_shiye = ["26事业","26三支","26社区","26辅警","26书记员","26国企","30事业","30三支","30社区","30辅警","30书记员","30国企", "27事业","27三支","27社区","27辅警","27书记员","27国企","28事业","28三支","28社区","28辅警","28书记员","28国企", "29事业","29三支","29社区","29辅警","29书记员","29国企"]
        source_df['事业辅助列'] = source_df['备注名'].apply(lambda x: self.check_keyword_flag(x, kw_shiye))
        source_df['教师'] = source_df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["26教师","26特岗","26教资","30教师","30特岗","30教资","27教师","27特岗","27教资","28教师","28特岗","28教资","29教师","29特岗","29教资"]))
        source_df['文职'] = source_df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30文职","26文职","27文职","28文职","29文职"]))
        source_df['医疗'] = source_df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30医疗","26医疗","27医疗","28医疗","29医疗"]))
        source_df['银行'] = source_df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30银行","26银行","27银行","28银行","29银行"]))
        source_df['考研'] = source_df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30考研","26考研","27考研","28考研","29考研"]))
        source_df['学历'] = source_df['备注名'].apply(lambda x: self.check_keyword_flag(x, ["30学历","26学历","27学历","28学历","29学历"]))
        source_df['其他'] = ""
        
        if custom_total_kws_str.strip():
            custom_keywords = [k.strip() for k in custom_total_kws_str.replace('，', ',').split(',') if k.strip()]
            def check_custom_total(remark):
                if pd.isna(remark): return 0
                return 1 if any(kw in str(remark) for kw in custom_keywords) else 0
            source_df['标备总数'] = source_df['备注名'].apply(check_custom_total)
        else:
            def check_default_total(row):
                standard_cols = ['公职', '事业辅助列', '教师', '文职', '医疗', '银行', '考研', '学历']
                for col in standard_cols:
                    if row[col] == 1:
                        return 1
                return 0
            source_df['标备总数'] = source_df.apply(check_default_total, axis=1)

        date_kws = [k.strip() for k in date_keyword.replace('，', ',').split(',') if k.strip()]
        if not date_kws:
            date_kws = ["FAIL_SAFE"]
            
        source_df['！！！是否新增'] = source_df['创建时间'].apply(lambda x: "是" if any(k in str(x) for k in date_kws) else "否")
        source_df['日期'] = source_df['好友添加时间'].apply(lambda x: f"{pd.to_datetime(x).month}月{pd.to_datetime(x).day}日" if pd.notna(x) else "")
        source_df['渠道'] = source_df['渠道活码分组'].apply(lambda x: self.excel_lookup_find(x, ["网络","社会","高校","线上","现场","线下"], ["线上平台","线下活动","高校","线上平台","考试现场","线下活动"], "其他"))
        source_df['线上渠道'] = source_df['渠道活码分组'].apply(lambda x: self.excel_lookup_find(x, ["网站","公众号","小红书","视频号","抖音"], ["网站","公众号","小红书","视频号","抖音"], "其他"))
        source_df['事业2'] = source_df.apply(lambda row: 1 if row['公职'] != 1 and row['事业辅助列'] == 1 else 0, axis=1)
        
        def check_valid_reply(row):
            has_reply_time = pd.notna(row['客户上次回复时间']) and str(row['客户上次回复时间']).strip() != ""
            if has_reply_time and row['标备总数'] == 1:
                return "是"
            return "否"
        source_df['客户回话'] = source_df.apply(check_valid_reply, axis=1)

        # 5个质检原生字段强制挂载
        qc_original_fields = ['好友添加来源', '添加渠道码', '员工发送消息数', '客户回复消息数', '是否同意会话存档']
        for field in qc_original_fields:
            if field not in source_df.columns:
                if '数' in field:
                    source_df[field] = 0
                else:
                    source_df[field] = "未知"

        if 'ExternalUserId' not in source_df.columns:
            if len(source_df.columns) > 1:
                source_df.rename(columns={source_df.columns[1]: 'ExternalUserId'}, inplace=True)

        final_columns = [
            '备注名', 'ExternalUserId', '分校', '地市', '所属类别', 
            '公职', '事业2', '教师', '文职', '医疗', '银行', '考研', '学历', '其他', 
            '标备总数', '！！！是否新增', '渠道', '线上渠道', '客户回话', '日期', 
            '添加好友状态', '好友添加来源', '添加渠道码', '员工发送消息数', '客户回复消息数', '是否同意会话存档'
        ]
        
        actual_extract_cols = [c for c in final_columns if c in source_df.columns]
        df_clean = source_df[actual_extract_cols].copy()
        
        for num_col in ['公职', '事业2', '标备总数', '员工发送消息数', '客户回复消息数']:
            if num_col in df_clean.columns:
                df_clean[num_col] = pd.to_numeric(df_clean[num_col], errors='coerce').fillna(0)

        self.log("数据去重逻辑计算中...")
        df_dedup_province = df_clean.drop_duplicates(subset=['ExternalUserId', '分校'], keep='first')
        df_dedup_city = df_clean.drop_duplicates(subset=['ExternalUserId', '分校', '地市'], keep='first')
        df_retention_all = df_clean[df_clean['添加好友状态'].astype(str).str.contains("已添加", na=False)].copy()

        return df_clean, df_dedup_province, df_dedup_city, df_retention_all

    # --- 统计辅助函数 ---
    def calculate_summary_stats(self, raw_df, dedup_df):
        """计算统计数据，生成统一名称的字典"""
        new_added_subset = dedup_df[dedup_df['！！！是否新增'].astype(str).str.contains("是", na=False)]
        
        stats_dict = {
            "raw_count": len(raw_df),
            "dedup_count": len(dedup_df),
            "new_added_count": len(new_added_subset),
            "gz_sum": new_added_subset['公职'].sum(),
            "sy_sum": new_added_subset['事业2'].sum(),
            "total_tb_sum": new_added_subset['标备总数'].sum(),
            "reply_count": len(new_added_subset[new_added_subset['客户回话'].astype(str).str.contains("是", na=False)])
        }
        return stats_dict

    def format_output_row_list(self, stats, is_province_mode=False):
        """格式化输出行，比率取整"""
        def safe_div(a, b):
            return a / b if b != 0 else 0.0
            
        def to_pct(v):
            return f"{int(round(v * 100))}%"

        rep_rate = safe_div((stats["dedup_count"] - stats["new_added_count"]), stats["dedup_count"])
        other_tb = stats["total_tb_sum"] - stats["gz_sum"] - stats["sy_sum"]
        tb_rate = safe_div(stats["total_tb_sum"], stats["new_added_count"])
        reply_rate = safe_div(stats["reply_count"], stats["new_added_count"])
        
        return [
            int(stats["raw_count"]), int(stats["dedup_count"]), int(stats["new_added_count"]),
            to_pct(rep_rate), int(stats["gz_sum"]), int(stats["sy_sum"]),
            int(other_tb), int(stats["total_tb_sum"]), to_pct(tb_rate),
            int(stats["reply_count"]), to_pct(reply_rate)
        ]

    def calc_retention_stats_only(self, retention_df_subset):
        """留存版统计"""
        if len(retention_df_subset) == 0:
            return {"gz": 0, "sy": 0, "tb": 0, "ret": 0}
        
        gz_sum = retention_df_subset['公职'].sum()
        sy_sum = retention_df_subset['事业2'].sum()
        tb_sum = retention_df_subset['标备总数'].sum()
        
        return {
            "gz": gz_sum, "sy": sy_sum, "tb": tb_sum, "ret": len(retention_df_subset)
        }

    def format_retention_row_list(self, stats_dict):
        """留存版输出行格式化"""
        def safe_div(a, b):
            return a / b if b != 0 else 0.0
        def to_pct(v):
            return f"{int(round(v * 100))}%"
            
        other_tb = stats_dict["tb"] - stats_dict["gz"] - stats_dict["sy"]
        tb_rate = safe_div(stats_dict["tb"], stats_dict["ret"])
        
        return [
            int(stats_dict["ret"]), int(stats_dict["gz"]), int(stats_dict["sy"]),
            int(other_tb), int(stats_dict["tb"]), to_pct(tb_rate)
        ]

    # --- 常规报表业务逻辑 (01-09) ---
    def gen_prov_long_merged(self, df_raw, df_dedup, output_file):
        self.log(f"生成报表 [01_省份一维]: {os.path.basename(output_file)}")
        groups = BRANCH_GROUPS
        iter_list = [("线上平台", "网站", "网站"), ("线上平台", "小红书", "小红书"), ("线上平台", "公众号", "公众号"), ("线上平台", "抖音", "抖音"), ("线上平台", "视频号", "视频号"), ("线上平台", "其他", "其他"), ("线下活动", "线下活动", "线下活动"), ("考试现场", "考试现场", "考试现场"), ("高校", "高校", "高校"), ("其他", "其他", "其他")]
        
        final_rows = []
        # 初始化全国汇总
        grand_national_accumulator = {key: {"raw_count":0, "dedup_count":0,"new_added_count":0,"gz_sum":0,"sy_sum":0,"total_tb_sum":0,"reply_count":0} for key in [x[2] for x in iter_list]}
        
        for g_name, branches in groups.items():
            group_accumulator = {key: {"raw_count":0, "dedup_count":0,"new_added_count":0,"gz_sum":0,"sy_sum":0,"total_tb_sum":0,"reply_count":0} for key in [x[2] for x in iter_list]}
            for branch in branches:
                branch_raw = df_raw[df_raw['分校'] == branch]
                branch_dedup = df_dedup[df_dedup['分校'] == branch]
                for p_name, f_kw, s_name in iter_list:
                    if p_name == "线上平台":
                        current_raw = branch_raw[(branch_raw['渠道'] == '线上平台') & (branch_raw['线上渠道'] == f_kw)]
                        current_dedup = branch_dedup[(branch_dedup['渠道'] == '线上平台') & (branch_dedup['线上渠道'] == f_kw)]
                    else:
                        current_raw = branch_raw[branch_raw['渠道'] == f_kw]
                        current_dedup = branch_dedup[branch_dedup['渠道'] == f_kw]
                    
                    stats = self.calculate_summary_stats(current_raw, current_dedup)
                    # 累加统计数据
                    for key_name in stats:
                        group_accumulator[s_name][key_name] += stats[key_name]
                        grand_national_accumulator[s_name][key_name] += stats[key_name]
                    
                    final_rows.append([branch, p_name, s_name] + self.format_output_row_list(stats, True))
            
            # 写入类别总计行
            for p_name, f_kw, s_name in iter_list:
                final_rows.append([f"{g_name}类总计", p_name, s_name] + self.format_output_row_list(group_accumulator[s_name], True))
        
        # 写入全国总计行
        for p_name, f_kw, s_name in iter_list:
            final_rows.append(["全国", p_name, s_name] + self.format_output_row_list(grand_national_accumulator[s_name], True))
            
        pd.DataFrame(final_rows, columns=["分校", "平台", "渠道", "新增", "去重", "净增", "重率", "公职", "事业", "其他", "标备总额", "标备率", "回话", "回话率"]).to_excel(output_file, index=False)
        self._style_excel(output_file)

    def gen_prov_wide(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成报表 [省份宽表]: {os.path.basename(output_file)}")
        groups = BRANCH_GROUPS
        final_rows = []
        grand_national_stats = {ct[0]: {"raw_count":0, "dedup_count":0,"new_added_count":0,"gz_sum":0,"sy_sum":0,"total_tb_sum":0,"reply_count":0} for ct in channel_map}
        
        for group_name, branches in groups.items():
            group_stats = {ct[0]: {"raw_count":0, "dedup_count":0,"new_added_count":0,"gz_sum":0,"sy_sum":0,"total_tb_sum":0,"reply_count":0} for ct in channel_map}
            for branch in branches:
                branch_raw = df_raw[df_raw['分校'] == branch]
                branch_dedup = df_dedup[df_dedup['分校'] == branch]
                row_data = [branch]
                for ct_title, cf_keyword in channel_map:
                    if cf_keyword:
                        if strict_online:
                            c_raw = branch_raw[(branch_raw['渠道']=='线上平台')&(branch_raw['线上渠道']==cf_keyword)]
                            c_dedup = branch_dedup[(branch_dedup['渠道']=='线上平台')&(branch_dedup['线上渠道']==cf_keyword)]
                        else:
                            c_raw = branch_raw[branch_raw['渠道']==cf_keyword]
                            c_dedup = branch_dedup[branch_dedup['渠道']==cf_keyword]
                    else:
                        c_raw = branch_raw
                        c_dedup = branch_dedup
                        
                    stats = self.calculate_summary_stats(c_raw, c_dedup)
                    for key_name in stats:
                        group_stats[ct_title][key_name] += stats[key_name]
                        grand_national_stats[ct_title][key_name] += stats[key_name]
                    row_data.extend(self.format_output_row_list(stats, True))
                final_rows.append(row_data)
            
            group_summary_row = [f"{group_name}类总计"]
            for ct_title, _ in channel_map:
                group_summary_row.extend(self.format_output_row_list(group_stats[ct_title], True))
            final_rows.append(group_summary_row)
            
        national_row = ["全国"]
        for ct_title, _ in channel_map:
            national_row.extend(self.format_output_row_list(grand_national_stats[ct_title], True))
        final_rows.append(national_row)
        self._write_wide_excel_with_headers(output_file, final_rows, channel_map, 2, False)

    def gen_city_long_merged(self, df_raw, df_dedup, df_ret, output_file, do_ret):
        self.log(f"生成报表 [04_地市一维]: {os.path.basename(output_file)}")
        iter_list = [("线上平台", "网站", "网站"), ("线上平台", "小红书", "小红书"), ("线上平台", "公众号", "公众号"), ("线上平台", "抖音", "抖音"), ("线上平台", "视频号", "视频号"), ("线上平台", "其他", "其他"), ("线下活动", "线下活动", "线下活动"), ("考试现场", "考试现场", "考试现场"), ("高校", "高校", "高校"), ("其他", "其他", "其他")]
        final_rows = []
        grand_national_accumulator = {key: {"raw_count":0, "dedup_count":0,"new_added_count":0,"gz_sum":0,"sy_sum":0,"total_tb_sum":0,"reply_count":0} for key in [x[2] for x in iter_list]}
        grand_retention_accumulator = {key: {"gz": 0, "sy": 0, "tb": 0, "ret": 0} for key in [x[2] for x in iter_list]}
        
        for branch_name in sorted(df_raw['分校'].dropna().unique()):
            branch_raw_subset = df_raw[df_raw['分校'] == branch_name]
            branch_dedup_subset = df_dedup[df_dedup['分校'] == branch_name]
            branch_retention_subset = df_ret[df_ret['分校'] == branch_name] if do_ret else None
            
            for city_name in sorted(branch_raw_subset['地市'].dropna().unique()):
                city_raw = branch_raw_subset[branch_raw_subset['地市'] == city_name]
                city_dedup = branch_dedup_subset[branch_dedup_subset['地市'] == city_name]
                category = city_raw['所属类别'].iloc[0] if not city_raw.empty else "其它"
                
                for p_name, f_kw, s_name in iter_list:
                    if p_name == "线上平台":
                        current_raw = city_raw[(city_raw['渠道'] == '线上平台') & (city_raw['线上渠道'] == f_kw)]
                        current_dedup = city_dedup[(city_dedup['渠道'] == '线上平台') & (city_dedup['线上渠道'] == f_kw)]
                    else:
                        current_raw = city_raw[city_raw['渠道'] == f_kw]
                        current_dedup = city_dedup[city_dedup['渠道'] == f_kw]
                    
                    stats = self.calculate_summary_stats(current_raw, current_dedup)
                    for key_name in stats:
                        grand_national_accumulator[s_name][key_name] += stats[key_name]
                    
                    retention_cols = []
                    if do_ret:
                        city_retention_base = branch_retention_subset[branch_retention_subset['地市'] == city_name]
                        if p_name == "线上平台":
                            city_retention_final = city_retention_base[(city_retention_base['渠道'] == '线上平台') & (city_retention_base['线上渠道'] == f_kw)]
                        else:
                            city_retention_final = city_retention_base[city_retention_base['渠道'] == f_kw]
                        
                        rs_stats = self.calc_retention_stats_only(city_retention_final)
                        for key_name in rs_stats:
                            grand_retention_accumulator[s_name][key_name] += rs_stats[key_name]
                        retention_cols = [int(rs_stats["ret"]), int(rs_stats["gz"]), int(rs_stats["sy"]), int(rs_stats["tb"]-rs_stats["gz"]-rs_stats["sy"]), int(rs_stats["tb"])]
                    
                    final_rows.append([branch_name, city_name, category, p_name, s_name] + self.format_output_row_list(stats, False) + retention_cols)
        
        # 写入全国维度
        for p_name, f_kw, s_name in iter_list:
            national_row = ["全国", "总计", "-", p_name, s_name] + self.format_output_row_list(grand_national_accumulator[s_name], False)
            if do_ret:
                rs_national = grand_retention_accumulator[s_name]
                national_row += [int(rs_national["ret"]), int(rs_national["gz"]), int(rs_national["sy"]), int(rs_national["tb"]-rs_national["gz"]-rs_national["sy"]), int(rs_national["tb"])]
            final_rows.append(national_row)
            
        headers = ["分校", "地市", "类别", "平台", "渠道", "新增", "去重", "净增", "重率", "公职", "事业", "其他", "总量", "标备率", "回话", "回话率"]
        if do_ret:
            headers += ["留存量", "总公职", "总事业", "总其他", "总标备"]
        pd.DataFrame(final_rows, columns=headers).to_excel(output_file, index=False)
        self._style_excel(output_file)

    def gen_city_wide(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成报表 [地市宽表]: {os.path.basename(output_file)}")
        final_rows = []
        grand_national_accumulator = {ct[0]: {"raw_count":0, "dedup_count":0,"new_added_count":0,"gz_sum":0,"sy_sum":0,"total_tb_sum":0,"reply_count":0} for ct in channel_map}
        for branch_name in sorted(df_raw['分校'].dropna().unique()):
            br_raw = df_raw[df_raw['分校'] == branch_name]
            br_dedup = df_dedup[df_dedup['分校'] == branch_name]
            for city_name in sorted(br_raw['地市'].dropna().unique()):
                row_data = [branch_name, city_name]
                city_raw = br_raw[br_raw['地市'] == city_name]
                city_dedup = br_dedup[br_dedup['地市'] == city_name]
                for ct_title, cf_keyword in channel_map:
                    if cf_keyword:
                        if strict_online:
                            current_raw = city_raw[(city_raw['渠道']=='线上平台')&(city_raw['线上渠道']==cf_keyword)]
                            current_dedup = city_dedup[(city_dedup['渠道']=='线上平台')&(city_dedup['线上渠道']==cf_keyword)]
                        else:
                            current_raw = city_raw[city_raw['渠道']==cf_keyword]
                            current_dedup = city_dedup[city_dedup['渠道']==cf_keyword]
                    else:
                        current_raw = city_raw
                        current_dedup = city_dedup
                    
                    stats = self.calculate_summary_stats(current_raw, current_dedup)
                    for key_name in stats:
                        grand_national_accumulator[ct_title][key_name] += stats[key_name]
                    row_data.extend(self.format_output_row_list(stats, False))
                final_rows.append(row_data)
        
        national_row = ["全国", "总计"]
        for ct_title, _ in channel_map:
            national_row.extend(self.format_output_row_list(grand_national_accumulator[ct_title], False))
        final_rows.append(national_row)
        self._write_wide_excel_with_headers(output_file, final_rows, channel_map, 3, True)

    def gen_special_city_report(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成报表 [独立24地市]: {os.path.basename(output_file)}")
        final_rows = []
        grand_national_stats = {ct[0]: {"raw_count":0, "dedup_count":0,"new_added_count":0,"gz_sum":0,"sy_sum":0,"total_tb_sum":0,"reply_count":0} for ct in channel_map}
        for city in FIXED_ORDER_CITIES:
            br_raw = df_raw[df_raw['地市'] == city]
            br_dedup = df_dedup[df_dedup['地市'] == city]
            branch_name = br_raw['分校'].iloc[0] if not br_raw.empty else DEFAULT_BRANCH_MAP.get(city, "未知分校")
            row = [branch_name, city]
            for ct_title, cf_keyword in channel_map:
                if cf_keyword:
                    if strict_online:
                        c_raw = br_raw[(br_raw['渠道']=='线上平台')&(br_raw['线上渠道']==cf_keyword)]
                        c_dedup = br_dedup[(br_dedup['渠道']=='线上平台')&(br_dedup['线上渠道']==cf_keyword)]
                    else:
                        c_raw = br_raw[br_raw['渠道']==cf_keyword]
                        c_dedup = br_dedup[br_dedup['渠道']==cf_keyword]
                else:
                    c_raw = br_raw
                    c_dedup = br_dedup
                
                stats = self.calculate_summary_stats(c_raw, c_dedup)
                for key_name in stats:
                    grand_national_stats[ct_title][key_name] += stats[key_name]
                row.extend(self.format_output_row_list(stats, False))
            final_rows.append(row)
            
        national_row = ["全国", "总计"]
        for ct_title, _ in channel_map:
            national_row.extend(self.format_output_row_list(grand_national_stats[ct_title], False))
        final_rows.append(national_row)
        self._write_wide_excel_with_headers(output_file, final_rows, channel_map, 3, True)

    def gen_wide_retention_tables(self, df_ret, output_file, channel_map, strict_online, is_prov, is_spec=False):
        self.log(f"生成报表 [留存版看板]: {os.path.basename(output_file)}")
        final_rows = []
        grand_national_accumulator = {ct[0]: {"gz":0,"sy":0,"tb":0,"ret":0} for ct in channel_map}
        
        if is_spec:
            for city in FIXED_ORDER_CITIES:
                city_ret = df_ret[df_ret['地市'] == city]
                branch_name = city_ret['分校'].iloc[0] if not city_ret.empty else DEFAULT_BRANCH_MAP.get(city, "未知分校")
                row = [branch_name, city]
                for ct_title, cf_keyword in channel_map:
                    if cf_keyword:
                        if strict_online:
                            current_ret = city_ret[(city_ret['渠道']=='线上平台')&(city_ret['线上渠道']==cf_keyword)]
                        else:
                            current_ret = city_ret[city_ret['渠道']==cf_keyword]
                    else:
                        current_ret = city_ret
                    
                    stats = self.calc_retention_stats_only(current_ret)
                    for key in stats:
                        grand_national_accumulator[ct_title][key] += stats[key]
                    row.extend(self.format_retention_row_list(stats))
                final_rows.append(row)
        elif is_prov:
            groups = BRANCH_GROUPS
            for gn, bns in groups.items():
                group_stats_accumulator = {ct[0]: {"gz":0,"sy":0,"tb":0,"ret":0} for ct in channel_map}
                for branch_n in bns:
                    branch_ret = df_ret[df_ret['分校'] == branch_n]
                    row = [branch_n]
                    for ct_title, cf_keyword in channel_map:
                        if cf_keyword:
                            if strict_online:
                                current_ret = branch_ret[(branch_ret['渠道']=='线上平台')&(branch_ret['线上渠道']==cf_keyword)]
                            else:
                                current_ret = branch_ret[branch_ret['渠道']==cf_keyword]
                        else:
                            current_ret = branch_ret
                        
                        stats = self.calc_retention_stats_only(current_ret)
                        for key in stats:
                            group_stats_accumulator[ct_title][key] += stats[key]
                            grand_national_accumulator[ct_title][key] += stats[key]
                        row.extend(self.format_retention_row_list(stats))
                    final_rows.append(row)
                
                gr_row = [f"{gn}类总计"]
                for ct_title, _ in channel_map:
                    gr_row.extend(self.format_retention_row_list(group_stats_accumulator[ct_title]))
                final_rows.append(gr_row)
        else:
            for branch_n in sorted(df_ret['分校'].dropna().unique()):
                branch_ret = df_ret[df_ret['分校'] == branch_n]
                for city_n in sorted(branch_ret['地市'].dropna().unique()):
                    city_ret_subset = branch_ret[branch_ret['地市'] == city_n]
                    row = [branch_n, city_n]
                    for ct_title, cf_keyword in channel_map:
                        if cf_keyword:
                            if strict_online:
                                current_ret = city_ret_subset[(city_ret_subset['渠道']=='线上平台')&(city_ret_subset['线上渠道']==cf_keyword)]
                            else:
                                current_ret = city_ret_subset[city_ret_subset['渠道']==cf_keyword]
                        else:
                            current_ret = city_ret_subset
                        
                        stats = self.calc_retention_stats_only(current_ret)
                        for key in stats:
                            grand_national_accumulator[ct_title][key] += stats[key]
                        row.extend(self.format_retention_row_list(stats))
                    final_rows.append(row)
                    
        final_national_row = ["全国", "总计"] if not is_prov else ["全国"]
        for ct_title, _ in channel_map:
            final_national_row.extend(self.format_retention_row_list(grand_national_accumulator[ct_title]))
        final_rows.append(final_national_row)
        self._write_wide_excel_retention_structured(output_file, final_rows, channel_map, 2 if is_prov else 3, not is_prov)

    def gen_date_summary_report(self, df_dedup, d_str, output_file, is_city):
        if not d_str.strip():
            return
        self.log(f"生成报表 [09_日期汇总]: {os.path.basename(output_file)}")
        target_dates = [d.strip() for d in d_str.replace('，', ',').split(',') if d.strip()]
        filtered_df = df_dedup[df_dedup['日期'].isin(target_dates)]
        group_keys = ['分校', '地市'] if is_city else ['分校']
        
        count_all = filtered_df.groupby(group_keys).size().reset_index(name='total')
        count_web = filtered_df[filtered_df['线上渠道'] == '网站'].groupby(group_keys).size().reset_index(name='web')
        summary_result = pd.merge(count_all, count_web, on=group_keys, how='left').fillna(0)
        
        sum_total = summary_result['total'].sum()
        sum_web = summary_result['web'].sum()
        
        national_summary = {col: "" for col in summary_result.columns}
        national_summary['分校'] = "全国"
        if is_city: national_summary['地市'] = "总计"
        national_summary['total'] = sum_total
        national_summary['web'] = sum_web
        
        summary_result = pd.concat([summary_result, pd.DataFrame([national_summary])], ignore_index=True)
        summary_result.rename(columns={'total': f"汇总({','.join(target_dates)})", 'web': "其中:网站"}, inplace=True)
        summary_result.to_excel(output_file, index=False)
        self._style_excel(output_file)

    # --- 质检核心模块 4.20 ---
    def gen_quality_inspection_suite(self, df_clean, df_retention, output_file, remark_exclude, channel_exclude, qc_dates_str):
        self.log(">>> 开始质检综合报表生成业务流...")
        base_filename = os.path.splitext(output_file)[0]
        
        # 1. 导出全量留存明细底稿
        df_retention.to_csv(f"{base_filename}_表1_详情明细底稿.csv", index=False, encoding='utf-8-sig')

        # 2. 准备日期过滤池（受日期限制）
        target_retention_pool = df_retention.copy()
        target_clean_pool = df_clean.copy()
        if qc_dates_str.strip():
            date_list = [d.strip() for d in qc_dates_str.replace('，', ',').split(',') if d.strip()]
            target_retention_pool = target_retention_pool[target_retention_pool['日期'].isin(date_list)]
            target_clean_pool = target_clean_pool[target_clean_pool['日期'].isin(date_list)]

        # 3. 虚假备注判定逻辑
        df_abnormal_pool = target_retention_pool[
            (target_retention_pool['标备总数'] == 1) & 
            (target_retention_pool['客户回复消息数'] == 0) & 
            (target_retention_pool['好友添加来源'].astype(str).str.contains('扫描渠道二维码', na=False))
        ].copy()
        
        if remark_exclude.strip():
            rk_list = [k.strip() for k in remark_exclude.replace('，', ',').split(',') if k.strip()]
            def filter_remark_fn(val):
                return any(k in str(val) for k in rk_list)
            df_abnormal_pool = df_abnormal_pool[~df_abnormal_pool['备注名'].apply(filter_remark_fn)]
            
        if channel_exclude.strip():
            ck_list = [k.strip() for k in channel_exclude.replace('，', ',').split(',') if k.strip()]
            def filter_channel_fn(val):
                return any(k in str(val) for k in ck_list)
            df_abnormal_pool = df_abnormal_pool[~df_abnormal_pool['添加渠道码'].apply(filter_channel_fn)]

        # 导出表2明细
        df_fake_details_final = df_abnormal_pool[df_abnormal_pool['是否同意会话存档'].astype(str).str.strip() == '同意'].copy()
        df_fake_details_final.to_csv(f"{base_filename}_表2_虚假备注详情明细.csv", index=False, encoding='utf-8-sig')
        self.log(f"✅ 质检表2明细已生成: {len(df_fake_details_final)} 行")

        groups = BRANCH_GROUPS
        s_div_fn = lambda a, b: a / b if b else 0.0
        t_pct_fn = lambda v: f"{int(round(v * 100))}%"

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet 1: 表2汇总
            rows_t2 = []
            grand_total_t2 = 0
            for group_n, branches in groups.items():
                group_total = 0
                for b_name in branches:
                    c = len(df_fake_details_final[df_fake_details_final['分校'] == b_name])
                    group_total += c
                    rows_t2.append([b_name, c])
                rows_t2.append([f"{group_n}类总计", group_total])
                grand_total_t2 += group_total
            rows_t2.append(["全国总计", grand_total_t2])
            pd.DataFrame(rows_t2, columns=['分校','虚假备注']).to_excel(writer, sheet_name='2_虚假备注筛选', index=False)

            # Sheet 2: 跨分校逻辑
            df3_base_dedup = target_retention_pool.drop_duplicates(subset=['ExternalUserId', '分校'])
            df3_counts = df3_base_dedup.groupby('ExternalUserId').size().reset_index(name='跨分校')
            df3_counts.to_csv(f"{base_filename}_表3_加几个分校明细底稿.csv", index=False, encoding='utf-8-sig')
            
            abnormal_id_list = df3_counts[df3_counts['跨分校'] >= 10]['ExternalUserId']
            df4_pool_all_adds = target_retention_pool[target_retention_pool['ExternalUserId'].isin(abnormal_id_list)]
            rows_t4 = []
            grand_total_t4 = 0
            for group_n, branches in groups.items():
                group_total = 0
                for b_name in branches:
                    c = len(df4_pool_all_adds[df4_pool_all_adds['分校'] == b_name])
                    group_total += c
                    rows_t4.append([b_name, c])
                rows_t4.append([f"{group_n}类总计", group_total])
                grand_total_t4 += group_total
            rows_t4.append(["全国总计", grand_total_t4])
            pd.DataFrame(rows_t4, columns=['分校','添加客服量']).to_excel(writer, sheet_name='4_跨10个以上分校', index=False)

            # Sheet 3: 会话看板
            rows_t5 = []
            national_acc_t5 = {'ret':0,'has':0,'no':0,'db':0}
            for group_n, branches in groups.items():
                group_acc_t5 = {'ret':0,'has':0,'no':0,'db':0}
                for b_name in branches:
                    grp = target_retention_pool[target_retention_pool['分校'] == b_name]
                    ret = len(grp)
                    has = (grp['客户回复消息数'] > 0).sum()
                    no = ((grp['客户回复消息数'] == 0) & (grp['是否同意会话存档'].astype(str).str.strip() == '同意')).sum()
                    db = ((grp['员工发送消息数'] == 0) & (grp['客户回复消息数'] == 0) & (grp['是否同意会话存档'].astype(str).str.strip() == '同意')).sum()
                    
                    group_acc_t5['ret']+=ret; group_acc_t5['has']+=has; group_acc_t5['no']+=no; group_acc_t5['db']+=db
                    national_acc_t5['ret']+=ret; national_acc_t5['has']+=has; national_acc_t5['no']+=no; national_acc_t5['db']+=db
                    
                    rows_t5.append([b_name, ret, has, no, t_pct_fn(s_div_fn(no, ret)), db, t_pct_fn(s_div_fn(db, ret))])
                rows_t5.append([f"{group_n}类总计", group_acc_t5['ret'], group_acc_t5['has'], group_acc_t5['no'], t_pct_fn(s_div_fn(group_acc_t5['no'], group_acc_t5['ret'])), group_acc_t5['db'], t_pct_fn(s_div_fn(group_acc_t5['db'], group_acc_t5['ret']))])
            rows_t5.append(["全国总计", national_acc_t5['ret'], national_acc_t5['has'], national_acc_t5['no'], t_pct_fn(s_div_fn(national_acc_t5['no'], national_acc_t5['ret'])), national_acc_t5['db'], t_pct_fn(s_div_fn(national_acc_t5['db'], national_acc_t5['ret']))])
            pd.DataFrame(rows_t5, columns=['分校','留存量','有会话','0会话','0会话率','双向无会话','双无率']).to_excel(writer, sheet_name='5_会话情况', index=False)

            # Sheet 4: 来源分布
            source_names = ["获客助手", "名片分享", "其他", "群聊", "扫描名片二维码", "扫描渠道二维码", "搜索手机号", "微信联系人", "未知"]
            rows_t6 = []
            national_acc_t6 = {s: 0 for s in source_names}
            for group_n, branches in groups.items():
                group_acc_t6 = {s: 0 for s in source_names}
                for b_name in branches:
                    grp = target_clean_pool[target_clean_pool['分校'] == b_name]
                    v_counts = grp['好友添加来源'].fillna('未知').apply(lambda x: x if x in source_names else '其他').value_counts().to_dict()
                    branch_counts = [v_counts.get(s, 0) for s in source_names]
                    branch_total = sum(branch_counts)
                    rows_t6.append([b_name] + branch_counts + [branch_total] + [b_name] + [t_pct_fn(s_div_fn(c, branch_total)) for c in branch_counts] + [branch_total])
                    for i, s_name in enumerate(source_names):
                        group_acc_t6[s_name] += branch_counts[i]
                        national_acc_t6[s_name] += branch_counts[i]
                
                group_cnts = [group_acc_t6[s] for s in source_names]
                group_total_sum = sum(group_cnts)
                rows_t6.append([f"{group_n}类总计"] + group_cnts + [group_total_sum] + [f"{group_n}类总计"] + [t_pct_fn(s_div_fn(c, group_total_sum)) for c in group_cnts] + [group_total_sum])
            
            nat_cnts = [national_acc_t6[s] for s in source_names]
            nat_total_sum = sum(nat_cnts)
            rows_t6.append(["全国总计"] + nat_cnts + [nat_total_sum] + ["全国总计"] + [t_pct_fn(s_div_fn(c, nat_total_sum)) for c in nat_cnts] + [nat_total_sum])
            pd.DataFrame(rows_t6, columns=["分校"]+source_names+["总计"]+["分校"]+source_names+["总计"]).to_excel(writer, sheet_name='6_好友添加来源', index=False)

        self._style_excel(output_file)
        self.log("🎯 质检汇总分析看板生成完毕！")

    # --- 通用渲染 ---
    def _write_wide_excel_with_headers(self, out_f, rows, c_map, s_idx, has_c):
        sub_h = ["新增", "去重", "净增", "重复率", "公职", "事业", "其他", "总量", "标备率", "回话", "回话率"]
        with pd.ExcelWriter(out_f, engine='openpyxl') as writer:
            pd.DataFrame(rows).to_excel(writer, index=False, header=False, startrow=2)
            ws = writer.sheets['Sheet1']; ws['A1'] = "分校"; ws.merge_cells('A1:A2')
            if has_c: ws['B1'] = "地市"; ws.merge_cells('B1:B2')
            curr = s_idx
            for ctn, _ in c_map:
                ws.cell(row=1, column=curr).value = ctn
                for s in sub_h: ws.cell(row=2, column=curr).value = s; curr += 1
                ws.merge_cells(start_row=1, start_column=curr-len(sub_h), end_row=1, end_column=curr-1)
            self._style_excel_ws(ws)

    def _write_wide_excel_retention_structured(self, out_f, rows, c_map, s_idx, has_c):
        sub_h = ["留存量", "公职", "事业", "其他", "总量", "标备率"]
        with pd.ExcelWriter(out_f, engine='openpyxl') as writer:
            pd.DataFrame(rows).to_excel(writer, index=False, header=False, startrow=2)
            ws = writer.sheets['Sheet1']; ws['A1'] = "分校"; ws.merge_cells('A1:A2')
            if has_c: ws['B1'] = "地市"; ws.merge_cells('B1:B2')
            curr = s_idx
            for ctn, _ in c_map:
                ws.cell(row=1, column=curr).value = ctn
                for s in sub_h: ws.cell(row=2, column=curr).value = s; curr += 1
                ws.merge_cells(start_row=1, start_column=curr-len(sub_h), end_row=1, end_column=curr-1)
            self._style_excel_ws(ws)

    def _style_excel(self, path):
        try:
            wb = load_workbook(path)
            for sn in wb.sheetnames: self._style_excel_ws(wb[sn])
            wb.save(path)
        except: pass

    def _style_excel_ws(self, ws):
        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        align = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows():
            for cell in row: cell.border = border; cell.alignment = align
        for row in ws.iter_rows(min_row=1, max_row=2):
            for cell in row: cell.font = Font(bold=True); cell.fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
        ws.column_dimensions['A'].width = 18
        if ws['B1'].value == '地市': ws.column_dimensions['B'].width = 15

# ==============================================================================
# GUI 界面
# ==============================================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("数据自动化统计工具 4.20")
        self.root.geometry("980x520")
        self.root.configure(bg="#f8f9fa")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TNotebook.Tab', font=("Microsoft YaHei", 10, "bold"), padding=[15, 5])
        tk.Label(root, text="🚀 数据自动化统计工具 4.20 (终极全量稳定版)", font=("Microsoft YaHei", 15, "bold"), bg="#007bff", fg="white").pack(fill="x", pady=0, ipady=10)
        
        main_frame = tk.Frame(root, bg="#f8f9fa")
        main_frame.pack(fill="both", expand=True, padx=15, pady=10)
        left_side = tk.Frame(main_frame, bg="#f8f9fa"); left_side.pack(side="left", fill="both", expand=True)
        right_side = tk.Frame(main_frame, bg="#f8f9fa", width=350); right_side.pack(side="right", fill="y", padx=(15, 0)); right_side.pack_propagate(False)

        self.notebook = ttk.Notebook(left_side); self.notebook.pack(fill="both", expand=True)
        self.tab_import = ttk.Frame(self.notebook); self.tab_regular = ttk.Frame(self.notebook); self.tab_qc = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_import, text=" 📂 数据导入 "); self.notebook.add(self.tab_regular, text=" 📈 常规报表 "); self.notebook.add(self.tab_qc, text=" 📊 质检分析 ")

        # 导入
        s1 = ttk.LabelFrame(self.tab_import, text=" 步骤1：选择源文件 "); s1.pack(fill="x", pady=20, padx=15, ipady=15)
        self.ent_f = ttk.Entry(s1); self.ent_f.pack(side="left", fill="x", expand=True, padx=10)
        ttk.Button(s1, text="📂 浏览文件", command=self.browse, width=12).pack(side="right", padx=10)

        # 常规
        s2 = ttk.LabelFrame(self.tab_regular, text=" 参数设置 "); s2.pack(fill="x", pady=10, padx=10)
        fp = ttk.Frame(s2); fp.pack(fill="x", padx=10, pady=5)
        ttk.Label(fp, text="📅 月份:").pack(side="left"); self.ent_m = ttk.Entry(fp, width=15); self.ent_m.insert(0, "2026-01"); self.ent_m.pack(side="left", padx=5)
        ttk.Label(fp, text="📆 特定日期:").pack(side="left"); self.ent_d = ttk.Entry(fp); self.ent_d.pack(side="left", fill="x", expand=True, padx=5)
        
        s3 = ttk.LabelFrame(self.tab_regular, text=" 任务选择 "); s3.pack(fill="x", pady=5, padx=10)
        self.v_ret = tk.BooleanVar(); self.v_pl = tk.BooleanVar(); self.v_pw = tk.BooleanVar(); self.v_cl = tk.BooleanVar(); self.v_cw = tk.BooleanVar(); self.v_spec = tk.BooleanVar()
        ttk.Checkbutton(s3, text="生成留存版报表", variable=self.v_ret).pack(anchor="w", padx=15)
        gf = ttk.Frame(s3); gf.pack(fill="x", padx=15, pady=2)
        ttk.Checkbutton(gf, text="省份一维", variable=self.v_pl).grid(row=0, column=0, sticky="w", pady=5); ttk.Checkbutton(gf, text="省份宽表", variable=self.v_pw).grid(row=0, column=1, sticky="w", pady=5)
        ttk.Checkbutton(gf, text="地市一维", variable=self.v_cl).grid(row=1, column=0, sticky="w", pady=5); ttk.Checkbutton(gf, text="地市宽表", variable=self.v_cw).grid(row=1, column=1, sticky="w", pady=5)
        ttk.Checkbutton(gf, text="★单独地市报表", variable=self.v_spec).grid(row=2, column=0, columnspan=2, sticky="w", pady=5)
        self.btn_reg = ttk.Button(self.tab_regular, text="📈 执行常规任务", command=lambda: self.run('regular')); self.btn_reg.pack(side="bottom", fill="x", pady=10, padx=15, ipady=8)

        # 质检
        s4 = ttk.LabelFrame(self.tab_qc, text=" 质检配置 "); s4.pack(fill="both", expand=True, pady=10, padx=10)
        f_q0 = ttk.Frame(s4); f_q0.pack(fill="x", padx=15, pady=6); ttk.Label(f_q0, text="📆 过滤日期:").pack(side="left"); self.ent_q_date = ttk.Entry(f_q0); self.ent_q_date.pack(side="left", fill="x", expand=True, padx=5)
        f_q1 = ttk.Frame(s4); f_q1.pack(fill="x", padx=15, pady=6); ttk.Label(f_q1, text="🚫 备注排除:").pack(side="left"); self.ent_q_rem = ttk.Entry(f_q1); self.ent_q_rem.insert(0, "学26,学27,报26,课26,前台,到店,报27,课27"); self.ent_q_rem.pack(side="left", fill="x", expand=True, padx=5)
        f_q2 = ttk.Frame(s4); f_q2.pack(fill="x", padx=15, pady=6); ttk.Label(f_q2, text="🚫 渠道排除:").pack(side="left"); self.ent_q_ch = ttk.Entry(f_q2); self.ent_q_ch.insert(0, "前台,到店"); self.ent_q_ch.pack(side="left", fill="x", expand=True, padx=5)
        self.btn_qc = ttk.Button(self.tab_qc, text="🎯 执行质检分析", command=lambda: self.run('qc')); self.btn_qc.pack(side="bottom", fill="x", pady=10, padx=15, ipady=8)

        self.log_txt = scrolledtext.ScrolledText(right_side, font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4", relief="flat"); self.log_txt.pack(fill="both", expand=True)

    def log(self, msg):
        self.log_txt.config(state='normal'); self.log_txt.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n"); self.log_txt.see(tk.END); self.log_txt.config(state='disabled'); self.root.update_idletasks()

    def browse(self):
        f_list = filedialog.askopenfilenames(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if f_list:
            self.ent_f.delete(0, tk.END); self.ent_f.insert(0, ";".join(f_list))

    def run(self, t_type):
        threading.Thread(target=self.process, args=(t_type,), daemon=True).start()

    def process(self, task_type):
        files_p = self.ent_f.get().split(";"); 
        if not files_p or files_p[0] == "": return messagebox.showerror("错误", "未选源文件！")
        start_ts = time.time(); self.btn_reg.config(state="disabled"); self.btn_qc.config(state="disabled")
        self.log_txt.config(state='normal'); self.log_txt.delete(1.0, tk.END); self.log_txt.config(state='disabled')
        try:
            proc = UniversalProcessor(self.log); dfs = []
            self.log(f"正在载入 {len(files_p)} 个文件...")
            for f in files_p:
                if f.strip().lower().endswith('.csv'): dfs.append(pd.read_csv(f.strip(), encoding='utf-8-sig'))
                else: dfs.append(pd.read_excel(f.strip()))
            full_df = pd.concat(dfs, ignore_index=True); 
            clean_df, dp_df, dc_df, ret_df = proc.process_step1(full_df, self.ent_m.get(), "")
            b_dir = os.path.dirname(files_p[0]); b_name = os.path.splitext(os.path.basename(files_p[0]))[0]

            if task_type == 'regular':
                dc_df.to_csv(os.path.join(b_dir, f"{b_name}_00_全量备份.csv"), index=False, encoding='utf-8-sig')
                m_ch = [("微信总量", None), ("线上", "线上平台"), ("线下", "线下活动"), ("考试", "考试现场"), ("高校", "高校"), ("其他", "其他")]
                m_on = [("微信总量", None), ("网站", "网站"), ("红书", "小红书"), ("公号", "公众号"), ("抖音", "抖音"), ("视频", "视频号"), ("其他", "其他")]
                do_r = self.v_ret.get() and not ret_df.empty
                if self.v_pl.get(): proc.gen_prov_long_merged(clean_df, dp_df, os.path.join(b_dir, f"{b_name}_01_省一维.xlsx"))
                if self.v_pw.get():
                    proc.gen_prov_wide(clean_df, dp_df, os.path.join(b_dir, f"{b_name}_02_省宽_渠道.xlsx"), m_ch, False); proc.gen_prov_wide(clean_df, dp_df, os.path.join(b_dir, f"{b_name}_03_省宽_线上.xlsx"), m_on, True)
                    if do_r: 
                        proc.gen_wide_retention_tables(ret_df, os.path.join(b_dir, f"{b_name}_02_省宽_渠道_留存.xlsx"), m_ch, False, True)
                        proc.gen_wide_retention_tables(ret_df, os.path.join(b_dir, f"{b_name}_03_省宽_线上_留存.xlsx"), m_on, True, True)
                if self.v_cl.get(): proc.gen_city_long_merged(clean_df, dc_df, ret_df, os.path.join(b_dir, f"{b_name}_04_地一维.xlsx"), do_r)
                if self.v_cw.get():
                    proc.gen_city_wide(clean_df, dc_df, os.path.join(b_dir, f"{b_name}_05_地宽_渠道.xlsx"), m_ch, False); proc.gen_city_wide(clean_df, dc_df, os.path.join(b_dir, f"{b_name}_06_地宽_线上.xlsx"), m_on, True)
                    if do_r:
                        proc.gen_wide_retention_tables(ret_df, os.path.join(b_dir, f"{b_name}_05_地宽_渠道_留存.xlsx"), m_ch, False, False)
                        proc.gen_wide_retention_tables(ret_df, os.path.join(b_dir, f"{b_name}_06_地宽_线上_留存.xlsx"), m_on, True, False)
                if self.v_spec.get():
                    proc.gen_special_city_report(clean_df, dc_df, os.path.join(b_dir, f"{b_name}_07_单独_渠道.xlsx"), m_ch, False); proc.gen_special_city_report(clean_df, dc_df, os.path.join(b_dir, f"{b_name}_08_单独_线上.xlsx"), m_on, True)
                    if do_r:
                        proc.gen_wide_retention_tables(ret_df, os.path.join(b_dir, f"{b_name}_07_单独_渠道_留存.xlsx"), m_ch, False, False, True)
                        proc.gen_wide_retention_tables(ret_df, os.path.join(b_dir, f"{b_name}_08_单独_线上_留存.xlsx"), m_on, True, False, True)
                if self.ent_d.get().strip(): proc.gen_date_summary_report(dc_df, self.ent_d.get(), os.path.join(b_dir, f"{b_name}_09_汇总.xlsx"), True)
            else:
                proc.gen_quality_inspection_suite(clean_df, ret_df, os.path.join(b_dir, f"{b_name}_10_看板.xlsx"), self.ent_q_rem.get(), self.ent_q_ch.get(), self.ent_q_date.get())
            self.log(f"🎉 全部处理完成！耗时: {time.time()-start_ts:.2f}s"); messagebox.showinfo("成功", "完毕")
        except Exception as e:
            self.log(f"💥 错误: {str(e)}"); messagebox.showerror("错误", str(e))
        finally:
            self.btn_reg.config(state="normal"); self.btn_qc.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk(); app = App(root); root.mainloop()
