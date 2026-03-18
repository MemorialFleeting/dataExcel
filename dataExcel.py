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

        df['渠道'] = df['渠道活码分组'].apply(lambda x: self.excel_lookup_find(x, ["网络","社会","高校","线上","现场","线下"], ["线上平台","线下活动","高校","线上平台","考试现场","线下活动"], "其他"))
        df['线上渠道'] = df['渠道活码分组'].apply(lambda x: self.excel_lookup_find(x, ["网站","公众号","小红书","视频号","抖音"], ["网站","公众号","小红书","视频号","抖音"], "其他"))

        df['事业2'] = df.apply(lambda row: 1 if row['公职'] not in [1, "1"] and row['事业辅助列'] in [1, "1"] else 0, axis=1)

        def check_reply(row):
            has_reply = pd.notna(row['客户上次回复时间']) and str(row['客户上次回复时间']).strip() != ""
            return "是" if has_reply and str(row['标备总数']) == "1" else "否"
        df['客户回话'] = df.apply(check_reply, axis=1)

        def fmt_date(val):
            try: return f"{pd.to_datetime(val).month}月{pd.to_datetime(val).day}日"
            except: return str(val)
        df['日期'] = df['好友添加时间'].apply(fmt_date)

        # 质检所需的 5 个原生字段
        qc_required_cols = ['好友添加来源', '添加渠道码', '员工发送消息数', '客户回复消息数', '是否同意会话存档']
        for m in qc_required_cols:
            if m not in df.columns:
                df[m] = 0 if '数' in m else "未知"

        col_z_name = df.columns[25] if len(df.columns) > 25 else "ExternalUserId"
        if 'ExternalUserId' not in df.columns:
            df.rename(columns={df.columns[1]: 'ExternalUserId'}, inplace=True)
            col_z_name = 'ExternalUserId'

        final_cols = ['备注名', 'ExternalUserId', '分校', '地市', '所属类别', '公职', '事业2', '教师', '文职', '医疗', '银行', '考研', '学历', '其他', '标备总数', '！！！是否新增', '渠道', '线上渠道', '客户回话', '日期', '添加好友状态']
        actual_cols = [c for c in final_cols if c in df.columns]
        for f in qc_required_cols:
            if f not in actual_cols: actual_cols.append(f)

        df_clean = df[actual_cols].copy()
        for col in ['公职', '事业2', '标备总数', '员工发送消息数', '客户回复消息数']:
            if col in df_clean.columns: df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

        df_dp = df_clean.drop_duplicates(subset=['ExternalUserId', '分校'], keep='first')
        df_dc = df_clean.drop_duplicates(subset=['ExternalUserId', '分校', '地市'], keep='first')
        df_ret = df_clean[df_clean['添加好友状态'].astype(str).str.contains("已添加", na=False)].copy()

        return df_clean, df_dp, df_dc, df_ret

    def calc_stats(self, df_raw, df_dedup):
        new_df = df_dedup[df_dedup['！！！是否新增'].astype(str).str.contains("是", na=False)]
        new_add = len(new_df)
        gz = new_df['公职'].sum(); sy = new_df['事业2'].sum(); tb = new_df['标备总数'].sum()
        rep = len(new_df[new_df['客户回话'].astype(str).str.contains("是", na=False)])
        return {"raw_cnt": len(df_raw), "dedup_cnt": len(df_dedup), "new_add": new_add, "gz": gz, "sy": sy, "tb": tb, "rep": rep}

    def fmt_stats(self, s, is_province_mode=False):
        def safe_div(a, b): return a / b if b != 0 else 0.0
        def to_pct(v): return f"{v:.2%}" if is_province_mode else f"{int(round(v*100))}%"
        dup_rate = safe_div((s["dedup_cnt"] - s["new_add"]), s["dedup_cnt"])
        other = s["tb"] - s["gz"] - s["sy"]
        tb_rate = safe_div(s["tb"], s["new_add"]); rep_rate = safe_div(s["rep"], s["new_add"])
        return [int(s["raw_cnt"]), int(s["dedup_cnt"]), int(s["new_add"]), to_pct(dup_rate), int(s["gz"]), int(s["sy"]), int(other), int(s["tb"]), to_pct(tb_rate), int(s["rep"]), to_pct(rep_rate)]

    def calc_stats_retention(self, df_subset):
        retention_count = len(df_subset)
        if retention_count == 0: return {"gz":0, "sy":0, "tb":0, "ret":0}
        gz = df_subset['公职'].sum(); sy = df_subset['事业2'].sum(); tb = df_subset['标备总数'].sum()
        return {"gz": gz, "sy": sy, "tb": tb, "ret": retention_count}

    def fmt_stats_retention(self, s, is_province_mode=False):
        def safe_div(a, b): return a / b if b != 0 else 0.0
        def to_pct(v): return f"{v:.2%}" if is_province_mode else f"{int(round(v*100))}%"
        other = s["tb"] - s["gz"] - s["sy"]; tb_rate = safe_div(s["tb"], s["ret"])
        return [int(s["ret"]), int(s["gz"]), int(s["sy"]), int(other), int(s["tb"]), to_pct(tb_rate)]

    # --- ★★★ 质检模块 4.7 核心逻辑 ★★★ ---
    def gen_quality_inspection_report(self, df_clean, df_ret, output_file, remark_exclude, channel_exclude, qc_dates_str):
        self.log(">>> 开始生成质检分析报表...")
        base_path = os.path.splitext(output_file)[0]
        
        # --- 表 1：详情明细底稿 (不受日期限制，全量留存) ---
        df_ret.to_csv(f"{base_path}_表1_详情明细底稿.csv", index=False, encoding='utf-8-sig')
        self.log(f"✅ 表 1 已完成，共导出 {len(df_ret)} 行。")

        # 准备受日期限制的数据池（用于表 2-6）
        df_ret_qc = df_ret.copy()
        df_clean_qc = df_clean.copy()
        if qc_dates_str.strip():
            target_dates = [d.strip() for d in qc_dates_str.replace('，', ',').split(',') if d.strip()]
            df_ret_qc = df_ret_qc[df_ret_qc['日期'].isin(target_dates)]
            df_clean_qc = df_clean_qc[df_clean_qc['日期'].isin(target_dates)]
            self.log(f"🔍 质检结果表已应用日期过滤: {target_dates}")

        # 准备“虚假备注统计池”
        df_abnormal = df_ret_qc.copy()
        df_abnormal = df_abnormal[(df_abnormal['标备总数'] == 1) & (df_abnormal['客户回复消息数'] == 0)]
        df_abnormal = df_abnormal[df_abnormal['好友添加来源'].astype(str).str.contains('扫描渠道二维码', na=False)]
        if remark_exclude.strip():
            rem_kws = [k.strip() for k in remark_exclude.replace('，', ',').split(',') if k.strip()]
            df_abnormal = df_abnormal[~df_abnormal['备注名'].astype(str).apply(lambda x: any(k in x for k in rem_kws))]
        if channel_exclude.strip():
            ch_kws = [k.strip() for k in channel_exclude.replace('，', ',').split(',') if k.strip()]
            df_abnormal = df_abnormal[~df_abnormal['添加渠道码'].astype(str).apply(lambda x: any(k in x for k in ch_kws))]

        groups = {"A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"], "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"], "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]}
        def safe_div(a, b): return a / b if b else 0.0
        def to_pct_int(v): return f"{int(round(v * 100))}%"

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # --- 表 2：虚假备注筛选 (核心条件：存档 == 同意) ---
            df_t2_pool = df_abnormal[df_abnormal['是否同意会话存档'].astype(str).str.strip() == '同意']
            rows_t2 = []; grand_t2 = 0
            for g_name, branches in groups.items():
                g_total = 0
                for branch in branches:
                    c = len(df_t2_pool[df_t2_pool['分校'] == branch])
                    g_total += c; rows_t2.append([branch, c])
                rows_t2.append([f"{g_name}类总计", g_total]); grand_t2 += g_total
            rows_t2.append(["全国总计", grand_t2])
            pd.DataFrame(rows_t2, columns=['分校', '虚假备注']).to_excel(writer, sheet_name='2_虚假备注筛选', index=False)
            
            # --- 表 3/4：跨分校 ---
            df3_base = df_ret_qc.drop_duplicates(subset=['ExternalUserId', '分校'])
            df3 = df3_base.groupby('ExternalUserId').size().reset_index(name='跨分校')
            df3.to_csv(f"{base_path}_表3_加几个分校明细底稿.csv", index=False, encoding='utf-8-sig')
            over10 = df3[df3['跨分校'] >= 10]['ExternalUserId']
            df4_pool = df3_base[df3_base['ExternalUserId'].isin(over10)]
            rows_t4 = []; grand_t4 = 0
            for g_name, branches in groups.items():
                g_total = 0
                for branch in branches:
                    c = len(df4_pool[df4_pool['分校'] == branch])
                    g_total += c; rows_t4.append([branch, c])
                rows_t4.append([f"{g_name}类总计", g_total]); grand_t4 += g_total
            rows_t4.append(["全国总计", grand_t4])
            pd.DataFrame(rows_t4, columns=['分校', '添加客服量']).to_excel(writer, sheet_name='4_跨10个以上分校', index=False)
            
            # --- 表 5：会话情况 ---
            rows_t5 = []; grand_t5 = {'ret':0, 'has':0, 'no':0, 'db':0}
            for g_name, branches in groups.items():
                gt = {'ret':0, 'has':0, 'no':0, 'db':0}
                for branch in branches:
                    grp = df_ret_qc[df_ret_qc['分校'] == branch]; ret = len(grp)
                    has = (grp['客户回复消息数'] > 0).sum(); no = (grp['客户回复消息数'] == 0).sum()
                    db = ((grp['员工发送消息数'] == 0) & (grp['客户回复消息数'] == 0)).sum()
                    gt['ret']+=ret; gt['has']+=has; gt['no']+=no; gt['db']+=db
                    grand_t5['ret']+=ret; grand_t5['has']+=has; grand_t5['no']+=no; grand_t5['db']+=db
                    rows_t5.append([branch, ret, has, no, to_pct_int(safe_div(no, ret)), db, to_pct_int(safe_div(db, ret))])
                rows_t5.append([f"{g_name}类总计", gt['ret'], gt['has'], gt['no'], to_pct_int(safe_div(gt['no'], gt['ret'])), gt['db'], to_pct_int(safe_div(gt['db'], gt['ret']))])
            rows_t5.append(["全国总计", grand_t5['ret'], grand_t5['has'], grand_t5['no'], to_pct_int(safe_div(grand_t5['no'], grand_t5['ret'])), grand_t5['db'], to_pct_int(safe_div(grand_t5['db'], grand_t5['ret']))])
            pd.DataFrame(rows_t5, columns=['分校','留存量','客户有会话','客户0会话','客户0会话率','双向无会话','双向无会话率']).to_excel(writer, sheet_name='5_会话情况', index=False)

            # --- 表 6：好友添加来源 ---
            srcs = ["获客助手", "名片分享", "其他", "群聊", "扫描名片二维码", "扫描渠道二维码", "搜索手机号", "微信联系人", "未知"]
            rows_t6 = []; g6 = {s: 0 for s in srcs}
            for g_name, branches in groups.items():
                t6 = {s: 0 for s in srcs}
                for branch in branches:
                    grp = df_clean_qc[df_clean_qc['分校'] == branch]
                    cv = grp['好友添加来源'].fillna('未知').apply(lambda x: x if x in srcs else '其他').value_counts().to_dict()
                    bc = [cv.get(s, 0) for s in srcs]; bt = sum(bc)
                    for i, s in enumerate(srcs): t6[s]+=bc[i]; g6[s]+=bc[i]
                    rows_t6.append([branch] + bc + [bt] + [branch] + [to_pct_int(safe_div(c, bt)) for c in bc] + [bt])
                tc = [t6[s] for s in srcs]; tt = sum(tc)
                rows_t6.append([f"{g_name}类总计"] + tc + [tt] + [f"{g_name}类总计"] + [to_pct_int(safe_div(c, tt)) for c in tc] + [tt])
            gc = [g6[s] for s in srcs]; gt = sum(gc)
            rows_t6.append(["全国总计"] + gc + [gt] + ["全国总计"] + [to_pct_int(safe_div(c, gt)) for c in gc] + [gt])
            pd.DataFrame(rows_t6, columns=["分校"]+srcs+["总计"]+["分校"]+srcs+["总计"]).to_excel(writer, sheet_name='6_好友添加来源', index=False)

        self._style_excel(output_file)
        self.log("🎯 质检分析综合报表已全部完成！")

    def gen_prov_long_merged(self, df_raw, df_dedup, output_file):
        self.log(f"生成 [省份-一维]: {os.path.basename(output_file)}")
        groups = {"A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"], "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"], "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]}
        sub_headers = ["新增好友", "本月分校内去重", "净新增", "重复率", "公职标备", "事业标备", "其他标备", "标备总量", "标备率", "回话备注", "回话备注率"]
        iter_list = [("线上平台", "网站", "网站"), ("线上平台", "小红书", "小红书"), ("线上平台", "公众号", "公众号"), ("线上平台", "抖音", "抖音"), ("线上平台", "视频号", "视频号"), ("线上平台", "其他", "其他"), ("线下活动", "线下活动", "线下活动"), ("考试现场", "考试现场", "考试现场"), ("高校", "高校", "高校"), ("其他", "其他", "其他")]
        final_rows = []; grand_acc = {key: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for key in [x[2] for x in iter_list]}
        for g_name, branches in groups.items():
            g_acc = {key: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for key in [x[2] for x in iter_list]}
            for branch in branches:
                br = df_raw[df_raw['分校'] == branch]; bd = df_dedup[df_dedup['分校'] == branch]
                for pn, fk, sn in iter_list:
                    cr = br[(br['渠道']=='线上平台')&(br['线上渠道']==fk)] if pn=="线上平台" else br[br['渠道']==fk]
                    cd = bd[(bd['渠道']=='线上平台')&(bd['线上渠道']==fk)] if pn=="线上平台" else bd[bd['渠道']==fk]
                    stats = self.calc_stats(cr, cd)
                    for k in stats: g_acc[sn][k] += stats[k]; grand_acc[sn][k] += stats[k]
                    final_rows.append([branch, pn, sn] + self.fmt_stats(stats, True))
            for pn, fk, sn in iter_list: final_rows.append([f"{g_name}类总计", pn, sn] + self.fmt_stats(g_acc[sn], True))
        for pn, fk, sn in iter_list: final_rows.append(["全国", pn, sn] + self.fmt_stats(grand_acc[sn], True))
        pd.DataFrame(final_rows, columns=["分校", "所属平台", "所属渠道"]+sub_headers).to_excel(output_file, index=False); self._style_excel(output_file)

    def gen_prov_wide(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成: {os.path.basename(output_file)}")
        groups = {"A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"], "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"], "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]}
        final_rows = []; grand_acc = {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map}
        for g_n, branches in groups.items():
            g_acc = {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map}
            for branch in branches:
                br = df_raw[df_raw['分校'] == branch]; bd = df_dedup[df_dedup['分校'] == branch]; row = [branch]
                for ct, cf in channel_map:
                    cr = (br[(br['渠道']=='线上平台')&(br['线上渠道']==cf)] if strict_online else br[br['渠道']==cf]) if cf else br
                    cd = (bd[(bd['渠道']=='线上平台')&(bd['线上渠道']==cf)] if strict_online else bd[bd['渠道']==cf]) if cf else bd
                    stats = self.calc_stats(cr, cd)
                    for k in stats: grand_acc[ct][k] += stats[k]; g_acc[ct][k] += stats[k]
                    row.extend(self.fmt_stats(stats, True))
                final_rows.append(row)
            gr = [f"{g_n}类总计"]
            for ct, _ in channel_map: gr.extend(self.fmt_stats(g_acc[ct], True))
            final_rows.append(gr)
        n_row = ["全国"]
        for ct, _ in channel_map: n_row.extend(self.fmt_stats(grand_acc[ct], True))
        final_rows.append(n_row); self._write_wide_excel(output_file, final_rows, channel_map, 2, False)

    def gen_city_long_merged(self, df_raw, df_dedup, df_ret, output_file, do_ret):
        self.log(f"生成 [地市-一维]: {os.path.basename(output_file)}")
        base_h = ["新增好友", "本月分校内去重", "净新增", "重复率", "净-公职标备", "净-事业标备", "净-其他标备", "净-标备总量", "净-标备率", "回话备注", "回话备注率"]
        ext_h = ["总-留存量", "总-公职标备", "总-事业标备", "总-其他标备", "总-标备总量"] if do_ret else []
        iter_list = [("线上平台", "网站", "网站"), ("线上平台", "小红书", "小红书"), ("线上平台", "公众号", "公众号"), ("线上平台", "抖音", "抖音"), ("线上平台", "视频号", "视频号"), ("线上平台", "其他", "其他"), ("线下活动", "线下活动", "线下活动"), ("考试现场", "考试现场", "考试现场"), ("高校", "高校", "高校"), ("其他", "其他", "其他")]
        final_rows = []; grand_acc = {key: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for key in [x[2] for x in iter_list]}
        gr_acc = {key: {"gz":0,"sy":0,"tb":0,"ret":0} for key in [x[2] for x in iter_list]}
        for branch in sorted(df_raw['分校'].dropna().unique()):
            br = df_raw[df_raw['分校'] == branch]; bd = df_dedup[df_dedup['分校'] == branch]; ret = df_ret[df_ret['分校'] == branch] if do_ret else None
            for city in sorted(br['地市'].dropna().unique()):
                crb = br[br['地市'] == city]; cdb = bd[bd['地市'] == city]; cat = crb['所属类别'].iloc[0] if not crb.empty else "其它"
                for pn, fk, sn in iter_list:
                    cr = crb[(crb['渠道']=='线上平台')&(crb['线上渠道']==fk)] if pn=="线上平台" else crb[crb['渠道']==fk]
                    cd = cdb[(cdb['渠道']=='线上平台')&(cdb['线上渠道']==fk)] if pn=="线上平台" else cdb[cdb['渠道']==fk]
                    stats = self.calc_stats(cr, cd)
                    for k in stats: grand_acc[sn][k] += stats[k]
                    ex_c = []
                    if do_ret:
                        cr_b = ret[ret['地市'] == city]
                        crr = cr_b[(cr_b['渠道']=='线上平台')&(cr_b['线上渠道']==fk)] if pn=="线上平台" else cr_b[cr_b['渠道']==fk]
                        rs = self.calc_stats_retention(crr)
                        for k in rs: gr_acc[sn][k] += rs[k]
                        ex_c = [int(rs["ret"]), int(rs["gz"]), int(rs["sy"]), int(rs["tb"]-rs["gz"]-rs["sy"]), int(rs["tb"])]
                    final_rows.append([branch, city, cat, pn, sn] + self.fmt_stats(stats, False) + ex_c)
        for pn, fk, sn in iter_list:
            row = ["全国", "总计", "-", pn, sn] + self.fmt_stats(grand_acc[sn], False)
            if do_ret:
                ra = gr_acc[sn]; row += [int(ra["ret"]), int(ra["gz"]), int(ra["sy"]), int(ra["tb"]-ra["gz"]-ra["sy"]), int(ra["tb"])]
            final_rows.append(row)
        pd.DataFrame(final_rows, columns=["分校", "地市", "所属类别", "所属平台", "所属渠道"] + base_h + ext_h).to_excel(output_file, index=False); self._style_excel(output_file)

    def gen_city_wide(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成: {os.path.basename(output_file)}")
        final_rows = []; grand_acc = {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map}
        for branch in sorted(df_raw['分校'].dropna().unique()):
            br = df_raw[df_raw['分校'] == branch]; bd = df_dedup[df_dedup['分校'] == branch]
            for city in sorted(br['地市'].dropna().unique()):
                row = [branch, city]; crb = br[br['地市'] == city]; cdb = bd[bd['地市'] == city]
                for ct, cf in channel_map:
                    cr = (crb[(crb['渠道']=='线上平台')&(crb['线上渠道']==cf)] if strict_online else crb[crb['渠道']==cf]) if cf else crb
                    cd = (cdb[(cdb['渠道']=='线上平台')&(cdb['线上渠道']==cf)] if strict_online else cdb[cdb['渠道']==cf]) if cf else cdb
                    stats = self.calc_stats(cr, cd)
                    for k in stats: grand_acc[ct][k] += stats[k]
                    row.extend(self.fmt_stats(stats, False))
                final_rows.append(row)
        n_row = ["全国", "总计"]
        for ct, _ in channel_map: n_row.extend(self.fmt_stats(grand_acc[ct], False))
        final_rows.append(n_row); self._write_wide_excel(output_file, final_rows, channel_map, 3, True)

    def gen_special_city_report(self, df_raw, df_dedup, output_file, channel_map, strict_online):
        self.log(f"生成 [单独地市]: {os.path.basename(output_file)}")
        final_rows = []; grand_acc = {ch[0]: {"raw_cnt":0, "dedup_cnt":0,"new_add":0,"gz":0,"sy":0,"tb":0,"rep":0} for ch in channel_map}
        for city in FIXED_ORDER_CITIES:
            row_raw = df_raw[df_raw['地市'] == city]; row_dedup = df_dedup[df_dedup['地市'] == city]
            bn = row_raw['分校'].iloc[0] if not row_raw.empty else DEFAULT_BRANCH_MAP.get(city, "未知分校")
            row_data = [bn, city]
            for ct, cf in channel_map:
                cr = (row_raw[(row_raw['渠道']=='线上平台')&(row_raw['线上渠道']==cf)] if strict_online else row_raw[row_raw['渠道']==cf]) if cf else row_raw
                cd = (row_dedup[(row_dedup['渠道']=='线上平台')&(row_dedup['线上渠道']==cf)] if strict_online else row_dedup[row_dedup['渠道']==cf]) if cf else row_dedup
                stats = self.calc_stats(cr, cd)
                for k in stats: grand_acc[ct][k] += stats[k]
                row_data.extend(self.fmt_stats(stats, False))
            final_rows.append(row_data)
        n_row = ["全国", "总计"]
        for ct, _ in channel_map: n_row.extend(self.fmt_stats(grand_acc[ct], False))
        final_rows.append(n_row); self._write_wide_excel(output_file, final_rows, channel_map, 3, True)

    def gen_wide_retention(self, df_ret, output_file, channel_map, strict_online, is_prov, is_spec):
        self.log(f"生成 [留存版]: {os.path.basename(output_file)}")
        groups = {"A": ["山东分校", "广东分校", "河南分校", "河北分校", "湖北分校", "吉林分校", "山西分校", "陕西分校", "安徽分校", "辽宁分校", "云南分校"], "B": ["江苏分校", "湖南分校", "贵州分校", "四川分校", "黑龙江分校", "广西分校", "新疆分校", "浙江分校", "江西分校", "福建分校", "北京分校"], "C": ["甘肃分校", "海南分校", "内蒙古分校", "宁夏分校", "青海分校", "厦门分校", "上海分校", "天津分校", "西藏分校", "重庆分校"]}
        final_rows = []; grand_acc = {ch[0]: {"gz":0,"sy":0,"tb":0,"ret":0} for ch in channel_map}
        if is_spec:
            for city in FIXED_ORDER_CITIES:
                rr = df_ret[df_ret['地市'] == city]; bn = rr['分校'].iloc[0] if not rr.empty else DEFAULT_BRANCH_MAP.get(city, "未知分校")
                row = [bn, city]
                for ct, cf in channel_map:
                    c_ret = (rr[(rr['渠道']=='线上平台')&(rr['线上渠道']==cf)] if strict_online else rr[rr['渠道']==cf]) if cf else rr
                    stats = self.calc_stats_retention(c_ret)
                    for k in stats: grand_acc[ct][k] += stats[k]
                    row.extend(self.fmt_stats_retention(stats, False))
                final_rows.append(row)
        elif is_prov:
            for g_n, branches in groups.items():
                g_acc = {ch[0]: {"gz":0,"sy":0,"tb":0,"ret":0} for ch in channel_map}
                for branch in branches:
                    br = df_ret[df_ret['分校'] == branch]; row = [branch]
                    for ct, cf in channel_map:
                        c_ret = (br[(br['渠道']=='线上平台')&(br['线上渠道']==cf)] if strict_online else br[br['渠道']==cf]) if cf else br
                        stats = self.calc_stats_retention(c_ret)
                        for k in stats: grand_acc[ct][k] += stats[k]; g_acc[ct][k] += stats[k]
                        row.extend(self.fmt_stats_retention(stats, True))
                    final_rows.append(row)
                gr = [f"{g_n}类总计"]
                for ct, _ in channel_map: gr.extend(self.fmt_stats_retention(g_acc[ct], True))
                final_rows.append(gr)
        else:
            for branch in sorted(df_ret['分校'].dropna().unique()):
                br = df_ret[df_ret['分校'] == branch]
                for city in sorted(br['地市'].dropna().unique()):
                    cr = br[br['地市'] == city]; row = [branch, city]
                    for ct, cf in channel_map:
                        c_ret = (cr[(cr['渠道']=='线上平台')&(cr['线上渠道']==cf)] if strict_online else cr[cr['渠道']==cf]) if cf else cr
                        stats = self.calc_stats_retention(c_ret)
                        for k in stats: grand_acc[ct][k] += stats[k]
                        row.extend(self.fmt_stats_retention(stats, False))
                    final_rows.append(row)
        n_row = (["全国", "总计"] if not is_prov else ["全国"])
        for ct, _ in channel_map: n_row.extend(self.fmt_stats_retention(grand_acc[ct], is_prov))
        final_rows.append(n_row); self._write_wide_excel_retention(output_file, final_rows, channel_map, 2 if is_prov else 3, not is_prov)

    def gen_date_summary(self, df_dedup, d_str, output_file, is_city):
        if not d_str.strip(): return
        self.log(f"生成日期汇总: {os.path.basename(output_file)}")
        dates = [d.strip() for d in d_str.replace('，', ',').split(',') if d.strip()]
        f = df_dedup[df_dedup['日期'].isin(dates)]; gc = ['分校', '地市'] if is_city else ['分校']
        c_all = f.groupby(gc).size().reset_index(name='total'); c_web = f[f['线上渠道'] == '网站'].groupby(gc).size().reset_index(name='web')
        res = pd.merge(c_all, c_web, on=gc, how='left').fillna(0)
        res.loc['Sum'] = pd.Series(res[['total', 'web']].sum(), index=['total', 'web']); res.at['Sum', '分校'] = '全国'
        if is_city: res.at['Sum', '地市'] = '总计'
        res.rename(columns={'total': f"汇总({','.join(dates)})", 'web': "其中:网站"}, inplace=True)
        res.to_excel(output_file, index=False); self._style_excel(output_file)

    def _write_wide_excel(self, output_file, rows, channel_map, start_col_idx, has_city_col):
        sub_headers = ["新增好友", "本月分校内去重", "净新增", "重复率", "公职标备", "事业标备", "其他标备", "标备总量", "标备率", "回话备注", "回话备注率"]
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            pd.DataFrame(rows).to_excel(writer, index=False, header=False, startrow=2)
            ws = writer.sheets['Sheet1']; ws['A1'] = "分校"; ws.merge_cells('A1:A2')
            if has_city_col: ws['B1'] = "地市"; ws.merge_cells('B1:B2')
            curr = start_col_idx
            for ct, _ in channel_map:
                ws.cell(row=1, column=curr).value = ct
                for sub in sub_headers: ws.cell(row=2, column=curr).value = sub; curr += 1
                ws.merge_cells(start_row=1, start_column=curr-len(sub_headers), end_row=1, end_column=curr-1)
            self._style_excel_ws(ws)

    def _write_wide_excel_retention(self, output_file, rows, channel_map, start_col_idx, has_city_col):
        sub_headers = ["留存量", "公职标备", "事业标备", "其他标备", "标备总量", "标备率"]
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            pd.DataFrame(rows).to_excel(writer, index=False, header=False, startrow=2)
            ws = writer.sheets['Sheet1']; ws['A1'] = "分校"; ws.merge_cells('A1:A2')
            if has_city_col: ws['B1'] = "地市"; ws.merge_cells('B1:B2')
            curr = start_col_idx
            for ct, _ in channel_map:
                ws.cell(row=1, column=curr).value = ct
                for sub in sub_headers: ws.cell(row=2, column=curr).value = sub; curr += 1
                ws.merge_cells(start_row=1, start_column=curr-len(sub_headers), end_row=1, end_column=curr-1)
            self._style_excel_ws(ws)

    def _style_excel(self, file_path):
        try:
            wb = load_workbook(file_path)
            for sn in wb.sheetnames: self._style_excel_ws(wb[sn])
            wb.save(file_path)
        except: pass

    def _style_excel_ws(self, ws):
        thin = Side(border_style="thin", color="000000"); border = Border(top=thin, left=thin, right=thin, bottom=thin)
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
        self.root.title("数据自动化统计工具 4.7")
        self.root.geometry("980x520")
        self.root.configure(bg="#f8f9fa")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background="#f8f9fa")
        style.configure('TLabelframe', background="#f8f9fa")
        style.configure('TLabelframe.Label', font=("Microsoft YaHei", 10, "bold"), background="#f8f9fa")
        style.configure('TButton', font=("Microsoft YaHei", 10), background="#007bff", foreground="white")
        style.configure('TNotebook.Tab', font=("Microsoft YaHei", 10, "bold"), padding=[15, 5])

        tk.Label(root, text="🚀 数据自动化统计工具 4.7", font=("Microsoft YaHei", 15, "bold"), bg="#007bff", fg="white").pack(fill="x", ipady=10)

        main_frame = tk.Frame(root, bg="#f8f9fa")
        main_frame.pack(fill="both", expand=True, padx=15, pady=10)

        left_frame = tk.Frame(main_frame, bg="#f8f9fa"); left_frame.pack(side="left", fill="both", expand=True)
        right_frame = tk.Frame(main_frame, bg="#f8f9fa", width=350); right_frame.pack(side="right", fill="y", padx=(15, 0)); right_frame.pack_propagate(False)

        self.notebook = ttk.Notebook(left_frame); self.notebook.pack(fill="both", expand=True)
        tab_import = ttk.Frame(self.notebook); tab_regular = ttk.Frame(self.notebook); tab_qc = ttk.Frame(self.notebook)
        self.notebook.add(tab_import, text=" 📂 数据源导入 "); self.notebook.add(tab_regular, text=" 📈 常规报表 "); self.notebook.add(tab_qc, text=" 📊 质检数据分析 ")

        # 标签 1：导入
        s1f = ttk.LabelFrame(tab_import, text=" 步骤 1：导入数据 "); s1f.pack(fill="x", pady=20, padx=15, ipady=15)
        f_file = ttk.Frame(s1f); f_file.pack(fill="x", padx=15)
        self.ent_f = ttk.Entry(f_file); self.ent_f.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(f_file, text="📂 浏览文件", command=self.browse, width=12).pack(side="right")

        # 标签 2：常规
        s2f = ttk.LabelFrame(tab_regular, text=" 步骤 2：参数设置 "); s2f.pack(fill="x", pady=10, padx=10)
        f_p1 = ttk.Frame(s2f); f_p1.pack(fill="x", padx=10, pady=5)
        ttk.Label(f_p1, text="📅 判断月份:").pack(side="left"); self.ent_m = ttk.Entry(f_p1, width=15); self.ent_m.insert(0, "2026-01"); self.ent_m.pack(side="left", padx=10)
        ttk.Label(f_p1, text="📆 特定日期:").pack(side="left"); self.ent_d = ttk.Entry(f_p1); self.ent_d.pack(side="left", fill="x", expand=True, padx=5)
        s3f = ttk.LabelFrame(tab_regular, text=" 步骤 3：常规任务 "); s3f.pack(fill="x", pady=5, padx=10)
        self.v_retention = tk.BooleanVar(); self.v_pl = tk.BooleanVar(); self.v_pw = tk.BooleanVar(); self.v_cl = tk.BooleanVar(); self.v_cw = tk.BooleanVar(); self.v_special = tk.BooleanVar()
        ttk.Checkbutton(s3f, text="🔥 同时生成“留存版”", variable=self.v_retention).pack(anchor="w", padx=15)
        grid_f = ttk.Frame(s3f); grid_f.pack(fill="x", padx=15, pady=2)
        ttk.Checkbutton(grid_f, text="省份一维", variable=self.v_pl).grid(row=0, column=0, sticky="w", pady=5); ttk.Checkbutton(grid_f, text="省份宽表", variable=self.v_pw).grid(row=0, column=1, sticky="w", pady=5)
        ttk.Checkbutton(grid_f, text="地市一维", variable=self.v_cl).grid(row=1, column=0, sticky="w", pady=5); ttk.Checkbutton(grid_f, text="地市宽表", variable=self.v_cw).grid(row=1, column=1, sticky="w", pady=5)
        ttk.Checkbutton(grid_f, text="★ 单独地市", variable=self.v_special).grid(row=2, column=0, columnspan=2, sticky="w", pady=5)
        self.btn_regular = ttk.Button(tab_regular, text="📈 仅生成常规报表", command=lambda: self.run('regular')); self.btn_regular.pack(side="bottom", fill="x", pady=10, padx=15, ipady=8)

        # 标签 3：质检
        s4f = ttk.LabelFrame(tab_qc, text=" 质检专属配置 (表1全量明细，表2按'同意'过滤) "); s4f.pack(fill="both", expand=True, pady=10, padx=10)
        f_q0 = ttk.Frame(s4f); f_q0.pack(fill="x", padx=15, pady=8); ttk.Label(f_q0, text="📆 独立过滤日期:").pack(side="left"); self.ent_q_date = ttk.Entry(f_q0); self.ent_q_date.pack(side="left", fill="x", expand=True, padx=5)
        f_q1 = ttk.Frame(s4f); f_q1.pack(fill="x", padx=15, pady=8); ttk.Label(f_q1, text="🚫 备注排除:").pack(side="left"); self.ent_q_rem = ttk.Entry(f_q1); self.ent_q_rem.insert(0, "学26,学27,报26,课26,前台"); self.ent_q_rem.pack(side="left", fill="x", expand=True, padx=5)
        f_q2 = ttk.Frame(s4f); f_q2.pack(fill="x", padx=15, pady=8); ttk.Label(f_q2, text="🚫 渠道排除:").pack(side="left"); self.ent_q_ch = ttk.Entry(f_q2); self.ent_q_ch.insert(0, "前台"); self.ent_q_ch.pack(side="left", fill="x", expand=True, padx=5)
        self.btn_qc = ttk.Button(tab_qc, text="🎯 仅生成质检报表", command=lambda: self.run('qc')); self.btn_qc.pack(side="bottom", fill="x", pady=10, padx=15, ipady=8)

        # 日志
        l_lbl_f = ttk.LabelFrame(right_frame, text=" 实时运行日志 "); l_lbl_f.pack(fill="both", expand=True)
        self.log_txt = scrolledtext.ScrolledText(l_lbl_f, font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4", relief="flat"); self.log_txt.pack(fill="both", expand=True, padx=2, pady=2)

    def log(self, msg):
        self.log_txt.config(state='normal'); self.log_txt.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n"); self.log_txt.see(tk.END); self.log_txt.config(state='disabled')
        self.root.update_idletasks()

    def browse(self):
        fs = filedialog.askopenfilenames(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")]); 
        if fs: self.ent_f.delete(0, tk.END); self.ent_f.insert(0, ";".join(fs))

    def run(self, task_type): threading.Thread(target=self.process, args=(task_type,), daemon=True).start()

    def process(self, task_type):
        files = self.ent_f.get().split(";"); 
        if not files or files[0] == "": return messagebox.showerror("错误", "请先选择数据文件！")
        start_t = time.time(); self.btn_regular.config(state="disabled"); self.btn_qc.config(state="disabled")
        self.log_txt.config(state='normal'); self.log_txt.delete(1.0, tk.END); self.log_txt.config(state='disabled')
        try:
            proc = UniversalProcessor(self.log); dfs = []
            self.log(f"任务启动，正在读取 {len(files)} 个文件...")
            for f in files:
                if f.strip().lower().endswith('.csv'): dfs.append(pd.read_csv(f.strip(), encoding='utf-8-sig'))
                else: dfs.append(pd.read_excel(f.strip()))
            full_df = pd.concat(dfs, ignore_index=True)
            self.log(f"读取完成，共计 {len(full_df)} 行数据。")
            
            df_clean, df_dp, df_dc, df_ret = proc.process_step1(full_df, self.ent_m.get(), "")
            base_dir = os.path.dirname(files[0]); base_name = os.path.splitext(os.path.basename(files[0]))[0]

            if task_type == 'regular':
                df_dc.to_csv(os.path.join(base_dir, f"{base_name}_00_源数据备份.csv"), index=False, encoding='utf-8-sig')
                m_ch = [("微信好友总量", None), ("线上平台", "线上平台"), ("线下平台", "线下活动"), ("考试现场", "考试现场"), ("高校", "高校"), ("其他", "其他")]
                m_on = [("微信好友总量", None), ("网站", "网站"), ("小红书", "小红书"), ("公众号", "公众号"), ("抖音", "抖音"), ("视频号", "视频号"), ("其他", "其他")]
                do_ret = self.v_retention.get() and not df_ret.empty
                if self.v_pl.get(): proc.gen_prov_long_merged(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_01_省份一维.xlsx"))
                if self.v_pw.get():
                    proc.gen_prov_wide(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_02_省份宽表_渠道.xlsx"), m_ch, False)
                    proc.gen_prov_wide(df_clean, df_dp, os.path.join(base_dir, f"{base_name}_03_省份宽表_线上.xlsx"), m_on, True)
                    if do_ret:
                        proc.gen_wide_retention(df_ret, os.path.join(base_dir, f"{base_name}_02_省份宽表_渠道_留存.xlsx"), m_ch, False, True, False)
                        proc.gen_wide_retention(df_ret, os.path.join(base_dir, f"{base_name}_03_省份宽表_线上_留存.xlsx"), m_on, True, True, False)
                if self.v_cl.get(): proc.gen_city_long_merged(df_clean, df_dc, df_ret, os.path.join(base_dir, f"{base_name}_04_地市一维.xlsx"), do_ret)
                if self.v_cw.get():
                    proc.gen_city_wide(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_05_地市宽表_渠道.xlsx"), m_ch, False)
                    proc.gen_city_wide(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_06_地市宽表_线上.xlsx"), m_on, True)
                if self.v_special.get(): proc.gen_special_city_report(df_clean, df_dc, os.path.join(base_dir, f"{base_name}_07_单独地市.xlsx"), m_ch, False)
                if self.ent_d.get().strip(): proc.gen_date_summary(df_dc, self.ent_d.get(), os.path.join(base_dir, f"{base_name}_09_日期汇总.xlsx"), True)
            else:
                if df_ret.empty: self.log("❌ 留存数据为空。")
                else: proc.gen_quality_inspection_report(df_clean, df_ret, os.path.join(base_dir, f"{base_name}_10_质检汇总看板.xlsx"), self.ent_q_rem.get(), self.ent_q_ch.get(), self.ent_q_date.get())
            
            self.log(f"🎉 全部完成！耗时: {time.time()-start_t:.2f}s"); messagebox.showinfo("成功", "任务处理完毕")
        except Exception as e:
            self.log(f"💥 错误: {str(e)}"); messagebox.showerror("错误", str(e))
        finally:
            self.btn_regular.config(state="normal"); self.btn_qc.config(state="normal")

if __name__ == "__main__": root = tk.Tk(); app = App(root); root.mainloop()
