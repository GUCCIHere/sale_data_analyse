# pip install pandas openpyxl flask numpy
import pandas as pd
import numpy as np
import os
from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

# ===================== 基础配置 =====================
SOURCE_EXCEL = r"C:\Users\zhangjie33\Desktop\VScode\sales_data_N.xlsx"
FILTER_RULES = {
    "事业部描述": "集团食品工业事业部",
    "MKT": ["外部", "工厂自销"]
}
# 指标区分数量/金额
CORE_METRICS = [
    {"code": "sales_qty", "name": "销售数量(MT)", "label": "销量(MT)", "type": "qty"},
    {"code": "sales_revenue", "name": "销售收入", "label": "销售收入", "type": "amount"},
    {"code": "gross_profit", "name": "销售毛利(扣除运费)", "label": "销售毛利", "type": "amount"},
    {"code": "factory_cost", "name": "工厂费用", "label": "工厂费用", "type": "amount"},
    {"code": "marketing_cost", "name": "营销公司费用", "label": "营销公司费用", "type": "amount"},
    {"code": "pretax_profit", "name": "税前利润-损益（含套保）", "label": "税前利润", "type": "amount"}
]
TON_METRICS = [
    {"label": "吨均收入"},
    {"label": "吨均毛利"},
    {"label": "吨均工厂费用"},
    {"label": "吨均营销费用"},
    {"label": "吨均税前利润"}
]
MOM_DIMENSIONS = [
    {"code": "current", "label": "本月"},
    {"code": "last", "label": "上月"},
    {"code": "change", "label": "增减额"},
    {"code": "rate", "label": "变动率(%)"}
]
YOY_DIMENSIONS = [
    {"code": "current", "label": "本年累计"},
    {"code": "last", "label": "去年同期"},
    {"code": "change", "label": "增减额"},
    {"code": "rate", "label": "变动率(%)"}
]

# 全局缓存
df_cache = None

# ===================== 工具函数 =====================
def fmt_num(x, is_pct=False):
    """仅做小数保留，万元转换放到前端统一处理，后端保存原始数值"""
    if pd.isna(x) or np.isinf(x):
        val = 0.00
    else:
        val = float(x)
    if is_pct:
        return round(val * 100, 2)
    return round(val, 2)

def load_all_data():
    global df_cache
    if df_cache is not None:
        return df_cache.copy()
    fp = SOURCE_EXCEL
    if not os.path.exists(fp):
        raise Exception("源Excel文件不存在")
    wb = load_workbook(fp, read_only=True, data_only=True, keep_links=False)
    ws = wb.active
    header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    headers = list(header)
    col_idx = {n.strip():i for i,n in enumerate(headers) if n}
    required = [
        "事业部描述","MKT","销售组织","会计年度","期间","2017分类","所属集团重分类",
        "销售数量(MT)","销售收入","销售毛利(扣除运费)",
        "工厂税金及附加","工厂销售费用","研发费用","剩余工厂管理费用","工厂财务费用",
        "营销销售费用","营销管理费用","营销财务费用","营销附加税金","税前利润-损益（含套保）"
    ]
    miss = [f for f in required if f not in col_idx]
    if miss:
        wb.close()
        raise Exception(f"Excel缺失字段：{','.join(miss)}")
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        div = str(r[col_idx["事业部描述"]]).strip() if r[col_idx["事业部描述"]] else ""
        mkt = str(r[col_idx["MKT"]]).strip() if r[col_idx["MKT"]] else ""
        if div == FILTER_RULES["事业部描述"] and mkt in FILTER_RULES["MKT"]:
            rows.append(r)
    wb.close()
    df = pd.DataFrame(rows, columns=headers)
    num_cols = [
        "销售数量(MT)","销售收入","销售毛利(扣除运费)",
        "工厂税金及附加","工厂销售费用","研发费用","剩余工厂管理费用","工厂财务费用",
        "营销销售费用","营销管理费用","营销财务费用","营销附加税金","税前利润-损益（含套保）"
    ]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["工厂费用"] = df["工厂税金及附加"] + df["工厂销售费用"] + df["研发费用"] + df["剩余工厂管理费用"] + df["工厂财务费用"]
    df["营销公司费用"] = df["营销销售费用"] + df["营销管理费用"] + df["营销财务费用"] + df["营销附加税金"]
    df["2017分类"] = df["2017分类"].fillna("未分类产品")
    df["所属集团重分类"] = df["所属集团重分类"].fillna("未分类客户")
    df["销售组织"] = df["销售组织"].astype(str).str.strip()
    df_cache = df
    print(f"缓存加载完成，总行数{len(df)}")
    return df.copy()

# ===================== 接口 =====================
@app.route("/get_branch_list")
def get_branch():
    try:
        df = load_all_data()
        branches = sorted(df["销售组织"].unique().tolist())
        return jsonify({"branches": branches})
    except Exception as e:
        return jsonify({"error": str(e), "branches": []}),500

@app.route("/query_data", methods=["POST"])
def query_data():
    try:
        req = request.get_json()
        if not req:
            return jsonify({"error":"无参数"}),400
        target_branch = str(req.get("branch")).strip()
        target_year = int(req.get("year"))
        target_month = int(req.get("month"))
        last_year = target_year - 1
        df = load_all_data()
        df_branch = df[df["销售组织"] == target_branch].copy()
        if len(df_branch) == 0:
            return jsonify({"error":"该分公司无数据"}),400
        metric_names = [m["name"] for m in CORE_METRICS]
        amount_names = [m["name"] for m in CORE_METRICS if m["type"]=="amount"]

        # 单月/上月/累计区间
        df_curr_month = df_branch[(df_branch["会计年度"]==target_year)&(df_branch["期间"]==target_month)]
        if target_month == 1:
            df_last_month = df_branch[(df_branch["会计年度"]==target_year-1)&(df_branch["期间"]==12)]
        else:
            df_last_month = df_branch[(df_branch["会计年度"]==target_year)&(df_branch["期间"]==target_month-1)]
        df_curr_cum = df_branch[(df_branch["会计年度"]==target_year)&(df_branch["期间"]<=target_month)]
        df_last_cum = df_branch[(df_branch["会计年度"]==last_year)&(df_branch["期间"]<=target_month)]

        # 聚合函数：原始数值求和，万元转换交给前端渲染
        def agg_cat(sub_df):
            if len(sub_df)==0:
                return pd.DataFrame(columns=["2017分类"]+metric_names)
            agg = sub_df.groupby("2017分类")[metric_names].sum().reset_index()
            agg = agg.sort_values("销售数量(MT)", ascending=False).reset_index(drop=True)
            total = {"2017分类":"合计"}
            for m in metric_names:
                s = float(agg[m].sum())
                total[m] = fmt_num(s)
            agg = pd.concat([agg, pd.DataFrame([total])], ignore_index=True)
            return agg

        agg_curr_month = agg_cat(df_curr_month)
        agg_last_month = agg_cat(df_last_month)
        agg_curr_cum = agg_cat(df_curr_cum)
        agg_last_cum = agg_cat(df_last_cum)

        # 环比表格构造：后端返回原始数值，前端统一处理万元
        def build_mom(curr_df, last_df):
            data = []
            cats = curr_df["2017分类"].tolist()
            for cat in cats:
                cr = curr_df[curr_df["2017分类"]==cat].iloc[0]
                lr = last_df[last_df["2017分类"]==cat].iloc[0] if cat in last_df["2017分类"].values else pd.Series({m:0 for m in metric_names})
                row = {"分类":cat}
                for m in CORE_METRICS:
                    mn = m["name"]
                    cv = float(cr[mn])
                    lv = float(lr[mn])
                    diff = cv - lv
                    rate = diff/lv if lv !=0 else np.nan
                    row[f"{m['code']}_current"] = fmt_num(cv)
                    row[f"{m['code']}_last"] = fmt_num(lv)
                    row[f"{m['code']}_change"] = fmt_num(diff)
                    row[f"{m['code']}_rate"] = fmt_num(rate, is_pct=True)
                data.append(row)
            return data

        # 同比表格构造
        def build_yoy(curr_df, last_df):
            data = []
            cats = curr_df["2017分类"].tolist()
            for cat in cats:
                cr = curr_df[curr_df["2017分类"]==cat].iloc[0]
                lr = last_df[last_df["2017分类"]==cat].iloc[0] if cat in last_df["2017分类"].values else pd.Series({m:0 for m in metric_names})
                row = {"分类":cat}
                for m in CORE_METRICS:
                    mn = m["name"]
                    cv = float(cr[mn])
                    lv = float(lr[mn])
                    diff = cv - lv
                    rate = diff/lv if lv !=0 else np.nan
                    row[f"{m['code']}_current"] = fmt_num(cv)
                    row[f"{m['code']}_last"] = fmt_num(lv)
                    row[f"{m['code']}_change"] = fmt_num(diff)
                    row[f"{m['code']}_rate"] = fmt_num(rate, is_pct=True)
                data.append(row)
            return data

        # 吨均表格
        def build_ton(sub_df):
            data = []
            cats = sub_df["2017分类"].tolist()
            for cat in cats:
                r = sub_df[sub_df["2017分类"]==cat].iloc[0]
                qty = float(r["销售数量(MT)"])
                def ton(v):
                    return fmt_num(v/qty) if qty>0 else 0.00
                data.append({
                    "分类":cat,
                    "吨均收入":ton(r["销售收入"]),
                    "吨均毛利":ton(r["销售毛利(扣除运费)"]),
                    "吨均工厂费用":ton(r["工厂费用"]),
                    "吨均营销费用":ton(r["营销公司费用"]),
                    "吨均税前利润":ton(r["税前利润-损益（含套保）"])
                })
            return data

        mom_table = build_mom(agg_curr_month, agg_last_month)
        yoy_table = build_yoy(agg_curr_cum, agg_last_cum)
        ton_mom_curr = build_ton(agg_curr_month)
        ton_mom_last = build_ton(agg_last_month)
        ton_yoy_curr = build_ton(agg_curr_cum)
        ton_yoy_last = build_ton(agg_last_cum)

        # ========== 饼图数据（累计区间2017分类占比） ==========
        pie_base = df_curr_cum.groupby("2017分类").agg({
            "销售数量(MT)":"sum",
            "销售毛利(扣除运费)":"sum"
        }).reset_index()
        pie_qty_labels = pie_base["2017分类"].tolist()
        pie_qty_data = [float(x) for x in pie_base["销售数量(MT)"].tolist()]
        pie_gross_labels = pie_base["2017分类"].tolist()
        pie_gross_data = [float(x) for x in pie_base["销售毛利(扣除运费)"].tolist()]

        # ========== TOP20客户（所属集团重分类）：按销量降序取前20 ==========
        cust_agg = df_curr_cum.groupby("所属集团重分类").agg({
            "销售数量(MT)":"sum",
            "销售收入":"sum",
            "销售毛利(扣除运费)":"sum",
            "工厂费用":"sum",
            "营销公司费用":"sum",
            "税前利润-损益（含套保）":"sum"
        }).reset_index()
        # 按销量降序
        cust_agg = cust_agg.sort_values("销售数量(MT)", ascending=False).reset_index(drop=True)
        top20 = cust_agg.head(20).copy()
        total_all = cust_agg.sum(numeric_only=True)
        top20_sum = top20.sum(numeric_only=True)
        # 其他 = 合计 - TOP20
        other_row = {
            "所属集团重分类": "其他",
            "销售数量(MT)": fmt_num(total_all["销售数量(MT)"] - top20_sum["销售数量(MT)"]),
            "销售收入": fmt_num(total_all["销售收入"] - top20_sum["销售收入"]),
            "销售毛利(扣除运费)": fmt_num(total_all["销售毛利(扣除运费)"] - top20_sum["销售毛利(扣除运费)"]),
            "工厂费用": fmt_num(total_all["工厂费用"] - top20_sum["工厂费用"]),
            "营销公司费用": fmt_num(total_all["营销公司费用"] - top20_sum["营销公司费用"]),
            "税前利润-损益（含套保）": fmt_num(total_all["税前利润-损益（含套保）"] - top20_sum["税前利润-损益（含套保）"])
        }
        # 合计行
        total_row = {
            "所属集团重分类": "合计",
            "销售数量(MT)": fmt_num(total_all["销售数量(MT)"]),
            "销售收入": fmt_num(total_all["销售收入"]),
            "销售毛利(扣除运费)": fmt_num(total_all["销售毛利(扣除运费)"]),
            "工厂费用": fmt_num(total_all["工厂费用"]),
            "营销公司费用": fmt_num(total_all["营销公司费用"]),
            "税前利润-损益（含套保）": fmt_num(total_all["税前利润-损益（含套保）"])
        }
        cust_list = []
        for _,r in top20.iterrows():
            cust_list.append({
                "所属集团重分类": r["所属集团重分类"],
                "销售数量(MT)": fmt_num(r["销售数量(MT)"]),
                "销售收入": fmt_num(r["销售收入"]),
                "销售毛利(扣除运费)": fmt_num(r["销售毛利(扣除运费)"]),
                "工厂费用": fmt_num(r["工厂费用"]),
                "营销公司费用": fmt_num(r["营销公司费用"]),
                "税前利润-损益（含套保）": fmt_num(r["税前利润-损益（含套保）"])
            })
        cust_list.append(other_row)
        cust_list.append(total_row)

        return jsonify({
            "target_branch": target_branch,
            "year": target_year,
            "month": target_month,
            "mom_table": mom_table,
            "yoy_table": yoy_table,
            "ton_mom_curr": ton_mom_curr,
            "ton_mom_last": ton_mom_last,
            "ton_yoy_curr": ton_yoy_curr,
            "ton_yoy_last": ton_yoy_last,
            "core_metrics": CORE_METRICS,
            "mom_dimensions": MOM_DIMENSIONS,
            "yoy_dimensions": YOY_DIMENSIONS,
            "pie_qty_labels": pie_qty_labels,
            "pie_qty_data": pie_qty_data,
            "pie_gross_labels": pie_gross_labels,
            "pie_gross_data": pie_gross_data,
            "customer_table": cust_list
        })
    except Exception as e:
        print("查询异常：", str(e))
        return jsonify({"error": str(e)}),500

if __name__ == "__main__":
    app.run(debug=True)