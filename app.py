# app.py - 完整修正版
import pandas as pd
import numpy as np
from flask import Flask, render_template_string, request, jsonify
import os
import warnings
warnings.filterwarnings('ignore')

app = Flask(__name__)

EXCEL_FILE_PATH = "sales_data.xlsx"  # 请修改为实际路径

cached_agg_df = None
cached_metadata = None

def load_and_process_data():
    global cached_agg_df, cached_metadata
    if cached_agg_df is not None:
        return cached_agg_df, cached_metadata
    if not os.path.exists(EXCEL_FILE_PATH):
        raise FileNotFoundError(f"Excel文件不存在: {EXCEL_FILE_PATH}")
    print("正在加载Excel文件...")
    required_columns = ['会计年度', '期间', '销售组织', '事业部描述', 'MKT',
                        '销售数量(MT)', '销售毛利(扣除运费)', '税前利润-损益（含套保）']
    df = pd.read_excel(EXCEL_FILE_PATH)
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"缺少列: {missing_cols}")
    df = df[df['事业部描述'] == '食品工业事业部']
    df = df[df['MKT'].isin(['外部', '工厂自销'])]
    def parse_period(p):
        if pd.isna(p):
            return None
        s = str(p).strip().replace('月', '')
        try:
            return int(s)
        except:
            return None
    df['期间_clean'] = df['期间'].apply(parse_period)
    df = df.dropna(subset=['期间_clean'])
    df['会计年度'] = pd.to_numeric(df['会计年度'], errors='coerce').astype('Int64')
    df = df.dropna(subset=['会计年度'])
    for col in ['销售数量(MT)', '销售毛利(扣除运费)', '税前利润-损益（含套保）']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    agg_df = df.groupby(['销售组织', '会计年度', '期间_clean'], as_index=False)[
        ['销售数量(MT)', '销售毛利(扣除运费)', '税前利润-损益（含套保）']].sum()
    agg_df.rename(columns={'期间_clean': '期间'}, inplace=True)
    sales_orgs = sorted(agg_df['销售组织'].unique())
    years = sorted(agg_df['会计年度'].unique())
    months = sorted(agg_df['期间'].unique())
    cached_agg_df = agg_df
    cached_metadata = {'sales_orgs': sales_orgs, 'years': [int(y) for y in years], 'months': [int(m) for m in months]}
    print(f"加载成功，共{len(agg_df)}条记录")
    return agg_df, cached_metadata

def get_month_data(agg_df, sales_org, year, month):
    subset = agg_df[(agg_df['销售组织'] == sales_org) & (agg_df['会计年度'] == year) & (agg_df['期间'] == month)]
    if subset.empty:
        return None
    return subset.iloc[0].to_dict()

def calculate_ratio(current, previous):
    if previous is None or pd.isna(previous) or previous == 0:
        return None
    if current is None or pd.isna(current):
        return None
    return round((current - previous) / previous * 100, 2)

@app.route('/')
def index():
    try:
        _, metadata = load_and_process_data()
    except Exception as e:
        return f"<h3>加载失败</h3><p>{str(e)}</p>", 500
    html_template = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>销售数据分析看板</title>
        <style>
            body { font-family: 'Segoe UI', sans-serif; margin: 20px; background: #f5f5f5; }
            .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            h1 { color: #2c3e50; border-bottom: 2px solid #3498db; }
            .filters { background: #ecf0f1; padding: 15px; border-radius: 6px; margin-bottom: 25px; display: flex; gap: 15px; flex-wrap: wrap; align-items: flex-end; }
            .filter-group { display: flex; flex-direction: column; gap: 5px; }
            select, button { padding: 8px 12px; font-size: 14px; border-radius: 4px; }
            button { background: #3498db; color: white; border: none; cursor: pointer; }
            button:hover { background: #2980b9; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #ddd; padding: 12px; text-align: center; }
            th { background: #3498db; color: white; }
            .positive { color: #27ae60; font-weight: bold; }
            .negative { color: #e74c3c; font-weight: bold; }
            .loading { text-align: center; padding: 20px; display: none; }
            .error { color: #e74c3c; margin-top: 15px; padding: 10px; background: #fadbd8; border-radius: 4px; display: none; }
            .info-note { margin-top: 15px; font-size: 12px; color: #7f8c8d; text-align: center; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 销售数据同比环比分析</h1>
            <div class="filters">
                <div class="filter-group"><label>🏢 分公司</label><select id="salesOrg">{% for org in sales_orgs %}<option value="{{ org }}">{{ org }}</option>{% endfor %}</select></div>
                <div class="filter-group"><label>📅 年份</label><select id="year">{% for y in years %}<option value="{{ y }}">{{ y }}</option>{% endfor %}</select></div>
                <div class="filter-group"><label>🗓️ 月份</label><select id="month">{% for m in months %}<option value="{{ m }}">{{ m }}月</option>{% endfor %}</select></div>
                <div class="filter-group"><button id="queryBtn">🔍 查询</button></div>
            </div>
            <div id="loading" class="loading">⏳ 计算中...</div>
            <div id="resultTable"></div>
            <div id="errorMsg" class="error"></div>
            <div class="info-note">💡 环比 = (本月-上月)/上月×100% ；同比 = (本月-去年同月)/去年同月×100%<br>过滤条件：事业部描述="食品工业事业部"，MKT="外部"或"工厂自销"</div>
        </div>
        <script>
            const queryBtn = document.getElementById('queryBtn');
            const loadingDiv = document.getElementById('loading');
            const resultDiv = document.getElementById('resultTable');
            const errorDiv = document.getElementById('errorMsg');
            async function fetchData() {
                const salesOrg = document.getElementById('salesOrg').value;
                const year = parseInt(document.getElementById('year').value);
                const month = parseInt(document.getElementById('month').value);
                loadingDiv.style.display = 'block';
                resultDiv.innerHTML = '';
                errorDiv.style.display = 'none';
                try {
                    const resp = await fetch('/api/analysis', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ sales_org: salesOrg, year: year, month: month })
                    });
                    const data = await resp.json();
                    if (!resp.ok) throw new Error(data.error);
                    renderTable(data);
                } catch(err) {
                    errorDiv.style.display = 'block';
                    errorDiv.innerHTML = '❌ ' + err.message;
                } finally {
                    loadingDiv.style.display = 'none';
                }
            }
            function renderTable(data) {
                if (!data.metrics || data.metrics.length === 0) {
                    resultDiv.innerHTML = '<div style="padding:20px;text-align:center;">无数据</div>';
                    return;
                }
                let html = '<table><thead><tr><th>指标</th><th>本月值</th><th>环比(%)</th><th>同比(%)</th></tr></thead><tbody>';
                for (let item of data.metrics) {
                    let val = item.current_value.toLocaleString('zh-CN', { minimumFractionDigits: 2 });
                    let mom = (item.mom !== null) ? item.mom.toFixed(2)+'%' : '—';
                    let yoy = (item.yoy !== null) ? item.yoy.toFixed(2)+'%' : '—';
                    let momClass = (item.mom !== null && item.mom>=0) ? 'positive' : (item.mom!==null?'negative':'');
                    let yoyClass = (item.yoy !== null && item.yoy>=0) ? 'positive' : (item.yoy!==null?'negative':'');
                    html += `<tr><td><strong>${item.name}</strong></td><td>${val}</td><td class="${momClass}">${mom}</td><td class="${yoyClass}">${yoy}</td></tr>`;
                }
                html += '</tbody></table>';
                resultDiv.innerHTML = html;
            }
            window.addEventListener('load', fetchData);
            queryBtn.addEventListener('click', fetchData);
        </script>
    </body>
    </html>
    """
    return render_template_string(html_template,
                                 sales_orgs=metadata['sales_orgs'],
                                 years=metadata['years'],
                                 months=metadata['months'])

@app.route('/api/analysis', methods=['POST'])
def get_analysis():
    try:
        data = request.json
        sales_org = data.get('sales_org')
        year = int(data.get('year'))
        month = int(data.get('month'))
        agg_df, _ = load_and_process_data()
        current = get_month_data(agg_df, sales_org, year, month)
        if current is None:
            return jsonify({'error': f'无数据 {sales_org} {year}年{month}月'}), 404
        prev_year = year if month>1 else year-1
        prev_month = month-1 if month>1 else 12
        prev_data = get_month_data(agg_df, sales_org, prev_year, prev_month)
        last_year_data = get_month_data(agg_df, sales_org, year-1, month)
        metrics_config = [
            {'name': '销售数量 (MT)', 'key': '销售数量(MT)'},
            {'name': '销售毛利 (扣除运费)', 'key': '销售毛利(扣除运费)'},
            {'name': '税前利润-损益 (含套保)', 'key': '税前利润-损益（含套保）'}
        ]
        result = []
        for m in metrics_config:
            cur = current.get(m['key'], 0)
            prev = prev_data.get(m['key']) if prev_data else None
            ly = last_year_data.get(m['key']) if last_year_data else None
            mom = calculate_ratio(cur, prev)
            yoy = calculate_ratio(cur, ly)
            result.append({'name': m['name'], 'current_value': cur, 'mom': mom, 'yoy': yoy})
        return jsonify({'sales_org': sales_org, 'year': year, 'month': month, 'metrics': result})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 在文件末尾
if __name__ == '__main__':
    # 检查是否在 Vercel 环境运行
    is_vercel = os.environ.get('VERCEL') or os.environ.get('NOW_REGION')
    
    if not is_vercel:
        # 本地开发环境
        if not os.path.exists(EXCEL_FILE_PATH):
            print(f"\n⚠️ 未找到文件: {EXCEL_FILE_PATH}\n")
        else:
            print("启动本地开发服务器: http://127.0.0.1:5000")
            app.run(debug=True, host='127.0.0.1', port=5000)
    else:
        # Vercel 环境不需要启动服务器
        print("Running on Vercel - serverless mode")

# if __name__ == '__main__':
#     if not os.path.exists(EXCEL_FILE_PATH):
#         print(f"\n⚠️ 未找到文件: {EXCEL_FILE_PATH}\n请修改 EXCEL_FILE_PATH 变量\n")
#     print("启动服务: http://127.0.0.1:5000")
#     app.run(debug=True, host='127.0.0.1', port=5000)