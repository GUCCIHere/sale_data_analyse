import * as XLSX from 'xlsx';
// 从 Excel 解析数据
async function loadExcelData(env) {
  // 从 KV 获取缓存数据（首次从 R2 或内联数据读取）
  let cached = await env.DATA_STORE.get('sales_data', 'json');
  if (cached) return cached;

  // 首次加载：从 R2 或内联数据读取
  // 方式1：如果 Excel 放在 R2
  // const file = await env.DATA_BUCKET.get('sales_data_N.xlsx');
  // const buffer = await file.arrayBuffer();
  // const workbook = XLSX.read(buffer, { type: 'array' });

  // 方式2：如果 Excel 作为内联数据（需要转换为 Base64）
  // 这里为了演示，提供一个空数组占位
  const workbook = XLSX.read(/* 你的 Excel 数据 */);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(sheet);

  // 过滤数据
  const filtered = rawData.filter(row =>
    row["事业部描述"] === "集团食品工业事业部" &&
    ["外部", "工厂自销"].includes(row["MKT"])
  );

  // 缓存到 KV（有效期1小时）
  await env.DATA_STORE.put('sales_data', JSON.stringify(filtered), { expirationTtl: 3600 });
  return filtered;
}

// 工具函数：数值格式化
function fmtNum(val) {
  if (isNaN(val) || !isFinite(val)) return 0;
  return Math.round(val * 100) / 100;
}

// 分组聚合
function groupSum(data, groupKey, metricKeys) {
  const map = new Map();
  data.forEach(row => {
    const key = String(row[groupKey] || "未分类");
    if (!map.has(key)) {
      const obj = { [groupKey]: key };
      metricKeys.forEach(m => obj[m] = 0);
      map.set(key, obj);
    }
    const current = map.get(key);
    metricKeys.forEach(m => {
      current[m] = (current[m] || 0) + (Number(row[m]) || 0);
    });
  });
  return Array.from(map.values());
}

export async function onRequest(context) {
  try {
    const { branch, year, month } = await context.request.json();
    const allData = await loadExcelData(context.env);

    // 筛选分公司数据
    const branchData = allData.filter(row => String(row["销售组织"]) === String(branch));
    if (branchData.length === 0) {
      return new Response(JSON.stringify({ error: "该分公司无数据" }), {
        status: 400,
        headers: { "Content-Type": "application/json" }
      });
    }

    const targetYear = Number(year);
    const targetMonth = Number(month);
    const lastYear = targetYear - 1;

    // 定义指标
    const coreMetrics = [
      { code: "sales_qty", name: "销售数量(MT)", type: "qty" },
      { code: "sales_revenue", name: "销售收入", type: "amount" },
      { code: "gross_profit", name: "销售毛利(扣除运费)", type: "amount" },
      { code: "factory_cost", name: "工厂费用", type: "amount" },
      { code: "marketing_cost", name: "营销公司费用", type: "amount" },
      { code: "pretax_profit", name: "税前利润-损益（含套保）", type: "amount" }
    ];
    const metricNames = coreMetrics.map(m => m.name);

    // 计算辅助费用字段
    const enrichedData = branchData.map(row => ({
      ...row,
      "工厂费用": (Number(row["工厂税金及附加"]) || 0) +
        (Number(row["工厂销售费用"]) || 0) +
        (Number(row["研发费用"]) || 0) +
        (Number(row["剩余工厂管理费用"]) || 0) +
        (Number(row["工厂财务费用"]) || 0),
      "营销公司费用": (Number(row["营销销售费用"]) || 0) +
        (Number(row["营销管理费用"]) || 0) +
        (Number(row["营销财务费用"]) || 0) +
        (Number(row["营销附加税金"]) || 0)
    }));

    // 数据筛选函数
    function filterData(data, year, month, isCumulative = false) {
      let result = data.filter(row => Number(row["会计年度"]) === year);
      if (isCumulative) {
        result = result.filter(row => Number(row["期间"]) <= month);
      } else {
        result = result.filter(row => Number(row["期间"]) === month);
      }
      return result;
    }

    // 聚合函数
    function aggregateData(data) {
      const grouped = groupSum(data, "2017分类", metricNames);
      // 计算合计行
      const total = { "2017分类": "合计" };
      metricNames.forEach(m => {
        total[m] = fmtNum(grouped.reduce((sum, row) => sum + row[m], 0));
      });
      grouped.push(total);
      return grouped;
    }

    // 获取各数据集
    const currMonth = filterData(enrichedData, targetYear, targetMonth);
    const lastMonth = filterData(enrichedData, targetMonth === 1 ? targetYear - 1 : targetYear,
      targetMonth === 1 ? 12 : targetMonth - 1);
    const currCum = filterData(enrichedData, targetYear, targetMonth, true);
    const lastCum = filterData(enrichedData, lastYear, targetMonth, true);

    // 构建环比/同比表格
    function buildComparison(currData, lastData) {
      const currAgg = aggregateData(currData);
      const lastAgg = aggregateData(lastData);
      const result = [];

      currAgg.forEach(currRow => {
        const cat = currRow["2017分类"];
        const lastRow = lastAgg.find(r => r["2017分类"] === cat) || {};
        const row = { "分类": cat };

        coreMetrics.forEach(m => {
          const cv = currRow[m.name] || 0;
          const lv = lastRow[m.name] || 0;
          const diff = cv - lv;
          const rate = lv !== 0 ? (diff / lv) * 100 : null;
          row[`${m.code}_current`] = fmtNum(cv);
          row[`${m.code}_last`] = fmtNum(lv);
          row[`${m.code}_change`] = fmtNum(diff);
          row[`${m.code}_rate`] = rate !== null ? Math.round(rate * 100) / 100 : null;
        });
        result.push(row);
      });
      return result;
    }

    // 吨均表格
    function buildTonTable(data) {
      const agg = aggregateData(data);
      return agg.map(row => {
        const qty = row["销售数量(MT)"] || 0;
        const result = { "分类": row["2017分类"] };
        if (qty === 0) {
          ["吨均收入", "吨均毛利", "吨均工厂费用", "吨均营销费用", "吨均税前利润"].forEach(k => result[k] = 0);
        } else {
          result["吨均收入"] = fmtNum(row["销售收入"] / qty);
          result["吨均毛利"] = fmtNum(row["销售毛利(扣除运费)"] / qty);
          result["吨均工厂费用"] = fmtNum(row["工厂费用"] / qty);
          result["吨均营销费用"] = fmtNum(row["营销公司费用"] / qty);
          result["吨均税前利润"] = fmtNum(row["税前利润-损益（含套保）"] / qty);
        }
        return result;
      });
    }

    // 构建返回数据
    const responseData = {
      target_branch: branch,
      year: targetYear,
      month: targetMonth,
      mom_table: buildComparison(currMonth, lastMonth),
      yoy_table: buildComparison(currCum, lastCum),
      ton_mom_curr: buildTonTable(currMonth),
      ton_mom_last: buildTonTable(lastMonth),
      ton_yoy_curr: buildTonTable(currCum),
      ton_yoy_last: buildTonTable(lastCum),
      core_metrics: coreMetrics,
      mom_dimensions: [
        { code: "current", label: "本月" },
        { code: "last", label: "上月" },
        { code: "change", label: "增减额" },
        { code: "rate", label: "变动率(%)" }
      ],
      yoy_dimensions: [
        { code: "current", label: "本年累计" },
        { code: "last", label: "去年同期" },
        { code: "change", label: "增减额" },
        { code: "rate", label: "变动率(%)" }
      ]
    };

    return new Response(JSON.stringify(responseData), {
      headers: { "Content-Type": "application/json" }
    });

  } catch (error) {
    return new Response(JSON.stringify({ error: error.message }), {
      status: 500,
      headers: { "Content-Type": "application/json" }
    });
  }
}