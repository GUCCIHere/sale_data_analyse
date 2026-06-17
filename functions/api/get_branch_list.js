export async function onRequest(context) {
  try {
    // 从 KV 存储读取数据（首次从 Excel 解析并缓存）
    const data = await getCachedData(context.env);
    const branches = [...new Set(data.map(row => row["销售组织"]))].sort();
    return new Response(JSON.stringify({ branches }), {
      headers: { "Content-Type": "application/json" }
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: error.message, branches: [] }), {
      status: 500,
      headers: { "Content-Type": "application/json" }
    });
  }
}