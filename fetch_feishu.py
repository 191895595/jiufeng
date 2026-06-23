"""
从飞书电子表格拉取数据 → 生成 data.json
用于 GitHub Actions 定时运行
"""
import json
import os
import urllib.request
import urllib.parse
import re
from datetime import datetime, timedelta

FEISHU_APP_ID = os.environ["FEISHU_APP_ID"]
FEISHU_APP_SECRET = os.environ["FEISHU_APP_SECRET"]
SPREADSHEET_TOKEN = os.environ["SPREADSHEET_TOKEN"]
TARGET_SHEET_NAME = "Sheet1"  # 始终只读 Sheet1，忽略其他工作簿


def get_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    data = json.dumps({"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    resp = json.loads(urllib.request.urlopen(req).read())
    return resp["tenant_access_token"]


def excel_date(serial):
    """将 Excel 数字日期转为 YYYY-MM-DD"""
    return (datetime(1899, 12, 30) + timedelta(days=int(serial))).strftime("%Y-%m-%d")


def parse_date(val):
    """尝试解析日期，返回 YYYY-MM-DD 或 None"""
    if val is None:
        return None
    
    # 数字类型 → Excel 序列号
    if isinstance(val, (int, float)):
        try:
            return excel_date(val)
        except:
            return None
    
    # 字符串类型
    s = str(val).strip()
    if not s:
        return None
    
    # 尝试匹配各种日期格式
    patterns = [
        (r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})', lambda m: f"{m[1]}-{m[2].zfill(2)}-{m[3].zfill(2)}"),
        (r'^(\d{1,2})[-/](\d{1,2})[-/](\d{4})', lambda m: f"{m[3]}-{m[1].zfill(2)}-{m[2].zfill(2)}"),
    ]
    for pat, fmt in patterns:
        m = re.match(pat, s)
        if m:
            try:
                d = fmt(m.groups() if len(m.groups())==3 else [m.group(1), m.group(2), m.group(3)])
                # 验证日期合法性
                datetime.strptime(d, "%Y-%m-%d")
                return d
            except:
                pass
    
    # 尝试直接用 datetime 解析
    for fmt_str in ["%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"]:
        try:
            return datetime.strptime(s, fmt_str).strftime("%Y-%m-%d")
        except:
            continue
    
    return None


def fetch_all():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # 获取所有工作表，找到 Sheet1
    meta_url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}/sheets/query"
    meta = json.loads(urllib.request.urlopen(urllib.request.Request(meta_url, headers=headers)).read())
    sheets = meta["data"]["sheets"]
    
    sheet1 = next((s for s in sheets if s["title"] == TARGET_SHEET_NAME), None)
    if not sheet1:
        raise Exception(f"未找到工作表「{TARGET_SHEET_NAME}」，当前工作表：{[s['title'] for s in sheets]}")
    
    sheet_id = sheet1["sheet_id"]
    total = sheet1["grid_properties"]["row_count"]
    print(f"📋 读取工作表: {TARGET_SHEET_NAME} (id={sheet_id}, {total}行)")

    # 分批读取
    all_rows = []
    for start in range(1, total + 1, 5000):
        end = min(start + 4999, total)
        params = urllib.parse.urlencode({
            'valueRenderOption': 'ToString',
            'dateTimeRenderOption': 'FormattedString'
        })
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{urllib.parse.quote(f'{sheet_id}!A{start}:D{end}')}?{params}"
        resp = json.loads(urllib.request.urlopen(urllib.request.Request(url, headers=headers)).read())
        vals = resp["data"]["valueRange"]["values"]
        all_rows.extend(vals if start > 1 else vals)

    print(f"📋 API 返回 {len(all_rows)} 行（含表头）")

    # 解析（跳过表头第一行）
    records = []
    stats = {"total": 0, "no_date": 0, "no_channel": 0, "no_sales": 0, "bad_date": 0, "empty": 0}
    
    for row in all_rows[1:]:
        stats["total"] += 1
        
        # 跳过空行（4列全空）
        if not row or all(c is None for c in row) or len(row) < 4:
            stats["empty"] += 1
            continue
        
        try:
            # 日期
            d = parse_date(row[0])
            if not d:
                stats["no_date"] += 1
                continue
            if not d.startswith("202"):
                stats["bad_date"] += 1
                continue
            
            # 渠道
            ch = str(row[1]).strip() if row[1] else ''
            if not ch:
                stats["no_channel"] += 1
                continue
            # 跳过汇总行
            if any(kw in ch for kw in ['合计', '总计', '小计', '汇总', '平均', 'sum', 'total']):
                continue
            
            # 细分渠道
            sub = str(row[2]).strip().replace("\n", "").replace("\r", "") if row[2] else ''
            
            # 销售额（允许负数，表示退款）
            s = float(row[3]) if row[3] is not None else 0
            if s == 0:  # 只跳过真正的0值，负数保留
                stats["no_sales"] += 1
                continue
            
            records.append({
                "日期": d,
                "渠道": ch,
                "细分渠道": sub,
                "销售额": round(s, 2)
            })
        except Exception as e:
            continue

    print(f"📋 总行数: {stats['total']} | 有效: {len(records)} | 跳过: 空行{stats['empty']} 无日期{stats['no_date']} 日期异常{stats['bad_date']} 无渠道{stats['no_channel']} 销售额≤0{stats['no_sales']}")

    # 按 (日期, 渠道, 细分渠道) 合并去重（飞书表格中同一店铺同一天可能有多行不同产品）
    before_merge = len(records)
    merged = {}
    for r in records:
        key = (r["日期"], r["渠道"], r["细分渠道"])
        if key not in merged:
            merged[key] = r
        else:
            merged[key]["销售额"] = round(merged[key]["销售额"] + r["销售额"], 2)
    records = list(merged.values())
    
    if before_merge != len(records):
        print(f"📋 去重合并: {before_merge} → {len(records)} 条")

    total_sales = sum(r["销售额"] for r in records)
    dates = sorted(set(r["日期"] for r in records))

    # 输出每月汇总（方便对比透视表）
    monthly = {}
    for r in records:
        m = r["日期"][:7]
        monthly[m] = monthly.get(m, 0) + r["销售额"]
    print("📋 每月汇总:")
    for m in sorted(monthly.keys()):
        print(f"  {m}: ¥{monthly[m]:,.2f}")
    print(f"  💰 总计: ¥{total_sales:,.2f}")

    return {
        "records": records,
        "total_sales": round(total_sales, 2),
        "record_count": len(records),
        "date_range": f"{dates[0]} ~ {dates[-1]}",
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }


if __name__ == "__main__":
    result = fetch_all()
    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)
    print(f"✅ {result['record_count']} records, ¥{result['total_sales']:,.2f}, {result['date_range']}")
