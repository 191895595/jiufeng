"""
从飞书电子表格拉取数据 → 生成 data.json
用于 GitHub Actions 定时运行
"""
import json
import os
import urllib.request
import urllib.parse
from datetime import datetime, timedelta

FEISHU_APP_ID = os.environ["FEISHU_APP_ID"]
FEISHU_APP_SECRET = os.environ["FEISHU_APP_SECRET"]
SPREADSHEET_TOKEN = os.environ["SPREADSHEET_TOKEN"]
SHEET_ID = "0XGxqd"


def get_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    data = json.dumps({"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    resp = json.loads(urllib.request.urlopen(req).read())
    return resp["tenant_access_token"]


def excel_date(serial):
    return (datetime(1899, 12, 30) + timedelta(days=int(serial))).strftime("%Y-%m-%d")


def fetch_all():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # 获取行数
    meta_url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}/sheets/query"
    meta = json.loads(urllib.request.urlopen(urllib.request.Request(meta_url, headers=headers)).read())
    total = meta["data"]["sheets"][0]["grid_properties"]["row_count"]

    # 分批读取
    all_rows = []
    for start in range(1, total + 1, 5000):
        end = min(start + 4999, total)
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{urllib.parse.quote(f'{SHEET_ID}!A{start}:D{end}')}"
        resp = json.loads(urllib.request.urlopen(urllib.request.Request(url, headers=headers)).read())
        vals = resp["data"]["valueRange"]["values"]
        all_rows.extend(vals if start > 1 else vals)

    # 解析
    records = []
    for row in all_rows[1:]:
        if len(row) < 4:
            continue
        try:
            d = excel_date(row[0]) if isinstance(row[0], (int, float)) else str(row[0]).strip()
            ch = str(row[1]).strip()
            sub = str(row[2]).strip().replace("\n", "")
            s = float(row[3])
            if s > 0 and d.startswith("202"):
                records.append({"日期": d, "渠道": ch, "细分渠道": sub, "销售额": round(s, 2)})
        except:
            continue

    total_sales = sum(r["销售额"] for r in records)
    dates = sorted(set(r["日期"] for r in records))

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
