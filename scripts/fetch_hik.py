"""抓取海康机器人（hikrobotics.com）机器视觉产品真实型号与参数。

数据来源：官网产品页同一接口 VisionProductContent
  POST https://www.hikrobotics.com/cn/Api/Foreground/Vision/VisionProductContent
  body: {firstModuleId, secondaryModuleId, page, size, screening, showEol}

用法：python scripts/fetch_hik.py
输出：data/hik_products.json（相机 / 镜头 / 光源 + 列名 + 抓取时间）
"""

import json
import os
import time
import urllib.request

API = "https://www.hikrobotics.com/cn/Api/Foreground/Vision/VisionProductContent"

# (key, 名称, firstModuleId, secondaryModuleId)
GROUPS = [
    ("camera", "工业面阵相机", 27, 259),
    ("lens_fa", "FA 镜头", 40, 49),
    ("lens_tele", "远心镜头", 40, 50),
    ("light_std", "标准系列光源", 245, 296),
    ("light_hik", "海康风格系列光源", 245, 294),
]

PAGE = 100


def call(fm, sm, page, size):
    body = json.dumps({
        "firstModuleId": fm,
        "secondaryModuleId": sm,
        "page": page,
        "size": size,
        "screening": [],
        "showEol": False,
    }).encode()
    req = urllib.request.Request(API, data=body, headers={
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://www.hikrobotics.com/cn/machinevision/visionproduct",
    })
    with urllib.request.urlopen(req, timeout=40) as r:
        return json.loads(r.read().decode())


def fetch_group(fm, sm):
    first = call(fm, sm, 1, PAGE)
    data = first.get("data") or {}
    content = data.get("VisionProductContent") or {}
    total = content.get("total", 0)
    cols = [c["name"] for c in sorted(
        data.get("VisionConfigDisplay") or [], key=lambda x: x.get("sort", 0))]
    records = list(content.get("records") or [])
    configs = list(data.get("VisionProductConfig") or [])
    pages = (total + PAGE - 1) // PAGE
    for p in range(2, pages + 1):
        try:
            more = call(fm, sm, p, PAGE).get("data") or {}
            c2 = more.get("VisionProductContent") or {}
            records.extend(c2.get("records") or [])
            configs.extend(more.get("VisionProductConfig") or [])
            time.sleep(0.2)
        except Exception as e:  # 单页失败不影响整体
            print("  page %d failed: %s" % (p, e))
    rows = []
    for rec, cfg in zip(records, configs):
        if not isinstance(cfg, list):
            continue
        rows.append({
            "model": (rec.get("productModel") or "").strip(),
            "desc": (rec.get("productName") or "").strip(),
            "id": rec.get("id"),
            "status": rec.get("productTrends"),
            "img": rec.get("previewUrl"),
            "spec": dict(zip(cols, cfg)),
        })
    return {"columns": cols, "total": total, "rows": rows}


def main():
    out = {"source": "https://www.hikrobotics.com/cn/machinevision/visionproduct",
           "api": API, "fetched_at": time.strftime("%Y-%m-%d %H:%M:%S"),
           "groups": {}}
    for key, name, fm, sm in GROUPS:
        print("fetching %s (%d/%d) ..." % (name, fm, sm))
        g = fetch_group(fm, sm)
        g["name"] = name
        out["groups"][key] = g
        print("  -> %d / %d rows, columns=%s" % (len(g["rows"]), g["total"], g["columns"]))
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    dst = os.path.join(here, "data")
    os.makedirs(dst, exist_ok=True)
    path = os.path.join(dst, "hik_products.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print("written %s  (%.1f KB)" % (path, os.path.getsize(path) / 1024))


if __name__ == "__main__":
    main()
