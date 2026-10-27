"""为筛选后的相机补抓详情参数（像元尺寸、靶面尺寸、快门等）。

数据来源：官网详情页同一接口
  GET https://www.hikrobotics.com/cn/Api/Foreground/Vision/VisionProductConfig?id=<productId>

用法：python scripts/fetch_camera_params.py
输出：data/hik_camera_params.json
"""

import json
import os
import re
import time
import urllib.request
from concurrent.futures import ThreadPoolExecutor

API = "https://www.hikrobotics.com/cn/Api/Foreground/Vision/VisionProductConfig?id="
KEEP = ["传感器型号", "分辨率", "像元尺寸", "靶面尺寸", "最大帧率", "数据接口",
        "黑白/彩色", "快门", "信噪比", "曝光时间", "功耗", "镜头接口", "尺寸"]


def res_mp(s):
    m = re.match(r"\s*(\d+)\s*[×xX]\s*(\d+)", s or "")
    return round(int(m.group(1)) * int(m.group(2)) / 1e6, 2) if m else 0


def load():
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    with open(os.path.join(here, "data", "hik_products.json"), encoding="utf-8") as f:
        return here, json.load(f)


def pick(cam):
    """选型工具需要的子集：黑白、GigE/USB3、1.2MP~20MP。"""
    s = cam["spec"]
    mp = res_mp(s.get("分辨率") or "")
    iface = s.get("数据接口") or ""
    mono = "黑白" in (s.get("黑白/彩色") or "")
    return mono and 1.2 <= mp <= 20 and ("GigE" in iface or "USB" in iface)


def fetch(cam):
    try:
        req = urllib.request.Request(API + str(cam["id"]), headers={
            "User-Agent": "Mozilla/5.0", "Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=30) as r:
            j = json.loads(r.read().decode())
        params = {}
        for item in j.get("data") or []:
            n = (item.get("name") or "").strip()
            if n in KEEP:
                params[n] = (item.get("value") or "").strip()
        return {"model": cam["model"], "desc": cam["desc"], "id": cam["id"],
                "img": cam["img"], "params": params}
    except Exception as e:
        return {"model": cam["model"], "id": cam["id"], "error": str(e)[:80]}


def main():
    here, data = load()
    cams = [c for c in data["groups"]["camera"]["rows"] if pick(c)]
    print("selected %d cameras" % len(cams))
    out, done = [], 0
    with ThreadPoolExecutor(max_workers=8) as ex:
        for r in ex.map(fetch, cams):
            out.append(r)
            done += 1
            if done % 25 == 0:
                print("  %d/%d" % (done, len(cams)))
    ok = [r for r in out if "params" in r and r["params"]]
    path = os.path.join(here, "data", "hik_camera_params.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"updated_at": time.strftime("%Y-%m-%d %H:%M:%S"), "rows": ok},
                  f, ensure_ascii=False, separators=(",", ":"))
    print("ok=%d fail=%d -> %s (%.1f KB)" % (len(ok), len(out) - len(ok), path,
                                             os.path.getsize(path) / 1024))


if __name__ == "__main__":
    main()
