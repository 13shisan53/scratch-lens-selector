"""把抓取到的海康数据整理成选型工具可用的紧凑结构，并注入 index.html。

输入：data/hik_products.json、data/hik_camera_params.json
输出：data/hik-data.js（便于单独引用）、index.html（把数据内联进 __DATA__ 占位符）

用法：python scripts/build_site.py
"""

import json
import os
import re

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def rnum(s, default=0.0):
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s or "")
    return float(m.group(0)) if m else default


def parse_res(s):
    m = re.search(r"(\d+)\s*[×xX*]\s*(\d+)", s or "")
    return (int(m.group(1)), int(m.group(2))) if m else (0, 0)


def norm_iface(s):
    s = (s or "").lower()
    if "10 gigabit" in s or "10gige" in s:
        return "10GigE"
    if "coaxpress" in s or "cxp" in s:
        return "CoaXPress"
    if "camera link" in s or "cameralink" in s:
        return "CameraLink"
    if "usb" in s:
        return "USB3.0"
    if "gigabit" in s or "gige" in s:
        return "GigE"
    return "其他"


def sensor_inch(s):
    """'1/2.9”' -> '1/2.9\"'；统一引号，去掉空格。"""
    t = (s or "").strip().replace("’’", '"').replace("”", '"').replace("''", '"').replace("’", '"')
    return t


def light_size(model):
    """从型号里解析环光外径 / 条形长度（mm）：MV-LRDS-H-230-60-W2 -> 230"""
    m = re.match(r"MV-L[A-Z]{3}(?:-H)?-(\d+(?:\.\d+)?)", model or "")
    return float(m.group(1)) if m else 0.0


def light_angle(desc):
    m = re.search(r"(\d+)\s*°", desc or "")
    return int(m.group(1)) if m else 0


def circle_mm(s):
    """'Φ9 mm(1/1.8'')' -> 9.0；'Φ11 mm' -> 11.0"""
    m = re.search(r"[Φφ]\s*(\d+(?:\.\d+)?)", s or "")
    return float(m.group(1)) if m else 0.0


LIGHT_TYPE = [
    ("MV-LGSS", "穹顶光源"), ("MV-LGES", "平面无影光源"), ("MV-LRSS", "环形无影光源"),
    ("MV-LRDS", "环形光源"), ("MV-LLDS", "条形光源"), ("MV-LLAS", "组合条形光源"),
    ("MV-LTHS", "高亮线光源"), ("MV-LTDS", "线光源"), ("MV-LTVS", "VCSEL线光源"),
    ("MV-LTVM", "VCSEL线光源"), ("MV-LBSS", "面光源(背光)"), ("MV-LBPS", "平行面光源"),
    ("MV-LBES", "中孔面光源"), ("MV-LBFH", "侧导平行面光源"), ("MV-LCDS", "同轴光源"),
    ("MV-LTZ", "同轴光源"), ("MV-LDSS", "点光源"), ("MV-LDFM", "点光源"),
    ("MV-LDFS", "点光源"), ("MV-CTH", "开孔面光源"),
]


def light_type(model, desc):
    for pre, name in LIGHT_TYPE:
        if model.startswith(pre):
            return name
    if "环" in (desc or ""):
        return "环形光源"
    if "面" in (desc or ""):
        return "面光源"
    if "线" in (desc or ""):
        return "线光源"
    return "其他"


def light_color(model, desc):
    d = desc or ""
    for k, v in (("红外", "IR"), ("白色", "白光"), ("红色", "红光"), ("蓝色", "蓝光"), ("绿色", "绿光")):
        if k in d:
            return v
    tail = model.split("-")[-1].upper()
    if tail.startswith("IR"):
        return "IR"
    return {"W": "白光", "R": "红光", "B": "蓝光", "G": "绿光"}.get(tail[0] if tail else "", "其他")


def build():
    with open(os.path.join(HERE, "data", "hik_products.json"), encoding="utf-8") as f:
        raw = json.load(f)
    with open(os.path.join(HERE, "data", "hik_camera_params.json"), encoding="utf-8") as f:
        params = json.load(f)

    cameras = []
    for r in params["rows"]:
        p = r["params"]
        w, h = parse_res(p.get("分辨率"))
        px = rnum(p.get("像元尺寸"))
        if not (w and h and px):
            continue
        cameras.append({
            "m": r["model"], "w": w, "h": h, "px": px,
            "inch": sensor_inch(p.get("靶面尺寸")),
            "fps": rnum(p.get("最大帧率")),
            "if": norm_iface(p.get("数据接口")),
            "mount": p.get("镜头接口") or "",
            "exp": (p.get("曝光时间") or "").replace(" ", ""),
            "mp": round(w * h / 1e6, 1),
            "d": r.get("desc") or "",
        })
    cameras.sort(key=lambda c: (c["w"] * c["h"], c["m"]))

    lenses = []
    for r in raw["groups"]["lens_fa"]["rows"]:
        s = r["spec"]
        f = rnum(s.get("焦距"))
        if not f:
            continue
        lenses.append({"m": r["model"], "f": f, "fn": s.get("F数") or "",
                       "circle": circle_mm(s.get("像面尺寸")),
                       "mount": s.get("接口类型") or "", "d": r["desc"]})
    lenses.sort(key=lambda x: (x["f"], x["m"]))

    tele = []
    for r in raw["groups"]["lens_tele"]["rows"]:
        s = r["spec"]
        mag = rnum(s.get("放大倍率"))
        wd = s.get("工作距离") or ""
        if not mag:
            continue
        tele.append({"m": r["model"], "mag": mag, "wd": rnum(wd), "wdTxt": wd,
                     "circle": circle_mm(s.get("像面尺寸")),
                     "mount": s.get("接口类型") or "", "d": r["desc"]})
    tele.sort(key=lambda x: (x["mag"], x["m"]))

    lights, seen = [], set()
    for g in ("light_std", "light_hik"):
        for r in raw["groups"][g]["rows"]:
            if r["model"] in seen:
                continue
            seen.add(r["model"])
            s = r["spec"]
            lights.append({
                "m": r["model"], "t": light_type(r["model"], r["desc"]),
                "color": light_color(r["model"], r["desc"]),
                "if": s.get("接口类型") or "", "p": rnum(s.get("功率")),
                "pTx": s.get("功率") or "", "d": r["desc"],
                "size": light_size(r["model"]), "angle": light_angle(r["desc"]),
            })
    lights.sort(key=lambda x: (x["t"], x["m"]))

    data = {
        "meta": {
            "source": raw["source"],
            "api": raw["api"],
            "fetchedAt": raw["fetched_at"],
            "counts": {"cameras": len(cameras), "lenses": len(lenses),
                       "tele": len(tele), "lights": len(lights)},
        },
        "cameras": cameras, "lenses": lenses, "tele": tele, "lights": lights,
    }
    js = "window.HIK = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n"
    with open(os.path.join(HERE, "data", "hik-data.js"), "w", encoding="utf-8") as f:
        f.write(js)

    tpl_path = os.path.join(HERE, "scripts", "index.template.html")
    with open(tpl_path, encoding="utf-8") as f:
        tpl = f.read()
    html = tpl.replace("/*__DATA__*/", js)
    with open(os.path.join(HERE, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)
    print("cameras=%d lenses=%d tele=%d lights=%d" % (len(cameras), len(lenses), len(tele), len(lights)))
    print("hik-data.js %.1f KB / index.html %.1f KB" % (
        os.path.getsize(os.path.join(HERE, "data", "hik-data.js")) / 1024,
        os.path.getsize(os.path.join(HERE, "index.html")) / 1024))


if __name__ == "__main__":
    build()
