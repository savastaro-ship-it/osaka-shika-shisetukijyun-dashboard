# -*- coding: utf-8 -*-
"""
大阪府内の地区別集計モジュール。

医療機関コード（E列・医療機関番号）の先頭2桁を、大阪府歯科医師会の地区に
振り分けて、地区×施設基準の集計を出力する。

振り分けロジック（VBA「地域別ベア集計ツール」＋出力シートから移植）：
- 1対1マッピング（REGION_MAP）
- 合算：堺(01,60-65)・高槻(09,39)・泉佐野泉南(12,45,56,95)・大東四條畷(19,57)
        富田林(49,35)・高石忠岡(53,54)・狭山美原(93,66)
- 住所判定：47豊能郡 → 住所に「能勢町」→箕面／「豊能町」→池田
- 未知コード（98,99,将来の新規）→「その他」（医療機関コード+名称を内訳に保持）

出力：
- data/region/27_current.json  地区×基準の最新値
- data/region/27_history.json  地区×基準の時系列
"""

import io
import json
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from lib import DATA as DIR_DATA, _atomic_write_text, _find_asof_in_sheet, _detect_header_row

DIR_REGION = DIR_DATA / "region"

# ============ 府歯地区マッピング（行政コード全角2桁 → 地区名） ============
REGION_MAP: Dict[str, str] = {
    # --- 大阪市内 ---
    "４１": "北（北＋大淀）",
    "５２": "都島", "０２": "福島", "２８": "此花",
    "９４": "中央区（東＋南）",
    "１８": "西", "０４": "港", "２７": "大正", "１７": "天王寺", "４３": "浪速",
    "１０": "西淀川", "９１": "淀川", "３０": "東淀川", "１５": "東成", "２２": "生野",
    "３１": "旭", "４４": "城東", "９２": "鶴見", "２３": "阿倍野", "５９": "住之江",
    "２０": "住吉", "０８": "東住吉", "５８": "平野", "３３": "西成",
    # --- 北摂 ---
    "１４": "箕面",          # +47豊能郡の能勢町（住所判定）
    "２５": "池田",          # +47豊能郡の豊能町（住所判定）
    "４０": "豊中", "１６": "吹田", "４２": "茨木", "３７": "摂津",
    "０９": "高槻", "３９": "高槻",           # 三島郡（島本町）→高槻
    # --- 北河内 ---
    "２４": "枚方", "３６": "交野", "０３": "寝屋川", "３２": "守口", "２６": "門真",
    "１９": "大東・四條畷", "５７": "大東・四條畷",
    # --- 中河内 ---
    "５０": "東大阪（東大阪東＋東大阪西）",
    "５５": "八尾", "４６": "柏原",
    # --- 南河内 ---
    "３４": "藤井寺", "４８": "松原", "３８": "羽曳野",
    "４９": "富田林", "３５": "富田林",       # 南河内郡→富田林
    "９３": "狭山美原", "６６": "狭山美原",   # 堺市美原区→狭山美原
    "０７": "河内長野",
    # --- 堺 ---
    "０１": "堺",
    "６０": "堺", "６１": "堺", "６２": "堺", "６３": "堺", "６４": "堺", "６５": "堺",
    # --- 泉州 ---
    "５３": "高石忠岡", "５４": "高石忠岡",   # 泉北郡（忠岡町）→高石忠岡
    "０６": "泉大津", "０５": "和泉", "１１": "岸和田", "１３": "貝塚",
    "１２": "泉佐野泉南", "４５": "泉佐野泉南", "５６": "泉佐野泉南", "９５": "泉佐野泉南",
}

TOYONO_CODE = "４７"   # 豊能郡：住所で箕面/池田に振り分け
OTHER_NAME = "その他"

# 地区の表示順（出力シート順）
DISTRICT_ORDER: List[str] = [
    "北（北＋大淀）", "都島", "福島", "此花", "中央区（東＋南）", "西", "港", "大正",
    "天王寺", "浪速", "西淀川", "淀川", "東淀川", "東成", "生野", "旭", "城東", "鶴見",
    "阿倍野", "住之江", "住吉", "東住吉", "平野", "西成",
    "箕面", "池田", "豊中", "吹田", "茨木", "摂津", "高槻",
    "枚方", "交野", "寝屋川", "守口", "門真", "大東・四條畷",
    "東大阪（東大阪東＋東大阪西）", "八尾", "柏原",
    "藤井寺", "松原", "羽曳野", "富田林", "狭山美原", "河内長野",
    "堺", "高石忠岡", "泉大津", "和泉", "岸和田", "貝塚", "泉佐野泉南",
    OTHER_NAME,
]


def _to_wide2(code_str: str) -> str:
    """医療機関番号の先頭2文字を全角化（VBAのStrConv(...,vbWide)相当）"""
    head = str(code_str).strip()[:2]
    return "".join(chr(ord(c) + 0xFEE0) if "0" <= c <= "9" else c for c in head)


def resolve_district(mcode: str, address: str) -> Tuple[str, bool]:
    """医療機関番号と住所から地区名を返す。戻り値: (地区名, 未知コードか)"""
    w2 = _to_wide2(mcode)
    if w2 == TOYONO_CODE:
        addr = str(address or "")
        if "能勢町" in addr:
            return "箕面", False
        if "豊能町" in addr:
            return "池田", False
        return OTHER_NAME, True
    if w2 in REGION_MAP:
        return REGION_MAP[w2], False
    return OTHER_NAME, True


def aggregate_osaka_regions(xlsx_bytes: bytes) -> Optional[dict]:
    """大阪の名簿xlsxから地区×基準の集計を作る。

    戻り値：
    {
      "version": "2026.5", "asof": "令和8年5月1日現在",
      "districts": [
        {"name": "堺", "clinics": 191,
         "standards": [{"kigo","name","count","count_uniq"}, ...]},
        ...
      ],
      "other_members": [{"code": "9812345", "name": "..."}, ...]
    }
    """
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), header=None, dtype=str)
    asof_tuple = _find_asof_in_sheet(raw)
    if not asof_tuple:
        return None
    y, m, d = asof_tuple
    version = f"{2018 + y}.{m}"
    asof_str = f"令和{y}年{m}月{d}日現在"

    hdr = _detect_header_row(raw)
    df = pd.read_excel(io.BytesIO(xlsx_bytes), header=hdr, dtype=str)

    need = ["項番", "医療機関番号", "受理記号"]
    if any(c not in df.columns for c in need):
        return None
    if "区分" in df.columns:
        df = df[df["区分"] == "歯科"]
    df = df.dropna(subset=["医療機関番号"])
    if df.empty:
        return None

    addr_col = "医療機関所在地（住所）" if "医療機関所在地（住所）" in df.columns else None
    name_col = "医療機関名称" if "医療機関名称" in df.columns else None

    # 地区判定（行単位）
    districts = []
    unknown = []
    for _, row in df.iterrows():
        dist, is_unknown = resolve_district(
            row["医療機関番号"],
            row[addr_col] if addr_col else "",
        )
        districts.append(dist)
        if is_unknown:
            unknown.append(row)
    df = df.assign(_district=districts)

    # その他の内訳（医療機関コード＋名称、ユニーク）
    other_members: List[dict] = []
    seen = set()
    for row in unknown:
        code = str(row["医療機関番号"]).strip()
        if code in seen:
            continue
        seen.add(code)
        nm = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
        other_members.append({"code": code, "name": nm})

    # 地区ごとに集計（母数=項番のユニーク数、共通パーサと同一ロジック）
    out_districts = []
    for dist_name, g in df.groupby("_district"):
        total = int(g["項番"].dropna().nunique())
        sub = g.dropna(subset=["受理記号"]).copy()
        sub["受理記号"] = sub["受理記号"].astype(str).str.strip()
        sub = sub[sub["受理記号"] != ""]
        standards = []
        if not sub.empty:
            agg = (sub.groupby("受理記号")
                      .agg(name=("受理届出名称", "first"),
                           count=("項番", "size"),
                           count_uniq=("項番", "nunique"))
                      .reset_index()
                      .sort_values("count", ascending=False))
            standards = [
                {"kigo": str(r["受理記号"]),
                 "name": (str(r["name"]) if pd.notna(r["name"]) else ""),
                 "count": int(r["count"]),
                 "count_uniq": int(r["count_uniq"])}
                for _, r in agg.iterrows()
            ]
        out_districts.append({"name": dist_name, "clinics": total, "standards": standards})

    # 表示順に整列（DISTRICT_ORDERにないものは末尾）
    order = {n: i for i, n in enumerate(DISTRICT_ORDER)}
    out_districts.sort(key=lambda dct: order.get(dct["name"], 999))

    return {
        "version": version,
        "asof": asof_str,
        "districts": out_districts,
        "other_members": other_members,
    }


def _vkey(v: str):
    y, m = v.split(".")
    return (int(y), int(m))


def update_osaka_region_outputs(xlsx_bytes: bytes) -> bool:
    """大阪xlsxから region/27_current.json と region/27_history.json を更新。

    - current は常に上書き（ただし既存より古い月ならスキップ）
    - history は version 単位で追記（既にある月はスキップ＝冪等）
    """
    agg = aggregate_osaka_regions(xlsx_bytes)
    if not agg:
        print("[region] 大阪xlsxの地区集計に失敗")
        return False

    DIR_REGION.mkdir(parents=True, exist_ok=True)
    version = agg["version"]

    # ---- current ----
    cur_path = DIR_REGION / "27_current.json"
    write_current = True
    if cur_path.exists():
        try:
            old = json.loads(cur_path.read_text(encoding="utf-8"))
            if old.get("version") and _vkey(old["version"]) > _vkey(version):
                write_current = False  # 既存の方が新しい
        except Exception:
            pass
    if write_current:
        _atomic_write_text(cur_path, json.dumps(agg, ensure_ascii=False, separators=(",", ":")))

    # ---- history ----
    hist_path = DIR_REGION / "27_history.json"
    if hist_path.exists():
        try:
            hist = json.loads(hist_path.read_text(encoding="utf-8"))
        except Exception:
            hist = {"versions": [], "districts": {}}
    else:
        hist = {"versions": [], "districts": {}}

    if version in hist.get("versions", []):
        print(f"[region] 大阪 {version}: history既存、スキップ")
        return True

    hist.setdefault("versions", []).append(version)
    hist["versions"].sort(key=_vkey)
    dd = hist.setdefault("districts", {})
    for dist in agg["districts"]:
        slot = dd.setdefault(dist["name"], {"totals": {}, "kigo": {}})
        slot["totals"][version] = dist["clinics"]
        for s in dist["standards"]:
            krec = slot["kigo"].setdefault(s["kigo"], {"name": s["name"], "series": []})
            krec["name"] = s["name"]  # 最新名を保持
            krec["series"].append({"v": version, "c": s["count"],
                                   "u": s.get("count_uniq", s["count"])})
            krec["series"].sort(key=lambda p: _vkey(p["v"]))

    _atomic_write_text(hist_path, json.dumps(hist, ensure_ascii=False, separators=(",", ":")))
    print(f"[region] 大阪 {version}: 地区別集計を出力（{len(agg['districts'])}地区、"
          f"その他{len(agg['other_members'])}機関）")
    return True
