#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
近畿厚生局「施設基準の届出受理状況（歯科）」を毎日チェックし、新しい月が出たら
近畿・北陸の全府県を集計してダッシュボード用データを更新するスクリプト。

設計の要点
- 1つの歯科ZIPに全府県の xlsx が同梱されているので、府県ごとにループして集計する。
- 施設基準の主キーは「受理記号（略称）」。名称が変わっても記号が同じなら同一基準として
  時系列を連続させる。表示名はその月の名称（記号↔名称は月内で1:1）。
- 「施設基準の種類」は、その府県で届出のある記号だけ（=名簿に現れた記号）を数える。
- 月次スナップショットを history/<府県>.json に追記し、推移（時系列）を育てる。
  名称が変わった月は names[] に記録する。掲載の無い月は単に点が増えないだけ（=欠測）。

出力（data/ 配下）
  current/<府県>.json     最新スナップショット（一覧表示用）
  history/<府県>.json     記号キーの時系列＋名称履歴（推移用）
  source/<府県>.xlsx      元データ（ダウンロード同梱用、最新で差し替え）
  prefectures.json        メタ（版・現在日・最終チェック・府県一覧・公表月一覧）
  state.json              前回処理したファイル名など

依存: pandas, openpyxl
"""

import io
import json
import os
import re
import sys
import zipfile
import urllib.request
from datetime import datetime, timezone, timedelta
from pathlib import Path

import pandas as pd

INDEX_URL = "https://kouseikyoku.mhlw.go.jp/kinki/gyomu/gyomu/hoken_kikan/shitei_jokyo_00004.html"
BASE = "https://kouseikyoku.mhlw.go.jp/kinki/"
UA = "Mozilla/5.0 (compatible; kinki-shika-dashboard/2.0)"
JST = timezone(timedelta(hours=9))

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
DIR_CURRENT = DATA / "current"
DIR_HISTORY = DATA / "history"
DIR_SOURCE = DATA / "source"
STATE = DATA / "state.json"
PREFS_JSON = DATA / "prefectures.json"

# romaji（ファイル名） -> 日本語表示名（近畿厚生局＝近畿2府5県＋北陸 富山・石川）
PREF_NAMES = {
    "fukui": "福井", "toyama": "富山", "ishikawa": "石川", "shiga": "滋賀",
    "kyoto": "京都", "osaka": "大阪", "hyogo": "兵庫", "nara": "奈良", "wakayama": "和歌山",
}
PREF_ORDER = ["osaka", "hyogo", "kyoto", "shiga", "nara", "wakayama", "fukui", "ishikawa", "toyama"]


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=90) as r:
        return r.read()


def find_sika_zip(html: str):
    """一覧ページから歯科ZIPのURL・ファイル名・版（年月）を取り出す。"""
    m = re.search(r'href="([^"]*?((?:s)?(\d{4})\.(\d{1,2})_sisetukijun_sika\.zip))"', html)
    if not m:
        return None
    href, fname, year, month = m.group(1), m.group(2), int(m.group(3)), int(m.group(4))
    url = href if href.startswith("http") else (BASE + href.lstrip("/").split("/")[-1])
    return {"url": url, "filename": fname, "version": f"{year}.{month}", "year": year, "month": month}


def reiwa_asof(year: int, month: int) -> str:
    return f"令和{year - 2018}年{month}月1日現在"


def detect_header_row(raw: pd.DataFrame) -> int:
    for i in range(min(10, len(raw))):
        if str(raw.iloc[i, 0]).strip() == "項番":
            return i
    return 3


def aggregate_pref(xlsx_bytes: bytes) -> dict:
    """1府県分を受理記号キーで集計。"""
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), header=None, dtype=str)
    hdr = detect_header_row(raw)
    df = pd.read_excel(io.BytesIO(xlsx_bytes), header=hdr, dtype=str)

    total = int(df["項番"].dropna().nunique())  # 母数＝歯科医療機関数
    sub = df.dropna(subset=["受理記号"]).copy()
    g = (sub.groupby("受理記号")
            .agg(name=("受理届出名称", "first"),
                 count=("項番", "size"),
                 count_uniq=("項番", "nunique"))
            .reset_index()
            .rename(columns={"受理記号": "kigo"}))
    g = g.sort_values("count", ascending=False)
    standards = [
        {"kigo": r["kigo"],
         "name": (r["name"] if pd.notna(r["name"]) else ""),
         "count": int(r["count"]),
         "count_uniq": int(r["count_uniq"])}
        for _, r in g.iterrows()
    ]
    return {"total_clinics": total, "n_standards": len(standards), "standards": standards}


def update_history(pref: str, version: str, total: int, standards: list):
    """history/<pref>.json に当月分を追記。記号キーで時系列＋名称履歴を持つ。"""
    path = DIR_HISTORY / f"{pref}.json"
    hist = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {"versions": [], "totals": {}, "kigo": {}}
    hist.setdefault("totals", {})
    if version not in hist["versions"]:
        hist["versions"].append(version)
    hist["totals"][version] = total      # 月ごとの母数（届出率の分母）
    for s in standards:
        k = s["kigo"]
        rec = hist["kigo"].get(k)
        if rec is None:
            rec = {"name": s["name"], "names": [{"v": version, "name": s["name"]}], "series": []}
            hist["kigo"][k] = rec
        elif s["name"] and s["name"] != rec["name"]:
            rec["names"].append({"v": version, "name": s["name"]})   # 名称変更を記録
            rec["name"] = s["name"]
        point = {"v": version, "c": s["count"], "u": s["count_uniq"]}
        if rec["series"] and rec["series"][-1]["v"] == version:
            rec["series"][-1] = point         # 同版の再公表（訂正）は上書き
        else:
            rec["series"].append(point)
    path.write_text(json.dumps(hist, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def main():
    for d in (DATA, DIR_CURRENT, DIR_HISTORY, DIR_SOURCE):
        d.mkdir(exist_ok=True)
    now = datetime.now(JST).strftime("%Y/%m/%d %H:%M")
    state = json.loads(STATE.read_text(encoding="utf-8")) if STATE.exists() else {}

    print(f"[{now}] checking {INDEX_URL}")
    html = fetch(INDEX_URL).decode("utf-8", "replace")
    info = find_sika_zip(html)
    if not info:
        print("歯科ZIPリンクが見つかりません。ページ構造が変わった可能性があります。")
        sys.exit(2)
    print("found:", info["filename"], "->", info["url"])

    data_changed = info["filename"] != state.get("signature")
    prefs_meta = json.loads(PREFS_JSON.read_text(encoding="utf-8")) if PREFS_JSON.exists() else {}

    if data_changed:
        print("新規/訂正データを検出。全府県を集計します…")
        zbytes = fetch(info["url"])
        found = {}
        with zipfile.ZipFile(io.BytesIO(zbytes)) as zf:
            for nm in zf.namelist():
                m = re.search(r"_sisetukijun_([a-z]+)_sika\.xlsx$", nm)
                if m:
                    found[m.group(1)] = zf.read(nm)
        if not found:
            print("ZIP内に府県別xlsxが見つかりません")
            sys.exit(3)

        per_pref = []
        for code in sorted(found, key=lambda c: PREF_ORDER.index(c) if c in PREF_ORDER else 99):
            agg = aggregate_pref(found[code])
            jp = PREF_NAMES.get(code, code)
            (DIR_SOURCE / f"{code}.xlsx").write_bytes(found[code])
            (DIR_CURRENT / f"{code}.json").write_text(json.dumps({
                "code": code, "name": jp, "version": info["version"],
                "asof": reiwa_asof(info["year"], info["month"]),
                "total_clinics": agg["total_clinics"],
                "n_standards": agg["n_standards"],
                "standards": agg["standards"],
            }, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
            update_history(code, info["version"], agg["total_clinics"], agg["standards"])
            per_pref.append({"code": code, "name": jp,
                             "total_clinics": agg["total_clinics"], "n_standards": agg["n_standards"]})
            print(f"  {jp}: 母数={agg['total_clinics']} / 種類={agg['n_standards']}")

        versions = prefs_meta.get("versions", [])
        if info["version"] not in versions:
            versions.append(info["version"])
        prefs_meta = {
            "version": info["version"],
            "asof": reiwa_asof(info["year"], info["month"]),
            "versions": versions,
            "prefectures": per_pref,
        }
    else:
        print("前回から変化なし。最終チェック日時のみ更新します。")

    prefs_meta["checked"] = now
    PREFS_JSON.write_text(json.dumps(prefs_meta, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    STATE.write_text(json.dumps(
        {"signature": info["filename"], "version": info["version"], "last_checked": now},
        ensure_ascii=False, indent=2), encoding="utf-8")

    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a", encoding="utf-8") as f:
            f.write(f"changed={'true' if data_changed else 'false'}\n")
            f.write(f"version={info['version']}\n")
    print("done. data_changed =", data_changed)


if __name__ == "__main__":
    main()
