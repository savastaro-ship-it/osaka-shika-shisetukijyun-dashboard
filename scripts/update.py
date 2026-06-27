#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
近畿厚生局「施設基準の届出受理状況」(歯科) を毎日チェックして、
新しい月のデータが公開されていれば大阪府分を集計し、
ダッシュボード(index.html)とdata/を更新するスクリプト。

やること:
  1. 一覧ページをスクレイプして「歯科」のZIPリンクを取得
       例: https://kouseikyoku.mhlw.go.jp/kinki/2026.4_sisetukijun_sika.zip
       （訂正版は s2026.4_..._sika.zip のように s が付くことがある）
  2. 前回処理したファイル名(signature)と比較し、変わっていれば更新
  3. ZIPを取得 → 中の {年月}_sisetukijun_osaka_sika.xlsx を抽出
  4. 集計（母数=項番のユニーク数 / 各施設基準=受理届出名称×受理記号の届出医療機関数）
  5. index.html の /*DATA:START*/ ... /*DATA:END*/ を置換、data/source.xlsx を更新
  6. 最終チェック日時は毎回更新（＝毎日コミットが入り、cron無効化(60日)も防げる）

依存: pandas, openpyxl （標準ライブラリ: urllib, zipfile, json, re, datetime）
"""

import io
import json
import re
import sys
import zipfile
import urllib.request
from datetime import datetime, timezone, timedelta
from pathlib import Path

import pandas as pd

INDEX_URL = "https://kouseikyoku.mhlw.go.jp/kinki/gyomu/gyomu/hoken_kikan/shitei_jokyo_00004.html"
BASE = "https://kouseikyoku.mhlw.go.jp/kinki/"
UA = "Mozilla/5.0 (compatible; osaka-shika-dashboard/1.0; +https://github.com/)"

ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML = ROOT / "index.html"
DATA_DIR = ROOT / "data"
SOURCE_XLSX = DATA_DIR / "source.xlsx"
STATE_JSON = DATA_DIR / "state.json"

JST = timezone(timedelta(hours=9))


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def find_sika_zip(html: str):
    """一覧ページから歯科ZIPのURLとファイル名・版(年月)を取り出す。"""
    # 例: 2026.4_sisetukijun_sika.zip / s2026.4_sisetukijun_sika.zip
    m = re.search(r'href="([^"]*?((?:s)?(\d{4})\.(\d{1,2})_sisetukijun_sika\.zip))"', html)
    if not m:
        return None
    href, fname, year, month = m.group(1), m.group(2), int(m.group(3)), int(m.group(4))
    url = href if href.startswith("http") else (BASE + href.lstrip("/").split("/")[-1])
    version = f"{year}.{month}"           # 表示用の版 (例 "2026.4")
    return {"url": url, "filename": fname, "version": version, "year": year, "month": month}


def reiwa_asof(year: int, month: int) -> str:
    return f"令和{year - 2018}年{month}月1日現在"


def detect_header_row(raw: pd.DataFrame) -> int:
    for i in range(min(10, len(raw))):
        if str(raw.iloc[i, 0]).strip() == "項番":
            return i
    return 3  # フォールバック（観測値）


def aggregate(xlsx_bytes: bytes) -> dict:
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), header=None, dtype=str)
    hdr = detect_header_row(raw)
    df = pd.read_excel(io.BytesIO(xlsx_bytes), header=hdr, dtype=str)

    total = df["項番"].dropna().nunique()           # 母数 = 歯科医療機関数
    sub = df.dropna(subset=["受理届出名称"]).copy()
    g = (sub.groupby(["受理届出名称", "受理記号"])["項番"]
            .nunique().reset_index())
    g.columns = ["name", "abbr", "count"]
    g = g.sort_values("count", ascending=False)

    standards = [
        {"name": r["name"],
         "abbr": (r["abbr"] if pd.notna(r["abbr"]) else ""),
         "count": int(r["count"])}
        for _, r in g.iterrows()
    ]
    return {"total_clinics": int(total),
            "n_standards": int(len(standards)),
            "standards": standards}


def load_existing_data() -> dict:
    html = INDEX_HTML.read_text(encoding="utf-8")
    m = re.search(r"/\*DATA:START\*/(.*?)/\*DATA:END\*/", html, re.S)
    return json.loads(m.group(1)) if m else {}


def inject_data(data: dict):
    html = INDEX_HTML.read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = re.sub(r"/\*DATA:START\*/.*?/\*DATA:END\*/",
                  "/*DATA:START*/" + payload + "/*DATA:END*/",
                  html, flags=re.S)
    INDEX_HTML.write_text(html, encoding="utf-8")


def main():
    DATA_DIR.mkdir(exist_ok=True)
    now = datetime.now(JST).strftime("%Y/%m/%d %H:%M")
    state = json.loads(STATE_JSON.read_text(encoding="utf-8")) if STATE_JSON.exists() else {}

    print(f"[{now}] checking {INDEX_URL}")
    html = fetch(INDEX_URL).decode("utf-8", "replace")
    info = find_sika_zip(html)
    if not info:
        print("歯科ZIPリンクが見つかりませんでした。ページ構造が変わった可能性があります。")
        sys.exit(2)
    print("found:", info["filename"], "->", info["url"])

    data = load_existing_data()
    data_changed = info["filename"] != state.get("signature")

    if data_changed:
        print("新しい（または訂正された）データを検出。取得して集計します…")
        zbytes = fetch(info["url"])
        with zipfile.ZipFile(io.BytesIO(zbytes)) as zf:
            member = next((n for n in zf.namelist()
                           if n.endswith("_sisetukijun_osaka_sika.xlsx")), None)
            if not member:
                print("ZIP内に大阪歯科xlsxが見つかりません:", zf.namelist())
                sys.exit(3)
            xlsx_bytes = zf.read(member)

        agg = aggregate(xlsx_bytes)
        SOURCE_XLSX.write_bytes(xlsx_bytes)   # 元データを差し替え（ダウンロード同梱用）

        today = datetime.now(JST)
        created = f"令和{today.year - 2018}年{today.month}月{today.day}日作成"
        data.update({
            "asof": reiwa_asof(info["year"], info["month"]),
            "version": info["version"],
            "created": created,
            "source_url": "./data/source.xlsx",
            "total_clinics": agg["total_clinics"],
            "n_standards": agg["n_standards"],
            "standards": agg["standards"],
        })
        print(f"集計完了: 母数={agg['total_clinics']} / 基準数={agg['n_standards']}")
    else:
        print("前回から変化なし。最終チェック日時のみ更新します。")

    data["checked"] = now
    inject_data(data)

    state = {"signature": info["filename"], "version": info["version"], "last_checked": now}
    STATE_JSON.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    # GitHub Actions 出力（changed=true のときだけ意味のある更新があった）
    out = __import__("os").environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a", encoding="utf-8") as f:
            f.write(f"changed={'true' if data_changed else 'false'}\n")
            f.write(f"version={info['version']}\n")

    print("done. data_changed =", data_changed)


if __name__ == "__main__":
    main()
