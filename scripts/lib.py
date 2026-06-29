# -*- coding: utf-8 -*-
"""
全国 歯科 施設基準ダッシュボード 共通ライブラリ

- 定数（47都道府県、7厚生局の管轄）
- アダプタ基底クラス（局ごとの取得仕様の差を吸収する）
- 共通xlsxパーサ（受理記号キーで集計、区分=歯科フィルタ、全シート対応）
- 出力（current/history/source、prefectures.json、state.json）

府県キーは JISコード文字列（"01"〜"47"）に統一。
"""

import io
import json
import re
import unicodedata
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import List, Tuple, Dict, Optional

import pandas as pd

# ============ パス・定数 ============

UA = "Mozilla/5.0 (compatible; shika-dashboard/3.0)"
JST = timezone(timedelta(hours=9))

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
DIR_CURRENT = DATA / "current"
DIR_HISTORY = DATA / "history"
DIR_SOURCE = DATA / "source"
STATE_FILE = DATA / "state.json"
PREFS_FILE = DATA / "prefectures.json"

# JISコード "01"〜"47" → 表示名
PREF_NAMES: Dict[str, str] = {
    "01": "北海道", "02": "青森", "03": "岩手", "04": "宮城", "05": "秋田", "06": "山形",
    "07": "福島", "08": "茨城", "09": "栃木", "10": "群馬", "11": "埼玉", "12": "千葉",
    "13": "東京", "14": "神奈川", "15": "新潟", "16": "富山", "17": "石川", "18": "福井",
    "19": "山梨", "20": "長野", "21": "岐阜", "22": "静岡", "23": "愛知", "24": "三重",
    "25": "滋賀", "26": "京都", "27": "大阪", "28": "兵庫", "29": "奈良", "30": "和歌山",
    "31": "鳥取", "32": "島根", "33": "岡山", "34": "広島", "35": "山口", "36": "徳島",
    "37": "香川", "38": "愛媛", "39": "高知", "40": "福岡", "41": "佐賀", "42": "長崎",
    "43": "熊本", "44": "大分", "45": "宮崎", "46": "鹿児島", "47": "沖縄",
}

# 表示順：北から順
BUREAU_ORDER = ["hokkaido", "tohoku", "kantoshinetsu", "tokaihokuriku",
                "kinki", "chugokushikoku", "kyushu"]

# 局ごとの管轄
BUREAU_PREFS: Dict[str, List[str]] = {
    "hokkaido":       ["01"],
    "tohoku":         ["02", "03", "04", "05", "06", "07"],
    "kantoshinetsu":  ["08", "09", "10", "11", "12", "13", "14", "15", "19", "20"],
    "tokaihokuriku":  ["16", "17", "21", "22", "23", "24"],
    "kinki":          ["18", "25", "26", "27", "28", "29", "30"],
    "chugokushikoku": ["31", "32", "33", "34", "35", "36", "37", "38", "39"],
    "kyushu":         ["40", "41", "42", "43", "44", "45", "46", "47"],
}
CODE_TO_BUREAU: Dict[str, str] = {
    code: bureau for bureau, codes in BUREAU_PREFS.items() for code in codes
}

# 「特別ポジション」：このリストにあるコードは局並び順より前に出す。
# 当面は大阪のみ。最終UXで「大阪」「全国」を主軸にする構想の布石。
FEATURED_TOP: List[str] = ["27"]


# ============ データ型 ============

@dataclass
class FileRef:
    """発見した1ファイル分の参照"""
    url: str
    filename: str
    extra: dict = field(default_factory=dict)  # アダプタ固有のメタ


@dataclass
class DiscoveryResult:
    """1局分の発見結果"""
    bureau: str
    file_refs: List[FileRef]
    version: str       # "2026.5"（URLから推定。確定は parse 後のセル内容）
    year: int
    month: int
    signature: str     # 差分検知キー


@dataclass
class PrefRecord:
    """1府県・1月分のパース結果"""
    pref_code: str        # "27"
    pref_name: str        # "大阪"
    asof: str             # "令和8年5月1日現在"
    version: str          # "2026.5"
    total_clinics: int
    standards: List[dict] # [{kigo, name, count, count_uniq}, ...]
    raw_xlsx: bytes


# ============ HTTP ============

def http_get(url: str, timeout: int = 90) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


# ============ 共通パーサ ============

_RE_ASOF = re.compile(r"令和(\d+)年(\d+)月(\d+)日\s*現在")


def _parse_asof_from_text(text) -> Optional[Tuple[int, int, int]]:
    if text is None:
        return None
    s = unicodedata.normalize("NFKC", str(text))
    m = _RE_ASOF.search(s)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2)), int(m.group(3))


def _find_asof_in_sheet(raw: pd.DataFrame) -> Optional[Tuple[int, int, int]]:
    """シート上部のセルから「令和Y年M月D日現在」を探す"""
    for i in range(min(6, len(raw))):
        for v in raw.iloc[i].dropna().tolist():
            res = _parse_asof_from_text(v)
            if res:
                return res
    return None


def _detect_header_row(raw: pd.DataFrame) -> int:
    """「項番」を含む行のindexを返す。見つからなければ3"""
    for i in range(min(12, len(raw))):
        v = raw.iloc[i, 0]
        if pd.notna(v) and str(v).strip() == "項番":
            return i
    return 3


def _aggregate_pref_df(df: pd.DataFrame) -> Tuple[int, List[dict]]:
    """1府県分のDataFrame → (母数, standards)"""
    total = int(df["項番"].dropna().nunique())  # 母数 = 医療機関数

    sub = df.dropna(subset=["受理記号"]).copy()
    # 受理記号は strip のみ（NFKCはhistory継続性を壊すのでやらない）
    sub["受理記号"] = sub["受理記号"].astype(str).str.strip()
    sub = sub[sub["受理記号"] != ""]

    if sub.empty:
        return total, []

    g = (sub.groupby("受理記号")
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
        for _, r in g.iterrows()
    ]
    return total, standards


def parse_xlsx_to_records(xlsx_bytes: bytes) -> List[PrefRecord]:
    """
    1つのxlsxから府県レコードのリストを返す。
    - 全シートをループ（東北=6シート, 他=1シート）
    - 「項番」を含む行をヘッダ検出
    - 区分=="歯科" でフィルタ（中国四国対応・他は無害）
    - 都道府県コードでgroupby（通常1府県）
    - 母数 = 項番のユニーク数、施設基準 = 受理記号でgroupby
    """
    xf = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    records: List[PrefRecord] = []

    for sheet in xf.sheet_names:
        raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet,
                            header=None, dtype=str)
        asof_tuple = _find_asof_in_sheet(raw)
        if not asof_tuple:
            continue
        y, m, d = asof_tuple
        version = f"{2018 + y}.{m}"
        asof_str = f"令和{y}年{m}月{d}日現在"

        hdr = _detect_header_row(raw)
        df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet,
                           header=hdr, dtype=str)

        if "項番" not in df.columns or "受理記号" not in df.columns:
            continue

        # 区分=="歯科"フィルタ
        if "区分" in df.columns:
            df = df[df["区分"] == "歯科"]

        if df.empty:
            continue

        # 都道府県コードでgroupby
        if "都道府県コード" in df.columns:
            groups = list(df.dropna(subset=["都道府県コード"])
                            .groupby("都道府県コード"))
        else:
            groups = [(None, df)]

        for code_val, g in groups:
            code = str(code_val).strip().zfill(2) if code_val else ""
            if code not in PREF_NAMES:
                continue
            total, standards = _aggregate_pref_df(g)
            if total == 0 and not standards:
                continue
            records.append(PrefRecord(
                pref_code=code,
                pref_name=PREF_NAMES[code],
                asof=asof_str,
                version=version,
                total_clinics=total,
                standards=standards,
                raw_xlsx=xlsx_bytes,
            ))
    return records


def merge_records_by_pref(records: List[PrefRecord]) -> List[PrefRecord]:
    """
    同じ（pref_code, version）のレコードが複数あればマージして1つにする。
    想定ケース：1府県が複数xlsxに分かれている場合（ファイル分割・区分別出力など）。
    通常は何もしない（pass-through）。
    """
    if not records:
        return []
    grouped: Dict[Tuple[str, str], PrefRecord] = {}
    for rec in records:
        key = (rec.pref_code, rec.version)
        if key not in grouped:
            grouped[key] = rec
        else:
            grouped[key] = _merge_two(grouped[key], rec)
    return list(grouped.values())


def _merge_two(a: PrefRecord, b: PrefRecord) -> PrefRecord:
    """同じpref/versionの2レコードを足し合わせる。受理記号で集計する単純合算。"""
    # 項番（医療機関数）は重複なしと仮定して合算
    new_total = a.total_clinics + b.total_clinics
    # 受理記号ごとに合算
    by_kigo: Dict[str, dict] = {s["kigo"]: dict(s) for s in a.standards}
    for s in b.standards:
        if s["kigo"] in by_kigo:
            by_kigo[s["kigo"]]["count"] += s["count"]
            by_kigo[s["kigo"]]["count_uniq"] += s["count_uniq"]
            # 名前は先に来た方を採用（同月内なので一致する想定）
        else:
            by_kigo[s["kigo"]] = dict(s)
    new_standards = sorted(by_kigo.values(), key=lambda s: -s["count"])
    return PrefRecord(
        pref_code=a.pref_code,
        pref_name=a.pref_name,
        asof=a.asof,
        version=a.version,
        total_clinics=new_total,
        standards=new_standards,
        raw_xlsx=a.raw_xlsx,  # source xlsx は先頭のものだけ保持
    )


# ============ 出力 ============

def _atomic_write_text(path: Path, text: str):
    """書き換え途中で読まれて壊れないように一時ファイル経由で書く"""
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(path)


def write_pref_outputs(rec: PrefRecord, bureau: str):
    """1府県分の current/history/source を書き出す"""
    # source xlsx
    (DIR_SOURCE / f"{rec.pref_code}.xlsx").write_bytes(rec.raw_xlsx)

    # current
    cur = {
        "code": rec.pref_code,
        "name": rec.pref_name,
        "bureau": bureau,
        "version": rec.version,
        "asof": rec.asof,
        "total_clinics": rec.total_clinics,
        "n_standards": len(rec.standards),
        "standards": rec.standards,
    }
    _atomic_write_text(
        DIR_CURRENT / f"{rec.pref_code}.json",
        json.dumps(cur, ensure_ascii=False, separators=(",", ":")),
    )

    # history（追記）
    _update_history(rec)


def _update_history(rec: PrefRecord):
    path = DIR_HISTORY / f"{rec.pref_code}.json"
    hist = (json.loads(path.read_text(encoding="utf-8"))
            if path.exists() else
            {"versions": [], "totals": {}, "kigo": {}})
    hist.setdefault("totals", {})
    if rec.version not in hist["versions"]:
        hist["versions"].append(rec.version)
    hist["totals"][rec.version] = rec.total_clinics

    for s in rec.standards:
        k = s["kigo"]
        slot = hist["kigo"].get(k)
        if slot is None:
            slot = {"name": s["name"],
                    "names": [{"v": rec.version, "name": s["name"]}],
                    "series": []}
            hist["kigo"][k] = slot
        elif s["name"] and s["name"] != slot["name"]:
            slot["names"].append({"v": rec.version, "name": s["name"]})
            slot["name"] = s["name"]
        point = {"v": rec.version, "c": s["count"], "u": s["count_uniq"]}
        if slot["series"] and slot["series"][-1]["v"] == rec.version:
            slot["series"][-1] = point  # 同版の訂正は上書き
        else:
            slot["series"].append(point)

    _atomic_write_text(
        path,
        json.dumps(hist, ensure_ascii=False, separators=(",", ":")),
    )


def rebuild_prefectures_json(now_str: str, state: dict):
    """current/*.json を全部読んで prefectures.json を組み立て直す"""
    def sort_key(code: str):
        # 1. FEATURED_TOP の順（無ければ大きな値で後ろ送り）
        primary = FEATURED_TOP.index(code) if code in FEATURED_TOP else len(FEATURED_TOP)
        # 2. 局の並び順
        bureau = CODE_TO_BUREAU.get(code, "")
        bi = BUREAU_ORDER.index(bureau) if bureau in BUREAU_ORDER else 99
        # 3. JISコード昇順
        return (primary, bi, code)

    prefs = []
    versions_set = set()
    files = sorted(DIR_CURRENT.glob("*.json"), key=lambda p: sort_key(p.stem))
    latest_asof = ""
    for p in files:
        c = json.loads(p.read_text(encoding="utf-8"))
        prefs.append({
            "code": c["code"],
            "name": c["name"],
            "bureau": c.get("bureau", ""),
            "total_clinics": c["total_clinics"],
            "n_standards": c["n_standards"],
            "version": c["version"],
            "asof": c["asof"],
        })
        versions_set.add(c["version"])

    versions = sorted(versions_set,
                      key=lambda v: tuple(int(x) for x in v.split(".")))
    latest = versions[-1] if versions else ""
    for p in files:
        c = json.loads(p.read_text(encoding="utf-8"))
        if c["version"] == latest:
            latest_asof = c["asof"]
            break

    meta = {
        "version": latest,
        "asof": latest_asof,
        "versions": versions,
        "prefectures": prefs,
        "bureaus": state.get("bureaus", {}),
        "checked": now_str,
    }
    _atomic_write_text(
        PREFS_FILE,
        json.dumps(meta, ensure_ascii=False, separators=(",", ":")),
    )


# ============ state ============

def load_state() -> dict:
    """state.json を読む。旧形式（{"signature":...}）は捨てて新形式の空っぽを返す。"""
    if not STATE_FILE.exists():
        return {"bureaus": {}}
    try:
        s = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"bureaus": {}}
    if not isinstance(s, dict) or "bureaus" not in s:
        return {"bureaus": {}}
    return s


def save_state(state: dict, now_str: str):
    state["last_checked"] = now_str
    _atomic_write_text(
        STATE_FILE,
        json.dumps(state, ensure_ascii=False, indent=2),
    )


def now_jst_str() -> str:
    return datetime.now(JST).strftime("%Y/%m/%d %H:%M")


def ensure_dirs():
    for d in (DATA, DIR_CURRENT, DIR_HISTORY, DIR_SOURCE):
        d.mkdir(parents=True, exist_ok=True)


# ============ アダプタ基底 ============

class Adapter:
    """局アダプタの基底。各局はこれを継承して discover / extract_xlsxs を実装する。"""
    bureau: str = ""

    def discover(self) -> Optional[DiscoveryResult]:
        raise NotImplementedError

    def fetch(self, ref: FileRef) -> bytes:
        return http_get(ref.url)

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        """生バイト → [(filename, xlsx_bytes), ...]"""
        raise NotImplementedError
