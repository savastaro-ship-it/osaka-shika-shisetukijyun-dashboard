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
    "00": "全国",  # 特別エントリ：47府県の集計
    "01": "北海道", "02": "青森", "03": "岩手", "04": "宮城", "05": "秋田", "06": "山形",
    "07": "福島", "08": "茨城", "09": "栃木", "10": "群馬", "11": "埼玉", "12": "千葉",
    "13": "東京", "14": "神奈川", "15": "新潟", "16": "富山", "17": "石川", "18": "福井",
    "19": "山梨", "20": "長野", "21": "岐阜", "22": "静岡", "23": "愛知", "24": "三重",
    "25": "滋賀", "26": "京都", "27": "大阪", "28": "兵庫", "29": "奈良", "30": "和歌山",
    "31": "鳥取", "32": "島根", "33": "岡山", "34": "広島", "35": "山口", "36": "徳島",
    "37": "香川", "38": "愛媛", "39": "高知", "40": "福岡", "41": "佐賀", "42": "長崎",
    "43": "熊本", "44": "大分", "45": "宮崎", "46": "鹿児島", "47": "沖縄",
}

# 47府県のコード（"00" を含まない）
ALL_47_CODES: List[str] = [c for c in PREF_NAMES.keys() if c != "00"]

# 表示順：北から順
BUREAU_ORDER = ["hokkaido", "tohoku", "kantoshinetsu", "tokaihokuriku",
                "kinki", "chugoku", "shikoku", "kyushu"]

# 局ごとの管轄
BUREAU_PREFS: Dict[str, List[str]] = {
    "hokkaido":       ["01"],
    "tohoku":         ["02", "03", "04", "05", "06", "07"],
    "kantoshinetsu":  ["08", "09", "10", "11", "12", "13", "14", "15", "19", "20"],
    "tokaihokuriku":  ["16", "17", "21", "22", "23", "24"],
    "kinki":          ["18", "25", "26", "27", "28", "29", "30"],
    "chugoku":        ["31", "32", "33", "34", "35"],
    "shikoku":        ["36", "37", "38", "39"],
    "kyushu":         ["40", "41", "42", "43", "44", "45", "46", "47"],
}
CODE_TO_BUREAU: Dict[str, str] = {
    code: bureau for bureau, codes in BUREAU_PREFS.items() for code in codes
}

# 「特別ポジション」：このリストにあるコードは局並び順より前に出す。
# 「全国」が一番上、その次に「大阪」。
FEATURED_TOP: List[str] = ["00", "27"]


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


def build_national_aggregates() -> bool:
    """47府県の current/*.json と history/*.json から、
    全国集計を data/current/00.json と data/history/00.json に書き出す。

    速報値（current/00.json）：各府県の最新月を単純合算（月混在OK、鮮度優先）
    確定値（history/00.json）：全47府県が共通して持つ月だけ集計（時系列グラフ用）

    47府県全部分のデータが揃ってないと「集計範囲が狭くなる」が、走らせて損は無い。
    戻り値：書き出しが成功したら True。
    """
    # --- 各府県の current/history を読み込み ---
    pref_currents: Dict[str, dict] = {}
    pref_histories: Dict[str, dict] = {}
    for code in ALL_47_CODES:
        cp = DIR_CURRENT / f"{code}.json"
        if cp.exists():
            try:
                pref_currents[code] = json.loads(cp.read_text(encoding="utf-8"))
            except Exception:
                pass
        hp = DIR_HISTORY / f"{code}.json"
        if hp.exists():
            try:
                pref_histories[code] = json.loads(hp.read_text(encoding="utf-8"))
            except Exception:
                pass

    if not pref_currents:
        print("[national] 府県データが1件もないのでスキップ")
        return False

    n_prefs = len(pref_currents)
    print(f"[national] {n_prefs}/47府県のデータから集計")

    # === 速報値（current）：各府県の最新月を単純合算 ===
    total_clinics = sum(c.get("total_clinics", 0) for c in pref_currents.values())
    # kigo別合計
    kigos: Dict[str, dict] = {}
    for c in pref_currents.values():
        for std in c.get("standards", []):
            k = std.get("kigo", "")
            if not k:
                continue
            if k not in kigos:
                kigos[k] = {
                    "kigo": k,
                    "name": std.get("name", k),
                    "count": 0,
                    "count_uniq": 0,
                }
            kigos[k]["count"] += int(std.get("count", 0))
            kigos[k]["count_uniq"] += int(std.get("count_uniq", std.get("count", 0)))

    # version: 府県の最新月のうち最新（max）。混在してたら「（速報）」表記
    versions_set = {c.get("version", "") for c in pref_currents.values() if c.get("version")}
    versions_sorted = sorted(
        versions_set,
        key=lambda v: tuple(int(x) for x in v.split(".")) if v and "." in v else (0, 0),
    )
    if not versions_sorted:
        version_label = ""
        asof_label = ""
    elif len(versions_sorted) == 1:
        version_label = versions_sorted[0]
        # asof は最新月のものを採用
        asof_label = next(
            (c.get("asof", "") for c in pref_currents.values()
             if c.get("version") == version_label),
            "",
        )
    else:
        # 月が混在 → 最新月を version、asofは「速報」表記
        version_label = versions_sorted[-1]
        asof_label = f"各府県の最新月の合算（速報、{versions_sorted[0]}〜{versions_sorted[-1]}）"

    standards_sorted = sorted(
        kigos.values(),
        key=lambda s: (-int(s["count"]), s["kigo"]),
    )

    current_data = {
        "code": "00",
        "name": "全国",
        "bureau": "national",
        "total_clinics": total_clinics,
        "n_standards": len(standards_sorted),
        "version": version_label,
        "asof": asof_label,
        "standards": standards_sorted,
        "n_prefs_aggregated": n_prefs,  # 何府県分から集計したか
    }
    _atomic_write_text(
        DIR_CURRENT / "00.json",
        json.dumps(current_data, ensure_ascii=False, separators=(",", ":")),
    )
    print(f"[national] current: 母数={total_clinics} 種類={len(standards_sorted)} "
          f"version={version_label}")

    # === 確定値（history）：全47府県が共通して持つ月だけ ===
    if not pref_histories:
        print("[national] historyデータが無いので確定値はスキップ")
        return True

    # ALL_47_CODES 全府県が history を持っていない場合、「持ってる府県の共通月」になる
    common_months = None
    for code, h in pref_histories.items():
        months = set(h.get("versions", []))
        if common_months is None:
            common_months = months
        else:
            common_months &= months
    common_months = sorted(
        common_months or [],
        key=lambda v: tuple(int(x) for x in v.split(".")),
    )

    if not common_months:
        print("[national] 全府県共通の月がないのでhistory空")
        # 空のhistoryでも書く
        _atomic_write_text(
            DIR_HISTORY / "00.json",
            json.dumps({"versions": [], "totals": {}, "kigo": {}},
                       ensure_ascii=False, separators=(",", ":")),
        )
        return True

    print(f"[national] history: 全府県共通の月 = {len(common_months)}月分"
          f" ({common_months[0]}〜{common_months[-1]})")

    history_versions = common_months
    history_totals: Dict[str, int] = {v: 0 for v in history_versions}
    history_kigos: Dict[str, dict] = {}  # kigo → {name, names, series_dict{v→{c,u}}}

    for code, h in pref_histories.items():
        totals = h.get("totals", {})
        for v in history_versions:
            history_totals[v] += int(totals.get(v, 0))
        for kigo, rec in (h.get("kigo") or {}).items():
            for sp in rec.get("series", []):
                v = sp.get("v", "")
                if v not in history_versions:
                    continue
                slot = history_kigos.setdefault(kigo, {
                    "name": rec.get("name", kigo),
                    "names": [{"v": history_versions[0], "name": rec.get("name", kigo)}],
                    "_series_dict": {v: {"c": 0, "u": 0} for v in history_versions},
                })
                slot["_series_dict"][v]["c"] += int(sp.get("c", 0))
                slot["_series_dict"][v]["u"] += int(sp.get("u", sp.get("c", 0)))

    # series_dict → series list（versions順）
    for kigo, slot in history_kigos.items():
        slot["series"] = [
            {"v": v, "c": slot["_series_dict"][v]["c"], "u": slot["_series_dict"][v]["u"]}
            for v in history_versions
            if slot["_series_dict"][v]["c"] > 0  # 0件の月は省略
        ]
        slot.pop("_series_dict", None)

    history_data = {
        "versions": history_versions,
        "totals": history_totals,
        "kigo": history_kigos,
    }
    _atomic_write_text(
        DIR_HISTORY / "00.json",
        json.dumps(history_data, ensure_ascii=False, separators=(",", ":")),
    )
    print(f"[national] history書き出し完了（kigo {len(history_kigos)}種類）")
    return True


def rebuild_prefectures_json(now_str: str, state: dict):
    """current/*.json と history/*.json を読んで prefectures.json を組み立て直す"""
    def sort_key(code: str):
        # 1. FEATURED_TOP の順（無ければ大きな値で後ろ送り）
        primary = FEATURED_TOP.index(code) if code in FEATURED_TOP else len(FEATURED_TOP)
        # 2. 局の並び順
        bureau = CODE_TO_BUREAU.get(code, "")
        bi = BUREAU_ORDER.index(bureau) if bureau in BUREAU_ORDER else 99
        # 3. JISコード昇順
        return (primary, bi, code)

    prefs = []
    files = sorted(DIR_CURRENT.glob("*.json"), key=lambda p: sort_key(p.stem))
    latest_asof = ""
    latest_version = ""
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

    # 公表月一覧は history を全部見て union を取る（current は最新だけなので不十分）
    all_versions = set()
    for hp in DIR_HISTORY.glob("*.json"):
        try:
            h = json.loads(hp.read_text(encoding="utf-8"))
        except Exception:
            continue
        all_versions.update(h.get("versions", []))
    versions = sorted(all_versions,
                      key=lambda v: tuple(int(x) for x in v.split(".")))
    if versions:
        latest_version = versions[-1]
    # asof は latest_version を持つ府県のものを使う
    for p in files:
        c = json.loads(p.read_text(encoding="utf-8"))
        if c["version"] == latest_version:
            latest_asof = c["asof"]
            break

    meta = {
        "version": latest_version,
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
    """state.json を読む。旧形式（{"signature":...}）を新形式（{"signatures":[...]}）に自動変換。"""
    if not STATE_FILE.exists():
        return {"bureaus": {}}
    try:
        s = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"bureaus": {}}
    if not isinstance(s, dict) or "bureaus" not in s:
        return {"bureaus": {}}

    # 旧形式 → 新形式に変換
    for bureau, b_state in s["bureaus"].items():
        if not isinstance(b_state, dict):
            continue
        if "signatures" not in b_state and "signature" in b_state:
            b_state["signatures"] = [b_state["signature"]]
            if "version" in b_state and "latest_version" not in b_state:
                b_state["latest_version"] = b_state["version"]
            # 古いキーは残してもOK、無害
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
    """局アダプタの基底。各局はこれを継承して discover / extract_xlsxs を実装する。

    discover() は **利用可能な月のリスト** を返す：
      - 通常の局（過去アーカイブなし、もしくは未対応）は最新1月だけのリスト
      - 過去月backfill対応の局は全月を返す（順不同でOK、update.py側で時系列順に並べる）
    """
    bureau: str = ""

    def discover(self) -> List[DiscoveryResult]:
        raise NotImplementedError

    def fetch(self, ref: FileRef) -> bytes:
        return http_get(ref.url)

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        """生バイト → [(filename, xlsx_bytes), ...]"""
        raise NotImplementedError
