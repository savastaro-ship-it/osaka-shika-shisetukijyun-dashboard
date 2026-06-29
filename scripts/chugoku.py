# -*- coding: utf-8 -*-
"""
中国厚生局アダプタ（中国四国厚生局のうち中国5県分）。

掲載ページ: chugokushikoku/chousaka/shisetsukijunjuri.html
府県     : 31鳥取・32島根・33岡山・34広島・35山口
形       : 1ZIPに5県分のxlsxが入っている
注意     : **区分混在のxlsx**（医科・歯科・薬局が同じファイルに入ってる）。
            共通パーサの「区分==歯科」フィルタが歯科だけ拾う。
過去月   : アーカイブなし

ページ構造：
  ## 中国四国厚生局管内の届出受理医療機関名簿（令和X年Y月1日現在）  ← この見出し
  | 鳥取県 | 島根県 | 岡山県 | 広島県 | 山口県 | 各県分エクセルデータ |
  |--------|--------|--------|--------|--------|----------------------|
  | (PDF)  | (PDF)  | (PDF)  | (PDF)  | (PDF)  | (ZIP) ← これを拾う  |
  ## 中国四国厚生局管内の届出受理医療機関名簿（主な届出項目別: ...）  ← ここで打ち切る
  ## 中国四国厚生局管内の保険外併用療養費医療機関名簿（...）  ← または ここ
"""

import io
import re
import zipfile
from typing import List, Tuple
from urllib.parse import urljoin

from lib import Adapter, DiscoveryResult, FileRef, http_get


INDEX_URL = ("https://kouseikyoku.mhlw.go.jp/chugokushikoku/chousaka/"
             "shisetsukijunjuri.html")

# 「中国四国厚生局管内の届出受理医療機関名簿（令和X年Y月1日現在）」見出し
# 「届出受理医療機関名簿」と「（令和...」の間に <br>・改行等が入っても対応するよう緩める。
# ※「（主な届出項目別」はスキップ（セクション違い）
_RE_SECTION = re.compile(
    r'中国四国厚生局管内の届出受理医療機関名簿'
    r'(?!（主な|\(主な)'                       # 「（主な届出項目別」は除外
    r'[^（(]{0,80}?'                           # 間のHTMLタグ・空白等
    r'[（(]\s*令和\s*(\d+)\s*年\s*(\d+)\s*月\s*1\s*日\s*現在\s*[)）]',
    re.DOTALL | re.IGNORECASE,
)
# このセクションの「終わり」を示す次の見出し
_RE_NEXT_SECTION = re.compile(
    r'中国四国厚生局管内の(?:届出受理医療機関名簿（主な|保険外併用療養費|届出受理指定訪問)'
)


class ChugokuAdapter(Adapter):
    bureau = "chugoku"

    def discover(self) -> List[DiscoveryResult]:
        html = http_get(INDEX_URL).decode("utf-8", "replace")

        m = _RE_SECTION.search(html)
        if not m:
            print("[chugoku] 「届出受理医療機関名簿（令和X年Y月1日現在）」見出しが見つかりません")
            return []
        y, mo = int(m.group(1)), int(m.group(2))

        # このセクション範囲を切り出す（次の見出しまで）
        start = m.end()
        m_next = _RE_NEXT_SECTION.search(html, start)
        end = m_next.start() if m_next else len(html)
        section = html[start:end]

        # その中の最初の.zipリンクを取得（この範囲には「各県分エクセルデータ」のZIP1つだけ）
        m_zip = re.search(r'href="([^"]+\.zip)"', section)
        if not m_zip:
            print("[chugoku] セクション内にZIPリンクが見つかりません")
            return []

        full = urljoin(INDEX_URL, m_zip.group(1))
        return [DiscoveryResult(
            bureau=self.bureau,
            file_refs=[FileRef(url=full, filename=full.rsplit("/", 1)[-1])],
            version=f"{2018 + y}.{mo}",
            year=2018 + y,
            month=mo,
            signature=full,
        )]

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        """ZIP内の全xlsxを取り出す。歯科の判別は共通パーサの「区分==歯科」フィルタに任せる。

        中国の名簿xlsxは区分混在（医科・歯科・薬局が同じファイル）。
        """
        out: List[Tuple[str, bytes]] = []
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            for nm in zf.namelist():
                if nm.lower().endswith(".xlsx"):
                    out.append((nm, zf.read(nm)))
        print(f"[chugoku] {ref.filename}: xlsx {len(out)}件抽出")
        return out
