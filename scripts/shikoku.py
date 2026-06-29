# -*- coding: utf-8 -*-
"""
四国厚生支局アダプタ。

掲載ページ: shikoku/gyomu/gyomu/hoken_kikan/shitei/index.html
府県     : 36徳島・37香川・38愛媛・39高知
形       : 歯科専用ZIPに4県分のxlsxが入っている（医科とは別ZIP）
過去月   : アーカイブなし

ページ構造（4.施設基準の届出受理状況（全体）セクション）:
  ## 4. 施設基準の届出受理状況（全体）
  | 届出受理医療機関名簿（令和X年Y月1日現在） | 香川県 | 徳島県 | 愛媛県 | 高知県 | 各県分エクセルデータ |
  | 医　科 | PDF | PDF | PDF | PDF | (ZIP・医科)              |
  | 歯　科 | PDF | PDF | PDF | PDF | **(ZIP・歯科) ← これ**  |
  | 薬　局 | PDF | PDF | PDF | PDF | (ZIP・薬局)              |

注：「届出受理医療機関名簿（届出項目別）（令和X年Y月1日現在）」というセクション7
   も存在するが、こちらは途中に「（届出項目別）」が入るので正規表現でマッチしない。
"""

import io
import re
import zipfile
from typing import List, Tuple
from urllib.parse import urljoin

from lib import Adapter, DiscoveryResult, FileRef, http_get


INDEX_URL = ("https://kouseikyoku.mhlw.go.jp/shikoku/gyomu/gyomu/"
             "hoken_kikan/shitei/index.html")

# 「届出受理医療機関名簿（令和X年Y月1日現在）」見出し
# ※「届出受理医療機関名簿（届出項目別）（令和...）」は途中で構造が違うのでマッチしない
_RE_HEADER = re.compile(
    r'届出受理医療機関名簿\s*[（(]令和\s*(\d+)\s*年\s*(\d+)\s*月\s*1\s*日\s*現在\s*[)）]'
)

# テーブル中の「歯　科」（全角空白あり/なし両対応）行の最初の.zipリンク
_RE_SHIKA_ZIP = re.compile(
    r'歯\s*科.*?href="([^"]+\.zip)"',
    re.DOTALL | re.IGNORECASE,
)


class ShikokuAdapter(Adapter):
    bureau = "shikoku"

    def discover(self) -> List[DiscoveryResult]:
        html = http_get(INDEX_URL).decode("utf-8", "replace")

        m = _RE_HEADER.search(html)
        if not m:
            print("[shikoku] 「届出受理医療機関名簿（令和X年Y月1日現在）」見出しが見つかりません")
            return []
        y, mo = int(m.group(1)), int(m.group(2))

        # このセクション範囲を切り出す（次の節見出しまで）
        # ##### 5.訪問看護事業所... や、別の届出受理医療機関名簿セクションまで
        start = m.end()
        m_next = re.search(r'#{4,5}\s+\*?\d+[\.．]|<h[2-4]|届出受理医療機関名簿（届出項目別）',
                           html[start:])
        end = start + m_next.start() if m_next else len(html)
        section = html[start:end]

        m_zip = _RE_SHIKA_ZIP.search(section)
        if not m_zip:
            print("[shikoku] 歯科行のZIPリンクが見つかりません")
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
        """ZIP内の全xlsxを取り出す。歯科専用ZIPなので中身は歯科データのみ。"""
        out: List[Tuple[str, bytes]] = []
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            for nm in zf.namelist():
                if nm.lower().endswith(".xlsx"):
                    out.append((nm, zf.read(nm)))
        print(f"[shikoku] {ref.filename}: xlsx {len(out)}件抽出")
        return out
