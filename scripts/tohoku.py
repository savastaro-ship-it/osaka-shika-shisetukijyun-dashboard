# -*- coding: utf-8 -*-
"""
東北厚生局アダプタ。

掲載ページ: tohoku/gyomu/gyomu/hoken_kikan/documents/201805koushin.html
府県     : 02青森・03岩手・04宮城・05秋田・06山形・07福島（6県）
形       : **xlsx直リンク**で、6県分が1ファイルにまとまってる
過去月   : アーカイブなし（最新月のみ）

ファイル名規則：
  shisetsu-touhoku-shika-r{令和年2桁}{月2桁}.xlsx
  例: 令和8年4月公表 → shisetsu-touhoku-shika-r0804.xlsx

注：
  ・「shisetsu-touhoku-ika-」（医科）・「shisetsu-touhoku-yakkyoku-」（薬局）と衝突しない
  ・「koumoku{N}-touhoku-shika-」（届出項目別の歯科）は「shisetsu」を含まないので衝突しない
"""

import re
from typing import List, Tuple
from urllib.parse import urljoin

from lib import Adapter, DiscoveryResult, FileRef, http_get


INDEX_URL = ("https://kouseikyoku.mhlw.go.jp/tohoku/gyomu/gyomu/"
             "hoken_kikan/documents/201805koushin.html")

# 月見出し「令和X年Y月1日現在」
_RE_MONTH = re.compile(r'令和\s*(\d+)\s*年\s*(\d+)\s*月\s*1\s*日\s*現在')

# 6県分の歯科xlsx「shisetsu-touhoku-shika-...xlsx」
# - shisetsu-touhoku-shika-r0804.xlsx ← これを拾う
# - shisetsu-touhoku-ika-...xlsx（医科）・shisetsu-touhoku-yakkyoku-...xlsx（薬局）はマッチしない
# - koumoku11-touhoku-shika-...xlsx（届出項目別の歯科）は「shisetsu」を含まないので衝突しない
_RE_SHIKA_XLSX = re.compile(
    r'href="([^"]+shisetsu-touhoku-shika[^"]*\.xlsx)"',
    re.IGNORECASE,
)


class TohokuAdapter(Adapter):
    bureau = "tohoku"

    def discover(self) -> List[DiscoveryResult]:
        """東北は最新月のみ（過去アーカイブなし）。"""
        html = http_get(INDEX_URL).decode("utf-8", "replace")

        m_date = _RE_MONTH.search(html)
        if not m_date:
            print("[tohoku] 月見出しが見つかりません")
            return []
        y, mo = int(m_date.group(1)), int(m_date.group(2))

        m_xlsx = _RE_SHIKA_XLSX.search(html)
        if not m_xlsx:
            print("[tohoku] 「shisetsu-touhoku-shika-...xlsx」リンクが見つかりません")
            return []

        full = urljoin(INDEX_URL, m_xlsx.group(1))
        return [DiscoveryResult(
            bureau=self.bureau,
            file_refs=[FileRef(url=full, filename=full.rsplit("/", 1)[-1])],
            version=f"{2018 + y}.{mo}",
            year=2018 + y,
            month=mo,
            signature=full,
        )]

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        """xlsx直リンクなので、blobをそのまま返す。

        ※ 中身は6県分まとめのxlsx。共通パーサの「都道府県コード」でgroupbyして
          6県（02青森〜07福島）に分離される。
        """
        return [(ref.filename, blob)]
