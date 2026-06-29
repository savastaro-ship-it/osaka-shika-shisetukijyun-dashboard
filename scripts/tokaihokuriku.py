# -*- coding: utf-8 -*-
"""
東海北陸厚生局アダプタ。

掲載ページ: tokaihokuriku/newpage_00349.html
構造     : 「届出受理医療機関名簿（歯科）」リンクが1つあり、それが歯科専用ZIP
形       : 1ZIPに6県分（富山・石川・岐阜・静岡・愛知・三重）のxlsxが入っている
府県     : 16富山・17石川・21岐阜・22静岡・23愛知・24三重
過去月   : アーカイブなし（最新月のみ）

※「（富山歯科）」のようにPDFリンクには県名が入る。ZIPリンクは「（歯科）」のみ
  なので、`[（(]歯科[）)]`（県名なし）で絞り込めば誤マッチしない。
"""

import io
import re
import zipfile
from typing import List, Tuple
from urllib.parse import urljoin

from lib import Adapter, DiscoveryResult, FileRef, http_get


INDEX_URL = "https://kouseikyoku.mhlw.go.jp/tokaihokuriku/newpage_00349.html"

# 月見出し「令和X年Y月1日現在」
_RE_MONTH = re.compile(r'令和\s*(\d+)\s*年\s*(\d+)\s*月\s*1\s*日\s*現在')

# 「届出受理医療機関名簿（歯科）」（県名なしの単独「歯科」）のZIPリンク。
# HTML では <a href="...zip">届出受理医療機関名簿（歯科）...</a> という構造で、
# href がテキストより先に来るので、その順序でマッチさせる。
# 「（富山歯科）」「（石川歯科）」などのPDFリンクは「県名+歯科」なのでマッチしない。
_RE_SHIKA_ZIP = re.compile(
    r'href="([^"]+\.zip)"[^>]*>\s*届出受理医療機関名簿\s*[（(]歯科[）)]',
    re.DOTALL | re.IGNORECASE,
)


class TokaihokurikuAdapter(Adapter):
    bureau = "tokaihokuriku"

    def discover(self) -> List[DiscoveryResult]:
        """東海北陸は最新月のみ（過去アーカイブなし）。"""
        html = http_get(INDEX_URL).decode("utf-8", "replace")

        m_date = _RE_MONTH.search(html)
        if not m_date:
            return []
        y, mo = int(m_date.group(1)), int(m_date.group(2))

        m_zip = _RE_SHIKA_ZIP.search(html)
        if not m_zip:
            print("[tokaihokuriku] 「届出受理医療機関名簿（歯科）」のZIPリンクが見つかりません")
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
        """ZIP内の全xlsxを取り出す。府県の判別は共通パーサ側の 都道府県コード で。

        東海北陸の歯科ZIPには6県分のxlsxが入っている想定。
        """
        out: List[Tuple[str, bytes]] = []
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            for nm in zf.namelist():
                if nm.lower().endswith(".xlsx"):
                    out.append((nm, zf.read(nm)))
        print(f"[tokaihokuriku] {ref.filename}: xlsx {len(out)}件抽出")
        return out
