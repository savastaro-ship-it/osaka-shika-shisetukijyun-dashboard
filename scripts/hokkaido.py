# -*- coding: utf-8 -*-
"""
北海道厚生局アダプタ。

掲載ページ: hokkaido/gyomu/gyomu/hoken_kikan/todokede_juri_ichiran.html
構造     : 最新月見出し「【令和X年Y月1日現在】」の下に、医科/歯科/薬局の3行
            それぞれ PDF と Excel(.xlsx) のリンクが並ぶシンプルな表
形       : ZIPではなくxlsx直リンク（九州とは違う）
府県     : 北海道（01）のみ
過去月   : アーカイブなし（最新月のみ）
"""

import re
from typing import List, Tuple
from urllib.parse import urljoin

from lib import Adapter, DiscoveryResult, FileRef, http_get


INDEX_URL = ("https://kouseikyoku.mhlw.go.jp/hokkaido/gyomu/gyomu/"
             "hoken_kikan/todokede_juri_ichiran.html")

# 月見出し「【令和X年Y月1日現在】」（カギカッコは任意、空白は無視）
_RE_MONTH = re.compile(r'令和\s*(\d+)\s*年\s*(\d+)\s*月\s*1\s*日\s*現在')
# 「保険医療機関（歯科）」が出現してから次に来るxlsxリンクを拾う
# 全角・半角カッコ両対応、間に他のタグや改行があってもOK
_RE_SHIKA_XLSX = re.compile(
    r'保険医療機関\s*[（(]\s*歯科\s*[）)].*?href="([^"]+\.xlsx)"',
    re.DOTALL | re.IGNORECASE,
)


class HokkaidoAdapter(Adapter):
    bureau = "hokkaido"

    def discover(self) -> List[DiscoveryResult]:
        """北海道は最新月のみ（過去アーカイブなし）。"""
        html = http_get(INDEX_URL).decode("utf-8", "replace")

        m_date = _RE_MONTH.search(html)
        if not m_date:
            return []
        y, mo = int(m_date.group(1)), int(m_date.group(2))

        m_xlsx = _RE_SHIKA_XLSX.search(html)
        if not m_xlsx:
            print("[hokkaido] 「保険医療機関（歯科）」のxlsxリンクが見つかりません")
            return []

        full = urljoin(INDEX_URL, m_xlsx.group(1))
        return [DiscoveryResult(
            bureau=self.bureau,
            file_refs=[FileRef(url=full, filename=full.rsplit("/", 1)[-1])],
            version=f"{2018 + y}.{mo}",
            year=2018 + y,
            month=mo,
            signature=full,  # 月が変わるとファイルIDも変わる
        )]

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        """xlsx直リンクなので、blobをそのまま返す。"""
        return [(ref.filename, blob)]
