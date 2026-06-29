# -*- coding: utf-8 -*-
"""
近畿厚生局アダプタ。

掲載ページ: kinki/gyomu/gyomu/hoken_kikan/shitei_jokyo_00004.html
発見方法 : 静的HTMLを正規表現 `(s?)\\d{4}\\.\\d{1,2}_sisetukijun_sika\\.zip` で探す
形      : 1つのZIPに府県別xlsxが同梱（`_sisetukijun_{romaji}_sika.xlsx`）
"""

import io
import re
import zipfile
from typing import List, Tuple, Optional

from lib import Adapter, DiscoveryResult, FileRef, http_get


INDEX_URL = ("https://kouseikyoku.mhlw.go.jp/kinki/gyomu/gyomu/"
             "hoken_kikan/shitei_jokyo_00004.html")
BASE = "https://kouseikyoku.mhlw.go.jp/kinki/"

# ZIP内の府県別xlsxを拾う
_RE_PREF_XLSX = re.compile(r"_sisetukijun_([a-z]+)_sika\.xlsx$")

# 訂正版は先頭に `s` が付く（例: s2026.5_sisetukijun_sika.zip）
_RE_ZIP_HREF = re.compile(
    r'href="([^"]*?((?:s)?(\d{4})\.(\d{1,2})_sisetukijun_sika\.zip))"'
)


class KinkiAdapter(Adapter):
    bureau = "kinki"

    def discover(self) -> Optional[DiscoveryResult]:
        html = http_get(INDEX_URL).decode("utf-8", "replace")
        m = _RE_ZIP_HREF.search(html)
        if not m:
            return None
        href, fname = m.group(1), m.group(2)
        year, month = int(m.group(3)), int(m.group(4))
        url = href if href.startswith("http") else (BASE + href.lstrip("/").split("/")[-1])
        return DiscoveryResult(
            bureau=self.bureau,
            file_refs=[FileRef(url=url, filename=fname)],
            version=f"{year}.{month}",
            year=year,
            month=month,
            signature=fname,
        )

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        out: List[Tuple[str, bytes]] = []
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            for nm in zf.namelist():
                if _RE_PREF_XLSX.search(nm):
                    out.append((nm, zf.read(nm)))
        return out
