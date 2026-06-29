# -*- coding: utf-8 -*-
"""
九州厚生局アダプタ。

掲載ページ: kyushu/gyomu/gyomu/hoken_kikan/index_00007.html
構造     : 月ごとに「令和X年Y月1日現在」見出し → 8府県分のZIPリンクが続く
形       : 1府県=1ZIP。各ZIP内に医科・歯科・薬局のxlsxが入っており、
            ファイル名に "shika" を含むものが歯科。
府県順   : 福岡→佐賀→長崎→熊本→大分→宮崎→鹿児島→沖縄（公式掲載順）

差分検知 : 全URLの結合をsignatureとする（月単位の更新で全URLが変わる想定）
"""

import io
import re
import zipfile
from typing import List, Tuple, Optional

from lib import Adapter, DiscoveryResult, FileRef, http_get


INDEX_URL = ("https://kouseikyoku.mhlw.go.jp/kyushu/gyomu/gyomu/"
             "hoken_kikan/index_00007.html")
BASE = "https://kouseikyoku.mhlw.go.jp/kyushu/"

# 掲載順（公式注記による）。9府県目以降が出ることはない。
_PREF_ORDER = ["40", "41", "42", "43", "44", "45", "46", "47"]
# 40:福岡 41:佐賀 42:長崎 43:熊本 44:大分 45:宮崎 46:鹿児島 47:沖縄

# 月見出しパターン：「令和Y年M月D日現在」（全角/半角混在に NFKC は不要、re で吸収）
_RE_MONTH_HEADER = re.compile(
    r'令和\s*(\d+)\s*年\s*(\d+)\s*月\s*\d+\s*日\s*現在'
)
# href="...zip"
_RE_ZIP_HREF = re.compile(r'href="([^"]+\.zip)"', re.IGNORECASE)


class KyushuAdapter(Adapter):
    bureau = "kyushu"

    def discover(self) -> Optional[DiscoveryResult]:
        html = http_get(INDEX_URL).decode("utf-8", "replace")

        # 最新月見出しを探す（最初に出てくるものが最新）
        m = _RE_MONTH_HEADER.search(html)
        if not m:
            return None
        y, mo = int(m.group(1)), int(m.group(2))

        # 次の月見出しまでの範囲を切り出す
        start = m.end()
        m2 = _RE_MONTH_HEADER.search(html, start)
        section = html[start:m2.start()] if m2 else html[start:]

        zip_urls = _RE_ZIP_HREF.findall(section)
        if len(zip_urls) != 8:
            # 想定外。8府県分そろってない場合は失敗扱い
            print(f"[kyushu] 想定外のZIP数: {len(zip_urls)} (期待値: 8)")
            return None

        # 順序を pref code にマップ
        file_refs = []
        for code, url in zip(_PREF_ORDER, zip_urls):
            full = url if url.startswith("http") else (BASE + url.lstrip("/"))
            file_refs.append(FileRef(
                url=full,
                filename=full.rsplit("/", 1)[-1],
                extra={"pref_code": code},
            ))

        # signature: 全URLの結合（どれか1つでも変われば再処理）
        signature = "|".join(r.url for r in file_refs)

        return DiscoveryResult(
            bureau=self.bureau,
            file_refs=file_refs,
            version=f"{2018 + y}.{mo}",
            year=2018 + y,
            month=mo,
            signature=signature,
        )

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        """ZIPから "shika" を名前に含むxlsxを抜き出す。"""
        out: List[Tuple[str, bytes]] = []
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            for nm in zf.namelist():
                low = nm.lower()
                if "shika" in low and low.endswith(".xlsx"):
                    out.append((nm, zf.read(nm)))
        return out
