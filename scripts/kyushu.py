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
from urllib.parse import urljoin

from lib import Adapter, DiscoveryResult, FileRef, http_get


INDEX_URL = ("https://kouseikyoku.mhlw.go.jp/kyushu/gyomu/gyomu/"
             "hoken_kikan/index_00007.html")

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

    def discover(self) -> List[DiscoveryResult]:
        """インデックスページから利用可能な全月を返す（過去月backfill対応）。"""
        html = http_get(INDEX_URL).decode("utf-8", "replace")

        # すべての月見出しを順に拾う（ページ上は新→旧の順で並んでいる）
        headers = list(_RE_MONTH_HEADER.finditer(html))
        if not headers:
            return []

        results: List[DiscoveryResult] = []
        for i, m in enumerate(headers):
            y, mo = int(m.group(1)), int(m.group(2))
            start = m.end()
            end = headers[i + 1].start() if i + 1 < len(headers) else len(html)
            section = html[start:end]

            zip_urls = _RE_ZIP_HREF.findall(section)
            if len(zip_urls) != 8:
                # 8府県分そろわないセクションはスキップ
                # （例：「略称一覧（令和8年2月1日現在）」のPDFリンクなど）
                continue

            full_urls = [urljoin(INDEX_URL, u) for u in zip_urls]
            file_refs = [
                FileRef(
                    url=full,
                    filename=full.rsplit("/", 1)[-1],
                    extra={"pref_code": code},
                )
                for code, full in zip(_PREF_ORDER, full_urls)
            ]
            # signature: 全URLの結合（どれか1つでも変われば再処理対象）
            signature = "|".join(full_urls)

            results.append(DiscoveryResult(
                bureau=self.bureau,
                file_refs=file_refs,
                version=f"{2018 + y}.{mo}",
                year=2018 + y,
                month=mo,
                signature=signature,
            ))

        return results

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        """ZIP内の全xlsxを取り出す。歯科の判別は共通パーサ側の 区分=='歯科' フィルタに任せる。

        ※命名規則が県ごとに揺れる可能性に備えて広めに拾う。
          医科・薬局のxlsxを読んでも、区分フィルタで0行になるだけで害は無い。
        """
        out: List[Tuple[str, bytes]] = []
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            for nm in zf.namelist():
                if nm.lower().endswith(".xlsx"):
                    out.append((nm, zf.read(nm)))
        print(f"[kyushu] {ref.filename}: xlsx {len(out)}件抽出")
        return out
