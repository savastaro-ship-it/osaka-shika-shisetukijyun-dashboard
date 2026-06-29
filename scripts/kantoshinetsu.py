# -*- coding: utf-8 -*-
"""
関東信越厚生局アダプタ。

掲載ページ: kantoshinetsu/chousa/kijyun.html
構造     : 1ZIPに10都県分のxlsxが入っている（東海北陸と同パターン）
府県     : 08-15 + 19, 20（10都県）
過去月   : アーカイブなし

⚠️ 特殊事情：
  HTMLのリンク更新が遅延する運用らしく、サーバにZIPが上がっていても
  HTMLには古いリンクのままのことがある。
  そのため、HTMLをスクレイピングせず、**URLパターンを推測**して
  HEADリクエストで存在確認し、最新のZIPを採用する。

ファイル名規則：
  shisetsu_shika_r{令和年2桁}{月2桁}.zip
  例: 令和8年6月公表 → shisetsu_shika_r0806.zip
"""

import io
import urllib.error
import urllib.request
import zipfile
from datetime import datetime, timezone, timedelta
from typing import List, Tuple

from lib import Adapter, DiscoveryResult, FileRef, UA


ZIP_URL_TEMPLATE = (
    "https://kouseikyoku.mhlw.go.jp/kantoshinetsu/"
    "shisetsu_shika_r{reiwa:02d}{month:02d}.zip"
)

JST = timezone(timedelta(hours=9))

# 何ヶ月探索するか
LOOK_AHEAD_MONTHS = 1   # 念のため未来1ヶ月先まで（厚生局が早めに公表する場合の保険）
LOOK_BACK_MONTHS = 12   # 過去12ヶ月遡る


def _url_exists(url: str, timeout: int = 15) -> bool:
    """HEAD で200を確認。HEAD不可ならGET(Range:0-0)でフォールバック。"""
    for method in ("HEAD", "GET"):
        req = urllib.request.Request(url, method=method)
        req.add_header("User-Agent", UA)
        if method == "GET":
            req.add_header("Range", "bytes=0-0")
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                if r.status in (200, 206):
                    return True
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return False
            # 405 (Method Not Allowed) などは次のmethodで再試行
            continue
        except Exception:
            return False
    return False


def _iter_candidate_months(now: datetime, ahead: int, back: int):
    """nowの ahead ヶ月先から back ヶ月遡って、(year, month) を新→旧の順で返す。"""
    y, m = now.year, now.month + ahead
    while m > 12:
        y += 1
        m -= 12
    for _ in range(ahead + back + 1):
        yield y, m
        m -= 1
        if m <= 0:
            y -= 1
            m += 12


class KantoshinetsuAdapter(Adapter):
    bureau = "kantoshinetsu"

    def discover(self) -> List[DiscoveryResult]:
        now = datetime.now(JST)
        for y, m in _iter_candidate_months(now, LOOK_AHEAD_MONTHS, LOOK_BACK_MONTHS):
            reiwa = y - 2018
            if reiwa < 1:
                break
            url = ZIP_URL_TEMPLATE.format(reiwa=reiwa, month=m)
            if _url_exists(url):
                print(f"[kantoshinetsu] 公表ZIP発見: "
                      f"r{reiwa:02d}{m:02d}.zip (令和{reiwa}年{m}月公表)")
                return [DiscoveryResult(
                    bureau=self.bureau,
                    file_refs=[FileRef(url=url, filename=url.rsplit("/", 1)[-1])],
                    version=f"{y}.{m}",  # 公表月、データas-ofは共通パーサが上書きする
                    year=y,
                    month=m,
                    signature=url,
                )]
        print("[kantoshinetsu] URL推測で公表ZIPが見つかりませんでした")
        return []

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        """ZIP内の全xlsxを取り出す。府県の判別は共通パーサ側の 都道府県コード で。"""
        out: List[Tuple[str, bytes]] = []
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            for nm in zf.namelist():
                if nm.lower().endswith(".xlsx"):
                    out.append((nm, zf.read(nm)))
        print(f"[kantoshinetsu] {ref.filename}: xlsx {len(out)}件抽出")
        return out
