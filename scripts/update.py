#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全国 歯科 施設基準ダッシュボード 更新エントリポイント。

各厚生局のアダプタを順に呼び出し、新しいファイルが見つかれば取得・集計して
府県別 current/history/source、および全体メタ (prefectures.json) を更新する。

- 各局のアダプタは scripts/<bureau>.py に置く（lib.Adapter を継承）
- 新しい局を追加するときは下の ADAPTERS リストに追加するだけ
- 1局のエラーは他局に波及しない（fail-soft：その局はスキップして次へ）
"""

import os
import sys
import traceback
from pathlib import Path

# scripts/ ディレクトリを import path に追加（lib, <bureau>.py を見えるように）
sys.path.insert(0, str(Path(__file__).resolve().parent))

from lib import (
    Adapter,
    ensure_dirs, load_state, save_state, now_jst_str,
    parse_xlsx_to_records, merge_records_by_pref,
    write_pref_outputs, rebuild_prefectures_json,
)
from kinki import KinkiAdapter
from kyushu import KyushuAdapter

# 有効化するアダプタ。新規追加はここに足すだけ。
ADAPTERS: list = [
    KinkiAdapter(),
    KyushuAdapter(),
]


def run_adapter(adapter: Adapter, state: dict) -> bool:
    """1局分を処理。新規データがあれば True、無ければ False。例外は呼び出し側で捕捉。

    - discover() は利用可能な全月のリストを返す
    - 既に処理済みのsignatureは飛ばす
    - 残り（新規）を古い月から順に処理（history を時系列順に追記するため）
    """
    bureau = adapter.bureau
    bureau_state = state["bureaus"].get(bureau, {})
    processed_sigs = set(bureau_state.get("signatures", []))

    print(f"[{bureau}] discover…")
    discoveries = adapter.discover()
    if not discoveries:
        print(f"[{bureau}] 利用可能なファイルが見つかりません")
        return False

    new_discs = [d for d in discoveries if d.signature not in processed_sigs]
    if not new_discs:
        print(f"[{bureau}] 新規データなし、スキップ（{len(discoveries)}月分すべて処理済み）")
        return False

    # 古い月から処理（history を時系列順に追記）
    new_discs.sort(key=lambda d: (d.year, d.month))
    print(f"[{bureau}] {len(new_discs)}月分の新規データを処理: "
          f"{[d.version for d in new_discs]}")

    n_succeeded_months = 0
    for disc in new_discs:
        print(f"\n[{bureau}] === {disc.version} ===")
        all_records = []
        for ref in disc.file_refs:
            print(f"[{bureau}] fetch {ref.filename}")
            try:
                blob = adapter.fetch(ref)
            except Exception as e:
                print(f"[{bureau}] fetch失敗 ({ref.filename}): {e!r}")
                continue
            for name, xlsx_bytes in adapter.extract_xlsxs(blob, ref):
                recs = parse_xlsx_to_records(xlsx_bytes)
                all_records.extend(recs)

        merged = merge_records_by_pref(all_records)
        if not merged:
            print(f"[{bureau}] {disc.version}: 0府県分しか取れず、この月はスキップ")
            continue

        for rec in merged:
            write_pref_outputs(rec, bureau=bureau)
            rate = (rec.standards[0]["count"] / rec.total_clinics * 100
                    if rec.standards and rec.total_clinics else 0)
            print(f"  {rec.pref_code} {rec.pref_name}: "
                  f"母数={rec.total_clinics} 種類={len(rec.standards)} "
                  f"trend1={rate:.2f}%")
        processed_sigs.add(disc.signature)
        n_succeeded_months += 1

    if n_succeeded_months == 0:
        print(f"[{bureau}] 1月分も成功せず、stateは更新しません")
        return False

    # 処理済みの中で最新月を求める
    all_processed_discs = [d for d in discoveries if d.signature in processed_sigs]
    latest = max(all_processed_discs, key=lambda d: (d.year, d.month))

    state["bureaus"][bureau] = {
        "signatures": sorted(processed_sigs),
        "latest_version": latest.version,
        "asof": f"令和{latest.year - 2018}年{latest.month}月1日現在",
        "checked": now_jst_str(),
    }
    return True


def main():
    ensure_dirs()
    now = now_jst_str()
    state = load_state()

    any_changed = False
    failures: list = []
    for ad in ADAPTERS:
        try:
            if run_adapter(ad, state):
                any_changed = True
        except Exception as e:
            traceback.print_exc()
            failures.append((ad.bureau, repr(e)))
            print(f"[{ad.bureau}] エラー: {e!r}（この局はスキップ、前回値を保持）")

    rebuild_prefectures_json(now, state)
    save_state(state, now)

    print(f"done. any_changed={any_changed} failures={[b for b,_ in failures]}")

    # GitHub Actions の outputs に出力
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a", encoding="utf-8") as f:
            f.write(f"changed={'true' if any_changed else 'false'}\n")
            f.write(f"failures={','.join(b for b,_ in failures)}\n")

    # 全局失敗なら CI を赤くしたい（cron が無効化されないために exit 0 にする手もあるが、
    # 一旦 1 にして気付ける形にする）
    if failures and not any_changed:
        sys.exit(1)


if __name__ == "__main__":
    main()
