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
    parse_xlsx_to_records, write_pref_outputs, rebuild_prefectures_json,
)
from kinki import KinkiAdapter

# 有効化するアダプタ。新規追加はここに足すだけ。
ADAPTERS: list = [
    KinkiAdapter(),
]


def run_adapter(adapter: Adapter, state: dict) -> bool:
    """1局分を処理。変更があれば True、無ければ False。例外は呼び出し側で捕捉。"""
    bureau = adapter.bureau
    print(f"[{bureau}] discover …")
    disc = adapter.discover()
    if not disc:
        print(f"[{bureau}] 歯科ファイルが見つかりません（ページ構造変更の可能性）")
        return False

    sig_prev = state["bureaus"].get(bureau, {}).get("signature")
    print(f"[{bureau}] signature: {disc.signature} (prev={sig_prev})")
    if disc.signature == sig_prev:
        print(f"[{bureau}] 前回から変化なし、スキップ")
        return False

    n_prefs = 0
    for ref in disc.file_refs:
        print(f"[{bureau}] fetch {ref.filename}")
        blob = adapter.fetch(ref)
        for name, xlsx_bytes in adapter.extract_xlsxs(blob, ref):
            recs = parse_xlsx_to_records(xlsx_bytes)
            for rec in recs:
                write_pref_outputs(rec, bureau=bureau)
                rate = (rec.standards[0]["count"] / rec.total_clinics * 100
                        if rec.standards and rec.total_clinics else 0)
                print(f"  {rec.pref_code} {rec.pref_name}: "
                      f"母数={rec.total_clinics} 種類={len(rec.standards)} "
                      f"trend1={rate:.2f}%")
                n_prefs += 1

    if n_prefs == 0:
        print(f"[{bureau}] パース結果ゼロ。失敗扱い、stateは更新しません")
        return False

    state["bureaus"][bureau] = {
        "signature": disc.signature,
        "version": disc.version,
        "asof": f"令和{disc.year - 2018}年{disc.month}月1日現在",
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
