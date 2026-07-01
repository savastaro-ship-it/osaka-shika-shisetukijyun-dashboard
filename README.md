# 歯科施設基準 届出状況ダッシュボード

47都道府県の歯科医療機関について、施設基準の届出状況を可視化するダッシュボード。全国8つの地方厚生局が公開している「施設基準の届出状況」データを自動収集・集計している。

**ライブ URL**: https://savastaro-ship-it.github.io/osaka-shika-shisetukijyun-dashboard/

---

## 全体アーキテクチャ

**Python（データ収集） + HTML/JS（表示） が完全分離** している。

```
┌────────────────────────────────────────┐
│ 全国8つの地方厚生局サイト               │
│  hokkaido / tohoku / kantoshinetsu /   │
│  tokaihokuriku / kinki / chugoku /     │
│  shikoku / kyushu                      │
└────────────┬───────────────────────────┘
             │ 毎日 cron
             ▼
┌────────────────────────────────────────┐
│ GitHub Actions (.github/workflows/)    │
│ scripts/update.py が全adapter駆動      │
│  ・discover: 最新月のZIP/xlsxを検出    │
│  ・fetch: バイト列で取得               │
│  ・parse: xlsxから届出情報を抽出       │
└────────────┬───────────────────────────┘
             │ 出力
             ▼
┌────────────────────────────────────────┐
│ data/*.json （GitHubレポにコミット）    │
│  ・prefectures.json  全府県メタ + 版   │
│  ・current/<code>.json  府県別最新値   │
│  ・history/<code>.json  府県別時系列   │
│  ・state.json  Actions内部状態         │
│  ・source/<code>.xlsx  最新取得xlsx    │
└────────────┬───────────────────────────┘
             │ HTTP fetch
             ▼
┌────────────────────────────────────────┐
│ GitHub Pages (index.html)              │
│  ・府県別ビュー                         │
│  ・受理記号別ビュー（横断分析）         │
│  ・全国集計タブ（速報/確定値切替）     │
└────────────────────────────────────────┘
```

**この分離のメリット**:
- 表示の改良は `index.html` だけ触れば済む（Python 不要）
- 別クライアント（スマホアプリ・Slack Bot 等）を作るときも同じ JSON を読めばOK
- JSON URL がそのまま実質的な API になっている

---

## ディレクトリ構成

```
osaka-shika-shisetukijyun-dashboard/
├── README.md                    # このファイル
├── index.html                   # ダッシュボード本体（HTML/CSS/JS 1ファイル）
├── scripts/                     # Python: データ収集・パース
│   ├── lib.py                   # 共通ライブラリ（Adapter基底・パーサ・データモデル）
│   ├── update.py                # エントリポイント（全adapter駆動）
│   ├── hokkaido.py              # 各厚生局アダプタ
│   ├── tohoku.py
│   ├── kantoshinetsu.py
│   ├── tokaihokuriku.py
│   ├── kinki.py
│   ├── chugoku.py
│   ├── shikoku.py
│   └── kyushu.py
├── .github/workflows/
│   └── update.yml               # cron設定（毎日06:00 JST頃）
└── data/                        # Actions が自動生成
    ├── prefectures.json         # 全府県メタ（並び順・版など）
    ├── state.json               # Actions内部状態（処理済signature）
    ├── current/
    │   ├── 00.json              # 全国（47府県の集計）
    │   ├── 01.json              # 北海道
    │   ├── ...
    │   └── 47.json              # 沖縄
    ├── history/                 # 時系列（同構造）
    └── source/                  # 最新取得xlsx（デバッグ用）
```

---

## Adapter パターン

`scripts/lib.py` の `Adapter` 基底クラスを継承して、各局のデータ取得ロジックを実装している。

```python
class Adapter:
    bureau: str  # "kyushu" など

    def discover(self) -> List[DiscoveryResult]:
        """厚生局サイトから最新月のZIP/xlsxを見つける"""

    def fetch(self, ref: FileRef) -> bytes:
        """ZIP/xlsxをHTTP取得"""

    def extract_xlsxs(self, blob: bytes, ref: FileRef) -> List[Tuple[str, bytes]]:
        """ZIPの中からxlsxを取り出す（xlsx直リンクの場合はそのまま返す）"""
```

- `discover()` が **各局のサイト構造の違いを吸収** する部分
- 抽出後の xlsx は `parse_xlsx_to_records()` （共通パーサ）に流れる
- パーサは「区分==歯科」フィルタと「都道府県コード」でgroupbyして府県別レコードを作る

### 新しい局・データソースを追加するには

1. `scripts/<name>.py` を作り、`Adapter` を継承したクラスを実装
2. `scripts/update.py` の `ADAPTERS` リストに追加
3. `scripts/lib.py` の `BUREAU_PREFS` / `BUREAU_ORDER` を更新（局ごとの管轄府県を定義）

---

## 各厚生局の xlsx 構造（実装ハマりポイント集）

各局でxlsxのフォーマットや配布方法が微妙に違う。それぞれの特徴と対策：

| 局 | 府県 | 配布形式 | 特徴・注意点 |
|---|---|---|---|
| **北海道** | 01 | xlsx直リンク | シンプル。HTMLに直接xlsxリンク |
| **東北** | 02-07 | xlsx直リンク（6県分1ファイル） | 1xlsxに東北6県分のデータが混在。共通パーサの「都道府県コード」でgroupbyして分離 |
| **関東信越** | 08-15, 19, 20 | 1ZIP+10都県 | **HTMLの更新が遅延することがある**。実際にZIPがサーバに上がっていてもHTMLのリンクが古いまま。→ URL推測型（`shisetsu_shika_r{令和年}{月}.zip`）でHEADリクエスト確認 |
| **東海北陸** | 16, 17, 21-24 | 1ZIP+6県分 | 1ZIPに東海北陸6県分のxlsxが同梱 |
| **近畿** | 18, 25-30 | 1ZIP+7府県分 | データ訂正時は `s{version}_...zip` のように`s`プレフィックス付き。両方対応 |
| **中国** | 31-35 | 1ZIP+5県分 | **区分混在xlsx**（医科・歯科・薬局が同ファイル）。共通パーサの「区分==歯科」フィルタが必須 |
| **四国** | 36-39 | 歯科専用ZIP | 中国とは別組織（**四国厚生支局**）。医科・歯科・薬局が別ZIP |
| **九州** | 40-47 | 8ZIP×1県ずつ | **過去12ヶ月分のアーカイブがHTMLに残る** → backfillで全月取得可能 |

### 特殊対応が必要だった事例

**関東信越の HTML遅延問題**: HTMLに載ってる最新リンクが古い月のままでも、サーバには新月のZIPが上がってることが多い。ファイル名規則が予測可能（`shisetsu_shika_r{令和年2桁}{月2桁}.zip`）なので、現在月から未来1ヶ月・過去12ヶ月をHEADリクエストで確認して最新を採用する。

**九州のbackfill後 current 上書きバグ**: 「未処理月を古い順に処理」した結果、既に処理済みの最新月ではなく最後に処理した過去月で current が上書きされていた。対策：
- `write_pref_outputs` で「既存より古い月の処理なら current は上書きしない」（再発防止）
- `reconcile_currents_with_history()` で毎回Actions実行時に current と history の整合性チェック・自動修復（既存バグの自動修復）

**中国四国の組織分離**: 「中国四国厚生局」という組織はなく、中国5県は「中国四国厚生局」（広島本局）、四国4県は「四国厚生支局」（別組織）。URLパス・ページ構造が別。アダプタは `chugoku` と `shikoku` で分離。

---

## データフォーマット

### `data/prefectures.json`

全府県のメタ情報。フロントが最初に読むファイル。

```json
{
  "version": "2026.5",
  "asof": "令和8年5月1日現在",
  "checked": "2026/06/30 09:40",
  "versions": ["2025.6", "2025.7", "...", "2026.5"],
  "prefectures": [
    {"code": "00", "name": "全国", "total_clinics": 65953},
    {"code": "27", "name": "大阪", "total_clinics": 5260}
  ]
}
```

並び順は `FEATURED_TOP` (`["00", "27"]` = 全国、大阪) が先頭、次に `BUREAU_ORDER` 順にJIS昇順。

### `data/current/<code>.json`

各府県の最新月データ。1府県1ファイル。

```json
{
  "code": "27",
  "name": "大阪",
  "bureau": "kinki",
  "version": "2026.5",
  "asof": "令和8年5月1日現在",
  "total_clinics": 5260,
  "n_standards": 113,
  "standards": [
    {"kigo": "補管", "name": "クラウン・ブリッジ維持管理料", "count": 5145, "count_uniq": 5145},
    {"kigo": "医療ＤＸ", "name": "医療ＤＸ推進体制整備加算", "count": 2350, "count_uniq": 2350}
  ]
}
```

- `code: "00"` は 47府県の合算値（速報値ベース、現時点合算）

### `data/history/<code>.json`

各府県の時系列データ。

```json
{
  "versions": ["2025.6", "2025.7", "...", "2026.5"],
  "totals": {"2025.6": 5240, "2025.7": 5245, "2026.5": 5260},
  "kigo": {
    "補管": {
      "name": "クラウン・ブリッジ維持管理料",
      "names": [{"v": "2025.6", "name": "クラウン・ブリッジ維持管理料"}],
      "series": [
        {"v": "2025.6", "c": 5100, "u": 5100},
        {"v": "2025.7", "c": 5110, "u": 5110}
      ]
    }
  }
}
```

`count_uniq` は将来の重複排除用（現状は `count` と同じ値が多い）。

### `data/state.json`

Actions内部の処理状態。フロントは読まない。

```json
{
  "bureaus": {
    "kyushu": {
      "signatures": ["<URL集合>"],
      "latest_version": "2026.5",
      "asof": "令和8年5月1日現在",
      "checked": "2026/06/30 09:40"
    }
  }
}
```

- `signatures`: 処理済みZIP/xlsxの識別子（URL）リスト。差分実行の判定に使う
- **`latest_version` は discover() のヒント値（URL/ページから推定）** で、実データas-of と一致しないことがある。真のas-ofは各 `current/<code>.json` の `version` を参照

---

## フロントエンド機能

### 府県別ビュー（デフォルト）

- 47府県 + 全国タブ（合計48タブ）
- タブ順：★全国 → ★大阪 → 北海道 → ... → 沖縄（JIS順）
- KPI・受理記号一覧・時系列グラフ
- ヘッダーの「府県名 ／ 歯科」「データ版XXXX.X」は選択府県で動的更新

### 全国タブ専用：速報値 / 確定値トグル

- **確定値（デフォルト）**: `history/00.json` の最終月のスナップショット
- **速報値**: `current/00.json`（各府県の最新月を単純合算）
- 混在月時は「内訳：2026.5：39府県／2026.4：8府県」のように月別府県数を表示（府県名も直接展開）

### 受理記号別ビュー

- 上部のビューモード切替で「府県別／受理記号別」
- 受理記号タブ（届出件数上位50個）
- 選択記号について、47府県の届出率を横棒グラフ表示
- 全国平均線（赤の点線）、色分け（全国+5pt/-5pt）
- ソート切替：届出率順／厚生局順／JIS順
- TOP3 / WORST3 ハイライト
- **Excel ダウンロード**（現在ソート順、末尾に全国合算行）

---

## 開発・運用

### 運用スタイル

- GitHub Web画面のみで運用（git コマンド不使用でも可）
- Actions は毎日 cron 実行（約06:00 JST）
- 手動実行も可能（「Actions」タブ → workflow → Run workflow）

### ローカルテスト

`scripts/` フォルダに `cd` して、`update.py` を実行：

```bash
cd scripts
python3 update.py
```

初回は全府県のデータを取得するので数分かかる。差分実行時は10〜30秒。

### 差分実行の判定

各局の `state.json` に `signatures`（URL集合）が保存されており、これと `discover()` の戻り値を照合して未処理のものだけ処理する。同じsignatureは再取得しない（差分実行）。

### state.json をリセットしたい時

- 該当局のエントリを削除するとフルリセット（全月再取得）
- signatures だけクリアすると次回のみ全月扱い

### `reconcile_currents_with_history()` の存在

毎回 update.py 実行時に、各府県の `current/<code>.json` の `version` が `history/<code>.json` の最新月と一致しているかチェックし、ズレていれば history から作り直す。冪等（既に正しければ何もしない）。**過去のbackfillバグの自動修復** として組み込まれている。

---

## 既知の制約・落とし穴

### xlsx as-of の判定

`_RE_ASOF = r"令和(\d+)年(\d+)月(\d+)日\s*現在"` で xlsx セルから読み取る。厚生局によっては「令和X年Y月作成」「令和X年Y月1日現在」が同じシートに混在するが、「現在」の有無で識別できる。

### 「公表月」と「データas-of月」

例：関東信越の `shisetsu_shika_r0806.zip`（令和8年6月公表）の中身は「令和8年5月1日現在」のデータ。
- `state.json.bureaus.kantoshinetsu.latest_version = "2026.6"` （公表月）
- `current/13.json.version = "2026.5"` （データas-of）

両者が食い違うのは仕様。フロントに出るのは後者。

### GitHub Pages のキャッシュ

`index.html` を更新しても、ブラウザキャッシュで古いのが表示されることがある。ユーザーには「強制リロード（Ctrl+Shift+R / Cmd+Shift+R）」を案内する。

---

## 将来の拡張案

### 実装が軽い（フロントのみ）

- 日本地図（コロプレス）
- 受理記号カテゴリ分類（基本・DX・技術・感染対策など）
- 全国平均との乖離マップ
- 増減ハイライト（最近1年で急増/減少した基準）
- 受理記号の詳細情報（点数、要件、告示リンク）

### バックエンド改修が必要

- **医療機関別インデックス**: 現状パーサは「基準別集計」のみ。xlsxには「医療機関名」列がある。パーサに列追加すれば、医院検索ポータル化が可能
- **要件マスタ**: 厚労省の告示ページから受理記号ごとの要件をスクレイプ →「取得手順ガイド」化
- **月次スナップショット強化**: 現状は各府県の history が「単一府県の時系列」。ある月の全国断面図（府県×基準の全マトリクス）を別途生成すると分析が楽

### マネタイズ想定

- 検索ポータル（フリーミアム、月額¥1,980程度）
- ターゲティングデータ販売（メーカー・コンサル向け）
- 月次PDFレポート（歯科経営コンサル向け）
- API 提供（研究機関・シンクタンク向け）

---

## 転用時の情報伝達（別プロジェクトで流用する場合）

このシステムのスクレイピング＋パース部分は他分野にも流用しやすい構造。転用時は以下を用意すると新規実装の詰まりが減る：

1. **既存コード一式**: `scripts/` フォルダ + `.github/workflows/update.yml`
2. **実データサンプル**: パースしたい実xlsx を2〜3個（構造バリエーション把握用）
3. **現在の稼働状態**: 最新の `state.json` + `data/current/*.json` の1府県分
4. **要件1文**: 「何を作りたいか」「既存の何を流用したいか」

---

## クレジット

- データ元: 全国8つの地方厚生局が公開している「施設基準の届出状況」
- ライセンス: このリポジトリのコードは MIT ライセンス（元データは各厚生局に帰属）
- お問い合わせ: (issue やPR歓迎)
