---
name: takimoto-valuation
description: 瀧本ゼミ形式のバリュエーションシート(.xlsx)を生成する。入力はストーリーmarkdown + 決算短信/有価証券報告書PDFs。出力はGoogle Sheetsにインポート可能な.xlsxで、因数分解+P&L+シナリオ別目標株価+ピア比較+グラフを含む。「バリュエーションシート作って」「スプレッドシート作って」「投資推奨のスプシ作って」「○○のスプシ作って」のような依頼で起動。スライド(.pptx)が欲しい場合はtakimoto-slidesスキルを使う。
---

# Takimoto Valuation Sheet Generator

瀧本ゼミ形式の株式投資推奨用バリュエーションシート(`.xlsx`)を生成するスキル。**ストーリーmarkdown + 複数の決算開示PDF** から、Google Sheets にインポート可能な構造化されたバリュエーション資料を出力する。

## 起動条件

- 「バリュエーションシート作って」「スプレッドシート作って」「.xlsxにして」
- 「○○(銘柄名)のスプシ作って」「投資推奨のスプシ作って」
- `/takimoto-valuation` を明示的に呼び出した時
- 入力markdownと決算PDFを渡してきて「数理モデル作って」と依頼

**注**: スライド(.pptx)が欲しい場合は **`takimoto-slides`** スキルを使う。両方欲しい場合は両方起動する。

## 必要な入力

| 入力 | 必須? | 形式 | 用途 |
| :---- | :---- | :---- | :---- |
| ストーリーmarkdown | **必須** | `samples/*.md`の構造 | 因数分解・前提・シナリオ |
| 決算開示PDFのフォルダ | **必須** | 複数PDF(yuho_101/eigyo/hanpo/yuho2Q) | 過去実績、セグメント、KPI |
| PER感応度 / ピア企業 | 任意 | AskUserQuestionで取得 | 目標株価のレンジ確定 |

**PDFを複数渡してもらう理由**: 有報・決算短信・決算説明資料はそれぞれ違う情報を持っている。例えば、有報は5年指標+詳細P&L、招集通知は中期計画の数値、決算説明資料は事業セグメント別KPIや業界比較データ。スキルは`extract_folder()`で全部読んで横断的に最良値を採用する。

## 設計原則

**Python(決定的) と Opus(柔軟) で役割を分ける**:

| 担当 | やること |
| :---- | :---- |
| **Python (`sheet_renderer.py`)** | 色・カンマ・インデント・グラフ・固定枠・グリッド線オフ・備考必須チェック |
| **Python (`extract_yuho.py`)** | PDFのテキスト化、P&L/SG&A/セグメント/5年指標/中計目標の構造化抽出 |
| **Python (`validate_consistency.py`)** | 売上-原価=粗利、粗利-販管費=営利、粗利率=粗利/売上、増収率、青セル備考必須、TODO残検出 |
| **Opus (あなた)** | **業績タイプ判定**、**因数分解の式評価**、**FY26+の数値計算**、**ピア企業の多軸選定**、**「織り込み済み反論」のPER計算**、**推論で埋めた数値の正当化** |

つまり**Opusの仕事は「数学」**。Pythonに計算を任せない。

## 品質第一・5パス検証

**絶対に1パスで仕上げない**。Maxプラン前提なのでトークン気にせず多パス回す。

### Pass 1: 入力理解 + PDF抽出

```python
import sys
sys.path.insert(0, "<skill_dir>/scripts")
from extract_yuho import extract_folder
fin = extract_folder("~/Downloads/4971-yuhos/")
# → indicators_5y, pl, sga, product_breakdown, segment_info,
#    shares_outstanding, mid_term_plan, _sources
```

markdownを読み:
- **業績タイプを判定** (`reference/factor_decomp_patterns.md` の5パターンに当てはめる)
- **因数分解の式**を式表記で抽出(0️⃣セクションの ``` ブロック)
- **各因数の方向と数値**(因数表)
- **シナリオ(Bear/Base/Bull)の前提**
- **言及されている過去業績の数値**(yuho と突合に使う)

### Pass 2: 骨組み生成

```python
from sheet_renderer import Row, SheetConfig, ChartSpec, render
config = SheetConfig(
    name="XXXX_valuation",
    year_columns=["FY22 実", "FY23 実", "FY24 実", "FY25 実", "FY26 予", "FY27 予"],
    rows=[
        Row(label="【売上高】", is_section_header=True),
        Row(label="売上高 合計", values=[...PDFから...], note="yuho開示"),
        # ↓ 業績タイプに応じた因数分解
        Row(label="【因数分解】", is_section_header=True),
        Row(label="ドライバー1", indent=1, values=[...自分で計算...],
            is_independent=True, importance="★★★", note="markdown仮定"),
        # ...
        # P&L
        # 工場稼働率 等
        # EPS, PER 感応度
        # 「織り込み済み?」反論ブロック
        # ピア比較ブロック
    ],
    charts=[
        ChartSpec(title="売上・利益・利益率",
                  bar_rows=["売上高 合計", "売上総利益", "営業利益"],
                  line_rows=["営業利益率"],
                  anchor="J2"),
        # 業績タイプ別の主要ドライバーチャート
        # シナリオ別目標株価チャート
    ],
)
todos = render(config, "output.xlsx")
```

### Pass 3: FY26+ プロジェクションの計算

**ここがOpusの本領**。markdownの因数分解式と仮定を読み取り、毎年の数値を**自分の頭で計算して**埋める。

例(メック):
- 「GPU生産+30/20/10%、Hopper/Blackwell/Rubin比率変化、面積/層数加重平均、使用量補正50%」
- → FY25/FY26/FY27のGPU向けCZ需要量(相対値)を計算
- → 単価据置なら需要量比 × FY24実績売上 で FY25-FY27 GPU向け売上

業績タイプ別の計算ロジックは `reference/factor_decomp_patterns.md` を参照。

### Pass 4: 整合性検証

```python
sys.path.insert(0, "<skill_dir>/scripts")
from validate_consistency import validate, print_findings
findings = validate("output.xlsx")
print_findings(findings)
```

検出される問題:
- ERROR: 算術不整合 (売上-原価≠粗利 等) → **必ず修正**
- WARN: 比率不整合 (粗利率と粗利÷売上 が違う) → 確認して修正
- INFO: TODO残数 → 推論で埋めるか AskUserQuestion

### Pass 5: 1対1対応検証

markdownの各主張に対して:
- スプシのどの行に反映されているか
- 逆にスプシの行は markdownのどこに根拠を持つか

漏れがあれば追加行、根拠が無ければ削除 or 備考で明示。

## 「PDFに無い数値」をどう扱うか

優先順序(SKILL.md上部の "PDFに無い数値" セクション参照):

1. **複数PDFの横断検索** — yuhoに無くても招集通知/決算説明資料にあるかも
2. **同業他社/業界統計からの推論** — 業界平均、競合IR、政府統計
3. **markdown内の手がかり** — 「○○の比率は約○○」のような言及
4. **フェルミ推定 / 隣接指標からの逆算** — 純利益÷EPS で発行済株式数を逆算 等
5. **TODO残し + AskUserQuestion** — 最後の手段

推論で埋めた値は備考に必ず明記:
```
備考: 「IR電話 / 純利益÷EPSから逆算 / 業界平均(15%)から保守的に14%とおく」
```

## スプシ必須セクション

`reference/sheet_structure.md` に従い、以下を必ず含める:

1. **売上高 + 増収率**
2. **セグメント別売上**(業績タイプによって柔軟)
3. **因数分解**(業績タイプに応じた構造)
4. **売上原価 + 粗利 + 粗利率**(製品別粗利率があれば内訳)
5. **販管費 + 内訳**(給料/荷造/研究開発/減価償却/その他)
6. **営業利益 + 営業利益率**
7. **経常利益 + 営業外損益**
8. **税前純利益 + 法人税率(実効) + 純利益**
9. **EPS + 発行済株式数**
10. **PER感応度** (Bear/Base/Bull の3シナリオ × 各年度)
11. **「織り込み済み?」反論ブロック** — 時価総額据置で純利益増→PER低下を機械的に示す
12. **ピア比較** — 自社現状PER + 多軸選定の3社程度
13. **業績タイプ別 補助KPI**(例: メック=工場稼働率、AlbaLink=営業1人粗利)
14. **チャート3つ程度** — 売上+利益+利益率の棒+線、業績ドライバー、目標株価シナリオ

## 業績タイプの判定

入力markdownの 0️⃣ セクション ``` ブロックを見て、`reference/factor_decomp_patterns.md` の5パターンから判定:

| パターン | 判定キーワード | 例 |
| :---- | :---- | :---- |
| ① 物量×物理量 | 「面積」「層数」「mm²」「kg」等の物理単位 | メック |
| ② コホート | 「1年目」「2年目」「経験年数」「定着率」 | ムービン |
| ③ セグメント×粗利率改善 | 「新規」「既存」「リニューアル」「セグメントA × 粗利率」 | かわでん |
| ④ 1人粗利×営業人数 | 「営業1人当たり」「店舗数」「店舗あたり人数」 | AlbaLink |
| ⑤ 持分法 | 「持分法利益」「出資比率」「関連会社」 | 保土谷 |

迷ったら `AskUserQuestion` で確認。

## PER感応度 + ピア比較の作法

### PER感応度

3シナリオ(Bear/Base/Bull)を作り、それぞれ:
- PER水準を独立変数(青)として明記
- 想定純利益 × PER ÷ 発行済株式数 = 目標株価
- 各シナリオの前提を備考に書く

### 「織り込み済み?」反論ブロック

時価総額据置のシナリオで、将来純利益が増えた時の PER がどれくらい低下するかを示す:

```
時価総額 (据置)          [独立]   現値の時価総額を全年度固定
連結純利益 Base (再掲)   [従属]
時価総額据置PER          [従属]   = 時価総額 ÷ 純利益 → 利益増で機械的に低下
```

この一連の行で「もしPER 据置なら純利益増で株価上昇」を視覚化。

### ピア比較(3社程度)

- 純粋同業がいる場合 → 同業3社のPER
- 純粋同業がいない場合 → **多軸選定** (`reference/asking_user.md`):
  - 業態/事業内容で似た会社
  - ビジネスモデル/価格戦略で似た会社
  - 地域戦略/規模戦略で似た会社
  - それぞれ備考に「○○軸での類似」と明記

## 実行手順

### Step 0: 入力の所在を確認

ユーザーに次を確認(明示済みなら省略):
1. **ストーリーmarkdown のパス** (例: `samples/4971-mec.md`)
2. **決算PDFを入れたフォルダのパス** (例: `~/Downloads/4971-yuhos/`)

### Step 1: 環境セットアップ

```bash
which pdftotext || sudo apt-get install -y poppler-utils
python3 -c "import openpyxl" 2>/dev/null || pip3 install openpyxl
```

### Step 2: PDF抽出

```python
from extract_yuho import extract_folder
fin = extract_folder("<pdf_folder>")
print(f"Sources: {len(fin['_sources'])}, P&L years: {list(fin['pl'].keys())}")
```

### Step 3: 業績タイプ判定

markdown の 0️⃣ 因数分解式を読み、`reference/factor_decomp_patterns.md` の5パターンから1つ選ぶ。決まらなければ AskUserQuestion。

### Step 4: 行構造の組み立て

`reference/sheet_structure.md` を読み、業績タイプに応じた行構造を `Row` データクラスのリストに展開。

### Step 5: FY26+ の数値を**自分で計算**

markdownの因数分解式に従って、FY26/FY27 等の予想値を**自分の頭で計算**。Pythonに計算ロジックを書かない。結果を `Row.values` に直接埋め込む。

### Step 6: グラフ仕様の定義

`ChartSpec` を3つ程度作る:
- 売上・利益(棒) + 利益率(線・第二軸ドット付き)
- 業績ドライバーの主要因数(棒+線)
- シナリオ別目標株価(棒)

### Step 7: レンダリング + 検証

```python
todos = render(config, "<output_path>.xlsx")
findings = validate("<output_path>.xlsx")
print_findings(findings)
```

ERRORは必ず修正。WARNとTODOは備考で説明できるなら残してOK。

### Step 8: ユーザーへ報告

- 出力ファイルのパス (例: `outputs/XXXX-company-valuation.xlsx`)
- 推論で埋めた数値のリスト + その根拠
- 残TODO数 と AskUserQuestion で確認したい項目
- 「Google Driveにアップしてシート形式で開いてください」

## やっちゃダメなこと

- **markdownにない数値を捏造しない** — 不明なら TODO のまま残し、推論なら備考で明示
- **スタイルルールを破らない** — 黒/青、カンマ、インデント、グリッド線オフ、固定枠、★、備考は機械的に守る
- **業績タイプを保土谷型(持分法)に決めつけない** — markdownの因数分解を素直に読む
- **計算をPythonに書かない** — FY26+の予想値はOpus自身が頭で計算
- **PER議論を比較なしで終わらせない** — ピア(多軸選定)+「織り込み済み反論」両方入れる
- **青セルに備考無しを許さない** — レンダラーが警告を出すので、見つかったら必ず埋める
- **検証ERRORを残さない** — `validate_consistency.py` の ERROR は必ず解消

## ファイル構造

```
.claude/skills/takimoto-valuation/
├── SKILL.md                          ← 本ファイル
├── reference/
│   ├── sheet_structure.md            ← スプシ行構造規約
│   ├── factor_decomp_patterns.md     ← 業績タイプ5種のパターン
│   └── asking_user.md                ← AskUserQuestion テンプレ(PER, ピア企業 等)
└── scripts/
    ├── sheet_renderer.py             ← スプシスタイル強制
    ├── extract_yuho.py               ← yuho PDF抽出(複数PDF対応)
    └── validate_consistency.py       ← 整合性検証
```

## 参考例

開発リポジトリ `kanaeda0225/my-claude-project` の以下:

- `samples/*.md` — 5パターンの業績タイプの入力markdown例
- `samples/_template.md` — 新規執筆用の雛形
- `outputs/4971-mec-valuation.xlsx` — メックの生成済みスプシ
  (FY24実績がyuhoと完全一致、FY25実績で予想を+15.7%上振れ確認)
- `outputs/4112-hodogaya-valuation.xlsx` — 保土谷の骨組みスプシ
- `scripts/build_mec_sheet.py` — メックbuildスクリプト(few-shot として参考に)
- `scripts/build_hodogaya_sheet.py` — 保土谷buildスクリプト

## 一言

**Opusの仕事は「推論+計算+判断」、Pythonの仕事は「型と検証」**。型を破らないよう Python が見張る一方で、数値の中身は徹底的に Opus 自身が判断して埋める。これが「型にはまらず成長ストーリーにベストな数理モデル」を実現する設計思想。
