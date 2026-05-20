# my-claude-project — 日本株投資リサーチ + Takimoto Skills

このレポジトリには **2種類のもの** が入っている:

1. **Takimoto Skills** — 瀧本ゼミ形式の株式投資推奨資料(スライド+バリュエーションシート)をClaude Codeで自動生成する2つのスキル
2. **業界リサーチ + 銘柄スクリーニング成果物** — Claude Code on the Web の各セッションで蓄積した日本株リサーチノート群

スキルだけ使いたい人は下記「インストール」へ。リサーチ成果物だけ読みたい人は `research/`, `screening/`, `heat_matrix/` 配下を直接参照。

## 2つのスキル

| スキル | 出力 | 必要な入力 | 起動フレーズ例 |
|---|---|---|---|
| **`takimoto-slides`** | `.pptx` (Google Slides用) | ストーリーmarkdownのみ | 「○○のスライド作って」「投資推奨スライド作って」 |
| **`takimoto-valuation`** | `.xlsx` (Google Sheets用) | markdown + **決算短信/有報PDF(複数)** | 「○○のスプシ作って」「バリュエーションシート作って」 |

スライドとスプシで必要なものが違う(スライドは markdown 1つ、スプシは PDF を Opus が自分で漁って統合する必要がある)ため、別スキルにしてある。両方欲しい時は両方順に起動。

## インストール

```bash
git clone <this repo>
cd my-claude-project
bash install.sh
```

これで `~/.claude/skills/takimoto-slides/` と `~/.claude/skills/takimoto-valuation/` の両方が配置され、**どのプロジェクトのClaude Codeセッションからも呼び出せる**ようになる。

開発中で書き換えながら使うなら:

```bash
bash install.sh --link    # シンボリックリンク
```

アンインストール:

```bash
bash install.sh --uninstall
```

### 依存

`install.sh` が自動チェックする:

| 依存 | 用途 | インストール |
|---|---|---|
| `poppler-utils` (`pdftotext`) | 決算短信PDFからテキスト抽出 | `brew install poppler` / `apt-get install poppler-utils` |
| `openpyxl` | .xlsx生成 | `pip3 install openpyxl` |
| `python-pptx` | .pptx生成 | `pip3 install python-pptx` |

## 使い方

### スライドだけ欲しい時

```
「メックのスライド作って、入力は samples/4971-mec.md」
```

→ `takimoto-slides` が起動。markdown を読んでスライドを組み立てて `.pptx` 出力。

### スプシだけ欲しい時

```
「メックのスプシ作って、markdownは samples/4971-mec.md、
 決算短信は ~/Downloads/4971-yuhos/ にある」
```

→ `takimoto-valuation` が起動。markdownの因数分解+PDF抽出+FY26+の数値計算+整合性検証で `.xlsx` 出力。

### 両方欲しい時

両スキルを順に呼ぶ(あるいは「○○のスライドとスプシ両方作って」と頼めば Claude が両方起動する)。

## 入力markdownの書き方

`samples/4112-hodogaya.md` を**参照モデル**として参照。最低限必要なセクション:

- **タイトル + 推奨 + 投資期間 + 発表者**
- **0️⃣ 発表サマリー**(ストーリー1行 + 因数分解式 + ミスプライシング表)
- **1️⃣ 基本情報** / **2️⃣ マクロ** / **3️⃣ ミクロ** / **4️⃣ バリュエーション** / **5️⃣ Appendix**

5つの業績ドライバータイプ(samples/に各1本ずつあり)を参考にして自分の銘柄を書く:

| 業績タイプ | サンプル |
|---|---|
| 物量 × 物理量 | `samples/4971-mec.md` |
| 経験年数コホート | `samples/421A-movin.md` |
| セグメント × 粗利率改善 | `samples/6648-kawaden.md` |
| 1人粗利 × 営業人数 | `samples/5537-albalink.md` |
| 持分法利益 | `samples/4112-hodogaya.md` |

新規執筆は `samples/_template.md` を雛形に。

## リポジトリ構造

```
.
├── .claude/skills/                     ← Takimoto Skills (install.sh で配布対象)
│   ├── takimoto-slides/                  スライド生成スキル (.pptx)
│   │   ├── SKILL.md
│   │   ├── reference/                    章立て / 視覚原則 / AskUserQuestion
│   │   └── scripts/pptx_renderer.py      統一レイアウト強制
│   └── takimoto-valuation/               バリュエーションシート生成スキル (.xlsx)
│       ├── SKILL.md
│       ├── reference/                    行構造 / 業績タイプ5種 / AskUserQuestion
│       └── scripts/
│           ├── sheet_renderer.py         スプシスタイル強制
│           ├── extract_yuho.py           決算PDFフォルダ丸ごと抽出
│           └── validate_consistency.py   整合性検証
│
├── samples/                            投資テーゼmarkdown(スキル入力例) — 業績タイプ5種
├── scripts/                            銘柄別buildスクリプト(few-shot参考) — 7銘柄
├── outputs/                            スキル生成済みデモ(.pptx / .xlsx) — 7銘柄
│
├── research/                           業界・銘柄deep dive レポート
│   ├── analog-mature-node-supply-japan.md      アナログ半導体逼迫テーマ
│   ├── companies/                              個別銘柄deep dive (6616/6707/6769/6875)
│   ├── cpo-optical-integration-universe.md     CPO/光電融合 日本ユニバース
│   ├── 5232-soc-ln-modulator-triangulation.md  SOC LNモジュレータ三角測量
│   ├── glass-substrate-packaging-2026.md       ガラス基板パッケージング転換
│   └── power-infrastructure-transformer-macro.md  変圧器/電力インフラマクロ
│
├── screening/                          横断スクリーニング(マクロテーマ → 銘柄リスト)
│   ├── ai-datacenter-stocks-japan.md          AIデータセンター関連 ~120銘柄
│   ├── macro-thematic-research.md             34マクロテーマ × 銘柄評価
│   ├── grid-construction-stocks-japan.md      系統工事業界 銘柄マップ
│   └── grid-construction-primer.md            系統工事業界 体系プライマー
│
├── heat_matrix/                        業界アツさ × 市場注目度マトリクス
│   ├── _index.md                              全業界スナップショット
│   ├── _meta_ranking.md                       横断TOP10銘柄ランキング
│   ├── _growth_screen_skill.md                スクリーニング手法
│   └── heat_*.md (×13)                        業界別deep dive
│
├── install.sh                          ~/.claude/skills/ への配置スクリプト
└── README.md                           本ファイル
```

## リサーチ成果物の使い方

`research/` `screening/` `heat_matrix/` はそれぞれ別セッションで蓄積したリサーチノート。投資テーマ別に横断検索したい時は ripgrep などでgrepするのが早い:

```bash
rg -l "CPO" research/ screening/ heat_matrix/
rg -l "変圧器" .
```

Obsidian にレポジトリごと開いてグラフビューで見るのも有効。

## 動作確認

サンプル銘柄でテストするには:

```bash
# スプシ(yuho 必要 — リポ外なので各自準備、~/Downloads/4971-yuhos/ 等を想定)
python3 scripts/build_mec_sheet.py
python3 scripts/build_kanamic_sheet.py    # Excel関数版(リファクタ済)

# スライド(markdown だけで動く、全銘柄)
for s in mec movin kawaden albalink hodogaya nitta kanamic; do
  python3 scripts/build_${s}_slides.py
done

# 整合性検証
python3 .claude/skills/takimoto-valuation/scripts/validate_consistency.py outputs/4971-mec-valuation.xlsx
```

## デザイン哲学

| 担当 | やること |
|---|---|
| **Python** | スタイル強制(色/インデント/カンマ/グラフ/タイトル位置)、PDFパース、ファイル書出し、整合性検証 |
| **Opus(LLM)** | 因数分解の式評価、FY26+の数値計算、ナラティブ生成、ピア企業の多軸選定、用語解説、ストーリー整合性 |

つまり**「決まったこと」(型)と「判断が要ること」(数学+選定)を分離**。これにより:

- 業績タイプ無限通り対応 (Pythonでハードコードしない)
- 視覚・書式は常に統一感 (Python が機械的に強制)
- 推論が必要な数値はOpusが判断 + 備考で根拠を明記

を両立する。

詳細は各スキルの `SKILL.md` を参照:
- `.claude/skills/takimoto-slides/SKILL.md`
- `.claude/skills/takimoto-valuation/SKILL.md`
