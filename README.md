# Takimoto Presentation Generator

瀧本ゼミ形式の株式投資推奨資料(スライド+バリュエーションシート)をClaude Codeで自動生成するスキル。

## 何ができるか

入力:
- ストーリーmarkdown(投資テーゼ・因数分解・章立て)
- 決算短信・有価証券報告書のPDF(複数可、フォルダ指定OK)

出力:
- `.pptx` — 統一感のあるスライドデッキ(Google Slidesへインポート可)
- `.xlsx` — 因数分解+P&L+シナリオ別目標株価のバリュエーションシート(Google Sheetsへインポート可)

Python が決まりごと(色・カンマ・インデント・グラフ・タイトル位置の厳密一致)を機械的に強制し、Opus が因数分解の計算・ナラティブ・ピア企業選定など"判断"の部分を担う設計。

## インストール

```bash
git clone <this repo>
cd my-claude-project
bash install.sh
```

これで `~/.claude/skills/takimoto-presentation/` に配置され、**どのプロジェクトのClaude Codeセッションからも呼び出せる**ようになる。

開発中で書き換えながら使うなら:

```bash
bash install.sh --link    # シンボリックリンク
```

アンインストール:

```bash
bash install.sh --uninstall
```

### 依存

スクリプトが自動チェックする:

| 依存 | 用途 | インストール |
|---|---|---|
| `poppler-utils` (`pdftotext`) | 決算短信PDFからテキスト抽出 | `brew install poppler` / `apt-get install poppler-utils` |
| `openpyxl` | .xlsx生成 | `pip3 install openpyxl` |
| `python-pptx` | .pptx生成 | `pip3 install python-pptx` |

## 使い方

Claude Code セッション内で次のどれかを言う:

- `/takimoto-presentation`
- 「瀧本ゼミ形式で資料作って」
- 「この markdown と yuho からスライドとバリュエーションシート作って」
- 「投資推奨のスプシを作って」

スキルが起動して、

1. **入力markdownのパス**を確認
2. **決算PDFの場所**を確認(スプシ生成時)
3. PDFから過去業績を自動抽出
4. ストーリーに沿ってスライド/スプシを構成
5. 推論で埋めた数値や TODO をリストアップして報告

を行う。

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

## リポジトリ構造

```
.
├── .claude/skills/takimoto-presentation/   ← スキル本体(install.shで~/.claude/にコピー)
│   ├── SKILL.md
│   ├── reference/
│   │   ├── slide_structure.md              スライド章立て規約
│   │   ├── slide_design.md                 視覚原則(フォント・余白・色)
│   │   ├── sheet_structure.md              スプシ行構造規約
│   │   ├── factor_decomp_patterns.md       業績タイプ5種のパターン
│   │   └── asking_user.md                  AskUserQuestion テンプレ
│   └── scripts/
│       ├── pptx_renderer.py                スライドマスター的な統一レイアウト強制
│       ├── sheet_renderer.py               スプシスタイル強制
│       ├── extract_yuho.py                 yuho PDF抽出(複数PDF対応)
│       └── validate_consistency.py         整合性検証
│
├── samples/                                 入力markdownの参照例 (5パターン)
├── scripts/                                 ビルドスクリプトの参考実装
├── outputs/                                 生成済みデモ(参考)
├── install.sh                               ~/.claude/skills/ への配置スクリプト
└── README.md                                本ファイル
```

## 動作確認

サンプル銘柄でテストするには:

```bash
# スプシ(MEC、yuho 必要 — リポ外なので各自準備)
python3 scripts/build_mec_sheet.py

# スライド(markdown だけで動く)
python3 scripts/build_mec_slides.py
python3 scripts/build_movin_slides.py
python3 scripts/build_kawaden_slides.py
python3 scripts/build_albalink_slides.py
python3 scripts/build_hodogaya_slides.py

# 整合性検証
python3 .claude/skills/takimoto-presentation/scripts/validate_consistency.py outputs/4971-mec-valuation.xlsx
```

## デザイン哲学

| 担当 | やること |
|---|---|
| **Python** | スタイル強制(色/インデント/カンマ/グラフ/タイトル位置)、PDFパース、ファイル書出し、検証 |
| **Opus(LLM)** | 因数分解の式評価、FY26+の数値計算、ナラティブ生成、ピア企業選定、用語解説、ストーリー整合性 |

つまり**「決まったこと」と「判断が要ること」を分離**。これにより業績タイプ無限通り対応 + 視覚は常に統一感、を両立する。

詳細は `.claude/skills/takimoto-presentation/SKILL.md` を参照。
