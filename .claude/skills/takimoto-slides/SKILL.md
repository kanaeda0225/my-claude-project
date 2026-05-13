---
name: takimoto-slides
description: 瀧本ゼミ形式の株式投資推奨スライド(.pptx)を生成する。入力はストーリーmarkdownのみ。出力はGoogle Slidesにインポート可能な.pptx。「瀧本ゼミ形式でスライド作って」「投資推奨のスライド作って」「発表資料のスライドだけ欲しい」のような依頼で起動。バリュエーションシート(.xlsx)が欲しい場合はtakimoto-valuationスキルを使う。
---

# Takimoto Slides Generator

瀧本ゼミ形式の株式投資推奨スライドを `.pptx` で生成するスキル。markdown1つで動く(財務PDF不要)。

## 起動条件

- 「瀧本ゼミ形式でスライド作って」「発表資料のスライドを作って」
- 「投資推奨のスライドだけ欲しい」「.pptxにして」
- `/takimoto-slides` を明示的に呼び出した時
- 入力markdownを渡してきて「スライドにして」と依頼

**注**: スプシ(バリュエーションシート)が欲しい場合は **`takimoto-valuation`** スキルを使う。両方欲しい場合は両方起動して順番に走らせる。

## 必要な入力

| 入力 | 必須? | 形式 | 用途 |
| :---- | :---- | :---- | :---- |
| ストーリーmarkdown | **必須** | `samples/*.md`の構造 | スライドの章立て・本文・図表 |
| 発表者名 / 推奨内容 | 任意 | markdown末尾 or AskUserQuestion | 表紙に入る情報 |

## 設計原則

**Python(決定的) と Opus(柔軟) で役割を分ける**:

| 担当 | やること |
| :---- | :---- |
| **Python (`pptx_renderer.py`)** | タイトル位置・フォントサイズ・色・テーブル装飾・章ディバイダー・固定座標 |
| **Opus (あなた)** | markdownを読んでスライド分割の判断、subtitle(1メッセージ)の作文、用語解説の文言、画像プレースホルダの具体指示 |

## 品質第一・多パス検証

**1パスで仕上げない**。

1. **First Pass: 骨組み生成** — markdownを読んで第一稿の `.pptx` を生成
2. **Second Pass: 構造検査** — `python-pptx`で読み返して以下確認:
   - 各スライドのテキスト量 (200字未満)
   - サブタイトル(1メッセージ)が入っているか
   - 画像プレースホルダの具体性
   - 章境目で目次再掲が入っているか
3. **Third Pass: ドキュメント1対1対応の検証** — markdownの主張・章ごとに対応するスライドがあるか
4. **Fourth Pass: トーン整合** — markdownの文体(熱量型/分析型/物語型/体感報告型)がスライドにも反映されているか

## 実行手順

### Step 0: 入力の所在を確認

ユーザーに **ストーリーmarkdown のパス** を確認(明示済みなら省略)。

### Step 1: 環境セットアップ

```bash
python3 -c "import pptx" 2>/dev/null || pip3 install python-pptx
```

### Step 2: markdownの読み解き

ユーザーが渡したmarkdownを読み:
- タイトル(銘柄コード+社名)・推奨・目標株価・投資期間・発表者 (表紙)
- 0️⃣ サマリーセクション(ストーリー1行、因数分解式、ミスプライシング表)
- 1️⃣〜5️⃣ の章構造
- 文体・熱量(BUY!!! 型 / 分析型 / 物語型 / 体感報告型)

参照: `reference/slide_structure.md`(章立て規約)

### Step 3: スライドリストの組み立て

`scripts/pptx_renderer.py` を import:

```python
import sys
sys.path.insert(0, "<skill_dir>/scripts")
from pptx_renderer import Slide, SlideDeckConfig, render

slides = [
    Slide(kind="content", title="ストーリー",
          subtitle="1メッセージ(1行サマリー)",
          bullets=["...", "...", "...", "...", "..."],
          footer="出典"),
    Slide(kind="ascii", title="成長を因数分解する",
          subtitle="...",
          ascii_art="..."),
    Slide(kind="table", title="...", subtitle="...",
          table_headers=[...], table=[[...], [...]]),
    Slide(kind="image_placeholder", title="...", subtitle="...",
          image_placeholder="ここに○○の××グラフ(縦軸: ○○、横軸: ○○)。出典 ○○"),
    # ...
]

deck = SlideDeckConfig(
    title="【XXXX】社名",
    subtitle="推奨：BUY\n目標株価：...\n投資期間：...",
    presenter="○○期 ○○",
    slides=slides,
)
render(deck, "output.pptx")
```

### Step 4: 視覚原則の厳守

`reference/slide_design.md` を読み、以下を必ず守る:

- **フォント最小値**: タイトル28pt / 本文18pt / 表13pt / フッター10pt
- **1スライド1メッセージ**: 各 `content`/`table` スライドに `subtitle` を必ず設定
- **本文制限**: 箇条書き5項目以内、各30字以内、全体180字以内
- **画像プレースホルダ**: 「何の画像/どこから入手/サイズ感」の3点を必ず含める
- **章境目で目次再掲**: 1️⃣→2️⃣ など章遷移ごとに `toc` スライド挿入
- **出典フッター**: markdownに出典記述ある場合はスライドにも反映

### Step 5: スライド構成チェックリスト

生成後、自分で確認:

- [ ] 表紙 (銘柄+推奨+目標株価+投資期間+発表者) は揃っているか
- [ ] 0️⃣ ストーリー要旨・因数分解スライド・ミスプライシング表 がある
- [ ] 目次スライドがある
- [ ] 1️⃣ 基本情報 → 2️⃣ マクロ → 3️⃣ ミクロ → 4️⃣ バリュエーション → 5️⃣ Appendix の順
- [ ] 章境目4箇所 (基本→マクロ→ミクロ→Val→Appx) に目次再掲がある
- [ ] markdown内の表は対応する `table` スライドになっている
- [ ] markdown内の ASCII図は対応する `ascii` スライドになっている
- [ ] markdown内の「○○とは」ブロックは独立した `content` スライドになっている
- [ ] バリュエーションシートが別に存在する場合、本文中で参照している
  (例: 「詳細はバリュエーションシート参照」)

### Step 6: ユーザーへ報告

- 出力ファイルのパス (例: `outputs/XXXX-company-slides.pptx`)
- スライド枚数
- レンダラーが出した警告(あれば)
- 「Google Driveにアップしてプレゼン形式で開いてください」と一言

## やっちゃダメなこと

- **画像生成しない** — `[画像: 具体指示]` プレースホルダで止める。生成は絶対しない
- **markdownにない数値・引用を捏造しない** — 不明なら省略
- **スライドに字を詰め込みすぎない** — 200字超えたら分割するか捨象する
- **フォントを小さくして詰め込まない** — 最小値を絶対に守る
- **subtitle を省略しない** — 各 content/table スライドに必ず1メッセージを置く
- **章境目で目次再掲をスキップしない** — ナビゲーションの肝

## トーン合わせ

入力 markdown の文体を読み取って出力に反映:

| 文体タイプ | 例 | 出力での反映 |
| :---- | :---- | :---- |
| 熱量型 | 「BUY!!!」「STRONG BUY!!」 | エクスクラメーション残す、強気な見出し |
| 分析型 | 「アナリスト0社、構造的転換点」 | 数値・論理中心、感情語は控えめ |
| 物語型 | 「25年保守経営からの覚醒」 | 時系列の語り、変曲点強調 |
| 体感報告型 | 「現地で査定申し込んでみた」 | 体験談・一人称・具体引用 |

## ファイル構造

```
.claude/skills/takimoto-slides/
├── SKILL.md                      ← 本ファイル
├── reference/
│   ├── slide_structure.md        ← 章立て規約
│   ├── slide_design.md           ← 視覚原則(フォント・余白・色)
│   └── asking_user.md            ← AskUserQuestion テンプレ
└── scripts/
    └── pptx_renderer.py          ← スライドマスター的な統一レイアウト強制
```

## 参考例

開発リポジトリ `kanaeda0225/my-claude-project` の以下:

- `samples/*.md` — 5パターンの業績タイプの入力markdown例
- `samples/_template.md` — 新規執筆用の雛形
- `outputs/*-slides.pptx` — 5銘柄の生成済みデモ
- `scripts/build_*_slides.py` — 5銘柄分のbuildスクリプト(few-shot として参考に)
