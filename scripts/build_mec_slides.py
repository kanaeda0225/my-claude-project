"""Build the 4971 MEC slide deck (.pptx).

Demo of the takimoto-slides skill's slide renderer.
Source: samples/4971-mec.md (narrative) + financial PDFs (numbers).
"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO = Path(__file__).parent.parent
sys.path.insert(0, str(_REPO / ".claude/skills/takimoto-slides/scripts"))
from pptx_renderer import Slide, SlideDeckConfig, render  # noqa: E402


def build_deck() -> SlideDeckConfig:
    s: list[Slide] = []

    # 0️⃣ ストーリー / 因数分解 / ミスプライシング
    s.append(Slide(
        kind="content", title="ストーリー",
        subtitle="GPUパッケージ基板の大判化・多層化を独占する密着向上剤メーカー",
        bullets=[
            "メックはパッケージ基板の銅箔向け薬剤(CZシリーズ)を製造販売、その市場を独占",
            "GPU向け基板は通常基板に比べ層数2-3倍・面積10-15倍、今後2年でさらに面積2倍",
            "1枚あたり薬剤消費は面積×層数で爆発、GPU出荷台数の伸び以上のスピードで需要拡大",
            "1年で純利益1.4倍、株価1.6倍",
            "よってBUY!!!",
        ],
        footer="メックIR、モルガン・スタンレーレポート、イビデンIR",
    ))

    s.append(Slide(
        kind="ascii", title="成長を因数分解する",
        subtitle="GPU向け = 物量 × 物理量(層数×面積) × 補正",
        ascii_art=(
            "CZ売上 = 汎用(スマホ・パソコン)向け + GPU用PKG向け\n"
            "\n"
            "GPU用PKG向け = 前年売上 × GPU生産伸び × 面積増加率\n"
            "                              × 層数増加率 × 使用量補正\n"
        ),
        footer="メックIR資料、モルガン・スタンレーレポート",
    ))

    s.append(Slide(
        kind="table", title="因数分解：各因数の方向",
        subtitle="面積×層数の累乗効果が市場平均成長を大幅に超える",
        table_headers=["因数", "方向", "なぜ動くか"],
        table=[
            ["GPU生産伸び", "↑ 30/20/10%", "AIサーバー需要、電力制約での鈍化織り込み"],
            ["銅箔面積(mm²)", "↑ 4,000→12,000", "Hopper→Blackwell→Rubinで大判化"],
            ["銅箔層数", "↑ 12→18層", "信号・電源密度上昇でz方向に多層化"],
            ["使用量補正", "50%固定", "GPU向けは薄く粗化(汎用の半分)"],
            ["汎用向け", "横ばい(CAGR 0%)", "保守側に倒す"],
        ],
        footer="モルガン・スタンレーレポート",
    ))

    s.append(Slide(
        kind="content", title="なぜ市場が気づいていないか",
        subtitle="「ニッチな化学品メーカー」認識が累乗効果を見落としている",
        bullets=[
            "業界認識: マニアックな化学薬剤メーカーとして見られている",
            "独占の認識: 経産省も認める独占だが広く知られていない",
            "GPU恩恵: 通常基板の延長として見られ、面積×層数の累乗効果が反映されない",
            "工場稼働率: 現状43%、新工場で+3割増 — 物理的余地は十分",
        ],
    ))

    s.append(Slide(
        kind="toc", title="目次",
        bullets=["1️⃣ 基本情報", "2️⃣ マクロ", "3️⃣ ミクロ",
                 "4️⃣ バリュエーション", "5️⃣ Appendix"],
    ))

    # 1️⃣ 基本情報
    s.append(Slide(kind="chapter", title="1️⃣ 基本情報"))

    s.append(Slide(
        kind="table", title="会社の概要",
        subtitle="薬剤事業が99%、CZシリーズが薬剤の7割以上",
        table_headers=["項目", "内容"],
        table=[
            ["銘柄コード", "4971 (東証プライム)"],
            ["主事業", "銅箔表面処理薬剤(密着向上剤・エッチング剤)"],
            ["売上構成", "薬剤事業 99% / 機械・資材その他 1%以下"],
            ["主力製品", "CZシリーズ(密着向上剤)が薬剤の7割以上"],
            ["投資期間", "26/12通期決算発表まで(約1年)"],
        ],
    ))

    s.append(Slide(
        kind="content", title="(補足) CZシリーズとは",
        subtitle="パッケージ基板製造に必須、無いと歩留まり悪化",
        bullets=[
            "パッケージ基板の銅箔と樹脂の界面を密着させる薬剤",
            "無いと端がめくれて製品が成立しない(添加剤ではなく必須材料)",
            "だからこそ価格交渉力が強い",
        ],
    ))

    s.append(Slide(
        kind="ascii", title="CZシリーズの生誕ストーリー",
        subtitle="業界の根本問題を解決した会社=規格を握るポジション",
        ascii_art=(
            "1990年代、基板の材料が セラミック → 有機樹脂に切り替わった\n"
            "                              ↓\n"
            "         「銅箔と樹脂が剥がれる」という新しい問題が浮上\n"
            "                              ↓\n"
            "     メックが銅の表面を粗化することで密着度を高める手法を提案\n"
            "                              ↓\n"
            "             その後デファクトスタンダード化した\n"
        ),
    ))

    s.append(Slide(
        kind="content", title="半導体パッケージ基板とは",
        subtitle="半導体とPCB基板をつなぐ中継機",
        bullets=[
            "半導体(CPU/GPU/HBM)とマザーボードの間に挟まれる基板",
            "パッケージ基板メーカー: イビデン、新光電気工業など",
            "メックはこの基板メーカーに薬剤(CZシリーズ)を販売",
            "イビデンは過去IRに登場した納入先、現在生産量大幅増強 = 直接的追い風",
        ],
        footer="メックIR、イビデンIR",
    ))

    s.append(Slide(
        kind="ascii", title="今後パッケージ基板は大判化と多層化が並行で進む",
        subtitle="x-y方向は規格固定 → z方向(層数)に増やすしかない",
        ascii_art=(
            "         [大判化]                       [多層化]\n"
            "従来:  [SoC][HBM×2]      ━━━━━>   ビルドアップ層\n"
            "                                       コア層\n"
            "これから: [SoC大][HBM×4-8]              ビルドアップ層\n"
            "                                          (層数 増加)\n"
        ),
        footer="14.0蔦野の日東紡資料より引用",
    ))

    # 2️⃣ マクロ
    s.append(Slide(kind="toc", title="目次(マクロへ)", bullets=["▶ 2️⃣ マクロ"]))
    s.append(Slide(kind="chapter", title="2️⃣ マクロ"))

    s.append(Slide(
        kind="table", title="GPU向け基板生産数(YoY)",
        subtitle="汎用は据置(保守)、GPU向けは+30/20/10%",
        table_headers=["世代", "汎用向け CAGR", "AI用途GPU向け YoY"],
        table=[
            ["FY25", "0% (バリュエーション保守)", "+30%"],
            ["FY26", "0%", "+20%"],
            ["FY27", "0%", "+10%"],
        ],
        footer="各種半導体レポート、電力制約による鈍化を織り込み",
    ))

    s.append(Slide(
        kind="ascii", title="現時点の基板生産数の比率",
        subtitle="汎用:GPU = 60:1 (24/12時点) → 27/12には完全に主役交代",
        ascii_art=(
            "生産台数/年\n"
            " ① スマホ：12.2億    ② パソコン：2.65億    ③ AIサーバー：181万\n"
            "\n"
            "パッケージ基板数(1台あたり)\n"
            " ① スマホ：10        ② パソコン：15        ③ AIサーバー：120\n"
            "\n"
            " ↓\n"
            "\n"
            " 汎用(スマホ・パソコン)向け：GPU向け  =  60 : 1\n"
        ),
    ))

    s.append(Slide(
        kind="table", title="銅箔面積: GPU向けは大判化で爆発",
        subtitle="HopperからRubinまでで面積3倍 = 薬剤消費も3倍効く",
        table_headers=["世代", "インターポーザー面積", "余白込み仕様"],
        table=[
            ["Hopper", "2,831 mm²", "4,000 mm²"],
            ["Blackwell", "4,719 mm²", "7,500 mm²"],
            ["Rubin", "7,722 mm²", "12,000 mm²"],
        ],
        footer="各種レポートで予想されるレンジ",
    ))

    s.append(Slide(
        kind="table", title="銅箔層数: GPU向けは多層化で増える",
        subtitle="汎用6.2層 vs Rubin 18層 = 約3倍の差",
        table_headers=["世代", "層数"],
        table=[
            ["汎用(スマホ・パソコン)加重平均", "6.2"],
            ["Hopper", "12"],
            ["Blackwell", "14"],
            ["Rubin", "18"],
        ],
        footer="モルガン・スタンレーレポート",
    ))

    s.append(Slide(
        kind="table", title="Hopper / Blackwell / Rubin 比率推移",
        subtitle="FY27にRubinが75%占有 — もう一段薬剤需要が跳ねる",
        table_headers=["年", "Hopper", "Blackwell", "Rubin"],
        table=[
            ["FY24/12", "50.0%", "50.0%", "0.0%"],
            ["FY25/12", "20.0%", "80.0%", "0.0%"],
            ["FY26/12", "10.0%", "70.0%", "25.0%"],
            ["FY27/12", "5.0%", "20.0%", "75.0%"],
        ],
        footer="モルガン・スタンレーレポート、NVIDIA公表ロードマップ",
    ))

    s.append(Slide(
        kind="content", title="使用量補正：GPU向けは50%",
        subtitle="多層化のため薄く粗化 → 汎用の半分しか消費しない",
        bullets=[
            "汎用向け: CZ8101 (通常の粗化量)",
            "GPU向け: CZ8201 (薄く粗化する専用品)",
            "GPU向けは多層化のため可能な限り薄く粗化する必要",
            "→ 使用量は汎用比 50% で固定",
            "面積と層数の伸びの一部を相殺するが、最終売上は3年で4倍以上",
        ],
    ))

    # 3️⃣ ミクロ
    s.append(Slide(kind="toc", title="目次(ミクロへ)", bullets=["▶ 3️⃣ ミクロ"]))
    s.append(Slide(kind="chapter", title="3️⃣ ミクロ"))

    s.append(Slide(
        kind="content", title="シェア独占の3つの根拠",
        subtitle="「ニッチで重要で安い」という最強の組み合わせ",
        bullets=[
            "① メック自身と経産省が「独占」を公式に認めている",
            "② ニッチだが致命的に重要 → 顧客の再評価コスト(時間+金)が大きい",
            "③ PKG基板原価に占める割合は低い → 安さで乗り換えるインセンティブ弱い",
        ],
        footer="経産省資料、メックIR",
    ))

    s.append(Slide(
        kind="table", title="粗利率の高さ: 密着向上剤75% / エッチング剤45%",
        subtitle="GPU向けCZ比率が上がるほど全社粗利率が押し上がる",
        table_headers=["製品", "粗利率"],
        table=[
            ["密着向上剤 (CZシリーズ)", "75%"],
            ["エッチング剤", "45%"],
        ],
        footer="四半期データから方程式立てて推計",
    ))

    s.append(Slide(
        kind="table", title="生産キャパ: 物理的な余力がたっぷり",
        subtitle="新工場で 97,200 → 127,200 トン(+31%増)",
        table_headers=["年", "生産量(トン)", "工場キャパ(トン)", "稼働率"],
        table=[
            ["FY24/12", "42,075", "97,200", "43.3%"],
            ["FY25/12", "47,153", "97,200", "48.5%"],
            ["FY26/12", "57,084", "97,200", "58.7%"],
            ["FY27/12", "68,379", "127,200", "53.8%"],
        ],
        footer="メックIR(max稼働率7-8割の供給責任あり)",
    ))

    # 4️⃣ バリュエーション
    s.append(Slide(kind="toc", title="目次(バリュエーションへ)",
                   bullets=["▶ 4️⃣ バリュエーション"]))
    s.append(Slide(kind="chapter", title="4️⃣ バリュエーション"))

    s.append(Slide(
        kind="table", title="売上・利益のレバレッジ",
        subtitle="売上+51% × 利益率改善 × 税率正常化 = 純利益×3.1倍",
        table_headers=["指標", "FY24 実", "FY25 実", "FY27 予"],
        table=[
            ["売上高 (百万円)", "18,234", "20,948", "27,561"],
            ["営業利益 (百万円)", "4,562", "5,748", "9,465"],
            ["営業利益率", "25.0%", "27.4%", "34.3%"],
            ["純利益 (百万円)", "2,291", "5,028", "7,099"],
        ],
        footer="yuho開示、本発表バリュエーション",
    ))

    s.append(Slide(
        kind="content", title="目標株価とアップサイド",
        subtitle="1年で純利益×3.1倍 + 株価×1.6倍を想定",
        bullets=[
            "目標株価: 7,443円",
            "投資期間: 26/12通期決算発表まで(約1年)",
            "売上成長 × 営業レバレッジ × 税率正常化の三段ロケット",
        ],
    ))

    s.append(Slide(
        kind="image_placeholder",
        title="営業利益と利益率の推移",
        subtitle="営利+107% / 営利率25% → 34% の急改善を視覚化",
        image_placeholder=(
            "営業利益(棒、左軸、百万円)と営業利益率(線+ドット、右軸、%)の二軸グラフ。\n"
            "横軸: FY21実 / FY22実 / FY23実 / FY24実 / FY25実 / FY26予 / FY27予。\n"
            "サイズ目安: スライド中央に 横6inch × 縦4inch 程度。\n"
            "ソース: outputs/4971-mec-valuation.xlsx のチャート1 をエクスポート"
        ),
        image_caption="目視で2025-2027の急上昇トレンドが分かるように",
        footer="outputs/4971-mec-valuation.xlsx",
    ))

    # 5️⃣ Appendix
    s.append(Slide(kind="toc", title="目次(Appendixへ)", bullets=["▶ 5️⃣ Appendix"]))
    s.append(Slide(kind="chapter", title="5️⃣ Appendix"))

    s.append(Slide(
        kind="content", title="投資期間の根拠",
        subtitle="ストーリーが市場に伝わって株価が反応するまで1年",
        bullets=[
            "26/12は増収率が停滞するから",
            "26/12は工場稼働による人件費がのり、営利率の伸びが鈍化",
            "半導体マクロに対する不透明感",
        ],
    ))

    s.append(Slide(
        kind="content", title="ダウンサイド①: 半導体マクロの減速",
        subtitle="投資期間1年では発生確率低い",
        bullets=[
            "GPU需要そのものが急減するシナリオが最大リスク",
            "ただし投資期間が1年 → その間に半導体マクロが急減速する確率は低い",
            "→ まぁ大丈夫",
        ],
    ))

    s.append(Slide(
        kind="content",
        title="ダウンサイド②: ガラス基板採用 / 無粗化シフト",
        subtitle="技術的にはあり得るが製造現場の実装は時間がかかる",
        bullets=[
            "ガラス基板 / 無粗化技術 が普及すればCZシリーズは不要に",
            "ただし量産現場での実装には時間がかかる",
            "→ まぁ一年以内にはない",
        ],
    ))

    s.append(Slide(
        kind="content", title="原材料リスクについて",
        subtitle="ナフサのような変動激しい材料ではない (IR)",
        bullets=[
            "IR: 「原材料は企業秘密だが、ナフサのような価格変動が激しい材料では無い」",
            "原材料起因で粗利率75%(密着向上剤)が大幅低下するシナリオは想定不要",
        ],
        footer="メックIR電話",
    ))

    return SlideDeckConfig(
        title="【4971】メック",
        subtitle="推奨：BUY\n目標株価：7,443円\n投資期間：26/12通期決算まで",
        presenter="14.0期 中村陽佑",
        slides=s,
    )


def main():
    config = build_deck()
    out = Path(__file__).parent.parent / "outputs" / "4971-mec-slides.pptx"
    out.parent.mkdir(parents=True, exist_ok=True)
    warnings = render(config, str(out))
    print(f"\nWrote: {out}")
    print(f"Slides: {len(config.slides) + 1} (including cover)")
    print(f"Warnings: {len(warnings)}")


if __name__ == "__main__":
    main()
