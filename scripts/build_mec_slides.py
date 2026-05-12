"""Build the 4971 MEC slide deck (.pptx).

Demo of the takimoto-presentation skill's slide renderer.
Uses samples/4971-mec.md as the narrative source.

This script is what Opus would write while invoking the skill:
each Slide object corresponds to one slide in the output deck.
"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO = Path(__file__).parent.parent
sys.path.insert(0, str(_REPO / ".claude/skills/takimoto-presentation/scripts"))
from pptx_renderer import Slide, SlideDeckConfig, render  # noqa: E402


def build_deck() -> SlideDeckConfig:
    slides: list[Slide] = []

    # ---- 0️⃣ ストーリー / 因数分解 / ミスプライシング ----
    slides.append(
        Slide(
            kind="content",
            title="ストーリー",
            subtitle="GPUパッケージ基板の大判化・多層化を独占",
            bullets=[
                "メックはパッケージ基板の銅箔向け薬剤(CZシリーズ)を製造販売しており、その市場を独占している",
                "GPU向けパッケージ基板は通常基板に比べ層数2-3倍・面積10-15倍、今後2年でさらに面積2倍",
                "1枚あたりの薬剤消費量は面積×層数で増えるためGPU出荷台数の伸び以上の速度で需要爆発",
                "1年で純利益1.4倍、株価1.6倍",
                "よってBUY!!!",
            ],
            footer="メックIR、モルガン・スタンレーレポート 他",
        )
    )

    slides.append(
        Slide(
            kind="ascii",
            title="成長を因数分解する",
            ascii_art=(
                "CZ売上 = 汎用(スマホ・パソコン)向け + GPU用PKG向け\n"
                "\n"
                "GPU用PKG向け = 前年売上 × GPU生産伸び × 面積増加率\n"
                "                            × 層数増加率 × 使用量補正\n"
            ),
            footer="メック IR資料、モルガン・スタンレーレポート",
        )
    )

    slides.append(
        Slide(
            kind="table",
            title="因数分解：各因数の方向",
            table_headers=["因数", "方向", "なぜ動くか"],
            table=[
                ["GPU生産伸び", "↑ 30/20/10%", "AIサーバー需要。電力制約での鈍化織り込み"],
                ["銅箔面積(mm²)", "↑ 4,000→12,000", "Hopper→Blackwell→Rubin で大判化"],
                ["銅箔層数", "↑ 12→18層", "信号・電源密度の上昇でz方向に多層化"],
                ["使用量補正", "50%固定", "GPU向けは多層化のため薄く粗化(汎用の半分)"],
                ["汎用向け", "横ばい(CAGR 0%)", "保守側に倒す"],
            ],
            footer="モルガン・スタンレーレポート 他",
        )
    )

    slides.append(
        Slide(
            kind="content",
            title="なぜ市場が気づいていないか",
            bullets=[
                "業界認識: 「マニアックな化学薬剤メーカー」として認識されている",
                "独占の認識: メック自身と経産省が「市場独占」を明言しているのに広く知られていない",
                "GPU恩恵の織り込み: 通常基板の延長として見られ、「面積×層数」の累乗効果が反映されていない",
                "工場稼働率: 現状43%、新工場で+3割増 — 物理的余地は十分",
            ],
        )
    )

    # ---- 目次 ----
    slides.append(
        Slide(
            kind="toc",
            title="目次",
            bullets=[
                "1️⃣ 基本情報",
                "2️⃣ マクロ",
                "3️⃣ ミクロ",
                "4️⃣ バリュエーション",
                "5️⃣ Appendix",
            ],
        )
    )

    # ---- 1️⃣ 基本情報 ----
    slides.append(Slide(kind="chapter", title="1️⃣ 基本情報"))

    slides.append(
        Slide(
            kind="table",
            title="会社の概要",
            table_headers=["項目", "内容"],
            table=[
                ["主事業", "銅箔表面処理薬剤(密着向上剤・エッチング剤)"],
                ["売上構成", "薬剤事業が99%、機械・資材その他は1%以下"],
                ["主力製品", "CZシリーズ(密着向上剤)が薬剤の7割以上"],
                ["投資期間", "26/12通期決算発表まで(約1年)"],
            ],
        )
    )

    slides.append(
        Slide(
            kind="content",
            title="(補足) CZシリーズとは",
            bullets=[
                "パッケージ基板製造時に必須の薬剤",
                "無いと銅箔と樹脂の界面が剥がれて歩留まりが悪化",
                "「あっても無くてもいい添加剤」ではなく、無いと製品ができない",
                "だからこそ価格交渉力が強い",
            ],
        )
    )

    slides.append(
        Slide(
            kind="ascii",
            title="CZシリーズの生誕ストーリー",
            ascii_art=(
                "1990年代、基板の材料がセラミック → 有機樹脂に切り替わった\n"
                "   ↓\n"
                "「銅箔と樹脂が剥がれる」という新しい問題が浮上\n"
                "   ↓\n"
                "メックが銅の表面を粗化することで密着度を高める手法を提案\n"
                "   ↓\n"
                "その後デファクトスタンダード化した\n"
            ),
        )
    )

    # ---- 2️⃣ マクロ ----
    slides.append(Slide(kind="toc", title="目次(マクロへ)", bullets=["▶ 2️⃣ マクロ"]))
    slides.append(Slide(kind="chapter", title="2️⃣ マクロ"))

    slides.append(
        Slide(
            kind="table",
            title="現時点の基板生産数の比率",
            subtitle="汎用:GPU = 60:1",
            table_headers=["用途", "台数/年", "1台あたり基板数"],
            table=[
                ["スマホ", "12.2億", "10枚"],
                ["パソコン", "2.65億", "15枚"],
                ["AIサーバー", "181万", "120枚"],
            ],
            footer="各種統計より試算",
        )
    )

    slides.append(
        Slide(
            kind="table",
            title="銅箔面積: GPU向けは大判化で爆発",
            subtitle="HopperからRubinまでで面積3倍",
            table_headers=["世代", "インターポーザー面積", "余白込み仕様"],
            table=[
                ["Hopper", "2,831 mm²", "4,000 mm²"],
                ["Blackwell", "4,719 mm²", "7,500 mm²"],
                ["Rubin", "7,722 mm²", "12,000 mm²"],
            ],
            footer="各種レポートで予想されるレンジ",
        )
    )

    slides.append(
        Slide(
            kind="table",
            title="銅箔層数: GPU向けは多層化で増える",
            subtitle="汎用6.2層 vs Rubin 18層(約3倍)",
            table_headers=["世代", "層数"],
            table=[
                ["Hopper", "12"],
                ["Blackwell", "14"],
                ["Rubin", "18"],
            ],
            footer="モルガン・スタンレーレポート",
        )
    )

    slides.append(
        Slide(
            kind="table",
            title="Hopper / Blackwell / Rubin 比率推移",
            subtitle="FY27にRubinが75%を占める — もう一段跳ねる",
            table_headers=["年", "Hopper", "Blackwell", "Rubin"],
            table=[
                ["FY24/12", "50%", "50%", "0%"],
                ["FY25/12", "20%", "80%", "0%"],
                ["FY26/12", "10%", "70%", "25%"],
                ["FY27/12", "5%", "20%", "75%"],
            ],
        )
    )

    # ---- 3️⃣ ミクロ ----
    slides.append(Slide(kind="toc", title="目次(ミクロへ)", bullets=["▶ 3️⃣ ミクロ"]))
    slides.append(Slide(kind="chapter", title="3️⃣ ミクロ"))

    slides.append(
        Slide(
            kind="content",
            title="シェア独占の3つの根拠",
            bullets=[
                "① メック自身と経産省が「独占」を公式に認めている",
                "② ニッチだが致命的に重要な工程 → 顧客側の再評価コスト大",
                "③ PKG基板原価に占める割合は低い → 安さで乗り換えるインセンティブが弱い",
                "→「ニッチで重要で安い」という最強の組み合わせ",
            ],
        )
    )

    slides.append(
        Slide(
            kind="table",
            title="生産キャパ: 物理的な余力がたっぷり",
            subtitle="新工場で 97,200トン → 127,200トン(+31%)",
            table_headers=["年", "生産量(トン)", "工場キャパ(トン)", "稼働率"],
            table=[
                ["FY24/12", "42,075", "97,200", "43.3%"],
                ["FY25/12", "47,153", "97,200", "48.5%"],
                ["FY26/12", "57,084", "97,200", "58.7%"],
                ["FY27/12", "68,379", "127,200", "53.8%"],
            ],
            footer="メックIR",
        )
    )

    # ---- 4️⃣ バリュエーション ----
    slides.append(Slide(kind="toc", title="目次(バリュエーションへ)",
                        bullets=["▶ 4️⃣ バリュエーション"]))
    slides.append(Slide(kind="chapter", title="4️⃣ バリュエーション"))

    slides.append(
        Slide(
            kind="table",
            title="売上・利益のレバレッジ",
            subtitle="売上+51% → 純利益×3.1倍の三段ロケット",
            table_headers=["指標", "FY24", "FY25実", "FY27予"],
            table=[
                ["売上高 (百万円)", "18,234", "20,947", "27,561"],
                ["営業利益 (百万円)", "4,562", "5,748", "9,465"],
                ["営業利益率", "25.0%", "27.4%", "34.3%"],
                ["純利益 (百万円)", "2,291", "5,028", "7,099"],
            ],
            footer="yuho、メックIR、本発表バリュエーション",
        )
    )

    slides.append(
        Slide(
            kind="content",
            title="目標株価とアップサイド",
            bullets=[
                "目標株価: 7,443円",
                "投資期間: 26/12通期決算発表まで(約1年)",
                "純利益 ×3.1倍 + 株価 ×1.6倍を想定",
            ],
        )
    )

    # ---- 5️⃣ Appendix ----
    slides.append(Slide(kind="toc", title="目次(Appendixへ)", bullets=["▶ 5️⃣ Appendix"]))
    slides.append(Slide(kind="chapter", title="5️⃣ Appendix"))

    slides.append(
        Slide(
            kind="content",
            title="投資期間の根拠",
            subtitle="26/12通期決算発表まで(約1年)を想定",
            bullets=[
                "26/12は増収率が停滞するから",
                "26/12は工場稼働による人件費がのり、営利率の伸びが鈍化",
                "半導体マクロに対する不透明感",
            ],
        )
    )

    slides.append(
        Slide(
            kind="content",
            title="ダウンサイド①: 半導体マクロの減速",
            bullets=[
                "GPU需要そのものが急減するシナリオが最大リスク",
                "ただし投資期間が1年なので、その間にAIサーバーマクロが急減速する可能性は低い",
                "→ まぁ大丈夫",
            ],
        )
    )

    slides.append(
        Slide(
            kind="content",
            title="ダウンサイド②: ガラス基板採用 / 無粗化シフト",
            bullets=[
                "技術的には密着向上剤を不要にする方向の動きはありえる",
                "ガラス基板 / 無粗化技術 が普及すればメックの主力製品は不要に",
                "ただし製造現場での実装には時間がかかる",
                "→ まぁ一年以内にはない",
            ],
        )
    )

    slides.append(
        Slide(
            kind="content",
            title="原材料リスクについて",
            bullets=[
                "IR: 「原材料は企業秘密だが、ナフサのような価格変動が激しい材料では無い」",
                "原材料価格変動による粗利率の急激な悪化リスクは低い",
                "粗利率75%(密着向上剤)が原材料起因で大幅に落ちるシナリオは想定不要",
            ],
        )
    )

    return SlideDeckConfig(
        title="【4971】メック",
        subtitle="推奨：BUY\n目標株価：7,443円\n投資期間：26/12通期決算まで",
        presenter="14.0期 中村陽佑",
        slides=slides,
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
