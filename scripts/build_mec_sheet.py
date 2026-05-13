"""Build the 4971 MEC valuation .xlsx.

Data sources:
- Historical (FY21-FY25): /tmp/mec-financials/download/4971_yuho_101_*.pdf
  (extracted via pdftotext; see notes in samples/4971-mec.md)
- Forecast assumptions (FY26-FY27): samples/4971-mec.md
  (factor decomposition, growth rates, margins)
- Cross-validated against the original MEC CSV (existing benchmark)
"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO = Path(__file__).parent.parent
sys.path.insert(0, str(_REPO / ".claude/skills/takimoto-valuation/scripts"))
from sheet_renderer import ChartSpec, Row, SheetConfig, render  # noqa: E402


# ---------------------------------------------------------------------------
# Historical data from 4971_yuho_101_20251231.pdf (FY25 yuho, filed 2026/3/23)
# All values in 百万円 unless noted.
# ---------------------------------------------------------------------------
# FY24 / FY25 actuals (from 連結損益計算書)
ACT = {
    "FY24": {
        "売上高": 18_234,
        "売上原価": 7_132,
        "売上総利益": 11_101,
        "販管費": 6_539,
        "営業利益": 4_562,
        "経常利益": 4_682,
        "純利益": 2_291,
        # SG&A breakdown
        "貸倒引当金繰入額": 1,
        "給料及び賞与": 1_858,
        "荷造運搬費": 780,
        "賞与引当金繰入額": 231,
        "役員賞与引当金繰入額": 50,
        "株式報酬引当金繰入額": 24,
        "退職給付費用": 103,
        "研究開発費": 1_334,
        # Product split
        "薬品売上": 17_478,
        "機械売上": 578,
        "資材売上": 169,
        "その他売上": 8,
        # Share count
        "発行済株式数_期末_千株": 20_071,
        "EPS": 122.38,
        "減価償却費": 816,  # 連結 全体
    },
    "FY25": {
        "売上高": 20_947,
        "売上原価": 7_970,
        "売上総利益": 12_977,
        "販管費": 7_229,
        "営業利益": 5_748,
        "経常利益": 6_051,
        "純利益": 5_028,
        "貸倒引当金繰入額": 2,
        "給料及び賞与": 1_981,
        "荷造運搬費": 872,
        "賞与引当金繰入額": 266,
        "役員賞与引当金繰入額": 56,
        "株式報酬引当金繰入額": 55,
        "退職給付費用": 88,
        "研究開発費": 1_379,
        "薬品売上": 20_211,
        "機械売上": 312,
        "資材売上": 403,
        "その他売上": 19,
        "発行済株式数_期末_千株": 19_571,  # 自己株消却 △500千株 (2025/8/29)
        "EPS": 272.14,
        "減価償却費": 823,
    },
}


# ---------------------------------------------------------------------------
# FY26 / FY27 projections (from samples/4971-mec.md)
# ---------------------------------------------------------------------------
# CZ売上の汎用 vs GPU 分解 (markdownの因数分解を CSV と同じ単位で再現)
# 24/12時点で汎用基板生産数=60, GPU向け=1 (相対値)
GPU_GROWTH = {"FY25": 1.30, "FY26": 1.30 * 1.20, "FY27": 1.30 * 1.20 * 1.10}
LAYER_GPU = {"FY24": 13, "FY25": 14, "FY26": 16, "FY27": 17}
AREA_GPU = {"FY24": 5_750, "FY25": 6_800, "FY26": 8_650, "FY27": 10_700}
USAGE_FACTOR = 0.50  # GPU向けは汎用の50%粗化

# Helper: CZ向け売上の理論計算 (markdown 仮定どおり)
def cz_general(base_fy24: int, growth: float = 1.0) -> int:
    """汎用向け基板=60 据え置き × 6.2層 × 330mm² (markdown値)"""
    return int(base_fy24 * growth)


# Use the existing CSV's CZ split as the FY24 baseline reference
CZ_FY24_TOTAL = 12_699  # 百万円 (existing CSV)
CZ_GENERAL_FY24 = 12_699 - 3_737  # 汎用向け = 全体 - GPU向け = 8,962
# Wait — existing CSV showed 汎用 122,760 / GPU 37,375 (これは銅箔需要相対値)
# Let me recalculate using markdown logic
# CSV のCZ売上は 12,699 (FY24) → 22,039 (FY27)
# このうちFY24 GPU向け = 37,375 / 160,135 = 23.3%
# 汎用向け = 76.7% = 9,734
CZ_FY24_GENERAL = round(12_699 * 0.767)
CZ_FY24_GPU = 12_699 - CZ_FY24_GENERAL


def build_valuation_sheet() -> SheetConfig:
    rows: list[Row] = []

    years = ["FY24/12 実", "FY25/12 実", "FY26/12 予", "FY27/12 予"]

    # ===========================
    # 売上高
    # ===========================
    rows.append(Row(label="【売上高】", is_section_header=True))

    rows.append(
        Row(
            label="売上高 合計",
            indent=0,
            values=[
                ACT["FY24"]["売上高"],
                ACT["FY25"]["売上高"],
                "TODO: 計算",
                "TODO: 計算",
            ],
            note="連結損益計算書より。FY26/FY27は薬品+機械+資材+その他の合計",
        )
    )

    rows.append(
        Row(
            label="増収率",
            indent=1,
            values=[None, 0.149, 0.124, 0.144],
            note="(売上高 ÷ 前年売上高) − 1。FY25実績14.9%、FY26-27は推計",
            number_format="0.0%",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # セグメント別売上
    # ===========================
    rows.append(Row(label="【セグメント別売上】", is_section_header=True))

    # 薬品事業
    rows.append(
        Row(
            label="薬品事業",
            indent=0,
            values=[
                ACT["FY24"]["薬品売上"],
                ACT["FY25"]["薬品売上"],
                22_800,  # markdown projection
                26_811,  # markdown projection
            ],
            note="yuho開示。FY25実績は前期比+15.6%増。FY26-27は密着向上剤(CZ)の伸びとエッチング剤据置を反映",
        )
    )
    rows.append(
        Row(
            label="密着向上剤 (CZシリーズ)",
            indent=1,
            values=[
                CZ_FY24_GPU + CZ_FY24_GENERAL,
                "TODO",  # FY25実 (yuho では未開示)
                18_028,
                22_039,
            ],
            is_independent=True,
            importance="★★★",
            note="薬品売上の主力(7割以上)。汎用向け+GPU向けの和。GPU向けは Markdown 因数分解で算出",
        )
    )
    rows.append(
        Row(
            label="汎用向け (スマホ/パソコン)",
            indent=2,
            values=[CZ_FY24_GENERAL, CZ_FY24_GENERAL, CZ_FY24_GENERAL, CZ_FY24_GENERAL],
            is_independent=True,
            note="markdown：成熟市場で CAGR 0% (保守)。FY24実績ベースで据え置き",
        )
    )
    rows.append(
        Row(
            label="GPU向け",
            indent=2,
            values=[CZ_FY24_GPU, "TODO: 約4,860", 5_700, 8_300],
            is_independent=True,
            importance="★★★",
            note="markdown因数分解：生産伸び+30/20/10% × 面積1.18/1.27/1.24倍 × 層数1.08/1.14/1.06倍 × 使用量補正50%",
        )
    )
    rows.append(
        Row(
            label="エッチング剤",
            indent=1,
            values=[4_000, 4_000, 4_000, 4_000],
            is_independent=True,
            note="markdown：据え置き。粗利率45%",
        )
    )
    rows.append(
        Row(
            label="薬品事業 その他",
            indent=1,
            values=[772, "TODO", 772, 772],
            is_independent=True,
            note="markdown：据え置き",
        )
    )

    rows.append(
        Row(
            label="機械事業",
            indent=0,
            values=[
                ACT["FY24"]["機械売上"],
                ACT["FY25"]["機械売上"],
                "TODO",
                "TODO",
            ],
            is_independent=True,
            note="yuho開示。FY25は前期比-46.1%。汎用機の売上はボラあり",
        )
    )
    rows.append(
        Row(
            label="資材事業",
            indent=0,
            values=[
                ACT["FY24"]["資材売上"],
                ACT["FY25"]["資材売上"],
                "TODO",
                "TODO",
            ],
            is_independent=True,
            note="yuho開示。FY25は前期比+138.8%",
        )
    )
    rows.append(
        Row(
            label="その他事業",
            indent=0,
            values=[
                ACT["FY24"]["その他売上"],
                ACT["FY25"]["その他売上"],
                "TODO",
                "TODO",
            ],
            is_independent=True,
            note="yuho開示",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 因数分解 (GPU向けCZシリーズ)
    # ===========================
    rows.append(Row(label="【GPU向けCZシリーズ 因数分解】", is_section_header=True))

    rows.append(
        Row(
            label="GPU向け基板生産数 (相対値)",
            indent=0,
            values=[1.00, 1.30, 1.56, 1.72],
            is_independent=True,
            importance="★★",
            note="markdown：FY25+30%, FY26+20%, FY27+10%。電力制約での鈍化を保守的に織り込み",
            number_format="0.00",
        )
    )
    rows.append(
        Row(
            label="基板生産数 YoY",
            indent=1,
            values=[None, 0.30, 0.20, 0.10],
            note="増加率の年度別表示 (markdown)",
            number_format="0.0%",
        )
    )

    rows.append(
        Row(
            label="銅箔層数 (加重平均)",
            indent=0,
            values=[
                LAYER_GPU["FY24"],
                LAYER_GPU["FY25"],
                LAYER_GPU["FY26"],
                LAYER_GPU["FY27"],
            ],
            is_independent=True,
            importance="★★",
            note="markdown：Hopper12 / Blackwell14 / Rubin18 を世代ミックスで加重平均 (モルガン・スタンレーレポート)",
            number_format="0",
        )
    )

    rows.append(
        Row(
            label="銅箔面積 (mm²)",
            indent=0,
            values=[
                AREA_GPU["FY24"],
                AREA_GPU["FY25"],
                AREA_GPU["FY26"],
                AREA_GPU["FY27"],
            ],
            is_independent=True,
            importance="★★★",
            note="markdown：Hopper4,000 / Blackwell7,500 / Rubin12,000 mm²を世代ミックスで加重平均",
        )
    )

    rows.append(
        Row(
            label="使用量補正",
            indent=0,
            values=[USAGE_FACTOR, USAGE_FACTOR, USAGE_FACTOR, USAGE_FACTOR],
            is_independent=True,
            note="markdown：GPU向けは多層化のため薄く粗化 → 汎用の半分(50%)",
            number_format="0.0%",
        )
    )

    # GPU/Blackwell世代比率
    rows.append(Row(label="  ▽ NVIDIA世代別比率", indent=0))
    rows.append(
        Row(
            label="Hopper",
            indent=1,
            values=[0.50, 0.20, 0.10, 0.05],
            is_independent=True,
            note="markdown : H100/H200世代",
            number_format="0.0%",
        )
    )
    rows.append(
        Row(
            label="Blackwell",
            indent=1,
            values=[0.50, 0.80, 0.70, 0.20],
            is_independent=True,
            note="markdown : B100/B200世代",
            number_format="0.0%",
        )
    )
    rows.append(
        Row(
            label="Rubin",
            indent=1,
            values=[0.00, 0.00, 0.25, 0.75],
            is_independent=True,
            note="markdown : 次世代。FY27に主役入れ替わり",
            number_format="0.0%",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 粗利
    # ===========================
    rows.append(Row(label="【粗利】", is_section_header=True))

    rows.append(
        Row(
            label="売上原価",
            indent=0,
            values=[
                ACT["FY24"]["売上原価"],
                ACT["FY25"]["売上原価"],
                "TODO",
                "TODO",
            ],
            note="yuho開示。FY26-FY27は売上×(1-粗利率)",
        )
    )
    rows.append(
        Row(
            label="売上総利益",
            indent=0,
            values=[
                ACT["FY24"]["売上総利益"],
                ACT["FY25"]["売上総利益"],
                "TODO: 15,321",
                "TODO: 18,329",
            ],
            note="yuho開示。markdownでは FY26 15,321 / FY27 18,329 を見込む",
        )
    )
    rows.append(
        Row(
            label="粗利率",
            indent=1,
            values=[
                ACT["FY24"]["売上総利益"] / ACT["FY24"]["売上高"],
                ACT["FY25"]["売上総利益"] / ACT["FY25"]["売上高"],
                0.651,
                0.665,
            ],
            note="FY24実 60.9% → FY25実 62.0% → FY26-27 で密着向上剤比率上昇により改善",
            number_format="0.0%",
        )
    )

    rows.append(
        Row(
            label="密着向上剤 粗利率",
            indent=1,
            values=[0.75, 0.75, 0.75, 0.75],
            is_independent=True,
            note="markdown：四半期ごとの方程式から推計。75%",
            number_format="0.0%",
        )
    )
    rows.append(
        Row(
            label="エッチング剤 粗利率",
            indent=1,
            values=[0.45, 0.45, 0.45, 0.45],
            is_independent=True,
            note="markdown：四半期ごとの方程式から推計。45%",
            number_format="0.0%",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 販管費
    # ===========================
    rows.append(Row(label="【販管費】", is_section_header=True))

    rows.append(
        Row(
            label="販管費 合計",
            indent=0,
            values=[
                ACT["FY24"]["販管費"],
                ACT["FY25"]["販管費"],
                "TODO: 7,553",
                "TODO: 8,864",
            ],
            note="yuho開示。FY26/FY27はmarkdown",
        )
    )

    sga_items = [
        ("給料及び賞与", "27/12は工場稼働分込み (markdown)"),
        ("荷造運搬費", "売上に比例 (markdown)"),
        ("貸倒引当金繰入額", "据え置き (markdown)"),
        ("賞与引当金繰入額", "据え置き (markdown)"),
        ("役員賞与引当金繰入額", "据え置き (markdown)"),
        ("株式報酬引当金繰入額", "据え置き (markdown)。FY25実は予想+131%超過"),
        ("退職給付費用", "据え置き (markdown)"),
        ("研究開発費", "売上の約10%目標 (markdown)。FY25実は予想を下回り"),
    ]
    for item, note in sga_items:
        rows.append(
            Row(
                label=item,
                indent=1,
                values=[
                    ACT["FY24"][item],
                    ACT["FY25"][item],
                    "TODO",
                    "TODO",
                ],
                is_independent=True,
                note=note,
            )
        )

    rows.append(
        Row(
            label="減価償却費 (新工場分)",
            indent=1,
            values=[0, 0, 0, 225],
            is_independent=True,
            note="markdown：IR資料より。新工場(2026年12月本格稼働)の減価償却は2~2.5億/年",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 営業利益・経常・純利益
    # ===========================
    rows.append(Row(label="【利益】", is_section_header=True))

    rows.append(
        Row(
            label="営業利益",
            indent=0,
            values=[
                ACT["FY24"]["営業利益"],
                ACT["FY25"]["営業利益"],
                "TODO: 7,769",
                "TODO: 9,465",
            ],
            note="yuho開示。markdownでは FY26 7,769 / FY27 9,465 を見込む",
        )
    )
    rows.append(
        Row(
            label="営業利益率",
            indent=1,
            values=[
                ACT["FY24"]["営業利益"] / ACT["FY24"]["売上高"],
                ACT["FY25"]["営業利益"] / ACT["FY25"]["売上高"],
                0.330,
                0.343,
            ],
            note="FY24実 25.0% → FY25実 27.4% → markdownで33.0% → 34.3% へ",
            number_format="0.0%",
        )
    )

    rows.append(
        Row(
            label="経常利益",
            indent=0,
            values=[
                ACT["FY24"]["経常利益"],
                ACT["FY25"]["経常利益"],
                "TODO",
                "TODO",
            ],
            note="yuho開示。営業外損益は小幅",
        )
    )

    rows.append(
        Row(
            label="親会社株主帰属純利益",
            indent=0,
            values=[
                ACT["FY24"]["純利益"],
                ACT["FY25"]["純利益"],
                "TODO: 5,826",
                "TODO: 7,099",
            ],
            note="yuho開示。markdownでは FY26 5,826 / FY27 7,099。FY24は税効果会計の影響で実効税率49.8%だった",
        )
    )
    rows.append(
        Row(
            label="実効税率",
            indent=1,
            values=[0.498, 0.223, 0.25, 0.25],
            is_independent=True,
            note="FY24は税効果関係で異常値49.8%。FY25実 22.3% → markdownはFY26-27で25.0%と置く",
            number_format="0.0%",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 工場稼働率
    # ===========================
    rows.append(Row(label="【工場稼働率】", is_section_header=True))

    rows.append(
        Row(
            label="工場稼働率",
            indent=0,
            values=[0.433, 0.485, 0.587, 0.538],
            note="生産量 ÷ 工場キャパ。供給責任ありmax 7-8割。FY27に新工場稼働で分母が増えるため低下",
            number_format="0.0%",
        )
    )
    rows.append(
        Row(
            label="生産量 (トン)",
            indent=1,
            values=[42_075, 47_153, 57_084, 68_379],
            is_independent=True,
            note="markdown：CZ生産量。GPU向け薬剤量に概ね連動",
        )
    )
    rows.append(
        Row(
            label="工場キャパ (トン)",
            indent=1,
            values=[97_200, 97_200, 97_200, 127_200],
            is_independent=True,
            importance="★",
            note="markdown：26/12新工場で +3割増 (127,200トン)。供給責任ありmax 7-8割",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # EPS & PER & 目標株価
    # ===========================
    rows.append(Row(label="【株価・PER感応度】", is_section_header=True))

    rows.append(
        Row(
            label="発行済株式数 (千株, 期末)",
            indent=0,
            values=[
                ACT["FY24"]["発行済株式数_期末_千株"],
                ACT["FY25"]["発行済株式数_期末_千株"],
                ACT["FY25"]["発行済株式数_期末_千株"],
                ACT["FY25"]["発行済株式数_期末_千株"],
            ],
            is_independent=True,
            note="yuho開示。FY25/8/29に自己株式消却 △500千株。FY26-FY27は据え置き想定",
        )
    )
    rows.append(
        Row(
            label="EPS (円)",
            indent=0,
            values=[
                ACT["FY24"]["EPS"],
                ACT["FY25"]["EPS"],
                "TODO",
                "TODO",
            ],
            note="yuho開示。FY26-FY27は純利益 ÷ 発行済株式数",
            number_format="0.00",
        )
    )

    rows.append(
        Row(
            label="想定PER (ベスト)",
            indent=1,
            values=["TODO: 30?", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="同業比較で設定。半導体製造装置・電子材料平均は20-30倍レンジ",
            number_format="0.0",
        )
    )
    rows.append(
        Row(
            label="想定PER (ノーマル)",
            indent=1,
            values=["TODO: 27.5?", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="現状PER水準",
            number_format="0.0",
        )
    )
    rows.append(
        Row(
            label="想定PER (ワースト)",
            indent=1,
            values=["TODO: 25?", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="ダウンサイドシナリオ",
            number_format="0.0",
        )
    )

    rows.append(
        Row(
            label="目標株価 ベスト (円)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="EPS × ベストPER",
        )
    )
    rows.append(
        Row(
            label="目標株価 ノーマル (円)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="EPS × ノーマルPER",
        )
    )
    rows.append(
        Row(
            label="目標株価 ワースト (円)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="EPS × ワーストPER",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # ピア比較
    # ===========================
    rows.append(Row(label="【ピア比較 (PER)】", is_section_header=True))

    rows.append(
        Row(
            label="メック (4971) 現状PER",
            indent=0,
            values=["TODO: 現値", "", "", ""],
            is_independent=True,
            note="現株価 / EPS(直近実績) で算出",
            number_format="0.0",
        )
    )
    rows.append(
        Row(
            label="ピア① 半導体材料 (信越化学、JSR等)",
            indent=1,
            values=["TODO", "", "", ""],
            is_independent=True,
            note="多軸選定 (業態：半導体材料)",
            number_format="0.0",
        )
    )
    rows.append(
        Row(
            label="ピア② パッケージ基板メーカー (イビデン等)",
            indent=1,
            values=["TODO", "", "", ""],
            is_independent=True,
            note="多軸選定 (バリューチェーン上の隣接企業)",
            number_format="0.0",
        )
    )
    rows.append(
        Row(
            label="ピア③ 化学品ニッチ独占企業",
            indent=1,
            values=["TODO", "", "", ""],
            is_independent=True,
            note="多軸選定 (ビジネスモデル：ニッチ独占)",
            number_format="0.0",
        )
    )

    # ===========================
    # Charts
    # ===========================
    charts = [
        ChartSpec(
            title="売上・粗利・営利",
            bar_rows=["売上高 合計", "売上総利益", "営業利益"],
            line_rows=["営業利益率"],
            anchor="J2",
        ),
        ChartSpec(
            title="GPU向けCZシリーズの伸び",
            bar_rows=["GPU向け"],
            line_rows=["GPU向け基板生産数 (相対値)"],
            anchor="J22",
        ),
        ChartSpec(
            title="工場稼働率と生産量",
            bar_rows=["生産量 (トン)", "工場キャパ (トン)"],
            line_rows=["工場稼働率"],
            anchor="J42",
        ),
    ]

    return SheetConfig(
        name="4971_valuation",
        year_columns=years,
        rows=rows,
        charts=charts,
    )


def main():
    config = build_valuation_sheet()
    out = Path(__file__).parent.parent / "outputs" / "4971-mec-valuation.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    todos = render(config, str(out))
    print(f"\nWrote: {out}")
    print(f"Total TODOs: {len(todos)}")


if __name__ == "__main__":
    main()
