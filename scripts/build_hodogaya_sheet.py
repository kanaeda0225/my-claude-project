"""Build the 4112 Hodogaya valuation .xlsx from the input markdown.

Numbers used here are pulled directly from samples/4112-hodogaya.md.
Where the markdown is silent (e.g., historical SFC financials, exact
PER assumptions, SG&A), the cell is left as "TODO: <ask>" and reported
by the renderer.
"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO = Path(__file__).parent.parent
sys.path.insert(0, str(_REPO / ".claude/skills/takimoto-valuation/scripts"))
from sheet_renderer import ChartSpec, Row, SheetConfig, render  # noqa: E402


# ---------------------------------------------------------------------------
# Constants pulled directly from samples/4112-hodogaya.md
# ---------------------------------------------------------------------------
STOCK_PRICE = 2464  # 円, 2026/5/11
MARKET_CAP_JPY = 38_100_000_000  # 約381億円
SHARES_OUTSTANDING = round(MARKET_CAP_JPY / STOCK_PRICE)  # ~15.46M shares
SFC_STAKE = 0.50  # 保土谷の出資比率

# Tandem adoption rates (from markdown table)
TANDEM_2024 = 0.30  # 30%超 (iPad Pro登場後)
TANDEM_2026 = 0.36

# Segment growth rates 2026 (Omdia)
SEG_GROWTH = {
    "スマホ": -0.03,
    "タブレット": 0.39,
    "ノートPC": 0.46,
    "モニター": 0.61,
}

# LG Chem litigation
LG_DAMAGES_JPY = 2_200_000_000  # 約22億円
LG_SCENARIOS = {"A": 0.40, "B": 0.45, "C": 0.15}


# ---------------------------------------------------------------------------
# Sheet 1: Valuation
# ---------------------------------------------------------------------------
def build_valuation_sheet() -> SheetConfig:
    """Main valuation sheet for 4112 Hodogaya."""
    rows: list[Row] = []

    # ===========================
    # Section: 売上ドライバー (SFC側)
    # ===========================
    rows.append(Row(label="【売上ドライバー：SFC OLED材料需要】", is_section_header=True))

    # IT-OLED panel shipments (decomposed)
    rows.append(
        Row(
            label="IT-OLED総需要 (相対指数)",
            indent=0,
            values=["TODO: 24基準", "TODO", "TODO", "TODO"],
            note="IT-OLEDパネル枚数 × 面積 × 積層数 ÷ 歩留 で算出される需要相対値",
            importance="★★★",
            number_format="#,##0.00",
        )
    )

    rows.append(
        Row(
            label="IT-OLEDパネル出荷枚数 (千枚)",
            indent=1,
            values=["TODO", "TODO", "TODO", "TODO"],
            is_independent=True,
            importance="★★★",
            note="Omdia CAGR+37%。MacBook Pro 200万枚(2026年中)を含む",
        )
    )
    rows.append(
        Row(
            label="MacBook Pro 14/16inch (千枚)",
            indent=2,
            values=[0, 0, "TODO: 2000?", "TODO"],
            is_independent=True,
            note="Digitimes：2026年中 200万枚目標(Apple向け)",
        )
    )
    rows.append(
        Row(
            label="iPad Pro (千枚)",
            indent=2,
            values=["TODO", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="2024年以降タンデム採用継続",
        )
    )
    rows.append(
        Row(
            label="その他IT-OLED (千枚)",
            indent=2,
            values=["TODO", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="2027年以降 iPad Air OLED・MacBook Air OLED など追加",
        )
    )

    rows.append(
        Row(
            label="平均面積補正 (vs スマホ基準)",
            indent=1,
            values=[1.0, "TODO", "TODO", "TODO"],
            is_independent=True,
            importance="★",
            note="スマホ比 約25倍 (一時的な水準切り上げ)。MacBook Pro 16inchで約3,400cm²",
            number_format="0.00",
        )
    )

    rows.append(
        Row(
            label="タンデム積層補正",
            indent=1,
            values=[TANDEM_2024, "TODO", TANDEM_2026, "TODO"],
            is_independent=True,
            importance="★★",
            note="Omdia(2025年4月)：2024年 30%超 → 2026年 36%。タンデムは発光層2層で材料消費 +50〜75%",
            number_format="0.0%",
        )
    )

    rows.append(
        Row(
            label="歩留まり (A6ライン)",
            indent=1,
            values=["TODO", "TODO", 0.85, 0.90],
            is_independent=True,
            importance="★",
            note="Digitimes 2026年4月下旬時点で85%超。目標90%",
            number_format="0.0%",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # Section: SFC損益 (50%出資)
    # ===========================
    rows.append(Row(label="【SFC損益 (Samsung Fine Chemicals)】", is_section_header=True))

    rows.append(
        Row(
            label="SFC売上高 (百万円換算)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="ホスト材料単価 × 出荷量。為替も影響",
        )
    )
    rows.append(
        Row(
            label="ホスト材料平均単価",
            indent=1,
            values=["TODO", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="青色ホスト材料の㎏あたり単価。タンデム比率上昇で構成比変化",
        )
    )

    rows.append(
        Row(
            label="SFC粗利",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
        )
    )
    rows.append(
        Row(
            label="SFC原価率",
            indent=1,
            values=["TODO", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="高純度品 (不純物 1ppm以下) の原価構造",
            number_format="0.0%",
        )
    )

    rows.append(
        Row(
            label="SFC販管費",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="天安工場運営費 + 韓国本社経費。生産量に概ね比例",
        )
    )

    rows.append(
        Row(
            label="SFC営業利益",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
        )
    )
    rows.append(
        Row(
            label="SFC法人税率 (韓国)",
            indent=1,
            values=[0.22, 0.22, 0.22, 0.22],
            is_independent=True,
            note="韓国の法人税実効税率 概ね22%。為替・控除によって変動",
            number_format="0.0%",
        )
    )

    rows.append(
        Row(
            label="SFC純利益",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # Section: 4112 連結損益
    # ===========================
    rows.append(Row(label="【保土谷化学 連結損益】", is_section_header=True))

    rows.append(
        Row(
            label="国内化成品事業 売上高",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="染料・農薬中間体など。成熟。成長ドライバーではない",
        )
    )

    rows.append(
        Row(
            label="国内化成品事業 純利益",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="据え置き想定。詳細は会社開示のセグメント情報",
        )
    )

    rows.append(
        Row(
            label="SFCからの持分法利益 (×50%)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="SFC純利益 × 50% (保土谷の出資比率)",
        )
    )

    rows.append(
        Row(
            label="連結 親会社株主帰属純利益",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # Section: 株価 & PER
    # ===========================
    rows.append(Row(label="【株価・PER感応度】", is_section_header=True))

    rows.append(
        Row(
            label="発行済株式数 (千株)",
            indent=0,
            values=[
                SHARES_OUTSTANDING // 1000,
                SHARES_OUTSTANDING // 1000,
                SHARES_OUTSTANDING // 1000,
                SHARES_OUTSTANDING // 1000,
            ],
            is_independent=True,
            note=f"時価総額 {MARKET_CAP_JPY/1e8:.0f}億 / 株価 {STOCK_PRICE}円 から逆算",
        )
    )

    rows.append(
        Row(
            label="EPS (円)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="連結純利益 ÷ 発行済株式数",
        )
    )

    rows.append(
        Row(
            label="Bear PER",
            indent=1,
            values=["TODO: 10?", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="MacBook Pro量産がFY27以降にズレ込み。SFCの寄与が限定的なシナリオ",
        )
    )
    rows.append(
        Row(
            label="Base PER",
            indent=1,
            values=["TODO: 15?", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="MacBook Pro 200万枚出荷。タンデム効果をフル享受",
        )
    )
    rows.append(
        Row(
            label="Bull PER",
            indent=1,
            values=["TODO: 20?", "TODO", "TODO", "TODO"],
            is_independent=True,
            note="MacBook Pro + iPad Pro継続 + 追加IT-OLED案件",
        )
    )

    rows.append(
        Row(
            label="目標株価 Bear (円)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="EPS × Bear PER",
        )
    )
    rows.append(
        Row(
            label="目標株価 Base (円)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="EPS × Base PER",
        )
    )
    rows.append(
        Row(
            label="目標株価 Bull (円)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="EPS × Bull PER",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # Section: 「織り込み済み?」反論用
    # ===========================
    rows.append(
        Row(label="【「織り込み済み?」反論：時価総額据え置きシナリオ】", is_section_header=True)
    )

    rows.append(
        Row(
            label="時価総額 (百万円・据え置き)",
            indent=0,
            values=[
                MARKET_CAP_JPY // 1_000_000,
                MARKET_CAP_JPY // 1_000_000,
                MARKET_CAP_JPY // 1_000_000,
                MARKET_CAP_JPY // 1_000_000,
            ],
            is_independent=True,
            note=f"現在時価総額 {MARKET_CAP_JPY/1e8:.0f}億円 を全年度で固定 (反論用)",
        )
    )

    rows.append(
        Row(
            label="連結純利益 Base (百万円)",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="上記Baseシナリオの純利益を再掲",
        )
    )

    rows.append(
        Row(
            label="時価総額据置PER",
            indent=0,
            values=["TODO", "TODO", "TODO", "TODO"],
            note="時価総額 ÷ Base純利益。利益増加で PER は機械的に低下する",
            number_format="0.0",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # Section: ピア比較
    # ===========================
    rows.append(Row(label="【ピア比較 (PER)】", is_section_header=True))

    rows.append(
        Row(
            label="保土谷化学 (4112)",
            indent=0,
            values=["TODO: 現PER", "", "", ""],
            is_independent=True,
            note=f"現株価 {STOCK_PRICE}円 / EPS(直近実績) で算出",
            number_format="0.0",
        )
    )
    rows.append(
        Row(
            label="ピア① OLED材料系上場会社",
            indent=1,
            values=["TODO", "", "", ""],
            is_independent=True,
            note="多軸選定：OLED材料同業。SDCの別サプライヤーが上場していれば理想",
            number_format="0.0",
        )
    )
    rows.append(
        Row(
            label="ピア② 化学 (染料/中間体) 業界平均",
            indent=1,
            values=["TODO", "", "", ""],
            is_independent=True,
            note="多軸選定：表向きの業態カテゴリ。市場の現在の認識を反映",
            number_format="0.0",
        )
    )
    rows.append(
        Row(
            label="ピア③ 持分法主導 (連結利益≒関連会社利益) 企業",
            indent=1,
            values=["TODO", "", "", ""],
            is_independent=True,
            note="多軸選定：ビジネス構造の類似(50%出資先が主力)。例：商社系の持分会社主導銘柄",
            number_format="0.0",
        )
    )

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # Section: KPI / Riskオフセット
    # ===========================
    rows.append(Row(label="【KPI & Riskオフセット】", is_section_header=True))

    rows.append(
        Row(
            label="A6ライン歩留まり",
            indent=0,
            values=["TODO", "TODO", 0.85, 0.90],
            is_independent=True,
            importance="★",
            note="Digitimes：2026年4月時点85%超。85%超は量産フェーズとして十分",
            number_format="0.0%",
        )
    )
    rows.append(
        Row(
            label="タンデム採用率 (IT-OLED全体)",
            indent=0,
            values=[TANDEM_2024, "TODO", TANDEM_2026, "TODO"],
            is_independent=True,
            importance="★★",
            note="Omdia：2024年30%超 → 2026年36%。今後も上昇見込み",
            number_format="0.0%",
        )
    )
    rows.append(
        Row(
            label="LG Chem訴訟 特別損失(シナリオB)",
            indent=0,
            values=[0, 0, LG_DAMAGES_JPY // 1_000_000, 0],
            is_independent=True,
            note=f"シナリオB確率45%。約22億円。判決2026年7〜8月予定。詳細はAppendix",
        )
    )
    rows.append(
        Row(
            label="シナリオB確率",
            indent=1,
            values=[0.45, 0.45, 0.45, 0.45],
            is_independent=True,
            note="LG Chem敗訴で損害賠償のみ。製造販売は継続。株価への影響は軽微",
            number_format="0.0%",
        )
    )

    # ===========================
    # Charts
    # ===========================
    charts = [
        ChartSpec(
            title="連結損益と利益率",
            bar_rows=[
                "連結 親会社株主帰属純利益",
            ],
            line_rows=[],
            anchor="K2",
        ),
        ChartSpec(
            title="目標株価シナリオ別",
            bar_rows=[
                "目標株価 Bear (円)",
                "目標株価 Base (円)",
                "目標株価 Bull (円)",
            ],
            line_rows=[],
            anchor="K22",
        ),
        ChartSpec(
            title="タンデム採用率と歩留まり",
            bar_rows=[],
            line_rows=[
                "タンデム採用率 (IT-OLED全体)",
                "A6ライン歩留まり",
            ],
            anchor="K42",
        ),
    ]

    return SheetConfig(
        name="4112_valuation",
        year_columns=["FY24/12 実", "FY25/12 実", "FY26/12 予", "FY27/12 予"],
        rows=rows,
        charts=charts,
    )


def main():
    config = build_valuation_sheet()
    out = Path(__file__).parent.parent / "outputs" / "4112-hodogaya-valuation.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    todos = render(config, str(out))
    print(f"\nWrote: {out}")
    print(f"Total TODOs: {len(todos)}")


if __name__ == "__main__":
    main()
