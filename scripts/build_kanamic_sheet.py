"""Build the 3939 カナミックネットワーク valuation .xlsx.

Data sources:
- Historical (FY21-FY25 actuals): 3939_yuho_101_20{23,24,25}0930.pdf
- FY26 会社予想: 3939_eigyo_101_20251217.pdf / 決算説明資料(2026年5月13日発表) p9
- FY26 2Q実績: 3939_hanpo_101_20260331.pdf
- FY27/FY28 シナリオ: /tmp/story_3939.md (ストーリーmd)

Pattern: SaaS型 = 有料ユーザーID数 × 月額ARPU × 12 ＋ 健康寿命延伸 ＋ ソリューション開発
"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO = Path(__file__).parent.parent
sys.path.insert(0, str(_REPO / ".claude/skills/takimoto-valuation/scripts"))
from sheet_renderer import ChartSpec, Row, SheetConfig, render  # noqa: E402


# ---------------------------------------------------------------------------
# Historical actuals (FY23-FY25) — all in 百万円
# Source: 3939_yuho_101_20{23,24,25}0930.pdf 連結損益計算書
# ---------------------------------------------------------------------------
ACT = {
    "FY23": {
        "売上高": 3_746, "売上原価": 1_152, "売上総利益": 2_594,
        "販管費": 1_498, "営業利益": 1_096, "経常利益": 1_107,
        "税前": 1_104, "法人税等": 341, "純利益": 763, "EPS": 16.08,
        # セグメント (FY23時点では「クラウドPF事業」に健康寿命延伸も含む)
        "クラウドPF事業セグ売上": 3_695,  # カナミック2,402 + プラットフォーム300 + 健康寿命845 + その他147
        "_カナミッククラウド": 2_402, "_プラットフォーム": 300,
        "_健康寿命延伸": 846, "_その他クラウドPF": 147,
        "ソリューション開発": 51,
        "健康寿命延伸事業": None,  # FY24からセグメント分離
        # SG&A内訳
        "従業員給料": 353, "役員報酬": 241, "地代家賃": 85,
        "法定福利費": 77, "賞与引当金繰入額": 37, "減価償却費": 44,
        "貸倒引当金繰入額": 4,
        # KPI
        "有料ユーザーID数_期末": 155_516,  # 2023/9 4Q
        "無料ユーザーID数_期末": 86_760,
        "導入地域数": 1_270,
        "アーバンフィット店舗数": 13,
    },
    "FY24": {
        "売上高": 5_007, "売上原価": 1_935, "売上総利益": 3_072,
        "販管費": 1_633, "営業利益": 1_439, "経常利益": 1_448,
        "税前": 1_352, "法人税等": 432, "純利益": 920, "EPS": 19.38,
        "クラウドPF事業セグ売上": 3_365,  # FY24からは健康寿命延伸を切り出し
        "_カナミッククラウド": 2_853, "_プラットフォーム": 424, "_その他クラウドPF": 88,
        "健康寿命延伸事業": 1_131,
        "ソリューション開発": 512,
        "従業員給料": 381, "役員報酬": 251, "地代家賃": 92,
        "法定福利費": 82, "賞与引当金繰入額": 40, "減価償却費": 46,
        "貸倒引当金繰入額": 6,
        "有料ユーザーID数_期末": 191_813,
        "無料ユーザーID数_期末": 101_309,
        "導入地域数": 1_320,
        "アーバンフィット店舗数": 20,
    },
    "FY25": {
        "売上高": 5_501, "売上原価": 1_964, "売上総利益": 3_536,
        "販管費": 1_930, "営業利益": 1_607, "経常利益": 1_613,
        "税前": 1_613, "法人税等": 501, "純利益": 1_112, "EPS": 23.42,
        "クラウドPF事業セグ売上": 3_583,
        "_カナミッククラウド": 2_997, "_プラットフォーム": 474, "_その他クラウドPF": 112,
        "健康寿命延伸事業": 1_188,
        "ソリューション開発": 730,
        "従業員給料": 486, "役員報酬": 256, "地代家賃": 85,
        "法定福利費": 99, "賞与引当金繰入額": 51, "減価償却費": 47,
        "貸倒引当金繰入額": 1,
        "有料ユーザーID数_期末": 221_822,
        "無料ユーザーID数_期末": 117_497,
        "導入地域数": 1_397,
        "アーバンフィット店舗数": 24,
    },
}

# ---------------------------------------------------------------------------
# FY26 (会社予想) / FY27 Base / FY28 Bear/Base/Bull projections
# Source: ストーリーmd 4-2節シナリオ別簡易PL + 決算説明資料 通期予想
# ---------------------------------------------------------------------------
# FY26 (会社予想): 売上 6,350 / 営利 2,050 / 純利 1,370 / EPS 28.5
# FY27 Base (推計): 売上 7,200 / 営利 2,376 / 純利 1,663 / EPS 34.6
# FY28 Bear: 売上 6,500 / 営利 1,950 / 純利 1,370 / EPS 28.4
# FY28 Base: 売上 8,100 / 営利 2,750 / 純利 1,930 / EPS 40.0
# FY28 Bull: 売上 10,100 / 営利 3,737 / 純利 2,620 / EPS 54.4

SHARES = 48_132  # 千株 (発行済株式総数、自己株なし)
PRICE_REF = 550  # 円 (2026年5月13日2Q決算発表時点、ストーリーmd記載)
MARKET_CAP_REF = round(PRICE_REF * SHARES / 1000)  # 百万円 = 26,473


def build_valuation_sheet() -> SheetConfig:
    rows: list[Row] = []

    # Year columns: 5 timeseries + 3 FY28 scenarios
    years = [
        "FY23/9 実", "FY24/9 実", "FY25/9 実",
        "FY26/9 予(会社)", "FY27/9 予(Base)",
        "FY28/9 Bear", "FY28/9 Base", "FY28/9 Bull",
    ]

    def v8(fy23, fy24, fy25, fy26, fy27, bear, base, bull):
        """Helper to build an 8-element values list."""
        return [fy23, fy24, fy25, fy26, fy27, bear, base, bull]

    # ===========================================
    # 売上高
    # ===========================================
    rows.append(Row(label="【売上高】", is_section_header=True))
    rows.append(Row(
        label="売上高 合計",
        values=v8(3_746, 5_007, 5_501, 6_350, 7_200, 6_500, 8_100, 10_100),
        note="連結損益計算書(FY23-25実績)、FY26は会社予想(2026年5月13日決算説明資料)、FY27/FY28はストーリーmd",
    ))
    rows.append(Row(
        label="増収率(YoY)",
        indent=1,
        values=v8(None, 0.337, 0.099, 0.154, 0.134, None, None, None),
        note="(売上高÷前期売上高)−1。FY28は3シナリオの並列(YoY概念外)につき空白。FY27 Base比でBear-9.7%/Base+12.5%/Bull+40.3%",
        number_format="0.0%",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # セグメント別売上
    # ===========================================
    rows.append(Row(label="【セグメント別売上】", is_section_header=True))
    rows.append(Row(
        label="医療・介護クラウドPF事業",
        values=v8(2_850, 3_365, 3_583, 4_120, 4_700, 4_300, 5_400, 6_800),
        is_independent=True,
        importance="★★★",
        note="yuho開示。FY23は健康寿命延伸込みのため845を控除した値(2,850)を併記。FY26+はクラウド+プラットフォーム+その他で構成し、AI課金でARPU上昇",
    ))
    rows.append(Row(
        label="カナミッククラウドサービス",
        indent=1,
        values=v8(2_402, 2_853, 2_997, 3_400, 3_900, 3_600, 4_500, 5_700),
        is_independent=True,
        importance="★★★",
        note="ストックビジネス。FY25は前期比+5.1%、FY24は+18.8%(M&A効果)、本ラインがAI課金の主受益",
    ))
    rows.append(Row(
        label="プラットフォームサービス",
        indent=1,
        values=v8(300, 424, 474, 550, 620, 550, 700, 850),
        is_independent=True,
        note="HP構築/介護労働安定センター受託/インターネット広告。FY25+11.9%",
    ))
    rows.append(Row(
        label="その他サービス(クラウドPF事業内)",
        indent=1,
        values=v8(147, 88, 112, 170, 180, 150, 200, 250),
        is_independent=True,
        note="大口顧客向けカスタマイズ開発受託。ボラあり",
    ))
    rows.append(Row(
        label="健康寿命延伸事業",
        values=v8(846, 1_131, 1_188, 1_300, 1_500, 1_400, 1_600, 1_900),
        is_independent=True,
        importance="★★",
        note="アーバンフィット24(2026年2Q末で27店舗)。FY23時点はクラウドPF事業内で開示、FY24以降にセグメント分離",
    ))
    rows.append(Row(
        label="ソリューション開発事業",
        values=v8(51, 512, 730, 930, 1_000, 800, 1_100, 1_400),
        is_independent=True,
        importance="★",
        note="Ruby開発(2023/8連結)+TWM(2025/2連結)。FY24+904%/FY25+42.6%。海外展開でも寄与",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # 因数分解 (SaaS型: 有料ユーザー × ARPU × 12)
    # ===========================================
    rows.append(Row(label="【因数分解: 有料ユーザー数 × 月額ARPU × 12ヶ月】", is_section_header=True))
    rows.append(Row(
        label="有料ユーザーID数 (期末、人)",
        values=v8(155_516, 191_813, 221_822, 240_000, 260_000, 260_000, 280_000, 310_000),
        is_independent=True,
        importance="★★★",
        note="決算説明資料1-8。FY23/9~FY25/9実績。FY26末はFY26-2Q時点247,836→+8%伸び想定。Base CAGR+9%/Bear+6%/Bull+12%",
    ))
    rows.append(Row(
        label="増加率 YoY",
        indent=1,
        values=v8(None, 0.233, 0.157, 0.082, 0.083, 0.083, 0.122, 0.241),
        note="FY24/9末+23.3%(M&A効果)、FY25/9末+15.7%。FY26-FY28はストーリーmd",
        number_format="0.0%",
    ))
    rows.append(Row(
        label="クラウドPF+ソリューション 月額ARPU (円)",
        values=v8(1_350, 1_500, 1_600, 1_660, 1_750, 1_560, 1_850, 2_100),
        is_independent=True,
        importance="★★★",
        note="(クラウドPF+ソリューション売上) ÷ 有料ユーザー数(年央概算) ÷ 12。FY28はストーリーmd前提。AI課金浸透がドライバー",
    ))
    rows.append(Row(
        label="ARPU 上昇率 YoY",
        indent=1,
        values=v8(None, 0.111, 0.067, 0.038, 0.054, -0.024, 0.158, 0.313),
        note="FY28 BaseはAI課金浸透25-30%でARPU+15.8%、BullはAI浸透40%超でARPU+31%超",
        number_format="0.0%",
    ))
    rows.append(Row(label="", is_blank=True))

    # AI課金サブブロック
    rows.append(Row(label="  ▽ AI課金浸透(変曲点ドライバー)", indent=0))
    rows.append(Row(
        label="AI機能浸透率 (有料IDの%)",
        indent=1,
        values=v8(0.0, 0.0, 0.0, 0.05, 0.15, 0.10, 0.27, 0.42),
        is_independent=True,
        importance="★★★",
        note="FY26から課金開始(AI100-Vision)。Base FY28で25-30%浸透、Bull 40%超。2027年改定で加算要件化なら非線形",
        number_format="0.0%",
    ))
    rows.append(Row(
        label="AI機能 平均月額単価 (円)",
        indent=1,
        values=v8(None, None, None, 800, 900, 700, 1_000, 1_200),
        is_independent=True,
        importance="★★",
        note="未開示。月額500-1,500円とストーリーmd記載。Base 1,000円/Bear 700/Bull 1,200で置く",
    ))
    rows.append(Row(
        label="AI課金 ARPU寄与 (円/月、ブレンド)",
        indent=1,
        values=v8(0, 0, 0, 40, 135, 70, 270, 504),
        note="浸透率 × 平均単価。Base FY28: 27%×1,000=270円が既存ARPUに上乗せ",
    ))
    rows.append(Row(label="", is_blank=True))

    # 健康寿命延伸サブブロック
    rows.append(Row(label="  ▽ 健康寿命延伸サブドライバー", indent=0))
    rows.append(Row(
        label="アーバンフィット 店舗数 (期末)",
        indent=1,
        values=v8(13, 20, 24, 30, 35, 32, 36, 42),
        is_independent=True,
        importance="★",
        note="2026/3末で27店舗(2026年5月時点)。Vision 2035ベースで年4-6店舗増想定",
        number_format="0",
    ))
    rows.append(Row(
        label="1店舗あたり年商 (百万円)",
        indent=1,
        values=v8(65, 57, 49, 43, 43, 44, 44, 45),
        is_independent=True,
        note="健康寿命延伸事業売上 ÷ 期中平均店舗数。新規出店で当初は低稼働、徐々に45百万へ収斂",
        number_format="0.0",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # 粗利
    # ===========================================
    rows.append(Row(label="【粗利】", is_section_header=True))
    rows.append(Row(
        label="売上原価",
        values=v8(1_152, 1_935, 1_964, 2_223, 2_448, 2_308, 2_673, 3_131),
        note="連結P&L(FY23-25実績)。FY26+は売上×(1-粗利率)。FY24は新規連結(Ruby/アーバンフィット)で粗利率低下",
    ))
    rows.append(Row(
        label="売上総利益",
        values=v8(2_594, 3_072, 3_536, 4_128, 4_752, 4_193, 5_427, 6_969),
        note="連結P&L。FY26+は売上×粗利率",
    ))
    rows.append(Row(
        label="粗利率",
        indent=1,
        values=v8(0.692, 0.614, 0.643, 0.650, 0.660, 0.645, 0.670, 0.690),
        is_independent=True,
        note="FY23 69.2%→FY24 61.4%(M&Aで一時悪化)→FY25 64.3%回復→AI課金限界利益≈100%で漸増",
        number_format="0.0%",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # 販管費
    # ===========================================
    rows.append(Row(label="【販管費】", is_section_header=True))
    rows.append(Row(
        label="販管費 合計",
        values=v8(1_498, 1_633, 1_930, 2_078, 2_376, 2_243, 2_677, 3_232),
        note="連結P&L(FY23-25実績)。FY26+は粗利−営利。営業利益率目標から逆算",
    ))
    sga_items = [
        ("従業員給料", [353, 381, 486, 580, 670, 620, 760, 950],
         "yuho販管費注記。FY25+27.6%は新規連結とTWM連結(2025/2)。SaaS人件費は売上に劣後成長"),
        ("役員報酬", [241, 251, 256, 270, 285, 280, 295, 310],
         "yuho販管費注記。FY24+4%、FY25+2.3%。安定推移"),
        ("地代家賃", [85, 92, 85, 90, 95, 95, 100, 110],
         "本社+営業所(7拠点)。FY25は減(再配置)"),
        ("法定福利費", [77, 82, 99, 110, 125, 120, 140, 170],
         "従業員給料に概ね連動"),
        ("賞与引当金繰入額", [37, 40, 51, 60, 70, 65, 80, 100],
         "業績連動。OPM拡大局面で増額"),
        ("減価償却費", [44, 46, 47, 50, 55, 55, 60, 70],
         "販管費分のみ。連結全体の減価償却費はFY25で332百万(B/S)"),
        ("貸倒引当金繰入額", [4, 6, 1, 3, 3, 3, 3, 3],
         "据え置き(過去レンジ1-6百万)"),
    ]
    for item, vals, note in sga_items:
        rows.append(Row(
            label=item, indent=1, values=vals,
            is_independent=True, note=note,
        ))
    # SG&A残余
    rows.append(Row(
        label="その他販管費",
        indent=1,
        values=v8(657, 735, 905, 915, 1_073, 1_005, 1_239, 1_519),
        is_independent=True,
        note="販管費合計から上記内訳を差引いた残余(外注費/広告/業務委託等)",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # 利益
    # ===========================================
    rows.append(Row(label="【利益】", is_section_header=True))
    rows.append(Row(
        label="営業利益",
        values=v8(1_096, 1_439, 1_607, 2_050, 2_376, 1_950, 2_750, 3_737),
        note="連結P&L(FY23-25実績)。FY26会社予想2,050。FY27-FY28シナリオはストーリーmd",
    ))
    rows.append(Row(
        label="営業利益率",
        indent=1,
        values=v8(0.293, 0.287, 0.292, 0.323, 0.330, 0.300, 0.340, 0.370),
        is_independent=True,
        importance="★★",
        note="FY25 29.2%→FY26予32.3%。Vision 2035で2030年33-35%目標。Base FY28 34%は妥当、AI限界利益反映",
        number_format="0.0%",
    ))
    rows.append(Row(
        label="経常利益",
        values=v8(1_107, 1_448, 1_613, 2_055, 2_380, 1_955, 2_755, 3_745),
        note="連結P&L(FY23-25実績)。営業外損益は小幅。FY26+は営利+5百万で簡略",
    ))
    rows.append(Row(
        label="営業外損益(純)",
        indent=1,
        values=v8(11, 9, 6, 5, 5, 5, 5, 5),
        is_independent=True,
        note="受取利息/手数料等。安定して+5〜10百万円",
    ))
    rows.append(Row(
        label="税金等調整前純利益",
        values=v8(1_104, 1_352, 1_613, 2_055, 2_380, 1_955, 2_755, 3_745),
        note="FY24は特別損失(減損96+移転7)で経常から減少。FY26+は特益・特損ゼロ想定",
    ))
    rows.append(Row(
        label="法人税等",
        values=v8(341, 432, 501, 617, 714, 587, 827, 1_124),
        note="実効税率を乗じた値",
    ))
    rows.append(Row(
        label="実効税率",
        indent=1,
        values=v8(0.309, 0.320, 0.311, 0.300, 0.300, 0.300, 0.300, 0.300),
        is_independent=True,
        note="FY23-25実績30-32%。法定実効税率約30%。FY26+は30%で固定",
        number_format="0.0%",
    ))
    rows.append(Row(
        label="親会社株主帰属純利益",
        values=v8(763, 920, 1_112, 1_437, 1_666, 1_369, 1_929, 2_622),
        note="連結P&L。FY26会社予想は1,370(28.5円)、ストーリー値とほぼ一致",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # 補助KPI (SaaS指標)
    # ===========================================
    rows.append(Row(label="【補助KPI: SaaS指標】", is_section_header=True))
    rows.append(Row(
        label="導入地域数 (期末)",
        values=v8(1_270, 1_320, 1_397, 1_440, 1_490, 1_460, 1_510, 1_580),
        is_independent=True,
        note="決算説明資料1-7。中学校区(人口3万人圏)単位。FY26 2Q時点1,420地域",
        number_format="0",
    ))
    rows.append(Row(
        label="無料ユーザーID数 (期末)",
        values=v8(86_760, 101_309, 117_497, 135_000, 152_000, 145_000, 158_000, 175_000),
        is_independent=True,
        note="決算説明資料1-8。無料→有料転換のパイプライン。介護情報基盤稼働で増加加速可能性",
    ))
    rows.append(Row(
        label="ユーザーID合計",
        values=v8(242_276, 293_122, 339_319, 375_000, 412_000, 405_000, 438_000, 485_000),
        note="有料+無料。FY26 2Q時点で376,972",
    ))
    rows.append(Row(
        label="EBITDA",
        values=v8(1_140, 1_740, 1_939, 2_500, 2_900, 2_400, 3_300, 4_400),
        note="営業利益+減価償却費(連結全体)。FY25実は1,939、FY26会社予想は2,500",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # 株価・PER感応度
    # ===========================================
    rows.append(Row(label="【株価・PER感応度】", is_section_header=True))
    rows.append(Row(
        label="発行済株式数 (千株、期末)",
        values=v8(SHARES, SHARES, SHARES, SHARES, SHARES, SHARES, SHARES, SHARES),
        is_independent=True,
        note="yuho開示。直近5年48,132千株で据置(自己株消却・新株発行なし)",
    ))
    rows.append(Row(
        label="EPS (円)",
        values=v8(16.08, 19.38, 23.42, 29.86, 34.62, 28.45, 40.08, 54.48),
        note="純利益÷発行済株式数。FY23-25はyuho開示と一致。FY28 Base 40.0/Bull 54.4はストーリーmd",
        number_format="0.00",
    ))
    rows.append(Row(
        label="想定PER ベスト(Bull水準)",
        indent=1,
        values=v8(None, None, None, None, None, None, None, 22.0),
        is_independent=True,
        note="Bullシナリオ:AI浸透確認でプレミアム。情報通信SaaS OPM30%超で18-25倍が相場、+α",
        number_format="0.0",
    ))
    rows.append(Row(
        label="想定PER ノーマル(Base水準)",
        indent=1,
        values=v8(None, None, None, None, None, None, 19.0, None),
        is_independent=True,
        note="Baseシナリオ:現状PER(19.3倍)維持。SaaS中央値かつカナミック直近の典型水準",
        number_format="0.0",
    ))
    rows.append(Row(
        label="想定PER ワースト(Bear水準)",
        indent=1,
        values=v8(None, None, None, None, None, 16.0, None, None),
        is_independent=True,
        note="Bearシナリオ:AI浸透遅延でディスカウント。過去レンジ下限14.9倍に若干上乗せ",
        number_format="0.0",
    ))
    rows.append(Row(
        label="目標株価 Bull (円)",
        values=v8(None, None, None, None, None, None, None, 1_199),
        note="FY28 Bull EPS 54.48 × 22倍 = 1,199円(現値+118%、ストーリー1,197と一致)",
        number_format="#,##0",
    ))
    rows.append(Row(
        label="目標株価 Base (円)",
        values=v8(None, None, None, None, None, None, 762, None),
        note="FY28 Base EPS 40.08 × 19倍 = 762円(現値+38%、ストーリー760と一致)",
        number_format="#,##0",
    ))
    rows.append(Row(
        label="目標株価 Bear (円)",
        values=v8(None, None, None, None, None, 455, None, None),
        note="FY28 Bear EPS 28.45 × 16倍 = 455円(現値-17%、ストーリー454と一致)",
        number_format="#,##0",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # 「織り込み済み?」反論ブロック
    # ===========================================
    rows.append(Row(label="【「織り込み済み?」反論】", is_section_header=True))
    rows.append(Row(
        label="時価総額 (百万円、現値据置)",
        values=v8(MARKET_CAP_REF, MARKET_CAP_REF, MARKET_CAP_REF, MARKET_CAP_REF,
                  MARKET_CAP_REF, MARKET_CAP_REF, MARKET_CAP_REF, MARKET_CAP_REF),
        is_independent=True,
        note="2026年5月13日株価550円 × 48,132千株 = 26,473百万円で全期間固定",
    ))
    rows.append(Row(
        label="連結純利益 Base (再掲)",
        indent=1,
        values=v8(763, 920, 1_112, 1_437, 1_666, 1_369, 1_929, 2_622),
        note="上記P&Lより再掲",
    ))
    rows.append(Row(
        label="時価総額据置PER (倍)",
        indent=1,
        values=v8(34.7, 28.8, 23.8, 18.4, 15.9, 19.3, 13.7, 10.1),
        note="時価総額 ÷ 純利益。株価変わらず利益増ならPERが機械的に低下→現値に純利益成長は織込まれていない",
        number_format="0.0",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # ピア比較 (PER)
    # ===========================================
    rows.append(Row(label="【ピア比較 (PER)】※純粋同業未上場、多軸選定", is_section_header=True))
    rows.append(Row(
        label="カナミック (3939) 現状PER",
        values=v8(None, None, None, 19.3, None, None, None, None),
        is_independent=True,
        note="株価550円 ÷ FY26予EPS 28.46円 = 19.3倍(2026年5月13日時点)",
        number_format="0.0",
    ))
    rows.append(Row(
        label="ピア① エス・エム・エス (2175)",
        indent=1,
        values=v8(None, None, None, "TODO", None, None, None, None),
        is_independent=True,
        note="医療介護領域のヘルスケアSaaS+人材プラットフォーム。軸:同業界(介護DX)。直近実績PERをIR確認推奨",
        number_format="0.0",
    ))
    rows.append(Row(
        label="ピア② ラクス (3923)",
        indent=1,
        values=v8(None, None, None, "TODO", None, None, None, None),
        is_independent=True,
        note="中小事業者向け低単価SaaS(楽楽精算等)。軸:ビジネスモデル類似(課金1,500円帯×多数の小口顧客)。高成長SaaSプレミアム参考",
        number_format="0.0",
    ))
    rows.append(Row(
        label="ピア③ JMDC (4483)",
        indent=1,
        values=v8(None, None, None, "TODO", None, None, None, None),
        is_independent=True,
        note="医療データプラットフォーム。軸:制度受益(医療DX政策・基盤)+データMoat。直近FY26予想ベースPERをIR確認推奨",
        number_format="0.0",
    ))
    rows.append(Row(
        label="ピア中央値(参考レンジ20-30倍)",
        indent=1,
        values=v8(None, None, None, "TODO", None, None, None, None),
        is_independent=True,
        note="プライム情報通信SaaS(OPM30%超・増益率20%超)のPER相場18-25倍を参考。実数値は要IR確認",
        number_format="0.0",
    ))
    rows.append(Row(label="", is_blank=True))

    # ===========================================
    # 仮定根拠サマリー
    # ===========================================
    rows.append(Row(label="【仮定根拠の要約】", is_section_header=True))
    rows.append(Row(
        label="業績タイプ判定",
        values=v8(None, None, None, None, None, None, None, None),
        note="SaaS型(ユーザー数×ARPU×12)。AlbaLink型(1人粗利×人数)+セグメント型のハイブリッド。主因数: 有料ID数とARPU",
    ))
    rows.append(Row(
        label="ストーリー核心",
        values=v8(None, None, None, None, None, None, None, None),
        note="既存22万有料ID×AI課金で獲得コストゼロのARPU階段上昇。2027年改定+介護情報基盤稼働の3つが2026-27に重なる構造変化",
    ))
    rows.append(Row(
        label="最大の独立変数",
        values=v8(None, None, None, None, None, None, None, None),
        note="(1)AI機能浸透率 (2)AI平均単価 (3)有料ユーザーCAGR。この3つでBear-Bullレンジが決まる",
    ))

    # ===========================================
    # Charts
    # ===========================================
    charts = [
        ChartSpec(
            title="売上・粗利・営業利益 と OPM",
            bar_rows=["売上高 合計", "売上総利益", "営業利益"],
            line_rows=["営業利益率"],
            anchor="K2",
        ),
        ChartSpec(
            title="有料ユーザー数 と 月額ARPU",
            bar_rows=["有料ユーザーID数 (期末、人)"],
            line_rows=["クラウドPF+ソリューション 月額ARPU (円)"],
            anchor="K22",
        ),
        ChartSpec(
            title="セグメント別売上(積上げ視覚)",
            bar_rows=["医療・介護クラウドPF事業", "健康寿命延伸事業", "ソリューション開発事業"],
            line_rows=["増収率(YoY)"],
            anchor="K42",
        ),
        ChartSpec(
            title="時価総額据置PERの低下 (織り込み済み反論)",
            bar_rows=["連結純利益 Base (再掲)"],
            line_rows=["時価総額据置PER (倍)"],
            anchor="K62",
        ),
    ]

    return SheetConfig(
        name="3939_valuation",
        year_columns=years,
        rows=rows,
        charts=charts,
    )


def main():
    config = build_valuation_sheet()
    out = Path(__file__).parent.parent / "outputs" / "3939-kanamic-valuation.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    todos = render(config, str(out))
    print(f"\nWrote: {out}")
    print(f"Total TODOs: {len(todos)}")


if __name__ == "__main__":
    main()
