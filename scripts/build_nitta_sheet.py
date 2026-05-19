"""Build the 5186 Nitta valuation .xlsx.

Inputs:
  - samples (markdown story) — 投資推奨ストーリー
  - 有報 2024/3, 2025/3 — 5年指標, 連結P&L, セグメント別, 販管費内訳, 関連会社合算
  - 決算説明資料 2025/3, 2026/3 — FY3/25, FY3/26 P&L概要, 持分法利益 GUA/NDI 内訳,
                                  FY3/27 会社予想, 中計目標
  - 半期報告書 2024/9, 2025/9

業績タイプ: ⑤ 持分法型 (保土谷と同じ構造、ただしGUA + NDIの2系列ある点が特徴)
"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO = Path(__file__).parent.parent
sys.path.insert(0, str(_REPO / ".claude/skills/takimoto-valuation/scripts"))
from sheet_renderer import ChartSpec, Row, SheetConfig, render  # noqa: E402


# ---------------------------------------------------------------------------
# Constants (markdown + IRBANK 暗黙の参照値)
# ---------------------------------------------------------------------------
# 株価想定: markdownの「現在比+35〜40%」「Bear ±0%」から逆算
# Base 7,350 / 1.35〜1.40 = 5,250〜5,440 → 中央値 5,400円 を想定基準株価とする
STOCK_PRICE = 5_400  # 円, markdown 「現在株価」相当
SHARES_OUTSTANDING_END = 29_272_503  # 期末発行済株式総数 (yuho 第96期)
TREASURY_SHARES = 1_464_792  # 自己株式 (yuho 第96期)
SHARES_OUTSTANDING_NET = SHARES_OUTSTANDING_END - TREASURY_SHARES  # 27,807,711
SHARES_AVG_PERIOD = 27_777_000  # EPS算定の期中平均株式数 (yuho)
MARKET_CAP_JPY = STOCK_PRICE * SHARES_OUTSTANDING_NET  # ≈ 1,502億円
NDI_STAKE = 0.50  # ニッタ・デュポン 出資比率


# ---------------------------------------------------------------------------
# Main valuation sheet
# ---------------------------------------------------------------------------
def build_valuation_sheet() -> SheetConfig:
    rows: list[Row] = []
    YEARS = ["FY3/23 実", "FY3/24 実", "FY3/25 実", "FY3/26 実", "FY3/27 予(会社)", "FY3/28 予(Base)"]

    # ===========================
    # 【売上高】
    # ===========================
    rows.append(Row(label="【売上高 (連結)】", is_section_header=True))

    rows.append(Row(
        label="売上高 合計 (百万円)",
        indent=0,
        values=[88_000, 88_609, 90_276, 91_834, 94_000, 96_200],
        note="FY3/23-25: yuho開示 / FY3/26: 決算説明資料(2026/5発表) / FY3/27: 会社予想(94,000) / FY3/28 Base: markdown シナリオ(+2.3%)",
        importance="★",
    ))
    rows.append(Row(
        label="増収率 (YoY)",
        indent=1,
        values=[None, 0.007, 0.019, 0.017, 0.024, 0.023],
        note="連結売上の前年比。本体事業は成熟、年率1-3%の低成長",
        number_format="0.0%",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【セグメント別売上】
    # ===========================
    rows.append(Row(label="【セグメント別売上 (連結)】", is_section_header=True))

    rows.append(Row(
        label="ベルト・ゴム製品事業",
        indent=0,
        values=[28_600, 29_535, 29_684, 30_597, 31_500, 32_000],
        is_independent=True,
        note="yuho セグメント情報。物流業界向けベルトが堅調。FY3/27+は会社想定+α",
        importance="★",
    ))
    rows.append(Row(
        label="ホース・チューブ製品事業",
        indent=0,
        values=[33_251, 31_697, 31_518, 32_983, 33_500, 34_000],
        is_independent=True,
        note="yuho。半導体製造装置向けが期末にかけて回復。FY3/26で+4.7%",
        importance="★",
    ))
    rows.append(Row(
        label="化工品事業",
        indent=0,
        values=[11_597, 11_822, 13_029, 11_681, 12_000, 12_300],
        is_independent=True,
        note="yuho。FY3/25は半導体回復で+10.2%、FY3/26は反動",
    ))
    rows.append(Row(
        label="その他産業用製品事業",
        indent=0,
        values=[10_449, 11_475, 11_527, 11_739, 12_000, 12_200],
        is_independent=True,
        note="yuho。空調製品(半導体クリーンルーム向け)を含む",
    ))
    rows.append(Row(
        label="不動産・経営指導・その他",
        indent=0,
        values=[4_100, 4_076, 4_516, 4_831, 5_000, 5_700],
        is_independent=True,
        note="不動産+経営指導+運転教習等。経営指導は関連会社業績連動",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【セグメント利益】
    # ===========================
    rows.append(Row(label="【セグメント利益】", is_section_header=True))

    rows.append(Row(
        label="ベルト・ゴム",
        indent=0,
        values=[3_122, 3_438, 3_473, 3_467, 3_500, 3_700],
        is_independent=True,
        note="yuho セグメント。安定的にOPM 10%超",
    ))
    rows.append(Row(
        label="ホース・チューブ",
        indent=0,
        values=[935, 7, 147, 1_073, 1_500, 1_900],
        is_independent=True,
        note="yuho。FY3/24は原料高で大幅減益、価格転嫁進捗で回復中",
        importance="★",
    ))
    rows.append(Row(
        label="化工品",
        indent=0,
        values=[202, 454, 1_015, 929, 1_000, 1_100],
        is_independent=True,
        note="yuho。半導体ミックス改善",
    ))
    rows.append(Row(
        label="その他産業用 + 不動産 + 経営指導 + その他",
        indent=0,
        values=[2_293, 2_267, 2_478, 2_506, 2_500, 2_700],
        is_independent=True,
        note="yuho セグメント合算。半導体クリーンルーム空調、賃貸、関連会社経営指導等",
    ))
    rows.append(Row(
        label="全社費用・調整",
        indent=0,
        values=[-1_566, -1_749, -1_960, -2_114, -2_200, -2_200],
        is_independent=True,
        note="本社管理部門費用等。売上拡大に伴い緩やかに増加",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【売上原価・粗利】
    # ===========================
    rows.append(Row(label="【売上原価・粗利 (連結)】", is_section_header=True))

    rows.append(Row(
        label="売上原価",
        indent=0,
        values=[65_537, 66_278, 66_030, 65_938, 67_000, 68_500],
        note="yuho 連結P&L。FY3/26は決算説明資料(売上総利益 25,896百万円から逆算)",
    ))
    rows.append(Row(
        label="売上総利益",
        indent=0,
        values=[22_463, 22_331, 24_246, 25_896, 27_000, 27_700],
        note="売上高 - 売上原価",
        importance="★",
    ))
    rows.append(Row(
        label="粗利率",
        indent=1,
        values=[0.2553, 0.2520, 0.2686, 0.2820, 0.2872, 0.2880],
        note="原材料価格転嫁の進捗、半導体製造装置向けプロダクトミックス改善で26→28%に上昇",
        number_format="0.0%",
        importance="★★",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【販管費】
    # ===========================
    rows.append(Row(label="【販管費 (連結)】", is_section_header=True))

    rows.append(Row(
        label="販管費 合計",
        indent=0,
        values=[17_474, 17_910, 19_090, 20_034, 20_800, 20_200],
        note="yuho P&L。FY3/26は決算説明資料(売上総利益25,896 - 営業利益5,862)",
    ))
    rows.append(Row(
        label="運賃及び賃借料",
        indent=1,
        values=[2_932, 2_549, 2_631, 2_700, 2_800, 2_900],
        is_independent=True,
        note="yuho 注記。FY3/26以降は売上比例で推計",
    ))
    rows.append(Row(
        label="従業員給料及び賞与",
        indent=1,
        values=[5_597, 5_669, 5_967, 6_200, 6_400, 6_500],
        is_independent=True,
        note="yuho 注記。年率+3〜5%で上昇トレンド",
    ))
    rows.append(Row(
        label="賞与引当金繰入額",
        indent=1,
        values=[589, 525, 578, 600, 620, 640],
        is_independent=True,
        note="yuho 注記。給料に概ね比例",
    ))
    rows.append(Row(
        label="退職給付費用",
        indent=1,
        values=[200, 230, 165, 200, 200, 200],
        is_independent=True,
        note="yuho 注記。年度変動あるが据置想定",
    ))
    rows.append(Row(
        label="貸倒引当金繰入額",
        indent=1,
        values=[43, -2, 1, 0, 0, 0],
        is_independent=True,
        note="yuho 注記。重要性低し",
    ))
    rows.append(Row(
        label="研究開発費",
        indent=1,
        values=[1_825, 2_058, 2_027, 2_283, 2_000, 2_200],
        is_independent=True,
        note="yuho 注記。FY3/27会社予想2,000、FY3/28はBase想定",
    ))
    rows.append(Row(
        label="その他販管費",
        indent=1,
        values=[6_288, 6_881, 7_722, 8_051, 8_780, 7_760],
        note="販管費合計 - 上記主要費目。減価償却・広告・出張等",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【営業利益】
    # ===========================
    rows.append(Row(label="【営業利益 (連結)】", is_section_header=True))

    rows.append(Row(
        label="営業利益",
        indent=0,
        values=[4_989, 4_421, 5_155, 5_862, 6_200, 7_500],
        note="売上総利益 - 販管費 / FY3/27は会社予想 / FY3/28はBaseシナリオ(+27.9%)",
        importance="★★",
    ))
    rows.append(Row(
        label="営業利益率",
        indent=1,
        values=[0.0567, 0.0499, 0.0571, 0.0638, 0.0660, 0.0780],
        note="SHIFT2030フェーズ2の2028年3月期目標が7.0%。Baseシナリオは目標達成を想定",
        number_format="0.0%",
        importance="★★",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【持分法投資利益】★★★ — 核心ドライバー
    # ===========================
    rows.append(Row(label="【持分法投資利益 (★★★ 核心ドライバー)】", is_section_header=True))

    rows.append(Row(
        label="持分法投資利益 合計",
        indent=0,
        values=[6_894, 7_001, 8_669, 8_592, 9_300, 11_300],
        note="連結P&L営業外収益。経常利益の約60%を占める。FY3/28 Baseは markdown想定値",
        importance="★★★",
    ))

    # 内訳
    rows.append(Row(
        label="うち NDI (ニッタ・デュポン:半導体研磨パッド)",
        indent=1,
        values=[3_000, 3_200, 4_005, 4_509, 5_200, 5_600],
        is_independent=True,
        note="FY3/25-26: 決算説明資料P.10〜11 四半期推移から合算 / FY3/23-24: 過去半導体需給と総持分法利益から逆算推定 / FY3/27+: Baseシナリオ",
        importance="★★★",
    ))
    rows.append(Row(
        label="うち GUA (ゲイツユニッタアジア:自動車・産業向けベルト)",
        indent=1,
        values=[3_894, 3_801, 4_724, 4_264, 4_300, 4_400],
        is_independent=True,
        note="FY3/25-26: 決算説明資料 / FY3/23-24: 総持分法利益からNDI推定値を控除して逆算 / EV化で中期横ばい想定",
        importance="★★",
    ))
    rows.append(Row(
        label="うち その他 (GNB等)",
        indent=1,
        values=[0, 0, -60, -181, -200, -200],
        is_independent=True,
        note="ゲイツニッタベルトカンパニー(米)等。FY3/26時点で赤字基調",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【NDI(ニッタ・デュポン)単体推計】 — α分解
    # ===========================
    rows.append(Row(label="【NDI(ニッタ・デュポン) 単体推計とα分解】", is_section_header=True))

    rows.append(Row(
        label="NDI売上 (CY、億円)",
        indent=0,
        values=[260, 280, 300, 327, 360, 410],
        is_independent=True,
        note="markdown: CY2025=327億円(+9%)。過去はNitta DuPont IR非開示部分は半導体出荷量と純利益率から推定。FY3/28はCAGR+13%(BSPDN/2nm本格化)",
        importance="★★★",
    ))
    rows.append(Row(
        label="ウェハ生産量 (CAGR寄与, +3〜5%)",
        indent=1,
        values=[None, None, None, None, 0.04, 0.04],
        is_independent=True,
        note="WSTS等。半導体ウェハ出荷量のマクロ成長率。AI需要で+3〜5%/年想定",
        number_format="0.0%",
        importance="★★",
    ))
    rows.append(Row(
        label="CMPステップ数/ウェハ (CAGR寄与, +4〜6%)",
        indent=1,
        values=[None, None, None, None, 0.05, 0.05],
        is_independent=True,
        note="28nm:12-15→2nm:25ステップ。BSPDN/GAA量産でα超過の最大要因。SemiEngineering根拠",
        number_format="0.0%",
        importance="★★★",
    ))
    rows.append(Row(
        label="先端品ASPミックス (CAGR寄与, +2〜3%)",
        indent=1,
        values=[None, None, None, None, 0.025, 0.025],
        is_independent=True,
        note="IC1000 → Ikonic/Emblem 世代交代。先端品単価1.5〜2倍",
        number_format="0.0%",
        importance="★★",
    ))
    rows.append(Row(
        label="世界シェア (CAGR寄与, 横ばい〜+1%)",
        indent=1,
        values=[None, None, None, None, 0.005, 0.005],
        is_independent=True,
        note="シェア66%。先端ノードでの認定優位。中国国産化リスクは中期的",
        number_format="0.0%",
    ))
    rows.append(Row(
        label="→ NDI売上CAGR (因数積)",
        indent=1,
        values=[None, None, None, None, 0.125, 0.125],
        note="(1+0.04)*(1+0.05)*(1+0.025)*(1+0.005) - 1 ≒ +12.5%。市場CAGR+7.5%を大幅上回り",
        number_format="0.0%",
        importance="★★★",
    ))

    rows.append(Row(
        label="NDI営業利益率(推定)",
        indent=0,
        values=[None, None, 0.30, 0.30, 0.32, 0.35],
        is_independent=True,
        note="CMPパッドはハイマージン消耗材。先端品ミックスでミックス改善。Base 35%(markdown想定 30〜40%レンジ中央)",
        number_format="0.0%",
        importance="★★",
    ))
    rows.append(Row(
        label="NDI純利益(推計, 億円)",
        indent=0,
        values=[None, None, 80, 90, 104, 112],
        note="NDI持分法利益 ÷ 出資比率50% で逆算 (FY3/26: 4,509百万 ÷ 0.50 = 90億)。NDI営業利益率と税率の整合チェック用",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【経常利益・純利益】
    # ===========================
    rows.append(Row(label="【経常利益・純利益 (連結)】", is_section_header=True))

    rows.append(Row(
        label="営業外収益その他 (受取利息・配当・為替差益等)",
        indent=0,
        values=[1_357, 1_110, 1_217, 1_300, 1_300, 1_300],
        note="yuho 営業外収益から持分法を除いた残り。横ばい想定",
    ))
    rows.append(Row(
        label="営業外費用 (支払利息・訴訟費用等)",
        indent=0,
        values=[-340, -525, -440, -2_944, -1_800, -1_500],
        note="yuho。FY3/26は訴訟関連費用が急増(単体経常利益が-47%減になった主因)",
        importance="★",
    ))
    rows.append(Row(
        label="経常利益",
        indent=0,
        values=[12_900, 12_007, 14_601, 14_810, 15_000, 18_800],
        note="営業利益 + 持分法 + 営外収益 - 営外費用 / FY3/27 会社予想 / FY3/28 markdown Base",
        importance="★★★",
    ))

    rows.append(Row(
        label="特別損益",
        indent=0,
        values=[-69, 2, 92, 100, -500, 0],
        is_independent=True,
        note="特別利益 - 特別損失。FY3/26は決算説明資料の経常→純利益から逆算 / FY3/27に訴訟関連の特損想定",
    ))
    rows.append(Row(
        label="税前純利益",
        indent=0,
        values=[12_831, 12_009, 14_693, 14_910, 14_500, 18_800],
        note="経常利益 + 特別損益",
    ))
    rows.append(Row(
        label="法人税等",
        indent=0,
        values=[1_908, 2_091, 2_492, 1_310, 2_200, 3_000],
        note="yuho 法人税等合計。FY3/26は決算説明資料から逆算(税前-当期純利益)。実効税率8%は持分法利益増加+訴訟費用の影響",
    ))
    rows.append(Row(
        label="実効税率",
        indent=1,
        values=[0.149, 0.174, 0.170, 0.088, 0.152, 0.160],
        is_independent=True,
        note="持分法利益(税後)が経常利益の半分超を占めるため通常より低い。FY3/26は訴訟費用で税前利益縮小により異常値",
        number_format="0.0%",
    ))
    rows.append(Row(
        label="少数株主損益",
        indent=1,
        values=[69, 59, 69, 71, 100, 0],
        note="yuho。重要性低し",
    ))
    rows.append(Row(
        label="親会社株主帰属純利益",
        indent=0,
        values=[10_853, 9_857, 12_131, 13_529, 12_300, 15_800],
        note="FY3/27 会社予想 / FY3/28 markdown Base。FY3/27会社予想は減益注意(訴訟費用想定織込)",
        importance="★★★",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【EPS・株価感応度】
    # ===========================
    rows.append(Row(label="【EPS・PER感応度・目標株価】", is_section_header=True))

    rows.append(Row(
        label="期中平均株式数 (千株)",
        indent=0,
        values=[28_021, 27_857, 27_777, 27_800, 27_800, 27_800],
        is_independent=True,
        note="yuho 1株当たり情報。FY3/25=27,777千株。自己株式控除後の平均。2024/9に100万株消却済",
    ))
    rows.append(Row(
        label="EPS (円)",
        indent=0,
        values=[387.3, 353.8, 436.7, 486.7, 442.4, 568.3],
        note="連結純利益 ÷ 期中平均株式数。FY3/28 Base=568円 (markdown525円は概算)",
        importance="★★★",
        number_format="0.0",
    ))

    # 3シナリオ PER
    rows.append(Row(
        label="想定PER Bear",
        indent=1,
        values=[None, None, None, None, None, 12.0],
        is_independent=True,
        note="ゴム製品セクター評価のまま据置。過去レンジ下限。markdown想定",
        number_format="0.0",
    ))
    rows.append(Row(
        label="想定PER Base",
        indent=1,
        values=[None, None, None, None, None, 14.0],
        is_independent=True,
        note="半導体材料企業としてのリレーティング途上。markdown想定",
        importance="★★★",
        number_format="0.0",
    ))
    rows.append(Row(
        label="想定PER Bull",
        indent=1,
        values=[None, None, None, None, None, 16.0],
        is_independent=True,
        note="ゴム製品 → 半導体材料企業として完全リレーティング。富士紡HD水準",
        number_format="0.0",
    ))

    # FY3/28 シナリオ別 純利益・EPS (markdownのシナリオから機械的計算)
    rows.append(Row(
        label="FY3/28 シナリオ別 純利益 (百万円)",
        indent=0,
        values=[None, None, None, None, None, "13,500 / 15,800 / 18,000"],
        is_independent=True,
        note="Bear / Base / Bull。markdown 4-3 シナリオ別業績表より。Base 158億 = 持分法113億 + 本体 65億",
    ))
    rows.append(Row(
        label="FY3/28 シナリオ別 EPS (円, 27,800千株ベース)",
        indent=0,
        values=[None, None, None, None, None, "486 / 568 / 648"],
        note="純利益 ÷ 期中平均27,800千株 (自己株式控除後)。markdownのEPS表記450/525/600は消却前30,272千株ベース(古い基準)",
    ))
    rows.append(Row(
        label="目標株価 Bear (円, 自社計算)",
        indent=0,
        values=[None, None, None, None, None, 5_832],
        note="486円×12倍 = 5,832円。markdown表記 5,400円(古いEPS×PER)よりも実態に近い。現在株価+8%",
        importance="★",
    ))
    rows.append(Row(
        label="目標株価 Base (円, 自社計算)",
        indent=0,
        values=[None, None, None, None, None, 7_956],
        note="568円×14倍 = 7,956円。markdown表記 7,350円(古いEPS×PER)より +8%。現在株価+47%",
        importance="★★★",
    ))
    rows.append(Row(
        label="目標株価 Bull (円, 自社計算)",
        indent=0,
        values=[None, None, None, None, None, 10_368],
        note="648円×16倍 = 10,368円。markdown表記 9,600円(古いEPS×PER)より +8%。現在株価+92%",
        importance="★★",
    ))
    rows.append(Row(
        label="(参考) markdown記載 目標株価",
        indent=1,
        values=[None, None, None, None, None, "5,400 / 7,350 / 9,600"],
        is_independent=True,
        note="markdown 4-3 シナリオ別表より。EPS分母を消却前29,272千株で計算した場合の値。自社計算より約8%低い",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【「織り込み済み?」反論ブロック】
    # ===========================
    rows.append(Row(
        label="【「織り込み済み?」反論：時価総額据置シナリオ】",
        is_section_header=True,
    ))

    rows.append(Row(
        label="時価総額 (百万円・据置)",
        indent=0,
        values=[
            MARKET_CAP_JPY // 1_000_000,
            MARKET_CAP_JPY // 1_000_000,
            MARKET_CAP_JPY // 1_000_000,
            MARKET_CAP_JPY // 1_000_000,
            MARKET_CAP_JPY // 1_000_000,
            MARKET_CAP_JPY // 1_000_000,
        ],
        is_independent=True,
        note=f"現在想定株価{STOCK_PRICE:,}円 × 自己株除発行済{SHARES_OUTSTANDING_NET/1e6:.1f}百万株 = {MARKET_CAP_JPY/1e8:.0f}億円を全年度固定",
    ))
    rows.append(Row(
        label="連結純利益 Base (百万円, 再掲)",
        indent=0,
        values=[10_853, 9_857, 12_131, 13_529, 12_300, 15_800],
        note="上記Baseシナリオの純利益を再掲",
    ))
    rows.append(Row(
        label="時価総額据置PER (倍)",
        indent=0,
        values=[13.84, 15.23, 12.38, 11.10, 12.21, 9.50],
        note="時価総額 ÷ 純利益。FY3/28に9.5倍まで機械的に低下 → 「現在価格は織り込み済み」なら 9.5倍が正当",
        number_format="0.0",
        importance="★★★",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【ピア比較 (PER)】
    # ===========================
    rows.append(Row(label="【ピア比較 (PER)】", is_section_header=True))

    rows.append(Row(
        label="ニッタ (5186) — 現状PER (FY3/26実績ベース)",
        indent=0,
        values=[11.1, None, None, None, None, None],
        is_independent=True,
        note=f"現在株価{STOCK_PRICE:,}円 / EPS486.7円 = 約11.1倍。markdown記載「10.3〜11.2倍」と整合",
        number_format="0.0",
    ))
    rows.append(Row(
        label="ピア① 富士紡HD (3104) — CMPパッド国内競合",
        indent=1,
        values=[18.0, None, None, None, None, None],
        is_independent=True,
        note="ビジネスモデル類似(CMPパッド世界2位、ニッタ・デュポンの直接競合)。PER 15〜20倍常態 → 推定18倍 ※実数値は要IR確認",
        number_format="0.0",
    ))
    rows.append(Row(
        label="ピア② バンドー化学 (5195) — ゴム製品セクター国内同業",
        indent=1,
        values=[10.0, None, None, None, None, None],
        is_independent=True,
        note="業態類似(ゴム製品セクター、ベルト主力)。市場の「現在の認識」を反映。PER 10倍前後 → 推定10倍 ※実数値は要IR確認",
        number_format="0.0",
    ))
    rows.append(Row(
        label="ピア③ 東京応化 (4186) — 半導体材料",
        indent=1,
        values=[22.0, None, None, None, None, None],
        is_independent=True,
        note="リレーティング先(半導体材料企業)。PER 18〜25倍 → 推定22倍。Bullシナリオ16倍はこれの保守版 ※実数値は要IR確認",
        number_format="0.0",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【SHIFT2030 中計目標 vs Base想定】
    # ===========================
    rows.append(Row(label="【SHIFT2030 フェーズ2 目標 vs Base想定 (FY3/28)】", is_section_header=True))

    rows.append(Row(
        label="売上高 (会社目標 / Base想定, 億円)",
        indent=0,
        values=[None, None, None, None, None, "1,150 / 962"],
        is_independent=True,
        note="会社目標1,150億は野心的(FY3/26 918→1,150は+25%)。Base想定は会社予想94,000→962億と保守的",
    ))
    rows.append(Row(
        label="営業利益率 (会社目標 / Base想定, %)",
        indent=0,
        values=[None, None, None, None, None, "7.0 / 7.8"],
        is_independent=True,
        note="Base想定は持分法成長で経常ベースも改善し、営業利益率は会社目標を上回る想定",
    ))
    rows.append(Row(
        label="事業ROIC (会社目標, %)",
        indent=0,
        values=[None, None, None, None, None, 7.0],
        is_independent=True,
        note="SHIFT2030フェーズ2新KPI。資本効率を意識した経営の証左",
        number_format="0.0",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【補助KPI・関連会社合算】
    # ===========================
    rows.append(Row(label="【関連会社合算財務 (GUA+NDI+GNB+DEMA)】", is_section_header=True))

    rows.append(Row(
        label="関連会社合算 売上高 (百万円)",
        indent=0,
        values=[None, 173_256, 206_737, 220_000, None, None],
        is_independent=True,
        note="FY3/24-25は yuho 重要な関連会社要約財務情報 / FY3/26 は持分法利益横ばいから推定 / 連結売上(918億)の2倍超の規模が「見えない」",
        importance="★",
    ))
    rows.append(Row(
        label="関連会社合算 当期純利益 (百万円)",
        indent=0,
        values=[None, 18_349, 23_371, 23_000, None, None],
        is_independent=True,
        note="FY3/24-25 yuho / FY3/26 持分法利益横ばいから推定。これに出資比率(主に50%)を乗じて持分法利益が算出",
    ))
    rows.append(Row(
        label="関連会社合算 純資産 (百万円)",
        indent=0,
        values=[None, 92_306, 112_967, 130_000, None, None],
        is_independent=True,
        note="FY3/24-25 yuho / FY3/26は内部留保増+為替で推定。BPS = 5,540円(FY3/25)で持分法投資の含み価値が大きい",
    ))

    rows.append(Row(label="", is_blank=True))

    # ===========================
    # 【リスク要因・カタリスト】
    # ===========================
    rows.append(Row(label="【リスク要因・カタリスト KPI】", is_section_header=True))

    rows.append(Row(
        label="本体単体 経常利益 (訴訟費用影響)",
        indent=0,
        values=[None, None, 7_359, 3_876, 5_500, 7_000],
        is_independent=True,
        note="FY3/26 -47.3%減益(yuho単体)。訴訟関連費用が単体に集中。FY3/27は訴訟費用減を想定、FY3/28で正常化",
        importance="★",
    ))
    rows.append(Row(
        label="海外売上高比率",
        indent=0,
        values=[None, None, 0.324, 0.332, 0.34, 0.36],
        is_independent=True,
        note="決算説明資料。北米物流ベルト堅調で上昇トレンド。中計目標は2021年比160%",
        number_format="0.0%",
    ))
    rows.append(Row(
        label="設備投資額 (百万円)",
        indent=0,
        values=[None, 2_535, 6_882, 4_445, 5_000, 4_500],
        is_independent=True,
        note="決算説明資料。FY3/25は半導体製造装置向け増産投資で急増、FY3/27会社予想5,000",
    ))
    rows.append(Row(
        label="配当性向 (%) ※単体ベース",
        indent=0,
        values=[None, 0.605, 0.615, 0.350, 0.350, 0.350],
        is_independent=True,
        note="yuho 提出会社経営指標等の配当性向(単体)。SHIFT2030フェーズ2は連結ベース30%以上、DOE2.5%以上が目標。FY3/26+は会社方針30%超を想定",
        number_format="0.0%",
    ))

    # ===========================
    # Charts
    # ===========================
    charts = [
        ChartSpec(
            title="連結 売上高・営業利益・経常利益・純利益",
            bar_rows=[
                "売上高 合計 (百万円)",
                "営業利益",
                "経常利益",
                "親会社株主帰属純利益",
            ],
            line_rows=["営業利益率"],
            anchor="J2",
        ),
        ChartSpec(
            title="持分法投資利益の分解 (GUA vs NDI)",
            bar_rows=[
                "うち GUA (ゲイツユニッタアジア:自動車・産業向けベルト)",
                "うち NDI (ニッタ・デュポン:半導体研磨パッド)",
            ],
            line_rows=[],
            anchor="J22",
        ),
        ChartSpec(
            title="目標株価シナリオ別 (FY3/28)",
            bar_rows=[
                "目標株価 Bear (円, 自社計算)",
                "目標株価 Base (円, 自社計算)",
                "目標株価 Bull (円, 自社計算)",
            ],
            line_rows=[],
            anchor="J42",
        ),
    ]

    return SheetConfig(
        name="5186_nitta_valuation",
        year_columns=YEARS,
        rows=rows,
        charts=charts,
    )


def main():
    config = build_valuation_sheet()
    out = _REPO / "outputs" / "5186-nitta-valuation.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    todos = render(config, str(out))
    print(f"\nWrote: {out}")
    print(f"Total TODOs: {len(todos)}")


if __name__ == "__main__":
    main()
