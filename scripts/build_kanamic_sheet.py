"""Build the 3939 カナミックネットワーク valuation .xlsx with Excel formulas.

設計方針:
- 独立変数(青、ベタ打ち): 各セグメント売上、ARPU、ユーザー数、AI浸透率/単価、
  粗利率、OPM、SG&A各項目、実効税率、想定PER、株価、株式数
- 従属変数(黒、Excel関数): 売上合計、増収率、売上総利益、売上原価、販管費合計、
  営業利益、経常利益、税前、法人税等、純利益、EPS、目標株価、時価総額据置PER、現状PER
- 関数化することで「粗利率を65%→67%に変えたい」が1セル変えるだけで全期波及

Data sources:
- Historical (FY23-FY25 actuals): 3939_yuho_101_20{23,24,25}0930.pdf
- FY26 会社予想: 決算説明資料(2026年5月13日発表) p9
- FY27/FY28 シナリオ: ストーリーmd
"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO = Path(__file__).parent.parent
sys.path.insert(0, str(_REPO / ".claude/skills/takimoto-valuation/scripts"))
from sheet_renderer import ChartSpec, Row, SheetConfig, render  # noqa: E402


# ===========================================================================
# 行構造プラン (ラベル/独立変数フラグ/重要度/備考をまずリスト化)
# Pass 1: 各行の row_idx を確定
# Pass 2: row_idx を使って関数を生成
# ===========================================================================

# 列構成: B〜I の 8列
#   B(idx=0)=FY23実, C(1)=FY24実, D(2)=FY25実, E(3)=FY26予, F(4)=FY27予,
#   G(5)=FY28 Bear, H(6)=FY28 Base, I(7)=FY28 Bull
YEAR_COLS = ["B", "C", "D", "E", "F", "G", "H", "I"]
N_YEARS = len(YEAR_COLS)
SHARES = 48_132  # 千株 (発行済株式総数)
PRICE_REF = 550  # 円 (2026/5/13)
MARKET_CAP_REF = round(PRICE_REF * SHARES / 1000)  # 26,473百万円


# 各行の論理キー → row_idx の対応を組み立てるためのレイアウト宣言
# (key, kind) のリスト。kind=section/blank/data。data行のみrow_idxを消費
def plan_rows():
    """全行をkey付きで列挙し、row_idx対応を返す。"""
    layout = [
        # 【売上高】
        ("h_sales", "section"),
        ("売上高 合計", "data"),
        ("増収率(YoY)", "data"),
        ("blank1", "blank"),
        # 【セグメント別売上】
        ("h_segment", "section"),
        ("医療・介護クラウドPF事業", "data"),
        ("カナミッククラウドサービス", "data"),
        ("プラットフォームサービス", "data"),
        ("その他サービス(クラウドPF内)", "data"),
        ("健康寿命延伸事業", "data"),
        ("ソリューション開発事業", "data"),
        ("blank2", "blank"),
        # 【因数分解】
        ("h_factor", "section"),
        ("有料ユーザーID数(期末、人)", "data"),
        ("増加率 YoY", "data"),
        ("クラウドPF+ソリューション 月額ARPU(円)", "data"),
        ("ARPU 上昇率 YoY", "data"),
        ("検算: ユーザー×ARPU×12 (百万円)", "data"),
        ("blank3", "blank"),
        ("subhead_ai", "subhead"),
        ("AI機能浸透率(有料IDの%)", "data"),
        ("AI機能 平均月額単価(円)", "data"),
        ("AI課金 ARPU寄与(円/月、ブレンド)", "data"),
        ("blank4", "blank"),
        ("subhead_fitness", "subhead"),
        ("アーバンフィット 店舗数(期末)", "data"),
        ("1店舗あたり年商(百万円)", "data"),
        ("blank5", "blank"),
        # 【粗利】
        ("h_gp", "section"),
        ("粗利率", "data"),
        ("売上総利益", "data"),
        ("売上原価", "data"),
        ("blank6", "blank"),
        # 【販管費】
        ("h_sga", "section"),
        ("販管費 合計", "data"),
        ("従業員給料", "data"),
        ("役員報酬", "data"),
        ("地代家賃", "data"),
        ("法定福利費", "data"),
        ("賞与引当金繰入額", "data"),
        ("減価償却費(販管費分)", "data"),
        ("貸倒引当金繰入額", "data"),
        ("その他販管費(差分)", "data"),
        ("blank7", "blank"),
        # 【利益】
        ("h_profit", "section"),
        ("営業利益率", "data"),
        ("営業利益", "data"),
        ("営業外損益(純)", "data"),
        ("経常利益", "data"),
        ("税金等調整前純利益", "data"),
        ("実効税率", "data"),
        ("法人税等", "data"),
        ("親会社株主帰属純利益", "data"),
        ("blank8", "blank"),
        # 【補助KPI】
        ("h_kpi", "section"),
        ("導入地域数(期末)", "data"),
        ("無料ユーザーID数(期末)", "data"),
        ("ユーザーID合計", "data"),
        ("EBITDA(連結減価償却足戻し)", "data"),
        ("連結減価償却費(全体、参考)", "data"),
        ("blank9", "blank"),
        # 【株価・PER感応度】
        ("h_per", "section"),
        ("発行済株式数(千株)", "data"),
        ("EPS(円)", "data"),
        ("想定PER ベスト(Bull)", "data"),
        ("想定PER ノーマル(Base)", "data"),
        ("想定PER ワースト(Bear)", "data"),
        ("目標株価 Bull(円)", "data"),
        ("目標株価 Base(円)", "data"),
        ("目標株価 Bear(円)", "data"),
        ("blank10", "blank"),
        # 【「織り込み済み?」反論】
        ("h_pricedin", "section"),
        ("時価総額(百万円、現値据置)", "data"),
        ("連結純利益 Base(再掲)", "data"),
        ("時価総額据置PER(倍)", "data"),
        ("blank11", "blank"),
        # 【ピア比較】
        ("h_peer", "section"),
        ("カナミック(3939) 現状PER", "data"),
        ("ピア① エス・エム・エス(2175)", "data"),
        ("ピア② ラクス(3923)", "data"),
        ("ピア③ JMDC(4483)", "data"),
        ("ピア中央値(参考)", "data"),
    ]
    # NOTE: renderer increments row_idx for blank rows too, so blank "consumes" a row
    pos = {}
    idx = 2  # row 1 = header
    for key, kind in layout:
        if kind in ("data", "section", "subhead"):
            pos[key] = idx
        idx += 1  # always increment (blank rows occupy a sheet row)
    return layout, pos


def cell(pos, key, year_idx):
    """Return Excel cell reference like 'B11' for the given (key, year_idx)."""
    return f"{YEAR_COLS[year_idx]}{pos[key]}"


def col_range(pos, key, start_idx=0, end_idx=None):
    """Return Excel range like B11:I11 for a single row across columns."""
    if end_idx is None:
        end_idx = N_YEARS - 1
    return f"{YEAR_COLS[start_idx]}{pos[key]}:{YEAR_COLS[end_idx]}{pos[key]}"


# ===========================================================================
# 独立変数(青)の値テーブル — ここを変えるとシートが更新される
# ===========================================================================

# 実績は yuho 開示値、予想は story md or 推計。
# 予想列のセグメント別売上は AI課金浸透率と ARPU 上昇を反映した我々の推計
INPUTS = {
    # ==== セグメント別売上 (百万円) ====
    "カナミッククラウドサービス": [2_402, 2_853, 2_997, 3_400, 3_900, 3_600, 4_500, 5_700],
    "プラットフォームサービス":   [300, 424, 474, 550, 620, 550, 700, 850],
    "その他サービス(クラウドPF内)": [147, 88, 112, 170, 180, 150, 200, 250],
    "健康寿命延伸事業":            [846, 1_131, 1_188, 1_300, 1_500, 1_400, 1_600, 1_900],
    "ソリューション開発事業":      [51, 512, 730, 930, 1_000, 800, 1_100, 1_400],
    # ==== 因数分解 (KPI) ====
    "有料ユーザーID数(期末、人)": [155_516, 191_813, 221_822, 240_000, 260_000,
                                   260_000, 280_000, 310_000],
    "クラウドPF+ソリューション 月額ARPU(円)":
        [1_350, 1_500, 1_600, 1_660, 1_750, 1_560, 1_850, 2_100],
    "AI機能浸透率(有料IDの%)": [0.0, 0.0, 0.0, 0.05, 0.15, 0.10, 0.27, 0.42],
    "AI機能 平均月額単価(円)": [None, None, None, 800, 900, 700, 1_000, 1_200],
    "アーバンフィット 店舗数(期末)": [13, 20, 24, 30, 35, 32, 36, 42],
    "1店舗あたり年商(百万円)": [65, 57, 49, 43, 43, 44, 44, 45],
    # ==== 粗利率・OPM (独立、シナリオ別仮定) ====
    # 実績は actual_粗利/actual_売上 から精度高く設定
    "粗利率":  [0.6925, 0.6135, 0.6428, 0.650, 0.660, 0.645, 0.670, 0.690],
    "営業利益率": [0.2926, 0.2874, 0.2922, 0.3228, 0.330, 0.300, 0.340, 0.370],
    # ==== SG&A 各項目 (百万円) ====
    "従業員給料":     [353, 381, 486, 580, 670, 620, 760, 950],
    "役員報酬":       [241, 251, 256, 270, 285, 280, 295, 310],
    "地代家賃":       [85, 92, 85, 90, 95, 95, 100, 110],
    "法定福利費":     [77, 82, 99, 110, 125, 120, 140, 170],
    "賞与引当金繰入額": [37, 40, 51, 60, 70, 65, 80, 100],
    "減価償却費(販管費分)": [44, 46, 47, 50, 55, 55, 60, 70],
    "貸倒引当金繰入額": [4, 6, 1, 3, 3, 3, 3, 3],
    # ==== 営業外損益・税率 ====
    "営業外損益(純)":   [11, 9, 6, 5, 5, 5, 5, 5],
    "実効税率":         [0.3088, 0.3198, 0.3108, 0.300, 0.300, 0.300, 0.300, 0.300],
    # ==== 補助KPI ====
    "導入地域数(期末)":      [1_270, 1_320, 1_397, 1_440, 1_490, 1_460, 1_510, 1_580],
    "無料ユーザーID数(期末)": [86_760, 101_309, 117_497, 135_000, 152_000, 145_000, 158_000, 175_000],
    "連結減価償却費(全体、参考)": [268, 301, 332, 450, 524, 450, 550, 663],
    # ==== 株価・PER ====
    "発行済株式数(千株)": [SHARES] * N_YEARS,
    # 想定PER は Bear/Base/Bull のそれぞれの列のみセット
    "想定PER ベスト(Bull)": [None]*7 + [22.0],
    "想定PER ノーマル(Base)": [None]*6 + [19.0, None],
    "想定PER ワースト(Bear)": [None]*5 + [16.0, None, None],
    "時価総額(百万円、現値据置)": [MARKET_CAP_REF] * N_YEARS,
    "カナミック(3939) 現状PER": [None, None, None, 19.3, None, None, None, None],
    "ピア① エス・エム・エス(2175)": [None, None, None, "TODO", None, None, None, None],
    "ピア② ラクス(3923)":         [None, None, None, "TODO", None, None, None, None],
    "ピア③ JMDC(4483)":           [None, None, None, "TODO", None, None, None, None],
    "ピア中央値(参考)":           [None, None, None, "TODO", None, None, None, None],
    # 増収率の FY24-FY27 は formula で計算。Bear/Base/Bull は YoY概念外 → 空白
    # AI課金 ARPU寄与は formula
    # 売上関連は全部 formula
}


# ===========================================================================
# 行構築 (rows リストを作る)
# ===========================================================================

def build() -> SheetConfig:
    layout, pos = plan_rows()
    rows: list[Row] = []

    # 各 layout 項目を対応する Row オブジェクトに変換
    for key, kind in layout:
        if kind == "blank":
            rows.append(Row(label="", is_blank=True))
            continue
        if kind == "section":
            rows.append(Row(label=SECTION_HEADERS[key], is_section_header=True))
            continue
        if kind == "subhead":
            rows.append(Row(label=SUBHEADS[key], indent=0))
            continue
        # data 行
        rows.append(make_data_row(key, pos))

    # ===========================
    # Charts
    # ===========================
    charts = [
        ChartSpec(
            title="売上・粗利・営業利益 と OPM",
            bar_rows=["売上高 合計", "売上総利益", "営業利益"],
            line_rows=["営業利益率"],
            anchor="K2",
        ),
        ChartSpec(
            title="有料ユーザー数 と 月額ARPU",
            bar_rows=["有料ユーザーID数(期末、人)"],
            line_rows=["クラウドPF+ソリューション 月額ARPU(円)"],
            anchor="K22",
        ),
        ChartSpec(
            title="セグメント別売上",
            bar_rows=["医療・介護クラウドPF事業", "健康寿命延伸事業", "ソリューション開発事業"],
            line_rows=["増収率(YoY)"],
            anchor="K42",
        ),
        ChartSpec(
            title="時価総額据置PERの低下 (織り込み済み反論)",
            bar_rows=["連結純利益 Base(再掲)"],
            line_rows=["時価総額据置PER(倍)"],
            anchor="K62",
        ),
    ]

    year_columns = [
        "FY23/9 実", "FY24/9 実", "FY25/9 実",
        "FY26/9 予(会社)", "FY27/9 予(Base)",
        "FY28/9 Bear", "FY28/9 Base", "FY28/9 Bull",
    ]
    return SheetConfig(
        name="3939_valuation",
        year_columns=year_columns,
        rows=rows,
        charts=charts,
    )


# ===========================================================================
# セクションヘッダ / サブヘッダのラベル
# ===========================================================================
SECTION_HEADERS = {
    "h_sales": "【売上高】",
    "h_segment": "【セグメント別売上】",
    "h_factor": "【因数分解: 有料ユーザー数 × 月額ARPU × 12ヶ月】",
    "h_gp": "【粗利】",
    "h_sga": "【販管費】",
    "h_profit": "【利益】",
    "h_kpi": "【補助KPI: SaaS指標】",
    "h_per": "【株価・PER感応度】",
    "h_pricedin": "【「織り込み済み?」反論】",
    "h_peer": "【ピア比較(PER) ※純粋同業未上場、多軸選定】",
}
SUBHEADS = {
    "subhead_ai":      "  ▽ AI課金浸透(変曲点ドライバー)",
    "subhead_fitness": "  ▽ 健康寿命延伸サブドライバー",
}


# ===========================================================================
# Row 生成: 独立変数 or 関数 を判定して values を作る
# ===========================================================================

# 独立変数(青、ベタ打ち)行の備考と重要度
INDEP_META = {
    "カナミッククラウドサービス": ("★★★",
        "ストックビジネス。FY25+5.1%実績、FY24+18.8%(M&A)。AI課金の主受益ライン"),
    "プラットフォームサービス": ("",
        "HP構築/介護労働安定センター受託/インターネット広告。FY25+11.9%"),
    "その他サービス(クラウドPF内)": ("",
        "大口顧客向けカスタマイズ開発受託。ボラあり"),
    "健康寿命延伸事業": ("★★",
        "アーバンフィット24(FY25末24店舗→Vision 2035で年4-6店舗増想定)"),
    "ソリューション開発事業": ("★",
        "Ruby開発(2023/8連結)+TWM(2025/2連結)。FY25+42.6%"),
    "有料ユーザーID数(期末、人)": ("★★★",
        "決算説明資料1-8。FY26末はFY26-2Q時点247,836→+8%伸び想定。Base CAGR+9%"),
    "クラウドPF+ソリューション 月額ARPU(円)": ("★★★",
        "(クラウドPF+ソリューション)÷ユーザー数(年央)÷12 から逆算。AI課金がドライバー"),
    "AI機能浸透率(有料IDの%)": ("★★★",
        "FY26課金開始(AI100-Vision)。Base FY28で25-30%、Bull 40%超。2027改定で加算要件化なら非線形"),
    "AI機能 平均月額単価(円)": ("★★",
        "未開示。ストーリーmdで月額500-1,500円帯。Base 1,000円で置く"),
    "アーバンフィット 店舗数(期末)": ("★",
        "2026/3末で27店舗(2026年5月時点)。Vision 2035で年4-6店舗増"),
    "1店舗あたり年商(百万円)": ("",
        "健康寿命延伸事業売上÷期中平均店舗数。新規出店は当初低稼働、徐々に45百万へ収斂"),
    "粗利率": ("★★",
        "FY23-25は actual_粗利/actual_売上 から計算。FY24はM&A連結で一時悪化。AI限界利益≒100%で漸増"),
    "営業利益率": ("★★",
        "FY23-25は actual_営利/actual_売上。Vision 2035で2030年33-35%目標、FY28 Base 34%は妥当"),
    "従業員給料": ("",
        "yuho販管費注記。FY25+27.6%は新規連結とTWM連結(2025/2)。SaaS人件費は売上に劣後成長"),
    "役員報酬": ("", "yuho販管費注記。安定推移"),
    "地代家賃": ("", "本社+営業所(7拠点)。FY25は再配置で減"),
    "法定福利費": ("", "従業員給料に概ね連動"),
    "賞与引当金繰入額": ("", "業績連動。OPM拡大局面で増額"),
    "減価償却費(販管費分)": ("", "販管費分のみ。連結全体はB/Sの減価償却費を別途参照"),
    "貸倒引当金繰入額": ("", "据え置き(過去レンジ1-6百万)"),
    "営業外損益(純)": ("", "受取利息/手数料等。安定して+5〜10百万円"),
    "実効税率": ("", "FY23-25 actual 30.9-32.0%。法定実効税率約30%。FY26+は30%固定"),
    "導入地域数(期末)": ("",
        "決算説明資料1-7。中学校区(人口3万人圏)単位。FY26 2Q時点1,420地域"),
    "無料ユーザーID数(期末)": ("",
        "無料→有料転換のパイプライン。介護情報基盤稼働で増加加速可能性"),
    "連結減価償却費(全体、参考)": ("",
        "B/S(連結CF表)。販管費分+原価分の合計。FY25は332百万"),
    "発行済株式数(千株)": ("",
        "yuho開示。直近5年48,132千株で据置(自己株消却・新株発行なし)。"
        "EPS算定上は期中平均≒47,460(自己株差)だが簡略化のため期末で固定"),
    "想定PER ベスト(Bull)": ("",
        "Bullシナリオ:AI浸透確認でプレミアム。SaaS OPM30%超で18-25倍が相場、+α"),
    "想定PER ノーマル(Base)": ("",
        "Baseシナリオ:現状PER(19.3倍)維持。SaaS中央値かつカナミック直近の典型水準"),
    "想定PER ワースト(Bear)": ("",
        "Bearシナリオ:AI浸透遅延でディスカウント。過去レンジ下限14.9倍に若干上乗せ"),
    "時価総額(百万円、現値据置)": ("",
        "2026/5/13株価550円 × 48,132千株 = 26,473百万円で全期間固定"),
    "カナミック(3939) 現状PER": ("",
        "株価550円 ÷ FY26予EPS 28.46円 ≒ 19.3倍(2026/5/13)"),
    "ピア① エス・エム・エス(2175)": ("",
        "医療介護領域のヘルスケアSaaS+人材PF。軸:同業界(介護DX)。直近実績PERをIR確認推奨"),
    "ピア② ラクス(3923)": ("",
        "中小事業者向け低単価SaaS(楽楽精算等)。軸:ビジネスモデル類似(低単価高浸透)。"
        "高成長SaaSプレミアム参考"),
    "ピア③ JMDC(4483)": ("",
        "医療データプラットフォーム。軸:制度受益(医療DX政策・基盤)+データMoat"),
    "ピア中央値(参考)": ("",
        "プライム情報通信SaaS(OPM30%超・増益率20%超)のPER相場18-25倍を参考"),
}

# 関数行(黒、=式)の備考と重要度
DEP_META = {
    "売上高 合計": ("",
        "=セグメント3区分の合計。FY23のクラウドPF事業内訳は健康寿命延伸を控除した値"),
    "増収率(YoY)": ("",
        "=(今期売上÷前期売上)−1。FY28 3シナリオはYoY概念外につき空白"),
    "医療・介護クラウドPF事業": ("★★★",
        "=カナミッククラウド+プラットフォーム+その他(クラウドPF内)。yuho開示と一致"),
    "検算: ユーザー×ARPU×12 (百万円)": ("",
        "=有料ユーザー × ARPU × 12 ÷ 1,000,000。クラウドPF+ソリューション売上と概ね一致を確認"),
    "増加率 YoY": ("",
        "=(今期ユーザー÷前期ユーザー)−1。FY28 3シナリオはYoY概念外につき空白"),
    "ARPU 上昇率 YoY": ("",
        "=(今期ARPU÷前期ARPU)−1。FY28 BaseはAI課金浸透25-30%でARPU+15.6%、Bull 40%超で+31%超"),
    "AI課金 ARPU寄与(円/月、ブレンド)": ("",
        "=AI浸透率 × AI平均単価。Base FY28: 27%×1,000=270円が既存ARPUに上乗せ"),
    "売上総利益": ("",
        "=売上高 × 粗利率"),
    "売上原価": ("",
        "=売上高 − 売上総利益"),
    "販管費 合計": ("",
        "=売上総利益 − 営業利益 (=粗利−営利)。SG&A項目はSUM参考のみで強制せず"),
    "その他販管費(差分)": ("",
        "=販管費合計 − 上記7項目のSUM。外注費/広告/業務委託等"),
    "営業利益": ("",
        "=売上高 × 営業利益率"),
    "経常利益": ("",
        "=営業利益 + 営業外損益(純)"),
    "税金等調整前純利益": ("",
        "=経常利益 (FY24は特別損失104百万で経常から減少、参考値1,352)"),
    "法人税等": ("",
        "=税前 × 実効税率"),
    "親会社株主帰属純利益": ("",
        "=税前 − 法人税等。FY26会社予想は1,370(28.5円)、ほぼ一致"),
    "ユーザーID合計": ("",
        "=有料ユーザーID + 無料ユーザーID"),
    "EBITDA(連結減価償却足戻し)": ("",
        "=営業利益 + 連結減価償却費"),
    "EPS(円)": ("",
        "=純利益 × 1,000 ÷ 発行済株式数(千株)。FY28 Base 40.1円 ≒ ストーリー40.0円"),
    "目標株価 Bull(円)": ("",
        "=FY28 Bull EPS × Bull想定PER。Bull EPS 54.5 × 22倍 ≒ 1,199円(現値+118%)"),
    "目標株価 Base(円)": ("",
        "=FY28 Base EPS × Base想定PER。Base EPS 40.1 × 19倍 ≒ 762円(現値+38%)"),
    "目標株価 Bear(円)": ("",
        "=FY28 Bear EPS × Bear想定PER。Bear EPS 28.5 × 16倍 ≒ 455円(現値-17%)"),
    "連結純利益 Base(再掲)": ("",
        "=親会社株主帰属純利益(上記P&Lより再掲)"),
    "時価総額据置PER(倍)": ("",
        "=時価総額 ÷ 純利益。株価据置で利益増ならPER低下→現値に純利益成長は織込まれていない"),
}


def make_data_row(key: str, pos: dict) -> Row:
    """key に応じて独立変数(ベタ打ち)か関数行かを判定し Row を生成。"""

    # === 独立変数(青) ===
    if key in INDEP_META:
        importance, note = INDEP_META[key]
        return Row(
            label=key,
            indent=indent_for(key),
            values=INPUTS[key],
            is_independent=True,
            importance=importance,
            note=note,
            number_format=number_format_for(key),
        )

    # === 関数行(黒) ===
    importance, note = DEP_META[key]
    values = formula_values_for(key, pos)
    return Row(
        label=key,
        indent=indent_for(key),
        values=values,
        is_independent=False,
        importance=importance,
        note=note,
        number_format=number_format_for(key),
    )


def indent_for(key: str) -> int:
    """ラベルに応じてインデントレベルを返す。"""
    if key in ("カナミッククラウドサービス", "プラットフォームサービス",
               "その他サービス(クラウドPF内)"):
        return 1
    if key in ("増収率(YoY)", "増加率 YoY", "ARPU 上昇率 YoY",
               "AI機能浸透率(有料IDの%)", "AI機能 平均月額単価(円)",
               "AI課金 ARPU寄与(円/月、ブレンド)",
               "アーバンフィット 店舗数(期末)", "1店舗あたり年商(百万円)",
               "粗利率", "営業利益率",
               "従業員給料", "役員報酬", "地代家賃", "法定福利費",
               "賞与引当金繰入額", "減価償却費(販管費分)", "貸倒引当金繰入額",
               "その他販管費(差分)",
               "営業外損益(純)", "実効税率",
               "想定PER ベスト(Bull)", "想定PER ノーマル(Base)", "想定PER ワースト(Bear)",
               "連結純利益 Base(再掲)", "時価総額据置PER(倍)",
               "ピア① エス・エム・エス(2175)", "ピア② ラクス(3923)",
               "ピア③ JMDC(4483)", "ピア中央値(参考)"):
        return 1
    return 0


def number_format_for(key: str) -> str:
    """ラベルに応じた number_format を返す。"""
    if key in ("増収率(YoY)", "増加率 YoY", "ARPU 上昇率 YoY",
               "粗利率", "営業利益率", "実効税率", "AI機能浸透率(有料IDの%)"):
        return "0.0%"
    if key in ("EPS(円)",):
        return "0.00"
    if key.startswith("想定PER") or "PER" in key:
        return "0.0"
    if key in ("有料ユーザーID数(期末、人)", "無料ユーザーID数(期末)", "ユーザーID合計",
               "導入地域数(期末)", "アーバンフィット 店舗数(期末)"):
        return "#,##0"
    if key == "1店舗あたり年商(百万円)":
        return "0.0"
    if key in ("目標株価 Bull(円)", "目標株価 Base(円)", "目標株価 Bear(円)",
               "クラウドPF+ソリューション 月額ARPU(円)", "AI機能 平均月額単価(円)",
               "AI課金 ARPU寄与(円/月、ブレンド)"):
        return "#,##0"
    return "#,##0"


def formula_values_for(key: str, pos: dict) -> list:
    """関数行の values リスト(各年度=Excel formula 文字列)を返す。"""
    out = []
    for y in range(N_YEARS):

        if key == "売上高 合計":
            f = (f"={cell(pos,'医療・介護クラウドPF事業',y)}"
                 f"+{cell(pos,'健康寿命延伸事業',y)}"
                 f"+{cell(pos,'ソリューション開発事業',y)}")
            out.append(f)

        elif key == "増収率(YoY)":
            # FY23(y=0): 空白。FY28 3シナリオ(y=5,6,7)はYoY概念外。それ以外はYoY。
            if y in (0, 5, 6, 7):
                out.append(None)
            else:
                cur = cell(pos, "売上高 合計", y)
                prev = cell(pos, "売上高 合計", y - 1)
                out.append(f"={cur}/{prev}-1")

        elif key == "医療・介護クラウドPF事業":
            f = (f"={cell(pos,'カナミッククラウドサービス',y)}"
                 f"+{cell(pos,'プラットフォームサービス',y)}"
                 f"+{cell(pos,'その他サービス(クラウドPF内)',y)}")
            out.append(f)

        elif key == "検算: ユーザー×ARPU×12 (百万円)":
            users = cell(pos, "有料ユーザーID数(期末、人)", y)
            arpu = cell(pos, "クラウドPF+ソリューション 月額ARPU(円)", y)
            out.append(f"={users}*{arpu}*12/1000000")

        elif key == "増加率 YoY":
            if y in (0, 5, 6, 7):
                out.append(None)
            else:
                cur = cell(pos, "有料ユーザーID数(期末、人)", y)
                prev = cell(pos, "有料ユーザーID数(期末、人)", y - 1)
                out.append(f"={cur}/{prev}-1")

        elif key == "ARPU 上昇率 YoY":
            if y in (0, 5, 6, 7):
                out.append(None)
            else:
                cur = cell(pos, "クラウドPF+ソリューション 月額ARPU(円)", y)
                prev = cell(pos, "クラウドPF+ソリューション 月額ARPU(円)", y - 1)
                out.append(f"={cur}/{prev}-1")

        elif key == "AI課金 ARPU寄与(円/月、ブレンド)":
            # 浸透率 × 単価。単価が None(空)なら 0
            penet = cell(pos, "AI機能浸透率(有料IDの%)", y)
            unit = cell(pos, "AI機能 平均月額単価(円)", y)
            # IFERRORで単価ない年は0扱い
            out.append(f"=IFERROR({penet}*{unit},0)")

        elif key == "売上総利益":
            sales = cell(pos, "売上高 合計", y)
            gpr = cell(pos, "粗利率", y)
            out.append(f"={sales}*{gpr}")

        elif key == "売上原価":
            sales = cell(pos, "売上高 合計", y)
            gp = cell(pos, "売上総利益", y)
            out.append(f"={sales}-{gp}")

        elif key == "販管費 合計":
            gp = cell(pos, "売上総利益", y)
            op = cell(pos, "営業利益", y)
            out.append(f"={gp}-{op}")

        elif key == "その他販管費(差分)":
            sga_total = cell(pos, "販管費 合計", y)
            items = ["従業員給料", "役員報酬", "地代家賃", "法定福利費",
                     "賞与引当金繰入額", "減価償却費(販管費分)", "貸倒引当金繰入額"]
            subtract = "-".join(cell(pos, it, y) for it in items)
            out.append(f"={sga_total}-{subtract}")

        elif key == "営業利益":
            sales = cell(pos, "売上高 合計", y)
            opm = cell(pos, "営業利益率", y)
            out.append(f"={sales}*{opm}")

        elif key == "経常利益":
            op = cell(pos, "営業利益", y)
            non = cell(pos, "営業外損益(純)", y)
            out.append(f"={op}+{non}")

        elif key == "税金等調整前純利益":
            out.append(f"={cell(pos,'経常利益',y)}")

        elif key == "法人税等":
            tax_pre = cell(pos, "税金等調整前純利益", y)
            rate = cell(pos, "実効税率", y)
            out.append(f"={tax_pre}*{rate}")

        elif key == "親会社株主帰属純利益":
            tax_pre = cell(pos, "税金等調整前純利益", y)
            tax = cell(pos, "法人税等", y)
            out.append(f"={tax_pre}-{tax}")

        elif key == "ユーザーID合計":
            p = cell(pos, "有料ユーザーID数(期末、人)", y)
            f = cell(pos, "無料ユーザーID数(期末)", y)
            out.append(f"={p}+{f}")

        elif key == "EBITDA(連結減価償却足戻し)":
            op = cell(pos, "営業利益", y)
            dep = cell(pos, "連結減価償却費(全体、参考)", y)
            out.append(f"={op}+{dep}")

        elif key == "EPS(円)":
            ni = cell(pos, "親会社株主帰属純利益", y)
            shares = cell(pos, "発行済株式数(千株)", y)
            out.append(f"={ni}*1000/{shares}")

        elif key == "目標株価 Bull(円)":
            if y == 7:  # Bull列のみ
                eps = cell(pos, "EPS(円)", 7)
                per = cell(pos, "想定PER ベスト(Bull)", 7)
                out.append(f"={eps}*{per}")
            else:
                out.append(None)

        elif key == "目標株価 Base(円)":
            if y == 6:
                eps = cell(pos, "EPS(円)", 6)
                per = cell(pos, "想定PER ノーマル(Base)", 6)
                out.append(f"={eps}*{per}")
            else:
                out.append(None)

        elif key == "目標株価 Bear(円)":
            if y == 5:
                eps = cell(pos, "EPS(円)", 5)
                per = cell(pos, "想定PER ワースト(Bear)", 5)
                out.append(f"={eps}*{per}")
            else:
                out.append(None)

        elif key == "連結純利益 Base(再掲)":
            out.append(f"={cell(pos,'親会社株主帰属純利益',y)}")

        elif key == "時価総額据置PER(倍)":
            mc = cell(pos, "時価総額(百万円、現値据置)", y)
            ni = cell(pos, "親会社株主帰属純利益", y)
            out.append(f"={mc}/{ni}")

        else:
            out.append(None)

    return out


def main():
    config = build()
    out = Path(__file__).parent.parent / "outputs" / "3939-kanamic-valuation.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    todos = render(config, str(out))
    print(f"\nWrote: {out}")
    print(f"Total TODOs: {len(todos)}")


if __name__ == "__main__":
    main()
