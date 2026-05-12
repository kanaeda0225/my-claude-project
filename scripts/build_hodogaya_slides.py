"""Build the 4112 Hodogaya slide deck (.pptx). Source: samples/4112-hodogaya.md"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO = Path(__file__).parent.parent
sys.path.insert(0, str(_REPO / ".claude/skills/takimoto-presentation/scripts"))
from pptx_renderer import Slide, SlideDeckConfig, render  # noqa: E402


def build_deck() -> SlideDeckConfig:
    s: list[Slide] = []

    s.append(Slide(
        kind="content", title="ストーリー",
        subtitle="MacBook Pro OLED量産 = SFC材料需要爆発、アナリスト0社",
        bullets=[
            "Appleが2026年MacBook ProにOLEDを初採用、Samsung Displayが製造",
            "SDCに青色OLEDホスト材料を供給するのがSFC(保土谷50%出資)",
            "ノートPC1台 ≒ iPhone 11-14台分の材料消費 (面積25倍 × タンデム2層)",
            "アナリスト0社・PBR 0.69倍 = 機関投資家が再評価していない",
            "「染料・化学メーカー」認識のままミスプライシング",
        ],
        footer="The Elec, Omdia, Digitimes, OLED-info, UBI Research",
    ))

    s.append(Slide(
        kind="content", title="発表を理解するための地図",
        subtitle="マクロ → ミクロ → バリュエーション の3つの問い",
        bullets=[
            "① なぜ「今」OLED材料需要が拡大するのか? (マクロ)",
            "② なぜ「保土谷」がその恩恵を受けるのか? (ミクロ)",
            "③ なぜ「今がまだ買い場」なのか? (バリュエーション)",
        ],
    ))

    s.append(Slide(
        kind="ascii", title="成長を因数分解する",
        subtitle="台数 × 面積 × 積層数 × (1/歩留) = SFC材料需要",
        ascii_art=(
            "SFCへの材料需要\n"
            "  ∝ IT-OLEDパネル枚数 × 画面の面積 × 発光層の積層数 ÷ 歩留まり\n"
            "\n"
            "市場が見ている: 「IT-OLED CAGR +37% (Omdia)」 ← 枚数しか捉えていない\n"
            "\n"
            "面積(スマホ比25倍) × 積層数(+50-75%タンデム) の累乗効果がさらに乗る\n"
            "→ SFC材料需要 推計 CAGR 45-65%\n"
        ),
    ))

    s.append(Slide(
        kind="table", title="因数分解：各因数の方向",
        subtitle="MacBook Pro 1台 ≒ iPhone 11-14台分の材料消費",
        table_headers=["因数", "方向", "なぜ動くか"],
        table=[
            ["IT-OLEDパネル枚数", "↑ CAGR 37%", "MacBook Pro / iPad Pro の普及"],
            ["画面の面積", "↑ スマホ比25倍", "スマホ→ノートPCへの大型化"],
            ["発光層積層数", "↑ +50-75% (タンデム)", "ノートPC向けは2層構造が標準"],
            ["(1/歩留まり)", "← 改善中", "A6ライン歩留85%超、目標90%"],
        ],
        footer="Omdia, Digitimes",
    ))

    s.append(Slide(
        kind="table", title="なぜ市場が気づいていないか",
        subtitle="「145年の染料メーカー」認識がSFCを覆い隠している",
        table_headers=["観点", "現状"],
        table=[
            ["アナリストカバレッジ", "0社 (みんかぶ・IFIS共に「対象外」)"],
            ["機関投資家の認識", "「145年の歴史を持つ染料・化学メーカー」"],
            ["OLED材料サプライヤー認識", "ほぼゼロ"],
        ],
    ))

    s.append(Slide(kind="toc", title="目次",
                   bullets=["1️⃣ 基本情報", "2️⃣ マクロ", "3️⃣ ミクロ",
                            "4️⃣ バリュエーション", "5️⃣ Appendix"]))

    # 1️⃣
    s.append(Slide(kind="chapter", title="1️⃣ 基本情報"))

    s.append(Slide(
        kind="table", title="会社の概要",
        subtitle="時価総額381億・PBR 0.69倍 = 純資産を下回る割安水準",
        table_headers=["項目", "内容"],
        table=[
            ["銘柄コード", "4112 (東証プライム)"],
            ["創業", "1878年(明治11年)、145年の歴史"],
            ["主事業", "有機合成化学品(染料・中間体・OLED材料)"],
            ["時価総額", "約381億円(2026/5)"],
            ["株価", "2,464円(2026/5/11)"],
            ["PBR", "0.69倍 (純資産を下回る割安)"],
            ["アナリストカバレッジ", "0社"],
        ],
    ))

    s.append(Slide(
        kind="content", title="(補足) PBRとは",
        subtitle="純資産対比の株価倍率、0.69 = 解散しても株価より戻る",
        bullets=[
            "会社の純資産(解散時に株主に戻る価値)に対して株価が何倍か",
            "0.69倍 = 今すぐ解散しても株価より多くお金が戻ってくる状態",
            "割安のサインとして使われる",
        ],
    ))

    s.append(Slide(
        kind="ascii", title="ビジネスモデルの全体像",
        subtitle="表面は化学メーカー、実態はSFC持分法利益が稼ぎ頭",
        ascii_art=(
            "保土谷化学工業 (4112)\n"
            " │\n"
            " ├─ 国内化成品事業(染料・農薬中間体など)\n"
            " │    → 安定的だが成熟、成長ドライバーではない\n"
            " │\n"
            " └─ SFC (Samsung Fine Chemicals)  ← 投資テーゼの核心\n"
            "      → 保有持分: 約50%\n"
            "      → 韓国法人、Samsung Display(SDC)と50:50の合弁\n"
            "      → 主力製品: 青色OLEDホスト材料・ドーパント\n"
            "      → 主要顧客: Samsung Display (ほぼ独占的)\n"
        ),
    ))

    s.append(Slide(
        kind="content", title="そもそもOLEDとは",
        subtitle="有機ELディスプレイ、自発光で薄く美しい",
        bullets=[
            "有機化合物の薄膜に電気を流すと自ら光る仕組み",
            "液晶(LCD)と違い自発光、薄く美しい発色",
            "iPhoneやSamsung Galaxyのスマホで主流",
        ],
    ))

    s.append(Slide(
        kind="ascii", title="OLEDの構造(超簡略版)",
        subtitle="発光層(EML)の品質を決めるのがホスト材料",
        ascii_art=(
            "[陰極(−)]\n"
            "  ↑ 電子が流れてくる\n"
            "[電子輸送層]\n"
            "┌──────────────────────────────┐\n"
            "│ 発光層(EML)                  │\n"
            "│  ・青色ホスト材料 ← SFCが供給│  ここで光が生まれる\n"
            "│  ・青色ドーパント(発光体)    │\n"
            "└──────────────────────────────┘\n"
            "[正孔輸送層]\n"
            "  ↑ 正孔が流れてくる\n"
            "[陽極(+)]\n"
        ),
    ))

    s.append(Slide(
        kind="content", title="青色ホスト材料とは(アナロジー)",
        subtitle="演者がどれだけ輝けるかは「会場」の品質で決まる",
        bullets=[
            "ホスト = コンサート会場、ドーパント = 演者(発光体)",
            "どんなに優れた演者でも、粗末な会場では実力発揮できない",
            "青色ホスト材料はOLEDの発光効率・寿命を決める「会場」",
            "→ ここをSFCが独占的に供給",
        ],
    ))

    s.append(Slide(
        kind="content", title="なぜ青色が難しいか",
        subtitle="RGB三色のうち青が最難関、OLEDの画質寿命を決定",
        bullets=[
            "青色発光体は劣化しやすく、発光効率も赤・緑より低い",
            "OLEDの画質・寿命は「青の品質」で決まると言って過言ではない",
            "SFCはこの最難関の材料を専門とする",
        ],
    ))

    # 2️⃣
    s.append(Slide(kind="toc", title="目次(マクロへ)", bullets=["▶ 2️⃣ マクロ"]))
    s.append(Slide(kind="chapter", title="2️⃣ マクロ"))

    s.append(Slide(
        kind="ascii", title="マクロ全体の流れ",
        subtitle="スマホ縮小 → IT-OLED急拡大、MacBook Pro量産が転換点",
        ascii_art=(
            "[今起きていること]\n"
            "スマホOLED市場は成熟・縮小 → 代わりにノートPC・タブレット向けOLEDが急拡大\n"
            "\n"
            "[2026年の最大イベント]\n"
            "Apple MacBook ProにOLEDが初採用 → Samsung Display A6ラインが量産開始\n"
            "\n"
            "[なぜ材料需要が「スマホの時代」とは桁違いか]\n"
            "大画面 × 発光層2枚(タンデム) → 1台あたり材料量がスマホの11-14倍\n"
        ),
    ))

    s.append(Slide(
        kind="table", title="主戦場の交代(2026年成長率)",
        subtitle="スマホは-3%、IT機器向けは+39〜61%",
        table_headers=["セグメント", "2026年成長率", "状況"],
        table=[
            ["スマホOLED材料", "-3%", "市場飽和、単価も低下"],
            ["タブレットOLEDパネル", "+39%", "iPad Pro、Galaxy Tab等"],
            ["ノートPC OLEDパネル", "+46%", "MacBook Pro OLEDが牽引"],
            ["OLEDモニター", "+61%", "ゲーミング・クリエイター向け"],
        ],
        footer="Omdia (2025-2026年データ)",
    ))

    s.append(Slide(
        kind="table", title="Samsung Display A6ラインの状況",
        subtitle="2026年5-6月量産開始、歩留85%超で安定フェーズ",
        table_headers=["項目", "内容"],
        table=[
            ["製造会社", "Samsung Display(韓国)"],
            ["工場・ライン", "天安工場 A6ライン(第8.6世代ガラス基板)"],
            ["量産開始", "2026年5〜6月にガラス投入、7月本格量産"],
            ["歩留まり", "4月下旬時点で85%超(目標90%)"],
            ["Apple向け出荷目標", "200万枚(2026年中)"],
            ["競合", "SDCがリジッド8.6G OLEDを独占供給"],
        ],
        footer="Digitimes報道、Samsung Display IR",
    ))

    s.append(Slide(
        kind="ascii", title="タンデムOLEDとは",
        subtitle="発光層2枚重ね = 材料消費が単純に倍",
        ascii_art=(
            "[従来のスマホ向けOLED(シングルスタック)]\n"
            "    [発光層 × 1枚]  ← 材料を1枚分だけ使う\n"
            "\n"
            "[ノートPC向けOLED(タンデム構造)]\n"
            "    [発光層 × 1枚]\n"
            "    [電荷生成層]    ← 2層を繋ぐ層\n"
            "    [発光層 × 1枚]  ← 合計で材料を2枚分使う\n"
        ),
    ))

    s.append(Slide(
        kind="content", title="なぜノートPCはタンデムが必要か",
        subtitle="長時間×大画面の輝度・寿命要求でタンデムが事実上必須",
        bullets=[
            "ノートPCは長時間使用、かつスマホより画面大 → 輝度・寿命要求高い",
            "タンデムで各発光層の負荷を半分に抑え、寿命が大幅に伸びる",
            "スマホOLEDをそのまま大きくしてもノートPC要求水準を満たせない",
            "→ タンデムが事実上の必須構造",
        ],
    ))

    s.append(Slide(
        kind="table", title="タンデム採用率の推移",
        subtitle="0% → 36%へ急上昇、今後も上昇見込み",
        table_headers=["時期", "採用率", "イベント"],
        table=[
            ["2023年以前", "約0%", "IT向けOLEDがほぼ存在しない"],
            ["2024年", "30%超", "iPad Pro(史上初のタンデム製品)が登場"],
            ["2026年(予測)", "36%", "MacBook Pro OLEDが加わる"],
        ],
        footer="Omdia (2025年4月)",
    ))

    s.append(Slide(
        kind="table", title="MacBook Pro 1台 vs iPhone 1台の材料消費",
        subtitle="面積25倍 × タンデム2層 = 約11-14台分の材料",
        table_headers=["比較項目", "iPhone", "MacBook Pro 16inch"],
        table=[
            ["画面面積", "約135 cm²", "約3,400 cm² (約25倍)"],
            ["OLED構造", "シングルスタック", "タンデム(2層)"],
            ["材料消費量(換算)", "1台分", "約11-14台分"],
        ],
        footer="MacBook Pro 200万台 ≒ スマホ換算 2,200-2,800万台分の材料需要",
    ))

    # 3️⃣
    s.append(Slide(kind="toc", title="目次(ミクロへ)", bullets=["▶ 3️⃣ ミクロ"]))
    s.append(Slide(kind="chapter", title="3️⃣ ミクロ"))

    s.append(Slide(
        kind="ascii", title="SFCのMoat①: 資本的一体化",
        subtitle="SDC = SFC共同出資者、切り替え動機が構造的にゼロ",
        ascii_art=(
            "Samsung Display ── 50%出資 ──→ SFC ←── 50%出資 ── 保土谷化学\n"
            "                                ↑\n"
            "                       「自分が株主の会社」\n"
            "                       から材料を買っている\n"
            "\n"
            "競合材料メーカーへの切り替え = 「自分が株主の会社を外す」意思決定\n"
            "→ SDCがSFCを切り替えるインセンティブは構造的にほぼゼロ\n"
        ),
    ))

    s.append(Slide(
        kind="content", title="SFCのMoat②: 昇華精製の暗黙知",
        subtitle="1ppm純度 × 数年-数十年の試行錯誤 = 模倣困難",
        bullets=[
            "OLEDホスト材料は不純物が1ppm(100万分の1)以下の超高純度が必要",
            "「昇華精製」プロセスは温度・圧力・速度の微妙な調整が必要",
            "最適パラメータ発見に数年〜数十年の試行錯誤、技術者の勘に依存",
            "保土谷145年の有機化学の知見がSFCに移転、後発が同品質を出すのは極めて困難",
        ],
    ))

    s.append(Slide(
        kind="ascii", title="SFCのMoat③: スペックイン2-3年保護",
        subtitle="切り替えに最短2-3年 = 一度組み込まれたら長期保護",
        ascii_art=(
            "新材料の評価開始\n"
            "  → 単体特性評価(発光効率・色・電流効率)    : 数ヶ月\n"
            "  → デバイス評価(パネルレベルの寿命・画質)  : 数ヶ月〜1年\n"
            "  → 量産試験(歩留・安定性の確認)            : 数ヶ月\n"
            "  → 量産ラインへの組み込み\n"
            "                              合計: 最短でも 2〜3年\n"
            "\n"
            "→ スペックインされた材料は 2-3年間、構造的に保護される\n"
        ),
    ))

    s.append(Slide(
        kind="ascii", title="今が変曲点(タイムライン)",
        subtitle="MacBook Proだけの一発イベントではなく、構造的転換点の始まり",
        ascii_art=(
            "2024年       → iPad Pro Tandem OLED採用(タンデム0% → 30%超に急上昇)\n"
            "\n"
            "2026年Q2     → Samsung Display A6ライン稼働、MacBook Pro向け材料出荷  ← 今ここ\n"
            "\n"
            "2026年Q3     → MacBook Pro発売(Apple秋イベント前後)\n"
            "\n"
            "2026年Q4     → SFC材料需要がピーク、業績反映開始\n"
            "\n"
            "2027年以降   → iPad Air OLED、MacBook Air OLED等 追加案件が続く見込み\n"
        ),
    ))

    s.append(Slide(
        kind="table", title="なぜSFCはマクロをoutperformするか",
        subtitle="顧客集中 × 非スマホ依存 × タンデム恩恵 で +45-65% CAGR",
        table_headers=["理由", "詳細"],
        table=[
            ["顧客がSDC集中", "SDCはIT-OLED独占サプライヤー、増産がそのままSFCに直結"],
            ["スマホOLED縮小の影響なし", "SFC主用途はIT-OLED、スマホ縮小と無関係"],
            ["タンデム採用率上昇", "採用率上昇で1枚あたり材料消費が増える"],
        ],
        footer="SFC材料消費量推計CAGR 45-65% (保守〜楽観)",
    ))

    # 4️⃣
    s.append(Slide(kind="toc", title="目次(バリュエーションへ)",
                   bullets=["▶ 4️⃣ バリュエーション"]))
    s.append(Slide(kind="chapter", title="4️⃣ バリュエーション"))

    s.append(Slide(
        kind="table", title="現在の株価水準",
        subtitle="2-3月比+35-45%上昇後でも、機関投資家の本格再評価はまだ",
        table_headers=["指標", "数値", "読み方"],
        table=[
            ["株価", "2,464円(2026/5/11)", "2-3月比+35-45%上昇済み"],
            ["時価総額", "約381億円", "中小型株の域"],
            ["PBR", "0.69倍", "純資産を下回る「割安水準」"],
            ["アナリストカバレッジ", "0社", "機関投資家の再評価余地が大きい"],
        ],
    ))

    s.append(Slide(
        kind="content", title="なぜまだ割安なのか",
        subtitle="染料メーカー認識 + 簿価のSFC持分 = 二重の隠れ資産",
        bullets=[
            "市場は依然「染料・化学メーカー」として見ている",
            "MacBook Pro向けSFC青色ホスト材料量産は利益予測に未反映",
            "SFC自身がKOSPI上場を計画、実現で保有持分が時価評価される",
            "現在は簿価(取得原価)でしか見えない「隠れた資産」が表面化する可能性",
        ],
        footer="OLED-info (2025年11月) 報道",
    ))

    s.append(Slide(
        kind="table", title="シナリオ別の方向感",
        subtitle="詳細数値はバリュエーションシート参照",
        table_headers=["シナリオ", "前提"],
        table=[
            ["Bear", "MacBook Pro量産がFY27以降にズレ込み、SFC寄与限定的"],
            ["Base", "MacBook Pro 200万枚出荷、タンデム効果フル享受"],
            ["Bull", "MacBook Pro + iPad Pro継続 + 追加IT-OLED案件"],
        ],
    ))

    # 5️⃣
    s.append(Slide(kind="toc", title="目次(Appendixへ)", bullets=["▶ 5️⃣ Appendix"]))
    s.append(Slide(kind="chapter", title="5️⃣ Appendix"))

    s.append(Slide(
        kind="content", title="A. LG Chem訴訟リスク(最重要テールリスク)",
        subtitle="判決2026年7-8月 = MacBook Pro量産と完全に重なる",
        bullets=[
            "LG ChemがSFC青色ホスト材料の特許侵害を訴訟(2019年〜)",
            "2024年に特許法院でもSFC敗訴(2連敗)",
            "2026年4月: LG Chemが製造販売輸入禁止+在庫廃棄を追加請求",
            "2026年7-8月: 最高裁が最終判断予定",
        ],
        footer="The Elec、UBI Research 等",
    ))

    s.append(Slide(
        kind="table", title="LG Chemの請求内容",
        subtitle="差止と在庫廃棄が認められると経営リスク",
        table_headers=["請求", "規模"],
        table=[
            ["損害賠償", "約300億ウォン(約22億円)、SFC年利益と同規模"],
            ["製造販売輸入禁止", "SFCの青色ホスト材料の製造販売輸入の差止"],
            ["在庫廃棄", "現在の在庫を全て廃棄(韓国では極めてまれな請求)"],
        ],
    ))

    s.append(Slide(
        kind="table", title="シナリオ分析",
        subtitle="差止確率15%、メインストーリーが純粋機能する確率40%",
        table_headers=["シナリオ", "推定確率", "保土谷株への影響"],
        table=[
            ["A. 和解 or SFC勝訴", "40%", "✓ 訴訟リスク消滅、メインストーリー機能"],
            ["B. 損害賠償のみで決着", "45%", "△ 約22億円特損、供給継続、影響軽微"],
            ["C. 製造差止命令", "15%", "✗ 量産タイミングで供給停止、株価-30〜-50%"],
        ],
    ))

    s.append(Slide(
        kind="content", title="なぜシナリオCの確率が低いか",
        subtitle="SDCの阻止動機 + 韓国司法慣行で差止は稀",
        bullets=[
            "①SDCはSFC共同出資者、量産ラインに直結する材料が止まることは許容できない",
            "→ SDCがLG Chemとクロスライセンス交渉に入る動機が極めて強い",
            "②韓国の特許侵害訴訟で即時製造差止は極めてまれ",
            "損害賠償が主な救済手段となることが多い",
        ],
    ))

    s.append(Slide(
        kind="content", title="B. 投資戦略: エントリー",
        subtitle="訴訟結果AorBで決着後にポジション拡大が安全",
        bullets=[
            "現在(訴訟結果前): ポジション軽め、15%確率の-30〜-50%リスクあり",
            "訴訟がA/Bで決着(2026年7-8月): ポジション拡大、メインストーリー加速",
            "MacBook Pro発表後反落時: 「事実で売り」の押し目買い機会",
        ],
    ))

    s.append(Slide(
        kind="table", title="B. カタリストと売却タイミング",
        subtitle="LG Chem判決 → MacBook Pro発売 → SFC IPO の順に上昇要因",
        table_headers=["カタリスト", "時期", "対応"],
        table=[
            ["LG Chem最高裁判決", "2026年7-8月", "A/Bならポジション拡大、Cなら即売却"],
            ["MacBook Pro発売", "2026年9-10月", "「噂で買い事実で売り」リスクあり、利確検討"],
            ["FY26 3Q-4Q決算", "26年11月-27年2月", "数字でバリュエーション再評価"],
            ["アナリストカバレッジ開始", "時期不明", "情報格差の窓が閉じる、フェードアウト検討"],
            ["SFC KOSPI上場", "時期未定", "PBR再評価完成、IPO後は割安感消滅"],
        ],
    ))

    s.append(Slide(
        kind="content", title="C. SFC KOSPI IPO計画",
        subtitle="簿価→時価のカタリスト、PBR割安の主因を解消",
        bullets=[
            "OLED-info(2025年11月)報道: SFCがKOSPI上場を計画中",
            "現在、保土谷BSのSFC持分は取得原価(簿価)で計上",
            "SFC上場で保有持分が時価評価され、PBR 0.69倍の主因が表面化",
            "リスク: LG Chem訴訟次第で計画が延期・中止の可能性",
        ],
    ))

    s.append(Slide(
        kind="content", title="D. BOEによる追加機会(参考)",
        subtitle="フレキシブル基板で技術別、現時点で採用エビデンスなし",
        bullets=[
            "BOEは成都に第8.6世代OLEDライン建設中(2025-2026年)",
            "ただしフレキシブルOLED(ポリイミド基板)で、リジッドOLEDと別技術",
            "SFCがBOEに採用されている確認エビデンスは現時点でない",
            "→ メインストーリーには組み込まず、オプション価値として保持",
        ],
    ))

    return SlideDeckConfig(
        title="【4112】保土谷化学工業",
        subtitle="推奨：BUY\nアナリスト0社・PBR 0.69倍 = 機関投資家未発見\n投資期間：2026年Q4までを軸に判断",
        presenter="瀧本ゼミ 発表者",
        slides=s,
    )


def main():
    config = build_deck()
    out = Path(__file__).parent.parent / "outputs" / "4112-hodogaya-slides.pptx"
    out.parent.mkdir(parents=True, exist_ok=True)
    warnings = render(config, str(out))
    print(f"\nWrote: {out}")
    print(f"Slides: {len(config.slides) + 1} (including cover)")
    print(f"Warnings: {len(warnings)}")


if __name__ == "__main__":
    main()
