"""Build the 6648 かわでん slide deck (.pptx). Source: samples/6648-kawaden.md"""
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
        subtitle="武器商人ポジ × 値上げ × ミックスシフト の三段レバレッジ",
        bullets=[
            "建設・電力工事需要が来る、需給ひっ迫で電気工事会社の採算性が大幅改善",
            "かわでんは関電工/きんでんに配電盤を卸す業界一強の武器商人",
            "値上げ × ミックスシフトの二重レバレッジで粗利率改善",
            "市場は「ただの電気部品会社」と認識、ポジションの強さを過小評価",
            "STRONG BUY!!",
        ],
        footer="かわでんIR、関電工・きんでんIR、電設業界各社IR",
    ))

    s.append(Slide(
        kind="ascii", title="成長を因数分解する",
        subtitle="売上 = 新規 + リニューアル / 粗利 = 各セグメント × 粗利率",
        ascii_art=(
            "売上高 = 新規案件売上 + リニューアル売上\n"
            "\n"
            "粗利  = 新規案件売上 × 新規粗利率\n"
            "       + リニューアル売上 × リニューアル粗利率\n"
        ),
    ))

    s.append(Slide(
        kind="table", title="因数分解：各因数の方向",
        subtitle="値上げ × ミックスシフトで営利率5% → 21% へ",
        table_headers=["因数", "方向", "なぜ動くか"],
        table=[
            ["新規案件売上", "↑ +10%/年", "マクロ需要連動、電設サブコン単価+5-10%"],
            ["リニューアル売上", "↑ +15%/年", "マクロ+10% + 営業注力+5%"],
            ["新規粗利率", "↑ 30→42%", "顧客の電設サブコン業績改善で値上げ通る"],
            ["リニューアル粗利率", "↑ 40→52%", "新規より構造的に10pt高い"],
            ["販管費率", "22→24% 横ばい", "業績の伸びに対し相対的に固定"],
        ],
    ))

    s.append(Slide(
        kind="content", title="なぜ市場が気づいていないか",
        subtitle="「ただの電気部品会社」認識が独占性と転換期を見落とし",
        bullets=[
            "市場の認識: ただの電気部品会社、典型的な値嵩株",
            "カスタム品事業: 日本唯一の上場会社という独占性が知られていない",
            "リニューアル: バブル期建設物の交換期到来という構造需要が織り込まれない",
            "顧客との力関係: 関電工・きんでんが大株主 → 優良案件集中が見えていない",
        ],
    ))

    s.append(Slide(kind="toc", title="目次",
                   bullets=["1️⃣ 基本情報", "2️⃣ マクロ", "3️⃣ ミクロ",
                            "4️⃣ バリュエーション", "5️⃣ Appendix"]))

    # 1️⃣
    s.append(Slide(kind="chapter", title="1️⃣ 基本情報"))

    s.append(Slide(
        kind="table", title="会社の概要",
        subtitle="国内最大手シェア2割、新規7割+リニューアル3割",
        table_headers=["項目", "内容"],
        table=[
            ["銘柄コード", "6648 (東証プライム)"],
            ["主事業", "配電盤・分電盤・DC向け電気設備の製造販売"],
            ["顧客", "関電工・きんでんなどの電設サブコン"],
            ["シェア", "国内最大手で約2割(競合は非上場で売上100億未満が大半)"],
            ["売上構成", "新規 約7割 / リニューアル 約3割"],
            ["創業", "1925年(今年で創業100周年)"],
        ],
    ))

    s.append(Slide(
        kind="content", title="配電盤・分電盤とは",
        subtitle="建物の電気の入り口、建設に必須の機器",
        bullets=[
            "配電盤: 電力会社からの高圧電気をビル内で使える低圧に変成",
            "分電盤: ブレーカーのこと",
            "→ 建物の電気の入り口、建設には必須",
        ],
    ))

    s.append(Slide(
        kind="table", title="カスタム品と汎用品の違い",
        subtitle="カスタム品の上場会社は日本ではかわでんのみ",
        table_headers=["用途", "製品"],
        table=[
            ["一般戸建/中小マンション/ビル", "汎用品で事足りる"],
            ["高層ビル/DC/半導体工場/その他工場", "カスタム品が必須(複雑構造)"],
        ],
        footer="他の上場配電盤会社は汎用品専門",
    ))

    s.append(Slide(
        kind="content", title="製品の使い先",
        subtitle="AIブームの恩恵を最も直接的に受ける用途構成",
        bullets=[
            "戸建ては全くなく、ほぼすべて非住宅(タワマンも数%)",
            "首都圏再開発の高層ビル、データセンター、半導体工場",
            "国内回帰した製造業の工場、自衛隊基地などが多め",
            "→ AIブームの恩恵を最も直接的に受ける用途構成",
        ],
        footer="かわでんIR",
    ))

    s.append(Slide(
        kind="content", title="リニューアル事業について",
        subtitle="バブル期の自顧客リスト = 超強力な営業資産",
        bullets=[
            "既導入建物のメンテナンス・点検・部品交換",
            "配電盤の耐久15-20年、バブル期建設の建物が一斉に交換期",
            "創業100周年、中計で構造改革を宣言",
            "他社はこのリニューアル事業を代替できない(後述)",
        ],
    ))

    # 2️⃣
    s.append(Slide(kind="toc", title="目次(マクロへ)", bullets=["▶ 2️⃣ マクロ"]))
    s.append(Slide(kind="chapter", title="2️⃣ マクロ"))

    s.append(Slide(
        kind="content", title="様々な需要が流入中",
        subtitle="これだけの追い風が一業界に同時に来るのは滅多にない",
        bullets=[
            "首都圏ビル再開発、データセンター、半導体工場",
            "自衛隊防衛関連施設、火力・太陽光発電所",
            "系統用蓄電池の設置",
            "電線などのインフラ更新需要",
        ],
    ))

    s.append(Slide(
        kind="content", title="業界全体の受注残が急増",
        subtitle="今後数年間、需給ひっ迫は変わらない見通し",
        bullets=[
            "電設業界各社IR: 受注は非常に好調",
            "受注残が積みあがる状態は今後数年変わりそうにない",
            "→ かわでんへの長期的需要が確保される",
        ],
    ))

    s.append(Slide(
        kind="table", title="しかし、肝心の電気工事士が足りない",
        subtitle="求人倍率は全国平均の約3倍",
        table_headers=["指標", "数値"],
        table=[
            ["国内全体の有効求人倍率", "1.3倍"],
            ["電気工事士の求人倍率", "3.6倍 (全国平均の約3倍)"],
        ],
        footer="IR: 人手不足は今後数年変わらない",
    ))

    s.append(Slide(
        kind="content", title="結果、工事単価と粗利率が爆上がり",
        subtitle="関電工・きんでんの粗利&株価が急上昇",
        bullets=[
            "業界最大手の関電工、きんでんの粗利と株価は急上昇",
            "顧客(電設サブコン)が儲かるからかわでんの値上げも通る",
            "→ コバンザメ戦略が成立する",
        ],
    ))

    # 3️⃣
    s.append(Slide(kind="toc", title="目次(ミクロへ)", bullets=["▶ 3️⃣ ミクロ"]))
    s.append(Slide(kind="chapter", title="3️⃣ ミクロ"))

    s.append(Slide(
        kind="content", title="①コバンザメ戦略: 顧客の電設サブコンに値上げ",
        subtitle="業界全体が儲かる時に業界一強の取り分が増える",
        bullets=[
            "顧客サブコンの業績改善で、かわでんの値上げも通りやすい",
            "業界一強なので関電工・きんでんから優良案件を紹介してもらいやすい",
            "(大阪IR、首都圏再開発、ラピダス工場など)",
            "競合は規模が小さく、かわでんに頼むような工事を頼めない",
            "関電工ときんでんが大株主 → 優良案件を独占する構造",
        ],
    ))

    s.append(Slide(
        kind="content", title="②既存案件の保守点検でミックス改善",
        subtitle="バブル期に作った自社製品が一斉に耐用年数を迎える",
        bullets=[
            "既導入顧客のメンテ・部品交換は作業効率良く粗利率が新規より10pt高い",
            "今期から営業活動を強化、既存顧客の案件獲得を加速",
            "バブル期建設物の自社製品が一斉に耐用年数 = 顧客リストが膨大",
            "→ ハードウェアばらまき+アフターサービスで稼ぐ配電盤業界のApple化",
        ],
    ))

    s.append(Slide(
        kind="content", title="今期上期でリニューアルの売上高大幅増",
        subtitle="特段大きな案件ではない = 構造的な改善",
        bullets=[
            "営業活動注力の結果、新規案件よりリニューアルの伸び率が高い",
            "特段大きな案件が入ったわけではない(IR)",
            "→ 構造的な改善であり、一時的な大型案件ではない",
        ],
        footer="IR",
    ))

    # 4️⃣
    s.append(Slide(kind="toc", title="目次(バリュエーションへ)",
                   bullets=["▶ 4️⃣ バリュエーション"]))
    s.append(Slide(kind="chapter", title="4️⃣ バリュエーション"))

    s.append(Slide(
        kind="table", title="売上の組み立て",
        subtitle="新規+10%/年、リニューアル+15%/年",
        table_headers=["年度", "新規案件", "リニューアル", "合計"],
        table=[
            ["FY24 実", "15,090", "6,244", "21,334"],
            ["FY25 実", "17,827", "6,391", "24,218"],
            ["FY26 予", "19,610", "7,669", "27,279"],
            ["FY27 予", "21,571", "8,820", "30,390"],
            ["FY28 予", "23,728", "10,143", "33,870"],
        ],
        footer="単位: 百万円",
    ))

    s.append(Slide(
        kind="table", title="粗利率の改善",
        subtitle="新規+リニューアルとも+6/+3/+3pt の改善 = ミックス効果",
        table_headers=["年度", "新規粗利率", "リニューアル粗利率", "全社粗利率"],
        table=[
            ["FY25 実", "30%", "40%", "33.1%"],
            ["FY26 予", "36%", "46%", "38.8%"],
            ["FY27 予", "39%", "49%", "41.9%"],
            ["FY28 予", "42%", "52%", "45.0%"],
        ],
    ))

    s.append(Slide(
        kind="table", title="利益の爆発",
        subtitle="営業利益 1,135 → 7,111 百万円(×6.3倍)",
        table_headers=["指標", "FY24 実", "FY25 実", "FY26 予", "FY27 予", "FY28 予"],
        table=[
            ["売上高", "21,334", "24,218", "27,279", "30,390", "33,870"],
            ["営業利益", "1,135", "2,590", "4,040", "5,440", "7,111"],
            ["営業利益率", "5.3%", "10.7%", "14.8%", "17.9%", "21.0%"],
        ],
        footer="単位: 百万円",
    ))

    s.append(Slide(
        kind="table", title="PER感応度",
        subtitle="同業他社上場なし、電設サブコンPER 15-20 を参考",
        table_headers=["シナリオ", "PER", "根拠"],
        table=[
            ["Worst", "12.5", "保守側"],
            ["Normal", "15", "電設サブコン下限"],
            ["Best", "18", "電設サブコン中央値"],
        ],
        footer="現状かわでんPER 14 → 上振れ余地あり",
    ))

    # 5️⃣
    s.append(Slide(kind="toc", title="目次(Appendixへ)", bullets=["▶ 5️⃣ Appendix"]))
    s.append(Slide(kind="chapter", title="5️⃣ Appendix"))

    s.append(Slide(
        kind="content", title="A. 投資期間: いつ売るか",
        subtitle="新工場稼働前に売却が資金効率的に最善",
        bullets=[
            "29年期4Qから新工場の減価償却・費用が掛かり始める(年数十億)",
            "値上げによる利益率改善が断言できるのは2-3年後まで",
            "→ 2027年5月(28年期ガイダンス時) もしくは 2028年2月(28年期3Q決算時)",
        ],
        footer="かわでんIR、関電工IR",
    ))

    s.append(Slide(
        kind="content", title="B. そんなに採用できるのか? → できる",
        subtitle="大手は採用絞り、落ちた人が来る構造",
        bullets=[
            "ターゲット人材は毎年3,000-8,000人いる(関電工IR)",
            "ここ数年新規+中途で毎年40-50人採用できている",
            "競合他社以上の給料を出している",
        ],
    ))

    s.append(Slide(
        kind="content", title="C. 大手に人手を吸われない? → 可能性低い",
        subtitle="関電工・きんでんはバブル期トラウマで採用絞り",
        bullets=[
            "関電工・きんでんは大規模採用を絞る(バブル期に大量採用→リストラのトラウマ)",
            "採用倍率: 関電工 約10倍、きんでん 約20倍",
            "→ そこで落ちた人がかわでんに来る = 数千人単位の潜在候補",
        ],
    ))

    s.append(Slide(
        kind="content", title="D. リニューアルは他社に代替されない",
        subtitle="顧客リスト・設計図・独自部品がかわでんだけにある",
        bullets=[
            "カスタム品: 顧客リストや設計図はかわでんしか持っていない",
            "独自部品も多く使われ、他社は調達できない",
            "建物全体の電気システムがかわでん配電盤に合うように作られている",
            "→ 既存顧客のメンテ・修理需要はほぼすべてかわでんでさばく",
        ],
        footer="IR",
    ))

    s.append(Slide(
        kind="content", title="E. 競合・外国企業の台頭? → 可能性低い",
        subtitle="カスタム品の参入障壁が高い",
        bullets=[
            "汎用品メーカーがカスタム品参入: 設計人材・設備を急には整えられない",
            "外国企業: 販売・設置体制を急には確保できない",
            "顧客も信用のない会社に配電盤を任せにくい",
            "→ かわでんの一強が続く",
        ],
    ))

    s.append(Slide(
        kind="content", title="F. 工場の生産キャパは大丈夫?",
        subtitle="既存2工場の稼働率80-90%、問題なし",
        bullets=[
            "既存の国内2工場のキャパは割と余裕がある",
            "現時点の工場稼働率は80-90%程度",
        ],
        footer="IR",
    ))

    return SlideDeckConfig(
        title="【6648】かわでん",
        subtitle="推奨：STRONG BUY\n株価：2,300円 → 4,978円\n投資期間：1-2年弱",
        presenter="14.0期 庭野純之介",
        slides=s,
    )


def main():
    config = build_deck()
    out = Path(__file__).parent.parent / "outputs" / "6648-kawaden-slides.pptx"
    out.parent.mkdir(parents=True, exist_ok=True)
    warnings = render(config, str(out))
    print(f"\nWrote: {out}")
    print(f"Slides: {len(config.slides) + 1} (including cover)")
    print(f"Warnings: {len(warnings)}")


if __name__ == "__main__":
    main()
