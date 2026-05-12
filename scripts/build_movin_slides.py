"""Build the 421A ムービン slide deck (.pptx). Source: samples/421A-movin.md"""
from __future__ import annotations

import sys
from pathlib import Path

_REPO = Path(__file__).parent.parent
sys.path.insert(0, str(_REPO / ".claude/skills/takimoto-presentation/scripts"))
from pptx_renderer import Slide, SlideDeckConfig, render  # noqa: E402


def build_deck() -> SlideDeckConfig:
    s: list[Slide] = []

    # 0️⃣ ストーリー / 因数分解 / ミスプライシング
    s.append(Slide(
        kind="content", title="ストーリー",
        subtitle="25年保守経営から覚醒、CAボトルネック解消で売上×4.5倍へ",
        bullets=[
            "コンサル特化の転職エージェント、創業28年で95%のコンサルとつながりあり",
            "25年間ヨコヨコ → AIや競合の出現で社長が危機感、本気の成長フェーズへ",
            "CA増員・クライアントシフト・SEO強化・DB掘り起こしの4施策が一斉発動",
            "FY24→FY27で売上 ×4.5、純利益 ×8.5",
            "市場は「市場連動の会社」と認識、施策起因の加速フェーズに気づいていない",
        ],
        footer="ムービンIR電話、目論見書、IMF、コンサル業界各社の被保険者数",
    ))

    s.append(Slide(
        kind="ascii", title="成長を因数分解する",
        subtitle="経験年数コホート別 + 検算用に成約数×単価ベース",
        ascii_art=(
            "売上 =  1年目人数 × 一人当たり売上\n"
            "      + 2年目人数 × 一人当たり売上\n"
            "      + 3年目以降人数 × 一人当たり売上\n"
            "\n"
            "別の見方(検算用):\n"
            "売上 = 成約数 × 成約単価\n"
            "成約単価 = 転職後年収 × 手数料率\n"
        ),
    ))

    s.append(Slide(
        kind="table", title="因数分解：各因数の方向",
        subtitle="採用と単価上昇の二重追い風",
        table_headers=["因数", "方向", "なぜ動くか"],
        table=[
            ["1年目CA人数", "↑ +35/+40/+40", "採用計画。手取金をCA人件費に充当"],
            ["2年目以降人数", "↑ 自動増", "定着率1年目92%/2年目以降82%でコホート厚く"],
            ["一人当たり売上", "↑ 経験曲線", "1年目→2年目で4.7倍、3年目で7.99倍"],
            ["成約単価", "↑ 1.00→1.49", "年収+5%×手数料率35→40%上昇"],
        ],
        footer="目論見書、IR電話",
    ))

    s.append(Slide(
        kind="content", title="なぜ市場が気づいていないか",
        subtitle="「市場連動の会社」認識が施策起因の加速を捉えていない",
        bullets=[
            "過去業績: 2025年に入るまで25年間ヨコヨコ → 成長しない会社と認識",
            "直近の伸びの解釈: 「コンサル市場の伸びに比例」と見られる",
            "実態: 施策起因の加速フェーズ。CAという律速が外れ、施策効果が積み上がる",
            "来期以降: 4施策はまだ途上段階、直近の変化は加速する",
        ],
    ))

    s.append(Slide(kind="toc", title="目次",
                   bullets=["1️⃣ 基本情報", "2️⃣ マクロ", "3️⃣ ミクロ",
                            "4️⃣ バリュエーション", "5️⃣ Appendix"]))

    # 1️⃣ 基本情報
    s.append(Slide(kind="chapter", title="1️⃣ 基本情報"))

    s.append(Slide(
        kind="table", title="会社の概要",
        subtitle="コンサル特化の転職エージェント、創業28年",
        table_headers=["項目", "内容"],
        table=[
            ["銘柄コード", "421A (東証グロース)"],
            ["主事業", "コンサル特化の転職エージェント"],
            ["モデル", "求職者にCAが無料支援、採用時に企業から年収×手数料率"],
            ["現在の株価/PER", "3,300円 / 27.5"],
            ["投資の核心", "25年保守経営から成長フェーズへの大転換"],
        ],
    ))

    s.append(Slide(
        kind="content", title="求職者サイドの構造",
        subtitle="高学歴&若者中心、8割が自社プラットフォーム経由",
        bullets=[
            "「コンサル 転職」検索でHPが上位表示(創業からのSEO)",
            "自社HP経由が80%、ビズリーチ等の外部PFが20%",
            "広告費ほぼ0で前期5,800人以上を自社集客",
            "→ 集客の仕組みが既に成立している",
        ],
    ))

    s.append(Slide(
        kind="content", title="クライアントの「優勝劣敗」",
        subtitle="ITコンサル/DX支援/PMO が伸び、純粋戦略/シンクタンクは伸び悩み",
        bullets=[
            "「優勝」側: ノースサンド、Dirbatoなど(ITコンサル/DX実行支援/PMO)",
            "「劣敗」側: マッキンゼー、三菱総研などの純粋戦略・シンクタンク",
            "ムービンの戦略: 積極採用している優勝側に集中送客",
            "→ 単価上昇と成約数増加に直結",
        ],
        footer="主要コンサルティングファームの被保険者数の変動率データ",
    ))

    # 2️⃣ マクロ
    s.append(Slide(kind="toc", title="目次(マクロへ)", bullets=["▶ 2️⃣ マクロ"]))
    s.append(Slide(kind="chapter", title="2️⃣ マクロ"))

    s.append(Slide(
        kind="content", title="コンサル業界は年率約10%成長",
        subtitle="IT/DX/PMO が牽引、潮流は中長期で変わらず",
        bullets=[
            "市場CAGR: 約10%",
            "牽引分野: ITコンサル、DX実行支援、PMO",
            "要因: 生成AI事業開発、DX・IT刷新ニーズの高まり",
            "見通し: 潮流は変わらず中長期的に高成長",
        ],
    ))

    s.append(Slide(
        kind="table", title="コンサル業界の人手不足",
        subtitle="他業界より3倍以上人手不足、構造的に解消しない",
        table_headers=["指標", "コンサル", "全国平均"],
        table=[
            ["転職求人倍率", "8.68倍", "2.51倍"],
        ],
    ))

    s.append(Slide(
        kind="content", title="インフレ+人手不足で年収+5%/年",
        subtitle="インフレ2% + 構造的人手不足 = 中長期で年収上昇",
        bullets=[
            "2022年から年率2%でインフレ発生(IMFは少なくとも2030年まで継続予想)",
            "インフレ + 人手不足で年率5%で年収上昇",
            "→ ムービンの成約単価が自動的に上がる",
        ],
        footer="IMF、コンサル業界各社IR",
    ))

    s.append(Slide(
        kind="ascii", title="マクロまとめ",
        subtitle="市場の苦しみがムービンの追い風になる構造",
        ascii_art=(
            "コンサル業界が人手不足とインフレで苦しめば苦しむほど\n"
            "    ↓\n"
            "求人数と年収は上昇する\n"
            "    ↓\n"
            "ムービンの成約数と成約単価は上昇する\n"
            "\n"
            "ただし市場は: 「市場に比例して成長する会社」と認識\n"
            "→ 内部要因による加速を見落としている (= ミスプライシング)\n"
        ),
    ))

    # 3️⃣ ミクロ
    s.append(Slide(kind="toc", title="目次(ミクロへ)", bullets=["▶ 3️⃣ ミクロ"]))
    s.append(Slide(kind="chapter", title="3️⃣ ミクロ"))

    s.append(Slide(
        kind="content", title="歴史: 創業と25年の保守経営",
        subtitle="集客はずっと回っていた、CAが足りなくて伸びなかった",
        bullets=[
            "余命宣告されて、家族を養うために28年前に創業",
            "「成長を指向しない、少数で高い果実」という方針",
            "SEOは創業からやっていて集客は十分、CAが足りずさばけない",
            "→ 25年間売上ヨコヨコ",
        ],
    ))

    s.append(Slide(
        kind="content", title="変曲点: 外部環境変化と危機感",
        subtitle="AI出現+競合台頭+事業承継 で本気の成長フェーズへ",
        bullets=[
            "AIの出現で競争環境がガラッと変わりそう",
            "ビズリーチ・リクルートと本格的に競合することに",
            "創業25年でメンバー高齢化、事業承継が必要",
            "→「これまでの方針では潰れる」と社長が危機感、成長に舵を切る",
        ],
    ))

    s.append(Slide(
        kind="content", title="IPOの理由",
        subtitle="事業承継 + BtoC スケール拡大に必要不可欠",
        bullets=[
            "事業承継には内部統制と資本問題解決が必要 → ほぼIPOと変わらない",
            "BtoCなので知名度向上・資本力で投資ができる環境が必要",
            "→ IPOがこの解決策として合理的",
        ],
    ))

    s.append(Slide(
        kind="table", title="施策① SEO対策の強化",
        subtitle="2022年メンバー増員、自社HP集客が3年で+38%",
        table_headers=["年", "自社HP経由の集客"],
        table=[
            ["2022年", "4,200人"],
            ["2023年", "4,600人"],
            ["2024年", "5,800人"],
        ],
    ))

    s.append(Slide(
        kind="content", title="施策② キャリアアドバイザー増員",
        subtitle="律速段階を解消、成果は来期以降に反映",
        bullets=[
            "これまではDBに求職者が貯まっても処理できていなかった",
            "2024年以降CA採用数を増加",
            "彼らが戦力になるのは来期以降",
            "→ いま採用したCAの成果はまだ業績未反映 = 加速フェーズの根拠",
        ],
    ))

    s.append(Slide(
        kind="content", title="施策③ クライアントシフト",
        subtitle="積極採用ファームに集中送客、手数料率35→40%超へ",
        bullets=[
            "2024年1Qからクライアントポートフォリオを入れ替え",
            "伸びているクライアントの募集をキャッチしてフォローアップ",
            "4Qから徐々に成果、まだ途上",
            "→ 手数料率上昇 × 成約数増加",
        ],
    ))

    s.append(Slide(
        kind="content", title="施策④ 過去DB掘り起こし",
        subtitle="過去にアプローチしてない8割に余地あり",
        bullets=[
            "それまでは一度支援したら終わりだった",
            "2023年から掘り起こしを開始、成果が出たため増員",
            "現状は過去DBの2割しかアプローチできていない",
            "→ まだ余地あり",
        ],
    ))

    s.append(Slide(
        kind="content", title="参入障壁",
        subtitle="集客力PFと過去DBは短期では獲得不可能",
        bullets=[
            "集客PFを持たない他社は他社PF利用料を払う必要 → 経費圧迫",
            "過去DBがないと新規集客に依存 → 広告宣伝費がかさむ",
            "経費が膨らむと年収に回せず、CAの採用力が落ちる",
            "→ 創業28年の蓄積が参入障壁",
        ],
    ))

    # 4️⃣ バリュエーション
    s.append(Slide(kind="toc", title="目次(バリュエーションへ)",
                   bullets=["▶ 4️⃣ バリュエーション"]))
    s.append(Slide(kind="chapter", title="4️⃣ バリュエーション"))

    s.append(Slide(
        kind="table", title="売上の組み立て(コホート別)",
        subtitle="3年目以降コホートが時間と共に勝手に厚くなる",
        table_headers=["年度", "1年目", "2年目", "3年目以降", "合計"],
        table=[
            ["FY24", "158", "508", "1,728", "2,387"],
            ["FY25", "243", "1,636", "2,545", "4,424"],
            ["FY26", "241", "2,347", "4,865", "7,453"],
            ["FY27", "268", "2,320", "8,060", "10,647"],
        ],
        footer="単位: 百万円。本発表バリュエーションシート",
    ))

    s.append(Slide(
        kind="table", title="利益のレバレッジ",
        subtitle="固定費752百万円据置 → 営業利益率36% → 69% へ跳躍",
        table_headers=["指標", "FY24", "FY25", "FY26", "FY27"],
        table=[
            ["売上高", "2,387", "4,424", "7,453", "10,647"],
            ["営業利益", "861", "2,384", "4,729", "7,349"],
            ["営業利益率", "36.07%", "53.89%", "63.45%", "69.02%"],
            ["純利益", "575", "1,592", "3,158", "4,908"],
        ],
        footer="単位: 百万円",
    ))

    s.append(Slide(
        kind="table", title="PER感応度",
        subtitle="ワーストでも株価4.43倍 — コホートの自然な厚みが効く",
        table_headers=["シナリオ", "PER", "株価", "倍率"],
        table=[
            ["ベスト", "30", "17,546円", "5.32倍"],
            ["ノーマル", "27.5", "16,084円", "4.87倍"],
            ["ワースト", "25", "14,621円", "4.43倍"],
        ],
        footer="現株価3,300円、25/12予想利益1,007百万円、発行済8.05百万株",
    ))

    # 5️⃣ Appendix
    s.append(Slide(kind="toc", title="目次(Appendixへ)", bullets=["▶ 5️⃣ Appendix"]))
    s.append(Slide(kind="chapter", title="5️⃣ Appendix"))

    s.append(Slide(
        kind="content", title="A. CA採用計画の根拠",
        subtitle="今期+35、来期+40、再来期+40 — 達成可能",
        bullets=[
            "手取金(IPO調達)はCA人件費に充当する予定",
            "IR電話: 上記の採用計画を立てている",
            "良い在庫さえあればCAの仕事は簡単、採用力も高い",
            "→ この採用計画は難なく達成しそう",
        ],
        footer="IR電話",
    ))

    s.append(Slide(
        kind="content", title="B. 一人当たり売上高の根拠",
        subtitle="入社1年目900万→2年目4.7倍→3年目7.99倍",
        bullets=[
            "目論見書: 入社1年目対比、2年目で4.7倍、3年目で7.99倍",
            "計算上、入社1年目の平均売上高は900万円",
            "増加は一人当たり成約数の伸びによるもので、今後も一定と仮定",
            "→ 一人当たり売上高は成約単価の増加に比例する",
        ],
        footer="目論見書",
    ))

    s.append(Slide(
        kind="content", title="C. 成約単価の根拠",
        subtitle="転職後年収 × 手数料率 で年率+10%級の伸び",
        bullets=[
            "成約単価 = 転職後の年収 × 手数料率",
            "転職後年収: インフレ+人手不足で年率5%成長と仮定",
            "手数料率: 35%(従来) → クライアントシフト後40%超(IR)",
            "今後さらに上昇見込み(積極採用企業への送客比率増加)",
        ],
        footer="IR電話、IMF",
    ))

    s.append(Slide(
        kind="content", title="D. 成約数の構造",
        subtitle="CA律速で現状5.7%しか送客できず、伸びしろ膨大",
        bullets=[
            "自社PF経由81%(新規80% + 過去DB20%)、外部PF19%",
            "新規履歴書数の5.7%しか送客できていない",
            "→ 新規履歴書数の6-7分の1までいけそう = 伸びしろ膨大",
            "過去DBはそろそろ限度、その分を外部集客に頼るパターンでバリュエーション",
        ],
    ))

    return SlideDeckConfig(
        title="【421A】ムービン・ストラテジック・キャリア",
        subtitle="推奨：BUY\n投資期間：3年(FY27まで)\nワーストでも株価4.43倍",
        presenter="13.0期 生田開都",
        slides=s,
    )


def main():
    config = build_deck()
    out = Path(__file__).parent.parent / "outputs" / "421A-movin-slides.pptx"
    out.parent.mkdir(parents=True, exist_ok=True)
    warnings = render(config, str(out))
    print(f"\nWrote: {out}")
    print(f"Slides: {len(config.slides) + 1} (including cover)")
    print(f"Warnings: {len(warnings)}")


if __name__ == "__main__":
    main()
