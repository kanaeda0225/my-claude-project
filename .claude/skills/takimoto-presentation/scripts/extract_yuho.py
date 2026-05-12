"""Extract key financials from a Japanese yuho (有価証券報告書) PDF.

Usage:
    from extract_yuho import extract_yuho
    data = extract_yuho("path/to/yuho.pdf")
    # data["pl"]["FY24"]["売上高"] -> 18234 (百万円)
    # data["pl"]["FY24"]["販管費内訳"]["給料及び賞与"] -> 1858

Notes:
- Calls `pdftotext -layout` under the hood (poppler-utils must be installed)
- Adobe-Japan1 font warnings are normal and don't affect extraction
- Returns numbers in 百万円 (rounded from 千円 in yuho)
- Supports the standard yuho layout (連結損益計算書, 主要な経営指標等, セグメント情報)
"""
from __future__ import annotations

import re
import subprocess
import tempfile
from pathlib import Path


def _run_pdftotext(pdf_path: str) -> str:
    """Run pdftotext -layout and return extracted text."""
    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        subprocess.run(
            ["pdftotext", "-layout", str(pdf_path), tmp_path],
            check=True,
            capture_output=True,
        )
        return Path(tmp_path).read_text()
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def _parse_thousand(s: str) -> int:
    """Parse '18,234,377' or '△18,074' style number to int (千円 → 百万円 if /1000)."""
    s = s.strip().replace(",", "")
    if s.startswith("△") or s.startswith("-"):
        return -int(s.lstrip("△").lstrip("-"))
    return int(s)


def _to_million(thousand: int) -> int:
    """Convert 千円 to 百万円 (rounded)."""
    return round(thousand / 1000)


# ---------------------------------------------------------------------------
# Extractors
# ---------------------------------------------------------------------------

def _find_5year_indicators(text: str) -> dict:
    """Find the 5-year main indicators block.

    Returns dict like:
      {"連結": {"FY21": {"売上高": 15038, "経常利益": 4104, ...}, "FY22": ...}}
    """
    out: dict = {"連結": {}, "単体": {}}
    # Look for the block headed by '主要な経営指標等の推移'
    m = re.search(r"主要な経営指標等の推移.*?\(1\) 連結経営指標等(.*?)(?:\(2\)|提出会社の経営指標等)",
                  text, re.DOTALL)
    if not m:
        return out

    block = m.group(1)
    # Extract years (e.g., 2021年12月 / 2022年12月 / ...)
    year_match = re.findall(r"(\d{4})年12月", block)
    if not year_match:
        return out
    years = [f"FY{int(y) % 100}" for y in year_match[:5]]

    # For each metric line, pick 5 numbers
    metrics_patterns = {
        "売上高": r"売上高\s*\(千円\)\s*([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)",
        "経常利益": r"経常利益\s*\(千円\)\s*([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)",
        "純利益": r"親会社株主に帰属する\s*\(千円\)\s*([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)",
        "包括利益": r"包括利益\s*\(千円\)\s*([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)",
        "純資産額": r"純資産額\s*\(千円\)\s*([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)",
        "総資産額": r"総資産額\s*\(千円\)\s*([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)",
        "EPS": r"１株当たり当期純利益\s*\(円\)\s*([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)",
        "BPS": r"１株当たり純資産額\s*\(円\)\s*([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)",
    }

    for metric, pat in metrics_patterns.items():
        m2 = re.search(pat, block)
        if not m2:
            continue
        for i, year in enumerate(years):
            raw = m2.group(i + 1)
            try:
                if metric in ("EPS", "BPS"):
                    val = float(raw.replace(",", ""))
                else:
                    val = _to_million(_parse_thousand(raw))
                out["連結"].setdefault(year, {})[metric] = val
            except (ValueError, IndexError):
                pass

    return out


def _find_consolidated_pl(text: str) -> dict:
    """Find the most recent 連結損益計算書 with 2-year comparison.

    Returns dict like:
      {"FY24": {"売上高": 18234, "売上原価": 7132, ...}, "FY25": {...}}
    """
    out: dict = {}
    # Find the section
    m = re.search(r"② 【連結損益計算書及び連結包括利益計算書】\s*【連結損益計算書】(.*?)【連結包括利益計算書】",
                  text, re.DOTALL)
    if not m:
        return out

    block = m.group(1)
    # Year headers like "(自 2024年１月１日" — find prior and current years
    year_matches = re.findall(r"自\s*(\d{4})年", block)
    if len(year_matches) < 2:
        return out
    prev_year = f"FY{int(year_matches[0]) % 100}"
    cur_year = f"FY{int(year_matches[1]) % 100}"

    pl_items = {
        "売上高": r"^売上高\s+([\d,]+)\s+([\d,]+)\s*$",
        "売上原価": r"^売上原価\s+(?:※\d+\s+)?([\d,]+)\s+(?:※\d+\s+)?([\d,]+)\s*$",
        "売上総利益": r"^売上総利益\s+([\d,]+)\s+([\d,]+)\s*$",
        "販管費": r"^販売費及び一般管理費\s+(?:※[\d,]+\s+)?([\d,]+)\s+(?:※[\d,]+\s+)?([\d,]+)\s*$",
        "営業利益": r"^営業利益\s+([\d,]+)\s+([\d,]+)\s*$",
        "経常利益": r"^経常利益\s+([\d,]+)\s+([\d,]+)\s*$",
        "税前純利益": r"税金等調整前当期純利益\s+([\d,]+)\s+([\d,]+)",
        "純利益": r"^親会社株主に帰属する当期純利益\s+([\d,]+)\s+([\d,]+)\s*$",
    }

    for item, pat in pl_items.items():
        for line in block.split("\n"):
            m2 = re.search(pat, line.strip())
            if m2:
                try:
                    prev = _to_million(_parse_thousand(m2.group(1)))
                    cur = _to_million(_parse_thousand(m2.group(2)))
                    out.setdefault(prev_year, {})[item] = prev
                    out.setdefault(cur_year, {})[item] = cur
                except ValueError:
                    pass
                break

    return out


def _find_sga_breakdown(text: str) -> dict:
    """Find the 販管費の主要な費目 breakdown (typically inside 注記事項).

    Returns dict like:
      {"FY24": {"給料及び賞与": 1857, "荷造運搬費": 779, ...}, "FY25": {...}}
    """
    out: dict = {}
    m = re.search(
        r"販売費及び一般管理費の主なもの.*?自\s*(\d{4})年.*?自\s*(\d{4})年(.*?)※３",
        text, re.DOTALL,
    )
    if not m:
        return out
    prev_year = f"FY{int(m.group(1)) % 100}"
    cur_year = f"FY{int(m.group(2)) % 100}"
    block = m.group(3)

    sga_items = [
        "貸倒引当金繰入額",
        "給料及び賞与",
        "荷造運搬費",
        "賞与引当金繰入額",
        "役員賞与引当金繰入額",
        "株式報酬引当金繰入額",
        "退職給付費用",
        "研究開発費",
    ]
    for item in sga_items:
        pat = rf"{item}\s+([\d,]+)(?:千円)?\s+([\d,]+)"
        m2 = re.search(pat, block)
        if m2:
            try:
                prev = _to_million(_parse_thousand(m2.group(1)))
                cur = _to_million(_parse_thousand(m2.group(2)))
                out.setdefault(prev_year, {})[item] = prev
                out.setdefault(cur_year, {})[item] = cur
            except ValueError:
                pass

    return out


def _find_product_breakdown(text: str) -> dict:
    """Try to find 売上高の内訳 by product (typical at start of 経営成績の分析).

    Returns dict like:
      {"FY25": {"薬品売上": 20211, "機械売上": 312, "資材売上": 403, "その他売上": 19}}
    """
    out: dict = {}
    # Looks for the line pattern: 「薬品売上高は202億11百万円（前期比27億33百万円、15.6％増）」
    # Tolerates whitespace inside product names, full-width digits, and line
    # wraps between 百万 and 円. Also handles small revenues without 億 part.
    fw_to_hw = str.maketrans("０１２３４５６７８９", "0123456789")
    normalized = text.translate(fw_to_hw)
    # Pattern with 億 part
    pat_with_oku = (
        r"(薬\s*品|機\s*械|資\s*材|その\s*他)\s*売上高\s*(?:は)?\s*"
        r"(\d+)\s*億\s*(\d+)\s*百万\s*円"
    )
    for m in re.finditer(pat_with_oku, normalized):
        name = m.group(1).replace(" ", "").replace("　", "")
        key = f"{name}売上"
        oku = int(m.group(2))
        man = int(m.group(3))
        value = oku * 100 + man  # in 百万円
        out.setdefault("FY_latest", {})[key] = value
    # Pattern without 億 part (small revenues like その他19百万円)
    pat_no_oku = (
        r"(薬\s*品|機\s*械|資\s*材|その\s*他)\s*売上高\s*(?:は)?\s*"
        r"(\d{1,3})\s*百万\s*円"
    )
    for m in re.finditer(pat_no_oku, normalized):
        name = m.group(1).replace(" ", "").replace("　", "")
        key = f"{name}売上"
        if key in out.get("FY_latest", {}):
            continue  # don't overwrite the more-precise match above
        value = int(m.group(2))
        # Conservatively assume this is the latest (current) fiscal year
        out.setdefault("FY_latest", {})[key] = value
    return out


def _find_shares_outstanding(text: str) -> int | None:
    """Find 発行済株式総数 (latest)."""
    m = re.search(r"発行済株式総数\s*\(株\)\s*([\d,]+)", text)
    if m:
        # The regex captured the first number (FY21). The next 4 numbers on the
        # same line are FY22-FY25 (latest). We extract them and take the 4th.
        rest = m.string[m.end():]
        # Stop at the first newline that has no numbers (table row ends)
        first_line = rest.split("\n", 1)[0]
        nums = re.findall(r"([\d,]+)", first_line)
        if len(nums) >= 4:
            return _parse_thousand(nums[3])
        # Try the multi-line block if all 4 weren't on same line
        block_nums = re.findall(r"([\d,]+)", rest[:200])
        if len(block_nums) >= 4:
            return _parse_thousand(block_nums[3])
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def extract_yuho(pdf_path: str) -> dict:
    """Extract all known sections from a yuho PDF.

    Returns:
      {
        "indicators_5y": {"連結": {"FY21": {...}, ...}, "単体": {...}},
        "pl": {"FY24": {"売上高": ..., ...}, "FY25": {...}},
        "sga": {"FY24": {"給料及び賞与": ..., ...}, "FY25": {...}},
        "product_breakdown": {"FY_latest": {"薬品売上": ..., ...}},
        "shares_outstanding": int (latest 期末発行済株式総数),
      }
    """
    text = _run_pdftotext(pdf_path)
    return {
        "indicators_5y": _find_5year_indicators(text),
        "pl": _find_consolidated_pl(text),
        "sga": _find_sga_breakdown(text),
        "product_breakdown": _find_product_breakdown(text),
        "shares_outstanding": _find_shares_outstanding(text),
    }


if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) != 2:
        print("Usage: python extract_yuho.py <path_to_yuho.pdf>")
        sys.exit(1)

    data = extract_yuho(sys.argv[1])
    print(json.dumps(data, ensure_ascii=False, indent=2))
