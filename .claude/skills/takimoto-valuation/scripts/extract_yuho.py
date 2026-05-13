"""Extract key financials from Japanese disclosure PDFs (yuho, eigyo, hanpo).

Supports two modes:
  1. extract_yuho(path) - single PDF
  2. extract_folder(path) - merge all PDFs in a folder

The folder mode prefers the most recent yuho_101 (annual securities report)
as the primary source, then falls back to other documents for missing data:
  - yuho_101 (有価証券報告書): most complete, P&L + segments + SG&A
  - eigyo_101 (招集通知): AGM summary, often has next-year guidance
  - hanpo_101 (半期報告書): interim, useful for current-year mid-point
  - yuho2Q (四半期報告書): quarterly, for QoQ trends

Returns numbers in 百万円.
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
    """Parse '18,234,377' or '△18,074' style number to int (raw 千円)."""
    s = s.strip().replace(",", "")
    if s.startswith("△") or s.startswith("-"):
        return -int(s.lstrip("△").lstrip("-"))
    return int(s)


def _to_million(thousand: int) -> int:
    """Convert 千円 to 百万円 (rounded)."""
    return round(thousand / 1000)


def _normalize_fw_digits(text: str) -> str:
    """Convert full-width digits to half-width."""
    return text.translate(str.maketrans("０１２３４５６７８９", "0123456789"))


# ---------------------------------------------------------------------------
# Section finders
# ---------------------------------------------------------------------------

def _find_5year_indicators(text: str) -> dict:
    """Find the 5-year main indicators block. Returns FY-keyed dict."""
    out: dict = {"連結": {}, "単体": {}}
    m = re.search(
        r"主要な経営指標等の推移.*?\(1\)\s*連結経営指標等(.*?)(?:\(2\)|提出会社の経営指標等)",
        text, re.DOTALL,
    )
    if not m:
        return out

    block = m.group(1)
    year_match = re.findall(r"(\d{4})年12月", block)
    if not year_match:
        return out
    years = [f"FY{int(y) % 100}" for y in year_match[:5]]

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
    """Find the 連結損益計算書 with 2-year comparison."""
    out: dict = {}
    m = re.search(
        r"② 【連結損益計算書及び連結包括利益計算書】\s*【連結損益計算書】(.*?)【連結包括利益計算書】",
        text, re.DOTALL,
    )
    if not m:
        return out

    block = m.group(1)
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
    """Find 販管費の主要な費目 breakdown."""
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
    """Find 売上高の内訳 by product."""
    out: dict = {}
    normalized = _normalize_fw_digits(text)
    pat_with_oku = (
        r"(薬\s*品|機\s*械|資\s*材|その\s*他)\s*売上高\s*(?:は)?\s*"
        r"(\d+)\s*億\s*(\d+)\s*百万\s*円"
    )
    for m in re.finditer(pat_with_oku, normalized):
        name = m.group(1).replace(" ", "").replace("　", "")
        key = f"{name}売上"
        oku = int(m.group(2))
        man = int(m.group(3))
        value = oku * 100 + man
        out.setdefault("FY_latest", {})[key] = value
    pat_no_oku = (
        r"(薬\s*品|機\s*械|資\s*材|その\s*他)\s*売上高\s*(?:は)?\s*"
        r"(\d{1,3})\s*百万\s*円"
    )
    for m in re.finditer(pat_no_oku, normalized):
        name = m.group(1).replace(" ", "").replace("　", "")
        key = f"{name}売上"
        if key in out.get("FY_latest", {}):
            continue
        value = int(m.group(2))
        out.setdefault("FY_latest", {})[key] = value
    return out


def _find_segment_info(text: str) -> dict:
    """Find segment information (geographic or business segment)."""
    out: dict = {}
    # Try to locate the 報告セグメント table block
    m = re.search(
        r"報告セグメントごとの売上高、利益または損失(.*?)(?:減価償却費|有形固定資産)",
        text, re.DOTALL,
    )
    if not m:
        return out
    block = m.group(1)
    # Get the year/period
    year_m = re.search(r"自\s*(\d{4})年", block)
    if year_m:
        year = f"FY{int(year_m.group(1)) % 100}"
    else:
        year = "FY_latest"
    # Look for "外部顧客への売上高" line with multiple numbers
    line_m = re.search(r"外部顧客への売上高\s+([\d,\s]+)", block)
    if line_m:
        nums = re.findall(r"([\d,]+)", line_m.group(1))
        # Drop the last number (sum) — keep individual segments
        if len(nums) >= 2:
            out.setdefault(year, {})["セグメント別売上(外部顧客)"] = [
                _to_million(_parse_thousand(n)) for n in nums
            ]
    return out


def _find_shares_outstanding(text: str) -> int | None:
    """Find latest 期末発行済株式総数."""
    m = re.search(r"発行済株式総数\s*\(株\)\s*([\d,]+)", text)
    if m:
        rest = m.string[m.end():]
        first_line = rest.split("\n", 1)[0]
        nums = re.findall(r"([\d,]+)", first_line)
        if len(nums) >= 4:
            return _parse_thousand(nums[3])
        block_nums = re.findall(r"([\d,]+)", rest[:200])
        if len(block_nums) >= 4:
            return _parse_thousand(block_nums[3])
    return None


def _find_capex_and_facilities(text: str) -> dict:
    """Find 設備投資の概要 and 主要な設備の状況."""
    out: dict = {}
    # Look for 当連結会計年度の設備投資 amount
    m = re.search(r"当連結会計年度.{0,40}?設備投資.{0,30}?([\d,]+).{0,5}百万円", text)
    if m:
        try:
            out["当期設備投資"] = int(m.group(1).replace(",", ""))
        except ValueError:
            pass
    return out


def _find_mid_term_plan(text: str) -> dict:
    """Find mid-term plan targets (中期経営計画)."""
    out: dict = {}
    # Common patterns: 「連結売上高 250億円（2027年12月期）」
    pat = r"連結売上高\s+(\d+)億円.*?(\d{4})年"
    m = re.search(pat, text)
    if m:
        out["中計売上目標(億円)"] = int(m.group(1))
        out["中計目標年度"] = f"FY{int(m.group(2)) % 100}"
    pat = r"研究開発に関する投資.*?売上高の約\s*(\d+)\s*[％%]"
    m = re.search(pat, text)
    if m:
        out["中計R&D比率目標"] = float(m.group(1)) / 100
    return out


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def extract_yuho(pdf_path: str) -> dict:
    """Extract all known sections from a single yuho PDF."""
    text = _run_pdftotext(pdf_path)
    return {
        "_source_file": Path(pdf_path).name,
        "indicators_5y": _find_5year_indicators(text),
        "pl": _find_consolidated_pl(text),
        "sga": _find_sga_breakdown(text),
        "product_breakdown": _find_product_breakdown(text),
        "segment_info": _find_segment_info(text),
        "shares_outstanding": _find_shares_outstanding(text),
        "capex": _find_capex_and_facilities(text),
        "mid_term_plan": _find_mid_term_plan(text),
    }


def _classify_pdf(filename: str) -> tuple[str, str]:
    """Classify a disclosure PDF by type and date.

    Returns (doc_type, yyyymmdd):
      doc_type ∈ {"yuho", "eigyo", "hanpo", "yuho2Q", "yuho2Q_supp", "other"}
    """
    name = Path(filename).name.lower()
    date_m = re.search(r"(\d{8})", name)
    date = date_m.group(1) if date_m else "00000000"
    if "yuho2q" in name:
        return ("yuho2Q", date)
    if "hanpo" in name:
        return ("hanpo", date)
    if "eigyo" in name:
        return ("eigyo", date)
    if "yuho_101" in name:
        return ("yuho", date)
    return ("other", date)


def extract_folder(folder_path: str) -> dict:
    """Extract financials from all PDFs in a folder, merging across documents.

    Strategy:
      - Process yuho_101 files first, oldest → newest (newer overrides older)
      - Then merge hanpo/yuho2Q for interim data points
      - Then merge eigyo for any guidance / next-period info
    """
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(folder_path)

    pdfs = sorted(folder.glob("**/*.pdf"))
    if not pdfs:
        raise ValueError(f"No PDFs found in {folder}")

    # Group by type
    by_type: dict = {}
    for pdf in pdfs:
        doc_type, date = _classify_pdf(pdf.name)
        by_type.setdefault(doc_type, []).append((date, pdf))
    for k in by_type:
        by_type[k].sort()  # oldest first

    # Merge
    merged: dict = {
        "indicators_5y": {"連結": {}, "単体": {}},
        "pl": {},
        "sga": {},
        "product_breakdown": {},
        "segment_info": {},
        "shares_outstanding": None,
        "capex": {},
        "mid_term_plan": {},
        "_sources": [],
    }

    # 1. Process yuho_101 (oldest → newest, so newer wins)
    for date, pdf in by_type.get("yuho", []):
        data = extract_yuho(str(pdf))
        merged["_sources"].append({"type": "yuho", "date": date, "file": pdf.name})
        _merge_inplace(merged, data)

    # 2. hanpo (oldest → newest)
    for date, pdf in by_type.get("hanpo", []):
        data = extract_yuho(str(pdf))
        merged["_sources"].append({"type": "hanpo", "date": date, "file": pdf.name})
        _merge_inplace(merged, data, only_missing=True)

    # 3. yuho2Q
    for date, pdf in by_type.get("yuho2Q", []):
        data = extract_yuho(str(pdf))
        merged["_sources"].append({"type": "yuho2Q", "date": date, "file": pdf.name})
        _merge_inplace(merged, data, only_missing=True)

    # 4. eigyo (招集通知): usually has guidance for next period
    for date, pdf in by_type.get("eigyo", []):
        data = extract_yuho(str(pdf))
        merged["_sources"].append({"type": "eigyo", "date": date, "file": pdf.name})
        _merge_inplace(merged, data, only_missing=True)

    return merged


def _merge_inplace(target: dict, source: dict, only_missing: bool = False):
    """Merge `source` into `target`. If only_missing, don't overwrite existing keys."""
    for section_key in ("indicators_5y", "pl", "sga", "product_breakdown",
                        "segment_info", "capex", "mid_term_plan"):
        if section_key not in source:
            continue
        src = source[section_key]
        dst = target.setdefault(section_key, {})
        if section_key == "indicators_5y":
            for level_key in ("連結", "単体"):
                src_level = src.get(level_key, {})
                dst_level = dst.setdefault(level_key, {})
                for year, metrics in src_level.items():
                    dst_year = dst_level.setdefault(year, {})
                    for k, v in metrics.items():
                        if only_missing and k in dst_year:
                            continue
                        dst_year[k] = v
        elif isinstance(src, dict):
            # Two-level dict (FY24, FY25 keys with sub-dicts)
            for year, vals in src.items():
                if isinstance(vals, dict):
                    dst_year = dst.setdefault(year, {})
                    for k, v in vals.items():
                        if only_missing and k in dst_year:
                            continue
                        dst_year[k] = v
                else:
                    if only_missing and year in dst:
                        continue
                    dst[year] = vals
    # shares_outstanding (single value)
    if source.get("shares_outstanding") is not None:
        if not (only_missing and target.get("shares_outstanding") is not None):
            target["shares_outstanding"] = source["shares_outstanding"]


if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) != 2:
        print("Usage: python extract_yuho.py <path_to_yuho.pdf_or_folder>")
        sys.exit(1)

    path = sys.argv[1]
    if Path(path).is_dir():
        data = extract_folder(path)
    else:
        data = extract_yuho(path)
    print(json.dumps(data, ensure_ascii=False, indent=2, default=str))
