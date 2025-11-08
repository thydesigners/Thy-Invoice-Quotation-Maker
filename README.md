# Thy-Invoice-Quotation-Maker
import os
import re
import time
import glob
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

from docx import Document
from docx.table import _Row
from copy import deepcopy

# Optional: words in Indian numbering system
try:
    from num2words import num2words
except ImportError:
    num2words = None

# Optional: DOCX -> PDF (requires MS Word on Windows/macOS)
try:
    from docx2pdf import convert
except Exception:
    convert = None


# ---------- Formatting helpers ----------

def amount_to_words(amount):
    if num2words is None:
        return ""
    try:
        amount = int(Decimal(str(amount)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        words = num2words(amount, lang="en_IN").replace(",", "").title()
        return f"{words} Only"
    except Exception:
        return ""


def indian_group(int_str):
    if len(int_str) <= 3:
        return int_str
    pre = int_str[-3:]
    rest = int_str[:-3]
    groups = []
    while len(rest) > 2:
        groups.insert(0, rest[-2:])
        rest = rest[:-2]
    if rest:
        groups.insert(0, rest)
    return ",".join(groups + [pre])


def fmt_money_indian(x, force_round=False):
    try:
        q = Decimal(str(x))
        if force_round:
            q = q.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
            s = f"{q:.2f}"
        else:
            q = q.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            s = f"{q:.2f}"
        int_part, dec_part = s.split(".")
        sign = ""
        if int_part.startswith("-"):
            sign = "-"
            int_part = int_part[1:]
        int_part_grp = indian_group(int_part if int_part else "0")
        return f"{sign}{int_part_grp}.{dec_part}"
    except Exception:
        return str(x)


def fmt_hsn(h):
    s = str(h).strip()
    if s == "" or s.lower() == "nan":
        return ""
    try:
        f = float(s)
        if abs(f - int(f)) < 1e-9:
            return str(int(f))
        return s
    except Exception:
        return s


def quantity_display(raw_q):
    if raw_q is None:
        return ""
    if isinstance(raw_q, str):
        s = raw_q.strip()
        if s == "":
            return ""
        try:
            f = float(s.replace(",", ""))
        except Exception:
            return s
        if "." in s:
            return f"{f:.10f}".rstrip("0").rstrip(".")
        else:
            return str(int(round(f)))
    try:
        f = float(raw_q)
        if f.is_integer():
            return str(int(f))
        else:
            return f"{f:.10f}".rstrip("0").rstrip(".")
    except Exception:
        return str(raw_q)


def coalesce(*vals):
    for v in vals:
        if v is None:
            continue
        if isinstance(v, float) and pd.isna(v):
            continue
        if str(v).strip().lower() == "nan":
            continue
        if str(v).strip() != "":
            return v
    return ""


def to_float(x, default=0.0):
    try:
        if x is None:
            return default
        if isinstance(x, str):
            x = x.replace(",", "").strip()
            if x == "":
                return default
        return float(x)
    except Exception:
        return default


def normalize_rate(rate):
    r = to_float(rate, 0.0)
    if r == 0:
        return 0.0
    if 0 < r <= 1:
        return r * 100.0
    return r


def percent_str(p):
    try:
        return f"{float(p):g}%"
    except Exception:
        return ""


def rupee_round(x):
    try:
        return int(Decimal(str(float(x))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except Exception:
        return int(float(x))


def int_only_str(x):
    """
    Return integer part as a string (for values like 101, '101', 101.0, '101.0').
    Empty string for None/blank. Falls back to original string if non-numeric.
    """
    if x is None:
        return ""
    s = str(x).strip()
    if s == "":
        return ""
    try:
        return str(int(float(s)))
    except Exception:
        return s


# ---------- File naming / saving helpers ----------

def sanitize_filename(name: str, maxlen: int = 180) -> str:
    name = name.replace("–", "-").replace("—", "-")
    name = re.sub(r'[<>:"/\\|?*]', "-", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = re.sub(r"[-_]{2,}", "-", name)
    return name[:maxlen].rstrip(". ")


def safe_save_document(doc, base_path: str, max_tries: int = 10, wait_secs: float = 0.4) -> str:
    root, ext = os.path.splitext(base_path)
    attempt = 0
    while attempt < max_tries:
        try_path = base_path if attempt == 0 else f"{root} ({attempt}){ext}"
        try:
            doc.save(try_path)
            return try_path
        except PermissionError:
            time.sleep(wait_secs)
            attempt += 1
        except OSError:
            if attempt == 0:
                safe_root = sanitize_filename(root)
                base_path = f"{safe_root}{ext}"
                root, ext = os.path.splitext(base_path)
            else:
                time.sleep(wait_secs)
                attempt += 1
    from time import strftime
    fallback = f"{root} ({strftime('%Y%m%d-%H%M%S')}){ext}"
    doc.save(fallback)
    return fallback


# ---------- Cross-run placeholder replacement (preserve fonts/sizes) ----------

def _build_run_index(paragraph):
    runs = paragraph.runs
    spans = []
    pos = 0
    for i, r in enumerate(runs):
        t = r.text or ""
        spans.append((i, pos, pos + len(t)))
        pos += len(t)
    full = "".join(r.text or "" for r in runs)
    return full, spans


def _find_all_occurrences(text, needle):
    i = 0
    while True:
        j = text.find(needle, i)
        if j == -1:
            return
        yield j
        i = j + len(needle)


def _replace_span_in_runs(paragraph, spans, start, end, replacement):
    runs = paragraph.runs
    i_start = i_end = None
    off_start = off_end = None
    for (ri, s, e) in spans:
        if s <= start < e:
            i_start, off_start = ri, start - s
        if s < end <= e:
            i_end, off_end = ri, end - s
    if i_start is None or i_end is None:
        return
    if i_start == i_end:
        t = runs[i_start].text or ""
        runs[i_start].text = t[:off_start] + replacement + t[off_end:]
    else:
        t0 = runs[i_start].text or ""
        runs[i_start].text = t0[:off_start] + replacement
        for ri in range(i_start + 1, i_end):
            runs[ri].text = ""
        tn = runs[i_end].text or ""
        runs[i_end].text = tn[off_end:]


def replace_in_paragraph(paragraph, mapping):
    if not paragraph.runs:
        return
    changed = True
    while changed:
        changed = False
        full, spans = _build_run_index(paragraph)
        for k, v in sorted(mapping.items(), key=lambda kv: len(kv[0]), reverse=True):
            placeholder = f"{{{{{k}}}}}"
            for pos in list(_find_all_occurrences(full, placeholder)):
                _replace_span_in_runs(paragraph, spans, pos, pos + len(placeholder), str(v))
                changed = True
                full, spans = _build_run_index(paragraph)


def replace_placeholders_everywhere(doc, mapping):
    for p in doc.paragraphs:
        replace_in_paragraph(p, mapping)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, mapping)


# ---------- Row cloning & item filling ----------

def fill_item_row_in_place(row, item_mapping):
    """Replace placeholders in the given row (preserves fonts/sizes/borders)."""
    for cell in row.cells:
        for p in cell.paragraphs:
            full = "".join(r.text or "" for r in p.runs)
            if "{{" in full and "}}" in full:
                replace_in_paragraph(p, item_mapping)


def insert_row_after_using_template(tbl, after_row, template_tr):
    """Insert a brand-new row cloned from the pristine template_tr."""
    new_tr = deepcopy(template_tr)
    after_row._tr.addnext(new_tr)
    return _Row(new_tr, tbl)


# ---------- Placeholder-driven item-row detection ----------

ITEM_PLACEHOLDERS_SETS = [
    {"S. No.", "S_No", "Sl.No"},
    {"Product", "Title", "Particulars"},
    {"Unit_Price"},
    {"Quantity"},
    {"Amount"},
]

def row_contains_item_placeholders(row_text):
    score = 0
    for group in ITEM_PLACEHOLDERS_SETS:
        if any(f"{{{{{g}}}}}" in row_text for g in group):
            score += 1
    return score


def find_item_placeholder_row(doc):
    best = (None, None, 0)
    for tbl in doc.tables:
        for i, row in enumerate(tbl.rows):
            row_text = ""
            for cell in row.cells:
                for p in cell.paragraphs:
                    row_text += "".join(r.text or "" for r in p.runs) + " "
            score = row_contains_item_placeholders(row_text)
            if score > best[2]:
                best = (tbl, i, score)
    tbl, idx, score = best
    if tbl is not None and score >= 2:
        return tbl, idx
    return None, None


# ---------- Company template selection ----------

def find_company_template(company: str, search_dir: str) -> str | None:
    if not company:
        return None
    target = f"{company}.docx".lower()
    docs = glob.glob(os.path.join(search_dir, "*.docx"))
    for p in docs:
        if os.path.basename(p).lower() == target:
            return p
    for p in docs:
        base = os.path.splitext(os.path.basename(p))[0].lower()
        comp = company.lower()
        if base.startswith(comp) or comp in base:
            return p
    return None


def pick_default_template(search_dir: str) -> str | None:
    preferred = os.path.join(search_dir, "invoice_template.docx")
    if os.path.exists(preferred):
        return preferred
    docs = glob.glob(os.path.join(search_dir, "*.docx"))
    return docs[0] if docs else None


# ---------- Main ----------

def main():
    cwd = os.getcwd()

    # Excel detection
    candidate_excels = [p for p in glob.glob(os.path.join(cwd, "*.xlsx"))
                        if os.path.basename(p).lower().startswith("datainvoice")]
    excels = candidate_excels or glob.glob(os.path.join(cwd, "*.xlsx"))
    if not excels:
        print("⚠️ No Excel file found in folder.")
        input("Press any key to exit...")
        return
    excel_path = excels[0]

    out_dir = os.path.join(cwd, "out_invoices")
    os.makedirs(out_dir, exist_ok=True)

    # --- Read Excel ---
    df = pd.read_excel(excel_path)
    df = df.astype(object).where(pd.notna(df), None)

    # --- Forward-fill invoice-level fields safely ---
    pd.set_option('future.no_silent_downcasting', True)
    for col in ["INV_No", "File_No", "Date", "Party_Address", "Party_GST", "Title", "Company", "GST %", "SGST", "CGST", "HSN_Code"]:
        if col in df.columns:
            df[col] = df[col].astype("object")
            df[col] = df[col].ffill()
            df[col] = df[col].infer_objects(copy=False)

    if "File_No" not in df.columns and "INV_No" not in df.columns:
        print("⚠️ 'File_No' or 'INV_No' column is required in Excel.")
        input("Press any key to exit...")
        return

    # Keep only groups that have at least File_No or INV_No
    df = df[(df.get("File_No").notna() if "File_No" in df.columns else False) |
            (df.get("INV_No").notna() if "INV_No" in df.columns else False)]
    if df.empty:
        print("⚠️ No rows with a File_No/INV_No group.")
        input("Press any key to exit...")
        return

    # Grouping key: prefer File_No, else INV_No, both as integer-only strings
    def group_key(row):
        fn = int_only_str(row.get("File_No"))
        if fn:
            return fn
        return int_only_str(row.get("INV_No"))

    df["_GROUP_KEY"] = df.apply(group_key, axis=1)

    for group_id, g in df.groupby("_GROUP_KEY"):
        first = g.iloc[0]

        # --------- Company-specific template selection ---------
        company = coalesce(first.get("Company"), "").strip()
        template_path = find_company_template(company, cwd)
        if not template_path:
            template_path = pick_default_template(cwd)
            if not template_path:
                print("⚠️ No DOCX template found in folder.")
                input("Press any key to exit...")
                return
            else:
                print(f"ℹ️ Template for company '{company}' not found. Using default: {os.path.basename(template_path)}")
        else:
            print(f"ℹ️ Using template for company '{company}': {os.path.basename(template_path)}")

        # --------- Build items (Amount = Unit×Qty) ---------
        items = []
        for _, r in g.iterrows():
            if not any(r.get(c) is not None for c in ["Product", "Unit_Price", "Quantity"]):
                continue

            unit_price = to_float(r.get("Unit_Price", 0))
            qty_raw = r.get("Quantity", 0)
            qty_num = to_float(qty_raw, 0)
            amount_num = unit_price * qty_num

            # S. No. strictly integer (fallback to index 1..N)
            sno_raw = coalesce(r.get("S. No."), r.get("Sl.No"), r.get("S_No"))
            sno = int_only_str(sno_raw) or str(len(items) + 1)

            items.append({
                "S. No.": sno,  # already integer string
                "Product": coalesce(r.get("Product"), ""),
                "HSN_Code": r.get("HSN_Code", ""),
                "Unit_Price": unit_price,
                "Quantity": qty_raw,
                "Amount": amount_num,
            })
        if not items:
            continue

        # --------- Totals & Taxes ---------
        total_amount = sum(to_float(it["Amount"]) for it in items)
        gst_pct = normalize_rate(first.get("GST %", 0))
        sgst_rate = normalize_rate(first.get("SGST", 0))
        cgst_rate = normalize_rate(first.get("CGST", 0))
        if sgst_rate <= 0 and cgst_rate <= 0 and gst_pct > 0:
            sgst_rate = cgst_rate = gst_pct / 2.0
        sgst_amount = total_amount * (sgst_rate / 100.0)
        cgst_amount = total_amount * (cgst_rate / 100.0)
        grand_total = total_amount + sgst_amount + cgst_amount
        grand_total_rounded = rupee_round(grand_total)

        # --------- Header mapping ---------
        inv_no_val  = coalesce(first.get("INV_No"), "")
        file_no_val = coalesce(first.get("File_No"), "")
        inv_no_str  = int_only_str(inv_no_val)  # integer-only if present
        file_no_str = int_only_str(file_no_val)  # integer-only

        title_val = coalesce(first.get("Title"), items[0].get("Product"), "")
        party_addr = coalesce(first.get("Party_Address"), "")
        party_gst = coalesce(first.get("Party_GST"), "")
        date_val  = coalesce(first.get("Date"), "")
        comp_val  = company or str(party_addr).split(",")[0][:40]

        mapping = {
            "INV_No": inv_no_str,        # still available if your template shows it
            "File_No": file_no_str,      # add {{File_No}} in template if you want it printed
            "Date": date_val,
            "Party_Address": party_addr,
            "Party_GST": party_gst,
            "Title": title_val,
            "Particulars": title_val,
            "Total_Amount": fmt_money_indian(total_amount),
            "SGST": percent_str(sgst_rate),
            "CGST": percent_str(cgst_rate),
            "SGSTAmount": fmt_money_indian(sgst_amount),
            "CGSTAmount": fmt_money_indian(cgst_amount),
            "Grand_Total": fmt_money_indian(grand_total_rounded, force_round=True),
            "Amount_in_Words": amount_to_words(grand_total_rounded),
            "Company": comp_val,
        }

        # --------- Open template & replace header placeholders ---------
        doc = Document(template_path)
        replace_placeholders_everywhere(doc, mapping)

        # --------- Find item placeholder row (by placeholders only) ---------
        items_tbl, item_row_idx = find_item_placeholder_row(doc)
        if items_tbl is None:
            print("⚠️ Could not find an item placeholder row (e.g., {{Product}}, {{Quantity}}, {{Amount}}). "
                  "Please ensure your template has one row with item placeholders.")
        else:
            # Keep a pristine copy of the placeholder row BEFORE filling
            placeholder_row = items_tbl.rows[item_row_idx]
            pristine_template_tr = deepcopy(placeholder_row._tr)

            def build_item_mapping(it):
                unit = fmt_money_indian(it["Unit_Price"])
                qty  = quantity_display(it["Quantity"])
                amt  = fmt_money_indian(it["Amount"])
                s_no = it.get("S. No.", "")  # already integer string
                prod = str(coalesce(it.get("Product"), it.get("Title"), ""))
                return {
                    "S. No.": s_no,
                    "S_No": s_no,
                    "Sl.No": s_no,
                    "Product": prod,
                    "Title": prod,
                    "Particulars": prod,
                    "HSN_Code": fmt_hsn(it.get("HSN_Code", "")),
                    "Unit_Price": unit,
                    "Quantity": qty,
                    "Amount": amt,
                }

            # First item into the existing row
            fill_item_row_in_place(placeholder_row, build_item_mapping(items[0]))

            # Subsequent items: clone pristine row then fill
            after_row = placeholder_row
            for it in items[1:]:
                new_row = insert_row_after_using_template(items_tbl, after_row, pristine_template_tr)
                fill_item_row_in_place(new_row, build_item_mapping(it))
                after_row = new_row

        # ---------- File naming: STRICTLY "File_No - Title.docx" ----------
        title_clean = sanitize_filename(title_val) or "Invoice"
        file_no_clean = sanitize_filename(file_no_str) if file_no_str else ""
        if not file_no_clean:
            print(f"ℹ️ Warning: File_No is blank for group '{group_id}'. Using '{title_clean}.docx' as filename.")
            out_name = title_clean
        else:
            out_name = f"{file_no_clean} - {title_clean}"

        out_docx = os.path.join(out_dir, out_name + ".docx")
        saved_path = safe_save_document(doc, out_docx)
        print(f"✅ Created DOCX: {saved_path}")

        out_pdf = os.path.splitext(saved_path)[0] + ".pdf"
        if convert:
            try:
                convert(saved_path, out_pdf)
                print(f"✅ Created PDF: {out_pdf}")
            except Exception as e:
                print(f"⚠️ Could not create PDF for {saved_path}: {e}")
        else:
            print("ℹ️ docx2pdf not available; only DOCX created.")

    print("\n🎉 All invoices processed successfully!")
    input("Press any key to exit...")


if __name__ == "__main__":
    import traceback, datetime
    try:
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        pass
    try:
        main()
    except Exception as e:
        log_path = os.path.join(os.getcwd(), "invoice_error.log")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now()}]\n")
            f.write(traceback.format_exc())
        print("\n❌ A fatal error occurred. Details saved to invoice_error.log")
        print(f"   Location: {log_path}\n")
        input("Press Enter to close...")
