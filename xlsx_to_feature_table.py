#!/usr/bin/env python3
"""
xlsx_to_feature_table.py — convert a filled copy of GenBank_Annotation_Template.xlsx
into NCBI's 5-column tab-delimited feature table: the exact file format uploaded
in the "Feature" step of NCBI BankIt/WebSub submissions.

Usage:
    python3 xlsx_to_feature_table.py input.xlsx [output.tbl.txt]

Format (per https://www.ncbi.nlm.nih.gov/genbank/feature_table/ and matching
the structure of a real lab feature table, Feature Moonfish_annotation_KP.txt):

    >Feature	<Locus/Sequence Name>
    <start>	<end>	gene
    			locus_tag	<Name>_gp<N>
    <start>	<end>	CDS
    			product	<descriptive product>
    			product	gp<N>
    			transl_table	11
    			codon_start	1
    			note	<note>            (only if a note was given)

Notes on conventions implemented, matching this lab's protocol:
  - Strand is NOT a separate column: if Start < End the feature is '+', if
    Start > End it's '-' — you enter the numbers in that order directly, and
    they're passed straight through to the two location columns.
  - Every CDS row automatically gets a paired 'gene' feature written above it
    (with a locus_tag, if a Gene # was given) — you do not enter a separate
    gene row for each CDS.
  - CDS features always get transl_table (default 11) and codon_start
    (default 1) qualifiers, matching standard phage annotation practice.
  - Partial features get the standard '<'/'>' markers on the start/end
    columns (column 1 gets '<', column 2 gets '>', regardless of strand).
  - There is no 'source' feature in this file — BankIt collects organism/
    strain/collection date/etc. through its own web form, not the feature
    table upload.
  - Qualifier lines always use real tabs (three leading tabs, then key, tab,
    value) and blank/empty qualifiers are simply omitted, so the output
    doesn't carry the stray blank-line and missing-tab artifacts sometimes
    seen in hand-edited feature tables.
"""
import sys
import re
from openpyxl import load_workbook

VALID_FEATURE_KEYS = {"gene", "CDS", "tRNA", "rRNA", "misc_feature"}
IUPAC_NT = set("ACGTUNRYSWKMBDHV")


class ConversionError(Exception):
    pass


def warn(msg, warnings):
    warnings.append(msg)


def read_sequence_length(wb, warnings):
    """Sequence lives on its own sheet (column A, from row 2 down, one chunk
    per cell) because Excel caps a single cell at ~32,767 characters, which
    real genomes routinely exceed. Only used here to sanity-check feature
    coordinates; the feature table itself doesn't carry the sequence."""
    if "Sequence" not in wb.sheetnames:
        warn("No 'Sequence' sheet found — skipping coordinate range checks.", warnings)
        return None
    ws = wb["Sequence"]
    chunks = []
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
        v = row[0].value
        if v:
            chunks.append(str(v).strip())
    raw = "".join(chunks)
    if not raw:
        warn("The 'Sequence' sheet is empty — skipping coordinate range checks.", warnings)
        return None
    seq = re.sub(r"[\s\d]", "", raw).upper()
    bad = set(seq) - IUPAC_NT
    if bad:
        warn(f"Sequence contains unexpected characters: {sorted(bad)}", warnings)
    return len(seq)


def read_record_info(ws):
    data = {}
    for row in ws.iter_rows(min_row=2, max_col=2):
        label_cell, value_cell = row[0], row[1]
        if label_cell.value is None:
            continue
        label = str(label_cell.value).replace("*", "").strip()
        value = value_cell.value
        value = "" if value is None else str(value).strip()
        data[label] = value
    return data


def read_features(ws):
    header_row = 2
    headers = {}
    for cell in ws[header_row]:
        if cell.value:
            headers[str(cell.value).strip()] = cell.column
    required = ["Feature Key", "Start", "End"]
    for r in required:
        if r not in headers:
            raise ConversionError(f"Features sheet is missing required column '{r}'.")

    def find_col(*names):
        for n in names:
            if n in headers:
                return headers[n]
        return None

    col_map = {
        "feature_key": find_col("Feature Key"),
        "start": find_col("Start"),
        "end": find_col("End"),
        "gene_num": find_col("Gene #"),
        "product": find_col("Product"),
        "transl_table": find_col("transl_table"),
        "codon_start": find_col("Codon Start"),
        "note": find_col("Note"),
        "other": find_col("Other Qualifiers (gb only, key=value|key=value)"),
        "partial5": find_col("Partial 5' end (Y/N)"),
        "partial3": find_col("Partial 3' end (Y/N)"),
    }

    def get(row, key):
        col = col_map.get(key)
        if col is None:
            return ""
        v = ws.cell(row=row, column=col).value
        return "" if v is None else (str(v).strip() if not isinstance(v, (int, float)) else v)

    features = []
    for r in range(header_row + 1, ws.max_row + 1):
        fkey = get(r, "feature_key")
        if not fkey or str(fkey).strip() == "":
            continue
        note_val = get(r, "note")
        if isinstance(note_val, str) and note_val.strip().lower().startswith("example row"):
            note_val = ""
        features.append({
            "row": r,
            "feature_key": str(fkey).strip(),
            "start": get(r, "start"),
            "end": get(r, "end"),
            "gene_num": get(r, "gene_num"),
            "product": get(r, "product"),
            "transl_table": get(r, "transl_table"),
            "codon_start": get(r, "codon_start"),
            "note": note_val,
            "other": get(r, "other"),
            "partial5": str(get(r, "partial5") or "N").strip().upper() == "Y",
            "partial3": str(get(r, "partial3") or "N").strip().upper() == "Y",
        })
    return features


def parse_other_qualifiers(text):
    pairs = []
    if not text:
        return pairs
    for chunk in str(text).split("|"):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "=" in chunk:
            k, v = chunk.split("=", 1)
            pairs.append((k.strip(), v.strip()))
        else:
            pairs.append((chunk, None))
    return pairs


def loc_cols(start, end, partial5, partial3):
    try:
        s = str(int(float(start)))
        e = str(int(float(end)))
    except (TypeError, ValueError):
        return None, None
    if partial5:
        s = "<" + s
    if partial3:
        e = ">" + e
    return s, e


def qual_line(key, value=None, warnings=None, row=None):
    # This format is tab/newline-delimited, so a stray tab or line break
    # inside a value would silently corrupt the column structure — collapse
    # any such whitespace to single spaces instead.
    if value is not None and ("\t" in str(value) or "\n" in str(value) or "\r" in str(value)):
        if warnings is not None:
            warn(f"Row {row}: qualifier '{key}' contained a tab/newline character — replaced with a space "
                 f"(the feature table format uses tabs/newlines as delimiters).", warnings)
        value = re.sub(r"[\t\r\n]+", " ", str(value)).strip()
    if value is None or value == "":
        return f"\t\t\t{key}"
    return f"\t\t\t{key}\t{value}"


def build_feature_table(record, features, seq_length, warnings):
    name = record.get("Locus/Sequence Name", "") or "SEQ1"
    lines = [f">Feature\t{name}"]

    for feat in features:
        key = feat["feature_key"]
        if key not in VALID_FEATURE_KEYS:
            warn(f"Row {feat['row']}: unrecognized feature key '{key}' — skipped.", warnings)
            continue

        def ql(k, v=None, _row=feat["row"]):
            return qual_line(k, v, warnings, _row)

        s, e = loc_cols(feat["start"], feat["end"], feat["partial5"], feat["partial3"])
        if s is None:
            warn(f"Row {feat['row']}: {key} has non-numeric Start/End — skipped.", warnings)
            continue
        if seq_length is not None:
            try:
                lo = min(float(feat["start"]), float(feat["end"]))
                hi = max(float(feat["start"]), float(feat["end"]))
                if lo < 1 or hi > seq_length:
                    warn(f"Row {feat['row']}: {key} location {feat['start']}..{feat['end']} is out of range "
                         f"for a {seq_length}-bp sequence — check this row.", warnings)
            except (TypeError, ValueError):
                pass

        gene_num = feat["gene_num"]
        locus_tag = f"{name}_gp{int(float(gene_num))}" if gene_num not in ("", None) else None
        gp_label = f"gp{int(float(gene_num))}" if gene_num not in ("", None) else None

        if key == "CDS":
            # paired gene feature first
            lines.append(f"{s}\t{e}\tgene")
            if locus_tag:
                lines.append(ql("locus_tag", locus_tag))
            else:
                warn(f"Row {feat['row']}: CDS has no Gene # — gene feature written without a locus_tag.", warnings)

            lines.append(f"{s}\t{e}\tCDS")
            if not feat["product"]:
                warn(f"Row {feat['row']}: CDS has no Product — BankIt requires one.", warnings)
            lines.append(ql("product", feat["product"] or "hypothetical protein"))
            if gp_label:
                lines.append(ql("product", gp_label))
            transl_table = feat["transl_table"] or "11"
            codon_start = feat["codon_start"] or "1"
            lines.append(ql("transl_table", transl_table))
            lines.append(ql("codon_start", codon_start))
            if feat["note"]:
                lines.append(ql("note", feat["note"]))
            for k, v in parse_other_qualifiers(feat["other"]):
                lines.append(ql(k, v))

        elif key == "gene":
            lines.append(f"{s}\t{e}\tgene")
            if locus_tag:
                lines.append(ql("locus_tag", locus_tag))
            if feat["note"]:
                lines.append(ql("note", feat["note"]))
            for k, v in parse_other_qualifiers(feat["other"]):
                lines.append(ql(k, v))

        else:  # tRNA, rRNA, misc_feature
            lines.append(f"{s}\t{e}\t{key}")
            if feat["product"]:
                lines.append(ql("product", feat["product"]))
            elif key in ("tRNA", "rRNA"):
                warn(f"Row {feat['row']}: {key} has no Product — recommended (e.g. \"tRNA-Ile\").", warnings)
            if feat["note"]:
                lines.append(ql("note", feat["note"]))
            for k, v in parse_other_qualifiers(feat["other"]):
                lines.append(ql(k, v))

    return "\n".join(lines) + "\n"


def convert(xlsx_path, out_path):
    warnings = []
    wb = load_workbook(xlsx_path, data_only=True)
    if "Record Info" not in wb.sheetnames:
        raise ConversionError("Workbook has no 'Record Info' sheet.")
    if "Features" not in wb.sheetnames:
        raise ConversionError("Workbook has no 'Features' sheet.")

    record = read_record_info(wb["Record Info"])
    features = read_features(wb["Features"])
    seq_length = read_sequence_length(wb, warnings)

    text = build_feature_table(record, features, seq_length, warnings)

    with open(out_path, "w", newline="\n") as f:
        f.write(text)

    return text, warnings


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 xlsx_to_feature_table.py input.xlsx [output.tbl.txt]")
        sys.exit(1)
    xlsx_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else re.sub(r"\.xlsx$", ".tbl.txt", xlsx_path)

    try:
        text, warnings = convert(xlsx_path, out_path)
    except ConversionError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    print(f"Wrote {out_path}")
    if warnings:
        print(f"\n{len(warnings)} warning(s):")
        for w in warnings:
            print(f"  - {w}")
    else:
        print("No warnings.")


if __name__ == "__main__":
    main()
