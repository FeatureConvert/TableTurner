#!/usr/bin/env python3
"""
codon_trna_report.py — companion analysis script (not a GenBank submission
converter) for comparative phage/virus genomics: computes codon usage across
all annotated CDS features in a filled GenBank Annotation Template workbook,
cross-references it against the amino acid identities of annotated tRNA
genes, and flags amino acids with real codon usage but no annotated cognate
tRNA. Produces one row of summary statistics per genome/segment — built as
a lightweight per-genome feature vector for exactly the kind of comparative
analysis in Sarah Doore's lab's published work relating phage codon usage
bias, infection style, and tRNA gene counts
(https://pmc.ncbi.nlm.nih.gov/articles/PMC13015707/).

This is NOT a submission format — it reads the same workbook the other two
converter scripts use (Record Info / Sequence / Features sheets, including
multi-segment workbooks), but produces a codon usage CSV and a text report,
not anything you'd upload to NCBI.

Usage:
    python3 codon_trna_report.py input.xlsx [output_prefix]

Requires xlsx_to_genbank.py in the same directory — this script reuses its
workbook-reading, Exon Group/splicing-aware coding-sequence assembly, and
multi-segment detection logic rather than duplicating it, so codon usage is
computed from the exact same coding sequences (correct strand, correct
splicing, correct codon_start) that the submission converters would use.

Outputs:
    <prefix>_codon_usage.csv   — one row per codon: count, per-mille, amino
                                 acid, and whether that amino acid has an
                                 annotated cognate tRNA gene in this genome.
    <prefix>_trna_report.txt   — human-readable summary: genome length, GC%,
                                 CDS count, codon count, tRNA gene list (with
                                 parsed amino acid identity), and any amino
                                 acids used in codons but missing a tRNA.

Limitations:
    - tRNA identity is parsed from the tRNA row's Product text (e.g.
      "tRNA-Ile", "tRNA-fMet", "tRNA-SeC") — if a lab's Product text doesn't
      follow that convention, the parser won't recognize it and will list it
      as "unparsed" rather than guessing.
    - "Infection style" (lytic/lysogenic/chronic) isn't inferable from
      sequence alone — the output CSV has a blank column for it so you can
      fill it in from your own metadata when assembling a comparative
      dataset across many genomes.
"""
import sys
import os
import re
import csv
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import xlsx_to_genbank as gb
except ImportError:
    print("ERROR: this script must be run from the same directory as xlsx_to_genbank.py "
          "(it reuses that script's workbook-reading and splicing/strand logic).")
    sys.exit(1)

from openpyxl import load_workbook

# Standard genetic code table 11 codon -> one-letter amino acid, reused
# directly from xlsx_to_genbank's CODON_TABLE for consistency.
CODON_TABLE = gb.CODON_TABLE

AA_1_TO_3 = gb.AA_1_TO_3
AA_3_TO_1_LOWER = {v.lower(): k for k, v in AA_1_TO_3.items()}
# A few additional tRNA-naming aliases seen in real Product text beyond the
# strict 3-letter code (initiator methionine, selenocysteine, stop-codon
# suppressor/readthrough tRNAs).
AA_ALIASES_LOWER = {
    "fmet": "M", "met-e": "M", "sec": "U", "pyl": "O", "sup": None,
    "ini": "M", "initiator": "M",
}

TRNA_NAME_RE = re.compile(r"tRNA[-_ ]?([A-Za-z]+)", re.IGNORECASE)


def parse_trna_amino_acid(product_text):
    """Best-effort parse of a tRNA row's Product text (e.g. "tRNA-Ile",
    "tRNA-fMet", "tRNA-SeC", "tRNA-Leu1") into a one-letter amino acid code.
    Returns None if it doesn't recognize the text (rather than guessing)."""
    if not product_text:
        return None
    m = TRNA_NAME_RE.search(str(product_text))
    if not m:
        return None
    raw = m.group(1).strip()
    # strip trailing isoacceptor digits some labs append, e.g. "Leu1"
    letters_only = re.sub(r"\d+$", "", raw)
    key = letters_only.lower()
    if key in AA_3_TO_1_LOWER:
        return AA_3_TO_1_LOWER[key]
    if key in AA_ALIASES_LOWER:
        return AA_ALIASES_LOWER[key]
    return None


def gc_content(seq):
    if not seq:
        return 0.0
    gc = sum(1 for c in seq if c in "GC")
    return 100.0 * gc / len(seq)


def codon_usage_for_segment(record, raw_features, seq, warnings):
    """Returns (codon_counter, cds_count, trna_list) for one segment, reusing
    xlsx_to_genbank's own grouping/interval/strand-aware coding-sequence
    assembly so results reflect exactly what the submission converters
    would translate (correct splicing, correct strand, correct codon_start)."""
    features = gb.group_exons(raw_features, warnings)
    codon_counter = Counter()
    cds_count = 0
    trna_list = []

    for feat in features:
        key = feat["feature_key"]
        intervals = feat["intervals"]
        try:
            norm_intervals = [(int(float(s)), int(float(e))) for s, e in intervals]
        except (TypeError, ValueError):
            continue

        if key == "CDS":
            cds_count += 1
            region = gb.build_coding_sequence(seq, norm_intervals, feat["strand"])
            try:
                cs = int(feat.get("codon_start") or 1)
            except (TypeError, ValueError):
                cs = 1
            trimmed = region[cs - 1:]
            for i in range(0, len(trimmed) - 2, 3):
                codon = trimmed[i:i + 3]
                if len(codon) == 3 and all(c in "ACGT" for c in codon):
                    codon_counter[codon] += 1
        elif key == "tRNA":
            aa = parse_trna_amino_acid(feat.get("product"))
            trna_list.append({
                "row": feat.get("row"),
                "product": feat.get("product") or "",
                "amino_acid_1": aa,
                "amino_acid_3": AA_1_TO_3.get(aa) if aa else None,
            })

    return codon_counter, cds_count, trna_list


def build_report(xlsx_path, out_prefix):
    warnings = []
    wb = load_workbook(xlsx_path, data_only=True)
    segments = gb.detect_segments(wb)

    total_codons = Counter()
    total_cds = 0
    total_length = 0
    all_trna = []
    per_segment_summaries = []

    if segments is None:
        record = gb.read_record_info(wb["Record Info"])
        raw_features = gb.read_features(wb["Features"])
        seq = gb.read_sequence_sheet(wb, warnings)
        name = record.get("Locus/Sequence Name", "") or "SEQ1"
        codons, cds_count, trna_list = codon_usage_for_segment(record, raw_features, seq, warnings)
        total_codons.update(codons)
        total_cds += cds_count
        total_length += len(seq)
        all_trna.extend(trna_list)
        per_segment_summaries.append((name, len(seq), cds_count, len(trna_list), gc_content(seq)))
    else:
        for label, rec_ws, seq_ws, feat_ws in segments:
            record = gb.read_record_info(rec_ws)
            raw_features = gb.read_features(feat_ws)
            seq = gb.read_sequence_ws(seq_ws, warnings)
            name = record.get("Locus/Sequence Name", "") or label
            codons, cds_count, trna_list = codon_usage_for_segment(record, raw_features, seq, warnings)
            total_codons.update(codons)
            total_cds += cds_count
            total_length += len(seq)
            all_trna.extend(trna_list)
            per_segment_summaries.append((name, len(seq), cds_count, len(trna_list), gc_content(seq)))

    # Amino acids actually used in the observed codons.
    aa_usage = Counter()
    for codon, count in total_codons.items():
        aa = CODON_TABLE.get(codon)
        if aa and aa != "*":
            aa_usage[aa] += count

    trna_covered_aas = {t["amino_acid_1"] for t in all_trna if t["amino_acid_1"]}
    missing_trna = sorted(aa for aa, count in aa_usage.items() if aa not in trna_covered_aas)
    unparsed_trna = [t for t in all_trna if t["amino_acid_1"] is None]

    total_codon_count = sum(total_codons.values())

    # --- codon usage CSV ---
    csv_path = f"{out_prefix}_codon_usage.csv"
    with open(csv_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["codon", "amino_acid", "count", "per_1000_codons", "has_annotated_trna",
                    "infection_style"])
        for codon in sorted(CODON_TABLE):
            aa = CODON_TABLE[codon]
            count = total_codons.get(codon, 0)
            per_mille = (1000.0 * count / total_codon_count) if total_codon_count else 0.0
            has_trna = "" if aa == "*" else ("yes" if aa in trna_covered_aas else "no")
            w.writerow([codon, aa, count, f"{per_mille:.2f}", has_trna, ""])

    # --- text report ---
    txt_path = f"{out_prefix}_trna_report.txt"
    lines = []
    lines.append(f"Codon usage / tRNA coverage report for {xlsx_path}")
    lines.append("=" * 60)
    lines.append(f"Total genome length (all segments): {total_length} bp")
    lines.append(f"Total CDS features: {total_cds}")
    lines.append(f"Total codons counted: {total_codon_count}")
    lines.append(f"Total tRNA genes annotated: {len(all_trna)}")
    lines.append("")
    lines.append("Per-segment summary:")
    for name, length, cds_count, trna_count, gc in per_segment_summaries:
        lines.append(f"  {name}: {length} bp, GC {gc:.1f}%, {cds_count} CDS, {trna_count} tRNA")
    lines.append("")
    lines.append("Annotated tRNA genes (parsed amino acid identity):")
    if all_trna:
        for t in all_trna:
            aa_label = t["amino_acid_3"] or "UNPARSED"
            lines.append(f"  Row {t['row']}: \"{t['product']}\" -> {aa_label}")
    else:
        lines.append("  (none annotated)")
    if unparsed_trna:
        lines.append("")
        lines.append("tRNA Product text this script could not parse into an amino acid "
                      "(listed as UNPARSED above) — check these by hand:")
        for t in unparsed_trna:
            lines.append(f"  Row {t['row']}: \"{t['product']}\"")
    lines.append("")
    if missing_trna:
        aa3 = ", ".join(AA_1_TO_3.get(a, a) for a in missing_trna)
        lines.append(f"Amino acids with real codon usage but NO annotated cognate tRNA gene: {aa3}")
        lines.append("(This may simply mean the corresponding host tRNA is used in trans — many "
                      "phages don't encode a full tRNA set — not necessarily a data-entry problem.)")
    else:
        lines.append("Every amino acid with nonzero codon usage has at least one annotated "
                      "cognate tRNA gene.")
    lines.append("")
    if warnings:
        lines.append(f"{len(warnings)} warning(s) from the underlying workbook read:")
        for w_ in warnings:
            lines.append(f"  - {w_}")

    with open(txt_path, "w") as f:
        f.write("\n".join(lines) + "\n")

    return csv_path, txt_path, warnings


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 codon_trna_report.py input.xlsx [output_prefix]")
        sys.exit(1)
    xlsx_path = sys.argv[1]
    out_prefix = sys.argv[2] if len(sys.argv) > 2 else re.sub(r"\.xlsx$", "", xlsx_path)

    csv_path, txt_path, warnings = build_report(xlsx_path, out_prefix)
    print(f"Wrote {csv_path}")
    print(f"Wrote {txt_path}")
    if warnings:
        print(f"\n{len(warnings)} warning(s) — see {txt_path} for details.")


if __name__ == "__main__":
    main()
