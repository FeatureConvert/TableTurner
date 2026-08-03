---
name: "genbank-converter"
description: "Convert a filled-in \"GenBank Annotation Template\" spreadsheet (.xlsx, with Record Info, Sequence, and Features sheets) into either NCBI's 5-column tab-delimited feature table (for BankIt/WebSub upload) or a spec-compliant GenBank flat file (.gb). Use when the user has phage/gene sequence and feature annotation data in an Excel spreadsheet and wants a feature table or GenBank flat file, mentions converting a filled annotation template, or asks to turn spreadsheet gene/CDS/tRNA feature data into a submission-ready file."
---

## What this skill does

Converts a filled-in copy of the "GenBank Annotation Template" workbook into
either of two outputs:

1. **NCBI's 5-column tab-delimited feature table** — the exact file format
   uploaded in the "Feature" step of NCBI BankIt/WebSub submissions. This is
   the format most wet-lab phage/small-genome annotation protocols actually
   submit (see "Design notes" below).
2. **A full GenBank flat file (.gb)** — LOCUS/DEFINITION/FEATURES/ORIGIN
   blocks, for record-keeping or non-BankIt routes.

Both scripts read the same workbook, which has three sheets:

- **Record Info** — key/value rows: Locus/Sequence Name, Definition,
  Organism, Molecule Type, Topology, Division, Accession, Strain, Isolate,
  Collection Date, Country, Comment. Locus/Sequence Name doubles as the
  feature table's sequence ID and the locus_tag prefix (e.g. "Moonfish_gp1").
- **Sequence** — the nucleotide sequence, in column A starting at row 2, ONE
  CHUNK PER CELL if it's long. Excel caps a single cell at ~32,767
  characters, which real genomes routinely exceed (e.g. a 69 kb phage
  genome needs ~4 cells at 20,000 chars each) — always read/write this
  sheet by concatenating all non-empty cells in column A from row 2 down, in
  order. Never assume the whole sequence fits in one cell.
- **Features** — one row per feature, in genome order. Columns: Feature Key
  (CDS / tRNA / rRNA / misc_feature / gene), Start, End, Gene # (drives
  locus_tag "{Name}_gp{N}" and the short label "gp{N}"), Product,
  transl_table (blank defaults to 11 — bacteria/archaea/phage code), Codon
  Start (blank defaults to 1), Note, plus optional gray-header columns
  (Protein ID, db_xref, Other Qualifiers as key=value|key=value, Partial 5'
  end Y/N, Partial 3' end Y/N) used by the .gb output.

  **Strand is not a separate column.** If Start < End the feature is on the
  + strand; if Start > End (a higher number first) it's on the - strand —
  users enter the two numbers in that order directly, matching how these
  wet-lab protocols already record gene coordinates. Don't add a strand
  column back in.

  **CDS rows do not need a paired gene row.** Both converters automatically
  write the paired `gene` feature (with locus_tag) above/before each CDS —
  confirmed one-to-one against a real published record (82 CDS, 82 gene).

## How to run it

This skill is bundled directly in the TableTurner repo at
`.claude/skills/genbank-converter/`, so it loads automatically for anyone
using Claude Code or Cowork with this repo folder open — no separate
install step needed.

1. Locate the user's filled .xlsx.
2. The converter scripts (`xlsx_to_feature_table.py`, `xlsx_to_genbank.py`)
   live at the repo root. Ask the user (or infer from context —
   "BankIt"/"WebSub"/"feature table" implies the first script; "GenBank flat
   file"/".gb" implies the second) which output they want, or produce both
   if unclear.
3. Run: `python3 xlsx_to_feature_table.py <input.xlsx> <output.tbl.txt>`
   and/or `python3 xlsx_to_genbank.py <input.xlsx> <output.gb>`
   (openpyxl must be available — `pip install openpyxl --break-system-packages`
   if the import fails).
4. Read back and report any warnings — they flag real data issues (missing
   CDS product, out-of-range coordinates, mismatched gene/CDS spans,
   translations that come out empty because of a coordinate typo) rather
   than being fatal. Treat a translation-came-out-empty warning as "go check
   this row in the source data," not a bug.
5. Share the resulting file(s) with the user.

For CDS-bearing .gb output, sanity-check with Biopython if available:
`python3 -c "from Bio import SeqIO; r = SeqIO.read('output.gb','genbank'); print(len(r.seq), len(r.features))"`
(`pip install biopython --break-system-packages` first if needed). A clean
parse plus translations that recompute identically from the DNA
(`f.extract(r.seq).translate(table=11, to_stop=True, cds=True)`) is strong
evidence the file is well-formed — note the `cds=True` flag, which applies
the same "first codon of a CDS is always Met" rule the converter itself
implements; without it, plain `.translate()` will show false mismatches on
any gene that starts with an alternate start codon like GTG or TTG.

## Design notes / known limitations

- This schema is modeled directly on a real lab's dsDNA phage sequencing +
  annotation protocol (GeneMark S ORF calling → BLAST/InterPro predictions
  tracked in a spreadsheet → hand-typed 5-column feature table → BankIt
  upload). transl_table=11 and codon_start=1 defaults, the Gene #-driven
  locus_tag/gp-label convention, and the "no source row in the feature
  table" behavior all mirror that real workflow.
- Single contiguous interval per feature row only — no spliced join()
  locations for multi-exon CDS/mRNA.
- One sequence record per workbook.
- Validated end-to-end against a real 85-feature, 69 kb phage annotation
  (see `examples/moonfish/` in this repo for the full worked example and
  validation report), and triple-checked feature-by-feature against that
  phage's actual published NCBI record.
