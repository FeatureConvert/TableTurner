# GenBank Annotation Template → Feature Table / GenBank Flat File

Two scripts that convert a filled-in "GenBank Annotation Template" Excel
workbook into NCBI submission-ready output:

| Script | Produces | Use for |
|---|---|---|
| `xlsx_to_feature_table.py` | NCBI's 5-column tab-delimited feature table (`.tbl.txt`) | Uploading via BankIt / WebSub |
| `xlsx_to_genbank.py` | A full GenBank flat file (`.gb`) | Record-keeping, or any non-BankIt submission route |

Both read the same workbook, so you only fill in the spreadsheet once and
can generate either or both outputs from it.

Originally built for non-spliced dsDNA phage genomes, both scripts now also
support polyprotein viruses, spliced eukaryotic viral genes, and RNA virus
genomes. See `PROTOCOL.md` for the annotation conventions to follow for each
virus type, and "Which virus types this covers" below for a quick summary.

## Requirements

```
pip install openpyxl --break-system-packages
pip install biopython --break-system-packages   # optional, for sanity-checking .gb output
```

## Quick start

```
python3 xlsx_to_feature_table.py my_annotation.xlsx my_annotation.tbl.txt
python3 xlsx_to_genbank.py my_annotation.xlsx my_annotation.gb
```

Each run prints any warnings it found — these flag real data issues (a CDS
with no Product, coordinates outside the sequence length, a translation
that came out empty, Exon Group rows that don't match up) and are worth
reading even when the file was written successfully. A warning means "check
this row," not "the run failed."

To double-check a `.gb` file's translations independently:

```
python3 -c "
from Bio import SeqIO
r = SeqIO.read('my_annotation.gb', 'genbank')
for f in r.features:
    if f.type == 'CDS':
        recomputed = f.extract(r.seq).translate(table=11, to_stop=True, cds=True)
        declared = f.qualifiers['translation'][0]
        print(f.location, 'OK' if str(recomputed) == declared else 'MISMATCH')
"
```

## The workbook

Three sheets, filled in by hand or generated from your own annotation
pipeline:

- **Record Info** — key/value fields: Locus/Sequence Name, Definition,
  Organism, Molecule Type, Topology, Division, Accession, Strain, Isolate,
  Collection Date, Country, Comment.
- **Sequence** — the nucleotide sequence, column A, one chunk per cell
  starting at row 2 (Excel caps a cell at ~32,767 characters — a 69 kb
  genome needs several cells).
- **Features** — one row per feature: Feature Key, Start, End, Gene #,
  Product, transl_table, Codon Start, Note, Exon Group, plus optional
  columns (Protein ID, db_xref, Other Qualifiers, Partial 5'/3' end).
  Strand isn't a separate column — Start > End means minus strand.

Full column-by-column conventions, defaults, and worked examples are in
`PROTOCOL.md`.

## Which virus types this covers

- **Non-spliced DNA/RNA viruses** (most phages and most other
  prokaryotic/archaeal viruses) — fully supported out of the box.
- **Polyprotein viruses** (coronaviruses, flaviviruses, picornaviruses) —
  supported via `mat_peptide` rows.
- **Spliced eukaryotic DNA viruses** (herpesviruses, adenoviruses,
  baculoviruses) — supported via the Exon Group column.
- **Ambisense/negative-sense segments** (arenaviruses, bunyaviruses) —
  supported with no extra steps; strand is read per-feature.
- **Segmented genomes** (influenza, bunyaviruses, reoviruses) — fill in
  and run one workbook per segment; each segment is its own GenBank record.
- **Not supported**: features that wrap around a circular genome's origin,
  multiple isoforms of the same gene in one workbook, RNA editing.

## Validation

The core (non-spliced) logic was checked feature-by-feature against a real
published 69 kb phage genome (NCBI accession PQ613263). The v3 additions
(mat_peptide, Exon Group splicing, genomic RNA) were validated with
hand-built test cases whose expected translations were computed by hand in
advance, then cross-checked by round-tripping the generated `.gb` files
through Biopython's own GenBank parser. The trickiest part — the
coordinate ordering inside a spliced minus-strand `join()` — was verified
directly against the official INSDC feature table specification's worked
examples and against NCBI's own published worked example of a spliced
tRNA gene, not just against one library's behavior. See `PROTOCOL.md` and
the comments in `xlsx_to_genbank.py` for the full reasoning and worked
examples if you need to trust or extend this further.
