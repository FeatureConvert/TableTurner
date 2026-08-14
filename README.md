# GenBank Annotation Template → Feature Table / GenBank Flat File

Scripts that convert a filled-in "GenBank Annotation Template" Excel
workbook into NCBI submission-ready output, plus a companion analysis
script for phage genomics.

| Script | Produces | Use for |
|---|---|---|
| `xlsx_to_feature_table.py` | NCBI's 5-column tab-delimited feature table (`.tbl.txt`), plus a companion multi-FASTA (`.fsa`) for multi-segment genomes | Uploading via BankIt / WebSub |
| `xlsx_to_genbank.py` | A full GenBank flat file (`.gb`) | Record-keeping, or any non-BankIt submission route |
| `codon_trna_report.py` | A codon-usage CSV + tRNA-coverage text report (NOT a submission format) | Comparative phage genomics (codon bias / tRNA content studies) |

The two converters read the same workbook, so you only fill in the
spreadsheet once and can generate either or both submission outputs from
it. `codon_trna_report.py` also reads that same workbook (it imports and
reuses `xlsx_to_genbank.py`'s logic) but produces an analysis report
instead of a submission file — see "Codon usage / tRNA coverage analysis"
below.

Originally built for non-spliced dsDNA phage genomes, the converters now
also support the full INSDC feature key vocabulary, polyprotein viruses,
spliced eukaryotic viral genes, RNA virus genomes, RNA editing and
stop-codon readthrough, native multi-segment genomes, and plus-strand
circular-origin-spanning features. See `PROTOCOL.md` for the annotation
conventions to follow for each virus type and worked examples, and "Which
virus types this covers" below for a quick summary. If something doesn't
work the way you expect, check `TROUBLESHOOTING.md` first.

## Requirements

```
pip install openpyxl --break-system-packages
pip install biopython --break-system-packages   # optional, for sanity-checking .gb output
```

## Quick start

```
python3 xlsx_to_feature_table.py my_annotation.xlsx my_annotation.tbl.txt
python3 xlsx_to_genbank.py my_annotation.xlsx my_annotation.gb
python3 codon_trna_report.py my_annotation.xlsx my_annotation   # optional, analysis only
```

Each converter run prints any warnings it found — these flag real data
issues (a CDS with no Product, coordinates outside the sequence length, a
translation that came out empty, Exon Group rows that don't match up, a
feature key that doesn't support `/product` folding its value into `/note`
instead, a Transl Except used on an unsupported spliced CDS) and are worth
reading even when the file was written successfully. A warning means "check
this row," not "the run failed." See `TROUBLESHOOTING.md` for what each
warning means and how to fix it.

To double-check a `.gb` file's translations independently:

```
python3 -c "
from Bio import SeqIO
for r in SeqIO.parse('my_annotation.gb', 'genbank'):
    for f in r.features:
        if f.type == 'CDS' and 'translation' in f.qualifiers:
            if f.qualifiers.get('exception') or f.qualifiers.get('transl_except'):
                continue  # verify these by hand instead — see PROTOCOL.md
            recomputed = f.extract(r.seq).translate(table=11, to_stop=True, cds=True)
            declared = f.qualifiers['translation'][0]
            print(f.location, 'OK' if str(recomputed) == declared else 'MISMATCH')
"
```

This loops over every record in the file (handles multi-segment `.gb`
output too, via `SeqIO.parse` instead of `SeqIO.read`). See "Pre-submission
validation" in `PROTOCOL.md`/`TROUBLESHOOTING.md` for the fuller checklist,
including what this check does NOT catch.

## The workbook

Three sheets, filled in by hand or generated from your own annotation
pipeline:

- **Record Info** — key/value fields: Locus/Sequence Name, Definition,
  Organism, Molecule Type (genomic DNA / genomic RNA / mRNA / etc.),
  Topology, Division, Accession, Strain, Isolate, Collection Date, Country,
  Comment.
- **Sequence** — the nucleotide sequence, column A, one chunk per cell
  starting at row 2 (Excel caps a cell at ~32,767 characters — a 69 kb
  genome needs several cells).
- **Features** — one row per feature: Feature Key (any of the ~50 INSDC
  keys), Start, End, Gene #, Product, transl_table, Codon Start, Note, Exon
  Group, Exception, Transl Except, Translation Override, plus optional
  gray-header columns (Protein ID, db_xref, Other Qualifiers, Partial
  5'/3' end). Strand isn't a separate column — Start > End means minus
  strand.

For a genome with multiple segments (influenza-style), duplicate the three
sheets per segment, suffixing each sheet name with a label in parentheses —
e.g. "Record Info (PB2)", "Sequence (PB2)", "Features (PB2)" — all in one
workbook. Both converters detect this automatically and produce one
combined multi-record output.

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
- **Segmented genomes** (influenza, bunyaviruses, reoviruses) — supported
  natively in one workbook (see above), or you can still run one workbook
  per segment if you prefer.
- **Stop-codon readthrough / selenocysteine incorporation** (alphaviruses
  and similar) — supported via the Transl Except column.
- **RNA editing** (paramyxovirus P/V/W genes) — supported via the Exception
  + Translation Override columns (you supply the real protein; the tool
  can't derive an edited transcript from genomic coordinates alone).
- **The full INSDC feature key vocabulary** (~50 keys: 5'UTR/3'UTR,
  repeat_region, regulatory, operon, mobile_element, stem_loop, and more) —
  supported, with keys that don't take `/product` per spec folding it into
  `/note` instead.
- **Overlapping reading frames** (HBV-style) — supported with no extra
  steps.
- **Circular-origin-spanning features, plus strand only** (e.g. a
  geminivirus Rep gene) — supported via Exon Group.
- **Not supported**: minus-strand circular-origin-spanning features,
  multiple isoforms of the same gene in one workbook, and (as always) the
  tool cannot algorithmically derive an RNA-edited transcript — you supply
  it via Translation Override.

## Codon usage / tRNA coverage analysis

`codon_trna_report.py` is a separate analysis tool, not a submission
converter, built for phage genomics work like Sarah Doore's lab's published
research on phage codon usage bias, infection style, and tRNA gene content
(https://pmc.ncbi.nlm.nih.gov/articles/PMC13015707/). Given the same
annotated workbook, it computes codon usage across every CDS and
cross-references it against the amino-acid identities of annotated tRNA
genes (parsed from their Product text, e.g. "tRNA-Ile"), flagging any amino
acid used in real codons but missing a cognate tRNA — a genuine biological
signal (many phages rely on host tRNAs) rather than necessarily an error.

```
python3 codon_trna_report.py my_annotation.xlsx my_annotation
```

Requires `xlsx_to_genbank.py` in the same directory (it imports that
script's workbook-reading and splicing/strand logic rather than duplicating
it). Produces `<prefix>_codon_usage.csv` (per-codon counts, per-mille
frequency, amino acid, tRNA coverage flag, and a blank column for you to
fill in infection style when building a multi-genome comparative dataset)
and `<prefix>_trna_report.txt` (a human-readable summary). See
`PROTOCOL.md` for a worked example with expected output.

## Anti-CRISPR (acr/aca) operon annotation

For labs doing anti-CRISPR gene discovery (e.g. Karen Maxwell's lab's
guilt-by-association approach, where an aca gene's recognizable
helix-turn-helix domain flags adjacent acr genes that have no detectable
domain of their own), annotate the acr-aca gene pair as an `operon` feature
spanning both CDSs, with the guilt-by-association rationale in Note and
`operon=<name>` in Other Qualifiers. This is standard INSDC vocabulary —
no code changes were needed to support it. See the worked example in
`PROTOCOL.md`.

## Validation

The core (non-spliced) logic was checked feature-by-feature against a real
published 69 kb phage genome (NCBI accession PQ613263). The v3 additions
(mat_peptide, Exon Group splicing, genomic RNA) and v4 additions (wide
feature-key vocabulary, RNA editing/transl_except, multi-segment workbooks,
circular-origin wrap, overlapping reading frames) were validated with
hand-built test cases whose expected translations were computed by hand in
advance, then cross-checked by round-tripping the generated `.gb` files
through Biopython's own GenBank parser. The trickiest part — the
coordinate ordering inside a spliced minus-strand `join()` — was verified
directly against the official INSDC feature table specification's worked
examples and against NCBI's own published worked example of a spliced
tRNA gene, not just against one library's behavior. See `PROTOCOL.md` for
the full worked-example test suite, `TROUBLESHOOTING.md` for common issues
and pre-submission validation steps, and the comments in
`xlsx_to_genbank.py`/`xlsx_to_feature_table.py` for the full reasoning if
you need to trust or extend this further.
