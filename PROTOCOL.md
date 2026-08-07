# Annotation Protocol: GenBank Annotation Template

How to fill in the Features sheet for each virus/genome type the converter
scripts support. Read `README.md` first for the basic workbook layout and
how to run the scripts.

## Baseline convention (applies to every virus type)

- **Strand** is never a separate column. Enter Start < End for the plus
  strand, Start > End (the larger number first) for the minus strand.
- **CDS rows don't need a paired gene row** — both scripts write it for
  you, using Gene # to build the locus_tag (`{Locus Name}_gp{N}`).
- **transl_table**: leave blank for bacteria/archaea/phage (defaults to
  11). Use `1` for eukaryotic DNA viruses (herpesviruses, adenoviruses,
  baculoviruses), which use the standard genetic code.
- **Molecule Type** (Record Info sheet): `genomic DNA` for DNA
  genomes, `genomic RNA` for RNA virus genomes (positive-sense,
  negative-sense, or dsRNA all use this same value — the distinction lives
  in how you annotate strand per-feature, not in Molecule Type).
- **Division** (Record Info sheet): use `VRL` for viral records instead of
  the default `UNA`.

## 1. Non-spliced genomes (phages, most other prokaryotic/archaeal viruses)

No special steps. One row per feature, Feature Key = CDS / tRNA / rRNA /
misc_feature / gene, Start/End per the strand convention above. This is the
original, most-tested path.

## 2. Polyprotein viruses (coronaviruses, flaviviruses, picornaviruses)

These genomes translate as one long CDS that's cleaved post-translationally
into multiple mature peptides. Enter:

- **One CDS row** spanning the entire polyprotein open reading frame.
  Product = "polyprotein" (or the specific polyprotein name).
- **One `mat_peptide` row per mature peptide**, each with its own
  Start/End (a sub-range within the CDS's span) and Product (the peptide's
  name, e.g. "capsid protein", "RNA-dependent RNA polymerase", "3C-like
  proteinase").

Leave Gene # blank on mat_peptide rows. They do not get a locus_tag or a
paired gene feature — this matches how real published polyprotein records
are annotated (the gene feature, if any, covers the whole polyprotein CDS;
individual mat_peptides don't repeat it).

Example (simplified, positions illustrative):

| Feature Key | Start | End | Gene # | Product |
|---|---|---|---|---|
| CDS | 100 | 900 | 1 | polyprotein |
| mat_peptide | 100 | 400 | | capsid protein |
| mat_peptide | 401 | 900 | | RNA-dependent RNA polymerase |

## 3. Spliced genes (herpesviruses, adenoviruses, baculoviruses)

For a gene split across introns, give every exon row the same Feature Key
(usually CDS) and the same **Exon Group** label (any short text you choose,
e.g. "g1" — just needs to be unique per gene, reused across that gene's
exon rows). Only the *first* exon row of the group needs Product/Gene
#/Note/etc.; later rows only need Start and End — anything else you put on
a continuation row is ignored (and the script will warn you about it, so
you'll know if you did this by mistake).

**Enter exon rows in transcription (5'→3') order:**

- **Plus-strand gene**: ascending genomic order — the intuitive order,
  first row = the exon closest to the start of the sequence.
- **Minus-strand gene**: *descending* genomic order — the first row is the
  exon at the *higher* coordinates. This is the same convention as a single
  minus-strand row's own Start > End (higher number first), just extended
  across multiple rows.

You do not need to think about how this becomes a `join()` in the final
flat file — both scripts handle that conversion internally. (If you're
curious why the two formats actually store multi-exon order differently
internally, see the "Why entry order differs from file order" section
below — you don't need to know this to use the template correctly.)

Example, minus-strand gene with two exons (Exon Group "g1"):

| Feature Key | Start | End | Gene # | Product | Exon Group |
|---|---|---|---|---|---|
| CDS | 5163 | 4918 | 3 | envelope glycoprotein | g1 |
| CDS | 4571 | 2691 | | | g1 |

This produces `CDS complement(join(2691..4571,4918..5163))` in the flat
file — INSDC always writes join() arguments in ascending genomic order
regardless of strand, even though you entered them in descending
(transcription) order.

Retroviral ribosomal frameshifts (e.g. gag-pol) can be modeled the same
way: an Exon Group CDS spanning the two reading-frame segments, plus a
`ribosomal_slippage` entry (no `=value`) in the Other Qualifiers column of
the first row.

## 4. Ambisense / negative-sense segments (arenaviruses, bunyaviruses)

No special steps needed. A single segment can carry genes in both
orientations — just enter each CDS row with the Start/End order matching
its own strand. Strand is inferred per-feature, not per-record.

## 5. Segmented genomes (influenza, bunyaviruses, reoviruses, rotaviruses)

Each segment is its own GenBank record. **Fill in and run one workbook per
segment** — this isn't a limitation of the schema so much as how these are
actually submitted (BankIt itself takes segments as separate uploads, not
one combined file). Name each workbook/output file for its segment (e.g.
`Segment1_PB2.xlsx`, `Segment2_PB1.xlsx`, ...).

## 6. Not supported

- **Circular-origin-spanning features** — a feature whose location wraps
  from the end of a circular sequence back to the beginning (e.g.
  `join(2004..2195,3..20)`). This is a distinct edge case from ordinary
  splicing.
- **Multiple isoforms of the same gene** (alternative splicing producing
  more than one mature transcript) in a single workbook.
- **RNA editing** (e.g. paramyxovirus P gene editing).

If you hit one of these, the workaround is to build the flat file's
FEATURES block for that one feature by hand and splice it into the
converter's output, or ask for the scripts to be extended.

## Why entry order differs from file order (background, not required reading)

The tab-delimited feature table and the GenBank flat file actually encode a
spliced minus-strand feature's exon order *differently*, which is why the
workbook uses one consistent "transcription order" input convention and
each script does its own internal translation to the right output order:

- The **tab-delimited feature table** lists a multi-exon feature's
  continuation lines in transcription order. NCBI's own worked example (a
  spliced tRNA-Phe at
  https://www.ncbi.nlm.nih.gov/genbank/feature_table/) shows this
  directly: the table lines read "4626 4590" then "4570 4535" — descending
  genomic order for a minus-strand feature.
- The **GenBank flat file's `join()` operator** is always written in
  ascending genomic order, regardless of strand, per the official INSDC
  feature table specification's own worked examples
  (`complement(join(2691..4571,4918..5163))`,
  `complement(join(21..349,567..795))` — lower range first in both). The
  same NCBI tRNA example converts to
  `complement(join(4535..4570,4590..4626))` in the resulting flat file —
  the reverse order from the tab table.

Both scripts take the same "transcription order" input so the workbook
convention doesn't depend on which output format you're generating, and
`xlsx_to_genbank.py` reverses a minus-strand group's order internally
before building the `join()` string. This distinction is easy to get
backwards (an earlier draft of the flat-file converter did), and a wrong
order silently produces a real-looking but incorrect protein translation
rather than an error — so if you're ever hand-editing a spliced feature's
location directly, double check it against these worked examples rather
than guessing.
