# Annotation Protocol: GenBank Annotation Template

How to fill in the Features sheet for each virus/genome type the converter
scripts support. Read `README.md` first for the basic workbook layout and
how to run the scripts. See `TROUBLESHOOTING.md` for warning-by-warning
explanations and a pre-submission validation checklist.

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
misc_feature / gene (or any other INSDC key — see section 7), Start/End per
the strand convention above. This is the original, most-tested path.

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

Each segment is its own GenBank record. Two supported approaches:

- **Native multi-segment workbook (recommended for most labs)**: in ONE
  workbook, duplicate the Record Info/Sequence/Features sheet triplet per
  segment, suffixing each sheet name with the same label in parentheses —
  e.g. "Record Info (PB2)", "Sequence (PB2)", "Features (PB2)" for one
  segment, "Record Info (PB1)"/etc. for the next. Both converters detect
  this automatically (a plain "Record Info"/"Sequence"/"Features" workbook
  with no suffix is still treated as a single segment, unchanged) and
  produce ONE combined output: a multi-record `.gb` (multiple concatenated
  LOCUS...// blocks, directly readable with Biopython's `SeqIO.parse`) or a
  multi-record `.tbl` (repeated `>Feature` blocks) plus a companion
  multi-FASTA `.fsa` of all segments' sequences — matching how BankIt batch
  submissions pair a multi-FASTA with a multi-record feature table.
- **One workbook per segment (the original workaround, still supported)**:
  fill in and run a separate plain workbook for each segment, producing
  separate output files. Useful if your lab's process submits segments one
  at a time, or if you're maintaining per-segment files for other reasons.

Either way, name segments clearly (e.g. "PB2", "Seg1") so the resulting
records are easy to tell apart.

## 6. Wide INSDC feature-key vocabulary

Beyond gene/CDS/tRNA/rRNA/misc_feature/mat_peptide, the Feature Key column
accepts the full ~50-key INSDC vocabulary — 5'UTR, 3'UTR, repeat_region,
regulatory, operon, mobile_element, stem_loop, oriT, misc_recomb,
polyA_site, and more (full list:
https://www.insdc.org/submitting-standards/feature-table/). Two qualifiers
that come up often are reachable through the Other Qualifiers column (as
`key=value`, `|`-separated for multiple):

- `regulatory_class=promoter` (or `terminator`, `ribosome_binding_site`,
  etc.) on a `regulatory` feature.
- `rpt_type=long_terminal_repeat` (or `inverted`, `direct`, `dispersed`,
  etc.) on a `repeat_region` feature — this is the modern, correct way to
  annotate retroviral LTRs and viral inverted terminal repeats (INSDC
  retired the old standalone "LTR"/"inverted_repeat" keys in favor of
  `repeat_region` + `rpt_type`).

Only feature keys the INSDC spec confirms take a `/product` qualifier will
get one from the Product column (CDS, mat_peptide, the RNA keys, and a
handful of others like misc_feature/sig_peptide/propeptide/
transit_peptide/D_segment/N_region); for any other key, a non-blank Product
value is folded into `/note` instead, with a warning, so nothing is
silently dropped.

## 7. Stop-codon readthrough and selenocysteine incorporation (Transl Except column)

For a CDS where a specific codon reads through a stop codon (common in
alphaviruses and some plant viruses) or incorporates selenocysteine, enter
`codonNumber:aminoAcid` in the Transl Except column — e.g. `142:Trp` means
"the 142nd codon of this CDS, which would otherwise be a stop, is read as
tryptophan." Multiple exceptions on one CDS are `|`-separated (e.g.
`50:Sec|142:Trp`). Both the amino acid's one-letter (`W`) and three-letter
(`Trp`) forms are accepted, plus `Sec`/`U` for selenocysteine and `Pyl`/`O`
for pyrrolysine.

Only works on a non-spliced (single Exon Group-less) CDS row — see
`TROUBLESHOOTING.md` for the workaround if you need this on a spliced CDS.

## 8. RNA editing (Exception + Translation Override columns)

For paramyxovirus-style P/V/W gene editing (where the mRNA gains inserted
nucleotides not present in the genome) or trans-splicing, the tool cannot
derive the correct protein from genomic coordinates alone — you supply it:

- **Exception** column: enter the INSDC exception text, e.g. `RNA editing`
  or `trans-splicing`.
- **Translation Override** column: enter the real protein sequence
  (one-letter codes) directly. Used verbatim; no translation is computed.

If you set Exception without a Translation Override, the `.gb` output
omits `/translation` for that CDS entirely (with a warning) rather than
computing a wrong one from the unedited genomic sequence.

## 9. Circular-origin-spanning features, plus strand only

For a feature (e.g. a geminivirus Rep gene) whose location wraps from near
the end of a circular genome back to position 1 — the INSDC spec's own
example is `join(2004..2195,3..20)` — use the Exon Group column exactly as
you would for ordinary splicing, but with the tail-end exon first and the
head-end exon second:

| Feature Key | Start | End | Gene # | Product | Exon Group |
|---|---|---|---|---|---|
| CDS | 2004 | 2195 | 1 | Rep protein | wrap1 |
| CDS | 3 | 20 | | | wrap1 |

The tool detects that these intervals aren't in ascending order for a
plus-strand feature and treats it as an origin wrap rather than an error,
giving the auto-generated gene feature the same multi-interval location as
the CDS. **Minus-strand origin wraps are not supported** (the correct
syntax needs `join(complement(...),complement(...))`).

## 10. Overlapping reading frames (HBV-style)

No special steps needed — just enter each overlapping CDS as its own row
with its own Start/End/Gene #/Product. Nothing in either converter checks
one feature's coordinates against another's, so overlapping (even
frame-shifted) CDS's on the same strand annotate correctly with zero extra
configuration. (Confirmed by direct testing, not just assumption — see
"Worked examples" below.)

## 11. Anti-CRISPR (acr/aca) operon annotation

Anti-CRISPR (acr) genes are notoriously hard to identify from sequence
alone — they're small, fast-evolving, and share no common motif. The
standard field workaround ("guilt by association," Pawluk et al. 2016) is:
anti-CRISPR-associated (aca) genes ARE recognizable, because they carry a
helix-turn-helix domain that autoregulates the acr-aca operon, and acr
genes are typically found immediately adjacent to an aca gene in the same
operon. To annotate this relationship explicitly:

- One `operon` feature spanning both the acr and aca genes, with a `Note`
  explaining the guilt-by-association rationale and
  `operon=<name>` in Other Qualifiers.
- One CDS row each for the acr and aca genes, annotated normally.

No code changes were needed for this — `operon` is standard INSDC
vocabulary and Other Qualifiers already supports arbitrary key=value pairs.
See "Worked examples" below for a full test case.

## 12. Codon usage / tRNA coverage analysis (not a submission format)

For phage genomics work relating codon usage bias to tRNA gene content
(e.g. Sarah Doore's lab's published research,
https://pmc.ncbi.nlm.nih.gov/articles/PMC13015707/), run
`codon_trna_report.py` against the same annotated workbook instead of (or
in addition to) the two submission converters. It computes codon usage
across every CDS (using the exact same splicing/strand-aware coding
sequence assembly as `xlsx_to_genbank.py`) and cross-references it against
tRNA gene Products to flag amino acids used in real codons but missing an
annotated cognate tRNA. See "Worked examples" below for expected output on
the standard regression test workbook.

## 13. Not supported

- **Minus-strand circular-origin-spanning features** — the correct INSDC
  syntax needs `join(complement(...),complement(...))`, not implemented.
- **Multiple isoforms of the same gene** (alternative splicing producing
  more than one mature transcript) in a single workbook.
- **Algorithmic derivation of RNA-edited transcripts** — you must supply
  the real protein via Translation Override; the tool won't guess.
- **Transl Except on a spliced (Exon Group) CDS** — mapping a codon number
  across an intron boundary isn't implemented.

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

## Worked examples

Every feature type above has a corresponding hand-built test workbook with
a hand-computed expected output, exercised during development (not just
"it ran without crashing"). These are the actual test cases used to find
and fix real bugs in this tool — reproduce them yourself if you want to
verify a copy of the scripts before trusting it with real data.

| Test workbook | Feature exercised | Expected result |
|---|---|---|
| `test1_regression.xlsx` | Baseline non-spliced CDS + tRNA | 2 CDS translate correctly; 1 tRNA-Ile annotated |
| `test2_spliced_plus.xlsx` | Plus-strand 2-exon join() | Correct spliced protein, ascending join() |
| `test3_spliced_minus.xlsx` | Minus-strand 2-exon join() | `complement(join(...))` in ascending genomic order despite descending entry order |
| `test4_polyprotein.xlsx` | mat_peptide cleavage products | One CDS + 2 mat_peptides, independent, no locus_tag on mat_peptides |
| `test5_spliced_partial.xlsx` | Partial (`<`/`>`) markers on a spliced feature | Markers land on the true biological outer ends, not every exon |
| `test6_three_exon_minus.xlsx` | 3-exon minus-strand join() | Correct 3-part `complement(join(...))`, correct concatenated protein |
| `test7_warnings.xlsx` | Deliberately bad data | Confirms every warning type fires when it should |
| `testA_wide_keys.xlsx` | 5'UTR / CDS / repeat_region / regulatory / 3'UTR | repeat_region's Product folds into /note (with warning); regulatory_class/rpt_type via Other Qualifiers |
| `testB_readthrough_plus.xlsx` / `testB_readthrough_minus.xlsx` | Transl Except stop-codon readthrough, CDS `ATGAAATGACCCTAA` | Both strands translate to `MKWP` (codon 3, normally a stop, reads as Trp) |
| `testC_rna_editing.xlsx` | Exception with/without Translation Override | No-override row: `/translation` omitted with a warning. Override row: protein used verbatim |
| `testD_multisegment.xlsx` | Native 2-segment workbook | One combined output; segment 1 → `MKPGF`, segment 2 → `MDYK` |
| `testE_origin_wrap.xlsx` | Plus-strand circular-origin wrap, 30 bp genome, CDS wraps 25-30 → 1-9 | Protein `MKPD`; gene feature gets the same join() as the CDS |
| `testF_overlap.xlsx` | HBV-style overlapping CDS's, no Exon Group | Both CDS's annotate independently: `MKPGF` and (frame-shifted) `NPGLKG` |
| `testG_acr_aca_operon.xlsx` | Karen Maxwell-style acr/aca operon | `operon` feature spans both genes with guilt-by-association note; acr CDS → `MKP`, aca CDS → `MDYK` — verified end to end below |

### Worked example: acr/aca operon (testG)

A synthetic 33 bp two-gene phage operon: an acr-like CDS (`ATGAAACCCTAG`,
positions 3-14, translating to `MKP`) immediately upstream of an aca-like
CDS (`ATGGATTACAAATAG`, positions 17-31, translating to `MDYK`), wrapped in
one `operon` feature spanning positions 3-31 with a Note describing the
guilt-by-association rationale and `operon=acrX1-aca1 operon` in Other
Qualifiers. Running both converters against this workbook produces:

```
operon          3..31
                /note="Predicted acr-aca operon, identified by
                guilt-by-association: acrX1 lacks a recognizable domain
                but sits immediately upstream of aca1, an HTH-domain
                autoregulator typical of anti-CRISPR operons (Pawluk et
                al. 2016 guilt-by-association strategy)."
                /operon="acrX1-aca1 operon"
gene            3..14
                /locus_tag="TestAcrAca_gp1"
CDS             3..14
                /locus_tag="TestAcrAca_gp1"
                ...
                /product="putative anti-CRISPR protein AcrX1"
                /translation="MKP"
gene            17..31
                /locus_tag="TestAcrAca_gp2"
CDS             17..31
                /locus_tag="TestAcrAca_gp2"
                ...
                /product="anti-CRISPR-associated protein Aca1"
                /translation="MDYK"
```

Both translations match the hand-computed expected proteins exactly, and
the `operon` feature carries through correctly to both the `.gb` and `.tbl`
outputs (as `/operon=` in the flat file, and as a plain `operon` qualifier
line in the feature table).

### Worked example: codon usage / tRNA report

Running `codon_trna_report.py` against `test1_regression.xlsx` (a 70 bp
synthetic genome with 2 CDS's and 1 tRNA-Ile gene) produces:

```
Codon usage / tRNA coverage report for test1_regression.xlsx
============================================================
Total genome length (all segments): 70 bp
Total CDS features: 2
Total codons counted: 11
Total tRNA genes annotated: 1

Per-segment summary:
  TestPhage: 70 bp, GC 57.1%, 2 CDS, 1 tRNA

Annotated tRNA genes (parsed amino acid identity):
  Row 5: "tRNA-Ile" -> Ile

Amino acids with real codon usage but NO annotated cognate tRNA gene: Ala, Asp, Phe, Gly, Lys, Met, Pro, Trp
(This may simply mean the corresponding host tRNA is used in trans — many
phages don't encode a full tRNA set — not necessarily a data-entry problem.)
```

The accompanying CSV lists all 64 codons with counts, per-mille frequency,
and a `has_annotated_trna` yes/no column. Running the same script against
`testD_multisegment.xlsx` correctly sums codon usage across both segments
(41 bp total, 2 CDS, matching the individual per-segment expected proteins
`MKPGF` and `MDYK`), confirming the multi-segment path works for analysis,
not just for submission-format conversion.
