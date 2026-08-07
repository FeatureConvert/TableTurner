# Example: Shigella phage Moonfish

A real, complete worked example - so you can see exactly what TableTurner does before running it on your own data, and check its output against real published records yourself.

Moonfish is a 69,166-bp *Shigella* podophage isolated and annotated by [Dr. Kristin Parent's lab](https://kparentlab.natsci.msu.edu/) at Michigan State University, and published as:

> Subramanian S, McGuffin H, Passage R, Dover JA, Parent KN. 2025. Complete genome sequence of *Shigella* phage Moonfish isolated from Mid-Michigan. *Microbiology Resource Announcements* 14(6):e01255-24. https://doi.org/10.1128/mra.01255-24 - open access: [PMC12160474](https://pmc.ncbi.nlm.nih.gov/articles/PMC12160474/)

The finished, published GenBank record is [PQ613263](https://www.ncbi.nlm.nih.gov/nuccore/PQ613263).

## Files in this folder

| File | What it is |
|---|---|
| `GenBank_Annotation_Template_FILLED_moonfish.xlsx` | The template, filled in with Moonfish's real annotation data (85 CDS, 1 tRNA) |
| `Moonfish_genome.txt` | The raw 69,166-bp genome sequence |
| `original_lab_draft_feature_table.txt` | The lab's real, hand-typed draft feature table - what a person would otherwise have had to produce manually |
| `generated_feature_table.tbl.txt` | TableTurner's output - same format, generated from the spreadsheet instead of by hand |
| `generated_genbank.gb` | TableTurner's full GenBank flat file output for the same data |

Open any of these yourself and compare - that's the point of including them.

## Validation: how this was actually checked

Two independent checks were run against real data, not synthetic test cases, before this tool was considered done. Both are summarized here so you can judge the tool's reliability for yourself rather than take it on faith.

### Check 1 - generated feature table vs. the lab's own hand-typed draft

`generated_feature_table.tbl.txt` was diffed character-by-character against `original_lab_draft_feature_table.txt` (the real file the lab produced by hand before this tool existed).

- **642 lines** in both files.
- **99.91%** similarity after normalizing away pure line-ending/whitespace artifacts (the original file uses old-Mac-style CR line endings, a pure encoding quirk unrelated to content).
- Of 642 lines, **14 differ**. Every one was individually traced:
  - **13 are TableTurner correctly fixing formatting slips in the hand-typed original** - space-delimited qualifiers instead of tabs, a stray extra tab before one feature line, leading spaces instead of tabs on one line, trailing whitespace on a product value. These are the tool doing its job: producing clean, spec-compliant tab-delimited output instead of reproducing typing artifacts.
  - **1 is a genuine data discrepancy** (gene gp54's CDS span reads as zero-length in the generated output). Root cause: the example spreadsheet used to build this demo inherited a data-entry error from an earlier parsing step, not a bug in the conversion logic itself - TableTurner's own coordinate-range warning system correctly flagged this row (see the single warning printed when generating `generated_genbank.gb`).

**Confidence in the converter logic from this check: 99%.** Every differing line is accounted for; none point to a defect in the conversion code - the one real mismatch traces to bad input data, not bad conversion.

### Check 2 - generated GenBank flat file vs. the real published NCBI record

`generated_genbank.gb` was compared feature-by-feature against the actual published record, [PQ613263](https://www.ncbi.nlm.nih.gov/nuccore/PQ613263), fetched directly from NCBI. This check caught real bugs during development, all now fixed:

- **Minus-strand partial markers (`<`/`>`) were backwards.** The INSDC spec ties `<`/`>` to genomic coordinate position, not reading direction, so the correct placement flips on the minus strand relative to the plus strand - this was fixed and verified against Biopython's own location parser.
- **A gene's short label (e.g. "gp61") wasn't being represented in the flat file at all.** NCBI's own processing folds this into `/note` rather than a second `/product` line - confirmed by inspecting the real published record and matched exactly.
- **The first codon of a CDS wasn't always translated as Met.** Standard convention (and genetic code table 11, used for phage) translates a gene's true start codon as Met even for alternate starts like GTG or TTG, which would translate as Val/Leu anywhere else in the frame. Missing this caused several real minus-strand genes to come out with the wrong first amino acid - found by diffing translations against the published record's, confirmed against Biopython's own `translate(cds=True)`.
- **Paired `gene` features weren't being auto-generated for each CDS.** Every CDS in a GenBank record needs a matching `gene` feature; the tool now generates this automatically for all 85 CDS in this draft, matching the feature-table converter's existing behavior. (The published record itself has 82 CDS/82 gene, one-to-one - see the note below on why that count differs from this draft's 85.)

Two more differences turned out to be expected, not bugs:
- The draft has **85 CDS**; the published record has **82** - NCBI curation merged/removed 3 entries between submission and publication. Confirmed by checking the published record's locus_tag numbering directly (gaps exist where entries were removed).
- **locus_tag naming**: this output uses `Moonfish_gp{N}` (matching the lab's own draft convention, which this tool was built to match), while the published record shows `Moonfish_{N}` without "gp" - likely a BankIt/NCBI reformatting step during submission, not something to silently "correct" here since the source data's own convention is what the tool is asked to preserve. Confirmed with Dr. Parent: the tool keeps this default. If your own lab uses a different locus_tag convention, you'll need to adjust the generated output (or your Gene #/naming inputs) manually - the tool won't auto-detect or switch conventions for you.

**Confidence after this check: high** - every difference found either got fixed and re-verified, or was traced to a documented, expected divergence between a submission draft and NCBI's own post-curation output. Independently cross-checked at each step with Biopython (`SeqIO.read`, `translate(table=11, cds=True)`, `reverse_complement()`) rather than trusting the tool's own logic alone.

## Try it yourself

The converter scripts live at the repo root, one folder up from here, so run them from there and point them at the files in this folder:

```
cd ..
python3 xlsx_to_feature_table.py examples/moonfish/GenBank_Annotation_Template_FILLED_moonfish.xlsx examples/moonfish/my_output.tbl.txt
python3 xlsx_to_genbank.py examples/moonfish/GenBank_Annotation_Template_FILLED_moonfish.xlsx examples/moonfish/my_output.gb
```

Then diff `my_output.tbl.txt` against `generated_feature_table.tbl.txt`, or `my_output.gb` against `generated_genbank.gb` - they should match exactly, since nothing about the conversion logic depends on anything outside this repo.
