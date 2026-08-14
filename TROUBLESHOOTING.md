# Troubleshooting

Common warnings, errors, and gotchas across `xlsx_to_feature_table.py`,
`xlsx_to_genbank.py`, and `codon_trna_report.py`, plus pre-submission
validation steps to run before actually uploading to NCBI. If you hit
something not covered here, check the comments in the relevant script —
every non-obvious decision is explained inline with the worked example or
bug report that motivated it.

## Reading warnings

Neither converter treats a warning as fatal — the script still writes its
output file. A warning means "a human should look at this row before
trusting the output," not "the run failed." Warnings print after the
"Wrote ..." line; if you don't see any, the run had none.

### "CDS has no Product — BankIt requires one" / "GenBank/BankIt requires one"

NCBI requires every CDS to have a `/product`. If left blank, both scripts
substitute "hypothetical protein" so the file is still syntactically valid,
but you should fill in a real product name before submitting — an
all-"hypothetical protein" genome looks unannotated even if your
coordinates and translations are perfect.

### "translation came out empty — check start/end/strand/codon_start"

The single most useful warning this tool produces. It means the Start/End/
strand/Codon Start combination produced zero amino acids before hitting a
stop codon — almost always a coordinate typo (off-by-one, wrong strand
direction, or a Start/End that lands mid-intron). In real testing, this
warning caught an actual typo in a hand-edited feature table (a gene line
whose start and end were the same number instead of the CDS's real span).
Treat it as "go check this row's coordinates," not a tool bug.

### "Exon Group '<name>' mixes rows whose Start/End imply different strands"

Every row in an Exon Group must agree on strand (all Start < End, or all
Start > End). This fires when they don't — check for a row where you
transposed Start and End, or copy-pasted the wrong pair of coordinates.

### "<field> on an Exon Group '<name>' continuation row is ignored"

Only the FIRST row of an Exon Group carries qualifiers (Product, Gene #,
Note, Other Qualifiers, and in `xlsx_to_genbank.py` also Protein ID,
db_xref, Exception, Transl Except, Translation Override). If you filled in
one of these fields on a later exon row, this warns you it was ignored —
move the value to the first row of the group.

### "Transl Except is only supported for a non-spliced (single-interval) CDS — skipped"

Transl Except only works on a CDS with no Exon Group (a single Start/End
pair). Computing which genomic interval a codon number falls in across an
intron boundary isn't implemented. If you need a readthrough/selenocysteine
exception on a spliced CDS, you'll need to compute the genomic location by
hand and add the `/transl_except` qualifier directly to the output file
(or split the CDS's Exon Group differently so the exception codon falls
in a single-interval piece, if the biology allows it).

### "codon <N> in Transl Except is out of range for this CDS — skipped"

The codon number you gave (e.g. "142" in "142:Trp") doesn't fall within the
CDS's length given its Codon Start. Double check: codon numbers count from
1 at the first codon AFTER any Codon Start offset, not from the raw
genomic Start. Also check you didn't transpose Start/End for a minus-strand
CDS — this warning is what an earlier version of `xlsx_to_feature_table.py`
spuriously showed for every minus-strand readthrough CDS due to a real bug
(now fixed) that computed a negative CDS length; if you see this on a
minus-strand row and are confident the codon number is right, make sure
you're running the current version of the script.

### "unrecognized amino acid code '<x>' in Transl Except — qualifier skipped"

Accepted forms: one-letter code (e.g. "W"), three-letter code (e.g. "Trp"),
"Sec"/"U" for selenocysteine, "Pyl"/"O" for pyrrolysine, "TERM" for a true
stop, or "OTHER". Anything else is rejected rather than guessed at — check
for a typo.

### "'<key>' doesn't take a /product qualifier per the INSDC spec — the Product column value was folded into Note instead"

Some feature keys (5'UTR, 3'UTR, repeat_region, regulatory, operon,
stem_loop, and others) don't support `/product` per the INSDC spec. If you
put something in the Product column for one of these, it's preserved (in
`/note`) rather than silently dropped, but you'll want to move
qualifier-appropriate data to Other Qualifiers instead — e.g.
`regulatory_class=promoter` for a `regulatory` feature, or
`rpt_type=long_terminal_repeat` for a `repeat_region`.

### "Exon Group intervals are not in ascending order — treating this as a circular-origin-spanning feature"

This fires when a plus-strand Exon Group's Start/End pairs aren't in
increasing genomic order — the tool assumes you meant a feature that wraps
around a circular genome's origin (like a geminivirus Rep gene) and handles
it accordingly. If that's NOT what you intended, you likely have exon rows
in the wrong order, or mixed up which row goes first — check your Start/End
values. Minus-strand origin wraps aren't detected this way (the correct
syntax there needs `join(complement(...),complement(...))`, not
implemented) — if you have one, you'll need to hand-edit the output.

### "unrecognized feature key '<key>' — skipped"

The Feature Key column has a value that isn't in the ~50-key INSDC
vocabulary (or is misspelled/miscapitalized — feature keys are
case-sensitive, e.g. "CDS" not "cds", "5'UTR" not "5'utr"). Check spelling
against https://www.insdc.org/submitting-standards/feature-table/.

### Sequence contains unexpected characters

The Sequence sheet has characters outside the IUPAC nucleotide alphabet
(ACGTUNRYSWKMBDHV). Usually a paste artifact (stray whitespace inside a
sequence chunk cell, a header row that got included, or the sequence in
protein rather than nucleotide letters). The scripts keep the characters
as-is (they don't reject the run) but flag them since they'll produce a
wrong or unusual translation.

## Known gaps (things the tool does NOT check for you)

- **mobile_element's mandatory `/mobile_element_type` qualifier** —
  reachable via Other Qualifiers, but the tool does not warn you if you
  forget it. NCBI's own validation (table2asn) will catch this at
  submission time if you miss it.
- **No cross-feature validation at all.** Neither converter checks whether
  two features make biological sense together (overlapping genes,
  a CDS entirely inside another feature, gene/CDS span mismatches beyond
  the auto-generated pairing). This is why HBV-style overlapping reading
  frames "just work" — nothing stops you from creating genuinely wrong
  overlaps too. Use table2asn / the NCBI Discrepancy Report for this class
  of check (see below).
- **RNA-edited transcripts are not derived, ever.** If you set Exception
  without also giving a Translation Override, the `.gb` output simply omits
  `/translation` for that CDS (with a warning) rather than guessing.
- **Minus-strand circular-origin wraps.** Only the plus-strand form is
  detected and handled.

## Pre-submission validation checklist

Do these roughly in order of effort, and don't skip straight to the last
one — each catches a different class of problem:

1. **Read every warning the script prints.** Free, and catches the most
   common mistakes (see above).
2. **Biopython round-trip** — confirms the file parses and CDS
   translations recompute identically from the DNA:
   ```
   python3 -c "
   from Bio import SeqIO
   for r in SeqIO.parse('output.gb', 'genbank'):
       print(r.id, len(r.seq), len(r.features))
       for f in r.features:
           if f.type == 'CDS' and 'translation' in f.qualifiers:
               if f.qualifiers.get('exception') or f.qualifiers.get('transl_except'):
                   continue
               declared = f.qualifiers['translation'][0]
               try:
                   recomputed = str(f.extract(r.seq).translate(table=11, to_stop=True, cds=True))
               except Exception as e:
                   print('  (cds=True check skipped:', e, ')')
                   continue
               print('  OK' if declared == recomputed else f'  MISMATCH: {declared} vs {recomputed}')
   "
   ```
   This does NOT check CDS's with `/transl_except` or `/exception` — verify
   those by hand (see "Worked examples" in `PROTOCOL.md`) or trust a
   supplied Translation Override. It also will not catch a
   syntactically-valid-but-wrong record, like the historical
   backwards-`join()`-order bug this tool had during development — a
   frameshifted protein is still a "valid" translation as far as Biopython
   is concerned.
3. **Hand-compute the expected output for at least one feature per feature
   type you used** (one mat_peptide boundary, one spliced join(), one
   transl_except position, the origin wrap, one segment of a multi-segment
   record) and compare to the tool's output. This is genuinely the
   strongest check available, and it's how every real bug in this tool's
   own development was caught.
4. **Run NCBI's `table2asn`, or review the Discrepancy Report it
   generates**, before actually submitting. This checks things this tool
   deliberately doesn't attempt: biological plausibility of overlaps,
   mandatory-qualifier completeness for specific feature types, and
   BankIt/WebSub's own submission-time validation rules. This tool's output
   should be "structurally correct and ready for human/NCBI review," not
   "guaranteed to pass with zero flags" — those are different bars.

## `codon_trna_report.py`-specific notes

- It requires `xlsx_to_genbank.py` to be in the same directory — it imports
  that script rather than duplicating its logic, so if you see
  `ModuleNotFoundError` or the "ERROR: this script must be run from the
  same directory" message, copy both files together.
- If a tRNA row's Product text doesn't match a recognizable pattern (e.g.
  "tRNA-Ile", "tRNA-fMet", "tRNA-SeC", "tRNA-Leu1"), it's reported as
  "UNPARSED" in the text report rather than guessed at — fix the Product
  text to a recognizable form, or check by hand.
- "Amino acids with real codon usage but NO annotated cognate tRNA" is
  informational, not necessarily an error — many phages rely on host
  tRNAs for some amino acids rather than encoding a complete set.
