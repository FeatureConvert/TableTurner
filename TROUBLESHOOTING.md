# Troubleshooting

**Note: this is beta software (v1.1.0-beta), not yet widely tested.** Validated against one real published genome so far, not across multiple labs or datasets - if you hit something not covered here, it may genuinely be new. See the "Still stuck?" section below.

Common errors, warnings, and gotchas - and what to do about each.

## Setup errors

**`ModuleNotFoundError: No module named 'openpyxl'`**
Install the one dependency the scripts need: `pip install openpyxl` (or `pip install openpyxl --break-system-packages` if your Python is externally managed).

**`ERROR: Workbook has no 'Record Info' sheet.` / `...no 'Features' sheet.`**
The scripts look for sheets named exactly `Record Info`, `Sequence`, and `Features`. This usually means you're pointing the script at a workbook that isn't a copy of `GenBank_Annotation_Template.xlsx`, or a sheet got renamed/deleted. Start from a fresh copy of the template if you're not sure.

**`ERROR: refusing to overwrite the input file (...)`**
You ran a script without an output filename and the auto-generated output name matched the input name (or you passed the same name for both). Give it an explicit output name, e.g.:
```
python3 xlsx_to_feature_table.py your_file.xlsx output.tbl.txt
```

## Warnings the scripts print (these are informational, not fatal)

**`translation came out empty - check start/end/strand/codon_start`**
The Start/End coordinates for that CDS row produce a zero-length or nonsensical span once strand is applied. Almost always a typo in the Start or End cell for that row - go check it. (This is exactly the kind of issue the warning system exists to catch; see the Moonfish example's validation report for a real instance.)

**`CDS has no Product - BankIt requires one`**
Fill in the Product column for that row before submitting - NCBI's BankIt will reject a CDS with no product.

**`... has no Gene # - gene feature written without a locus_tag`**
That CDS/gene row has no Gene # filled in, so the auto-generated paired `gene` feature has no `locus_tag`. Add a Gene # if this feature should have one, or ignore the warning if that's intentional.

**`location X..Y is out of range for a N-bp sequence`**
The Start/End for that row falls outside 1..N, where N is the total length of whatever you entered on the Sequence sheet. Usually means either a coordinate typo or the Sequence sheet is missing some chunks (see below).

**`Sequence contains unexpected characters: [...]`**
Something other than standard IUPAC nucleotide codes showed up in the Sequence sheet - could be a stray character from copy-pasting out of another tool. Worth a look before trusting the coordinates.

**`qualifier '...' contained a tab/newline character - replaced with a space`**
A cell (usually Note or Product) had a literal tab or line break in it, which would corrupt the tab-delimited feature table format. The script already fixed it for you; just an FYI.

## Data-entry gotchas

**Long sequences split across multiple cells**
Excel caps a single cell at ~32,767 characters, which real genomes (like Moonfish's 69,166 bp) exceed. Split the sequence across multiple cells in column A of the Sequence sheet, one chunk per cell, in order, starting at row 2 - the scripts concatenate them automatically. Don't try to cram the whole thing into one cell.

**Strand is inferred from Start vs. End, not a separate column**
If Start < End, the feature is `+` strand. If Start > End (the bigger number goes in the Start column), it's `-` strand. There's no separate Strand column to fill in - entering the numbers "backwards" for a minus-strand gene is correct, not a mistake.

**Don't add a separate `gene` row for every CDS**
Both scripts auto-generate the paired `gene` feature for each CDS row. Adding your own `gene` row for the same gene will produce a duplicate.

**Alternate start codons (GTG, TTG, etc.) are only recognized for `transl_table` 1 and 11**
The template defaults to `transl_table=11` (bacteria/archaea/phage), which recognizes TTG, CTG, ATT, ATC, ATA, ATG, and GTG as valid alternate start codons that translate as Met. `transl_table=1` (standard code) only recognizes ATG. If you're using a different genetic code table and a gene's translation looks wrong at the first residue, this is likely why - open a GitHub issue if you need another table's start codons supported.

**locus_tag naming (`Name_gp1` vs `Name_1`) doesn't match the published NCBI record**
This is expected, not a bug - see `examples/moonfish/README.md` for a full explanation. NCBI's own curation sometimes reformats locus_tags during submission processing; the tool preserves whatever convention you're already using in your source data rather than silently changing it. This is a confirmed design decision (the `Name_gp{N}` default matches this lab's own convention), not something the tool will auto-detect or switch for you - if your lab uses a different locus_tag convention, you'll need to adjust the generated output manually (or your source data's Gene # / naming) to match.

**CDS count in your draft doesn't match the eventually-published record**
Also expected - NCBI curation can merge or remove features between submission and publication (documented with a real example in `examples/moonfish/README.md`).

## Still stuck?

- **Tool bugs / script errors / anything that looks like a defect in the conversion logic:** open an issue at [github.com/FeatureConvert/TableTurner/issues](https://github.com/FeatureConvert/TableTurner/issues).
- **"Is this the right way to annotate this feature" / other biology judgment calls:** that's a GenBank/NCBI question, not the tool - see the README's "Questions / support" section for where to look.
- **Every time you use the tool:** verify the output yourself before submitting - no second reviewer is required, but the person running the tool is responsible for checking it. See "Verifying the output" in `GenBank_Protocol.docx`.
