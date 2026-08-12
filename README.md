# TableTurner

**Current release: v1.1.0 (Beta)** - not yet widely tested. Validated end-to-end against one real published genome so far (see Example below); hasn't been run across multiple labs, datasets, or annotation styles yet. Verify your own output carefully, especially early on - see "Verifying the output" in the protocol doc.

Turn a filled-in Excel spreadsheet into a submission-ready GenBank record - no hand-typing a feature table, no wrestling with column spacing, no re-learning the GenBank flat file spec every time you annotate a genome.

## What it does

Gene/CDS/tRNA annotation for a sequenced genome usually starts the same way: ORF calling, BLAST/InterPro predictions, tRNA calls - all tracked in a spreadsheet. Turning that spreadsheet into something NCBI will actually accept, though, has traditionally meant hand-typing a 5-column feature table or a full GenBank flat file, one qualifier line at a time. That's slow, and it's exactly the kind of tedious, mechanical work that introduces typos into otherwise-solid biology.

TableTurner does that conversion for you. You fill in three sheets of a template with the same information you're already tracking - nothing new to learn, no coordinate math - and it produces either:

- **A BankIt/WebSub feature table** - the exact 5-column tab-delimited file NCBI expects at submission, or
- **A complete GenBank flat file (.gb)** - LOCUS/DEFINITION/FEATURES/ORIGIN, for record-keeping, sharing with collaborators, or loading into Biopython/SnapGene.

The biochemistry - coordinates, strand, reading frame, translation, start codons - is handled the way GenBank itself handles it, not approximated. Validated feature-by-feature against a real published NCBI record, including catching cases where the first codon of a gene is an alternate start codon (GTG/TTG) rather than ATG.

## Quick start

1. Copy `GenBank_Annotation_Template.xlsx` and fill in the three sheets:
   - **Record Info** - organism, molecule type, topology, and the rest of the record-level metadata.
   - **Sequence** - the nucleotide sequence (split across multiple cells if it's a long genome - Excel caps a single cell around 32,767 characters).
   - **Features** - one row per gene/CDS/tRNA/rRNA/misc_feature, in genome order. Strand is inferred from whether Start or End is larger; you don't need to add a separate gene row for each CDS, that's generated automatically.
2. Convert it - two ways, pick whichever fits:
   - **Ask Claude (recommended).** This repo includes a bundled Claude skill (`.claude/skills/genbank-converter/`), so if you're working with Claude Code or Cowork in this repo folder, it loads automatically - just ask Claude to "convert my filled template to a BankIt feature table" (or "to a GenBank flat file"), and it'll run the right script and report back any warnings.
   - **Or run it yourself from the command line:**

     ```
     python3 xlsx_to_feature_table.py your_filled_template.xlsx output.tbl.txt
     python3 xlsx_to_genbank.py your_filled_template.xlsx output.gb
     ```

     (Needs `openpyxl`: `pip install openpyxl`.)
3. Read the warnings the script prints - they flag real issues (missing product on a CDS, coordinates out of range, a translation that came out empty because of a typo'd gene span) rather than being cosmetic noise. See `TROUBLESHOOTING.md` if anything looks off or a warning/error doesn't make sense.
4. **Spot-check the output before you trust it - every time, not just while you're new to it.** The person running the tool is responsible for reviewing its output before submission; no second reviewer is required. Confirm the Locus/Sequence Name, organism, and a handful of CDS products/coordinates match what you expect, and see "Verifying the output" in the protocol doc for a fuller checklist.
5. Upload the feature table to BankIt/WebSub, or keep the .gb file for your own records.

See `GenBank_Protocol.docx` for the full step-by-step lab protocol.

## Files

| File | Purpose |
|---|---|
| `GenBank_Annotation_Template.xlsx` | Blank template - start here |
| `xlsx_to_feature_table.py` | Converts a filled template to NCBI's BankIt feature table format |
| `xlsx_to_genbank.py` | Converts a filled template to a full GenBank flat file (.gb) |
| `GenBank_Protocol.docx` | Step-by-step lab protocol, from filling out the template through submission |
| `.claude/skills/genbank-converter/` | The bundled Claude skill - loads automatically when using Claude Code/Cowork in this repo folder |
| `TROUBLESHOOTING.md` | Common errors, warnings, and data-entry gotchas, with fixes |

## Example

[`examples/moonfish/`](examples/moonfish/) contains a full worked example on a real, published 69,166-bp phage genome (*Shigella* phage Moonfish, [PQ613263](https://www.ncbi.nlm.nih.gov/nuccore/PQ613263)) - the filled template, TableTurner's generated output, the lab's original hand-typed feature table for comparison, and a full write-up of how the output was validated character-by-character against the original draft and feature-by-feature against the real published NCBI record, including a confidence assessment and an explanation of every discrepancy found.

## Known limitations

- **Beta software - not yet widely tested.** Validated feature-by-feature and character-by-character against one real published genome (see Example below), but hasn't been run yet across multiple labs, users, or a wider variety of genomes/annotation styles. Treat every output as something to verify, not something to trust blindly.
- One sequence record per workbook - for multi-segment genomes, use a separate workbook per segment.
- One contiguous span per feature row - no spliced/multi-exon (`join()`) locations.
- Defaults assume `transl_table=11` and `codon_start=1` (bacteria/archaea/phage) - override per row if a feature genuinely needs something else.

## Coming soon

A broader feature set was built and mechanically tested (hand-computed expected outputs, cross-checked against Biopython's parser), but is being held back from release until it's reviewed against real annotation data by someone familiar with the underlying biology - not just verified against synthetic test cases:

- Full INSDC feature-key vocabulary - 5'UTR/3'UTR, repeat_region, regulatory, operon, mobile_element, and more, beyond today's gene/CDS/tRNA/rRNA/misc_feature/source set.
- Polyprotein cleavage products (`mat_peptide`).
- Spliced eukaryotic/viral genes - `join()`/`complement(join())` locations for multi-exon features.
- RNA virus genome support.
- RNA editing and stop-codon readthrough.
- Native multi-segment genomes in a single workbook (rather than one workbook per segment).
- Circular-origin-spanning features.
- A companion codon-usage / tRNA-coverage report script.

If you need one of these for real annotation work now, open an issue at [github.com/FeatureConvert/TableTurner/issues](https://github.com/FeatureConvert/TableTurner/issues) so we can prioritize getting it properly validated.

## Why "TableTurner"

It turns a spreadsheet *table* into a GenBank feature *table* - and turns the tables on the hours of manual formatting that used to take.

## Acknowledgments

TableTurner was built to help streamline a real lab workflow: [Dr. Kristin Parent](https://kparentlab.natsci.msu.edu/), Michigan State University, provided the real-world phage genome sequencing/annotation protocol and annotation data this tool was designed around and validated against, including feature-by-feature testing against a real published NCBI record. The tool itself was written by Robert Houston (not a domain scientist), to save a lab member from doing this conversion by hand - and shared publicly in the hope it's useful to anyone else facing the same tedious conversion.

The example dataset in `examples/moonfish/` comes from the real, published genome of *Shigella* phage Moonfish:

> Subramanian S, McGuffin H, Passage R, Dover JA, Parent KN. 2025. Complete genome sequence of *Shigella* phage Moonfish isolated from Mid-Michigan. *Microbiology Resource Announcements* 14(6):e01255-24. https://doi.org/10.1128/mra.01255-24 - [full text (open access)](https://pmc.ncbi.nlm.nih.gov/articles/PMC12160474/)

## How to cite

TableTurner is free to use under the MIT license (see below) - no legal obligation to cite it. But if it saved you time on an annotation that made it into a publication, an acknowledgment or citation is genuinely appreciated and helps justify continued work on tools like this. Suggested citation:

```
Houston, R. TableTurner: Excel-to-GenBank annotation conversion. https://github.com/FeatureConvert/TableTurner
```

See `CITATION.cff` for a machine-readable version - GitHub's "Cite this repository" button (top right of the repo page) will generate APA/BibTeX formats from it automatically.

## Questions / support

- **Hitting an error or a warning you don't understand?** Check `TROUBLESHOOTING.md` first - it covers the common ones with fixes.
- **Tool bugs, feature requests, anything about how the scripts/template work:** open an issue at [github.com/FeatureConvert/TableTurner/issues](https://github.com/FeatureConvert/TableTurner/issues) - Robert Houston maintains the tool.
- **Annotation/biology questions** - is this the right way to annotate a feature, should something be marked partial, what's the correct convention for a given case - that's a GenBank/NCBI question, not the tool maintainer's. See the [NCBI Feature Table spec](https://www.ncbi.nlm.nih.gov/genbank/feature_table/) and [sample GenBank record](https://www.ncbi.nlm.nih.gov/genbank/samplerecord/), or contact NCBI directly at info@ncbi.nlm.nih.gov for genuinely ambiguous cases.

## License

MIT - see `LICENSE`. Free to use, modify, and redistribute, including commercially; just keep the copyright notice.
