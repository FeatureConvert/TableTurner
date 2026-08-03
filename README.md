# TableTurner

Turn a filled-in Excel spreadsheet into a submission-ready GenBank record — no hand-typing a feature table, no wrestling with column spacing, no re-learning the GenBank flat file spec every time you annotate a genome.

## What it does

Gene/CDS/tRNA annotation for a sequenced genome usually starts the same way: ORF calling, BLAST/InterPro predictions, tRNA calls — all tracked in a spreadsheet. Turning that spreadsheet into something NCBI will actually accept, though, has traditionally meant hand-typing a 5-column feature table or a full GenBank flat file, one qualifier line at a time. That's slow, and it's exactly the kind of tedious, mechanical work that introduces typos into otherwise-solid biology.

TableTurner does that conversion for you. You fill in three sheets of a template with the same information you're already tracking — nothing new to learn, no coordinate math — and it produces either:

- **A BankIt/WebSub feature table** — the exact 5-column tab-delimited file NCBI expects at submission, or
- **A complete GenBank flat file (.gb)** — LOCUS/DEFINITION/FEATURES/ORIGIN, for record-keeping, sharing with collaborators, or loading into Biopython/SnapGene.

The biochemistry — coordinates, strand, reading frame, translation, start codons — is handled the way GenBank itself handles it, not approximated. Validated feature-by-feature against a real published NCBI record, including catching cases where the first codon of a gene is an alternate start codon (GTG/TTG) rather than ATG.

## Quick start

1. Copy `GenBank_Annotation_Template.xlsx` and fill in the three sheets:
   - **Record Info** — organism, molecule type, topology, and the rest of the record-level metadata.
   - **Sequence** — the nucleotide sequence (split across multiple cells if it's a long genome — Excel caps a single cell around 32,767 characters).
   - **Features** — one row per gene/CDS/tRNA/rRNA/misc_feature, in genome order. Strand is inferred from whether Start or End is larger; you don't need to add a separate gene row for each CDS, that's generated automatically.
2. Run whichever converter you need:

   ```
   python3 xlsx_to_feature_table.py your_filled_template.xlsx output.tbl.txt
   python3 xlsx_to_genbank.py your_filled_template.xlsx output.gb
   ```

   (Needs `openpyxl`: `pip install openpyxl`.)
3. Read the warnings the script prints — they flag real issues (missing product on a CDS, coordinates out of range, a translation that came out empty because of a typo'd gene span) rather than being cosmetic noise.
4. Upload the feature table to BankIt/WebSub, or keep the .gb file for your own records.

See `GenBank_Protocol_DRAFT.docx` for the full step-by-step lab protocol, including how to run this through Claude if you'd rather not touch the command line at all.

## Files

| File | Purpose |
|---|---|
| `GenBank_Annotation_Template.xlsx` | Blank template — start here |
| `xlsx_to_feature_table.py` | Converts a filled template to NCBI's BankIt feature table format |
| `xlsx_to_genbank.py` | Converts a filled template to a full GenBank flat file (.gb) |
| `GenBank_Protocol_DRAFT.docx` | Step-by-step lab protocol, from filling out the template through submission |

## Known limitations

- One sequence record per workbook — for multi-segment genomes, use a separate workbook per segment.
- One contiguous span per feature row — no spliced/multi-exon (`join()`) locations.
- Defaults assume `transl_table=11` and `codon_start=1` (bacteria/archaea/phage) — override per row if a feature genuinely needs something else.

## Why "TableTurner"

It turns a spreadsheet *table* into a GenBank feature *table* — and turns the tables on the hours of manual formatting that used to take.

## How to cite

TableTurner is free to use under the MIT license (see below) — no legal obligation to cite it. But if it saved you time on an annotation that made it into a publication, an acknowledgment or citation is genuinely appreciated and helps justify continued work on tools like this. Suggested citation:

```
Houston, R. TableTurner: Excel-to-GenBank annotation conversion. https://github.com/FeatureConvert/TableTurner
```

See `CITATION.cff` for a machine-readable version — GitHub's "Cite this repository" button (top right of the repo page) will generate APA/BibTeX formats from it automatically.

## License

MIT — see `LICENSE`. Free to use, modify, and redistribute, including commercially; just keep the copyright notice.
