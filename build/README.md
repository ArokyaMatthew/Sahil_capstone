# Build scripts for the Review-2 submission

These scripts generate everything in `../_submit/`. They are kept here so the files can be regenerated or edited later; they are not part of the submission.

| Script | Produces |
|---|---|
| `deck.py` | Review-2 presentation (GITAM Beamer "Boadilla" template, 16:9) |
| `make_poster_template.py` | A copy of the GITAM poster template (10.5 × 14 in), rebuilt from its recorded geometry |
| `poster.py` | Poster, filled into that template |
| `export_figs.py` | The five design diagrams from the deck as 300-dpi PNGs for the report |
| `report.py` | Project report, filled into the GITAM report template |
| `toc_pages.py` | Measures the page each chapter starts on, for the report's table of contents |
| `pp.py` | Shared drawing and text-measuring helpers |
| `ordercheck.py`, `docx_ordercheck.py` | Check that the XML opens cleanly in PowerPoint / Word |
| `node/icons.js` | Renders the slide icons (react-icons → PNG) |

## Working folder layout (`$W`)

```
$W/inputs/Project Report_Capstone Introduction (1).docx   # report template from the HoD's mail
$W/inputs/Poster Template.pptx                            # made by make_poster_template.py
$W/latex_zip/                                             # Review_2_Template__AIDS_.zip, unzipped
$W/icons/  $W/figs/  $W/out/                              # generated
```

## Rebuild

```bash
cd build/node && npm install && node icons.js "$W/icons" && cd ..
python3 make_poster_template.py "$W" "$W/inputs/Poster Template.pptx"
python3 deck.py "$W" "$W/out/review2.pptx"
python3 poster.py "$W" "$W/out/poster.pptx"
python3 export_figs.py "$W"
python3 report.py "$W" "$W/out/report.docx"
soffice --headless --convert-to pdf "$W/out/report.docx" --outdir "$W/out"
python3 toc_pages.py "$W/out/report.pdf" "$W/toc_pages.json"
python3 report.py "$W" "$W/out/report.docx"   # second pass writes the measured TOC page numbers
```

Python needs `python-pptx`, `python-docx`, `lxml` and `Pillow`, plus LibreOffice and poppler (`pdftoppm`, `pdftotext`). Text is measured with Cambria/Arial on Windows, or with their metric-compatible twins (Caladea, Liberation Sans) on Linux.
