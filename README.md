# Henry Russell – CV

Live at [henry.russell.ac](https://henry.russell.ac).

The web page **is** the CV: it is laid out as A4 sheets, `CV.pdf` is that page printed by headless Chromium, and `CV.docx` is generated from the same content. The Word version uses Arial for compatibility with macOS and Windows; the website and PDF use Open Sans. On a phone the same document reflows into a single readable column.

## Edit and rebuild

This repository builds the public CV from `data/cv.json`. In the full career workspace, the master record is the sibling `cv-variants/make.py`, which generates this JSON. Make content changes there and run `rebuild.py` from the career workspace root to keep the website, PDF and Word file aligned. The research is private workspace material and is not part of this public website repository.

To rebuild just this public website from its existing JSON:

```bash
pip install playwright python-docx && playwright install chromium
python tools/build.py data/cv.json
```

This rewrites `index.html`, `CV.pdf`, `CV.docx`, `robots.txt` and `sitemap.xml`. The build prints how much space is left on each A4 page and fails if a page overflows.

## Why it is built this way

- One column, real headings, contact details in the body text and no tables or text boxes, so applicant tracking systems read it in order.
- No framework and no JavaScript; one web font. The page works with scripts disabled.
- The banner blue is `#1976D2` rather than Bulma's `#209CEE`, so white text meets WCAG AA contrast.
