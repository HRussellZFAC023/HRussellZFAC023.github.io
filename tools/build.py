#!/usr/bin/env python3
"""Build the CV web page, PDF and Word document from one data file.

    python tools/build.py data/cv.json                    -> index.html, CV.pdf, CV.docx in the repo root
    python tools/build.py some.json --out some/folder     -> the same three files (plus CV.txt) elsewhere

The web page is laid out as A4 sheets, the PDF is that page printed by
headless Chromium, and the Word document is generated from the same blocks
with the same sizes, so all three read and look the same.

Requires: pip install playwright python-docx && playwright install chromium
"""
import html
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BLUE = "1976D2"  # a shade deeper than the original Bulma blue (#209CEE), so white text passes WCAG AA (4.6:1)

# One set of sizes (points) shared by the web page and the Word document.
PT = dict(body=9.75, small=9, contact=9, name=24, title=12.5, section=14, role=10.5)


# ---------------------------------------------------------------- inline text
def inline_html(text):
    out = html.escape(text, quote=False)
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", out)


def inline_runs(text):
    """Yield (text, bold) pieces for **bold** markup."""
    for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if part:
            yield part, i % 2 == 1


def project_runs(text):
    """Keep printed project addresses while making them usable in every format."""
    pattern = r"(?:github\.com/[A-Za-z0-9_-]+/[A-Za-z0-9_.-]+|yomureader\.com|barclaycardpayments\.com|coop\.co\.uk/funeralcare|zuehlke-skill-hub\.pages\.dev)"
    start = 0
    for match in re.finditer(pattern, text):
        yield text[start:match.start()], None
        yield match.group(), "https://" + match.group()
        start = match.end()
    yield text[start:], None


# ----------------------------------------------------------------------- HTML
CSS = """
      :root {
        --blue: #%(blue)s;
        --ink: #2b2b2b;
        --grey: #5a5a5a;
        --rule: #dbdbdb;
      }
      * { box-sizing: border-box; }
      html, body { margin: 0; background-color: #f5f7fa; }
      body {
        font-family: "Open Sans", "Segoe UI", -apple-system, BlinkMacSystemFont, "Helvetica Neue", Arial, sans-serif;
        font-size: %(body)spt;
        line-height: 1.28;
        color: var(--ink);
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      a { color: inherit; text-decoration: none; }
      a:hover { text-decoration: underline; }
      a:focus-visible { outline: 2px solid var(--blue); outline-offset: 2px; border-radius: 2px; }
      .contact a { text-decoration: underline; text-decoration-color: #b5b5b5; text-underline-offset: 2px; }

      /* The CV is a stack of A4 sheets, exactly as in CV.pdf and CV.docx */
      .cv-wrapper { padding: 2.5rem 0; }
      .sheet {
        width: 210mm;
        height: 297mm;
        padding: 7mm 13mm 7mm;
        margin: 0 auto 2rem;
        background: #fff;
        border-radius: 6px;
        box-shadow: 0 2px 3px rgba(10, 10, 10, 0.1), 0 0 0 1px rgba(10, 10, 10, 0.1);
        overflow: hidden;
      }
      .sheet:last-child { margin-bottom: 0; }

      .hero {
        background: var(--blue);
        color: #fff;
        text-align: center;
        padding: 3.5mm 4mm 3.5mm;
      }
      .contact {
        font-style: normal;
        margin: 2mm 0 0;
        padding: 1mm 0;
        border-bottom: 1px solid var(--rule);
        text-align: center;
        font-size: %(contact)spt;
        color: var(--grey);
      }
      .contact p { margin: 0; }
      .contact p + p { margin-top: 0.3mm; }
      .hero h1 { margin: 0; font-size: %(name)spt; line-height: 1.15; font-weight: 700; }
      .hero p { margin: 0.6mm 0 0; font-size: %(title)spt; }

      h2 {
        margin: 2.8mm 0 1.3mm;
        text-align: center;
        font-size: %(section)spt;
        font-weight: 400;
        line-height: 1.2;
        color: #363636;
      }
      p, ul, h3, h4 { margin: 0; }
      h3, h4 { font-size: inherit; line-height: inherit; }
      .para { margin-bottom: 1.2mm; }
      .role { display: flex; justify-content: space-between; align-items: baseline; margin-top: 2mm; font-size: %(role)spt; }
      .role .dates, .org, .tech, .meta { font-style: italic; font-weight: 400; color: var(--grey); }
      .org { margin-bottom: 0.6mm; }
      .org-inline { font-style: italic; font-weight: 400; color: var(--grey); font-size: %(body)spt; }
      .project { margin: 1.3mm 0 0 5mm; font-weight: 700; }
      .project a { text-decoration: underline; text-underline-offset: 2px; }
      .tech { margin-left: 5mm; font-size: %(small)spt; }
      .project.flat { margin-left: 0; margin-top: 2mm; font-size: %(role)spt; }
      .tech.flat { margin-left: 0; }
      ul { padding-left: 9.5mm; list-style-type: "•  "; }  /* a text bullet, so it survives PDF text extraction */
      ul.flat { padding-left: 4.5mm; }
      li { margin: 0.15mm 0; }
      .skill { margin-bottom: 0.7mm; }
      .edu { margin-top: 1.6mm; }

      /* Narrow screens: the same document, set for a phone instead of shrunk to fit one */
      @media screen and (max-width: 830px) {
        html, body { background: #fff; }
        body { font-size: 15px; line-height: 1.5; }
        .cv-wrapper { padding: 0; }
        .sheet { width: auto; height: auto; margin: 0; padding: 16px; border-radius: 0; box-shadow: none; overflow: visible; }
        .sheet + .sheet { padding-top: 0; }
        .hero { padding: 22px 12px 24px; }
        .hero h1 { font-size: 28px; }
        .hero p { font-size: 16px; }
        .contact, .tech { font-size: 13px; }
        .contact p { line-height: 1.9; }
        h2 { font-size: 20px; margin: 22px 0 8px; }
        .role, .project.flat { font-size: 16px; }
        .role { flex-wrap: wrap; column-gap: 12px; margin-top: 14px; }
        .project { margin-top: 10px; }
        li { margin: 4px 0; }
        nav.downloads { position: static; justify-content: center; padding: 8px 16px 28px; font-size: 15px; }
      }

      /* Download buttons: on screen only */
      .downloads { position: fixed; right: 1rem; bottom: 1rem; display: flex; gap: 0.5rem; font-size: 13px; }
      .downloads a {
        background: var(--blue); color: #fff; padding: 0.5em 0.9em; border-radius: 4px;
        box-shadow: 0 2px 3px rgba(10, 10, 10, 0.2);
      }
      .downloads a:hover { text-decoration: none; filter: brightness(1.08); }

      @page { size: A4; margin: 0; }
      @media print {
        html, body { background: #fff; }
        .cv-wrapper { padding: 0; animation: none !important; }
        .sheet { margin: 0; border-radius: 0; box-shadow: none; break-after: page; }
        .sheet:last-child { break-after: auto; }
        .downloads { display: none; }
      }

      /* Animate in */
      .scale-up-center {
        animation: scale-up-center 0.4s cubic-bezier(0.39, 0.575, 0.565, 1) both;
      }
      @keyframes scale-up-center {
        from { transform: scale(0.85); opacity: 0; }
        to { transform: scale(1); opacity: 1; }
      }
      @media (prefers-reduced-motion: reduce) { .scale-up-center { animation: none; } }
""" % dict(PT, blue=BLUE.lower())

def block_html(b):
    t = b["type"]
    if t == "section":
        return f'<h2>{inline_html(b["title"])}</h2>'
    if t == "para":
        return f'<p class="para">{inline_html(b["text"])}</p>'
    if t == "role":
        org_inline = (f'<span class="org-inline"> · {inline_html(b["org"])}</span>'
                      if b.get("inline") and b.get("org") else "")
        out = (f'<h3 class="role"><span>{inline_html(b["title"])}{org_inline}</span>'
               f'<span class="dates">{inline_html(b["dates"])}</span></h3>')
        if b.get("org") and not b.get("inline"):
            out += f'\n<p class="org">{inline_html(b["org"])}</p>'
        return out
    if t == "project":
        flat = " flat" if b.get("flat") else ""
        tag = "h3" if b.get("flat") else "h4"  # a project is a peer of a role; a client sits under one
        name = "".join(f'<a href="{html.escape(url, quote=True)}">{inline_html(label)}</a>'
                       if url else inline_html(label) for label, url in project_runs(b["name"]))
        out = f'<{tag} class="project{flat}">{name}</{tag}>'
        if b.get("tech"):
            out += f'\n<p class="tech{flat}">{inline_html(b["tech"])}</p>'
        return out
    if t == "bullets":
        cls = ' class="flat"' if b.get("flat") else ""
        items = "\n".join(f"  <li>{inline_html(i)}</li>" for i in b["items"])
        return f"<ul{cls}>\n{items}\n</ul>"
    if t == "skills":
        return "\n".join(f'<p class="skill"><strong>{inline_html(k)}:</strong> {inline_html(v)}</p>'
                         for k, v in b["rows"])
    if t == "edu":
        out = f'<h3 class="edu">{inline_html(b["title"])}</h3>'
        if b.get("meta"):
            out += f'\n<p class="meta">{inline_html(b["meta"])}</p>'
        if b.get("text"):
            out += f'\n<p>{inline_html(b["text"])}</p>'
        return out
    raise ValueError(t)


def indent(s, n):
    pad = " " * n
    return "\n".join(pad + line if line else line for line in s.split("\n"))


def build_html(cv, primary):
    def row(items):
        # Separators are real text (not CSS) so copied or parsed text keeps them apart.
        return "<p>" + "&ensp;·&ensp;".join(
            f'<a href="{html.escape(c["href"])}">{inline_html(c["text"])}</a>' if c.get("href")
            else inline_html(c["text"]) for c in items) + "</p>"

    contact = "\n".join("  " + row(r) for r in cv["contact"])
    sheets = []
    for n, page in enumerate(cv["pages"]):
        parts = []
        if n == 0:
            parts.append(f'<header class="hero">\n  <h1>{inline_html(cv["name"])}</h1>\n'
                         f'  <p>{inline_html(cv["title"])}</p>\n</header>')
            parts.append(f'<address class="contact">\n{contact}\n</address>')
        parts += [block_html(b) for b in page]
        sheets.append('<section class="sheet">\n' + indent("\n".join(parts), 2) + "\n</section>")
    robots = "" if primary else '\n    <meta name="robots" content="noindex" />'
    desc = html.escape(cv.get("description", ""), quote=True)
    social = ""
    if primary and cv.get("url"):
        person = {"@context": "https://schema.org", "@type": "Person", "name": cv["name"], "jobTitle": cv["title"],
                  "url": cv["url"], "email": cv["contact"][0][0]["href"],
                  "sameAs": [c["href"] for r in cv["contact"] for c in r if c.get("href", "").startswith("https://")
                             and c["href"] != cv["url"]],
                  "alumniOf": ["University College London", "Royal Holloway, University of London"]}
        social = f"""
    <link rel="canonical" href="{cv["url"]}/" />
    <meta property="og:type" content="profile" />
    <meta property="og:title" content="{html.escape(cv["name"])} – {html.escape(cv["title"])}" />
    <meta property="og:description" content="{desc}" />
    <meta property="og:url" content="{cv["url"]}/" />
    <script type="application/ld+json">{json.dumps(person, ensure_ascii=False)}</script>"""
    return f"""<!DOCTYPE html>
<html lang="en-GB">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{html.escape(cv["name"])} – {html.escape(cv["title"])} – CV</title>
    <meta name="description" content="{desc}" />{robots}{social}
    <!-- Generated by tools/build.py from {cv["_source"]}. Edit the data file, then rebuild. -->
    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
    <link
      rel="stylesheet"
      href="https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400&display=swap"
    />
    <style>{CSS}    </style>
  </head>

  <body>
    <main class="cv-wrapper scale-up-center">
{indent(chr(10).join(sheets), 6)}
    </main>

    <nav class="downloads" aria-label="Download this CV">
      <a href="{cv['_stem']}.pdf" download="{cv['file_stem']}.pdf">PDF</a>
      <a href="{cv['_stem']}.docx" download="{cv['file_stem']}.docx">Word</a>
    </nav>

  </body>
</html>
"""


# ----------------------------------------------------------------------- DOCX
def build_docx(cv, path):
    from docx import Document
    from docx.enum.section import WD_ORIENT  # noqa: F401
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor

    FONT = "Arial"  # Available in Word on both macOS and Windows.
    GREY = RGBColor(0x5A, 0x5A, 0x5A)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(13)
    sec.top_margin, sec.bottom_margin = Mm(8), Mm(8)
    text_width = Mm(210 - 26)

    def set_font(style, size, bold=None, color=None):
        f = style.font
        f.name, f.size = FONT, Pt(size)
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rfonts.set(qn(a), FONT)
        if bold is not None:
            f.bold = bold
        if color is not None:
            f.color.rgb = color

    # British English, so Word does not underline "colour" or "practised".
    rpr_default = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    lang = rpr_default.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr_default.append(lang)
    lang.set(qn("w:val"), "en-GB")

    normal = doc.styles["Normal"]
    set_font(normal, PT["body"], color=RGBColor(0x2B, 0x2B, 0x2B))
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.line_spacing = 1.02
    h1 = doc.styles["Heading 1"]
    set_font(h1, PT["section"], bold=False, color=RGBColor(0x36, 0x36, 0x36))
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before, h1.paragraph_format.space_after = Pt(9), Pt(4)
    h1.paragraph_format.keep_with_next = True
    for name in ("List Bullet",):
        set_font(doc.styles[name], PT["body"])

    def para(style=None):
        return doc.add_paragraph(style=style)

    def add_text(p, text, size=None, bold=None, italic=None, color=None):
        for chunk, is_bold in inline_runs(text):
            r = p.add_run(chunk)
            r.bold = True if (is_bold or bold) else None
            r.italic = italic
            if size:
                r.font.size = Pt(size)
            if color is not None:
                r.font.color.rgb = color
        return p

    def shade(p, fill):
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
        p._p.get_or_add_pPr().append(shd)

    def rule(p, side, color, sz, space):
        ppr = p._p.get_or_add_pPr()
        bdr = ppr.find(qn("w:pBdr"))
        if bdr is None:
            bdr = OxmlElement("w:pBdr")
            ppr.append(bdr)
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), str(space)); el.set(qn("w:color"), color)
        bdr.append(el)

    def hyperlink(p, text, url, size, color, bold=None):
        rid = p.part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                               is_external=True)
        link = OxmlElement("w:hyperlink")
        link.set(qn("r:id"), rid)
        r = p.add_run(text)
        r.font.size, r.font.color.rgb, r.font.underline = Pt(size), color, True
        r.bold = bold
        p._p.remove(r._r)
        link.append(r._r)
        p._p.append(link)

    def contact_lines():
        rows = cv["contact"]
        for n, items in enumerate(rows):
            p = para()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if n == 0:
                p.paragraph_format.space_before = Pt(6)
            if n == len(rows) - 1:
                rule(p, "bottom", "DBDBDB", 4, 3)
            for i, c in enumerate(items):
                if i:
                    add_text(p, "  ·  ", size=PT["contact"], color=GREY)
                if c.get("href"):
                    hyperlink(p, c["text"], c["href"], PT["contact"], GREY)
                else:
                    add_text(p, c["text"], size=PT["contact"], color=GREY)

    def hero():
        # Two shaded paragraphs; thick same-colour borders give the banner its padding,
        # so there are no empty paragraphs, tables or text boxes for a parser to trip on.
        white = RGBColor(0xFF, 0xFF, 0xFF)
        for text, size, bold, side in ((cv["name"], PT["name"], True, "top"), (cv["title"], PT["title"], None, "bottom")):
            p = para()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.0
            shade(p, BLUE)
            rule(p, side, BLUE, 96, 4)
            r = p.add_run(text)
            r.font.size, r.font.color.rgb, r.bold = Pt(size), white, bold

    def indent_mm(p, mm):
        p.paragraph_format.left_indent = Mm(mm)

    def block(b):
        t = b["type"]
        if t == "section":
            doc.add_paragraph(b["title"], style="Heading 1")
        elif t == "para":
            p = add_text(para(), b["text"])
            p.paragraph_format.space_after = Pt(3.5)
        elif t == "role":
            p = para()
            p.paragraph_format.space_before = Pt(5.5)
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.tab_stops.add_tab_stop(text_width, WD_TAB_ALIGNMENT.RIGHT)
            add_text(p, b["title"], size=PT["role"], bold=True)
            if b.get("inline") and b.get("org"):
                add_text(p, " · " + b["org"], italic=True, color=GREY)
            add_text(p, "\t" + b["dates"], size=PT["role"], italic=True, color=GREY)
            if b.get("org") and not b.get("inline"):
                p = add_text(para(), b["org"], italic=True, color=GREY)
                p.paragraph_format.space_after = Pt(1.5)
                p.paragraph_format.keep_with_next = True
        elif t == "project":
            flat = b.get("flat")
            p = para()
            for label, url in project_runs(b["name"]):
                if url:
                    hyperlink(p, label, url, PT["role"] if flat else PT["body"],
                              RGBColor(0x2B, 0x2B, 0x2B), bold=True)
                else:
                    add_text(p, label, bold=True, size=PT["role"] if flat else None)
            p.paragraph_format.space_before = Pt(5.5 if flat else 3.5)
            p.paragraph_format.keep_with_next = True
            indent_mm(p, 0 if flat else 5)
            if b.get("tech"):
                p = add_text(para(), b["tech"], size=PT["small"], italic=True, color=GREY)
                p.paragraph_format.keep_with_next = True
                indent_mm(p, 0 if flat else 5)
        elif t == "bullets":
            for item in b["items"]:
                p = add_text(para("List Bullet"), item)
                left = 4.5 if b.get("flat") else 9.5
                p.paragraph_format.left_indent = Mm(left)
                p.paragraph_format.first_line_indent = Mm(-4)
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0.6)
        elif t == "skills":
            for k, v in b["rows"]:
                p = para()
                add_text(p, k + ":", bold=True)
                add_text(p, " " + v)
                p.paragraph_format.space_after = Pt(2)
        elif t == "edu":
            p = add_text(para(), b["title"], bold=True)
            p.paragraph_format.space_before = Pt(4.5)
            p.paragraph_format.keep_with_next = True
            if b.get("meta"):
                add_text(para(), b["meta"], italic=True, color=GREY)
            if b.get("text"):
                add_text(para(), b["text"])
        else:
            raise ValueError(t)

    # python-docx starts with one empty paragraph; reuse the body from scratch.
    body = doc.element.body
    for el in list(body):
        if el.tag == qn("w:p"):
            body.remove(el)

    for n, page in enumerate(cv["pages"]):
        if n == 0:
            hero()
            contact_lines()
        for b in page:
            block(b)
        if n < len(cv["pages"]) - 1:
            doc.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)

    core = doc.core_properties
    core.author, core.title = cv["name"], f'{cv["name"]} – {cv["title"]} – CV'
    core.keywords = cv.get("keywords", "")
    doc.save(path)


# ------------------------------------------------------------------------ PDF
def build_pdf(html_path, pdf_path):
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1200, "height": 1000})
        page.goto(html_path.as_uri())
        page.evaluate("document.fonts.ready")
        page.wait_for_timeout(800)
        # A sheet whose content is taller than A4 would be clipped: fail loudly instead.
        over = page.evaluate("""() => [...document.querySelectorAll('.sheet')].map((s, i) => {
            const last = s.lastElementChild.getBoundingClientRect(), box = s.getBoundingClientRect();
            const pad = parseFloat(getComputedStyle(s).paddingBottom);
            return {page: i + 1, spare_px: Math.round(box.bottom - pad - last.bottom)};
        })""")
        for o in over:
            print(f"  page {o['page']}: {o['spare_px']}px spare" + ("  <-- OVERFLOW" if o["spare_px"] < 0 else ""))
        if "--preview" in sys.argv:
            page.screenshot(path=str(pdf_path.with_suffix(".preview.png")), full_page=True)
        page.pdf(path=str(pdf_path), format="A4", print_background=True, prefer_css_page_size=True, tagged=True, outline=True)
        browser.close()
    return all(o["spare_px"] >= 0 for o in over)


def build_txt(cv):
    """Plain-text copy in reading order: for pasting into application forms and for parser checks."""
    plain = lambda t: re.sub(r"\*\*(.+?)\*\*", r"\1", t)
    out = [cv["name"], cv["title"]] + [" | ".join(c["text"] for c in r) for r in cv["contact"]]
    for b in (b for page in cv["pages"] for b in page):
        t = b["type"]
        if t == "section":
            out += ["", b["title"]]
        elif t == "para":
            out.append(plain(b["text"]))
        elif t == "role":
            out.append(f'{b["title"]}, {b["dates"]}')
            if b.get("org"):
                out.append(b["org"])
        elif t == "project":
            tech = b.get("tech", "")
            if tech[:3] in ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"):
                out.append(f'{b["name"]} · {tech}')  # never start a line with a date
            else:
                out.append(b["name"])
                if tech:
                    out.append(tech)
        elif t == "bullets":
            out += ["- " + plain(i) for i in b["items"]]
        elif t == "skills":
            out += [f"{k}: {v}" for k, v in b["rows"]]
        elif t == "edu":
            out.append(", ".join(x for x in (b["title"], b.get("meta")) if x))
            if b.get("text"):
                out.append(plain(b["text"]))
    return "\n".join(out) + "\n"


def main(argv):
    args = [a for a in argv if not a.startswith("--")]
    src = Path(args[0]).resolve()
    out = Path(argv[argv.index("--out") + 1]).resolve() if "--out" in argv else None
    if out:
        args.remove(argv[argv.index("--out") + 1])
    cv = json.loads(src.read_text())
    primary = out is None
    cv["_source"] = str(src.relative_to(ROOT)) if primary else src.name
    out = out or ROOT
    out.mkdir(parents=True, exist_ok=True)
    cv.setdefault("file_stem", cv["name"].replace(" ", "-") + "-CV")
    cv.setdefault("description", f'{cv["name"]}: {cv["title"]}')
    stem = "CV" if primary else cv["file_stem"]  # copies built elsewhere are named ready to send
    cv["_stem"] = stem
    (out / "index.html").write_text(build_html(cv, primary))
    build_docx(cv, out / f"{stem}.docx")
    if not primary:
        (out / f"{stem}.txt").write_text(build_txt(cv))
    if primary and cv.get("url"):
        (out / "robots.txt").write_text(f"User-agent: *\nAllow: /\nSitemap: {cv['url']}/sitemap.xml\n")
        (out / "sitemap.xml").write_text('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
                                         f"  <url><loc>{cv['url']}/</loc></url>\n</urlset>\n")
    ok = build_pdf(out / "index.html", out / f"{stem}.pdf")
    print(("built " if ok else "BUILT WITH OVERFLOW ") + str(out))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
