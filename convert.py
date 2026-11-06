import docx
import re
from weasyprint import HTML
import argparse

def extract_line_components(p):
    """
    Parses a paragraph into (tag_html, body_html).
    Strips inline font colors from tags to ensure CSS overrides apply.
    """
    tag_regex = re.compile(r'^\s*(\[[^\]]+\])\s*(.*)')
    plain_text = p.text.strip()
    match = tag_regex.match(plain_text)

    if not match:
        return None, paragraph_runs_to_html(p.runs)

    tag_plain = match.group(1)

    tag_html_runs = []
    body_html_runs = []
    
    char_count = 0
    tag_length = len(tag_plain)

    for run in p.runs:
        text = run.text
        if not text:
            continue

        run_len = len(text)
        
        if char_count + run_len <= tag_length:
            tag_html_runs.append(format_run(run, ignore_color=True))
        elif char_count >= tag_length:
            body_html_runs.append(format_run(run, ignore_color=False))
        else:
            split_idx = tag_length - char_count
            tag_text = text[:split_idx]
            body_text = text[split_idx:]
            
            tag_html_runs.append(format_run(run, text_override=tag_text, ignore_color=True))
            body_html_runs.append(format_run(run, text_override=body_text, ignore_color=False))

        char_count += run_len

    tag_html = f'<span class="tag">{"".join(tag_html_runs)}</span>'
    body_html = "".join(body_html_runs).strip()

    return tag_html, body_html

def format_run(run, text_override=None, ignore_color=False):
    text = text_override if text_override is not None else run.text
    if not text:
        return ""

    text_escaped = (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
    )

    if run.bold:
        text_escaped = f"<b>{text_escaped}</b>"
    if run.italic:
        text_escaped = f"<i>{text_escaped}</i>"
    if run.underline:
        text_escaped = f"<u>{text_escaped}</u>"
    if run.font.superscript:
        text_escaped = f"<sup>{text_escaped}</sup>"
    if run.font.subscript:
        text_escaped = f"<sub>{text_escaped}</sub>"

    if not ignore_color and run.font.color and run.font.color.rgb:
        color_hex = f"#{run.font.color.rgb}"
        text_escaped = f'<span style="color: {color_hex};">{text_escaped}</span>'

    return text_escaped

def paragraph_runs_to_html(runs):
    return "".join(format_run(r) for r in runs)

def extract_cards_from_docx(docx_path):
    doc = docx.Document(docx_path)
    tag_regex = re.compile(r'^\s*\[[^\]]+\]')
    
    cards = []
    current_title = None
    current_items = []

    for p in doc.paragraphs:
        plain_text = p.text.strip()
        if not plain_text:
            continue

        tag_html, body_html = extract_line_components(p)

        if not tag_regex.match(plain_text):
            if current_title is not None:
                cards.append((current_title, current_items))
                current_items = []
            current_title = body_html
        else:
            if current_title is None:
                current_title = body_html
            current_items.append(f'<li>{tag_html} {body_html}</li>')

    if current_title is not None:
        cards.append((current_title, current_items))

    formatted_cards = []
    for idx, (title, items) in enumerate(cards, 1):
        card_html = f'''
        <div class="card">
          <div class="slide-header">Slide {idx}: {title}</div>
          <ul>
            {"".join(items)}
          </ul>
        </div>
        '''
        formatted_cards.append(card_html)

    return formatted_cards

def build_pdf_from_docx(input_docx, output_pdf):
    cards = extract_cards_from_docx(input_docx)
    chunk_size = 9
    pages_html = []

    total_pages = (len(cards) + chunk_size - 1) // chunk_size

    for page_idx in range(total_pages):
        chunk = cards[page_idx * chunk_size : (page_idx + 1) * chunk_size]
        grid_items = "".join(chunk)

        page_content = f'''
        <div class="page-container">
          <div class="cards-grid">
            {grid_items}
          </div>
        </div>
        '''
        pages_html.append(page_content)

    full_html = f'''
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto+Mono:wght@500;700&display=swap');

        @page {{
          size: A4 landscape;
          margin: 6mm;
        }}
        html, body {{
          margin: 0;
          padding: 0;
          height: 100%;
          font-family: "Helvetica Neue", Helvetica, "Liberation Sans", "DejaVu Sans", Arial, sans-serif;
          font-size: 8.0pt;
          line-height: 1.15;
          color: #1C252E;
        }}
        .page-container {{
          width: 100%;
          height: 198mm;
          page-break-after: always;
          box-sizing: border-box;
          display: flex;
          flex-direction: column;
        }}
        .cards-grid {{
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          grid-template-rows: repeat(3, 1fr);
          gap: 1mm;
          flex-grow: 1;
          box-sizing: border-box;
        }}
        .card {{
          background: #fff;
          border: 1px solid #d0d7de;
          border-radius: 2px;
          padding: 4px 6px;
          box-sizing: border-box;
          overflow: hidden;
          display: flex;
          flex-direction: column;
        }}
        .slide-header {{
          display: block;
          font-weight: bold;
          font-size: 7.8pt;
          color: #1a252f;
          background-color: #e2e7ec; /* Darkened from #f1f3f5 */
          padding: 3px 5px;
          margin-bottom: 3px;
          border-left: 3px solid #1c2d42;
          border-radius: 1px;
          line-height: 1.15;
          flex-shrink: 0;
        }}
        .slide-header * {{
          color: #1a252f !important;
        }}
        ul {{
          margin: 0;
          padding-left: 11px;
          list-style-type: disc;
          flex-grow: 1;
        }}
        li {{
          margin-bottom: 5px;
          line-height: 1.15;
        }}
        .tag {{
          font-family: 'Roboto Mono', monospace;
          font-weight: 700;
          font-size: 6.8pt;
          background-color: #E8F1FF;
          color: #0066FF !important;
          padding: 2px 2px;
          border-radius: 2px;
        }}
        .tag * {{
          color: #0066FF !important;
        }}
        b, strong {{
          font-weight: bold;
          color: #000000 !important;
        }}
        sup {{
          font-size: 0.75em;
          line-height: 0;
          vertical-align: super;
        }}
        sub {{
          font-size: 0.75em;
          line-height: 0;
          vertical-align: sub;
        }}
      </style>
    </head>
    <body>
      {"".join(pages_html)}
    </body>
    </html>
    '''

    HTML(string=full_html).write_pdf(output_pdf)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert a structured Word document into a formatted landscape microscript PDF grid.")
    parser.add_argument("input_docx", type=str, help="Path to the source .docx document.")
    parser.add_argument("output_pdf", type=str, help="Path to save the resulting .pdf file.")

    args = parser.parse_args()
    build_pdf_from_docx(args.input_docx, args.output_pdf)