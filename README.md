# Microscript Card & Grid PDF Generator

This repository provides an automated workflow to format structured notes and generate a compact, 3x3 grid landscape PDF (ideal for microscript slide decks, flashcards, or study guides) from Microsoft Word (`.docx`) documents.

---

## Screenshots

<table>
  <tr>
    <td align="center" width="33%"><img src="images/GDoc_before.jpg" width="230" alt="Google Docs"><br><sub><b>Google docs</b><br>Format your presentation notes like this in google docs</sub></td>
    <td align="center" width="33%"><img src="images/GDoc_after.jpg" width="230" alt=".gs script"><br><sub><b>.gs script</b><br>The tag_formatting.gs formats only the bracketed parts.</sub></td>
    <td align="center" width="33%"><img src="images/pdf_final.jpg" width="230" alt=".py script"><br><sub><b>.py script</b><br>The final output of convert.py. This is the topleft card in a 9x9 grid landscape layout.</sub></td>
  </tr>
</table>

---

## Workspace Overview

* **`environment.yml`**: Conda environment definition containing Python 3.11, `python-docx`, `weasyprint`, and necessary system render libraries (Cairo, Pango, Fontconfig).
* **`tag_formatting.gs`**: Google Apps Script to automatically highlight and style `[TAG]` markers within Google Docs prior to exporting.
* **`convert.py`**: Python script that parses the `.docx` file, builds styled HTML cards, and renders a multi-page A4 landscape grid PDF.

---

## 1. Environment Setup

### Prerequisites
* [Miniconda](https://docs.conda.io/en/latest/miniconda.html) or [Anaconda](https://www.anaconda.com/) installed on your system.

### Installation

1. Clone or download this repository to your local machine.
2. Open your terminal (or Anaconda Prompt on Windows) and navigate to the project directory:
   ```bash
   cd /path/to/repository
   ```
3. Create the Conda environment using the provided `environment.yml` file:
   ```bash
   conda env create -f environment.yml
   ```
4. Activate the newly created environment:
   ```bash
   conda activate weasyprint
   ```

---

## 2. Using the Google Apps Script (`tag_formatting.gs.txt`)

If you edit your original notes in Google Docs, you can use the Google Apps Script to visually standardize all bracketed tag elements (e.g., `[INFO]`, `[NOTE]`, `[KEY]`) before exporting.

### Steps:
1. Open your document in **Google Docs**.
2. In the top menu, go to **Extensions** > **Apps Script**.
3. Replace any code in the editor with the code inside `tag_formatting.gs.txt`:
   ```javascript
   function colorTagsGreen() {
     var body = DocumentApp.getActiveDocument().getBody();
     var pattern = "\\[[^\\]]+\\]"; // Matches anything inside [ square brackets ]
     var found = body.findText(pattern);
     
     while (found) {
       var element = found.getElement().asText();
       var start = found.getStartOffset();
       var end = found.getEndOffsetInclusive();
       
       element.setForegroundColor(start, end, "#008000"); // Dark Green
       element.setBold(start, end, true);                // Bold
       element.setFontFamily(start, end, "Roboto Mono"); // Monospace font
       
       found = body.findText(pattern, found);
     }
   }
   ```
4. Click **Save** (disk icon) and then click **Run**. Grant permissions if prompted.
5. Export your document to Word format via **File** > **Download** > **Microsoft Word (.docx)**.

---

## 3. Using `convert.py`

`convert.py` parses the structured `.docx` file, converts paragraphs starting with `[TAG]` into styled bullet items, groups paragraphs into discrete cards (one header per card), and compiles them 9-per-page (3x3 grid) into a printable A4 landscape PDF.

### Document Structure Guidelines:
* **Headers/Titles**: Any plain text line without a leading `[TAG]` acts as a card header (e.g., `Slide 1 Title`).
* **Tagged Items**: Any line starting with square brackets (e.g., `[DEF] Definition text...`) becomes a bullet point within the active card.

### Running the Conversion Script:

Make sure your Conda environment is active:
```bash
conda activate weasyprint
```

Run the script by providing the path to your input `.docx` file and the desired output `.pdf` path:

```bash
python convert.py input_document.docx output_deck.pdf
```

### Example Usage:
```bash
python convert.py ./notes/lecture_01.docx ./output/lecture_01_cards.pdf
```

> **Overflowing cards** You may notice, that some cards can appear to be overflowing and breaking the grid structure. This 'bug' is actually a feature, which enforces an upper limit for the amount of bulletpoints/text within a single card.


---

## Layout & Output Specifications

* **Page Format**: A4 Landscape (with 6mm margins).
* **Grid**: 3 columns × 3 rows (9 microscript cards per page).
* **Styling**:
  * Slide Headers: Dark grey/blue headers with strong contrast.
  * Tag Badges: Styled in blue background (`#E8F1FF`) with blue monospace font (`Roboto Mono`).
  * Supports rich text formats: bold, italic, underline, subscript, superscript, and custom text colors preserved from Word.