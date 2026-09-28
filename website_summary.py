# Website Summarizer

A Python script that summarizes a single web page (for example, a home page) and saves the result as an Excel table.

## What It Does

- Reads only the page you give it, not the whole website.
- Skips menus, footers, and hidden content.
- Creates an Excel table with two columns: **Short Summary** and **Description**.
- The first row summarizes the whole page, and the following rows summarize each section.
- Saves the file automatically in a separate folder and opens it when finished.

No API key or account is needed.

## Requirements

- Python 3.8 or newer
- VS Code (or any editor)

Install the required packages once:

```bash
python3 -m pip install -r requirements.txt
```

## How to Use

1. Open `website_summary.py`.
2. Change the page link at the top:

   ```python
   PAGE_URL = "https://www.bbc.com/news"
   ```

3. (Optional) Choose where to save the result:

   ```python
   SAVE_FOLDER = ""   # empty = Downloads/Website Summaries
   ```

4. Run the script:

   ```bash
   python3 website_summary.py
   ```

## Output

The Excel file is saved to `Downloads/Website Summaries` by default, with the site name and time in the file name, for example:

```
bbc.com_summary_2026-09-28_14-30-00.xlsx
```

| Short Summary | Description |
|---|---|
| Page title and main idea | Key sentences from the whole page |
| Section heading | Key sentences from that section |

## Limitations

- Pages that load their content with JavaScript may return little or no text.
- Some websites block automated visitors, so the script may show an error for them.

## Project Files

| File | Purpose |
|---|---|
| `website_summary.py` | The main script |
| `requirements.txt` | Required Python packages |
| `.gitignore` | Keeps generated files out of GitHub |
| `README.md` | This guide |

## Author

Jamil
