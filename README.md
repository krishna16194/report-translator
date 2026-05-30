# LinguaDoc — Document Translator

A web app that takes a **Word, Excel, or text file**, translates its contents into
a preferred language, and produces a report as a **Word (.docx)** and/or
**PowerPoint (.pptx)** document.

The UI is branded as **LinguaDoc** with an SVG logo ([static/logo.svg](static/logo.svg))
that doubles as the browser-tab favicon, a value-proposition description, feature
chips, and a 3-step "how it works" footer.

- **Interface:** browser web app (upload form)
- **Translation:** Google Translate via [`deep-translator`](https://pypi.org/project/deep-translator/) — free, no API key
- **Input formats:** `.docx`, `.xlsx`/`.xlsm`, `.txt`
- **Output formats:** `.docx`, `.pptx` (pick one or both → both come back as a `.zip`)
- **Design:** a built-in **standard professional template** (cover page, accent
  colours, styled headings/tables, page numbers, section-aware slides), or
  **bring your own** `.docx` / `.pptx` template whose theme & layouts are reused
- **OCR / scanned documents:** not included yet (planned — see below)

## Report design

By default the report uses a polished standard template:

- **Word:** centred cover page (title, accent rule, language + date), styled
  Heading 1/2/3, banded tables with a coloured header row, and page numbers.
- **PowerPoint (16:9):** title slide with accent bands, one slide per section
  (paragraphs become bullets, long sections continue on "(cont.)" slides),
  dedicated table slides, and section-divider slides — so content is spread out
  instead of crammed onto one slide.

To match your own branding, choose **"Use my own template"** and upload a
`.docx` and/or `.pptx`. The translated content is laid out using that file's
fonts, colours, and slide layouts. Leave a format's template blank to keep the
standard design for it.

## Setup

```powershell
cd doc-translator
pip install -r requirements.txt
```

## Run

```powershell
python -m uvicorn app.main:app --port 8000
```

Then open <http://127.0.0.1:8000> in your browser, choose a file, pick a target
language and output format, and click **Translate & download**.

## Deploy (Render)

> ⚠️ This is a **FastAPI (Python) app** — it cannot run on Cloudflare
> Workers/Pages, which only serve static files. Deploying just `static/` there
> makes `/api/*` return **404** because the Python backend isn't running. Use a
> Python-capable host. A `render.yaml` blueprint is included for [Render](https://render.com):

1. Push this repo to GitHub (already done: `krishna16194/report-translator`).
2. In Render: **New + → Blueprint**, connect the repo, and **Apply**. Render
   reads [`render.yaml`](render.yaml) and configures the build/start commands.
   (Or **New + → Web Service** and set Build = `pip install -r requirements.txt`,
   Start = `uvicorn app.main:app --host 0.0.0.0 --port $PORT`.)
3. Open the URL Render gives you (e.g. `https://report-translator.onrender.com`).
   The frontend and `/api/*` are served from the **same origin**, so it just works.

Any other Python host works too (Azure App Service, Railway, Fly.io) using the
same start command. On the free Render tier the service sleeps when idle, so the
first request after a pause can take ~30s to wake.

## How it works

```
upload  ->  app/parsers.py     extract a format-agnostic content model
        ->  app/translator.py  translate every text segment (cached, chunked)
        ->  app/generators.py  rebuild as .docx / .pptx
        ->  app/main.py        return a file (or .zip for both formats)
```

The content model is a list of blocks (`heading`, `paragraph`, `table`), which
keeps parsing, translation, and generation independent of each other.

## Project layout

| Path | Purpose |
|------|---------|
| `app/main.py` | FastAPI app + routes (`/api/languages`, `/api/translate`) |
| `app/parsers.py` | Read `.txt` / `.docx` / `.xlsx` into the content model |
| `app/translator.py` | Translate the content model; language list |
| `app/generators.py` | Build `.docx` and `.pptx` output |
| `static/index.html` | Frontend upload UI |
| `smoke_test.py` | Offline-ish pipeline check (`python smoke_test.py`) |

## Supported languages

English, Spanish, French, German, Italian, Portuguese, Dutch, Russian, Arabic,
Hindi, Tamil, Telugu, Kannada, Malayalam, Bengali, Japanese, Korean, Chinese
(Simplified & Traditional). Add more in `LANGUAGES` in `app/translator.py`.

## Roadmap / not yet done

- **Scanned documents & images (OCR):** add a parser branch for `.pdf`/`.png`/`.jpg`
  using Tesseract or a cloud OCR, returning the same content model.
- **Layout fidelity:** current output is a clean re-flow, not a pixel copy of the
  source. Preserving original styling would need richer parsing.
- **Production translation:** `deep-translator`'s free endpoint is rate-limited;
  swap to the official Google Cloud Translation API for higher volume.

## Notes

- Translation requires internet access (calls Google's public endpoint).
- The free endpoint is rate-limited; very large files may hit limits.
