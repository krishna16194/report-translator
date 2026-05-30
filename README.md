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

## Two backends

This repo ships **two** implementations of the same UI:

| | Local / full | Cloudflare (deployed) |
|---|---|---|
| Code | `app/` (Python, FastAPI) | `src/` (JavaScript, Worker) |
| Translation | `deep-translator` | Google endpoint via `fetch` |
| Input | `.docx`, `.xlsx`, `.txt` | `.docx`, `.txt` |
| Output design | cover, tables, page numbers, custom templates | cover + headings + bullets (no tables/templates) |
| Limits | none | 2 MB upload, free-tier CPU/subrequest caps |

The Cloudflare build is a **deliberately leaner** version that fits the free
Workers tier (10 ms CPU, 50 subrequests/request). `.xlsx` uploads and custom
templates aren't supported there — use the local Python app for those.

## Deploy (Cloudflare Workers — free)

The Worker serves the static frontend **and** the `/api/*` routes from one origin,
so `/api/languages` returns `200` (no more 404). No paid plan, no Docker.

```powershell
cd doc-translator
npm install                # installs wrangler + fflate
npx wrangler login         # opens a browser to authenticate to Cloudflare
npx wrangler deploy        # bundles src/ + static/ and deploys
```

Wrangler deploys to `report-translator.<your-subdomain>.workers.dev` — i.e. it
**replaces the current static-only deployment** with the full app. Open that URL
and the whole thing works.

Local development of the Worker:

```powershell
npx wrangler dev           # http://127.0.0.1:8787
```

Worker files: [src/index.js](src/index.js) (router), [src/translate.js](src/translate.js),
[src/parse.js](src/parse.js), [src/docx.js](src/docx.js), [src/pptx.js](src/pptx.js),
[wrangler.jsonc](wrangler.jsonc) (static-assets binding). Run `node scripts/test-gen.mjs`
to validate the document generators offline.

### Alternative: full Python app on Render (free)

To deploy the *full* version (Excel input, tables, custom templates) instead, use a
Python host. A [`render.yaml`](render.yaml) blueprint is included: in Render, choose
**New + → Blueprint**, connect this repo, and **Apply** (free tier; sleeps when idle).

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
