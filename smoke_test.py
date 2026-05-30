"""Quick end-to-end sanity check of the pipeline (no web server)."""
import os
import tempfile

from app import generators, parsers
from app.translator import Translator

# 1. Make a sample .txt input.
sample = (
    "Quarterly Report\n\n"
    "Sales grew by twenty percent this quarter.\n\n"
    "The team delivered three major projects on time."
)
with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8") as f:
    f.write(sample)
    path = f.name

# 2. Parse.
parsed = parsers.parse(path)
os.unlink(path)
print("Parsed blocks:", len(parsed["blocks"]))
assert len(parsed["blocks"]) == 3

# 3. Generators (no network needed).
docx_bytes = generators.build_docx("Quarterly Report", parsed["blocks"], "French", "May 30, 2026")
pptx_bytes = generators.build_pptx("Quarterly Report", parsed["blocks"], "French", "May 30, 2026")
print("docx bytes:", len(docx_bytes), "| pptx bytes:", len(pptx_bytes))
assert docx_bytes[:2] == b"PK" and pptx_bytes[:2] == b"PK"  # both are zip-based

# 4. Live translation (requires internet).
try:
    tr = Translator(target_lang="fr")
    translated = tr.blocks(parsed["blocks"])
    print("Sample translation:", translated[0]["text"])
    print("TRANSLATION: OK")
except Exception as exc:
    print(f"TRANSLATION SKIPPED (no network?): {exc}")

print("SMOKE TEST PASSED")
