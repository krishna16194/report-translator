# Launch the Document Translator web app.
python -m uvicorn app.main:app --port 8000
Write-Host "Open http://127.0.0.1:8000 in your browser."
