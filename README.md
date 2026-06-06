# AI-Assisted Editorial Review Tool

Automatically annotates `.docx` manuscripts with AI-powered editorial suggestions.

## What it does

- 🟡 **Yellow** — flags potential spelling mistakes and typos
- 🔵 **Cyan** — suggests sentences for emphasis (odd chapters)
- 🟢 **Green** — suggests sentences for block quotes (odd chapters)
- 🟠 **Closing heading** — inserts configurable heading near end of each chapter

## Files

| File | Description |
|------|-------------|
| `editorial_review_demo.py` | Main script |
| `sample_manuscript.docx` | Sample input (3 chapters with intentional errors) |
| `sample_annotated.docx` | Output — open in Word to see annotations |

## Setup

```bash
pip install python-docx
python editorial_review_demo.py input.docx output_annotated.docx
```

## Production

Replace `mock_analyze_chapter()` with `call_openai_api()` — the stub is already
in the script, just add your API key.
