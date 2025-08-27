# autoPPT
Turn bulk text, markdown, or prose into a fully formatted PowerPoint presentation that matches their chosen template’s look and feel.

---

## Features

* **Paste** long text or markdown
* **(Optional) Guidance**: one‑line goal like “turn into an investor pitch deck”
* **Choose a provider**: OpenAI, Anthropic, Gemini, **or** a local heuristic (no external calls)
* **Upload .pptx/.potx** template or sample presentation
* **Template‑true output**: slide layouts, fonts, colors, and existing images are applied
* **Smart slide count**: chooses a reasonable number of slides; you can nudge with a target slider
* **Privacy**: API keys are used only in‑memory for the active request; not stored or logged
* **Download .pptx** result

---

## How it works

1. **Outline generation**
   The app converts your input into a strict JSON outline: slide titles, concise bullets, optional speaker notes, and layout hints. This is done via your chosen LLM or a local markdown‑aware heuristic.

2. **Template introspection**
   The uploaded PowerPoint is loaded as the base so theme + fonts + layouts carry over. The app scans slides, masters, and layouts to collect **existing images** for reuse.

3. **Slide assembly**
   All existing slides are cleared (layouts remain). New slides are added using the template’s actual layouts; picture placeholders are filled from the collected image pool. If a slide has no body content, a tasteful image may be placed to balance the composition.

4. **Export**
   A fresh `.pptx` is produced for download—no template is modified on disk.

---

## Tech stack

* **Python** + **Streamlit** UI
* **python-pptx** for building slides and inheriting theme/layouts
* Optional SDKs: **openai**, **anthropic**, **google‑generativeai**

---

## Local run

```bash
# 1) Create and activate a virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate

# 2) Install dependencies
pip install -r requirements.txt

# 3) Run
streamlit run app.py
```

Open the local URL Streamlit prints, then:

1. Paste your text/markdown
2. (Optional) Add one‑line guidance
3. Pick provider and paste your API key (or choose **Local heuristic**)
4. Upload `.pptx` or `.potx` template
5. Click **Generate presentation** → **Download .pptx**

> **Note**: Keys are kept only in memory for the request and aren’t persisted by the app. If you host this yourself, ensure your platform’s logging doesn’t capture secrets (see Security below).

---



## 📄 License

Choose a license that fits your use (e.g., MIT). Add a `LICENSE` file in the repo.
