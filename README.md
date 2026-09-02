# PII Redaction Tool

This is a Python-based tool designed to automatically detect and redact Personally Identifiable Information (PII) from Word Documents (`.docx`) and replace it with realistic fake data.

## 🛠️ How it Works

The tool relies on **Microsoft Presidio** and **Faker**:
1. **Presidio Analyzer**: This engine combines Machine Learning (spaCy) to find unstructured text (like full names or cities) with Regular Expressions (Regex) to find structured text (like emails, phone numbers, and SSNs).
2. **Custom Recognizers**: To ensure it catches everything, we added custom regex rules for Indian Phone Numbers and specific Date of Birth formats.
3. **Faker**: Once PII is found, we ask Faker to generate a realistic replacement. We cache these replacements so that if a name like "John Doe" appears 5 times, it is always replaced with the exact same fake name, keeping the document readable.

## 📁 Files Included

- `redactor.py`: The main script that reads the DOCX, redacts the PII, and outputs the redacted DOCX and a log file.
- `evaluate_performance.py`: A testing script that compares the generated logs against a known list of PII (Ground Truth) to calculate Precision and Recall.
- `redaction_log.json`: (Generated) A log of every piece of text the script flagged and what it replaced it with.
- `Redacted_Output.docx`: (Generated) The final redacted document.

## ⚖️ Tradeoffs & Known Limitations

Building a PII redactor involves balancing accuracy with formatting. Here are the tradeoffs made in this implementation:

1. **Paragraph Formatting**: To safely replace text that might span across different formatting blocks (runs) in the `.docx` file, this tool flattens the text of each paragraph and rewrites it into a single run. The tradeoff is that mid-paragraph bolding or italics might be lost, but the document structure (headings, tables, lists) and the underlying text are perfectly preserved.
2. **False Positives vs. False Negatives**: We set the detection threshold in Presidio to `0.4`. This lower threshold ensures we catch more names (Recall), but the tradeoff is it occasionally flags non-PII text like financial terms (Precision). We prioritize missing zero PII (high recall) over perfectly clean text.

## 🚀 How to Run

Install the required dependencies:
```bash
pip install presidio-analyzer faker python-docx
python -m spacy download en_core_web_sm
```

Run the redactor:
```bash
python redactor.py
```

Run the evaluation report:
```bash
python evaluate_performance.py
```
