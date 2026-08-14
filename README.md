# PII Redaction Tool

A small Python tool to detect and replace personally identifiable information (PII) in Microsoft Word (.docx) documents.

## Features


Detects and redacts common PII types:

- Person names
- Organization names
- Email addresses
- Phone numbers
- IP addresses
- Credit card numbers
- Social Security Numbers (SSNs)
- Postal addresses

Redacted values are replaced with generated placeholders (examples):

- `Person 1`
- `Example Organization 1`
- `person1@example.com`
- `+91 90000 00001`
- `192.0.2.1`
- `Example Address 1`

## How It Works

The redaction pipeline works in the following stages:

1. Loads the input `.docx` document using `python-docx`.
2. Extracts text from document paragraphs and table cells.
3. Uses spaCy Named Entity Recognition to identify person and organization candidates.
4. Uses regular expressions and validation rules to detect structured PII such as emails, phone numbers, IP addresses, credit cards, and SSNs.
5. Filters false positives and already-generated placeholder values.
6. Generates replacement values for detected PII.
7. Replaces the detected values while preserving the DOCX document structure.
8. Saves the redacted document in the `output/` directory.
9. Runs validation checks to verify that detected PII values did not survive.

## Technology

- Python 3
- spaCy (NER)
- python-docx (DOCX manipulation)
- phonenumbers
- pytest
- Regular expressions

## Project Structure

```
PII-Redaction-Tool/
├── input/                      # place source .docx files here
├── output/                     # redacted .docx files are written here
├── tests/                      # pytest unit tests
├── pii_redactor.py             # main redaction script
├── validate_redaction.py       # document-level validation
├── review_candidates.py        # review low-confidence candidates
├── requirements.txt
└── README.md
```

## Installation

Create and activate a virtual environment, then install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run the Redaction Tool

1. Put the source Word document in the `input/` directory (for example `input/document.docx`).
2. Run the tool:

```bash
python pii_redactor.py
```

The redacted file will be written into `output/` (name is based on the input filename).

## Validation

After creating the redacted document run the validation scripts to confirm PII removal:

```bash
python validate_redaction.py
python review_candidates.py > review_output.txt
grep -n "STILL PRESENT" review_output.txt || echo "No configured important names remain"
```

`validate_redaction.py` compares detected original PII values to the redacted DOCX and reports any survivors.

## Unit Tests

Run unit tests with `pytest`:

```bash
python -m pytest -q
```

The test suite currently exercises detection, placeholder generation, and validation utilities.

## Example Validation Summary

An example validation report might look like:

```
EMAIL           original=26 survived=0
PHONE           original=23 survived=0
IP              original=0  survived=0
CREDIT_CARD     original=0  survived=0
SSN             original=0  survived=0
Original PII values checked : 49
PII values surviving        : 0
STATUS: PASS
```

And a final person-name verification summary:

```
FINAL PERSON NAME CHECK: PASS
```

## Output Behavior

- Replacement values are generated and cached so repeated occurrences of the same original PII receive the same placeholder within a run.
- The tool currently preserves DOCX structure (paragraphs and table cells) while replacing text.

## Recommended Workflow

1. Place document in `input/`
2. Run `python pii_redactor.py`
3. Run `python validate_redaction.py`
4. Review `review_output.txt` for low-confidence candidates
5. Run unit tests if you modified detection logic

## Limitations

- Detection accuracy depends on the input document and NLP model.
- Some ambiguous names or contextual PII may require manual review.
- Complex DOCX formatting should be visually checked after redaction.
- The current implementation is focused on Microsoft Word `.docx` documents.
- Uncommon or highly contextual PII may require additional detection rules or model improvements.

## Notes
- Exclude generated redacted documents and local virtual environments using `.gitignore`.
- Test the redacted output before using it in production.