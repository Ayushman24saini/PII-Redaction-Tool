# PII Redaction Tool — Evaluation Report

## Summary

This report summarizes the evaluation of the `PII-Redaction-Tool`. The evaluation covers unit tests, document-level validation, placeholder verification, and operational metrics for a validation set of structured PII values.

Key outcome: for the current validation set of 49 high-confidence structured PII values, 0 survived the redaction process (100% removal for that set). Unit tests pass locally (`13/13`).

---

## 1. Objective

Detect PII in Microsoft Word (`.docx`) documents and replace detected values with generated, non-sensitive placeholders. The evaluation focuses on detection correctness, replacement correctness, survival prevention, placeholder validity, and regression testing.

## 2. PII Categories

- Person names
- Organization names
- Email addresses
- Phone numbers
- IP addresses
- Credit card numbers
- Social Security Numbers (SSNs)
- Postal addresses

Detection approach: spaCy NER for person/organization candidates, plus regular expressions, validators, and contextual filters for structured PII.

---

## 3. Evaluation Methodology

The evaluation uses three complementary checks:

- Unit tests (`pytest`) for detection and placeholder generation.
- Document-level validation that searches for high-confidence structured PII from the original DOCX in the redacted DOCX.
- Person-name verification against a configured list of important names.

### 3.1 Unit Tests

Run locally and in CI with:

```bash
python -m pytest -q
```

Current result: `13 passed` (tests cover Luhn checks, phone/email placeholders, person/org/address generation, and invalid-value tests).

### 3.2 Document-Level Validation

`validate_redaction.py` compares original detected structured PII to the redacted document and reports survivors.

Latest validation (structured PII):

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

### 3.3 Person-Name Verification

Separate verification checks configured important names in the final redacted DOCX.

Latest result: `FINAL PERSON NAME CHECK: PASS`.

---

## 4. Placeholder Validation

Counts of generated placeholders observed in the redacted document (latest run):

```
email          : 70
phone          : 50
ip             : 0
credit_card    : 0
ssn            : 0
person         : 303
organization   : 354
address        : 0
```

Placeholders are cached during a run so repeated occurrences of the same original value receive the same replacement.

---

## 5. Metrics

Precision and recall require a manually labeled ground-truth dataset; current validation does not provide full ground truth for every candidate, so single-number precision/recall values are not claimed. The survival-based test is a practical operational check for structured PII.

- PII Survival Rate (validation set): 0 / 49 = 0%
- PII Removal Rate (validation set): 100%

### Current Evaluation Summary

| Metric / Check | Result |
|---|---:|
| Unit tests | 13/13 passed |
| Structured PII values checked | 49 |
| Structured PII values surviving | 0 |
| Structured PII removal rate | 100% |
| Email survivors | 0 |
| Phone survivors | 0 |
| IP survivors | 0 |
| Credit-card survivors | 0 |
| SSN survivors | 0 |
| Person-name verification | PASS |
| GitHub Actions CI | Configured |

---

## 6. False Positives / False Negatives

- False positives are mitigated via validators (phone validation, Luhn checks), contextual filters, and unit tests for invalid values.
- False negatives remain possible for ambiguous or highly contextual PII that NER and rules cannot confidently identify; survival testing alone cannot expose undetected PII.

---

## 7. Limitations

- Validation set is limited and not a full manually annotated ground truth; precision/recall cannot be fully summarized as single figures.
- Detection depends on the spaCy model and rule-based logic; model improvements will affect performance.
- Complex DOCX formatting may require visual inspection after redaction.
- Focus is on Microsoft Word `.docx` documents; other formats are not supported by the current implementation.
- Uncommon or highly contextual PII may require additional detection rules or training data.

---

## 8. Recommendations & Future Work

- Create a manually annotated ground-truth dataset to compute per-category precision, recall, and F1.
- Expand test cases for ambiguous contexts and document formatting edge cases.
- Add per-category benchmarks and visual diff artifacts to aid manual review.
- Consider integrating additional NLP models or fine-tuning spaCy for domain-specific names.

---

## 9. Final Assessment

The current implementation demonstrates reliable removal of the structured PII values included in the validation set and passes its unit tests. For a comprehensive evaluation, produce a representative, manually labeled benchmark and measure precision/recall per PII category.
