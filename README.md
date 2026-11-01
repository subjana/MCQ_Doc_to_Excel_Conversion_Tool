# MCQ Document to Excel

Portfolio owner: **Subjana Islam Tasnia**.

A Python utility for converting structured multiple-choice questions into an Excel question bank. This project supports educational content preparation by organizing questions, four answer options, and correct-answer indicators into spreadsheet columns.

## Features

- Handles Bangla and English question text and option labels.
- Removes question and option numbering.
- Exports an `MCQs` worksheet with a `1` for the correct option and `0` for other options when the answer is recognized.
- Includes paired Bangla DOCX inputs and XLSX outputs as examples.

## Install and run

Use Python on Windows. Place `requirements.txt` beside `billu_new_v4.0.py`, then run these commands from that folder:

```bash
python -m pip install -r requirements.txt
python billu_new_v4.0.py
```

At the prompts, enter:

1. The input folder's full Windows path, such as `C:\MCQ_Files`, without quotation marks.
2. The input filename without `.docx`.
3. `2` for a separately written answer.
4. `2` to remove source references.

The generated `.xlsx` is saved beside the input with the same base filename. An existing output with that name will be replaced.

## Input format

For this workflow, each question contains six nonempty lines: the question, four options, and an answer. For example:

```text
1. Which number is even?
a. 3
b. 4
c. 5
d. 7
Answer: b
```

Keep this sequence consistent throughout the document. Review the exported spreadsheet for parsing errors.

## Current scope

The supplied version defaults to DOCX input and uses Windows-style paths. It contains a PDF text-extraction branch, but PDF selection is not currently exposed as an active prompt. It does not perform OCR on scanned documents.

The instructions above use separate answers and remove sources. The bold-answer and source-preservation choices need numeric input conversion before those options work as intended.
