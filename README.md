# Simple_NLP_Analysis
This script measures lexical and syntactic development in writing excerpts across time, using three simple linguistics metrics.

The repository contains a Python script that analyzes student writing files (plain text) and generates per-student Word reports (.docx) with several linguistic measures.

Features
- Cleans and filters text.
- Computes Type-Token Ratio (TTR), Moving Average TTR (MATTR), and lexical density.
- Estimates Dependent Clause Ratio (DCR).
- Batch analyzes a student folder of .txt files.
- Produces a Word report for each student.
- Entrypoint that scans the 'mydata' directory and runs the pipeline.

Requirements
- Python 3.12
- spaCy 3.8.7 and an English model ( `en_core_web_sm`)
- python-docx 1.2.0
- pyspellchecker 0.8.3

Install (recommended inside a virtual environment)

# create & activate venv (Windows example)
py -m venv .myenv
.myenv\Scripts\activate

# upgrade packaging tools
py -m pip install --upgrade pip setuptools wheel

# install deps
py -m pip install -U spacy python-docx pyspellchecker
py -m spacy download en_core_web_sm
py -m spacy validate

_________________________________________________

Quick usage

Create a folder named "mydataé at the project root.
Inside mydata, create one folder per student. Put each writing as a .txt file. Filenames are used as the date field in reports.
Example: mydata/Student8/Week1.txt

Run the script:
python index.py 

Output

Generates a {student_name}_analysis_report.docx per student in the working directory. Report creation is handled by index.create_student_report.

Notes

The filtering stage uses index.filter_text and pyspellchecker; this may remove uncommon or domain-specific words.
Clause detection is heuristic and based on spaCy dependency labels in index.calculate_dcr. Expect edge cases with complex syntax.
The script assumes text files are UTF-8 encoded.

If you want to modify behavior:

Change cleaning rules in index.filter_text.
Adjust MATTR window size by passing a different segment_length to index.calculate_ttr_spacy.
