# PDF to MS Word Automation – Internship Assignment

## Overview
This project programmatically recreates a provided legal PDF document into a Microsoft Word (.docx) file using Python. The objective is to match the PDF’s layout, formatting, spacing, alignment, headings, and table structure as accurately as possible.

## Approach
1. Manually analyzed the PDF to understand layout, spacing, and structure.
2. Identified that the document is table-based and requires merged cells.
3. Used the `python-docx` library to construct the Word document from scratch.
4. Implemented tables, merged rows, controlled column widths, and formatted text runs.
5. Used multiple paragraphs inside table cells to prevent text clipping issues.
6. Preserved template placeholders exactly as shown in the PDF.

## Technologies Used
- Python 3.x
- python-docx
- Flask (for deployment)

## Features
- Accurate table-based layout replication
- Correct bold headings and spacing
- Fully automated Word document generation
- No manual editing of the output file
