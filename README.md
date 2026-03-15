# MSR Tape Cracking

Python scripts for parsing and processing Mortgage Servicing Rights (MSR) data tapes. Automates the extraction, cleaning, and summarization of raw loan-level data to accelerate portfolio analysis workflows.

## Overview

MSR tapes are large, often inconsistently formatted data files containing loan-level detail on mortgage servicing portfolios. Manually processing these files is time-intensive and error-prone. These scripts automate the ingestion and transformation process — handling raw tape data and producing clean, structured outputs ready for analysis or reporting.

## Features

- Parses raw MSR tape files (100K+ rows) into structured DataFrames
- Cleans and standardizes inconsistent field formats across tape sources
- Generates summary statistics and aggregations for portfolio-level analysis
- Outputs processed data to Excel for stakeholder reporting

## Files

| File | Description |
|------|-------------|
| `TapeCracking.py` | Initial version — core parsing and summarization logic |
| `TapeCracking2.0.py` | Enhanced version with improved field handling and output formatting |

## Dependencies

```
pandas
openpyxl
```

Install via pip:
```bash
pip install pandas openpyxl
```

## Usage

1. Place your raw MSR tape file in the working directory
2. Update the file path variable at the top of the script
3. Run the script — processed output will be saved as an Excel file in the same directory

## Context

Built during work at Pennymac Mortgage Trust to support recurring MSR portfolio analysis. Reduced manual data processing time by over 60% on tape files exceeding 100,000 rows.

## Author

**Emmanuel Lopez**
[LinkedIn](https://www.linkedin.com/in/lopez-emmanuel/) · [GitHub](https://github.com/elopez567)
