# CIS Benchmark Converter

**Author:** Octomany  
**LinkedIn:** [LinkedIn](https://www.linkedin.com/in/maxbeauchamp/)  

**Date Created:** 2024-11-06  
**Last Update:** 2025-14-03

## Description

`CIS Benchmark Converter` is a Python script that extracts recommendations from CIS Benchmark PDF documents and exports them into CSV, Excel, or JSON formats. The script converts unstructured PDF content into a structured table, simplifying compliance reviews and audits.

## Features

- **Configurable Extraction:**  
  Set the start page to skip tables of contents or disclaimers, and adjust the logging level via command-line options.

- **Multiple Output Formats:**  
  Export the extracted data as CSV (using a pipe `|` delimiter), Excel (with styled headers, dropdowns, and conditional formatting), or JSON for easy data integration.

- **Robust and Maintainable:**  
  Uses `pathlib` for modern file path management, extended type annotations for static type checking, enhanced exception handling, and a progress bar (via `tqdm`) for user feedback.

## Installation

1. Clone this repository.
2. Install dependencies using the provided `requirements.txt`:

   ```bash
   pip install -r requirements.txt
   ```

   **requirements.txt:**
   ```
   pdfplumber
   openpyxl
   tqdm
   ```

## Usage

Run the script from the command line as follows:

```bash
python cis_benchmark_converter.py \
  -i path/to/input_file.pdf \
  -o path/to/output_file \
  -f [csv|excel|json] \
  --start_page 10 \
  --log_level INFO
```

### Arguments

- `-i, --input` : Path to the input CIS Benchmark PDF file (required).
- `-o, --output` : Path to the output file (defaults to the input file name with a `.csv`, `.xlsx`, or `.json` extension).
- `-f, --format` : Output file format: `csv`, `excel`, or `json` (default: `excel`).
- `--start_page` : Page number to start extraction (default: 10).
- `--log_level` : Logging level (`DEBUG`, `INFO`, `WARNING`, or `ERROR`; default: `INFO`).

## Example

```bash
python cis_benchmark_converter.py -i ./CIS_AWS_Benchmark.pdf -o ./CIS_AWS_Benchmark.json -f json --start_page 10 --log_level INFO
```

For JSON output, the data is structured as a list of dictionaries, with each dictionary representing a recommendation and its associated sections.

> **Note:** The script automatically excludes any sections labeled "CIS Controls" to focus solely on the core recommendations.

## Acknowledgements

Special thanks to [Flavien Fouqueray (UnBonWhisky)](https://www.linkedin.com/in/ffouqueray/) for his valuable bug fixes and contributions in earlier versions of this script.

## License

This project is licensed under the MIT License. Please respect the copyrights
of the CIS Benchmark documents when using and sharing this tool.