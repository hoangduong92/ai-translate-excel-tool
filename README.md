# AI Excel Translation Tool

This Python script translates text within specified cells of an Excel spreadsheet using the Google Gemini API. It supports translation between Japanese and Vietnamese and offers various options for batching, concurrency, and selective translation.

## Prerequisites

1.  **Python 3.7+**
2.  **Google Gemini API Key**: You must have a `GEMINI_API_KEY` environment variable set.
    ```bash
    export GEMINI_API_KEY="YOUR_API_KEY_HERE"
    ```
    Alternatively, you can create a `.env` file in the project root directory with the following content:
    ```
    GEMINI_API_KEY=YOUR_API_KEY_HERE
    ```
3.  **Required Python Packages**: Install them using pip:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

The script is run from the command line. Here's the basic syntax:

```bash
python trans-tool.py --file_path <INPUT_EXCEL_FILE> --sheet_name <SHEET_TO_TRANSLATE> [OPTIONS]
```

### Key Command-Line Arguments

*   `--file_path` (Required): Path to the input Excel file (e.g., `data/my_data.xlsx`).
*   `--sheet_name` (Required): Name of the sheet within the Excel file to translate (e.g., `Sheet1`).
*   `--direction` (Optional): Translation direction. 
    *   `to_japanese` (Default): Translates from the detected language (typically Vietnamese or English) to Japanese.
    *   `to_vietnamese`: Translates from the detected language (typically Japanese or English) to Vietnamese.
    *   Example: `--direction to_vietnamese`
*   `--output_path` (Optional): Path to save the translated Excel file. If not provided, it defaults to `[input_filename]_translated.xlsx` in the same directory as the input file (e.g., `data/my_data_translated.xlsx`).
*   `--columns` (Optional): Specify which columns to translate. Provide a space-separated list of column letters (e.g., `A C E`). If not provided, the script attempts to translate all cells with translatable text.
*   `--start_row` (Optional): The row number to start translation from (1-indexed). Defaults to 1.
*   `--end_row` (Optional): The row number to end translation at (inclusive). If not provided, translates to the last row with data.

### Performance and Batching Arguments

*   `-b, --batch_size` (Optional): Number of items (cells or rows) to group together for a single API translation request. Default: `20`.
*   `-a, --api_delay` (Optional): Delay in seconds between consecutive API calls to avoid rate limiting. Default: `2.0`.
*   `--workers` (Optional): Number of concurrent worker threads for making API calls. Default: `3`.
*   `--group_by_row` (Optional): When this flag is present, the script groups all specified cells within a single row into one item for translation. This can provide better context for the API. If not set, each cell is treated as an individual item. `BATCH_SIZE` will then refer to the number of rows processed in a batch.

### Other Arguments

*   `--log_file` (Optional): Path to the log file. Defaults to `translation_log_[timestamp].log`.
*   `--cache_file` (Optional): Path to the cache file for storing translations (currently, caching logic seems to be commented out or partially implemented in the script).
*   `-h, --help`: Show the help message and exit.

## Examples

1.  **Translate Sheet1 of `input.xlsx` from Vietnamese to Japanese (default direction):**
    ```bash
    python trans-tool.py --file_path data/input.xlsx --sheet_name Sheet1
    ```

2.  **Translate columns A and D of `Sheet1` in `source_data.xlsx` from Japanese to Vietnamese, saving to `translated_output.xlsx`:**
    ```bash
    python trans-tool.py --file_path source_data.xlsx --sheet_name Sheet1 --columns A D --direction to_vietnamese --output_path results/translated_output.xlsx
    ```

3.  **Translate rows 10 to 50 in `Sheet2` of `complex_data.xlsx` to Japanese, grouping by row for context, using 5 workers and a batch size of 10 rows:**
    ```bash
    python trans-tool.py --file_path complex_data.xlsx --sheet_name Sheet2 --start_row 10 --end_row 50 --group_by_row --workers 5 --batch_size 10
    ```

## Logging

The script generates a log file (e.g., `translation_log_YYYYMMDD_HHMMSS.log`) in the script's directory for each run. This log contains detailed information about the translation process, including API calls, successes, and any errors encountered.
