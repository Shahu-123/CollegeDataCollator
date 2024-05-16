# College Data Extraction and Insertion

This project extracts specific information about various colleges using the Google Custom Search API and OpenAI's GPT-3.5-turbo, then inserts this information into an Excel sheet.

## Prerequisites

Make sure you have the following installed:
- Python 3.x
- `openpyxl` library
- `requests` library
- `openai` library

Install the required libraries using pip:
```bash
pip install openpyxl requests openai
```

## Configuration

Replace the placeholders with your actual API keys and file paths in the script:
- 'google_api_key': Your Google Custom Search API key.
- 'cx': Your Custom Search Engine ID.
- 'openai_api_key': Your OpenAI API key. 
- 'filename': Path to your Excel file.

## Usage

The main functions in the script are:

### get_search_results(query, api_key, cx)

Sends a request to the Google Custom Search API and returns the search results as a JSON object.

### print_snippets(search_results)

Extracts snippets from the search results and combines them into a single string.

### read_cells_from_excel(file_path, sheet_name, cell_range='A1:A52')

Reads cells from an Excel file and returns their values as a list of strings.

### extract_specific_info_from_chatgpt(text, required_info, api_key)

Sends text to OpenAI's GPT-3.5-turbo to extract specific information.

### write_to_excel(data_dict, file_path, sheet_name)

Writes data from a dictionary into the specified Excel sheet.