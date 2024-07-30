Here's the documentation for the provided script, which scrapes website data and saves it to an Excel file. The script uses various libraries such as `argparse`, `requests`, `BeautifulSoup`, `pandas`, `lxml`, `selenium`, and `openpyxl`.

## Documentation for `scraper-v3.py`

### Overview

This script scrapes data from web pages listed in sitemaps or specific URLs and saves the extracted information, along with screenshots, into an Excel file. It supports both regular and debug modes.

### Dependencies

Ensure you have the following Python libraries installed:

- `argparse`
- `requests`
- `beautifulsoup4`
- `pandas`
- `lxml`
- `selenium`
- `openpyxl`

You can install these dependencies using `pip`:

```sh
pip install argparse requests beautifulsoup4 pandas lxml selenium openpyxl
```

Additionally, you need to have a web driver for Selenium, such as ChromeDriver, available in your PATH.

### Usage

To run the script, use the following command:

```sh
python scraper-v3.py <sitemap_urls> <output_file> [--debug] [--limit <limit>] [--urls <urls>]
```

#### Arguments

- `sitemap_urls`: One or more URLs of sitemaps to scrape.
- `output_file`: Name of the output Excel file.
- `--debug`: (Optional) Activate debug mode. This limits the number of pages to scrape and/or scrapes specific URLs.
- `--limit`: (Optional) Limit the number of pages to scrape in debug mode (default: 5).
- `--urls`: (Optional) Specific URLs to scrape in debug mode.

### Functions

#### `get_all_pages_from_sitemaps(sitemap_urls)`

Fetches all page URLs listed in the provided sitemaps.

- `sitemap_urls`: List of sitemap URLs.
- Returns: List of page URLs.

#### `sanitize_text(text)`

Cleans and sanitizes text by removing control characters and unnecessary whitespace.

- `text`: Text to sanitize (can be a string or list of strings).
- Returns: Sanitized text.

#### `extract_html_structure(soup)`

Extracts the HTML structure and content from a BeautifulSoup object.

- `soup`: BeautifulSoup object of the webpage.
- Returns: Dictionary containing the title, headers (h1 to h6), paragraphs, links, and link texts.

#### `main(sitemap_urls, output_file, debug_mode=False, debug_limit=5, debug_urls=None)`

Main function that orchestrates the scraping and saving process.

- `sitemap_urls`: List of sitemap URLs.
- `output_file`: Name of the output Excel file.
- `debug_mode`: Boolean flag to activate debug mode.
- `debug_limit`: Limit on the number of pages to scrape in debug mode.
- `debug_urls`: List of specific URLs to scrape in debug mode.

### Workflow

1. **Fetch URLs**:
   - If in debug mode and specific URLs are provided, use those URLs.
   - Otherwise, fetch all page URLs from the provided sitemaps.

2. **Set Up Selenium**:
   - Initialize the Selenium WebDriver to take screenshots.

3. **Scrape Data**:
   - For each URL, fetch the page content, parse it with BeautifulSoup, and extract the HTML structure.
   - Take a screenshot of the page.
   - Save the extracted data and screenshot to an Excel sheet.

4. **Create Table of Contents**:
   - Create a table of contents in the Excel file with hyperlinks to the data for each page.

5. **Save and Quit**:
   - Save the workbook and close the Selenium WebDriver.

### Example

```sh
python scraper-v3.py "https://example.com/sitemap.xml" "output.xlsx" --debug --limit 10 --urls "https://example.com/page1" "https://example.com/page2"
```

This command runs the script in debug mode, limiting the scraping to 10 pages and specifically scraping `page1` and `page2`.

### Notes

- Ensure ChromeDriver or a similar web driver is installed and accessible in your PATH.
- Adjust the script for different web drivers if needed.
- The script assumes that the sitemaps are in XML format and follow the standard sitemap schema.

By following this documentation, users can understand the purpose and usage of the script, as well as its main functions and workflow.
