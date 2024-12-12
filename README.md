# Web Scraping e-Books and Saving as Word Documents

This project automates the process of scraping e-book content from a specified website and saving the extracted data as a well-formatted Word document. It uses Selenium and BeautifulSoup to extract text content, cleans it to remove unnecessary elements, and applies appropriate formatting to ensure readability.

## Features
- Extracts text content (headings and paragraphs) from specified pages of an e-book.
- Cleans the extracted content to remove unwanted tags and duplicate entries.
- Formats the text with styles (bold, italic, underline, highlights, etc.) based on the original HTML structure.
- Saves the cleaned and formatted content into a Word document (.docx).

## Prerequisites
- Python 3.x
- Google Chrome and ChromeDriver
- Required Python libraries:
  - selenium
  - bs4 (BeautifulSoup)
  - python-docx

## Installation
1. Clone this repository or download the script.
2. Install the required libraries using pip:
   ```bash
   pip install selenium beautifulsoup4 python-docx
   ```
3. Download and set up ChromeDriver for Selenium. Ensure it matches your installed version of Google Chrome.

## Usage
1. Run the script:
   ```bash
   python ebook_scraper.py
   ```
2. Enter the number of pages you want to scrape when prompted.
3. The script will navigate through the specified pages, extract the content, and save it as a Word document named `cleaned_extracted_content_with_headings.docx` in the current directory.

## Code Highlights
- **Dynamic Web Content Handling:** Uses Selenium to load and interact with dynamic web pages.
- **Text Cleaning:** Removes unnecessary elements like `[if !supportLists]` and `[endif]` from the extracted text.
- **Text Formatting:** Applies appropriate styles to text based on its HTML tags (e.g., bold, italic, highlight).

## Example Output
The extracted content is saved as a Word document with:
- Properly formatted headings and paragraphs.
- Clean and readable text.
- Styles that reflect the original HTML formatting.

## Notes
- Ensure that the target website allows scraping and respects its terms of service.
- Update the website URL in the script to match the e-book you want to scrape.

## Signature
- **User Name:** MMO23
- **E-mail:** mostafa.mamdouh.774@gmail.com

---
Feel free to use or modify this script for educational purposes or personal projects. For inquiries or collaboration, contact me via the email above.

