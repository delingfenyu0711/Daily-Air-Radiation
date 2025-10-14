# Daily Air Radiation Monitoring Data Scraping Tool

A Python tool for scraping and analyzing air radiation monitoring data, which automatically retrieves radiation monitoring information from specified sources and saves the data as Excel files.

## Features



* Automatically scrape air radiation monitoring data from specified websites

* Support two data parsing methods (HTML tag parsing and text line parsing)

* Implement dual data acquisition mechanism (requests first, Selenium as backup)

* Automatically generate Excel data files with timestamps

* Display acquired data in console table format

## Requirements



* Python 3.7+

* Required dependencies:


  * requests

  * beautifulsoup4

  * pandas

  * fake\_useragent

  * selenium

  * webdriver-manager

  * openpyxl (for Excel file processing)

## Installation Steps



1. Clone or download this project code

2. Install dependency packages:



```
pip install requests beautifulsoup4 pandas fake\_useragent selenium webdriver-manager openpyxl
```

## Usage



1. Run the main program directly:



```
python main.py
```



1. The program will automatically:

* Retrieve radiation monitoring data from the specified URL

* Display parsed data in the console

* Generate an Excel file to save the data (filename format: `辐射监测数据_YYYYMMDD_HHMMSS.xlsx`)

## Code Structure Explanation



* `get_radiation_data(url)`: Retrieve web content, prefer using requests, use Selenium if failed

* `parse_by_html_tags(html_content)`: Parse data through HTML tags (suitable for list structures)

* `parse_by_text_lines(html_content)`: Parse data through text lines (suitable for plain text structures)

* `parse_html(html_content)`: Intelligently select parsing method

* `save_to_excel(data)`: Save data to Excel file

* `display_data(data)`: Display data in the console

* `main()`: Main function, coordinating the work of various modules

## Notes



1. The program has a built-in random delay mechanism to reduce access pressure on the target website

2. If the website structure changes, you may need to adjust the tags and class names in the parsing functions

3. Network connection is required during operation, and the target website must be accessible

4. The first run of Selenium may require downloading ChromeDriver, please ensure network connectivity

## Disclaimer

This tool is for learning and research purposes only, and the data sources are publicly accessible websites. When using this tool, please comply with the robots protocol and relevant regulations of the target website, and reasonably control the access frequency.

> （注：文档部分内容可能由 AI 生成）