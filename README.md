# PubMed Data Extraction Script

This repository contains a Python script to search and extract PubMed records based on specific authors and topics. The extracted data is stored in a Pandas DataFrame and saved to an Excel file.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [Customization](#customization)
- [Output](#output)
- [License](#license)

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/TLDWTutorials/PubMed-Data-Extraction.git
    cd PubMed-Data-Extraction
    ```

2. Install the required dependencies:
    ```sh
    pip install pandas biopython
    ```

## Usage

1. Update the email address to your own to avoid potential issues with Entrez:
    ```python
    Entrez.email = 'your.email@example.com'
    ```

2. Customize the list of authors and topics as needed:
    ```python
    authors = ['Bryan Holland', 'Mehmet Oz', 'Anthony Fauci']
    topics = ['RNA', 'cardiovascular']
    ```

3. Run the script:
    ```sh
    python pubmed_extraction.py
    ```

## Customization

- **Authors**: Modify the `authors` list with the names of authors you want to include in the search.
- **Topics**: Modify the `topics` list with the topics you want to include in the search.
- **Date Range**: Adjust the `date_range` variable to the desired date range for your search.

## Output

The script will create an Excel file named `PubMed_results.xlsx` containing the following columns:
- PMID
- Title
- Abstract
- Authors
- Journal
- Keywords
- URL
- Affiliations

## License

This project is licensed under the MIT License.
