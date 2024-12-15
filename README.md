# PowerPoint Text Extractor Script

This Python script extracts text content from `.pptx` PowerPoint presentations and saves it as plain text files. It is useful for automating the extraction of slide text for analysis or archival purposes.

## Features
- Extracts XML content from `.pptx` files by converting them into ZIP format.
- Parses slide content to extract textual elements using BeautifulSoup.
- Saves text content from each slide into organized plain text files.
- Processes multiple `.pptx` files in a specified directory.

## Prerequisites
Ensure you have the following installed:

- Python 3.7 or later
- Required Python libraries:
  - `BeautifulSoup4`
  - `zipfile`
  - `glob`

Install the required libraries using pip if not already installed:
```bash
pip install beautifulsoup4
```

## Directory Structure
Place this script in the following directory structure:

```
project_root/
├── script.py  # The main script file
├── ppts/      # Directory containing .pptx files
├── extracts/  # Directory where extracted XML content will be stored
└── outputs/   # Directory where text files will be saved
```

## How to Use

1. Place your `.pptx` files in the `ppts` directory.
2. Run the script from the project root:
   ```bash
   python script.py
   ```
3. The extracted XML content will be stored in the `extracts` directory.
4. Text files containing slide content will be saved in the `outputs` directory.

## Output Format
Each output text file will be named based on the corresponding `.pptx` file (e.g., `output_ppt1.txt`) and will contain the text from all slides. Each slide's content is separated by a header:

```
--- Slide 1 ---
Slide content here

--- Slide 2 ---
Slide content here
```

## Limitations
- This script only extracts text content from slides and does not include other elements like images or formatting.
- It assumes that `.pptx` files are well-formed and follow the standard PowerPoint structure.

## License
This project is licensed under the MIT License. Feel free to use, modify, and distribute it.

## Contribution
If you have suggestions or want to improve the script, feel free to create a pull request or open an issue on GitHub.

## Acknowledgements
- Built with BeautifulSoup for XML parsing.
- Inspired by the need to automate PowerPoint text extraction for bulk analysis.

