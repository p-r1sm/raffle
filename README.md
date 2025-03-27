# Laabharti Card Generator

A Python application that generates cards from CSV data in a Word document, formatted with 8 cards per A4 page (2 columns and 4 rows).

## Features

- Reads beneficiary information from a CSV file
- Generates professional-looking cards with golden borders and styling
- Formats cards in a 2x4 grid on A4 paper (8 cards per page)
- Supports custom number of rows and columns
- Allows using a custom logo image for the cards
- Consistent golden text color for both headings and data fields
- Graphical user interface for easy operation

## Requirements

- Python 3.6 or higher
- Required packages (install with `pip install -r requirements.txt`):
  - python-docx
  - pandas
  - pillow

## Setup

1. Clone this repository or download the files
2. Install the required packages:

```bash
pip install -r requirements.txt
```

## Usage

### Quick Start (Recommended)

Simply run the launcher script which will check for dependencies and start the application:

```bash
# On Windows
python run.py

# On Mac/Linux
python3 run.py
```

### Graphical User Interface

You can also directly start the GUI:

```bash
python gui.py
```

The GUI allows you to:
- Select a CSV file
- Choose where to save the output document
- Provide your own logo image for the cards
- Adjust the number of rows and columns per page 
- Automatically convert data from different CSV formats
- Monitor progress with a status indicator

### Command Line Usage

You can also use the command line interface:

#### Basic Usage

```bash
python generate_cards.py
```

This will read data from the default `data/sample_data.csv` file and create a Word document `output_cards.docx` with the cards.

#### Custom Options

```bash
python generate_cards.py --csv your_data.csv --output your_output.docx --rows 4 --cols 2 --logo your_logo.png
```

Parameters:
- `--csv`: Path to your CSV file (default: data/sample_data.csv)
- `--output`: Path for the output Word document (default: output_cards.docx)
- `--rows`: Number of rows per page (default: 4)
- `--cols`: Number of columns per page (default: 2)
- `--logo`: Path to a custom logo image to use on the cards

## CSV Format

Your CSV file should contain the following columns:
- LAABHARTHI_NAME: Name of the beneficiary
- CONTACT_NUMBER: Contact phone number
- ARPIT_GROUP: Group name (if applicable)
- AREA: Location/area

Example:
```
LAABHARTHI_NAME,CONTACT_NUMBER,ARPIT_GROUP,AREA
"John Doe","9876543210","Group A","Delhi"
```

### Converting Your Data

If your existing CSV file doesn't match the required format, you can use the provided utility script to convert it:

```bash
python convert_data.py your_data.csv converted_data.csv
```

This utility will:
- Try to map your existing columns to the required format
- Prompt you for mappings if it can't determine them automatically
- Add empty columns for any missing required fields
- Save the converted data in the proper format

Alternatively, check the "Auto-convert CSV format" option in the GUI.

## Output

The generated Word document contains cards with:
- The beneficiary's name with a golden border
- Contact number with a golden border
- Arpit Group with a golden border
- Area with a golden border
- Custom logo image or a default meditation silhouette
- "CELEBRATING 125 YEARS" text
- "SHRIMAD RAJCHANDRAJI Gracing Dharampur" text
- "Amount Rs. 1000/-" text 