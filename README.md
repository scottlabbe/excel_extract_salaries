# Salary Data Extractor

A Python Jupyter notebook for consolidating salary data from multiple Excel files into a single dataset.

## Project Structure
- `notebook/`: Contains the main Python scripts
- `data/`: Place input Excel files here

## Setup
1. Clone the repository
```bash
git clone [repository-url]
cd salary-data-extractor
```

2. Install required packages
```bash
pip install -r requirements.txt
```

## Usage
1. Place your Excel files in the `data/` directory
2. Run the scripts in the notebook cells
```
3. Find the consolidated output in `output/compiled_data.xlsx`

## Input File Requirements
- Excel files must have two sheets:
  - 'Input Data' sheet with headers in A4:A5 and values in B4:B5
  - 'Salaries' sheet with headers in A3:F3 and data starting from row 4

## Security Notice
⚠️ This repository is designed to process salary data. Never commit actual salary data or sensitive information to this repository. Make sure to check files carefully before committing.