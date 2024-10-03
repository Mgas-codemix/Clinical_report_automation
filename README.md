# Automated Report Generation Script

This Python script automates the process of generating Word reports by extracting specific data from a MultiQC HTML report and tables from an Excel file. The final report includes both text (extracted data) and tables, and it's built using a pre-defined Word template.

## Step-by-Step Explanation

### 1. Argument Parsing:
The script uses `argparse` to handle command-line arguments, allowing the user to specify:
- `-inFile`: The path to the MultiQC HTML report.
- `-template`: The path to the Word template file.
- `-excelfile`: The path to the Excel file that contains additional tables.
- `-outdir`: The directory where the generated Word reports will be saved.

### 2. MultiQC Parsing:
The `parse_multiqc_report()` function uses `BeautifulSoup` to parse the MultiQC HTML file. It looks for specific sections in the HTML (e.g., coverage and quality metrics) and extracts the relevant values. The extracted values are returned in a dictionary, with keys like `'coverage'` and `'quality'`.

### 3. Excel Table Extraction:
The `extract_tables_from_excel()` function reads the entire Excel file using `Pandas` and returns the tables (as DataFrames) for each sheet in the Excel file. If the Excel file contains multiple sheets, each sheet is represented as a separate key in the dictionary returned by this function.

### 4. Inserting Tables into Word:
The `insert_table_from_dataframe()` function inserts a DataFrame into a Word document as a table. It creates a new table in the Word document and populates it with the data from the DataFrame (including headers and rows).

### 5. Replacing Placeholders in the Word Template:
The `replace_placeholders()` function scans through the paragraphs in the Word document and replaces placeholders (e.g., `<<SampleID>>`, `<<CoverageValue>>`) with actual values extracted from the MultiQC report.

### 6. Generating the Final Report:
The `generate_reports_from_multiqc_and_excel()` function creates a final Word report by:
- Replacing placeholders in the Word template with the extracted MultiQC data.
- Inserting tables from the Excel file.
- Saving the report as a `.docx` file in the specified output directory.

### 7. Main Function:
The `main()` function orchestrates the entire process:
- It parses the command-line arguments.
- Extracts data from the MultiQC report and Excel file.
- Generates the final report using the Word template.


