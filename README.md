# Excel Data Cleaner

Excel Data Cleaner is a Python script designed to clean, format, and organize Excel datasets efficiently. It handles tasks such as filling missing values, removing duplicates, cleaning column headers, formatting cell borders, centering headers, and removing images from Excel sheets.

### Features

The dataset contains the following features:
-  **Handles Missing Data**: Fills missing numeric values with the column mean and string values with "unknown".
-  **Duplicate Removal**: Automatically identifies and removes duplicate rows.
-  **Header Cleaning**: Cleans up column headers by removing symbols, standardizing names to lowercase, and replacing spaces with underscores.
-  **Date Conversion**: Converts columns with "date" in their name to a datetime format.
-  **Formatting:** Removes borders from all cells, Centers the headers and Standardizes column widths for better readability.
-  **Image Removal**: Removes any embedded images in the Excel sheet.

## How to Run the Project

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/excel-data-cleaner.git
    cd excel-data-cleaner
    ```
2. Run the script:
    ```bash
    python clean_excel_data.py
    ```

## Requirements

- openpyxl
- Pandas

## Contributing

Feel free to submit a pull request or file an issue if you find a bug or want to improve the project.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
