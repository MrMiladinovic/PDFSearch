
---

# PDF Keyword Search

## Overview

`PDF Keyword Search` is a Python application that allows you to search for specific keywords within PDF files located in a selected directory. This application features a graphical user interface (GUI) built with Tkinter and uses the PyPDF2 library to handle PDF files. You can perform case-sensitive or case-insensitive searches and view the results in a table format.

## Features

- **Directory Selection**: Choose the directory containing the PDF files to search through.
- **Keyword Search**: Input a keyword to search for within the selected PDF files.
- **Case Sensitivity**: Toggle case sensitivity for the keyword search.
- **Progress Tracking**: View the progress of the search operation.
- **Results Display**: See the filenames, file paths, and pages where the keyword was found.
- **Context Menu**: Right-click to copy the file path of the search results.

## Requirements

- Python 3.x
- PyPDF2
- Tkinter (usually included with Python standard library)

## Installation

### Option 1: Using Precompiled Release

You can download a precompiled release of the application from the [Releases](https://github.com/MrMiladinovic/PDFSearch/releases) page. This allows you to use the application without having to compile it from source.

1. **Download the Release**:
   Go to the [Releases](https://github.com/MrMiladinovic/PDFSearch/releases) page and download the latest release suitable for your operating system.

2. **Extract and Run**:
   Extract the downloaded file and follow the instructions included in the release notes to run the application.

### Option 2: Compile from Source

1. **Clone the Repository**:
    ```bash
    git clone https://github.com/MrMiladinovic/PDFSearch.git
    cd pdf-keyword-search
    ```

2. **Install Dependencies**:
    You can install the required Python packages using `pip`:
    ```bash
    pip install PyPDF2
    ```

## Usage

1. **Run the Application**:
    Execute the script to start the GUI application:
    ```bash
    python PDFSearch.py
    ```

2. **Select Directory**:
    Click the "Browse" button to select the directory containing your PDF files.

3. **Enter Keyword**:
    Type the keyword you want to search for in the "Keyword" entry field.

4. **Set Case Sensitivity**:
    Check or uncheck the "Case Sensitive" box based on your search needs.

5. **Start Search**:
    Click the "Search" button to begin the search operation. The progress bar will indicate the search status.

6. **View Results**:
    Results will be displayed in the table with columns for File Name, File Path, and Pages Found.

7. **Copy File Path**:
    Right-click on any result row to copy the file path to the clipboard.

## Contributing

Contributions are welcome! If you have suggestions, improvements, or fixes, please fork the repository and submit a pull request. For major changes, please open an issue to discuss them first.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.


## Acknowledgments

- **PyPDF2**: For handling PDF operations.
- **Tkinter**: For building the GUI components.

---
