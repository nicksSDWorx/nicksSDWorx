# Stamkaart Word naar Excel Converter

Extracts employee contract and salary history from a Dutch HR Word document ("stamkaarten") into a structured Excel file.

## Output format

| Naam | Begin contract | Einde contract | Dienstverband | Begindatum | Einddatum | Salaris |
|------|---------------|----------------|---------------|------------|-----------|---------|

- **Contract rows**: Naam + contract columns filled, salary columns empty
- **Salary rows**: Naam + salary columns filled, contract columns empty

## Build the .exe (Windows)

### Prerequisites
- Python 3.10 or newer installed on Windows: https://www.python.org/downloads/
- Make sure to check **"Add Python to PATH"** during installation

### Option A: Use the build script
```
build.bat
```

### Option B: Manual steps
```cmd
pip install python-docx openpyxl pyinstaller
pyinstaller --onefile --windowed --name "StamkaartConverter" app.py
```

The .exe will be in `dist\StamkaartConverter.exe`.

## Usage

1. Double-click `StamkaartConverter.exe`
2. Click **Bladeren...** to select the input Word (.docx) file
3. Click **Bladeren...** to choose where to save the Excel file
4. Click **Uitvoeren** to run the extraction
5. Check the status area for progress

## Debugging

If extraction fails or produces unexpected results, check the log file `stamkaart_debug.log` created next to the .exe. It logs every paragraph and table header encountered.

## Files

- `app.py` — Main application (GUI + Word parser + Excel export)
- `build.bat` — Windows build script
- `requirements.txt` — Python dependencies
- `Voorbeeld import historie.xlsx` — Example of the expected output format
