# WinMOD-Engineering-Assistant

The Engineering Assistant automatically creates WinMOD projects based on existing engineering data.

## PLC Signal Connector App

This Python GUI application loads PLC signal lists and macros from Excel files, filters the signals using rule files, and connects matching macros.

### Features
- Load PLC signals Excel file
- Load macros Excel file
- Load macro rule files to filter PLC signals before matching
- Process and connect signals based on rules and matching logic

### Requirements
- Python 3.x
- pandas
- openpyxl
- tkinter (built-in with Python)

### Installation
1. Install dependencies: `pip install -r requirements.txt`
2. Run the app: `python main.py`

### Console usage
- `python main.py signals.xlsx macros.xlsx output.txt`
- `python main.py signals.xlsx macros.xlsx rules.xlsx output.txt`

### Example files
Sample data is available in the `examples/` folder:
- `examples/signals_sample.xlsx`
- `examples/macros_sample.xlsx`
- `examples/rules_sample.xlsx`
- `examples/README.md`

### Usage
1. Click "Load PLC Signals Excel" to select the signals file.
2. Click "Load Macros Excel" to select the macros file.
3. Click "Load Rule Files" to select one or more rule files.
4. Click "Process and Connect" to filter signals and export matching macros.
