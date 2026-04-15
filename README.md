# WinMOD-Engineering-Assistant

The Engineering Assistant automatically creates WinMOD projects based on existing engineering data.

## PLC Signal Connector App

This Python GUI application loads PLC signal lists and macros from Excel files, filters the signals using rule files, and connects matching macros.

### Features
- Load PLC signals Excel file
- Load macros Excel file
- Load a separate rules Excel file to filter PLC signals before matching
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

### Rules file format
The app now supports two rule file styles:

1. Legacy filter rules:
   - `field`, `operator`, `value`, optional `active`
   - Filters the signal list before matching macros.

2. Macro rule format:
   - `macro type`, `macro signal`, `rule`
   - Optional columns: `field`, `active`
   - Default `field` is `name`.
   - Example rule syntax:
     - `*=F9-SGC` matches signals whose `name` contains `F9-SGC`
     - `name*=F9-SGC` also matches by `name`
     - `!=F9-SGC` means not equals
     - `F9-SGC` means contains `F9-SGC` in the selected field

Each macro rule row tells the app how to find the correct signal for the macro symbol.

### Usage
1. Click "Load PLC Signals Excel" to select the signals file.
2. Click "Load Macros Excel" to select the macros file.
3. Click "Load Rules Excel File" to select one or more external rule files.
4. Click "Process and Connect" to apply rules and export matching macros.
