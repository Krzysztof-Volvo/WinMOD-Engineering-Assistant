# Example Data for WinMOD Engineering Assistant

This folder contains sample Excel files you can use to try the PLC Signal Connector app.

Files:
- `signals_sample.xlsx` - sample PLC signal list with `name`, `type`, `address`, and `comment`.
- `macros_sample.xlsx` - sample macro definitions with `Symbol`, `MacroName`, `Input`, and `Output`.
- `rules_sample.xlsx` - sample macro rule definitions. Supports both legacy `field/operator/value` rules and the new `macro signal`/`rule` style.

Usage examples:
- GUI: load `signals_sample.xlsx`, `macros_sample.xlsx`, and `rules_sample.xlsx`.
- Console: `python main.py examples/signals_sample.xlsx examples/macros_sample.xlsx examples/rules_sample.xlsx output.txt`
