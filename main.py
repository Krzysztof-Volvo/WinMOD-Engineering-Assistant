import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import sys

class PLCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PLC Signal Connector")

        self.signals_df = None
        self.macros_df = None
        self.rules_df = None
        self.rule_files = []

        self.load_signals_btn = tk.Button(root, text="Load PLC Signals Excel", command=self.load_signals)
        self.load_signals_btn.pack(pady=4)

        self.load_macros_btn = tk.Button(root, text="Load Macros Excel", command=self.load_macros)
        self.load_macros_btn.pack(pady=4)

        self.load_rules_btn = tk.Button(root, text="Load Rule Files", command=self.load_rules)
        self.load_rules_btn.pack(pady=4)

        self.process_btn = tk.Button(root, text="Process and Connect", command=self.process, state=tk.DISABLED)
        self.process_btn.pack(pady=8)

        self.file_label_frame = tk.Frame(root)
        self.file_label_frame.pack(pady=8)
        self.signals_label = tk.Label(self.file_label_frame, text="Signals: not loaded")
        self.signals_label.grid(row=0, column=0, sticky="w", padx=4)
        self.macros_label = tk.Label(self.file_label_frame, text="Macros: not loaded")
        self.macros_label.grid(row=1, column=0, sticky="w", padx=4)
        self.rules_label = tk.Label(self.file_label_frame, text="Rules: none loaded")
        self.rules_label.grid(row=2, column=0, sticky="w", padx=4)

        self.editor_frame = tk.Frame(root)
        self.editor_frame.pack(pady=10)
        self.add_macro_btn = tk.Button(self.editor_frame, text="Add Macro", command=self.add_macro)
        self.add_macro_btn.pack(side=tk.LEFT, padx=5)
        self.edit_macro_btn = tk.Button(self.editor_frame, text="Edit Selected Macro", command=self.edit_macro)
        self.edit_macro_btn.pack(side=tk.LEFT, padx=5)
        self.save_macros_btn = tk.Button(self.editor_frame, text="Save Macros Excel", command=self.save_macros)
        self.save_macros_btn.pack(side=tk.LEFT, padx=5)

        self.macro_listbox = tk.Listbox(root, width=120, height=8)
        self.macro_listbox.pack(pady=5)

        self.rule_list_label = tk.Label(root, text="Loaded rule files:")
        self.rule_list_label.pack()
        self.rule_listbox = tk.Listbox(root, width=120, height=5)
        self.rule_listbox.pack(pady=5)
        self.refresh_rule_listbox()

        self.output_text = tk.Text(root, height=16, width=80)
        self.output_text.pack(pady=10)

    def load_signals(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        try:
            self.signals_df = pd.read_excel(file_path)
            self.signals_label.config(text=f"Signals: {file_path}")
            self.output_text.insert(tk.END, f"Loaded signals: {len(self.signals_df)} rows\n")
            self.output_text.insert(tk.END, f"Signal columns: {list(self.signals_df.columns)}\n\n")
            self.check_enable_process()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load signals: {e}")

    def load_macros(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        try:
            self.macros_df = pd.read_excel(file_path)
            self.macros_label.config(text=f"Macros: {file_path}")
            self.output_text.insert(tk.END, f"Loaded macros: {len(self.macros_df)} rows\n")
            self.output_text.insert(tk.END, f"Macro columns: {list(self.macros_df.columns)}\n\n")
            self.refresh_macro_listbox()
            self.check_enable_process()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load macros: {e}")

    def load_rules(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_paths:
            return
        new_rules = []
        for path in file_paths:
            try:
                rules_page = pd.read_excel(path)
                new_rules.append(rules_page)
                self.rule_files.append(path)
                self.output_text.insert(tk.END, f"Loaded rule file: {path} ({len(rules_page)} rows)\n")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load rule file {path}: {e}")
                return
        if self.rules_df is None:
            self.rules_df = pd.concat(new_rules, ignore_index=True)
        else:
            self.rules_df = pd.concat([self.rules_df] + new_rules, ignore_index=True)
        self.rules_label.config(text=f"Rules: {len(self.rule_files)} file(s) loaded")
        self.output_text.insert(tk.END, f"Total rules rows: {len(self.rules_df)}\n")
        self.output_text.insert(tk.END, f"Rule columns: {list(self.rules_df.columns)}\n\n")
        self.refresh_rule_listbox()
        self.check_enable_process()

    def refresh_rule_listbox(self):
        self.rule_listbox.delete(0, tk.END)
        if self.rule_files:
            for path in self.rule_files:
                self.rule_listbox.insert(tk.END, path)
        else:
            self.rule_listbox.insert(tk.END, "No rule files loaded")

    def refresh_macro_listbox(self):
        self.macro_listbox.delete(0, tk.END)
        if self.macros_df is None:
            return
        for idx, row in self.macros_df.iterrows():
            self.macro_listbox.insert(tk.END, f"{idx + 1}: {row.to_dict()}")

    def check_enable_process(self):
        if self.signals_df is not None and self.macros_df is not None and self.rules_df is not None:
            self.process_btn.config(state=tk.NORMAL)

    def process(self):
        self.output_text.insert(tk.END, "Processing signals and connecting to macros...\n")
        signal_col = 'name'
        symbol_col = 'Symbol'
        if self.rules_df is None:
            self.output_text.insert(tk.END, "No rules loaded.\n")
            return
        if self.signals_df is None or self.macros_df is None:
            self.output_text.insert(tk.END, "Load all files first.\n")
            return
        if signal_col not in self.signals_df.columns or symbol_col not in self.macros_df.columns:
            self.output_text.insert(tk.END, "Required columns not found in the loaded files.\n")
            return
        filtered_signals = apply_signal_rules(self.signals_df, self.rules_df, output_consumer=lambda msg: self.output_text.insert(tk.END, msg))
        if filtered_signals.empty:
            self.output_text.insert(tk.END, "No signals remain after applying rules.\n")
            return
        self.output_text.insert(tk.END, f"Signals after filtering: {len(filtered_signals)} rows\n")
        matched_macros = self.macros_df[self.macros_df[symbol_col].isin(filtered_signals[signal_col])]
        if matched_macros.empty:
            self.output_text.insert(tk.END, "No matching macros found after filtering.\n")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")], title="Save result as...")
        if not file_path:
            self.output_text.insert(tk.END, "Export cancelled.\n")
            return
        try:
            matched_macros.to_csv(file_path, sep=",", index=False)
            self.output_text.insert(tk.END, f"Exported {len(matched_macros)} rows to {file_path}\n")
        except Exception as e:
            self.output_text.insert(tk.END, f"Failed to export: {e}\n")

    def add_macro(self):
        if self.macros_df is None:
            messagebox.showerror("Error", "Load a macros Excel file first.")
            return
        columns = list(self.macros_df.columns)
        new_values = []
        for col in columns:
            val = simple_input_dialog(self.root, f"Enter value for '{col}':")
            if val is None:
                return
            new_values.append(val)
        self.macros_df.loc[len(self.macros_df)] = new_values
        self.refresh_macro_listbox()
        self.output_text.insert(tk.END, "Added new macro.\n")

    def edit_macro(self):
        if self.macros_df is None:
            messagebox.showerror("Error", "Load a macros Excel file first.")
            return
        selection = self.macro_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Select a macro to edit.")
            return
        idx = selection[0]
        row_idx = idx
        columns = list(self.macros_df.columns)
        for i, col in enumerate(columns):
            old_val = str(self.macros_df.iloc[row_idx, i])
            val = simple_input_dialog(self.root, f"Edit value for '{col}':", old_val)
            if val is None:
                return
            self.macros_df.iat[row_idx, i] = val
        self.refresh_macro_listbox()
        self.output_text.insert(tk.END, "Edited macro.\n")

    def save_macros(self):
        if self.macros_df is None:
            messagebox.showerror("Error", "No macros to save.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save macros as...")
        if not file_path:
            self.output_text.insert(tk.END, "Save cancelled.\n")
            return
        try:
            self.macros_df.to_excel(file_path, index=False)
            self.output_text.insert(tk.END, f"Macros saved to {file_path}\n")
        except Exception as e:
            self.output_text.insert(tk.END, f"Failed to save macros: {e}\n")


def simple_input_dialog(root, prompt, initial_value=""):
    dialog = tk.Toplevel(root)
    dialog.title("Input")
    tk.Label(dialog, text=prompt).pack(padx=10, pady=5)
    entry = tk.Entry(dialog, width=50)
    entry.pack(padx=10, pady=5)
    entry.insert(0, initial_value)
    result = {"value": None}

    def on_ok():
        result["value"] = entry.get()
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=5)
    tk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="Cancel", command=on_cancel).pack(side=tk.LEFT, padx=10)
    dialog.grab_set()
    root.wait_window(dialog)
    return result["value"]


def apply_signal_rules(signals_df, rules_df, output_consumer=None):
    if rules_df is None or rules_df.empty:
        return signals_df
    if not all(col in rules_df.columns for col in ['field', 'operator', 'value']):
        if output_consumer:
            output_consumer("Rule file must contain 'field', 'operator', and 'value' columns.\n")
        return signals_df
    mask = pd.Series(True, index=signals_df.index)
    for _, rule in rules_df.iterrows():
        active = str(rule.get('active', 'yes')).strip().lower()
        if active in ['no', 'false', '0', 'n']:
            continue
        field = str(rule.get('field', '')).strip()
        operator = str(rule.get('operator', 'equals')).strip().lower()
        value = rule.get('value')
        if not field:
            continue
        if field not in signals_df.columns:
            if output_consumer:
                output_consumer(f"Skipping rule because signal column '{field}' is not present.\n")
            continue
        mask &= evaluate_rule(signals_df[field], operator, value)
    return signals_df[mask]


def evaluate_rule(series, operator, value):
    if operator in ['=', '==', 'equals']:
        return series.astype(str) == str(value)
    if operator in ['!=', 'not equals', 'not equal']:
        return series.astype(str) != str(value)
    if operator == 'contains':
        return series.astype(str).str.contains(str(value), na=False)
    if operator == 'startswith':
        return series.astype(str).str.startswith(str(value), na=False)
    if operator == 'endswith':
        return series.astype(str).str.endswith(str(value), na=False)
    if operator == 'in':
        values = [v.strip() for v in str(value).split(',') if v.strip()]
        return series.astype(str).isin(values)
    if operator == 'not in':
        values = [v.strip() for v in str(value).split(',') if v.strip()]
        return ~series.astype(str).isin(values)
    if operator == 'regex':
        return series.astype(str).str.contains(str(value), na=False, regex=True)
    if operator in ['>', 'gt', 'greater than']:
        return pd.to_numeric(series, errors='coerce') > float(value)
    if operator in ['<', 'lt', 'less than']:
        return pd.to_numeric(series, errors='coerce') < float(value)
    if operator in ['>=', 'ge', 'greater or equal', 'greater than or equal']:
        return pd.to_numeric(series, errors='coerce') >= float(value)
    if operator in ['<=', 'le', 'less or equal', 'less than or equal']:
        return pd.to_numeric(series, errors='coerce') <= float(value)
    return series.astype(str) == str(value)


def console_mode(signals_file, macros_file, output_file, rules_file=None):
    signals_df = pd.read_excel(signals_file)
    macros_df = pd.read_excel(macros_file)
    print(f"Loaded signals: {len(signals_df)} rows")
    print(f"Loaded macros: {len(macros_df)} rows")
    if rules_file:
        rules_df = pd.read_excel(rules_file)
        print(f"Loaded macro rules: {len(rules_df)} rows")
        signals_df = apply_signal_rules(signals_df, rules_df, output_consumer=lambda msg: print(msg, end=''))
        print(f"Signals after rule filtering: {len(signals_df)} rows")
    signal_col = 'name'
    symbol_col = 'Symbol'
    if signal_col not in signals_df.columns or symbol_col not in macros_df.columns:
        print("Required columns not found.")
        return
    matched_macros = macros_df[macros_df[symbol_col].isin(signals_df[signal_col])]
    matched_macros.to_csv(output_file, sep=",", index=False)
    print(f"Exported {len(matched_macros)} rows to {output_file}")


if __name__ == "__main__":
    if len(sys.argv) == 4:
        console_mode(sys.argv[1], sys.argv[2], sys.argv[3])
    elif len(sys.argv) == 5:
        console_mode(sys.argv[1], sys.argv[2], sys.argv[4], rules_file=sys.argv[3])
    else:
        root = tk.Tk()
        app = PLCApp(root)
        root.mainloop()
