import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class PLCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PLC Signal Connector")
        self.signals_df = None
        self.macros_df = None
        self.macro_group_col = None
        self.macro_list_items = []

        # Buttons
        self.load_signals_btn = tk.Button(root, text="Load PLC Signals Excel", command=self.load_signals)
        self.load_signals_btn.pack(pady=10)

        self.load_macros_btn = tk.Button(root, text="Load Macros File", command=self.load_macros)
        self.load_macros_btn.pack(pady=10)

        self.process_btn = tk.Button(root, text="Process and Connect", command=self.process, state=tk.DISABLED)
        self.process_btn.pack(pady=10)

        self.exit_btn = tk.Button(root, text="Exit App", command=self.root.destroy)
        self.exit_btn.pack(pady=10)

        self.editor_frame = tk.Frame(root)
        self.editor_frame.pack(pady=10)
        self.add_macro_btn = tk.Button(self.editor_frame, text="Add Macro", command=self.add_macro)
        self.add_macro_btn.pack(side=tk.LEFT, padx=5)
        self.edit_macro_btn = tk.Button(self.editor_frame, text="Edit Selected Macro", command=self.edit_macro)
        self.edit_macro_btn.pack(side=tk.LEFT, padx=5)
        self.delete_macro_btn = tk.Button(self.editor_frame, text="Delete Selected Macro", command=self.delete_macro)
        self.delete_macro_btn.pack(side=tk.LEFT, padx=5)
        self.save_macros_btn = tk.Button(self.editor_frame, text="Save Macros Excel", command=self.save_macros)
        self.save_macros_btn.pack(side=tk.LEFT, padx=5)

        # Macro listbox for selection
        self.macro_label = tk.Label(root, text="Loaded macros: 0")
        self.macro_label.pack()
        self.macro_listbox = tk.Listbox(root, width=120, height=8)
        self.macro_listbox.pack(pady=5)

        # Text area for output
        self.output_text = tk.Text(root, height=20, width=80)
        self.output_text.pack(pady=10)

    def refresh_macro_listbox(self):
        self.macro_listbox.delete(0, tk.END)
        self.macro_list_items = []
        if self.macros_df is None:
            self.update_macro_label()
            return

        if self.macro_group_col is not None and self.macro_group_col in self.macros_df.columns:
            seen = set()
            for macro_type in self.macros_df[self.macro_group_col].astype(str):
                if macro_type in seen:
                    continue
                seen.add(macro_type)
                count = int((self.macros_df[self.macro_group_col].astype(str) == macro_type).sum())
                self.macro_list_items.append(macro_type)
                self.macro_listbox.insert(tk.END, f"{len(self.macro_list_items)}: {macro_type} ({count} lines)")
        else:
            for idx, row in self.macros_df.iterrows():
                self.macro_list_items.append(idx)
                self.macro_listbox.insert(tk.END, self.format_macro_row(idx, row))

        self.update_macro_label()

    def format_macro_row(self, idx, row):
        if 'Symbol' in row.index:
            return f"{idx + 1}: {row['Symbol']}"
        if len(row.index) > 0:
            first_col = row.index[0]
            return f"{idx + 1}: {row[first_col]}"
        return f"{idx + 1}: {row.to_dict()}"

    def update_macro_label(self):
        if self.macros_df is None:
            self.macro_label.config(text="Loaded macros: 0")
            return
        if self.macro_group_col is not None and self.macro_group_col in self.macros_df.columns:
            self.macro_label.config(text=f"Loaded macros: {len(self.macro_list_items)} groups ({len(self.macros_df)} rows)")
        else:
            self.macro_label.config(text=f"Loaded macros: {len(self.macros_df)}")

    def detect_signal_column(self):
        candidates = ['name', 'Name', 'Signal', 'signal', 'Tag', 'tag']
        for candidate in candidates:
            if candidate in self.signals_df.columns:
                return candidate
        return self.signals_df.columns[0]

    def detect_macro_input_column(self):
        candidates = ['Input', 'Inputs', 'Signal', 'Signals', 'Signal name', 'Name', 'Tag', 'Symbol']
        for candidate in candidates:
            if candidate in self.macros_df.columns:
                return candidate
        return self.macros_df.columns[0]

    def extract_symbol_parts(self, symbol):
        if pd.isna(symbol):
            return []
        value = str(symbol).strip()
        if '=' in value:
            left, right = value.split('=', 1)
            return [left.strip(), right.strip(), value]
        return [value]

    def build_macro_groups(self):
        if self.macro_group_col is not None and self.macro_group_col in self.macros_df.columns:
            if self.macro_group_col == 'Symbol':
                keys = self.macros_df['Symbol'].astype(str).str.split('=', 1).str[0]
                return dict(tuple(self.macros_df.groupby(keys)))
            return dict(tuple(self.macros_df.groupby(self.macro_group_col)))
        if 'Symbol' in self.macros_df.columns:
            keys = self.macros_df['Symbol'].astype(str).str.split('=', 1).str[0]
            return dict(tuple(self.macros_df.groupby(keys)))
        return {'ALL': self.macros_df}

    def collect_group_inputs(self, group_df, input_col):
        values = set()
        if input_col in group_df.columns:
            for cell in group_df[input_col].dropna().astype(str):
                for part in self.extract_symbol_parts(cell):
                    candidate = part.strip()
                    if candidate:
                        values.add(candidate)
            if values:
                return values

        if input_col == 'Symbol' and 'Symbol' in group_df.columns:
            symbol_values = [str(cell) for cell in group_df['Symbol'].dropna().astype(str)]
            if any('=' in val for val in symbol_values):
                for cell in symbol_values:
                    for part in self.extract_symbol_parts(cell):
                        candidate = part.strip()
                        if candidate:
                            values.add(candidate)
                if values:
                    return values
        return values

    def signal_matches_inputs(self, signal, inputs):
        sig = str(signal).strip()
        if not sig:
            return False
        for inp in inputs:
            if not inp:
                continue
            if '=' in inp:
                if inp == sig:
                    return True
            elif inp in sig or sig in inp:
                return True
        return False

    def extract_macro_channel(self, macro_key):
        if not isinstance(macro_key, str):
            return None
        if 'Ch1' in macro_key or 'ch1' in macro_key:
            return 1
        if 'Ch2' in macro_key or 'ch2' in macro_key:
            return 2
        return None

    def extract_signal_channel(self, signal):
        sig = str(signal)
        if '_12' in sig:
            return 1
        if '_22' in sig:
            return 2
        if 'AS01' in sig:
            return 1
        if 'AS02' in sig:
            return 2
        return None

    def extract_macro_name_from_symbol(self, symbol):
        text = str(symbol)
        if '=' in text:
            return text.split('=', 1)[0].strip()
        return text.strip()

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
            messagebox.showerror("Error", "Load a macros file first.")
            return
        selection = self.macro_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Select a macro to edit.")
            return
        idx = selection[0]
        if self.macro_group_col is not None and self.macro_group_col in self.macros_df.columns:
            group_key = self.macro_list_items[idx]
            group_rows = self.macros_df[self.macros_df[self.macro_group_col].astype(str) == group_key]
            if group_rows.empty:
                messagebox.showerror("Error", "Selected macro group not found.")
                return
            row_idx = group_rows.index[0]
        else:
            row_idx = self.macro_list_items[idx]

        columns = list(self.macros_df.columns)
        for i, col in enumerate(columns):
            old_val = str(self.macros_df.iloc[row_idx, i])
            val = simple_input_dialog(self.root, f"Edit value for '{col}':", old_val)
            if val is None:
                return
            self.macros_df.iloc[row_idx, i] = val
        self.refresh_macro_listbox()
        self.output_text.insert(tk.END, "Edited macro.\n")

    def delete_macro(self):
        if self.macros_df is None:
            messagebox.showerror("Error", "Load a macros file first.")
            return
        selection = self.macro_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Select a macro to delete.")
            return
        idx = selection[0]
        try:
            if self.macro_group_col is not None and self.macro_group_col in self.macros_df.columns:
                group_key = self.macro_list_items[idx]
                self.macros_df = self.macros_df[self.macros_df[self.macro_group_col].astype(str) != group_key].reset_index(drop=True)
                self.output_text.insert(tk.END, f"Deleted macro group '{group_key}'.\n")
            else:
                self.macros_df = self.macros_df.drop(self.macros_df.index[idx]).reset_index(drop=True)
                self.output_text.insert(tk.END, f"Deleted macro {idx + 1}.\n")
            self.refresh_macro_listbox()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete macro: {str(e)}")

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
            self.output_text.insert(tk.END, f"Failed to save macros: {str(e)}\n")

    def check_enable_process(self):
        if self.signals_df is not None and self.macros_df is not None:
            self.process_btn.config(state=tk.NORMAL)

    def process(self):
        self.output_text.insert(tk.END, "Processing signals and connecting to macros...\n")
        if self.signals_df is None or self.macros_df is None:
            self.output_text.insert(tk.END, "Load both signals and macros before processing.\n")
            return

        signal_col = self.detect_signal_column()
        input_col = self.detect_macro_input_column()
        macro_groups = self.build_macro_groups()

        signal_names = {str(value).strip() for value in self.signals_df[signal_col].dropna()}
        if not signal_names:
            self.output_text.insert(tk.END, f"Signal column '{signal_col}' is empty.\n")
            return

        default_pattern = '=F9-SGC1'
        used_signals = set()

        matched_groups = []
        for group_key, group_df in macro_groups.items():
            inputs = self.collect_group_inputs(group_df, input_col)
            group_rows = []
            channel = self.extract_macro_channel(group_key)
            if not inputs:
                for sig in sorted(signal_names):
                    if sig in used_signals:
                        continue
                    if default_pattern not in sig:
                        continue
                    signal_channel = self.extract_signal_channel(sig)
                    if channel is not None and signal_channel != channel:
                        continue
                    row = group_df.iloc[0].copy()
                    row['Symbol'] = sig
                    row['macro name'] = self.extract_macro_name_from_symbol(sig)
                    group_rows.append(row)
                    used_signals.add(sig)
                if not group_rows:
                    self.output_text.insert(tk.END, f"Macro '{group_key}' skipped: no signals matched pattern '{default_pattern}' for channel {channel}.\n")
            else:
                for sig in signal_names:
                    if sig in used_signals:
                        continue
                    if self.signal_matches_inputs(sig, inputs):
                        row = group_df.iloc[0].copy()
                        row['Symbol'] = sig
                        row['macro name'] = self.extract_macro_name_from_symbol(sig)
                        group_rows.append(row)
                        used_signals.add(sig)

            if group_rows:
                matched_groups.append((group_key, pd.DataFrame(group_rows)))
            else:
                self.output_text.insert(tk.END, f"Macro '{group_key}' skipped: no matching signals for inputs {sorted(inputs)}.\n")

        if not matched_groups:
            self.output_text.insert(tk.END, "No matching macro groups found.\n")
            return

        matched_macros = pd.concat([group_df for _, group_df in matched_groups], ignore_index=True)
        matched_macros = self.prepare_export_dataframe(matched_macros)
        self.output_text.insert(tk.END, f"Matched {len(matched_groups)} macro groups ({len(matched_macros)} rows).\n")

        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")], title="Save result as...")
        if not file_path:
            self.output_text.insert(tk.END, "Export cancelled.\n")
            return
        try:
            matched_macros.to_csv(file_path, sep=",", index=False)
            self.output_text.insert(tk.END, f"Exported {len(matched_macros)} connected macro rows to {file_path}\n")
        except Exception as e:
            self.output_text.insert(tk.END, f"Failed to export: {str(e)}\n")

    def prepare_export_dataframe(self, df):
        cols = ['Symbol', 'Type', 'Comment', 'Definition internal Format', 'Definition external Format', 'macro name', 'macro type', 'macro signal', 'target file name']
        export_df = pd.DataFrame(index=df.index)
        for col in cols:
            if col in df.columns:
                export_df[col] = df[col]
            else:
                export_df[col] = ''

        if 'macro name' not in df.columns or export_df['macro name'].eq('').all():
            if 'Symbol' in df.columns:
                export_df['macro name'] = df['Symbol'].astype(str).str.split('=', 1).str[0]

        if 'macro type' not in df.columns:
            if 'Macro type' in df.columns:
                export_df['macro type'] = df['Macro type']
            elif self.macro_group_col in ['macro type', 'Macro type'] and self.macro_group_col in df.columns:
                export_df['macro type'] = df[self.macro_group_col].astype(str)
            else:
                export_df['macro type'] = ''

        if 'macro signal' not in df.columns and 'Macro signal' in df.columns:
            export_df['macro signal'] = df['Macro signal']

        if 'target file name' not in df.columns and 'Target file name' in df.columns:
            export_df['target file name'] = df['Target file name']

        return export_df[cols]

    def load_signals(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.signals_df = pd.read_excel(file_path)
                self.output_text.insert(tk.END, f"Loaded signals: {len(self.signals_df)} rows\n")
                self.output_text.insert(tk.END, f"Columns: {list(self.signals_df.columns)}\n\n")
                self.check_enable_process()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load signals: {str(e)}")

    def load_macros(self):
        file_path = filedialog.askopenfilename(filetypes=[("Macro files", "*.txt *.csv *.xlsx *.xls")])
        if file_path:
            try:
                if file_path.lower().endswith((".txt", ".csv")):
                    self.macros_df = pd.read_csv(file_path, sep=None, engine="python")
                else:
                    self.macros_df = pd.read_excel(file_path)

                self.macro_group_col = None
                for candidate in ["macro signal", "Macro signal", "Symbol", "macro type", "Macro type", "MacroType", "macro_type", "Type", "Macro Type", "macro name", "Macro name"]:
                    if candidate in self.macros_df.columns:
                        self.macro_group_col = candidate
                        break

                self.output_text.insert(tk.END, f"Loaded macros: {len(self.macros_df)} rows\n")
                self.output_text.insert(tk.END, f"Columns: {list(self.macros_df.columns)}\n\n")
                if self.macro_group_col is not None:
                    self.output_text.insert(tk.END, f"Grouped by: {self.macro_group_col}\n\n")
                self.check_enable_process()
                self.refresh_macro_listbox()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load macros: {str(e)}")

def simple_input_dialog(root, prompt, initial_value=""):
    dialog = tk.Toplevel(root)
    dialog.title("Input")
    tk.Label(dialog, text=prompt).pack(padx=10, pady=5)
    entry = tk.Entry(dialog)
    entry.pack(padx=10, pady=5)
    entry.insert(0, initial_value)
    result = {"value": None}

    def on_ok():
        result["value"] = entry.get()
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    tk.Button(dialog, text="OK", command=on_ok).pack(side=tk.LEFT, padx=10, pady=5)
    tk.Button(dialog, text="Cancel", command=on_cancel).pack(side=tk.RIGHT, padx=10, pady=5)
    dialog.grab_set()
    root.wait_window(dialog)
    return result["value"]

if __name__ == "__main__":
    root = tk.Tk()
    app = PLCApp(root)
    root.mainloop()
