import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class PLCApp:
            # Macro editor controls
            self.editor_frame = tk.Frame(root)
            self.editor_frame.pack(pady=10)
            self.add_macro_btn = tk.Button(self.editor_frame, text="Add Macro", command=self.add_macro)
            self.add_macro_btn.pack(side=tk.LEFT, padx=5)
            self.edit_macro_btn = tk.Button(self.editor_frame, text="Edit Selected Macro", command=self.edit_macro)
            self.edit_macro_btn.pack(side=tk.LEFT, padx=5)
            self.save_macros_btn = tk.Button(self.editor_frame, text="Save Macros Excel", command=self.save_macros)
            self.save_macros_btn.pack(side=tk.LEFT, padx=5)

            # Macro listbox for selection
            self.macro_listbox = tk.Listbox(root, width=120, height=8)
            self.macro_listbox.pack(pady=5)

        def refresh_macro_listbox(self):
            self.macro_listbox.delete(0, tk.END)
            if self.macros_df is not None:
                for idx, row in self.macros_df.iterrows():
                    self.macro_listbox.insert(tk.END, f"{row.to_dict()}")

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
            columns = list(self.macros_df.columns)
            for i, col in enumerate(columns):
                old_val = str(self.macros_df.iloc[idx, i])
                val = simple_input_dialog(self.root, f"Edit value for '{col}':", old_val)
                if val is None:
                    return
                self.macros_df.iloc[idx, i] = val
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
                self.output_text.insert(tk.END, f"Failed to save macros: {str(e)}\n")
    def __init__(self, root):
        self.root = root
        self.root.title("PLC Signal Connector")
        self.signals_df = None
        self.macros_df = None

        # Buttons
        self.load_signals_btn = tk.Button(root, text="Load PLC Signals Excel", command=self.load_signals)
        self.load_signals_btn.pack(pady=10)

        self.load_macros_btn = tk.Button(root, text="Load Macros Excel", command=self.load_macros)
        self.load_macros_btn.pack(pady=10)

        self.process_btn = tk.Button(root, text="Process and Connect", command=self.process, state=tk.DISABLED)
        self.process_btn.pack(pady=10)

        # Text area for output
        self.output_text = tk.Text(root, height=20, width=80)
        self.output_text.pack(pady=10)

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
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.macros_df = pd.read_excel(file_path)
                self.output_text.insert(tk.END, f"Loaded macros: {len(self.macros_df)} rows\n")
                self.output_text.insert(tk.END, f"Columns: {list(self.macros_df.columns)}\n\n")
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

    def check_enable_process(self):
        if self.signals_df is not None and self.macros_df is not None:
            self.process_btn.config(state=tk.NORMAL)

    def process(self):
        self.output_text.insert(tk.END, "Processing signals and connecting to macros...\n")
        # Use 'name' as the signal column
        signal_col = 'name'
        macro_symbol_col = 'Symbol'
        # Filter macros to only those with a matching symbol in signals
        if signal_col not in self.signals_df.columns or macro_symbol_col not in self.macros_df.columns:
            self.output_text.insert(tk.END, "Required columns not found in the files.\n")
            return
        
        # Only keep macros where Symbol matches a signal name
        matched_macros = self.macros_df[self.macros_df[macro_symbol_col].isin(self.signals_df[signal_col])]
        if matched_macros.empty:
            self.output_text.insert(tk.END, "No matching signals found for macros.\n")
            return
        
        # Ask user where to save the txt file
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")], title="Save result as...")
        if not file_path:
            self.output_text.insert(tk.END, "Export cancelled.\n")
            return
        try:
            # Save as CSV with comma separator, no index, header included
            matched_macros.to_csv(file_path, sep=",", index=False)
            self.output_text.insert(tk.END, f"Exported {len(matched_macros)} connected macros to {file_path}\n")
        except Exception as e:
            self.output_text.insert(tk.END, f"Failed to export: {str(e)}\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = PLCApp(root)
    root.mainloop()