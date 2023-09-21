import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pyperclip

class App:
    def __init__(self, root):
        # Label and Textbox for input template
        tk.Label(root, text="PASTE TEMPLATE CODE HERE").pack(pady=10)
        self.template_text = tk.Text(root, height=10, width=50)
        self.template_text.pack(pady=10)
        
        # Label and Textbox for output
        tk.Label(root, text="GENERATED CODE").pack(pady=10)
        self.output_text = tk.Text(root, height=10, width=50)
        self.output_text.pack(pady=10)
        
        # Button to load Excel file
        self.load_button = tk.Button(root, text="Load XLSX File", command=self.load_xlsx)
        self.load_button.pack(pady=10)

        # Button to run the process
        self.run_button = tk.Button(root, text="RUN", command=self.generate_code)
        self.run_button.pack(pady=20)
        
        # Variable to store the loaded dataframe
        self.df = None

    def load_xlsx(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not filepath:
            return
        self.df = pd.read_excel(filepath)
    
    def generate_code(self):
        if self.df is None:
            messagebox.showerror("Error", "Please load an XLSX file first.")
            return
        
        template = self.template_text.get("1.0", "end-1c")
        generated_codes = []

        for _, row in self.df.iterrows():
            code = template
            for col in self.df.columns:
                code = code.replace(col, str(row[col]))
            generated_codes.append(code)
        
        output = ('\n' * 3).join(generated_codes)
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", output)
        
        pyperclip.copy(output)
        messagebox.showinfo("Info", "Code generated and saved to clipboard. Paste with ctrl-v.")
        
root = tk.Tk()
root.title("Code Generator")
app = App(root)
root.mainloop()
