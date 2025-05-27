import tkinter as tk
from tkinter import ttk
from formula_parser import parse_formula, format_formula_tree

class FormulaDebuggerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Formula Debugger")
        self.root.geometry("700x500")
        self.root.configure(bg="#f4f4f4")

        self.setup_widgets()

    def setup_widgets(self):
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 11))
        style.configure("TEntry", font=("Segoe UI", 11))

        self.input_label = ttk.Label(self.root, text="Enter Excel Formula:")
        self.input_label.pack(pady=(20, 5))

        self.formula_entry = ttk.Entry(self.root, width=80)
        self.formula_entry.pack(pady=5)

        self.analyze_button = ttk.Button(self.root, text="Analyze Formula", command=self.analyze_formula)
        self.analyze_button.pack(pady=10)

        self.output_text = tk.Text(self.root, wrap="word", height=20, font=("Consolas", 10))
        self.output_text.pack(padx=10, pady=(10, 20), fill="both", expand=True)

    def analyze_formula(self):
        formula = self.formula_entry.get()
        tree = parse_formula(formula)
        output = format_formula_tree(tree)
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, output)

if __name__ == "__main__":
    root = tk.Tk()
    app = FormulaDebuggerApp(root)
    root.mainloop()
