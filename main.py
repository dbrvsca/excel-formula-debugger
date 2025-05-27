import tkinter as tk
from tkinter import ttk
from formula_parser import parse_formula, format_formula_tree, format_formula_explanation

class FormulaDebuggerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Formula Debugger")
        self.root.geometry("900x600")
        self.root.configure(bg="#1e1e1e")  # dark VSCode-style background
        self.view_mode = "tree"  # or "explanation"

        self.setup_widgets()

    def setup_widgets(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TLabel",
                        font=("Segoe UI", 11),
                        foreground="#ffffff",
                        background="#1e1e1e")

        self.input_label = ttk.Label(self.root, text="Enter Excel Formula:")
        self.input_label.pack(pady=(20, 5))

        self.formula_entry = tk.Entry(self.root, font=("Courier New", 12), bg="#2d2d2d", fg="#ffffff",
                                      insertbackground="#ffffff", width=80, relief=tk.FLAT)
        self.formula_entry.pack(pady=5, ipady=6)
        self.formula_entry.bind("<Return>", lambda event: self.analyze_formula())  # Bind Enter key
        self.formula_entry.focus_set()  # Automatically focus on input

        button_frame = tk.Frame(self.root, bg="#1e1e1e")
        button_frame.pack(pady=10)

        button_style = {
            "bg": "#2d2d2d",
            "fg": "#2d2d2d",
            "activebackground": "#2d2d2d",
            "activeforeground": "#A41F1F",
            "font": ("Segoe UI", 10, "bold"),
            "relief": tk.FLAT,
            "padx": 12,
            "pady": 6,
            "bd": 0
        }

        self.analyze_button = tk.Button(button_frame, text="Analyze Formula", command=self.analyze_formula, **button_style)
        self.analyze_button.grid(row=0, column=0, padx=10)

        self.toggle_button = tk.Button(button_frame, text="Switch to Explanation View", command=self.toggle_view, **button_style)
        self.toggle_button.grid(row=0, column=1, padx=10)

        output_frame = tk.Frame(self.root, bg="#1e1e1e")
        output_frame.pack(padx=15, pady=(10, 20), fill="both", expand=True)

        self.output_text = tk.Text(output_frame, wrap="word", font=("Courier New", 11), bg="#252526",
                                    fg="#d4d4d4", insertbackground="#ffffff", relief=tk.FLAT)
        self.output_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(output_frame, command=self.output_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.output_text.config(yscrollcommand=scrollbar.set)

    def analyze_formula(self):
        self.formula_entry.focus_set()  # Ensure Entry is active
        raw_input = self.formula_entry.get().strip()
        if not raw_input:
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert(tk.END, "Please enter a formula to analyze.")
            return

        outputs = []
        for line in raw_input.split("="):
            line = line.strip()
            if not line:
                continue
            try:
                formula = "=" + line  # reattach the = for proper parsing
                tree = parse_formula(formula)
                if self.view_mode == "tree":
                    output = format_formula_tree(tree)
                else:
                    output = format_formula_explanation(tree)
                outputs.append(output)
            except Exception as e:
                outputs.append(f"Error parsing: {line}\n{str(e)}")

        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, "\n\n".join(outputs))

    def toggle_view(self):
        if self.view_mode == "tree":
            self.view_mode = "explanation"
            self.toggle_button.config(text="Switch to Tree View")
        else:
            self.view_mode = "tree"
            self.toggle_button.config(text="Switch to Explanation View")
        self.analyze_formula()

if __name__ == "__main__":
    root = tk.Tk()
    app = FormulaDebuggerApp(root)
    root.mainloop()
