# Excel Formula Debugger

A Python GUI tool for analyzing Excel formulas.

* Built with `tkinter`, this app provides two views:

- **Tree View** – shows the parsed formula structure
- **Explanation View** – provides a human-readable breakdown of what the formula does


---

## Features

- Analyze Excel formulas in real time
- Supports Enter key for quick evaluation
- Toggle between *tree* and *natural language* explanation
- Styled in dark mode (VSCode-style)
- Supports common functions like `IF`, `IFERROR`, `VLOOKUP`, `MATCH`, `TEXTJOIN`, `UNIQUE`, `LEFT`, `RIGHT`, `GETPIVOTDATA`, etc.
- Smart parsing and argument splitting

---

## Requirements

- Python 3.7+
- No additional libraries needed (uses standard library)

---

## File Structure

excel-formula-debugger/
├── main.py              # GUI logic
├── formula_parser.py    # Parsing + explanation logic
├── README.md

---

## Development Tips

* Add more Excel function explanations inside `formula_parser.py > format_formula_explanation`
* Extend parsing logic using regex in `_parse_expression`
* UI customizations are in `main.py > setup_widgets`

---

## LICENSE

MIT License. Feel free to fork and improve!

---



## Usage

```bash
# Clone the repository

git clone https://github.com/dbrvsca/excel-formula-debugger.git
cd excel-formula-debugger

# Run the app
python main.py
```
