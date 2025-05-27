import re

FUNCTION_PATTERN = re.compile(r"(?P<func>[A-Z]+)\((?P<args>.*)\)")
LITERAL_PATTERN = re.compile(r"^(TRUE|FALSE|\d+(\.\d+)?|\".*?\"|[A-Z]+\d+(:[A-Z]+\d+)?)$", re.IGNORECASE)


def parse_formula(formula):
    if formula.startswith("="):
        formula = formula[1:]
    return _parse_expression(formula.strip())


def _parse_expression(expr):
    expr = expr.strip()

    if LITERAL_PATTERN.fullmatch(expr):
        return expr

    match = FUNCTION_PATTERN.fullmatch(expr)
    if not match:
        return expr

    func = match.group("func")
    args_str = match.group("args")
    args = _split_arguments(args_str)
    parsed_args = [_parse_expression(arg.strip()) for arg in args]
    return {"function": func, "args": parsed_args}


def _split_arguments(args_str):
    args = []
    depth = 0
    current = []
    for c in args_str:
        if c == ',' and depth == 0:
            args.append(''.join(current))
            current = []
        else:
            if c == '(':
                depth += 1
            elif c == ')':
                depth -= 1
            current.append(c)
    if current:
        args.append(''.join(current))
    return args


def format_formula_tree(tree, indent=0):
    if isinstance(tree, str):
        return '  ' * indent + f"- {tree}\n"
    if isinstance(tree, dict):
        func = tree.get("function")
        args = tree.get("args", [])
        output = '  ' * indent + f"Function: {func}\n"
        for i, arg in enumerate(args):
            label = _arg_label(func, i)
            output += '  ' * (indent + 1) + f"{label}:\n"
            output += format_formula_tree(arg, indent + 2)
        return output


def format_formula_explanation(tree):
    if isinstance(tree, str):
        return tree
    if isinstance(tree, dict):
        func = tree.get("function")
        args = tree.get("args", [])

        if func == "IF" and len(args) == 3:
            return (f"Checks if {format_formula_explanation(args[0])}. "
                    f"If true, returns {format_formula_explanation(args[1])}; "
                    f"otherwise returns {format_formula_explanation(args[2])}.")

        elif func == "IFERROR" and len(args) == 2:
            return (f"Returns {format_formula_explanation(args[0])} unless it results in an error, "
                    f"in which case returns {format_formula_explanation(args[1])}.")

        elif func == "VLOOKUP" and len(args) >= 3:
            return (f"Looks up {format_formula_explanation(args[0])} in the table {format_formula_explanation(args[1])}, "
                    f"and returns the value from column {format_formula_explanation(args[2])}.")

        elif func == "MATCH" and len(args) >= 2:
            return (f"Returns the relative position of {format_formula_explanation(args[0])} in the range "
                    f"{format_formula_explanation(args[1])}.")

        elif func == "TEXTJOIN":
            return f"Joins multiple values using a delimiter: {', '.join(format_formula_explanation(arg) for arg in args)}."

        elif func == "CONCAT":
            return f"Concatenates the following values: {', '.join(format_formula_explanation(arg) for arg in args)}."

        elif func == "AND":
            conditions = [format_formula_explanation(arg) for arg in args]
            return f"Returns TRUE if all of the following are true: {', '.join(conditions)}."

        elif func == "OR":
            conditions = [format_formula_explanation(arg) for arg in args]
            return f"Returns TRUE if any of the following are true: {', '.join(conditions)}."

        elif func == "LEN" and args:
            return f"Returns the number of characters in {format_formula_explanation(args[0])}."

        elif func == "TRIM" and args:
            return f"Removes all extra spaces from {format_formula_explanation(args[0])}."

        elif func == "LEFT" and args:
            return f"Returns the first characters from the left in {format_formula_explanation(args[0])}."

        elif func == "RIGHT" and args:
            return f"Returns the last characters from the right in {format_formula_explanation(args[0])}."

        elif func == "UNIQUE" and args:
            return f"Returns a list of unique values from {format_formula_explanation(args[0])}."

        elif func == "GETPIVOTDATA" and len(args) >= 2:
            return f"Returns data stored in a pivot table for field {format_formula_explanation(args[0])} and data source {format_formula_explanation(args[1])}."

        return f"{func} function with arguments: {', '.join(format_formula_explanation(arg) for arg in args)}"


def _arg_label(func, index):
    labels = {
        "IF": ["Condition", "If True", "If False"],
        "IFERROR": ["Value", "If Error"],
        "VLOOKUP": ["Lookup Value", "Table Array", "Column Index", "Range Lookup"],
        "MATCH": ["Lookup Value", "Lookup Array", "Match Type"],
        "AND": ["Condition 1", "Condition 2"],
        "OR": ["Condition 1", "Condition 2"],
        "SUM": ["Value"],
        "LEN": ["Text"],
        "TRIM": ["Text"],
        "LEFT": ["Text", "Num Characters"],
        "RIGHT": ["Text", "Num Characters"],
        "TEXTJOIN": ["Delimiter", "Ignore Empty", "Text1", "Text2"],
        "UNIQUE": ["Array"],
        "GETPIVOTDATA": ["Data Field", "Pivot Table"]
    }
    if func in labels and index < len(labels[func]):
        return labels[func][index]
    return f"Arg {index + 1}"
