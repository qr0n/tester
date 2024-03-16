import re

def extract_function_signatures(file_path):
    with open(file_path, 'r') as file:
        content = file.read()

    # Regular expression to find all function signatures
    pattern = re.compile(r'\b(?:\w+\s+)?(\w+\s+\w+\([^)]*\))\s*{')
    matches = pattern.finditer(content)

    # Extract function signatures
    function_signatures = [match.group(1) for match in matches]

    return function_signatures

def extract(c_file_path):
    all_function_signatures = extract_function_signatures(c_file_path)

    if all_function_signatures:
        print("All Function Signatures Found:")
        for i, function_signature in enumerate(all_function_signatures, start=1):
            print(f"Function {i}: {function_signature}")
    else:
        print("No functions found in the provided C file.")
