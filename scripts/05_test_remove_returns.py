"""
Fix PytestReturnNotNoneWarning: remove return True/False from test_* functions.

Rules:
- return True  → delete the line (success path, test passes by default)
- return False → replace with assert False, "message" (extract from preceding print)
- Only modifies lines inside def test_* functions (not main/helpers)
"""

import re
import os
import sys

TEST_DIR = os.path.join(os.path.dirname(__file__), "..", "tests")

# Files to process
TARGET_FILES = [
    "test_converter.py",
    "test_day6_7.py",
    "test_day8_9.py",
    "test_day10_11.py",
    "test_excel.py",
    "test_final_integration.py",
    "test_ocr.py",
    "test_pdf.py",
    "test_ppt.py",
]


def extract_print_message(lines, idx):
    """Look backwards from idx to find the nearest print() and extract its message."""
    for i in range(idx - 1, max(idx - 5, -1), -1):
        stripped = lines[i].strip()
        # Match print("...") or print(f"...")
        m = re.match(r'print\(f?["\'](.+?)["\']', stripped)
        if m:
            msg = m.group(1)
            # Clean up f-string placeholders for static assert message
            msg = re.sub(r'\{[^}]+\}', '...', msg)
            # Remove leading symbols
            msg = msg.lstrip('✓✗× ·-')
            return msg.strip()
    return "test failed"


def get_indent(line):
    """Return leading whitespace."""
    return line[:len(line) - len(line.lstrip())]


def find_function_at_line(lines, target_idx):
    """Walk backwards to find the enclosing def name."""
    for i in range(target_idx, -1, -1):
        stripped = lines[i].strip()
        m = re.match(r'def (\w+)\s*\(', stripped)
        if m:
            return m.group(1)
    return None


def process_file(filepath):
    """Process a single test file."""
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()

    removed_true = 0
    replaced_false = 0
    new_lines = []
    skip_next_blank = False

    for idx, line in enumerate(lines):
        stripped = line.strip()

        # Only process return True/False inside test_* functions
        if stripped in ("return True", "return False"):
            func_name = find_function_at_line(lines, idx)
            if func_name and func_name.startswith("test_"):
                if stripped == "return True":
                    removed_true += 1
                    # Skip this line, also skip trailing blank line if any
                    skip_next_blank = True
                    continue
                else:  # return False
                    indent = get_indent(line)
                    msg = extract_print_message(lines, idx)
                    new_lines.append(f'{indent}assert False, "{msg}"\n')
                    replaced_false += 1
                    continue

        # Skip blank line right after a removed "return True"
        if skip_next_blank:
            skip_next_blank = False
            if stripped == "":
                continue

        new_lines.append(line)

    with open(filepath, "w", encoding="utf-8") as f:
        f.writelines(new_lines)

    return removed_true, replaced_false


def main():
    total_removed = 0
    total_replaced = 0

    for fname in TARGET_FILES:
        fpath = os.path.join(TEST_DIR, fname)
        if not os.path.exists(fpath):
            print(f"  SKIP {fname} (not found)")
            continue

        removed, replaced = process_file(fpath)
        total_removed += removed
        total_replaced += replaced
        print(f"  {fname}: removed {removed} return True, replaced {replaced} return False")

    print(f"\nTotal: removed {total_removed} return True, replaced {total_replaced} return False")
    print(f"Total lines changed: {total_removed + total_replaced}")


if __name__ == "__main__":
    main()
