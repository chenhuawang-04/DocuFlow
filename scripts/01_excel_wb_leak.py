#!/usr/bin/env python3
"""C1 Fix: Add try-finally blocks around load_workbook() calls in excel.py.

Transforms:
    wb = load_workbook(path)
    # ... code with early returns ...
    wb.close()
    return {success: True, ...}

Into:
    wb = load_workbook(path)
    try:
        # ... code with early returns ...
        return {success: True, ...}
    finally:
        wb.close()

Special handling for stats_summary() which has two load_workbook() calls:
- Keeps intermediate wb.close() before the second open (releases read lock)
- Wraps both opens in the same try-finally scope
"""

import sys
import os


def process_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 1-based line numbers of load_workbook calls to wrap with try-finally.
    # Line 1963 (stats_summary 2nd open) is EXCLUDED — it's inside 1st open's scope.
    wb_lines_1based = [
        134, 223, 296, 340, 373, 407, 445, 485, 540, 612, 706, 821, 877,
        927, 971, 1015, 1061, 1126, 1267, 1368, 1430, 1532, 1600, 1684,
        1804, 1902, 2032, 2154, 2271, 2399, 2496
    ]

    # stats_summary's intermediate wb.close() at line 1962 (1-based)
    # Must be KEPT because it releases read lock before second open.
    STATS_KEEP_CLOSE_LINE = 1962  # 1-based

    processed = 0
    removed_closes = 0

    # Process from bottom to top so earlier indices stay stable.
    for wb_line_1based in reversed(wb_lines_1based):
        wb_idx = wb_line_1based - 1
        wb_line = lines[wb_idx]
        indent = len(wb_line) - len(wb_line.lstrip())
        indent_str = ' ' * indent

        # Find the matching 'except' at outer try level (indent - 4)
        outer_indent = indent - 4
        except_idx = None
        for i in range(wb_idx + 1, len(lines)):
            line = lines[i]
            if line.strip() == '':
                continue
            line_indent = len(line) - len(line.lstrip())
            stripped = line.strip()
            if line_indent == outer_indent and stripped.startswith('except '):
                except_idx = i
                break

        if except_idx is None:
            print(f"WARNING: No except found for wb at line {wb_line_1based}",
                  file=sys.stderr)
            continue

        # Build replacement lines between wb and except
        new_middle = []
        new_middle.append(indent_str + 'try:\n')

        for i in range(wb_idx + 1, except_idx):
            line = lines[i]
            stripped = line.strip()

            # Handle wb.close() lines
            if stripped == 'wb.close()':
                # Special case: stats_summary intermediate close — keep it
                if (wb_line_1based == 1902
                        and (i + 1) == STATS_KEEP_CLOSE_LINE):
                    current_indent = len(line) - len(line.lstrip())
                    new_middle.append(
                        ' ' * (current_indent + 4) + stripped + '\n')
                    continue
                # All other wb.close() — remove (handled by finally)
                removed_closes += 1
                continue

            # Empty lines
            if stripped == '':
                new_middle.append('\n')
                continue

            # Indent by 4 more spaces
            current_indent = len(line) - len(line.lstrip())
            new_middle.append(' ' * (current_indent + 4) + line.lstrip())

        # Add finally block
        new_middle.append(indent_str + 'finally:\n')
        new_middle.append(indent_str + '    wb.close()\n')

        # Replace the range
        lines[wb_idx + 1:except_idx] = new_middle
        processed += 1

    # Write result
    with open(filepath, 'w', encoding='utf-8') as f:
        f.writelines(lines)

    print(f"Processed {processed} load_workbook calls, "
          f"removed {removed_closes} wb.close() calls")


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <excel.py path>")
        sys.exit(1)
    filepath = sys.argv[1]
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        sys.exit(1)
    process_file(filepath)
