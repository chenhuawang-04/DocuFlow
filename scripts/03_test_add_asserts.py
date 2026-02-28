#!/usr/bin/env python3
"""H7 Fix: Add assertions to test files.

Transforms:
    result = SomeOp.method(...)
    print(f"    结果: {result}")

Into:
    result = SomeOp.method(...)
    assert result["success"], f"Failed: {result}"

Also fixes M2 (hardcoded paths) by replacing absolute paths with tmp_path.
"""

import re
import os
import sys


def fix_test_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    original = content
    changes = 0

    # Pattern: print(f"    结果: {result}") or print(f"结果: {result}")
    # Replace with assert
    pattern = re.compile(
        r'(\s*)print\(f?"?\s*(?:结果|result|Result)[:：]?\s*\{(\w+)\}"?\)',
        re.IGNORECASE
    )

    def replacer(m):
        nonlocal changes
        indent = m.group(1)
        var = m.group(2)
        changes += 1
        return f'{indent}assert {var}.get("success"), f"Failed: {{{var}}}"'

    content = pattern.sub(replacer, content)

    # Fix hardcoded absolute paths (M2)
    content = content.replace(
        'sys.path.insert(0, os.path.join(os.path.dirname(__file__), \'src\'))',
        'sys.path.insert(0, os.path.join(os.path.dirname(__file__), \'..\', \'src\'))'
    )

    # Replace hardcoded E:/Project/DocuFlow paths with relative
    content = re.sub(
        r'"E:/Project/DocuFlow/test_output/([^"]+)"',
        r'os.path.join(os.path.dirname(__file__), "..", "test_output", "\1")',
        content
    )

    if content != original:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)

    return changes


if __name__ == '__main__':
    test_dir = os.path.join(os.path.dirname(__file__), '..', 'tests')
    test_dir = os.path.normpath(test_dir)

    total = 0
    for fname in sorted(os.listdir(test_dir)):
        if fname.startswith('test_') and fname.endswith('.py'):
            fpath = os.path.join(test_dir, fname)
            n = fix_test_file(fpath)
            if n:
                print(f"  {fname}: {n} assertions added")
            total += n

    print(f"Total: {total} assertions added")
