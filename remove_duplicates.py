"""
Remove duplicate routes block from app.py
Lines 5320-5871 (1-indexed) are duplicate routes that appear again after line 5322
"""

with open('app.py', encoding='utf-8') as f:
    lines = f.readlines()

print(f'Total lines before: {len(lines)}')

# Lines are 0-indexed in the list
# Remove lines 5319 (idx) to 5870 (idx) inclusive (1-indexed 5320 to 5871)
start_idx = 5319  # 0-indexed start (line 5320)
end_idx = 5870    # 0-indexed end inclusive (line 5871)

# Verify what we're cutting
print(f'Removing lines {start_idx+1} to {end_idx+1}:')
print(f'  First: {lines[start_idx].rstrip()}')
print(f'  Last:  {lines[end_idx].rstrip()}')
print(f'  Next:  {lines[end_idx+1].rstrip()}')

new_lines = lines[:start_idx] + lines[end_idx+1:]
print(f'Total lines after: {len(new_lines)}')

with open('app.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print('Done!')

# Quick syntax check
import ast
try:
    ast.parse(''.join(new_lines))
    print('Syntax check: OK')
except SyntaxError as e:
    print(f'Syntax error: {e}')
