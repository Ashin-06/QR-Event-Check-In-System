"""
Repair script for dashboard.html
Fixes:
  1. Corrupted lines 5310-5337 (0-indexed 5309-5336) - mid-function splice + invalid byte
  2. Duplicate/old forceQuarantineSelected at ~line 5446-5474
"""
import re

# Read file with error replacement
with open('templates/dashboard.html', encoding='utf-8', errors='replace') as f:
    content = f.read()

lines = content.split('\n')
print(f'Total lines: {len(lines)}')

# ── Step 1: Restore the damaged allPill.onclick code ─────────────────────────
# The truncated line is at index 5309 (1-indexed: 5310)
# It should be: "      window.activeSubgroupFilter = null;"
# Lines 5310-5336 (0-indexed) were replaced by our forceQuarantineSelected + corrupted byte

# Find the exact truncated line
trunc_idx = None
for i, line in enumerate(lines):
    if "document.querySelectorAll('.subgroup-pill').forEach(p => p.classList.r" in line:
        trunc_idx = i
        print(f'Found truncated line at index {i} (line {i+1})')
        break

# Find the corrupted byte line (the \ufffd</span> line)
corrupt_end_idx = None
for i in range(trunc_idx, trunc_idx + 50):
    if '\ufffd' in lines[i] or ('\u00ef\u00bf\u00bd' in lines[i]):
        corrupt_end_idx = i
        print(f'Found corrupted line at index {i} (line {i+1})')
        break

if trunc_idx is None or corrupt_end_idx is None:
    print('Could not find corruption markers - checking for partial match')
    for i, line in enumerate(lines[5295:5350], start=5295):
        print(f'  {i+1}: {line[:100]}')
else:
    print(f'Restoring lines {trunc_idx+1} through {corrupt_end_idx+1}')
    
    # What should replace lines[trunc_idx] through lines[corrupt_end_idx]
    restored = [
        "      window.activeSubgroupFilter = null;",
        "      document.querySelectorAll('.subgroup-pill').forEach(p => p.classList.remove('active'));",
        "      allPill.classList.add('active');",
        "      const dt = $('#registry-table').DataTable();",
        "      dt.draw();",
        "    };",
        "    container.appendChild(allPill);",
        "    ",
        "    // Create a pill for each unique group",
        "    Object.keys(counts).sort().forEach(groupName => {",
        "      const count = counts[groupName];",
        "      const pill = document.createElement('div');",
        "      const isThisActive = window.activeSubgroupFilter && window.activeSubgroupFilter.toLowerCase() === groupName.toLowerCase();",
        "      pill.innerHTML = `",
        "        <span class=\"group-pill-name\" style=\"font-weight: 500;\">${escapeHtml(groupName)}</span>",
        "        <span style=\"background: rgba(255,255,255,0.15); padding: 1px 5px; border-radius: 10px; font-size: 0.62rem; font-weight: 700; margin-right: 4px;\">${count}</span>",
        "        <span class=\"notify-group-icon\" style=\"opacity: 0.5; transition: opacity 0.2s; font-size: 0.75rem; margin-right: 4px; cursor: pointer;\" title=\"Notify Group\">\U0001f4e2</span>",
        "        <span class=\"edit-group-icon\" style=\"opacity: 0.5; transition: opacity 0.2s; font-size: 0.75rem; cursor: pointer;\" title=\"Rename group\">\u270f\ufe0f</span>",
    ]
    
    lines = lines[:trunc_idx] + restored + lines[corrupt_end_idx + 1:]
    print(f'Lines after step 1: {len(lines)}')

# ── Step 2: Remove old forceQuarantineSelected block ─────────────────────────
old_start = None
old_end = None
for i, line in enumerate(lines):
    if '// ── Force Quarantine selected / all checked-in' in line:
        old_start = i
        print(f'Found old forceQuarantine block at line {i+1}')
        break

if old_start is not None:
    # Find the closing }; 
    for i in range(old_start, old_start + 60):
        if lines[i].strip() == '};' or lines[i].strip() == '};':
            old_end = i
        if old_end and i > old_end + 1 and lines[i].strip() == '':
            break
    # Find the blank line after the closing block
    for i in range(old_start + 1, old_start + 60):
        if lines[i].strip() == '};':
            old_end = i
            break
    
    if old_end:
        print(f'Removing lines {old_start+1} through {old_end+1}')
        lines = lines[:old_start] + lines[old_end + 1:]
        print(f'Lines after step 2: {len(lines)}')
    else:
        print('Could not find end of old block - showing context:')
        for i in range(old_start, min(old_start + 50, len(lines))):
            print(f'  {i+1}: {lines[i][:100]}')

# ── Write fixed file ──────────────────────────────────────────────────────────
fixed = '\n'.join(lines)
# Remove any remaining replacement chars
fixed = fixed.replace('\ufffd', '')

with open('templates/dashboard.html', 'w', encoding='utf-8') as f:
    f.write(fixed)

print(f'\nDone! Final line count: {len(lines)}')

# Verify no corruption chars remain
remaining = fixed.count('\ufffd')
print(f'Remaining replacement chars: {remaining}')
