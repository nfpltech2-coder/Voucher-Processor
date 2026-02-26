
with open(r'c:\Projects\Reimbursement\reimbursement_app.py', 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

target = 'command=lambda val: [e.set(val) for e in trans_combos]'
replacement = 'command=lambda val, tc=trans_combos: [e.set(val) for e in tc]'

if target in content:
    new_content = content.replace(target, replacement)
    with open(r'c:\Projects\Reimbursement\reimbursement_app.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print("Successfully replaced.")
else:
    print("Target not found.")
