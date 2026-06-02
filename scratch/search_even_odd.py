import sys

sys.stdout.reconfigure(encoding='utf-8')
with open(r'C:\Users\The Linh\Desktop\Auto-Word\docx_processor.py', 'r', encoding='utf-8') as f:
    content = f.read()

count = 0
for idx, line in enumerate(content.splitlines()):
    if 'evenAndOddHeaders' in line:
        print(f"Line {idx+1}: {line.strip()}")
        count += 1
print(f"Total matches: {count}")
