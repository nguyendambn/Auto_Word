import os

print("Searching Desktop and subfolders for docx:")
for root, dirs, files in os.walk(r'C:\Users\The Linh\Desktop'):
    for f in files:
        if f.endswith('.docx') and 'verify' in f.lower():
            print(os.path.join(root, f))
