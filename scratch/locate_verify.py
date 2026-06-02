import os

path = r'C:\Users\The Linh\Desktop\Auto-Word\_verify_v2.docx'
print("Exists in C:\\Users\\The Linh\\Desktop\\Auto-Word\\_verify_v2.docx?", os.path.exists(path))
if os.path.exists(path):
    print("Size:", os.path.getsize(path))
else:
    # Let's search the whole user folder for verify files
    print("Searching for files...")
    for root, dirs, files in os.walk(r'C:\Users\The Linh'):
        for f in files:
            if 'verify' in f.lower() and f.endswith('.docx'):
                print(os.path.join(root, f))
