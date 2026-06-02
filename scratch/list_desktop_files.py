import os

print("Files in Desktop:")
for f in os.listdir(r'C:\Users\The Linh\Desktop'):
    if 'verify' in f or 'dam' in f.lower():
        print(f)
