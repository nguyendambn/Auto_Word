import os

print("cwd:", os.getcwd())
print("Files in cwd:")
for f in os.listdir('.'):
    print(f)
print("Files in C:\\Users\\The Linh\\Desktop\\Auto-Word:")
try:
    for f in os.listdir(r'C:\Users\The Linh\Desktop\Auto-Word'):
        print(f)
except Exception as e:
    print(e)
