import os
import tempfile
import sys

sys.stdout.reconfigure(encoding='utf-8')
print("Temp dir:", tempfile.gettempdir())
for f in os.listdir(tempfile.gettempdir()):
    if 'verify' in f.lower() or 'autoword' in f.lower():
        print(repr(f))
