import os
from tkinter.filedialog import askopenfilename

file_name = askopenfilename()

if file_name:
    file = b''
    with open(file_name, 'rb') as _file:
        file = _file.read()
    
    text = f"bathInput = {file}"
    
    with open("planilha.py", 'w')as _file:
        _file.write(text)