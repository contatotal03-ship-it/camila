import re

with open(r'c:\Users\Admin\Downloads\camila\VENTAS Y NOTAS\Conciliador\core\compras_logic.py', 'r', encoding='utf-8', errors='replace') as f:
    content = f.read()

matches = re.findall(r"\'[^\']*?mero\'|\"[^\"]*?mero\"", content)
print('Matches found:', matches)
for m in matches:
    print(m, 'Unicode points:', [ord(c) for c in m])
