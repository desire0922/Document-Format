import sys
sys.stdout.reconfigure(encoding="utf-8")
lines = open("UI对接.py", "r", encoding="utf-8").readlines()
for i in [444, 581, 1079, 1572, 1582]:
    if i < len(lines):
        ln = lines[i]
        if "QPushButton:hover" in ln and '") QPushButton' in ln:
            print(f"Line {i+1}: BROKEN - hover outside string")
        elif "QPushButton:hover" in ln:
            print(f"Line {i+1}: OK - hover inside string")
        else:
            print(f"Line {i+1}: NO HOVER")
