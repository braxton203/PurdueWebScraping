import openpyxl
import re


file_path = "/Users/braxtonconley/Downloads/Proof of Concept.xlsx"
sheet_name = "Chemicals"

wb = openpyxl.load_workbook(file_path)
ws = wb[sheet_name] 

temperature_set = set()  
pattern = r"<\s*(\d{1,4}C)"  # Regex pattern to match "< numberC" (1-4 digit numbers)
counter = 0

for row in ws.iter_rows():
    for cell in row:
        if isinstance(cell.value, str): 
            matches = re.findall(pattern, cell.value) 
            for match in matches:
                counter += 1
                temperature_set.add(match) 


temperature_list = sorted(temperature_set, key=lambda x: int(x[:-1])) 
print("Unique Temperatures Found:", temperature_list)
print("Total: ", counter)
