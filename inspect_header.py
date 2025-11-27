import csv

file_path = "miro_export_report.csv"

try:
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        header = next(reader)
        print(f"Header: {header}")
        
        print("First 5 rows:")
        for i, row in enumerate(reader):
            if i >= 5: break
            print(f"Row {i}: {row}")
except Exception as e:
    print(f"Error: {e}")
