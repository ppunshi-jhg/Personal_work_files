import csv

def merge_rows(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        rows = list(reader)
    
    merged_rows = []
    merged_row = ['', '']
    
    for row in rows[1:]:
        if not row[0].isdigit():
            merged_row[1] += ' ' + row[1]
        else:
            merged_rows.append(merged_row)
            merged_row = row
    
    merged_rows.append(merged_row)
    
    with open(output_file, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(merged_rows)

# Usage example
input_path = r'C:\Users\ppunshi\OneDrive - John Holland Group\Desktop\3307239REQIF.csv'
output_path = r'C:\Users\ppunshi\OneDrive - John Holland Group\Desktop\3307239REQIF2.csv'
merge_rows(input_path, output_path)