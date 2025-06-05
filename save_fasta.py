import pandas as pd
import os
excel_path = 'Your input file.xls'
output_base = 'Your output file folder'
xls = pd.ExcelFile(excel_path)
for sheet_name in xls.sheet_names:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    
    sheet_folder = os.path.join(output_base, sheet_name)
    os.makedirs(sheet_folder, exist_ok=True)
    
    text_data = '\n'.join(df[0].astype(str))
    entries = [entry.strip() for entry in text_data.split('>') if entry.strip()]
    
    for i, entry in enumerate(entries):
        lines = entry.splitlines()
        header = lines[0]
        sequence = ''.join(lines[1:])
        
        file_name = header.split('|')[1].replace('*', '_').replace(':', '_') + '.fasta'
        file_path = os.path.join(sheet_folder, file_name)
        
        with open(file_path, 'w') as f:
            f.write(f">{header}\n")
            f.write(sequence + '\n')
print("✅ Done! All sequences are saved by sheet.")
