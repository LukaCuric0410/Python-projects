import pandas as pd
import os

# definicija putanje
file_path = '/Users/CuricL/Desktop/Book2.xlsx'  # stavite svoju putanju

# direktorij di ce se spremiti csv fajlovi
output_directory = '/Users/CuricL/Desktop'  # zamjeni s svojom lokacijom

# provjera direktorija za spremanje ako ga nema napravi ga
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Učitaj Excel datoteku
excel_file = pd.ExcelFile(file_path)

# Petlja kroz svaki sheet u Excel datoteci
for sheet_name in excel_file.sheet_names:
    # Učitaj sheet u DataFrame
    df = excel_file.parse(sheet_name)

    # definicija putanje i ime CSV datoteke
    csv_file_path = os.path.join(output_directory, f"{sheet_name}.csv")

    # Sačuvaj DataFrame kao CSV datoteku u UTF-8 formatu
    df.to_csv(csv_file_path, index=False, encoding='utf-8-sig')

    print(f"Spremio: {csv_file_path}")

print("Sve datoteke su uspješno spremljene!")
