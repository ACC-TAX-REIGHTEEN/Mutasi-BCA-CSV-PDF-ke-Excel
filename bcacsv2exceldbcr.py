import pandas as pd
import os
import glob
import re
from openpyxl.utils import get_column_letter 

def split_db_cr_columns(df):
    processed_data = []

    pattern = re.compile(r'^\s*([\d.,]+)\s*(DB|CR)\s*$', re.IGNORECASE)

    for col in df.columns:
        col_data = df[col].astype(str)
        sample = col_data.head(20)
        matches = sample.apply(lambda x: bool(pattern.match(x)) if x and x.lower() != 'nan' else False)
        
        if matches.any():
            print(f"      -> Mendeteksi kolom '{col}' sebagai format DB/CR. Memecah kolom...")
            
            db_values = []
            cr_values = []
            
            for val in col_data:
                match = pattern.match(val) if isinstance(val, str) else None
                if match:
                    nominal_str = match.group(1).replace(',', '')
                    tipe = match.group(2).upper()
                    
                    try:
                        nominal = float(nominal_str)
                    except ValueError:
                        nominal = 0
                        
                    if tipe == 'DB':
                        db_values.append(nominal)
                        cr_values.append(None) 
                    elif tipe == 'CR':
                        db_values.append(None)
                        cr_values.append(nominal)
                else:
                    db_values.append(None)
                    cr_values.append(None)
                    
            temp_df = pd.DataFrame({'DB': db_values, 'CR': cr_values})
            processed_data.append(temp_df)
            
        else:
            processed_data.append(df[col])

    if processed_data:
        return pd.concat(processed_data, axis=1)
    return df

def convert_csv_to_excel_autofit():
    current_folder = os.path.dirname(os.path.abspath(__file__))
    print(f"Mencari file CSV di folder: {current_folder}")

    csv_files = glob.glob(os.path.join(current_folder, "*.csv"))

    if not csv_files:
        print("Tidak ditemukan file CSV di folder ini.")
        return

    print(f"Ditemukan {len(csv_files)} file CSV. Memulai proses...\n")

    for input_path in csv_files:
        try:
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_filename = f"{base_name}.xlsx"
            output_path = os.path.join(current_folder, output_filename)

            print(f"-> Memproses: {os.path.basename(input_path)} ...")

            df = pd.read_csv(input_path, sep=None, engine='python', quotechar='"', header=None)

            def clean_text(text):
                if isinstance(text, str):
                    return text.strip().strip(',')
                return text
            df_clean = df.map(clean_text)
            df_final = split_db_cr_columns(df_clean)

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                has_named_columns = any(isinstance(c, str) for c in df_final.columns)
                
                df_final.to_excel(writer, index=False, header=has_named_columns, sheet_name='Sheet1')
                
                worksheet = writer.sheets['Sheet1']

                for column in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    
                    for cell in column:
                        try:
                            if cell.value:
                                cell_len = len(str(cell.value))
                                if cell_len > max_length:
                                    max_length = cell_len
                        except:
                            pass
                    
                    adjusted_width = (max_length + 2) 
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            print(f"   [BERHASIL] Disimpan & Auto-fit: {output_filename}")

        except Exception as e:
            print(f"   [GAGAL] Error pada file {os.path.basename(input_path)}: {e}")
        
        print("-" * 40)

    print("\nSemua proses selesai!")

if __name__ == "__main__":
    convert_csv_to_excel_autofit()
