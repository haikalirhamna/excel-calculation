import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment

# Folder berisi file excel
folder_path = 'resources'
output_folder = 'output'  # Folder output untuk menyimpan file hasil

# Pastikan folder output ada, jika tidak, buat folder tersebut
os.makedirs(output_folder, exist_ok=True)

# Fungsi untuk menghitung skor
def calculate_score(row):
    attendance = row['Nilai Absensi']
    task = row['Nilai Bekerja']
    return ((attendance + task) / 192) * 100

# List semua file di folder, abaikan file sementara yang dimulai dengan `~$`
file_list = [file for file in os.listdir(folder_path) if file.endswith('.xlsx') and not file.startswith('~$')]

# List untuk menyimpan data hasil
final_data = []

# Loop melalui semua file dan proses datanya
for file in file_list:
    # Ambil nama bulan dengan menghapus bagian 'updated_bulan_' dan kapitalisasi nama bulan
    month_name = os.path.splitext(file)[0].replace('updated_bulan_', '').capitalize()
    
    # Baca file Excel
    df = pd.read_excel(os.path.join(folder_path, file))
    
    # Ganti nilai '-' dengan 0 di kolom "Nilai Bekerja" dan "Nilai Absensi"
    df['Nilai Bekerja'] = pd.to_numeric(df['Nilai Bekerja'].replace('-', 0), errors='coerce').fillna(0)
    df['Nilai Absensi'] = pd.to_numeric(df['Nilai Absensi'].replace('-', 0), errors='coerce').fillna(0)
    
    # Mengelompokkan data berdasarkan kolom "Nama"
    grouped_data = df.groupby('Nama').agg({
        'Pekerjaan': 'first',
        'Nilai Bekerja': 'sum',
        'Nilai Absensi': 'sum'
    }).reset_index()

    # Hitung skor untuk setiap individu
    grouped_data['Score'] = grouped_data.apply(calculate_score, axis=1)
    
    # Tambahkan nama bulan ke data
    grouped_data['Bulan'] = month_name
    
    # Pilih kolom yang dibutuhkan dan masukkan ke final_data
    final_data.append(grouped_data[['Nama', 'Pekerjaan', 'Bulan', 'Score']])

# Menggabungkan semua DataFrame dalam list menjadi satu DataFrame
final_data = pd.concat(final_data, ignore_index=True)

# Mengonversi nilai "Score" menjadi format persentase
final_data['Score'] = final_data['Score'].apply(lambda x: f'{x:.2f}%')

# Tentukan nama file output
output_file = os.path.join(output_folder, 'output_calculations.xlsx')

# Menyimpan hasil ke file excel baru
final_data.to_excel(output_file, index=False)

# Membuka file Excel yang baru saja dibuat untuk memodifikasi lebar kolom dan format
wb = load_workbook(output_file)
ws = wb.active

# Menyesuaikan lebar kolom agar otomatis sesuai dengan panjang isinya
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter  # Dapatkan huruf kolom
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)  # Menambahkan sedikit ruang ekstra
    ws.column_dimensions[column].width = adjusted_width

# Menyimpan file yang telah diperbarui
wb.save(output_file)