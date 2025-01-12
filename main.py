import os
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF

# Inisialisasi variabel
ideal_work_duration_seconds = 691200  # 192 jam dalam detik
ideal_task_score = 72  # Bobot ideal tugas
beta_0 = 0  # Intercept
beta_1 = 0.4  # Koefisien untuk Score Absensi
beta_2 = 0.6  # Koefisien untuk Score Tugas

# Tentukan jalur folder resources dan output
resources_path = os.path.join(os.getcwd(), 'resources')
output_path = os.path.join(os.getcwd(), 'output')

# Membuat folder output jika belum ada
if not os.path.exists(output_path):
    os.makedirs(output_path)

# Fungsi untuk memproses setiap file Excel
def process_excel(file_path, month):
    try:
        df = pd.read_excel(file_path)

        # Mapping status absensi ke nilai numerik
        status_mapping = {
            "hadir": 1,
            "sakit": 0.5,
            "izin": 0.25,
            "tidak hadir": 0
        }
        df['Status Absensi'] = df['Status Absensi'].map(status_mapping)

        # Mapping bobot tugas
        bobot_tugas_mapping = {
            "low": 1,
            "middle": 2,
            "high": 3
        }
        df['Bobot Tugas'] = df['Bobot Tugas'].map(bobot_tugas_mapping)

        # Replace data waktu menjadi time
        df['Scan Masuk 1'] = pd.to_datetime(df['Scan Masuk 1'], format='%H:%M:%S', errors='coerce')
        df['Scan Pulang 1'] = pd.to_datetime(df['Scan Pulang 1'], format='%H:%M:%S', errors='coerce')

        # Mengubah menjadi detik
        scan_in_seconds = (df['Scan Masuk 1'].dt.hour * 3600 +
                           df['Scan Masuk 1'].dt.minute * 60 +
                           df['Scan Masuk 1'].dt.second)
        scan_out_seconds = (df['Scan Pulang 1'].dt.hour * 3600 +
                            df['Scan Pulang 1'].dt.minute * 60 +
                            df['Scan Pulang 1'].dt.second)

        df['Durasi (detik)'] = (scan_out_seconds - scan_in_seconds)

        # Menghitung Score Absensi
        df['Attendance'] = df['Durasi (detik)'] * df['Status Absensi']

        # Grouping berdasarkan Nama
        grouped = df.groupby('Nama').agg({
            'Posisi': 'first',
            'Attendance': 'sum',
            'Bobot Tugas': 'sum',
            'Tugas': 'count',
            'Status Tugas': lambda x: (x == 'selesai').sum()  # Menghitung jumlah "selesai"
        }).reset_index()

        # Menghitung Attendance Percentage
        grouped['Attendance Percentage'] = (grouped['Attendance'] / ideal_work_duration_seconds) * 100
        grouped['Score Absensi'] = (grouped['Attendance Percentage'] * beta_1)

        # Menghitung Task Percentage
        grouped['Task Percentage'] = ((grouped['Bobot Tugas'] * grouped['Status Tugas']) / 
                                      (ideal_task_score * grouped['Tugas'])) * 100
        grouped['Score Tugas'] = (grouped['Task Percentage'] * beta_2)

        # Menghitung hasil regresi linear
        grouped['Regresi Linear'] = beta_0 + (grouped['Attendance Percentage'] * beta_1) + (
                grouped['Task Percentage'] * beta_2)

        # Menambahkan informasi bulan
        grouped['Month'] = month

        return grouped

    except Exception as e:
        print(f"Kesalahan pada file {os.path.basename(file_path)}: {e}")
        return None

# Loop untuk membaca dan memproses semua file Excel dalam folder 'resources'
all_results = []
months = []  # Menyimpan bulan untuk setiap file
for filename in os.listdir(resources_path):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(resources_path, filename)
        month = filename.split('.')[0]  # Menyimpan nama bulan dari nama file (asumsi nama file = bulan)
        results = process_excel(file_path, month)
        if results is not None:
            all_results.append(results)
            months.append(month)

# Gabungkan semua hasil menjadi satu DataFrame
all_results_df = pd.concat(all_results)

# Fungsi untuk membuat file PDF dengan hasil kalkulasi dan tabel
def create_pdf_with_results(dataframe, output_path):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Menambahkan halaman pertama
    pdf.add_page()
    
    # Menetapkan font untuk teks dalam PDF
    pdf.set_font("Arial", size=12)
    
    # Judul
    pdf.cell(200, 10, txt="Hasil Kalkulasi Skor Absensi dan Tugas", ln=True, align="C")
    pdf.ln(10)  # Menambah jarak setelah judul

    # Mendapatkan data berdasarkan bulan
    months = dataframe['Month'].unique()

    for month in months:
        # Menambahkan subjudul untuk setiap bulan
        pdf.set_font("Arial", size=12, style='B')
        pdf.cell(200, 10, txt=f"Hasil Bulan: {month}", ln=True, align="L")
        pdf.ln(5)  # Menambah jarak sebelum tabel
        
        # Tabel untuk data bulan ini
        monthly_data = dataframe[dataframe['Month'] == month]
        
        # Menambahkan tabel hasil
        pdf.set_font("Arial", size=10)
        pdf.cell(40, 10, 'Nama', border=1, align='C')
        pdf.cell(40, 10, 'Posisi', border=1, align='C')
        pdf.cell(40, 10, 'Score Absensi', border=1, align='C')
        pdf.cell(40, 10, 'Score Tugas', border=1, align='C')
        pdf.cell(40, 10, 'Regresi Linear', border=1, align='C')
        pdf.ln()

        for index, row in monthly_data.iterrows():
            pdf.cell(40, 10, row['Nama'], border=1, align='C')
            pdf.cell(40, 10, row.get('Posisi', 'N/A'), border=1, align='C')
            pdf.cell(40, 10, f"{row['Score Absensi']:.2f}", border=1, align='C')
            pdf.cell(40, 10, f"{row['Score Tugas']:.2f}", border=1, align='C')
            pdf.cell(40, 10, f"{row['Regresi Linear']:.2f}", border=1, align='C')
            pdf.ln()

    pdf_output_path = os.path.join(output_path, "hasil_kalkulasi.pdf")
    pdf.output(pdf_output_path)
    print("PDF berhasil dibuat di", pdf_output_path)

# Loop untuk setiap orang dan buatkan grafik regresi linear masing-masing dan simpan ke folder output
for person in all_results_df['Nama'].unique():
    person_data = all_results_df[all_results_df['Nama'] == person]
    
    # Visualisasi grafik per orang
    plt.figure(figsize=(10, 6))
    plt.plot(person_data['Month'], person_data['Score Absensi'], label='Score Absensi', marker='o')
    plt.plot(person_data['Month'], person_data['Score Tugas'], label='Score Tugas', marker='s')
    plt.plot(person_data['Month'], person_data['Regresi Linear'], label='Regresi Linear', marker='x')

    plt.title(f'Score for {person}: Attendance, Task, and Linear Regression')
    plt.xlabel('Month')
    plt.ylabel('Percentage (%)')
    plt.xticks(rotation=45)
    plt.legend()
    plt.grid(True)
    plt.tight_layout()

    # Simpan grafik ke dalam folder output dengan nama file yang sesuai
    output_file_path = os.path.join(output_path, f"{person}_score_plot.png")
    plt.savefig(output_file_path)

    # Menutup plot agar tidak ada masalah dengan grafik berikutnya
    plt.close()

# Panggil fungsi untuk membuat PDF setelah semua perhitungan selesai
create_pdf_with_results(all_results_df, output_path)
