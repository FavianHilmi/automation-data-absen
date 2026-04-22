import streamlit as st
import pandas as pd
import requests
import io
import zipfile

st.set_page_config(page_title="Automation Data Absen", layout="centered")

st.title("Automation Data Absen")

BASE_URL = "https://teko-cak.surabaya.go.id/cetak_new/lap_per_pegawai2/"
ID_INSTANSI = "3.10.00.00.00"

st.sidebar.header("Navigasi")
menu = st.sidebar.radio("Pilih Menu:", ["Download Template", "Proses Download Data"])

if menu == "Download Template":
    st.subheader("Download Template Excel")
    st.write("Silakan download template Excel berikut untuk mengisi data pegawai yang akan diproses:")
    data_template = {
        'Nama_Pegawai': ['ACHMAD SHOLIKIN', 'CONTOH PEGAWAI 2'],
        'ID_Pegawai': ['7f66455c-ee57-11ea-8acf-000c29766abb', '35951734-ece8-11ea-9cd7-000c29766abb'],
        'Tanggal_Akhir': [31, 31], 'Bulan': ['03', '03'], 'Tahun': [2026, 2026]
    }
    df_template = pd.DataFrame(data_template)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_template.to_excel(writer, index=False)
    
    st.download_button("Download Template (.xlsx)", data=buffer.getvalue(), file_name="template_absen.xlsx")

else:
    st.subheader("Proses Download Data Absen")
    
    col1, col2 = st.columns(2)
    with col1:
        file_type = st.selectbox("Format Laporan:", ["pdf", "xls"])
    with col2:
        output_mode = st.radio("Metode Download:", ["File ZIP (Rekomendasi)", "Satu per Satu"])

    uploaded_file = st.file_uploader("Upload Excel Data Pegawai:", type=["xlsx", "xls"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df['Bulan'] = df['Bulan'].astype(str).str.zfill(2)
        df['Tanggal_Akhir'] = df['Tanggal_Akhir'].astype(str).str.zfill(2)
        
        st.write(f"Total data: **{len(df)} pegawai**")

        if st.button("Mulai Proses"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            zip_buffer = io.BytesIO()
            
            success_files = []

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for index, row in df.iterrows():
                    nama = row['Nama_Pegawai']
                    status_text.text(f"Mendownload: {nama}...")
                    
                    params = {
                        'tgl': row['Tanggal_Akhir'], 'bulan': row['Bulan'], 'tahun': row['Tahun'],
                        'id_instansi': ID_INSTANSI, 'id_pegawai': row['ID_Pegawai'], 'type': file_type
                    }

                    try:
                        response = requests.get(BASE_URL, params=params, timeout=30)
                        if response.status_code == 200:
                            ext = "pdf" if file_type == "pdf" else "xls"
                            file_name = f"Laporan_{nama}_{params['tgl']}-{params['bulan']}.{ext}"
                            
                            if output_mode == "File ZIP (Rekomendasi)":
                                zip_file.writestr(file_name, response.content)
                            else:
                                success_files.append({"name": file_name, "content": response.content})
                        else:
                            st.error(f"Gagal: {nama} (Status: {response.status_code})")
                    except Exception as e:
                        st.warning(f"Error pada {nama}: {e}")
                    
                    progress_bar.progress((index + 1) / len(df))

            status_text.success("Proses selesai")

            if output_mode == "File ZIP (Rekomendasi)":
                st.download_button(
                    label="Download Semua (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Laporan_Absen_{file_type}.zip",
                    mime="application/zip"
                )
            else:
                for f in success_files:
                    st.download_button(label=f"Simpan {f['name']}", data=f['content'], file_name=f['name'])     

Tentu, ini ide yang bagus untuk melengkapi tool automasi kamu. Karena kamu sudah menggunakan Pandas dan Streamlit, kita bisa menambahkan menu baru dengan logika perhitungan yang kamu minta.

Berikut adalah tambahan kode untuk menu "Hitung Potongan". Kamu bisa menyisipkan bagian ini ke dalam struktur if-elif pada navigasi sidebar kamu.

Tambahan Logika Menu "Hitung Potongan"
Python
elif menu == "Hitung Potongan":
    st.subheader("Hitung Persentase Potongan Kehadiran")
    st.write("Upload file laporan (XLS) yang sudah didownload untuk dihitung potongannya.")

    uploaded_files = st.file_uploader("Upload File XLS Laporan:", type=["xls"], accept_multiple_files=True)

    if uploaded_files:
        all_results = []

        for uploaded_file in uploaded_files:
            # Baca excel mulai dari baris yang berisi data (header di baris ke-10 / index 9)
            df_absen = pd.read_excel(uploaded_file, skiprows=9)
            
            # Ambil Nama Pegawai dari cell B4 (karena skiprows 9, koordinat berubah)
            # Cara lebih aman: baca ulang tanpa skiprows hanya untuk ambil Nama
            df_info = pd.read_excel(uploaded_file, header=None, nrows=10)
            nama_pegawai = df_info.iloc[3, 1] # Cell B4

            for index, row in df_absen.iterrows():
                # Berhenti jika bertemu baris "TOTAL"
                if str(row['HARI']).strip() == "TOTAL" or pd.isna(row['HARI']):
                    break
                
                potongan = 0.0
                keterangan_potongan = []

                # 1. Logika Tidak Masuk (Simbol '*' biasanya berarti tidak masuk/libur tanpa keterangan)
                # Sesuaikan simbol ini dengan data asli di kolom 'ETERANGA'
                if row['ETERANGA'] == '*':
                    # Cek apakah ini hari kerja atau hari libur (Minggu/Sabtu)
                    if row['HARI'] not in ['MINGGU', 'SABTU']:
                        potongan = 3.0
                        keterangan_potongan.append("Tidak Masuk (3%)")
                
                else:
                    # 2. Cek Tidak Absen Masuk atau Pulang
                    if pd.isna(row['MASUK']):
                        potongan += 1.5
                        keterangan_potongan.append("Tidak Absen Masuk (1.5%)")
                    if pd.isna(row['PULANG']):
                        potongan += 1.5
                        keterangan_potongan.append("Tidak Absen Pulang (1.5%)")

                    # 3. Logika Terlambat (Kolom 'TELAT MASUK' JAM & MENIT)
                    jam_telat = row.get('JAM', 0) if pd.notna(row.get('JAM')) else 0
                    menit_telat = row.get('MENIT', 0) if pd.notna(row.get('MENIT')) else 0
                    total_menit = (jam_telat * 60) + menit_telat

                    if total_menit > 0:
                        if 1 <= total_menit <= 15:
                            potongan += 0.25
                            keterangan_potongan.append(f"Telat {total_menit}m (0.25%)")
                        elif 16 <= total_menit <= 60:
                            potongan += 0.5
                            keterangan_potongan.append(f"Telat {total_menit}m (0.5%)")
                        elif 61 <= total_menit <= 120:
                            potongan += 1.0
                            keterangan_potongan.append(f"Telat {total_menit}m (1%)")
                        elif total_menit > 120:
                            potongan += 1.5
                            keterangan_potongan.append(f"Telat {total_menit}m (1.5%)")

                if potongan > 0:
                    all_results.append({
                        "Nama Pegawai": nama_pegawai,
                        "Tanggal": row['TANGGAL'],
                        "Hari": row['HARI'],
                        "Potongan (%)": potongan,
                        "Detail": ", ".join(keterangan_potongan)
                    })

        if all_results:
            df_final = pd.DataFrame(all_results)
            st.write("### Hasil Perhitungan Potongan")
            st.dataframe(df_final)

            # Export options
            col_dl1, col_dl2 = st.columns(2)
            
            # Export Excel
            buffer_excel = io.BytesIO()
            with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Potongan')
            
            col_dl1.download_button("Download Hasil (Excel)", data=buffer_excel.getvalue(), file_name="rekap_potongan.xlsx")
            
            st.info("Catatan: Logika 'Tidak Masuk' mendeteksi simbol '*' pada hari kerja (Senin-Jumat).")
        else:
            st.success("Tidak ditemukan potongan pada file yang diupload.")