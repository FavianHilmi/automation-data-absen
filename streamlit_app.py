import streamlit as st
import pandas as pd
import requests
import io
import zipfile
from bs4 import BeautifulSoup

st.set_page_config(page_title="Automation Data Absen", layout="centered")

st.title("Automation Data Absen")

BASE_URL = "https://teko-cak.surabaya.go.id/cetak_new/lap_per_pegawai2/"
ID_INSTANSI = "3.10.00.00.00"

st.sidebar.header("Navigasi")
menu = st.sidebar.radio("Pilih Menu:", ["Download Template", "Proses Download Data" , "Hitung Potongan"])

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

elif menu == "Proses Download Data":
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

elif menu == "Hitung Potongan":
    st.subheader("Hitung Potongan via Web Scraping")
    st.write("Upload template excel untuk mengambil data langsung dari server.")

    uploaded_file = st.file_uploader("Upload Template Pegawai:", type=["xlsx", "xls"])

    if uploaded_file:
        df_pegawai = pd.read_excel(uploaded_file)
        all_results = []

        if st.button("Mulai Hitung Potongan"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, row_peg in df_pegawai.iterrows():
                nama = row_peg['Nama_Pegawai']
                status_text.text(f"Menganalisis: {nama}...")
                
                params = {
                    'tgl': row_peg['Tanggal_Akhir'],
                    'bulan': str(row_peg['Bulan']).zfill(2),
                    'tahun': row_peg['Tahun'],
                    'id_instansi': ID_INSTANSI,
                    'id_pegawai': row_peg['ID_Pegawai']
                }
                
                try:
                    response = requests.get(BASE_URL, params=params, timeout=30)
                    if response.status_code == 200:
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(response.text, 'html.parser')
    table = soup.find('table', {'border': '1'})
    
    if table:
        rows = table.find('tbody').find_all('tr')
        
        # --- DEBUG START ---
        # Lihat baris pertama yang ditemukan untuk memastikan index kolom benar
        test_cols = rows[0].find_all('td')
        st.write(f"DEBUG: Baris pertama ditemukan {len(test_cols)} kolom.")
        # Menampilkan isi setiap kolom agar kamu tahu index mana yang berisi jam/menit telat
        debug_data = {f"Col {i}": col.get_text(strip=True) for i, col in enumerate(test_cols)}
        st.json(debug_data) 
        # --- DEBUG END ---

        for row in rows:
            cols = row.find_all('td')
            # Lewati baris yang bukan data harian (seperti TOTAL atau baris kosong)
            if len(cols) < 10:
                continue

            hari_val = cols[0].get_text(strip=True).upper()
            jam_telat_raw = cols[4].get_text(strip=True)
            mnt_telat_raw = cols[5].get_text(strip=True)

            # Debug spesifik untuk melihat angka telat yang terbaca
            if jam_telat_raw != "-" or mnt_telat_raw != "-":
                st.write(f"🔍 Terdeteksi Potensi Telat: {hari_val} | Jam: {jam_telat_raw} | Menit: {mnt_telat_raw}")

                                # --- LOGIKA POTONGAN ---
                                if ket_val == 'M':
                                    if hari_val not in ['SABTU', 'MINGGU']:
                                        potongan = 3.0
                                        detail.append("Mangkir (3%)")
                                
                                elif ket_val in ['H', '', 'NAN']:
                                    # 1. Cek Absen Bolong
                                    if masuk_val in ['-', '']:
                                        potongan += 1.5
                                        detail.append("Tdk Absen Masuk (1.5%)")
                                    if pulang_val in ['-', '']:
                                        potongan += 1.5
                                        detail.append("Tdk Absen Pulang (1.5%)")

                                    # 2. Cek Telat (Ini bagian krusialnya)
                                    try:
                                        def to_float(val):
                                            val = val.replace('-', '0').strip()
                                            return float(val) if val else 0.0

                                        j_telat = to_float(jam_telat_val)
                                        m_telat = to_float(mnt_telat_val)
                                        tot_menit = (j_telat * 60) + m_telat

                                        if tot_menit > 0:
                                            p_telat = 0
                                            if 1 <= tot_menit <= 15: p_telat = 0.25
                                            elif 16 <= tot_menit <= 60: p_telat = 0.5
                                            elif 61 <= tot_menit <= 120: p_telat = 1.0
                                            else: p_telat = 1.5
                                            
                                            potongan += p_telat
                                            detail.append(f"Telat {int(tot_menit)}m ({p_telat}%)")
                                    except:
                                        pass

                                if potongan > 0:
                                    all_results.append({
                                        "Nama": nama,
                                        "Tanggal": tgl_val,
                                        "Hari": hari_val,
                                        "Potongan (%)": potongan,
                                        "Alasan": " + ".join(detail)
                                    })

                    progress_bar.progress((idx + 1) / len(df_pegawai))
                except Exception as e:
                    st.error(f"Gagal di {nama}: {e}")

            status_text.success("Selesai!")
            if all_results:
                res_df = pd.DataFrame(all_results)
                st.dataframe(res_df)
                
                # Download Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    res_df.to_excel(writer, index=False)
                st.download_button("Download Hasil", output.getvalue(), "potongan_absen.xlsx")
            else:
                st.info("Tidak ada potongan ditemukan.")