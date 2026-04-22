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
                            for row in rows:
                                cols = row.find_all('td')
                                
                                # 1. Pengaman Jumlah Kolom
                                if len(cols) < 10:
                                    continue

                                hari_val = cols[0].get_text(strip=True).upper()
                                tgl_val = cols[1].get_text(strip=True).upper()

                                # 2. PENGAMAN KETAT: Lewati jika baris TOTAL atau bukan data harian
                                # Kita cek apakah kolom hari mengandung kata 'TOTAL' atau kolom tanggal kosong/isi 'TOTAL'
                                if "TOTAL" in hari_val or "TOTAL" in tgl_val or not tgl_val or tgl_val == "-":
                                    continue
                                
                                # 3. Cek apakah ini benar-benar baris data (Tanggal biasanya mengandung angka)
                                # Ini mencegah baris footer/keterangan ikut terproses
                                if not any(char.isdigit() for char in tgl_val):
                                    continue

                                masuk_val = cols[3].get_text(strip=True)
                                jam_telat_val = cols[4].get_text(strip=True)
                                mnt_telat_val = cols[5].get_text(strip=True)
                                pulang_val = cols[6].get_text(strip=True)
                                ket_val = cols[-1].get_text(strip=True).upper()

                                # Skip baris header jika terbaca di tbody
                                if "HARI" in hari_val: continue

                                potongan = 0.0
                                detail = []

                                if ket_val == 'M':
                                    if hari_val not in ['SABTU', 'MINGGU']:
                                        potongan = 3.0
                                        detail.append("Mangkir (3%)")
                                
                                elif ket_val in ['H', '', 'NAN', '*']: # * adalah libur tapi bisa ada telat/absen
                                    # 1. Cek Absen Bolong (Hanya jika hari kerja/H)
                                    if ket_val == 'H':
                                        if masuk_val in ['-', '']:
                                            potongan += 1.5
                                            detail.append("Tdk Absen Masuk (1.5%)")
                                        if pulang_val in ['-', '']:
                                            potongan += 1.5
                                            detail.append("Tdk Absen Pulang (1.5%)")

                                    # 2. Cek Telat
                                    try:
                                        def to_f(v):
                                            v = v.replace('-', '0').strip()
                                            return float(v) if v else 0.0
                                        
                                        tot_m = (to_f(jam_telat_val) * 60) + to_f(mnt_telat_val)
                                        if tot_m > 0:
                                            if 1 <= tot_m <= 15: p = 0.25
                                            elif 16 <= tot_m <= 60: p = 0.5
                                            elif 61 <= tot_m <= 120: p = 1.0
                                            else: p = 1.5
                                            potongan += p
                                            detail.append(f"Telat {int(tot_m)}m ({p}%)")
                                    except: pass

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
                    st.error(f"Error {nama}: {e}")

            status_text.success("Selesai!")

            if all_results:
                df_detail = pd.DataFrame(all_results)
                
                # --- BAGIAN TOTAL PER PEGAWAI ---
                st.write("### 📊 Ringkasan Total Potongan")
                df_summary = df_detail.groupby("Nama")["Potongan (%)"].sum().reset_index()
                df_summary.columns = ["Nama Pegawai", "Total Potongan (%)"]
                # Tambahkan styling agar lebih menarik
                st.dataframe(df_summary.style.highlight_max(axis=0, color='#ffcccc'), use_container_width=True)

                st.write("### 📝 Rincian Harian")
                st.dataframe(df_detail, use_container_width=True)

                # Export ke Excel dengan dua Sheet
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                    df_summary.to_excel(writer, sheet_name='Summary', index=False)
                    df_detail.to_excel(writer, sheet_name='Detail_Harian', index=False)
                
                st.download_button(
                    label="Download Laporan Lengkap (Excel)",
                    data=buf.getvalue(),
                    file_name=f"rekap_potongan_{row_peg['Bulan']}_{row_peg['Tahun']}.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.info("Tidak ditemukan potongan pada periode ini.")