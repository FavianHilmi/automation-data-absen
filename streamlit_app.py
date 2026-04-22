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

    out_format = st.selectbox("Pilih Format Output:", ["Excel"])
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
                        # Parsing semua tabel di HTML
                        all_tables = pd.read_html(io.StringIO(response.text))
                        
                        # CARI TABEL ABSEN: Cari tabel yang punya kata 'TANGGAL' dan 'HARI'
                        df_absen = None
                        for t in all_tables:
                            t_str = t.astype(str)
                            if t_str.apply(lambda r: r.str.contains('TANGGAL').any(), axis=1).any():
                                df_absen = t
                                break
                        
                        if df_absen is not None:
                            # 1. Cari baris header (yang mengandung teks 'HARI')
                            mask = df_absen.apply(lambda r: r.astype(str).str.contains('HARI').any(), axis=1)
                            header_idx = df_absen.index[mask].tolist()[0]
                            
                            # 2. Set Header dan bersihkan tabel
                            df_absen.columns = [str(c).upper().strip() for c in df_absen.iloc[header_idx]]
                            df_absen = df_absen.iloc[header_idx + 1:].reset_index(drop=True)

                            for _, row in df_absen.iterrows():
                                hari_val = str(row.get('HARI', '')).upper()
                                
                                # Berhenti jika baris TOTAL atau kosong
                                if "TOTAL" in hari_val or hari_val == 'NAN' or not hari_val:
                                    break
                                
                                potongan = 0.0
                                detail = []

                                # --- LOGIKA 1: TIDAK MASUK (3%) ---
                                # Cek kolom terakhir (KETERANGAN)
                                ket = str(row.iloc[-1]).strip().upper()
                                if ket in ['M', '*', 'ALPHA']:
                                    if hari_val not in ['SABTU', 'MINGGU']:
                                        potongan = 3.0
                                        detail.append("Tidak Masuk (3%)")
                                
                                else:
                                    # --- LOGIKA 2: ABSEN BOLONG (1.5%) ---
                                    # Pastikan kolom MASUK dan PULANG ada
                                    masuk = str(row.get('MASUK', '-')).strip()
                                    pulang = str(row.get('PULANG', '-')).strip()
                                    
                                    if masuk in ['-', 'nan', '']:
                                        potongan += 1.5
                                        detail.append("Tdk Absen Masuk (1.5%)")
                                    if pulang in ['-', 'nan', '']:
                                        potongan += 1.5
                                        detail.append("Tdk Absen Pulang (1.5%)")

                                    # --- LOGIKA 3: TERLAMBAT ---
                                    # Cari index kolom JAM dan MENIT di bawah 'TELAT MASUK'
                                    # Biasanya JAM telat ada di index 5 dan MENIT di index 6
                                    try:
                                        def to_num(val):
                                            val = str(val).replace('-', '0').strip()
                                            return float(val) if val and val != 'nan' else 0

                                        j_telat = to_num(row.iloc[5])
                                        m_telat = to_num(row.iloc[6])
                                        tot_menit = (j_telat * 60) + m_telat

                                        if tot_menit > 0:
                                            if 1 <= tot_menit <= 15: p = 0.25
                                            elif 16 <= tot_menit <= 60: p = 0.5
                                            elif 61 <= tot_menit <= 120: p = 1.0
                                            else: p = 1.5
                                            potongan += p
                                            detail.append(f"Telat {int(tot_menit)}m ({p}%)")
                                    except:
                                        pass

                                if potongan > 0:
                                    all_results.append({
                                        "Nama Pegawai": nama,
                                        "Tanggal": row.get('TANGGAL'),
                                        "Hari": hari_val,
                                        "Potongan (%)": potongan,
                                        "Keterangan": " + ".join(detail)
                                    })
                        else:
                            st.warning(f"Tabel data tidak ditemukan untuk {nama}")
                    
                    progress_bar.progress((idx + 1) / len(df_pegawai))
                except Exception as e:
                    st.error(f"Gagal memproses {nama}: {e}")

            status_text.success("Analisis Selesai!")
            
            if all_results:
                df_final = pd.DataFrame(all_results)
                st.write("### Rekap Potongan Pegawai")
                st.dataframe(df_final, use_container_width=True)

                # Export Excel
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                st.download_button("Download Hasil Excel", buf.getvalue(), "rekap_potongan.xlsx")
            else:
                st.info("Hasil analisis: Tidak ditemukan potongan pada periode ini.")