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
    st.subheader("Hitung Persentase Potongan Kehadiran")
    st.write("Upload file laporan (XLS) untuk dihitung potongannya sesuai aturan.")

    # Gunakan file_uploader untuk menerima banyak file sekaligus
    uploaded_files = st.file_uploader("Upload File XLS Laporan:", type=["xls"], accept_multiple_files=True)

    if uploaded_files:
        all_results = []

        for uploaded_file in uploaded_files:
            try:
                # 1. Ambil Nama Pegawai dari Cell B4 (Row 4, Col 2)
                # Gunakan engine='xlrd' karena biasanya file dari portal tersebut adalah .xls lama
                df_info = pd.read_excel(uploaded_file, header=None, nrows=10)
                nama_pegawai = df_info.iloc[3, 1] 

                # 2. Baca data tabel (Header asli ada di baris 10, jadi skip 9 baris)
                df_absen = pd.read_excel(uploaded_file, skiprows=9)
                
                # Bersihkan kolom agar tidak ada spasi di nama kolom
                df_absen.columns = [str(c).strip() for c in df_absen.columns]

                for index, row in df_absen.iterrows():
                    # Stop jika baris kosong atau mencapai "TOTAL"
                    if pd.isna(row.get('HARI')) or "TOTAL" in str(row.get('HARI')).upper():
                        break
                    
                    potongan = 0.0
                    ket_list = []
                    
                    # LOGIKA 1: Tidak Masuk (3%)
                    # Jika kolom Keterangan (ETERANGA) berisi '*' dan bukan hari libur
                    if str(row.get('ETERANGA')) == '*' and row['HARI'] not in ['MINGGU', 'SABTU']:
                        potongan = 3.0
                        ket_list.append("Tidak Masuk (3%)")
                    
                    else:
                        # LOGIKA 2: Absen Bolong (1.5% masing-masing)
                        if pd.isna(row.get('MASUK')):
                            potongan += 1.5
                            ket_list.append("Tanpa Absen Masuk (1.5%)")
                        if pd.isna(row.get('PULANG')):
                            potongan += 1.5
                            ket_list.append("Tanpa Absen Pulang (1.5%)")

                        # LOGIKA 3: Terlambat (Berdasarkan kolom JAM dan MENIT di bagian TELAT MASUK)
                        # Nama kolom di pandas bisa berubah jika ada duplikasi, pastikan index kolom benar
                        # Berdasarkan gambar: JAM telat ada di kolom ke-6 (index 5) dan MENIT di kolom 7 (index 6)
                        jam_t = row.iloc[5] if pd.notna(row.iloc[5]) and isinstance(row.iloc[5], (int, float)) else 0
                        menit_t = row.iloc[6] if pd.notna(row.iloc[6]) and isinstance(row.iloc[6], (int, float)) else 0
                        
                        total_menit = (jam_t * 60) + menit_t

                        if total_menit > 0:
                            if 1 <= total_menit <= 15:
                                p_val = 0.25
                            elif 16 <= total_menit <= 60:
                                p_val = 0.5
                            elif 61 <= total_menit <= 120:
                                p_val = 1.0
                            else:
                                p_val = 1.5
                            
                            potongan += p_val
                            ket_list.append(f"Terlambat {int(total_menit)} mnt ({p_val}%)")

                    # Simpan hanya jika ada potongan
                    if potongan > 0:
                        all_results.append({
                            "Nama Pegawai": nama_pegawai,
                            "Tanggal": row.get('TANGGAL'),
                            "Hari": row.get('HARI'),
                            "Potongan (%)": potongan,
                            "Keterangan": " + ".join(ket_list)
                        })
            except Exception as e:
                st.error(f"Gagal memproses file {uploaded_file.name}: {e}")

        if all_results:
            df_hasil = pd.DataFrame(all_results)
            st.success(f"Berhasil menghitung {len(df_hasil)} baris potongan.")
            st.dataframe(df_hasil, use_container_width=True)

            # Download Buttons
            csv = df_hasil.to_csv(index=False).encode('utf-8')
            st.download_button("Download Rekap (CSV)", csv, "rekap_potongan.csv", "text/csv")
            
            # Excel Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_hasil.to_excel(writer, index=False, sheet_name='Potongan')
            st.download_button("Download Rekap (Excel)", output.getvalue(), "rekap_potongan.xlsx")
        else:
            st.info("Semua pegawai disiplin! Tidak ditemukan potongan.")