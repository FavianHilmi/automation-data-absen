import streamlit as st
import pandas as pd
import requests
import io
import zipfile
import streamlit.components.v1 as components
from streamlit_option_menu import option_menu
from bs4 import BeautifulSoup
import base64

st.set_page_config(page_title="Automation Data Absen", layout="centered")


BASE_URL = "https://teko-cak.surabaya.go.id/cetak_new/lap_per_pegawai2/"
ID_INSTANSI = "3.10.00.00.00"

with st.sidebar:
    # st.image("https://www.surabaya.go.id/images/logo.png", width=100)
    st.title("Automation Data Absen")
    # st.title("Main Menu")
    
    menu = option_menu(
        menu_title=None,
        options=["Download Template", "Proses Download Data", "Hitung Potongan"],
        icons=["cloud-download", "box-arrow-in-down", "calculator"],
        menu_icon="cast", 
        default_index=0,
        styles={
            "container": {"padding": "5px", "background-color": "#262730"},
            "icon": {"color": "#00d4ff", "font-size": "18px"}, 
            "nav-link": {
                "font-size": "14px", 
                "text-align": "left", 
                "margin": "5px", 
                "--hover-color": "#31333f"
            },
            "nav-link-selected": {"background-color": "#007bff"},
        }
    )

if menu == "Download Template":
    st.subheader("Download Template Excel")
    st.write("Template di bawah ini buat dipakai saat proses Download Data & Hitung Potongan Gaji:")
    data_template = {
        'Nama_Pegawai': ['ACHMAD SHOLIKIN', 'HERU WIJIANTO'],
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
                            file_name = f"Laporan_{nama}_{params['tgl']}-{params['bulan']}-{params['tahun']}.{ext}"
                            
                            if output_mode == "File ZIP":
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
                if success_files:
                    if st.button("Download Semua"):
                        js_script = ""
                        for f in success_files:
                            b64 = base64.b64encode(f['content']).decode()
                            mime_type = "application/pdf" if file_type == "pdf" else "application/vnd.ms-excel"
                            js_script += f"""
                                var link = document.createElement('a');
                                link.href = 'data:{mime_type};base64,{b64}';
                                link.download = '{f['name']}';
                                link.click();
                            """
                        components.html(f"<script>{js_script}</script>", height=0)
                        st.balloons()
                    
                    with st.expander("Daftar File"):
                        for f in success_files:
                            st.download_button(label=f"{f['name']}", data=f['content'], file_name=f['name'], key=f['name'])    

elif menu == "Hitung Potongan":
    st.subheader("Hitung Potongan Gaji")

    uploaded_file = st.file_uploader("Upload Template Pegawai:", type=["xlsx", "xls"])

    if uploaded_file:
        df_pegawai = pd.read_excel(uploaded_file)

        if "hasil_potongan" not in st.session_state:
            st.session_state.hasil_potongan = None

        if st.button("Hitung Potongan"):
            all_results = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, row_peg in df_pegawai.iterrows():
                nama = row_peg['Nama_Pegawai']
                status_text.text(f"Sekedhap, memproses: {nama}...")
                
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
                        soup = BeautifulSoup(response.text, 'html.parser')
                        table = soup.find('table', {'border': '1'})
                        
                        if table:
                            rows = table.find('tbody').find_all('tr')
                            for row in rows:
                                cols = row.find_all('td')
                                if len(cols) < 10: continue

                                hari_val = cols[0].get_text(strip=True).upper()
                                tgl_val = cols[1].get_text(strip=True).upper()

                                if "TOTAL" in hari_val or "TOTAL" in tgl_val or not tgl_val or tgl_val == "-":
                                    continue
                                if not any(char.isdigit() for char in tgl_val):
                                    continue

                                masuk_val = cols[3].get_text(strip=True)
                                jam_telat_val = cols[4].get_text(strip=True)
                                mnt_telat_val = cols[5].get_text(strip=True)
                                pulang_val = cols[6].get_text(strip=True)
                                ket_val = cols[-1].get_text(strip=True).upper()

                                potongan = 0.0
                                detail = []

                                if ket_val == 'M' and hari_val not in ['SABTU', 'MINGGU']:
                                    potongan = 3.0
                                    detail.append("Mangkir (3%)")
                                
                                elif ket_val in ['H', '', 'NAN', '*']:
                                    if ket_val == 'H':
                                        if masuk_val in ['-', '']:
                                            potongan += 1.5
                                            detail.append("Tdk Absen Masuk (1.5%)")
                                        if pulang_val in ['-', '']:
                                            potongan += 1.5
                                            detail.append("Tdk Absen Pulang (1.5%)")

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
                                        "Nama": nama, "Tanggal": tgl_val, "Hari": hari_val,
                                        "Potongan (%)": potongan
                                    })
                    progress_bar.progress((idx + 1) / len(df_pegawai))
                except Exception as e:
                    st.error(f"Error {nama}: {e}")

            st.session_state.hasil_potongan = all_results
            status_text.success("Selesai!")

        if st.session_state.hasil_potongan:
            df_detail = pd.DataFrame(st.session_state.hasil_potongan)
            
            with st.expander("Total Potongan Gaji", expanded=True):
                df_summary_raw = df_detail.groupby("Nama").agg({
                    "Potongan (%)": "sum",
                }).reset_index()

                df_summary_web = df_summary_raw.copy()
                df_summary_web["Ringkasan"] = df_summary_web.apply(
                    lambda row: f"{row['Potongan (%)']:.2f}%", axis=1
                )
                
                st.dataframe(df_summary_web[["Nama", "Ringkasan"]], use_container_width=True)
            
            with st.expander("Rincian Harian", expanded=False):
                st.dataframe(df_detail, use_container_width=True)

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                df_summary_raw.to_excel(writer, sheet_name='Summary', index=False)
                df_detail.to_excel(writer, sheet_name='Detail_Harian', index=False)
            
            st.write("")
            st.download_button(
                label="Download Laporan (Excel)",
                data=buf.getvalue(),
                file_name="rekap_potongan_{''}.xlsx",
                mime="application/vnd.ms-excel",
                use_container_width=True
            )
        elif st.session_state.hasil_potongan == []:
            st.info("Jos, Karyawan e disiplin!")