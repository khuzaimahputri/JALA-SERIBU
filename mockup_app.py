import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- CONFIG HALAMAN ---
st.set_page_config(
    page_title="JALA-SERIBU", 
    page_icon="🌊", 
    layout="wide"
)

# --- CUSTOM CSS ---
st.markdown(
    """
    <style>
        /* Mengurangi padding atas halaman */
        .block-container {
            padding-top: 1rem !important;
            padding-bottom: 1rem !important;
        }

        /* Menghilangkan jarak bawah pada judul utama (h1) */
        div[data-testid="stHeadingWithActionElements"] h1,
        .stMarkdown h1 {
            color: #002B6A !important;
            margin-bottom: 0px !important;
            padding-bottom: 5px !important;
        }

        /* Mengatur jarak atas caption agar mepet ke judul */
        div[data-testid="stCaptionContainer"] {
            margin-top: -5px !important;
            margin-bottom: -20px !important; 
            padding-top: 0px !important;
        }

        /* Mengubah wadah st.metric jadi card */
        div[data-testid="stMetric"] {
            background: linear-gradient(135deg, #F8FAFC 0%, #DBEAFE 100%);
            border: 1px solid #CBD5E1;
            border-top: 4px solid #002B6A; 
            margin-bottom: 0px !important;
            padding: 16px 20px;
            border-radius: 12px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        
        /* Efek Hover */
        div[data-testid="stMetric"]:hover {
            transform: translateY(-3px);
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
        }

        /* Judul metrik */
        div[data-testid="stMetricLabel"] > div {
            font-size: 14px !important;
            color: #475569 !important;
            font-weight: 600 !important;
        }

        /* Warna angka metrik */
        div[data-testid="stMetricValue"] > div {
            font-size: 28px !important;
            font-weight: 800 !important;
            color: #0F172A !important;
        }

        /* Aturan dasar badge delta */
        div[data-testid="stMetricDelta"] {
            padding: 2px 8px !important;
            border-radius: 20px !important;
            width: fit-content !important;
            font-weight: 700 !important;
        }

        /* Badge hijau (panah naik) */
        div[data-testid="stMetricDelta"]:has([data-testid="stMetricDeltaIcon-Up"]) {
            background-color: #DCFCE7 !important;
            color: #15803D !important;
        }
        div[data-testid="stMetricDelta"] [data-testid="stMetricDeltaIcon-Up"] {
            color: #16A34A !important;
        }

        /* Badge merah (panah turun) */
        div[data-testid="stMetricDelta"]:has([data-testid="stMetricDeltaIcon-Down"]) {
            background-color: #FEE2E2 !important;
            color: #B91C1C !important;
        }
        div[data-testid="stMetricDelta"] [data-testid="stMetricDeltaIcon-Down"] {
            color: #DC2626 !important;
        }

        /* Naikin garis pembatas */
        hr {
            margin-top: 5px !important;    /* Atur jarak atas garis (makin kecil makin naik) */
            margin-bottom: 5px !important; /* Jarak dari garis ke judul chart bawah */
        }

        div[data-testid="stVerticalBlockBorderWrapper"] > div {
            background: linear-gradient(135deg, #F8FAFC 0%, #EFF6FF 100%) !important;
            border: 1px solid #CBD5E1 !important;
            border-top: 4px solid #002B6A !important;
            border-radius: 12px !important;
            padding: 16px !important;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05) !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# --- HEADER/TITLE ---
st.title("🌊 JALA-SERIBU")
st.caption("Jaringan Agregasi Layanan dan Akuntabilitas BPS Kabupaten Kepulauan Seribu")


# --- TAB UTAMA ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Overview Kunjungan PST", 
    "📩 Saran & Pengaduan", 
    "❓ Pertanyaan & FAQ", 
    "📈 Progres SKD",
    "🌐 Pemutakhiran Kanal Digital"
])

# ==========================================
# TAB 1: OVERVIEW KUNJUNGAN PST
# ==========================================
with tab1:
    st.subheader("Ringkasan Kunjungan PST Agustus 2026")
    
    # Metric Cards
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Tamu", "42 Orang", delta="+12% dari bulan lalu")
    col2.metric("Keperluan Terbanyak", "Data", "45% dari total")
    col3.metric("Rata-Rata Waktu Kunjungan", "15 Menit", delta="-2 Menit")
    
    # Data Dummy Bar Chart Kategori Kunjungan
    data_kunjungan = pd.DataFrame({
        "Keperluan Kunjungan": ["Data", "Kepentingan Dinas", "Menemui Orang", "Lainnya"],
        "Jumlah Kunjungan": [19, 12, 6, 2]
    })

    col_grafik, col_info = st.columns([65, 35])

    PALETTE_BIRU = ["#C6E7FF", "#64B5F6", "#1E88E5", "#0D47A1"]

    with col_grafik:
        with st.container(border=True):
            fig = px.bar(
                data_kunjungan, 
                x="Jumlah Kunjungan", 
                y="Keperluan Kunjungan", 
                orientation='h',
                text = "Jumlah Kunjungan",
                title="Distribusi Kategori Keperluan Kunjungan Tamu PST",
                color="Jumlah Kunjungan",
                color_continuous_scale=PALETTE_BIRU
            )
            
            fig.update_traces(
                textposition="outside",
                textfont=dict(size=13),
                cliponaxis=False, # biar teks ga kepotong meskipun lewat dari batas sumbu x
            )

            fig.update_layout(
                height=300, 
                margin=dict(t=30, b=10, l=10, r=30),
                coloraxis_showscale=False,  # colorbar
                xaxis_title="Jumlah Kunjungan",  
                yaxis_title="",  
                plot_bgcolor="rgba(0,0,0,0)",  # background transparan
            )

            st.plotly_chart(fig, use_container_width=True)

    with col_info:
        with st.container(border=True):
            data_jenis = pd.DataFrame(
                {"Jenis Tamu": ["Mitra", "Tamu Biasa", "Instansi/Dinas"], "Jumlah": [28, 14, 5]}
            )

            fig_donut = px.pie(
                data_jenis,
                names="Jenis Tamu",
                values="Jumlah",
                hole=0.5,  # untuk donut chart
                title="Proporsi Jenis Tamu",
                color_discrete_sequence=PALETTE_BIRU,
            )

            fig_donut.update_traces(
                textinfo="label+percent",      # menampilkan teks kategori & persen
                textposition="outside",       
                textfont=dict(size=11),
                selector=dict(type='pie'),
                domain=dict(x=[0.1, 0.9], y=[0.1, 0.9]),
            )

            fig_donut.update_layout(
                height=300,
                margin=dict(t=30, b=10, l=50, r=25), 
                showlegend=False                      
            )

            st.plotly_chart(fig_donut, use_container_width=True)

# ==========================================
# TAB 2: SARAN & PENGADUAN 
# ==========================================
with tab2:
    st.subheader("Repositori Saran/Pengaduan")
    
    # Upload Box Screenshot
    with st.expander("📸 Upload Screenshot", expanded=True):
        uploaded_files = st.file_uploader(
            "Upload screenshot saran/pengaduan dari berbagai platform. Dapat upload multiple files.", 
            type=["jpg", "jpeg", "png"], 
            accept_multiple_files=True, 
            key="sp_uploader"
        )
        
        # Cek apakah list uploaded_files ada isinya (bukan pengecekan None)
        if uploaded_files:
            for idx, uploaded_file in enumerate(uploaded_files):
                if idx > 0:
                    st.divider()  # Garis pemisah kalau upload lebih dari 1 foto
                    
                col_img, col_info = st.columns([1, 2])
                with col_img:
                    st.image(uploaded_file, caption=f"Preview Screenshot {idx+1}", use_container_width=True)
                with col_info:
                    st.success(f"✅ Berhasil Membaca Screenshot #{idx+1}!")
                    st.write("**Hasil Ekstraksi Otomatis:**")
                    st.text_input("Platform Identified:", "Instagram (DM)", disabled=True, key=f"platform_{idx}")
                    st.text_area("Konteks Pesan/Pengaduan:", "Mohon info cara mendapatkan data PDRB Kepulauan Seribu 5 tahun terakhir.", disabled=True, key=f"konteks_{idx}")
                    
                    # AI Otomatis Deteksi Apakah Ada Tanggapan di Gambar
                    st.selectbox("Status Tanggapan (Auto-Detect):", ["🟢 Selesai (Ada Tanggapan)", "🔴 Belum Direspon"], index=0, disabled=True, key=f"status_{idx}")
                    st.text_area("Isi Tanggapan/Balasan (Jikalau ada):", "Halo Kak! Data PDRB dapat diakses gratis melalui web kepseribukab.bps.go.id atau silakan ajukan via PST.", disabled=True, key=f"tanggapan_{idx}")
                    
                    if st.button(f"💾 Simpan ke Database Monitoring (#{idx+1})", key=f"btn_save_{idx}"):
                        st.toast(f"Data pengaduan #{idx+1} berhasil tersimpan!")

    st.write("---")
    st.write("### 📑 Tabel Rekapitulasi Pengaduan")
    
    # Dummy Tabel Pengaduan dengan 2 Status & Kolom Tanggapan
    df_pengaduan = pd.DataFrame({
        "Tanggal": ["20/07/2026", "21/07/2026", "22/07/2026"],
        "Platform": ["Instagram", "LAPOR!", "Email"],
        "Topik": [
            "Data PDRB",
            "Apresiasi Petugas",
            "Kendala Portal"
        ],
        "Saran/Pengaduan": [
            "Tanya kelengkapan data PDRB 2025", 
            "Apresiasi pelayanan petugas PST ramah", 
            "Kendala akses login portal Romantik"
        ],
        "Status Tindak Lanjut": ["🔴 Belum Direspon", "🟢 Selesai", "🟢 Selesai"],
        "Tindak Lanjut": [
            "(Belum Ada Balasan)",
            "Terima kasih atas apresiasinya! Kami terus berkomitmen memberikan pelayanan terbaik.",
            "Halo, untuk kendala login akun Romantik telah diselesaikan oleh tim IT BPS."
        ],
        "Bukti":[
            "link",
            "link",
            "link"
        ]
    })
    
    st.dataframe(
        df_pengaduan, 
        use_container_width=True,
        hide_index=True,  
        )

# ==========================================
# TAB 3: PERTANYAAN DAN FAQ
# ==========================================
with tab3:
    st.subheader("Repositori Pertanyaan Publik dan FAQ")
    
    # 1. FORM INPUT SCREENSHOT PERTANYAAN BARU
    with st.expander("📸 Upload Screenshot", expanded=False):
            uploaded_faqs = st.file_uploader(
                "Upload screenshot pertanyaan baru dari berbagai platform. Dapat upload multiple files.", 
                type=["jpg", "jpeg", "png"], 
                accept_multiple_files=True, 
                key="faq_uploader"
            )
            
            # Pengecekan apakah list uploaded_faqs ada isinya
            if uploaded_faqs:
                for idx, uploaded_faq in enumerate(uploaded_faqs):
                    if idx > 0:
                        st.divider() # Garis pembatas jika upload lebih dari 1 file
                    
                    col_faq_img, col_faq_info = st.columns([1, 2])
                    with col_faq_img:
                        st.image(uploaded_faq, caption=f"Preview Pertanyaan #{idx+1}", use_container_width=True)
                    with col_faq_info:
                        st.success(f"✅ Berhasil Membaca Pertanyaan #{idx+1}!")
                        
                        # Widget dengan key unik menggunakan suffix idx
                        st.text_input("Platform:", "Instagram DM", disabled=True, key=f"faq_platform_{idx}")
                        st.text_input("Pertanyaan Di-ekstrak:", "Apakah data inflasi bulanan Kepulauan Seribu ada di website?", disabled=True, key=f"faq_q_{idx}")
                        st.text_area("Rekomendasi Jawaban Standar (FAQ):", "Halo Kak! Data inflasi DKI Jakarta & indikator strategis dapat diakses melalui website resmi kepseribukab.bps.go.id pada menu Publikasi/BRS.", disabled=True, key=f"faq_ans_{idx}")
                        
                        if st.button(f"💾 Simpan ke Database FAQ (#{idx+1})", key=f"faq_save_{idx}"):
                            st.toast(f"Pertanyaan #{idx+1} berhasil ditambahkan ke FAQ!")

    st.write("---")

    # 2. TABEL REKAPITULASI SEMUA PERTANYAAN
    st.write("### 📑 Tabel Rekapitulasi Pertanyaan Masuk")
    df_faq = pd.DataFrame({
        "Tanggal": ["18/07/2026", "20/07/2026", "21/07/2026", "22/07/2026"],
        "Platform": ["Instagram", "Email", "WhatsApp", "Instagram"],
        "Topik": ["Jam Operasional", "Cara Permohonan Data", "Syarat Romantik", "Akses Data PDRB"],
        "Pertanyaan": [
            "Jam operasional layanan PST offline", 
            "Cara permohonan data mikro/raw data", 
            "Syarat pengajuan Rekomendasi Statistik (Romantik)",
            "Akses data PDRB Kepulauan Seribu 2025"
        ],
        "Status Tindak Lanjut": ["🔴 Belum Direspon", "🟢 Selesai", "🟢 Selesai", "🟢 Selesai"],
        "Tindak Lanjut":[
            "a",
            "b",
            "c",
            "d"
        ],
        "Bukti": [
            "link", "link", "link", "link"
        ],
        "Frekuensi": ["28x Ditanyakan", "15x Ditanyakan", "8x Ditanyakan", "12x Ditanyakan"]
    })
    st.dataframe(df_faq, 
                 use_container_width=True,
                 hide_index=True,
                 )

    st.write("---")

    # 3. PERTANYAAN POPULER / TOP FAQ (KARTU RINGKASAN JAWABAN)
    st.write("### ⭐ Pertanyaan Paling Populer (FAQ)")
    
    with st.expander("❓ **Jam berapa pelayanan PST BPS Kepulauan Seribu buka?** (Ditanyakan 28x)", expanded=True):
        st.write("**Jawaban Resmi:**")
        st.info("PST BPS Kabupaten Kepulauan Seribu buka setiap hari kerja:\n- Senin - Kamis: 08.00 - 15.30 WIB\n- Jumat: 08.00 - 16.00 WIB\nHari Sabtu, Minggu, dan Libur Nasional Tutup.")

    with st.expander("❓ **Bagaimana cara mendapatkan data mikro / Raw Data BPS?** (Ditanyakan 15x)"):
        st.write("**Jawaban Resmi:**")
        st.info("Pemohon data mikro dapat mengajukan permohonan secara online melalui portal PST dengan melampirkan identitas KTP dan Surat Pengantar Lembaga/Kampus.")

    with st.expander("❓ **Dimana saya bisa mengunduh publikasi PDRB Kepulauan Seribu?** (Ditanyakan 12x)"):
        st.write("**Jawaban Resmi:**")
        st.info("Publikasi PDRB dapat diunduh gratis dalam format PDF melalui website resmi BPS Kabupaten Kepulauan Seribu (kepseribukab.bps.go.id) pada menu Publikasi.")

# ==========================================
# TAB 4: PROGRES SKD
# ==========================================
with tab4:
    st.subheader("Progres Survei Kebutuhan Data (SKD) Triwulan III 2026")
    
    col_skd1, col_skd2 = st.columns([3, 1])
    
    with col_skd1:
        sub_col_skd1, sub_col_skd2 = st.columns([1, 1])
        
        with sub_col_skd1:
            st.write("##### Jumlah Responden Triwulan Triwulan III 2026")
            st.progress(0.13, text="13,33% dari Target (4/10 Responden)")
        with sub_col_skd2:
            st.write("##### Jumlah Responden Dalam Setahun")
            st.progress(0.8, text="80% dari Target (24/30 Responden)")
        
        # FORM INPUT SCREENSHOT
        with st.expander("📸 Upload Screenshot", expanded=True):
                uploaded_skds = st.file_uploader(
                    "Upload screenshot progres pengisian survei. Dapat upload multiple files.", 
                    type=["jpg", "jpeg", "png"], 
                    accept_multiple_files=True, 
                    key="skd_uploader"
                )
                
                # Pengecekan apakah ada file dalam list uploaded_skds
                if uploaded_skds:
                    for idx, uploaded_skd in enumerate(uploaded_skds):
                        if idx > 0:
                            st.divider()  # Garis pemisah antar screenshot
                        
                        col_skd_img, col_skd_info = st.columns([1, 2])
                        with col_skd_img:
                            st.image(uploaded_skd, caption=f"Preview Bukti Update #{idx+1}", use_container_width=True)
                        with col_skd_info:
                            st.success(f"✅ Berhasil Membaca Bukti Update #{idx+1}!")

        # TABEL PROGRES PENGISIAN SURVEI
        df_skd = pd.DataFrame({
            "Tanggal Cacah": ["18/07/2026", "20/07/2026", "21/07/2026", "22/07/2026"],
            "Nama": ["Busro", "Rice Damayanti", "Sapitri", "Wahyudi"],
            "Status Kuesioner":["Belum terisi lengkap pada Blok 3", 
                                           "Belum terisi lengkap pada Blok 3", 
                                           "Belum diverifikasi", 
                                           "Belum diverifikasi"],
        })
        st.dataframe(df_skd, 
                     use_container_width=True, 
                     hide_index=True,)
                    

    with col_skd2:
        st.write("#### 📂 Akses Laporan Resmi SKD")
        st.link_button("📄 Laporan SKD Triwulan I 2026", "https://kepseribukab.bps.go.id")
        st.link_button("📄 Laporan SKD Triwulan II 2026", "https://kepseribukab.bps.go.id")
        st.button("📄 Laporan SKD Triwulan III 2026(Drafting)", disabled=True)

# ==========================================
# TAB 5: PEMUTAKHIRAN KANAL DIGITAL
# ==========================================
with tab5:
    st.subheader("Pemutakhiran Kanal Digital")
    
    # 1. Metric Cards Ringkasan Update
    col_web1, col_web2, col_web3 = st.columns(3)
    col_web1.metric("Total Update", "18 Konten", "+4 dari bulan lalu")
    col_web2.metric("Kanal Paling Aktif", "Website Utama BPS", "10 Update")
    col_web3.metric("Update Terakhir", "Hari ini (22/07/2026)")

    # 2. FORM INPUT SCREENSHOT UPDATE
    with st.expander("📸 Upload Screenshot", expanded=True):
            uploaded_webs = st.file_uploader(
                "Upload screenshot pemutakhiran berbagai kanal digital resmi. Dapat upload multiple files.", 
                type=["jpg", "jpeg", "png"], 
                accept_multiple_files=True, 
                key="web_uploader"
            )
            
            # Pengecekan apakah ada file yang diunggah dalam list uploaded_webs
            if uploaded_webs:
                for idx, uploaded_web in enumerate(uploaded_webs):
                    if idx > 0:
                        st.divider()  # Garis pemisah antar screenshot
                    
                    col_web_img, col_web_info = st.columns([1, 2])
                    with col_web_img:
                        st.image(uploaded_web, caption=f"Preview Bukti Update #{idx+1}", use_container_width=True)
                    with col_web_info:
                        st.success(f"✅ Berhasil Membaca Bukti Update #{idx+1}!")
                        
                        # Widget dengan key unik menggunakan suffix idx
                        st.text_input("Nama Website / Portal:", "Website Utama BPS (kepseribukab.bps.go.id)", disabled=True, key=f"web_portal_{idx}")
                        st.text_input("Judul / Konten yang Di-update:", "Publikasi Kabupaten Kepulauan Seribu Dalam Angka 2026", disabled=True, key=f"web_title_{idx}")
                        st.text_input("Tanggal Update:", "22/07/2026", disabled=True, key=f"web_date_{idx}")
                        st.text_input("Kategori Konten:", "Publikasi / Berita Resmi Statistik (BRS)", disabled=True, key=f"web_cat_{idx}")
                        
                        if st.button(f"💾 Simpan Log Update (#{idx+1})", key=f"web_save_{idx}"):
                            st.toast(f"Bukti update website #{idx+1} berhasil dicatat!")

    st.write("---")

    # 3. TABEL REKAPITULASI LOG UPDATE WEBSITE
    st.write("### 📑 Tabel Rekapitulasi Log Update Kanal Digital")
    
    df_website = pd.DataFrame({
        "Tanggal Update": ["22/07/2026", "21/07/2026", "19/07/2026", "15/07/2026"],
        "Kanal Digital": [
            "Website BPS", 
            "Portal PPID", 
            "Instagram", 
            "Website BPS"
        ],
        "Topik Konten": [
            "Publikasi Kepulauan Seribu Dalam Angka 2026", 
            "Update Laporan Akses Informasi Publik Q2", 
            "Informasi Magang", 
            "Berita Senam Bersama dan Layanan PST Keliling"
        ],
        "Bukti": ["link", "link", "link", "link"],
        
    })
    
    st.dataframe(df_website, 
                 use_container_width=True,
                 hide_index=True,)