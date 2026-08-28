import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from services.google_sheets import get_google_sheet_data
from datetime import datetime
from zoneinfo import ZoneInfo
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

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

        /* Caption khusus baris update + periode + download */
        div[data-testid="stHorizontalBlock"]:has(
            div[data-testid="stSelectbox"]
        ) div[data-testid="stCaptionContainer"] {
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
            min-height: 130px;
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

        div[data-testid="stDownloadButton"] button {
            width: 105px !important;
            min-width: 105px !important;
            max-width: 105px !important;

            height: 30px !important;
            min-height: 30px !important;

            padding: 3px 7px !important;
            white-space: nowrap !important;
        }

        /* Ukuran tulisan */
        div[data-testid="stDownloadButton"] button p {
            font-size: 13px !important;
        }

        div[data-testid="stDownloadButton"] button:hover {
            background: #F0F7FF !important;
            border-color: #64B5F6 !important;
            color: #0D47A1 !important;
        }

        div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
            min-height: 30px !important;
            height: 30px !important;
            font-size: 13px !important;
            padding-top: 0px !important;
            padding-bottom: 0px !important;
        }

        /* Wrapper isi selectbox: teks + ikon */
        div[data-testid="stSelectbox"]
        div[data-baseweb="select"] > div > div {
            min-height: 30px !important;
            height: 30px !important;

            display: flex !important;
            align-items: center !important;

            padding-top: 1px !important;
            padding-bottom: 0px !important;
            box-sizing: border-box !important;
        }

        /* Teks pilihan */
        div[data-testid="stSelectbox"]
        div[data-baseweb="select"] div[value] {
            line-height: 1 !important;
            padding-top: 8.5px !important;
            margin: 0 !important;
            transform: none !important;
        }

        /* =========================================
        UPDATE + PERIODE + DOWNLOAD AUTO RESPONSIVE
        ========================================= */

        div[data-testid="stHorizontalBlock"]:has(
            div[data-testid="stSelectbox"]
        ) {
            display: flex !important;
            flex-wrap: wrap !important;
            align-items: center !important;

            column-gap: 12px !important;
            row-gap: 6px !important;

            margin-top: -11px !important;
            margin-bottom: 1px !important;
        }

        /* Update fleksibel */
        div[data-testid="stHorizontalBlock"]:has(
            div[data-testid="stSelectbox"]
        ) > div:nth-child(1) {
            flex: 1 1 420px !important;
            width: auto !important;
            min-width: 280px !important;
        }

        /* Periode fleksibel */
        div[data-testid="stHorizontalBlock"]:has(
            div[data-testid="stSelectbox"]
        ) > div:nth-child(2) {
            flex: 0 1 230px !important;
            width: auto !important;
            min-width: 180px !important;
            max-width: 230px !important;
        }

        /* Download tetap */
        div[data-testid="stHorizontalBlock"]:has(
            div[data-testid="stSelectbox"]
        ) > div:nth-child(3) {
            flex: 0 0 105px !important;
            width: 105px !important;
            min-width: 105px !important;
            max-width: 105px !important;
            margin: 0 !important;
        }

        /* Selectbox ikut ruang yang tersedia */
        div[data-testid="stSelectbox"] {
            width: 100% !important;
            max-width: 230px !important;
        }

        /* Tombol download */
        div[data-testid="stDownloadButton"] {
            display: flex !important;
            justify-content: flex-start !important;
        }

        /* =========================================
        RESPONSIVE - MOBILE
        ========================================= */
        @media (max-width: 700px) {

            /*
            Periode + download mengikuti
            layout auto-responsive di atas.
            */

            /* Metric cards vertikal */
            div[data-testid="stHorizontalBlock"]:has(
                div[data-testid="stMetric"]
            ) {
                flex-direction: column !important;
                gap: 10px !important;
            }

            div[data-testid="stHorizontalBlock"]:has(
                div[data-testid="stMetric"]
            ) > div {
                width: 100% !important;
                min-width: 100% !important;
                flex: 0 0 auto !important;
            }

            /* Chart vertikal */
            div[data-testid="stHorizontalBlock"]:has(
                div[data-testid="stPlotlyChart"]
            ) {
                flex-direction: column !important;
                gap: 10px !important;
            }

            div[data-testid="stHorizontalBlock"]:has(
                div[data-testid="stPlotlyChart"]
            ) > div {
                width: 100% !important;
                min-width: 100% !important;
                flex: 0 0 auto !important;
            }
        }


        /* =========================================
        MOBILE SANGAT KECIL
        ========================================= */
        @media (max-width: 380px) {

            /* Kalau benar-benar sempit, baru kontrol ditumpuk */
            div[data-testid="stHorizontalBlock"]:has(
                div[data-testid="stSelectbox"]
            ) > div:nth-child(2),
            div[data-testid="stHorizontalBlock"]:has(
                div[data-testid="stSelectbox"]
            ) > div:nth-child(3) {
                flex: 0 0 100% !important;
                width: 100% !important;
                max-width: 100% !important;
                min-width: 0 !important;
            }

            div[data-testid="stSelectbox"] {
                width: 100% !important;
                max-width: 100% !important;
            }
        }        

    </style>
""",
    unsafe_allow_html=True,
)

# --- HEADER/TITLE ---
st.title("🌊 JALA-SERIBU")
st.caption("Jaringan Agregasi Layanan dan Akuntabilitas BPS Kabupaten Kepulauan Seribu")

# --- GLOBAL CONSTANTA ---
kategori_jenis_tamu = [
    "Masyarakat Umum/Mahasiswa",
    "Mitra Statistik",
    "Instansi/Dinas",
    "Lainnya"
]

kategori_keperluan = [
    "Konsultasi Data",
    "Tugas Mitra",
    "Kepentingan Dinas",
    "Bertemu Orang",
    "Lainnya"
]

# --- HELPER FUNCTIONS ---
def buat_excel_overview(
    df_ringkasan,
    data_kunjungan,
    data_jenis,
    df_mentah,
    rekap_bulanan=None, 
    nama_periode=None
):
    output = BytesIO()

    if nama_periode:
        bagian_periode = nama_periode.split("_")

        if bagian_periode[0] == "Triwulan":
            nomor_triwulan = int(bagian_periode[1])
            tahun_periode = bagian_periode[2]

            romawi = {
                1: "I",
                2: "II",
                3: "III",
                4: "IV"
            }

            judul_periode = (
                f"TRIWULAN {romawi[nomor_triwulan]} "
                f"TAHUN {tahun_periode}"
            )

        else:
            judul_periode = nama_periode.replace("_", " ").upper()

    else:
        judul_periode = ""

            # =========================
            # WARNA & STYLE DASAR
            # =========================
    biru = "1F4E78"
    putih = "FFFFFF"

    garis_tipis = Side(
        style="thin",
        color="000000"
    )

    border_tabel = Border(
        left=garis_tipis,
        right=garis_tipis,
        top=garis_tipis,
        bottom=garis_tipis
    )

    fill_header = PatternFill(
        fill_type="solid",
        fgColor=biru
    )

    font_header = Font(
        bold=True,
        color=putih
    )

    align_center = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df_ringkasan.to_excel(
            writer,
            index=False,
            sheet_name="Ringkasan"
        )

        ws_ringkasan = writer.sheets["Ringkasan"]

        # Sisipkan 2 baris di atas untuk judul
        ws_ringkasan.insert_rows(1, amount=2)

        # Judul periode
        ws_ringkasan.merge_cells("A1:B1")
        ws_ringkasan["A1"] = judul_periode
        ws_ringkasan["A1"].font = Font(
            bold=True,
            size=14
        )
        ws_ringkasan["A1"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        # Header sekarang pindah ke baris 3
        for cell in ws_ringkasan[3]:
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = align_center
            cell.border = border_tabel

        # Isi tabel
        for row in ws_ringkasan.iter_rows(
            min_row=4,
            max_row=ws_ringkasan.max_row,
            min_col=1,
            max_col=2
        ):
            for cell in row:
                cell.border = border_tabel
                cell.alignment = Alignment(
                    vertical="center"
                )

        # Kolom Nilai dibuat center
        for row in range(4, ws_ringkasan.max_row + 1):
            ws_ringkasan.cell(
                row=row,
                column=2
            ).alignment = align_center

        # Lebar kolom
        ws_ringkasan.column_dimensions["A"].width = 32
        ws_ringkasan.column_dimensions["B"].width = 28

        data_kunjungan.to_excel(
            writer,
            index=False,
            sheet_name="Keperluan Kunjungan"
        )

        ws_keperluan = writer.sheets["Keperluan Kunjungan"]

        # Sisipkan 2 baris untuk judul
        ws_keperluan.insert_rows(1, amount=2)

        # Judul periode
        ws_keperluan.merge_cells("A1:B1")
        ws_keperluan["A1"] = judul_periode
        ws_keperluan["A1"].font = Font(
            bold=True,
            size=14
        )
        ws_keperluan["A1"].alignment = align_center

        # Header
        for cell in ws_keperluan[3]:
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = align_center
            cell.border = border_tabel

        # Isi tabel
        for row in ws_keperluan.iter_rows(
            min_row=4,
            max_row=ws_keperluan.max_row,
            min_col=1,
            max_col=2
        ):
            for cell in row:
                cell.border = border_tabel
                cell.alignment = Alignment(
                    vertical="center"
                )

        # Kolom jumlah dibuat center
        for row in range(4, ws_keperluan.max_row + 1):
            ws_keperluan.cell(
                row=row,
                column=2
            ).alignment = align_center

        # Lebar kolom
        ws_keperluan.column_dimensions["A"].width = 28
        ws_keperluan.column_dimensions["B"].width = 20

        data_jenis.to_excel(
            writer,
            index=False,
            sheet_name="Jenis Tamu"
        )

        ws_jenis = writer.sheets["Jenis Tamu"]

        # Sisipkan 2 baris untuk judul
        ws_jenis.insert_rows(1, amount=2)

        # Judul periode
        ws_jenis.merge_cells("A1:B1")
        ws_jenis["A1"] = judul_periode
        ws_jenis["A1"].font = Font(
            bold=True,
            size=14
        )
        ws_jenis["A1"].alignment = align_center

        # Header
        for cell in ws_jenis[3]:
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = align_center
            cell.border = border_tabel

        # Isi tabel
        for row in ws_jenis.iter_rows(
            min_row=4,
            max_row=ws_jenis.max_row,
            min_col=1,
            max_col=2
        ):
            for cell in row:
                cell.border = border_tabel
                cell.alignment = Alignment(
                    vertical="center"
                )

        # Kolom jumlah dibuat center
        for row in range(4, ws_jenis.max_row + 1):
            ws_jenis.cell(
                row=row,
                column=2
            ).alignment = align_center

        # Lebar kolom
        ws_jenis.column_dimensions["A"].width = 30
        ws_jenis.column_dimensions["B"].width = 14

        df_mentah.to_excel(
            writer,
            index=False,
            sheet_name="Data Mentah"
        )

        ws_mentah = writer.sheets["Data Mentah"]

        # Freeze header
        ws_mentah.freeze_panes = "A2"

        # Autofilter
        ws_mentah.auto_filter.ref = ws_mentah.dimensions

        # Style header
        for cell in ws_mentah[1]:
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = align_center
            cell.border = border_tabel

        # Style isi tabel
        for row in ws_mentah.iter_rows(
            min_row=2,
            max_row=ws_mentah.max_row,
            min_col=1,
            max_col=ws_mentah.max_column
        ):
            for cell in row:
                cell.border = border_tabel
                cell.alignment = Alignment(
                    vertical="center"
                )

        # Mapping nama header ke nomor kolom
        header_map = {
            cell.value: cell.column
            for cell in ws_mentah[1]
        }


        # Format tanggal
        if "Tanggal Kehadiran" in header_map:
            kolom_tanggal = header_map["Tanggal Kehadiran"]

            for row in range(2, ws_mentah.max_row + 1):
                ws_mentah.cell(
                    row=row,
                    column=kolom_tanggal
                ).number_format = "dd/mm/yyyy"


        # Format jam
        for nama_kolom in ["Jam Datang", "Jam Pulang"]:
            if nama_kolom in header_map:
                nomor_kolom = header_map[nama_kolom]

                for row in range(2, ws_mentah.max_row + 1):
                    ws_mentah.cell(
                        row=row,
                        column=nomor_kolom
                    ).number_format = "hh:mm"

        # Lebar kolom
        lebar_kolom = {
            "Timestamp": 22,
            "Nama": 24,
            "Jenis Tamu": 30,
            "Keterangan Tamu": 30,
            "Keperluan": 24,
            "Detail Keperluan": 36,
            "Tanggal Kehadiran": 18,
            "Jam Datang": 14,
            "Jam Pulang": 14,
        }


        for nama_kolom, lebar in lebar_kolom.items():
            if nama_kolom in header_map:
                huruf_kolom = ws_mentah.cell(
                    row=1,
                    column=header_map[nama_kolom]
                ).column_letter

                ws_mentah.column_dimensions[
                    huruf_kolom
                ].width = lebar

        if rekap_bulanan is not None:

            rekap_bulanan.to_excel(
                writer,
                index=False,
                header=False,
                sheet_name="Rekap Bulanan",
                startrow=4
            )

            ws_rekap = writer.sheets["Rekap Bulanan"]

            # =========================
            # JUDUL PERIODE
            # =========================
            ws_rekap.merge_cells("A1:G1")
            ws_rekap["A1"] = judul_periode

            ws_rekap["A1"].font = Font(
                bold=True,
                size=14
            )

            ws_rekap["A1"].alignment = align_center

            # =========================
            # HEADER BERTINGKAT
            # =========================
            ws_rekap.merge_cells("A3:A4")
            ws_rekap.merge_cells("B3:F3")
            ws_rekap.merge_cells("G3:G4")

            ws_rekap["A3"] = "Bulan"
            ws_rekap["B3"] = "Keperluan"
            ws_rekap["G3"] = "Jumlah"

            ws_rekap["B4"] = "Konsultasi Data"
            ws_rekap["C4"] = "Tugas Mitra"
            ws_rekap["D4"] = "Kepentingan Dinas"
            ws_rekap["E4"] = "Bertemu Orang"
            ws_rekap["F4"] = "Lainnya"

            # Style header
            for row in ws_rekap.iter_rows(
                min_row=3,
                max_row=4,
                min_col=1,
                max_col=7
            ):
                for cell in row:
                    cell.fill = fill_header
                    cell.font = font_header
                    cell.alignment = align_center
                    cell.border = border_tabel

            # =========================
            # ISI TABEL
            # =========================
            for row in ws_rekap.iter_rows(
                min_row=5,
                max_row=ws_rekap.max_row,
                min_col=1,
                max_col=7
            ):
                for cell in row:
                    cell.border = border_tabel
                    cell.alignment = align_center

            # Bulan rata kiri
            for row in range(5, ws_rekap.max_row + 1):
                ws_rekap.cell(
                    row=row,
                    column=1
                ).alignment = Alignment(
                    horizontal="left",
                    vertical="center"
                )

            # =========================
            # LEBAR KOLOM
            # =========================
            ws_rekap.column_dimensions["A"].width = 14
            ws_rekap.column_dimensions["B"].width = 20
            ws_rekap.column_dimensions["C"].width = 16
            ws_rekap.column_dimensions["D"].width = 20
            ws_rekap.column_dimensions["E"].width = 18
            ws_rekap.column_dimensions["F"].width = 14
            ws_rekap.column_dimensions["G"].width = 12

    return output.getvalue()

def siapkan_data_excel(df_periode):
    # Total tamu
    total_tamu = len(df_periode)

    # Keperluan
    data_kunjungan = (
        df_periode["Keperluan"]
        .value_counts()
        .reindex(
            kategori_keperluan,
            fill_value=0
        )
        .rename_axis("Keperluan Kunjungan")
        .reset_index(name="Jumlah Kunjungan")
    )

    # Jenis tamu
    data_jenis = (
        df_periode["Jenis Tamu"]
        .value_counts()
        .reindex(
            kategori_jenis_tamu,
            fill_value=0
        )
        .rename_axis("Jenis Tamu")
        .reset_index(name="Jumlah")
    )

    # Keperluan terbanyak
    if not data_kunjungan.empty:
        keperluan_terbanyak = (
            data_kunjungan
            .sort_values("Jumlah Kunjungan", ascending=False)
            .iloc[0]["Keperluan Kunjungan"]
        )
    else:
        keperluan_terbanyak = "-"

    # Durasi
    tanggal_string = df_periode[
        "Tanggal Kehadiran"
    ].dt.strftime("%Y-%m-%d")

    waktu_datang = pd.to_datetime(
        tanggal_string + " " + df_periode["Jam Datang"].astype(str),
        errors="coerce"
    )

    waktu_pulang = pd.to_datetime(
        tanggal_string + " " + df_periode["Jam Pulang"].astype(str),
        errors="coerce"
    )

    durasi = (
        waktu_pulang - waktu_datang
    ).dt.total_seconds() / 60

    durasi = durasi[durasi >= 0]

    if not durasi.empty:
        nilai_durasi = f"{durasi.mean():.0f} Menit"
    else:
        nilai_durasi = "-"

    # Ringkasan
    df_ringkasan = pd.DataFrame({
        "Indikator": [
            "Total Tamu",
            "Keperluan Terbanyak",
            "Rata-Rata Waktu Kunjungan"
        ],
        "Nilai": [
            f"{total_tamu} Orang",
            keperluan_terbanyak,
            nilai_durasi
        ]
    })

    return (
        df_ringkasan,
        data_kunjungan,
        data_jenis
    )

def get_triwulan(bulan):
    if bulan <= 3:
        return 1
    elif bulan <= 6:
        return 2
    elif bulan <= 9:
        return 3
    else:
        return 4

def buat_rekap_bulanan(
    df_periode,
    nama_bulan,
    tambah_total=False,
    bulan_wajib=None
):

    df_temp = df_periode.copy()

    df_temp["Bulan"] = (
        df_temp["Tanggal Kehadiran"]
        .dt.month
        .map(lambda x: nama_bulan[x])
    )

    rekap = pd.crosstab(
        df_temp["Bulan"],
        df_temp["Keperluan"]
    )

    # Pastikan bulan yang tidak punya data tetap muncul
    if bulan_wajib is not None:
        rekap = rekap.reindex(
            bulan_wajib,
            fill_value=0
        )

    # Pastikan semua kategori selalu ada
    for kategori in kategori_keperluan:
        if kategori not in rekap.columns:
            rekap[kategori] = 0

    rekap = rekap[kategori_keperluan]

    rekap["Jumlah"] = rekap.sum(axis=1)

    rekap = rekap.reset_index()

    urutan_bulan = {
        nama_bulan[i]: i
        for i in range(1, 13)
    }

    rekap["_urutan"] = rekap["Bulan"].map(urutan_bulan)

    rekap = (
        rekap
        .sort_values("_urutan")
        .drop(columns="_urutan")
        .reset_index(drop=True)
    )

    # Tambahkan baris total
    if tambah_total:
        baris_total = {
            "Bulan": "Jumlah"
        }

        for kolom in kategori_keperluan:
            baris_total[kolom] = rekap[kolom].sum()

        baris_total["Jumlah"] = rekap["Jumlah"].sum()

        rekap = pd.concat(
            [
                rekap,
                pd.DataFrame([baris_total])
            ],
            ignore_index=True
        )

    return rekap

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
    from services.google_sheets import get_google_sheet_data

    @st.fragment(run_every="1h") # Auto-update setiap 1 jam
    def overview_kunjungan_pst():

        # Ambil Google Sheets
        df = get_google_sheet_data()
        
        if df.empty:
            st.warning("Belum ada data kunjungan.")
            return

        # Ambil waktu update
        waktu_update = datetime.now(
            ZoneInfo("Asia/Jakarta")
        ).strftime("%d/%m/%Y %H:%M WIB")

        # --- PREPROCESSING DATA ---
        df["Tanggal Kehadiran"] = pd.to_datetime(
            df["Tanggal Kehadiran"],
            errors="coerce"
        )

        # Buang baris yang tanggalnya tidak valid / NA
        df = df.dropna(subset=["Tanggal Kehadiran"])

        if df.empty:
            st.warning("Data tanggal kunjungan belum valid.")
            return

        nama_bulan = [
            "",
            "Januari", "Februari", "Maret", "April",
            "Mei", "Juni", "Juli", "Agustus",
            "September", "Oktober", "November", "Desember"
        ]

        periode_bulan_tersedia = (
            df["Tanggal Kehadiran"]
            .dt.to_period("M")
            .drop_duplicates()
            .sort_values(ascending=False)
            .tolist()
        )

        opsi_periode = []

        for periode in periode_bulan_tersedia:
            opsi_periode.append(
                f"{nama_bulan[periode.month]} {periode.year}"
            )

        # Opsi Triwulan
        triwulan_tersedia = set()

        for periode in periode_bulan_tersedia:
            nomor_triwulan = get_triwulan(periode.month)

            triwulan_tersedia.add(
                (periode.year, nomor_triwulan)
            )

        romawi = {
            1: "I",
            2: "II",
            3: "III",
            4: "IV"
        }

        romawi_ke_angka = {
            "I": 1,
            "II": 2,
            "III": 3,
            "IV": 4
        }

        for tahun, triwulan in sorted(
            triwulan_tersedia,
            reverse=True
        ):
            opsi_periode.append(
                f"Triwulan {romawi[triwulan]} {tahun}"
            )

        # =========================
        # PILIH PERIODE DASHBOARD
        # =========================

        # Placeholder judul supaya tetap tampil di atas baris kontrol
        judul_dashboard = st.empty()

        col_update, col_periode, col_download = st.columns(
            [7, 1.8, 0.65],
            gap="small",
            vertical_alignment="center"
        )

        with col_update:
            st.caption(
                f"🔄 *Data diperbarui otomatis setiap 1 jam* | "
                f"*Terakhir diperbarui: {waktu_update}*"
            )

        with col_periode:
            periode_pilih = st.selectbox(
                "Periode",
                opsi_periode,
                index=0,
                label_visibility="collapsed"
            )


        # =========================
        # FILTER PERIODE TERPILIH
        # =========================

        is_triwulan = periode_pilih.startswith("Triwulan")

        if is_triwulan:

            _, nomor_triwulan, tahun_periode = periode_pilih.split()

            nomor_triwulan = romawi_ke_angka[nomor_triwulan]
            tahun_periode = int(tahun_periode)

            bulan_awal = (nomor_triwulan - 1) * 3 + 1
            bulan_akhir = bulan_awal + 2

            bulan_triwulan = [
                nama_bulan[i]
                for i in range(bulan_awal, bulan_akhir + 1)
            ]

            # Data periode yang sedang ditampilkan
            df_periode = df[
                (df["Tanggal Kehadiran"].dt.year == tahun_periode) &
                (df["Tanggal Kehadiran"].dt.month >= bulan_awal) &
                (df["Tanggal Kehadiran"].dt.month <= bulan_akhir)
            ].copy()

            # Triwulan sebelumnya untuk delta
            if nomor_triwulan > 1:
                triwulan_lalu = nomor_triwulan - 1
                tahun_lalu = tahun_periode
            else:
                triwulan_lalu = 4
                tahun_lalu = tahun_periode - 1

            bulan_awal_lalu = (triwulan_lalu - 1) * 3 + 1
            bulan_akhir_lalu = bulan_awal_lalu + 2

            df_periode_lalu = df[
                (df["Tanggal Kehadiran"].dt.year == tahun_lalu) &
                (df["Tanggal Kehadiran"].dt.month >= bulan_awal_lalu) &
                (df["Tanggal Kehadiran"].dt.month <= bulan_akhir_lalu)
            ].copy()

            label_periode = (
                f"Triwulan {romawi[nomor_triwulan]} "
                f"{tahun_periode}"
            )

            nama_file_periode = (
                f"Triwulan_{nomor_triwulan}_{tahun_periode}"
            )

            label_perbandingan = "dari triwulan lalu"

        else:

            bulan_triwulan = None
            periode_terpilih = None

            for periode in periode_bulan_tersedia:
                label = (
                    f"{nama_bulan[periode.month]} {periode.year}"
                )

                if label == periode_pilih:
                    periode_terpilih = periode
                    break

            # Data bulan yang sedang ditampilkan
            df_periode = df[
                (df["Tanggal Kehadiran"].dt.year == periode_terpilih.year) &
                (df["Tanggal Kehadiran"].dt.month == periode_terpilih.month)
            ].copy()

            # Bulan sebelumnya untuk delta
            periode_lalu = periode_terpilih - 1

            df_periode_lalu = df[
                (df["Tanggal Kehadiran"].dt.year == periode_lalu.year) &
                (df["Tanggal Kehadiran"].dt.month == periode_lalu.month)
            ].copy()

            label_periode = (
                f"{nama_bulan[periode_terpilih.month]} "
                f"{periode_terpilih.year}"
            )

            nama_file_periode = (
                f"{nama_bulan[periode_terpilih.month]}_"
                f"{periode_terpilih.year}"
            )

            label_perbandingan = "dari bulan lalu"


        # Judul dashboard mengikuti periode
        judul_dashboard.subheader(
            f"Ringkasan Kunjungan PST {label_periode}"
        )


        # =========================
        # 1. TOTAL TAMU
        # =========================

        total_tamu = len(df_periode)
        total_periode_lalu = len(df_periode_lalu)

        if total_periode_lalu > 0:
            perubahan_tamu = (
                (total_tamu - total_periode_lalu)
                / total_periode_lalu
            ) * 100

            delta_total = (
                f"{perubahan_tamu:+.0f}% "
                f"{label_perbandingan}"
            )
        else:
            delta_total = None


        # =========================
        # 2. KEPERLUAN TERBANYAK
        # =========================

        distribusi_keperluan = (
            df_periode["Keperluan"]
            .dropna()
            .value_counts()
        )

        if not distribusi_keperluan.empty:
            keperluan_terbanyak = distribusi_keperluan.idxmax()
            jumlah_terbanyak = distribusi_keperluan.max()

            persen_terbanyak = (
                jumlah_terbanyak / total_tamu * 100
                if total_tamu > 0
                else 0
            )

            delta_keperluan = (
                f"{persen_terbanyak:.0f}% dari total"
            )

        else:
            keperluan_terbanyak = "-"
            delta_keperluan = None


        # =========================
        # 3. DURASI KUNJUNGAN
        # =========================

        tanggal_string = (
            df_periode["Tanggal Kehadiran"]
            .dt.strftime("%Y-%m-%d")
        )

        waktu_datang = pd.to_datetime(
            tanggal_string + " " +
            df_periode["Jam Datang"].astype(str),
            errors="coerce"
        )

        waktu_pulang = pd.to_datetime(
            tanggal_string + " " +
            df_periode["Jam Pulang"].astype(str),
            errors="coerce"
        )

        durasi = (
            waktu_pulang - waktu_datang
        ).dt.total_seconds() / 60

        durasi = durasi[durasi >= 0]

        if not durasi.empty:
            rata_rata_durasi = durasi.mean()
            nilai_durasi = f"{rata_rata_durasi:.0f} Menit"
        else:
            rata_rata_durasi = None
            nilai_durasi = "-"


        # Durasi periode sebelumnya
        delta_durasi = None

        if not df_periode_lalu.empty:

            tanggal_lalu = (
                df_periode_lalu["Tanggal Kehadiran"]
                .dt.strftime("%Y-%m-%d")
            )

            datang_lalu = pd.to_datetime(
                tanggal_lalu + " " +
                df_periode_lalu["Jam Datang"].astype(str),
                errors="coerce"
            )

            pulang_lalu = pd.to_datetime(
                tanggal_lalu + " " +
                df_periode_lalu["Jam Pulang"].astype(str),
                errors="coerce"
            )

            durasi_lalu = (
                pulang_lalu - datang_lalu
            ).dt.total_seconds() / 60

            durasi_lalu = durasi_lalu[
                durasi_lalu >= 0
            ]

            if (
                not durasi_lalu.empty
                and rata_rata_durasi is not None
            ):
                selisih = (
                    rata_rata_durasi -
                    durasi_lalu.mean()
                )

                delta_durasi = f"{selisih:+.0f} Menit"


        # =========================
        # DATA BAR CHART
        # =========================

        data_kunjungan = (
            df_periode["Keperluan"]
            .dropna()
            .value_counts()
            .reset_index()
        )

        data_kunjungan.columns = [
            "Keperluan Kunjungan",
            "Jumlah Kunjungan"
        ]

        data_kunjungan = (
            data_kunjungan
            .sort_values("Jumlah Kunjungan")
        )


        # =========================
        # DATA DONUT CHART
        # =========================

        data_jenis = (
            df_periode["Jenis Tamu"]
            .dropna()
            .value_counts()
            .reset_index()
        )

        data_jenis.columns = [
            "Jenis Tamu",
            "Jumlah"
        ]


        # =========================
        # DATA EXCEL
        # =========================

        (
            df_ringkasan_download,
            data_kunjungan_download,
            data_jenis_download
        ) = siapkan_data_excel(df_periode)

        rekap_bulanan = buat_rekap_bulanan(
            df_periode,
            nama_bulan,
            tambah_total=is_triwulan,
            bulan_wajib=bulan_triwulan
        )

        excel_overview = buat_excel_overview(
            df_ringkasan_download,
            data_kunjungan_download,
            data_jenis_download,
            df_periode,
            rekap_bulanan,
            nama_file_periode
        )


        # Tombol selalu aktif
        with col_download:
            st.download_button(
                label="📥 Unduh Data",
                data=excel_overview,
                file_name=(
                    f"Overview_Kunjungan_Tamu_"
                    f"{nama_file_periode}.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        # =========================
        # METRIC CARDS
        # =========================
        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Total Tamu",
            f"{total_tamu} Orang",
            delta=delta_total
        )

        col2.metric(
            "Keperluan Terbanyak",
            keperluan_terbanyak,
            delta=delta_keperluan
        )

        col3.metric(
            "Rata-Rata Waktu Kunjungan",
            nilai_durasi,
            delta=delta_durasi,
            delta_color="inverse"
        )
            
        # =========================
        # CHART
        # =========================
        col_grafik, col_info = st.columns([65, 35])

        PALETTE_BIRU = [
            "#C6E7FF",
            "#64B5F6",
            "#1E88E5",
            "#0D47A1"
        ]

        with col_grafik:

            with st.container(border=True):

                fig = px.bar(
                    data_kunjungan,
                    x="Jumlah Kunjungan",
                    y="Keperluan Kunjungan",
                    orientation="h",
                    text="Jumlah Kunjungan",
                    title=(
                        "Distribusi Keperluan "
                        "Kunjungan Tamu PST"
                    ),
                    color="Jumlah Kunjungan",
                    color_continuous_scale=PALETTE_BIRU
                )

                fig.update_traces(
                    textposition="outside",
                    textfont=dict(size=13),
                    cliponaxis=False
                )

                fig.update_layout(
                    height=250,
                    margin=dict(
                        t=30,
                        b=10,
                        l=10,
                        r=30
                    ),
                    coloraxis_showscale=False,
                    xaxis_title="Jumlah Kunjungan",
                    yaxis_title="",
                    plot_bgcolor="rgba(0,0,0,0)"
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )

        with col_info:

            with st.container(border=True):

                fig_donut = px.pie(
                    data_jenis,
                    names="Jenis Tamu",
                    values="Jumlah",
                    hole=0.5,
                    title="Proporsi Jenis Tamu",
                    color_discrete_sequence=PALETTE_BIRU
                )

                fig_donut.update_traces(
                    textinfo="label+percent",
                    textposition="outside",
                    textfont=dict(size=11),
                    selector=dict(type="pie"),
                    domain=dict(
                        x=[0.18, 0.88],
                        y=[0.12, 0.88]
                    )
                )

                fig_donut.update_layout(
                    height=250,
                    margin=dict(
                        t=30,
                        b=5,
                        l=75,
                        r=70
                    ),
                    showlegend=False,
                    uniformtext_minsize=9,
                    uniformtext_mode="show"
                )

                st.plotly_chart(
                    fig_donut,
                    use_container_width=True
                )

    overview_kunjungan_pst()

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