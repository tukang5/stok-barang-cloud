import io
from datetime import datetime
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# Set konfigurasi halaman web
st.set_page_config(
    page_title="Sistem Manajemen Stok & Log", page_icon="📦", layout="wide"
)

# === 1. KONFIGURASI LINK GOOGLE SHEETS ===
# SILAKAN PASTE URL GOOGLE SHEETS ANDA DI SINI
URL_SPREADSHEET = "https://docs.google.com/spreadsheets/d/1hCERQHti6BaAVJYqXWDtoubKuX_8EzHU_QdDKS3gB5w/edit?usp=sharing"

# Inisialisasi koneksi ke Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)


# === 2. SISTEM AUTENTIKASI / LOGIN ===
def sistem_login():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.subheader("🔒 Silakan Login Terlebih Dahulu")
            username = st.text_input("Username Administrator")
            password = st.text_input("Password", type="password")

            if st.button("Masuk 🔓", type="primary", use_container_width=True):
                if username == "tukang5" and password == "iduladha#15":
                    st.session_state["logged_in"] = True
                    st.success("Login Berhasil!")
                    st.rerun()
                else:
                    st.error("Username atau Password salah!")
        return False
    return True


# === 3. FUNGSI LOGIKA DATABASE GOOGLE SHEETS ===
def ambil_data(cari=""):
    try:
        df = conn.read(spreadsheet=URL_SPREADSHEET, worksheet="barang")
        # Bersihkan data kosong
        df = df.dropna(subset=["nama"])
        df["stok"] = pd.to_numeric(df["stok"]).astype(int)
        df["harga"] = pd.to_numeric(df["harga"]).astype(float)
        df["id"] = pd.to_numeric(df["id"]).astype(int)

        if cari:
            df = df[df["nama"].str.contains(cari, case=False, na=False)]
        return df
    except Exception:
        # Jika sheet kosong/baru dibuat, kembalikan dataframe kosong ber-struktur
        return pd.DataFrame(columns=["id", "nama", "stok", "harga"])


def ambil_riwayat():
    try:
        df = conn.read(spreadsheet=URL_SPREADSHEET, worksheet="riwayat")
        df = df.dropna(subset=["waktu"])
        # Urutkan dari yang terbaru (ID terbesar)
        if not df.empty:
            df["id"] = pd.to_numeric(df["id"]).astype(int)
            df = df.sort_values(by="id", ascending=False)
        return df
    except Exception:
        return pd.DataFrame(
            columns=["id", "waktu", "nama_barang", "tipe", "jumlah", "keterangan"]
        )


def tambah_barang_sheets(nama, stok, harga):
    df_lama = ambil_data()
    next_id = 1 if df_lama.empty else int(df_lama["id"].max()) + 1

    # Cek duplikasi nama
    if not df_lama.empty and nama.lower() in df_lama["nama"].str.lower().values:
        st.error(f"Barang dengan nama '{nama}' sudah ada!")
        return False

    df_baru = pd.DataFrame(
        [[next_id, nama, stok, harga]], columns=["id", "nama", "stok", "harga"]
    )
    df_total = pd.concat([df_lama, df_baru], ignore_index=True)

    conn.update(
        spreadsheet=URL_SPREADSHEET, worksheet="barang", data=df_total
    )
    catat_log(
        nama, "Barang Baru", stok, f"Pendaftaran barang baru dengan stok awal {stok}"
    )
    return True


def update_stok_sheets(id_barang, nama, stok_akhir, harga_baru, tipe_log, jumlah_mutasi, keterangan):
    df_lama = ambil_data()
    df_lama.loc[df_lama["id"] == id_barang, ["stok", "harga"]] = [
        stok_akhir,
        harga_baru,
    ]

    conn.update(
        spreadsheet=URL_SPREADSHEET, worksheet="barang", data=df_lama
    )
    catat_log(nama, tipe_log, jumlah_mutasi, keterangan)


def hapus_barang_sheets(nama):
    df_lama = ambil_data()
    df_baru = df_lama[df_lama["nama"].str.lower() != nama.lower()]

    conn.update(
        spreadsheet=URL_SPREADSHEET, worksheet="barang", data=df_baru
    )
    catat_log(nama, "Hapus", 0, "Barang dihapus permanen dari sistem")


def catat_log(nama_barang, tipe, jumlah, keterangan):
    df_riwayat_lama = ambil_riwayat()
    next_id = 1 if df_riwayat_lama.empty else int(df_riwayat_lama["id"].max()) + 1
    waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df_log_baru = pd.DataFrame(
        [[next_id, waktu_sekarang, nama_barang, tipe, jumlah, keterangan]],
        columns=["id", "waktu", "nama_barang", "tipe", "jumlah", "keterangan"],
    )
    df_total_log = pd.concat([df_riwayat_lama, df_log_baru], ignore_index=True)

    conn.update(
        spreadsheet=URL_SPREADSHEET, worksheet="riwayat", data=df_total_log
    )


# === 4. FUNGSI LAINNYA ===
def konversi_ke_excel(df, sheet_name="Data"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def beri_warna_stok(row):
    if row["Stok"] <= 2:
        return ["background-color: #ffcccc; color: #b30000; font-weight: bold"] * len(
            row
        )
    return [""] * len(row)


# === 5. LOGIKA UTAMA APLIKASI ===
if sistem_login():
    st.sidebar.title("📌 Menu Navigasi")
    menu = st.sidebar.radio(
        "Pilih Halaman:", ["Stok Barang Utama", "📋 Riwayat / Log Transaksi"]
    )

    st.sidebar.write("---")
    if st.sidebar.button("🚪 Keluar / Logout", use_container_width=True):
        st.session_state["logged_in"] = False
        st.rerun()

    # --- HALAMAN 1: STOK BARANG UTAMA ---
    if menu == "Stok Barang Utama":
        st.title("📦 Sistem Manajemen Stok Barang (Cloud Data)")
        st.markdown(
            "Terhubung otomatis dengan Google Sheets. Data aman anti-hilang."
        )
        st.markdown("---")

        kolom_kiri, kolom_kanan = st.columns([2, 3], gap="large")

        with kolom_kiri:
            st.subheader("📝 Formulir Barang")
            mode = st.radio(
                "Pilih Tindakan:",
                [
                    "Tambah Barang Baru",
                    "Update Stok Masuk/Keluar",
                    "Hapus Barang",
                ],
            )

            if mode == "Tambah Barang Baru":
                nama = st.text_input("Nama Barang")
                stok = st.number_input("Jumlah Stok Awal", min_value=0, step=1)
                harga = st.number_input("Harga Satuan (Rp)", min_value=0.0)

                if st.button("➕ Tambah Barang", type="primary"):
                    if not nama:
                        st.error("Nama barang tidak boleh kosong!")
                    else:
                        jika_sukses = tambah_barang_sheets(nama, stok, harga)
                        if jika_sukses:
                            st.success(f"Barang '{nama}' berhasil ditambahkan!")
                            st.rerun()

            elif mode == "Update Stok Masuk/Keluar":
                df_pilihan = ambil_data()
                if df_pilihan.empty:
                    st.warning("Belum ada data barang di Google Sheets.")
                else:
                    pilihan_barang = st.selectbox(
                        "Pilih Barang:", df_pilihan["nama"].tolist()
                    )
                    data_barang = df_pilihan[
                        df_pilihan["nama"] == pilihan_barang
                    ].iloc[0]

                    st.info(
                        f"Stok saat ini: **{data_barang['stok']} Pcs** | Harga: **Rp {data_barang['harga']:,.0f}**"
                    )

                    jenis_opsi = st.selectbox(
                        "Jenis Mutasi:", ["Stok Masuk (+)", "Stok Keluar (-)"]
                    )
                    jumlah_mutasi = st.number_input(
                        "Jumlah Perubahan Stok", min_value=1, step=1
                    )
                    keterangan = st.text_input(
                        "Keterangan / Catatan tambahan",
                        placeholder="Contoh: Restock / Terjual",
                    )
                    harga_baru = st.number_input(
                        "Perbarui Harga (Biarkan jika tetap)",
                        value=float(data_barang["harga"]),
                    )

                    if st.button("🔄 Proses Perubahan", type="primary"):
                        stok_akhir = int(data_barang["stok"])
                        tipe_log = "Masuk" if jenis_opsi == "Stok Masuk (+)" else "Keluar"

                        if jenis_opsi == "Stok Masuk (+)":
                            stok_akhir += jumlah_mutasi
                        else:
                            stok_akhir -= jumlah_mutasi

                        if stok_akhir < 0:
                            st.error("Gagal! Stok tidak boleh kurang dari 0.")
                        else:
                            update_stok_sheets(
                                int(data_barang["id"]),
                                pilihan_barang,
                                stok_akhir,
                                harga_baru,
                                tipe_log,
                                jumlah_mutasi,
                                keterangan,
                            )
                            st.success("Data berhasil diperbarui ke Cloud!")
                            st.rerun()

            elif mode == "Hapus Barang":
                df_pilihan = ambil_data()
                if df_pilihan.empty:
                    st.warning("Belum ada data barang.")
                else:
                    pilihan_barang = st.selectbox(
                        "Pilih Barang yang akan dihapus:",
                        df_pilihan["nama"].tolist(),
                    )
                    if st.button(
                        "🗑️ Hapus Permanen dari Gudang", type="secondary"
                    ):
                        hapus_barang_sheets(pilihan_barang)
                        st.success("Barang berhasil dihapus dari Cloud!")
                        st.rerun()

        with kolom_kanan:
            st.subheader("📋 Daftar Stok Gudang")
