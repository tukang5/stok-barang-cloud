import io
import pandas as pd
import streamlit as st
from supabase import create_client, Client

# Set konfigurasi halaman web
st.set_page_config(
    page_title="Sistem Manajemen Stok & Log Cloud", page_icon="📦", layout="wide"
)

# === 1. KONEKSI SUPABASE ===
def inisialisasi_supabase() -> Client:
    url: str = st.secrets["SUPABASE_URL"]
    key: str = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = inisialisasi_supabase()

# === 2. SISTEM KEAMANAN LISENSI & LOGIN DINAMIS ===
def sistem_login():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        # Cek apakah sudah ada akun pengguna terdaftar di database
        try:
            cek_user = supabase.table("pengguna").select("*").execute()
            ada_user = len(cek_user.data) > 0
        except Exception:
            ada_user = False

        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.write("")
            
            # JIKA BELUM ADA AKUN (Aplikasi Baru Pertama Kali Dibuka Klien)
            if not ada_user:
                st.subheader("🔑 Aktivasi Kunci Lisensi Aplikasi Baru")
                st.info("Aplikasi belum diaktivasi. Silakan masukkan Kunci Lisensi resmi dari Developer.")
                
                input_lisensi = st.text_input("Masukkan Kunci Lisensi (License Key)")
                buat_user = st.text_input("Buat Username Baru untuk Toko Anda")
                buat_pass = st.text_input("Buat Password Baru", type="password")
                
                if st.button("Aktifkan Aplikasi ✨", type="primary", use_container_width=True):
                    if not input_lisensi or not buat_user or not buat_pass:
                        st.error("Semua kolom pengisian wajib diisi!")
                    else:
                        cek_lisensi = supabase.table("lisensi").select("*").eq("kode_kunci", input_lisensi.strip().upper()).eq("status", "Tersedia").execute()
                        
                        if cek_lisensi.data:
                            try:
                                supabase.table("pengguna").insert({"username": buat_user.strip(), "password": buat_pass.strip()}).execute()
                                supabase.table("lisensi").update({"status": "Terpakai"}).eq("kode_kunci", input_lisensi.strip()).execute()
                                st.success("Aktivasi Sukses! Silakan muat ulang halaman untuk masuk.")
                                st.rerun()
                            except Exception:
                                st.error("Username tersebut sudah digunakan. Silakan pilih nama lain.")
                        else:
                            st.error("Kunci Lisensi Salah atau sudah kadaluwarsa/terpakai!")
            
            # JIKA SUDAH ADA AKUN (Kondisi Normal)
            else:
                st.subheader("🔒 Silakan Login Terlebih Dahulu")
                username = st.text_input("Username Toko")
                password = st.text_input("Password", type="password")

                if st.button("Masuk 🔓", type="primary", use_container_width=True):
                    fitur_cek = supabase.table("pengguna").select("*").eq("username", username.strip()).eq("password", password.strip()).execute()
                    if fitur_cek.data:
                        st.session_state["logged_in"] = True
                        st.success("Login Berhasil!")
                        st.rerun()
                    else:
                        st.error("Username atau Password salah!")
        return False
    return True

# === 3. FUNGSI LOGIKA DATABASE BARANG ===
def ambil_data(cari=""):
    try:
        if cari:
            response = supabase.table("barang").select("*").ilike("nama", f"%{cari}%").execute()
        else:
            response = supabase.table("barang").select("*").execute()
        
        df = pd.DataFrame(response.data)
        if df.empty:
            return pd.DataFrame(columns=["id", "nama", "stok", "harga"])
        return df[["id", "nama", "stok", "harga"]]
    except Exception:
        return pd.DataFrame(columns=["id", "nama", "stok", "harga"])

def ambil_riwayat():
    try:
        response = supabase.table("riwayat").select("*").order("id", desc=True).execute()
        df = pd.DataFrame(response.data)
        if df.empty:
            return pd.DataFrame(columns=["waktu", "nama_barang", "tipe", "jumlah", "keterangan"])
        df["waktu"] = pd.to_datetime(df["waktu"]).dt.strftime('%Y-%m-%d %H:%M:%S')
        return df[["waktu", "nama_barang", "tipe", "jumlah", "keterangan"]]
    except Exception:
        return pd.DataFrame(columns=["waktu", "nama_barang", "tipe", "jumlah", "keterangan"])

def catat_log(nama_barang, tipe, jumlah, keterangan):
    try:
        supabase.table("riwayat").insert({
            "nama_barang": nama_barang, "tipe": tipe, "jumlah": jumlah, "keterangan": keterangan
        }).execute()
    except Exception:
        pass

# === 4. FUNGSI PENDUKUNG ===
def konversi_ke_excel(df, sheet_name="Data"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def beri_warna_stok(row):
    if row["Stok"] <= 2:
        return ["background-color: #ffcccc; color: #b30000; font-weight: bold"] * len(row)
    return [""] * len(row)

# === 5. LOGIKA UTAMA TAMPILAN APLIKASI ===
if sistem_login():
    st.sidebar.title("📌 Menu Navigasi")
    menu = st.sidebar.radio("Pilih Halaman:", ["Stok Barang Utama", "📋 Riwayat / Log Transaksi"])
    st.sidebar.write("---")
    if st.sidebar.button("🚪 Keluar / Logout", use_container_width=True):
        st.session_state["logged_in"] = False
        st.rerun()

    if menu == "Stok Barang Utama":
        st.title("📦 Sistem Manajemen Stok Barang (Enterprise Cloud)")
        st.markdown("Aplikasi manajemen stok berskala industri retail, grosir, dan gudang pabrik.")
        st.markdown("---")

        kolom_kiri, kolom_kanan = st.columns([1, 2], gap="large")

        with kolom_kiri:
            st.subheader("📝 Formulir Barang")
            mode = st.radio("Pilih Tindakan:", ["Tambah Barang Baru", "Update Stok Masuk/Keluar", "Hapus Barang"])

            if mode == "Tambah Barang Baru":
                nama = st.text_input("Nama Barang")
                stok = st.number_input("Jumlah Stok Awal", min_value=0, step=1)
                harga = st.number_input("Harga Satuan (Rp)", min_value=0.0)

                if st.button("➕ Tambah Barang", type="primary"):
                    if not nama:
                        st.error("Nama barang tidak boleh kosong!")
                    else:
                        try:
                            supabase.table("barang").insert({"nama": nama, "stok": stok, "harga": harga}).execute()
                            catat_log(nama, "Barang Baru", stok, f"Pendaftaran barang baru dengan stok awal {stok}")
                            st.success(f"Barang '{nama}' berhasil didaftarkan di Cloud!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Eror Sistem: {str(e)}")

            elif mode == "Update Stok Masuk/Keluar":
                df_pilihan = ambil_data()
                if df_pilihan.empty:
                    st.warning("Belum ada data barang di database.")
                else:
                    pilihan_barang = st.selectbox("Pilih Barang:", df_pilihan["nama"].tolist())
                    data_barang = df_pilihan[df_pilihan["nama"] == pilihan_barang].iloc

                    st.info(f"Stok saat ini: **{data_barang['stok']} Pcs** | Harga: **Rp {float(data_barang['harga']):,.0f}**")
                    jenis_opsi = st.selectbox("Jenis Mutasi:", ["Stok Masuk (+)", "Stok Keluar (-)"])
                    jumlah_mutasi = st.number_input("Jumlah Perubahan Stok", min_value=1, step=1)
                    keterangan = st.text_input("Keterangan / Catatan tambahan", placeholder="Contoh: Restock Supplier")
                    harga_baru = st.number_input("Perbarui Harga (Biarkan jika tetap)", value=float(data_barang["harga"]))

                    if st.button("🔄 Proses Perubahan", type="primary"):
                        stok_akhir = int(data_barang["stok"])
                        tipe_log = "Masuk" if jenis_opsi == "Stok Masuk (+)" else "Keluar"
                        stok_akhir = stok_akhir + jumlah_mutasi if jenis_opsi == "Stok Masuk (+)" else stok_akhir - jumlah_mutasi

                        if stok_akhir < 0:
                            st.error("Gagal! Stok tidak boleh kurang dari 0.")
                        else:
                            supabase.table("barang").update({"stok": stok_akhir, "harga": harga_baru}).eq("id", int(data_barang["id"])).execute()
                            catat_log(pilihan_barang, tipe_log, jumlah_mutasi, keterangan)
                            st.success("Stok berhasil diperbarui!")
                            st.rerun()

            elif mode == "Hapus Barang":
                df_pilihan = ambil_data()
                if df_pilihan.empty:
                    st.warning("Belum ada data barang.")
                else:
                    pilihan_barang = st.selectbox("Pilih Barang yang akan dihapus:", df_pilihan["nama"].tolist())
                    if st.button("🗑️ Hapus Permanen", type="secondary"):
                        supabase.table("barang").delete().eq("nama", pilihan_barang).execute()
                        catat_log(pilihan_barang, "Hapus", 0, "Barang dihapus permanen dari sistem")
                        st.success("Barang berhasil dihapus!")
                        st.rerun()

        with kolom_kanan:
            st.subheader("📋 Daftar Stok Gudang Real-time")
            cari_input = st.text_input("🔍 Cari Nama Barang...", placeholder="Ketik untuk memfilter...")
