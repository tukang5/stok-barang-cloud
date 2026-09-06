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

        col1, col2, col3 = st.columns(3)
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
                        cek_lisensi = supabase.table("lisensi").select("*").ilike("kode_kunci", input_lisensi.strip()).execute()
                        lisensi_valid = False
                        if cek_lisensi.data and len(cek_lisensi.data) > 0:
                            data_kunci = cek_lisensi.data[0] 
                            if str(data_kunci.get("status", "")).lower() == "tersedia":
                                lisensi_valid = True

# 3. Jika lisensi ditemukan dan berstatus tersedia, loloskan sistem!
if lisensi_valid:

                        if lisensi_valid:
                            try:
                                supabase.table("pengguna").insert({"username": buat_user.strip(), "password": buat_pass.strip()}).execute()
                                supabase.table("lisensi").update({"status": "Terpakai"}).eq("kode_kunci", input_lisensi.strip().upper()).execute()
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
                            try:
                                supabase.table("barang").update({"stok": stok_akhir, "harga": harga_baru}).eq("id", int(data_barang["id"])).execute()
                                catat_log(pilihan_barang, tipe_log, jumlah_mutasi, keterangan)
                                st.success("Stok berhasil diperbarui!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Eror Perubahan: {str(e)}")

            elif mode == "Hapus Barang":
                df_pilihan = ambil_data()
                if df_pilihan.empty:
                    st.warning("Belum ada data barang.")
                else:
                    pilihan_barang = st.selectbox("Pilih Barang yang akan dihapus:", df_pilihan["nama"].tolist())
                    if st.button("🗑️ Hapus Permanen", type="secondary"):
                        try:
                            supabase.table("barang").delete().eq("nama", pilihan_barang).execute()
                            catat_log(pilihan_barang, "Hapus", 0, "Barang dihapus permanen dari sistem")
                            st.success("Barang berhasil dihapus!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Eror Hapus: {str(e)}")

        with kolom_kanan:
            st.subheader("📋 Daftar Stok Gudang Real-time")
            cari_input = st.text_input("🔍 Cari Nama Barang...", placeholder="Ketik untuk memfilter...")
            df_stok = ambil_data(cari_input)

            if not df_stok.empty:
                df_tampil = df_stok.copy()
                df_tampil.columns = ["ID", "Nama Barang", "Stok", "Harga (Rp)"]

                stok_kritis = df_tampil[df_tampil["Stok"] <= 2]
                if not stok_kritis.empty:
                    st.error(f"🚨 **Peringatan:** Ada {len(stok_kritis)} barang dengan stok kritis (≤ 2 pcs)!", icon="⚠️")

                df_tampil["Harga (Rp)"] = df_tampil["Harga (Rp)"].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
                st.table(df_tampil)

                st.markdown("---")
                m1, m2, m3 = st.columns(3)
                m1.metric(label="Total Jenis Barang", value=f"{len(df_tampil)} Item")
                m2.metric(label="Total Seluruh Stok", value=f"{df_stok['stok'].sum()} Pcs")

                with m3:
                    st.write("📥 **Unduh Laporan**")
                    data_excel = konversi_ke_excel(df_tampil, "Daftar Stok")
                    st.download_button(label="🟢 Ekspor ke Excel (.xlsx)", data=data_excel, file_name="laporan_stok_barang.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else:
                st.info("Tidak ada data barang ditemukan.")

    elif menu == "📋 Riwayat / Log Transaksi":
        st.title("📋 Log & Riwayat Mutasi Stok")
        st.markdown("---")
        df_riwayat = ambil_riwayat()

        if not df_riwayat.empty:
            df_riwayat_tampil = df_riwayat.copy()
            df_riwayat_tampil.columns = ["Waktu & Tanggal", "Nama Barang", "Tipe Aktivitas", "Jumlah (Pcs)", "Keterangan / Catatan"]
            
            st.table(df_riwayat_tampil)
            st.markdown("---")
            
            st.subheader("🖨️ Cetak Nota Transaksi Keluar")
            df_keluar = df_riwayat_tampil[df_riwayat_tampil["Tipe Aktivitas"] == "Keluar"]
            
            if df_keluar.empty:
                st.info("Belum ada transaksi 'Stok Keluar' yang bisa dibuatkan nota.")
            else:
                opsi_nota = [f"{row['Waktu & Tanggal']} - {row['Nama Barang']} ({row['Jumlah (Pcs)']} Pcs)" for _, row in df_keluar.iterrows()]
                transaksi_terpilih = st.selectbox("Pilih Transaksi yang Ingin Dicetak Notanya:", opsi_nota)
                
                indeks_pilihan = opsi_nota.index(transaksi_terpilih)
                data_nota = df_keluar.iloc[indeks_pilihan]
                
                html_nota = f"""
                <div style="font-family: 'Courier New', Courier, monospace; width: 280px; padding: 15px; border: 1px dashed #000; background-color: #fff; color: #000; margin: 0 auto;">
                    <div style="text-align: center; font-weight: bold; font-size: 14px;">NOTA PENGELUARAN STOK</div>
                    <div style="text-align: center; font-size: 11px; margin-bottom: 10px;">ENTERPRISE CLOUD SYSTEM</div>
                    <hr style="border-top: 1px dashed #000;">
                    <table style="width: 100%; font-size: 11px;">
                        <tr><td>Waktu</td><td>: {data_nota['Waktu & Tanggal']}</td></tr>
                        <tr><td>Barang</td><td>: {data_nota['Nama Barang']}</td></tr>
                        <tr><td>Jumlah</td><td>: {data_nota['Jumlah (Pcs)']} Pcs</td></tr>
                        <tr><td>Status</td><td>: BERHASIL (KELUAR)</td></tr>
                    </table>
                    <hr style="border-top: 1px dashed #000;">
                    <div style="font-size: 11px; word-wrap: break-word;">
                        <strong>Catatan/Keterangan:</strong><br>
                        {data_nota['Keterangan / Catatan'] if data_nota['Keterangan / Catatan'] else '-'}
                    </div>
                    <hr style="border-top: 1px dashed #000;">
                    <div style="text-align: center; font-size: 10px; margin-top: 10px; font-style: italic;">
                        Terima kasih atas kerja samanya.<br>Dokumen sah sistem cloud gudang.
                    </div>
                </div>
                """
                st.markdown("### 🔍 Pratinjau Nota:")
                st.html(html_nota)
                
                st.markdown("<br>", unsafe_allow_html=True)
                js_cetak = """
                <script>
                function cetakNota() {
                    window.print();
                }
                </script>
                <button onclick="cetakNota()" style="width: 100%; background-color: #ff4b4b; color: white; border: none; padding: 10px; border-radius: 8px; font-weight: bold; cursor: pointer;">
                    🖨️ Cetak Transaksi Sekarang / Simpan ke PDF
                </button>
                """
                st.components.v1.html(js_cetak, height=50)

            st.markdown("---")
            data_excel_log = konversi_ke_excel(df_riwayat_tampil, "Log Riwayat")
            st.download_button(label="🟢 Unduh Seluruh Log Riwayat (.xlsx)", data=data_excel_log, file_name="riwayat_mutasi_stok.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("Belum ada riwayat aktivitas.")
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
                            try:
                                supabase.table("barang").update({"stok": stok_akhir, "harga": harga_baru}).eq("id", int(data_barang["id"])).execute()
                                catat_log(pilihan_barang, tipe_log, jumlah_mutasi, keterangan)
                                st.success("Stok berhasil diperbarui!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Eror Perubahan: {str(e)}")

            elif mode == "Hapus Barang":
                df_pilihan = ambil_data()
                if df_pilihan.empty:
                    st.warning("Belum ada data barang.")
                else:
                    pilihan_barang = st.selectbox("Pilih Barang yang akan dihapus:", df_pilihan["nama"].tolist())
                    if st.button("🗑️ Hapus Permanen", type="secondary"):
                        try:
                            supabase.table("barang").delete().eq("nama", pilihan_barang).execute()
                            catat_log(pilihan_barang, "Hapus", 0, "Barang dihapus permanen dari sistem")
                            st.success("Barang berhasil dihapus!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Eror Hapus: {str(e)}")

        with kolom_kanan:
            st.subheader("📋 Daftar Stok Gudang Real-time")
            cari_input = st.text_input("🔍 Cari Nama Barang...", placeholder="Ketik untuk memfilter...")
            df_stok = ambil_data(cari_input)

            if not df_stok.empty:
                df_tampil = df_stok.copy()
                df_tampil.columns = ["ID", "Nama Barang", "Stok", "Harga (Rp)"]

                stok_kritis = df_tampil[df_tampil["Stok"] <= 2]
                if not stok_kritis.empty:
                    st.error(f"🚨 **Peringatan:** Ada {len(stok_kritis)} barang dengan stok kritis (≤ 2 pcs)!", icon="⚠️")

                df_tampil["Harga (Rp)"] = df_tampil["Harga (Rp)"].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
                st.table(df_tampil)

                st.markdown("---")
                m1, m2, m3 = st.columns(3)
                m1.metric(label="Total Jenis Barang", value=f"{len(df_tampil)} Item")
                m2.metric(label="Total Seluruh Stok", value=f"{df_stok['stok'].sum()} Pcs")

                with m3:
                    st.write("📥 **Unduh Laporan**")
                    data_excel = konversi_ke_excel(df_tampil, "Daftar Stok")
                    st.download_button(label="🟢 Ekspor ke Excel (.xlsx)", data=data_excel, file_name="laporan_stok_barang.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else:
                st.info("Tidak ada data barang ditemukan.")

    elif menu == "📋 Riwayat / Log Transaksi":
        st.title("📋 Log & Riwayat Mutasi Stok")
        st.markdown("---")
        df_riwayat = ambil_riwayat()

        if not df_riwayat.empty:
            df_riwayat_tampil = df_riwayat.copy()
            df_riwayat_tampil.columns = ["Waktu & Tanggal", "Nama Barang", "Tipe Aktivitas", "Jumlah (Pcs)", "Keterangan / Catatan"]
            
            st.table(df_riwayat_tampil)
            st.markdown("---")
            
            st.subheader("🖨️ Cetak Nota Transaksi Keluar")
            df_keluar = df_riwayat_tampil[df_riwayat_tampil["Tipe Aktivitas"] == "Keluar"]
            
            if df_keluar.empty:
                st.info("Belum ada transaksi 'Stok Keluar' yang bisa dibuatkan nota.")
            else:
                opsi_nota = [f"{row['Waktu & Tanggal']} - {row['Nama Barang']} ({row['Jumlah (Pcs)']} Pcs)" for _, row in df_keluar.iterrows()]
                transaksi_terpilih = st.selectbox("Pilih Transaksi yang Ingin Dicetak Notanya:", opsi_nota)
                
                indeks_pilihan = opsi_nota.index(transaksi_terpilih)
                data_nota = df_keluar.iloc[indeks_pilihan]
                
                html_nota = f"""
                <div style="font-family: 'Courier New', Courier, monospace; width: 280px; padding: 15px; border: 1px dashed #000; background-color: #fff; color: #000; margin: 0 auto;">
                    <div style="text-align: center; font-weight: bold; font-size: 14px;">NOTA PENGELUARAN STOK</div>
                    <div style="text-align: center; font-size: 11px; margin-bottom: 10px;">ENTERPRISE CLOUD SYSTEM</div>
                    <hr style="border-top: 1px dashed #000;">
                    <table style="width: 100%; font-size: 11px;">
                        <tr><td>Waktu</td><td>: {data_nota['Waktu & Tanggal']}</td></tr>
                        <tr><td>Barang</td><td>: {data_nota['Nama Barang']}</td></tr>
                        <tr><td>Jumlah</td><td>: {data_nota['Jumlah (Pcs)']} Pcs</td></tr>
                        <tr><td>Status</td><td>: BERHASIL (KELUAR)</td></tr>
                    </table>
                    <hr style="border-top: 1px dashed #000;">
                    <div style="font-size: 11px; word-wrap: break-word;">
                        <strong>Catatan/Keterangan:</strong><br>
                        {data_nota['Keterangan / Catatan'] if data_nota['Keterangan / Catatan'] else '-'}
                    </div>
                    <hr style="border-top: 1px dashed #000;">
                    <div style="text-align: center; font-size: 10px; margin-top: 10px; font-style: italic;">
                        Terima kasih atas kerja samanya.<br>Dokumen sah sistem cloud gudang.
                    </div>
                </div>
                """
                st.markdown("### 🔍 Pratinjau Nota:")
                st.html(html_nota)
                
                st.markdown("<br>", unsafe_allow_html=True)
                js_cetak = """
                <script>
                function cetakNota() {
                    window.print();
                }
                </script>
                <button onclick="cetakNota()" style="width: 100%; background-color: #ff4b4b; color: white; border: none; padding: 10px; border-radius: 8px; font-weight: bold; cursor: pointer;">
                    🖨️ Cetak Transaksi Sekarang / Simpan ke PDF
                </button>
                """
                st.components.v1.html(js_cetak, height=50)

            st.markdown("---")
            data_excel_log = konversi_ke_excel(df_riwayat_tampil, "Log Riwayat")
            st.download_button(label="🟢 Unduh Seluruh Log Riwayat (.xlsx)", data=data_excel_log, file_name="riwayat_mutasi_stok.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("Belum ada riwayat aktivitas.")
