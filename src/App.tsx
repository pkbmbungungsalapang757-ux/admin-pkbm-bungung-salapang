import React, { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";

// GANTI dengan URL Web App hasil deploy Google Apps Script kamu
const API_URL =
  "https://script.google.com/macros/s/AKfycbz8iQvJZhIAFS48tuJTjQZskEN9envcQAlfgU4HvZTF-3EFlBs_w-aA4M5vo2MdBgQHLw/exec";

// Kredensial login (disimpan langsung di kode)
const CREDENTIALS = {
  username: "admin",
  password: "1234",
};

interface Siswa {
  nisn: string;
  nama: string;
  kelas: string;
}

export default function App() {
  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(
    () => sessionStorage.getItem("isLoggedIn") === "true"
  );
  const [loginForm, setLoginForm] = useState({ username: "", password: "" });
  const [loginError, setLoginError] = useState<string>("");

  const [form, setForm] = useState<Siswa>({ nisn: "", nama: "", kelas: "" });
  const [dataSiswa, setDataSiswa] = useState<Siswa[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [submitting, setSubmitting] = useState<boolean>(false);
  const [pesan, setPesan] = useState<string>("");
  const [editIndex, setEditIndex] = useState<number | null>(null);
  const [nisnLama, setNisnLama] = useState<string>("");
  const [importing, setImporting] = useState<boolean>(false);
  const [deletingAll, setDeletingAll] = useState<boolean>(false);
  const [selected, setSelected] = useState<string[]>([]);
  const [deletingSelected, setDeletingSelected] = useState<boolean>(false);
  const [filterNama, setFilterNama] = useState<string>("");
  const [filterKelas, setFilterKelas] = useState<string>("");
  const [downloadingTemplate, setDownloadingTemplate] =
    useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isLoggedIn) {
      ambilData();
    }
  }, [isLoggedIn]);

  const ambilData = async () => {
    setLoading(true);
    try {
      const res = await fetch(`${API_URL}?action=getData`);
      const json = await res.json();
      if (json.status === "success") {
        setDataSiswa(json.data);
        setSelected([]);
      } else {
        setPesan("Gagal mengambil data: " + json.message);
      }
    } catch (err) {
      setPesan("Terjadi kesalahan saat mengambil data.");
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setForm((prev) => ({ ...prev, [name]: value }));
  };

  const resetForm = () => {
    setForm({ nisn: "", nama: "", kelas: "" });
    setEditIndex(null);
    setNisnLama("");
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!form.nisn.trim() || !form.nama.trim() || !form.kelas.trim()) {
      setPesan("Semua kolom wajib diisi.");
      return;
    }

    setSubmitting(true);
    setPesan("");

    try {
      const action = editIndex !== null ? "updateData" : "addData";
      const body = new URLSearchParams();
      body.append("action", action);
      body.append("nisn", form.nisn);
      body.append("nama", form.nama);
      body.append("kelas", form.kelas);
      if (editIndex !== null) {
        body.append("nisnLama", nisnLama);
      }

      const res = await fetch(API_URL, {
        method: "POST",
        body,
      });
      const json = await res.json();

      if (json.status === "success") {
        setPesan(
          editIndex !== null
            ? "Data berhasil diperbarui."
            : "Data berhasil ditambahkan."
        );
        resetForm();
        ambilData();
      } else {
        setPesan("Gagal menyimpan data: " + json.message);
      }
    } catch (err) {
      setPesan("Terjadi kesalahan saat menyimpan data.");
      console.error(err);
    } finally {
      setSubmitting(false);
    }
  };

  const handleEdit = (index: number) => {
    const item = dataSiswa[index];
    setForm({ nisn: item.nisn, nama: item.nama, kelas: item.kelas });
    setEditIndex(index);
    setNisnLama(item.nisn);
    setPesan("");
  };

  const handleDelete = async (nisn: string) => {
    if (!window.confirm("Hapus data siswa dengan NISN " + nisn + "?")) return;

    setSubmitting(true);
    try {
      const body = new URLSearchParams();
      body.append("action", "deleteData");
      body.append("nisn", nisn);

      const res = await fetch(API_URL, {
        method: "POST",
        body,
      });
      const json = await res.json();

      if (json.status === "success") {
        setPesan("Data berhasil dihapus.");
        ambilData();
      } else {
        setPesan("Gagal menghapus data: " + json.message);
      }
    } catch (err) {
      setPesan("Terjadi kesalahan saat menghapus data.");
      console.error(err);
    } finally {
      setSubmitting(false);
    }
  };

  const toggleSelect = (nisn: string) => {
    setSelected((prev) =>
      prev.indexOf(nisn) !== -1
        ? prev.filter((n) => n !== nisn)
        : [...prev, nisn]
    );
  };

  const toggleSelectAll = () => {
    if (selected.length === dataTampil.length) {
      setSelected([]);
    } else {
      setSelected(dataTampil.map((s) => s.nisn));
    }
  };

  const handleDeleteSelected = async () => {
    if (selected.length === 0) {
      setPesan("Belum ada data siswa yang diceklis.");
      return;
    }
    if (
      !window.confirm(
        `Hapus ${selected.length} data siswa yang diceklis? Tindakan ini tidak bisa dibatalkan.`
      )
    )
      return;

    setDeletingSelected(true);
    setPesan("");
    try {
      const body = new URLSearchParams();
      body.append("action", "deleteSelectedData");
      body.append("nisnList", JSON.stringify(selected));

      const res = await fetch(API_URL, {
        method: "POST",
        body,
      });
      const json = await res.json();

      if (json.status === "success") {
        setPesan(json.message);
        resetForm();
        ambilData();
      } else {
        setPesan("Gagal menghapus data terpilih: " + json.message);
      }
    } catch (err) {
      setPesan("Terjadi kesalahan saat menghapus data terpilih.");
      console.error(err);
    } finally {
      setDeletingSelected(false);
    }
  };

  const handleDeleteAll = async () => {
    if (dataSiswa.length === 0) {
      setPesan("Tidak ada data siswa untuk dihapus.");
      return;
    }
    if (
      !window.confirm(
        `Hapus SEMUA data siswa (${dataSiswa.length} data)? Tindakan ini tidak bisa dibatalkan.`
      )
    )
      return;
    if (!window.confirm("Apakah Anda benar-benar yakin?")) return;

    setDeletingAll(true);
    setPesan("");
    try {
      const body = new URLSearchParams();
      body.append("action", "deleteAllData");

      const res = await fetch(API_URL, {
        method: "POST",
        body,
      });
      const json = await res.json();

      if (json.status === "success") {
        setPesan(json.message);
        resetForm();
        ambilData();
      } else {
        setPesan("Gagal menghapus semua data: " + json.message);
      }
    } catch (err) {
      setPesan("Terjadi kesalahan saat menghapus semua data.");
      console.error(err);
    } finally {
      setDeletingAll(false);
    }
  };

  const handleLoginChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setLoginForm((prev) => ({ ...prev, [name]: value }));
  };

  const handleLoginSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (
      loginForm.username === CREDENTIALS.username &&
      loginForm.password === CREDENTIALS.password
    ) {
      sessionStorage.setItem("isLoggedIn", "true");
      setIsLoggedIn(true);
      setLoginError("");
    } else {
      setLoginError("Username atau password salah.");
    }
  };

  const handleKembali = () => {
    if (window.parent !== window) {
      // Dibuka lewat launcher (iframe): minta launcher kembali ke halaman sebelumnya
      window.parent.postMessage({ type: "SISWA_KEMBALI" }, "*");
    } else {
      // Dibuka langsung di browser: kembali ke halaman sebelumnya
      window.history.back();
    }
  };

  const handleLogout = () => {
    sessionStorage.removeItem("isLoggedIn");
    setIsLoggedIn(false);
    setLoginForm({ username: "", password: "" });

    // Kalau dibuka lewat launcher (iframe), kembali ke halaman awal launcher
    if (window.parent !== window) {
      window.parent.postMessage({ type: "PKBM_KEMBALI" }, "*");
    }
  };

  const handleDownloadTemplate = async () => {
    setDownloadingTemplate(true);
    setPesan("");
    try {
      const res = await fetch(`${API_URL}?action=downloadTemplate`);
      const json = await res.json();
      if (json.status === "success" && json.url) {
        // Buka URL export Excel di tab/window baru (aman untuk WebView Android)
        window.open(json.url, "_blank");
      } else {
        setPesan("Gagal membuat template: " + json.message);
      }
    } catch (err) {
      setPesan("Terjadi kesalahan saat membuat template.");
      console.error(err);
    } finally {
      setDownloadingTemplate(false);
    }
  };

  const handleImportClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setImporting(true);
    setPesan("");

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];

      // raw: false -> nilai dibaca sebagai string (NISN tidak berubah jadi number)
      const rows: any[] = XLSX.utils.sheet_to_json(sheet, {
        raw: false,
        defval: "",
      });

      if (rows.length === 0) {
        setPesan("File kosong atau tidak ada data yang bisa diimpor.");
        setImporting(false);
        return;
      }

      const dataSiswaBaru: Siswa[] = [];
      const barisInvalid: number[] = [];

      rows.forEach((row, idx) => {
        // Cocokkan nama kolom fleksibel (huruf besar/kecil, spasi)
        const keys = Object.keys(row);
        const nisnKey = keys.find((k) => k.trim().toLowerCase() === "nisn");
        const namaKey = keys.find((k) => k.trim().toLowerCase() === "nama");
        const kelasKey = keys.find((k) => k.trim().toLowerCase() === "kelas");

        const nisn = nisnKey ? String(row[nisnKey]).trim() : "";
        const nama = namaKey ? String(row[namaKey]).trim() : "";
        const kelas = kelasKey ? String(row[kelasKey]).trim() : "";

        if (!nisn || !nama || !kelas) {
          barisInvalid.push(idx + 2); // +2: baris ke-1 header, data mulai baris ke-2
          return;
        }

        dataSiswaBaru.push({ nisn, nama, kelas });
      });

      if (dataSiswaBaru.length === 0) {
        setPesan(
          "Tidak ada baris valid untuk diimpor. Pastikan kolom bernama nisn, nama, kelas terisi."
        );
        setImporting(false);
        if (fileInputRef.current) fileInputRef.current.value = "";
        return;
      }

      const body = new URLSearchParams();
      body.append("action", "importData");
      body.append("data", JSON.stringify(dataSiswaBaru));

      const res = await fetch(API_URL, {
        method: "POST",
        body,
      });
      const json = await res.json();

      if (json.status === "success") {
        let msg = `Import selesai: ${json.ditambahkan} data ditambahkan, ${json.dilewati} dilewati (duplikat NISN).`;
        if (barisInvalid.length > 0) {
          msg += ` ${
            barisInvalid.length
          } baris dilewati karena data tidak lengkap (baris: ${barisInvalid.join(
            ", "
          )}).`;
        }
        setPesan(msg);
        ambilData();
      } else {
        setPesan("Gagal mengimpor data: " + json.message);
      }
    } catch (err) {
      setPesan(
        "Terjadi kesalahan saat membaca file. Pastikan file berformat .xlsx atau .xls."
      );
      console.error(err);
    } finally {
      setImporting(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  const daftarKelas = Array.from(new Set(dataSiswa.map((s) => s.kelas))).sort(
    (a, b) => a.localeCompare(b, undefined, { numeric: true })
  );

  const dataTampil = dataSiswa.filter(
    (s) =>
      (filterNama.trim() === "" ||
        s.nama.toLowerCase().includes(filterNama.trim().toLowerCase())) &&
      (filterKelas === "" || s.kelas === filterKelas)
  );

  if (!isLoggedIn) {
    return (
      <div style={styles.loginContainer}>
        <form onSubmit={handleLoginSubmit} style={styles.loginForm}>
          <h1 style={styles.title}>Login</h1>
          <div style={styles.formGroup}>
            <label style={styles.label}>Username</label>
            <input
              type="text"
              name="username"
              value={loginForm.username}
              onChange={handleLoginChange}
              style={styles.input}
              placeholder="Masukkan username"
              autoFocus
            />
          </div>
          <div style={styles.formGroup}>
            <label style={styles.label}>Password</label>
            <input
              type="password"
              name="password"
              value={loginForm.password}
              onChange={handleLoginChange}
              style={styles.input}
              placeholder="Masukkan password"
            />
          </div>
          {loginError && <p style={styles.pesan}>{loginError}</p>}
          <button type="submit" style={styles.buttonPrimary}>
            Masuk
          </button>
          <button
            type="button"
            style={styles.buttonSecondary}
            onClick={handleKembali}
          >
            ← Kembali
          </button>
        </form>
      </div>
    );
  }

  return (
    <div style={styles.container}>
      <div style={styles.headerRow}>
        <h1 style={styles.title}>Input Data Siswa</h1>
        <button
          type="button"
          style={styles.buttonSecondary}
          onClick={handleLogout}
        >
          Logout
        </button>
      </div>

      <form onSubmit={handleSubmit} style={styles.form}>
        <div style={styles.formGroup}>
          <label style={styles.label}>NISN</label>
          <input
            type="text"
            name="nisn"
            value={form.nisn}
            onChange={handleChange}
            style={styles.input}
            placeholder="Masukkan NISN"
          />
        </div>

        <div style={styles.formGroup}>
          <label style={styles.label}>Nama</label>
          <input
            type="text"
            name="nama"
            value={form.nama}
            onChange={handleChange}
            style={styles.input}
            placeholder="Masukkan Nama"
          />
        </div>

        <div style={styles.formGroup}>
          <label style={styles.label}>Kelas</label>
          <input
            type="text"
            name="kelas"
            value={form.kelas}
            onChange={handleChange}
            style={styles.input}
            placeholder="Masukkan Kelas"
          />
        </div>

        <div style={styles.buttonRow}>
          <button
            type="submit"
            style={styles.buttonPrimary}
            disabled={submitting}
          >
            {submitting
              ? "Menyimpan..."
              : editIndex !== null
              ? "Update Data"
              : "Simpan Data"}
          </button>
          {editIndex !== null && (
            <button
              type="button"
              style={styles.buttonSecondary}
              onClick={resetForm}
            >
              Batal
            </button>
          )}
        </div>
      </form>

      {pesan && <p style={styles.pesan}>{pesan}</p>}

      <div style={styles.importBox}>
        <h2 style={styles.subtitle}>Import Data Excel</h2>
        <p style={styles.importText}>
          Unduh template, isi data siswa (kolom: nisn, nama, kelas), lalu unggah
          kembali file-nya.
        </p>
        <div style={styles.buttonRow}>
          <button
            type="button"
            style={styles.buttonSecondary}
            onClick={handleDownloadTemplate}
            disabled={downloadingTemplate}
          >
            {downloadingTemplate
              ? "Membuat template..."
              : "Download Template Excel"}
          </button>
          <button
            type="button"
            style={styles.buttonPrimary}
            onClick={handleImportClick}
            disabled={importing}
          >
            {importing ? "Mengimpor..." : "Import Data Excel"}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            style={{ display: "none" }}
            onChange={handleFileImport}
          />
        </div>
      </div>

      <div style={styles.buttonRow}>
        <button
          type="button"
          style={styles.buttonDeleteAll}
          onClick={handleDeleteAll}
          disabled={deletingAll}
        >
          {deletingAll ? "Menghapus..." : "Hapus Semua Data Siswa"}
        </button>
        <button
          type="button"
          style={styles.buttonDeleteAll}
          onClick={handleDeleteSelected}
          disabled={deletingSelected || selected.length === 0}
        >
          {deletingSelected
            ? "Menghapus..."
            : `Hapus Terpilih (${selected.length})`}
        </button>
      </div>

      <h2 style={styles.subtitle}>Daftar Siswa</h2>

      <div style={styles.filterRow}>
        <input
          type="text"
          list="daftar-nama"
          value={filterNama}
          onChange={(e) => {
            setFilterNama(e.target.value);
            setSelected([]);
          }}
          style={{ ...styles.input, flex: 1, minWidth: 180 }}
          placeholder="Cari / pilih nama siswa"
        />
        <datalist id="daftar-nama">
          {dataSiswa.map((s, i) => (
            <option key={s.nisn + i} value={s.nama} />
          ))}
        </datalist>

        <select
          value={filterKelas}
          onChange={(e) => {
            setFilterKelas(e.target.value);
            setSelected([]);
          }}
          style={styles.input}
        >
          <option value="">Semua Kelas</option>
          {daftarKelas.map((k) => (
            <option key={k} value={k}>
              Kelas {k}
            </option>
          ))}
        </select>

        {(filterNama !== "" || filterKelas !== "") && (
          <button
            type="button"
            style={styles.buttonSecondary}
            onClick={() => {
              setFilterNama("");
              setFilterKelas("");
              setSelected([]);
            }}
          >
            Reset Filter
          </button>
        )}
      </div>

      {loading ? (
        <p>Memuat data...</p>
      ) : dataSiswa.length === 0 ? (
        <p>Belum ada data siswa.</p>
      ) : dataTampil.length === 0 ? (
        <p>Tidak ada data yang cocok dengan filter.</p>
      ) : (
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>
                <input
                  type="checkbox"
                  checked={
                    dataTampil.length > 0 &&
                    selected.length === dataTampil.length
                  }
                  onChange={toggleSelectAll}
                />
              </th>
              <th style={styles.th}>No</th>
              <th style={styles.th}>NISN</th>
              <th style={styles.th}>Nama</th>
              <th style={styles.th}>Kelas</th>
              <th style={styles.th}>Aksi</th>
            </tr>
          </thead>
          <tbody>
            {dataTampil.map((item, index) => (
              <tr key={item.nisn + index}>
                <td style={styles.td}>
                  <input
                    type="checkbox"
                    checked={selected.indexOf(item.nisn) !== -1}
                    onChange={() => toggleSelect(item.nisn)}
                  />
                </td>
                <td style={styles.td}>{index + 1}</td>
                <td style={styles.td}>{item.nisn}</td>
                <td style={styles.td}>{item.nama}</td>
                <td style={styles.td}>{item.kelas}</td>
                <td style={styles.td}>
                  <button
                    style={styles.buttonEdit}
                    onClick={() => handleEdit(dataSiswa.indexOf(item))}
                  >
                    Edit
                  </button>
                  <button
                    style={styles.buttonDelete}
                    onClick={() => handleDelete(item.nisn)}
                  >
                    Hapus
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

const styles: { [key: string]: React.CSSProperties } = {
  headerRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
  },
  loginContainer: {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background: "#f3f4f6",
    fontFamily: "Arial, sans-serif",
  },
  loginForm: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    background: "#fff",
    padding: "28px",
    borderRadius: "10px",
    border: "1px solid #e5e7eb",
    width: "300px",
    boxShadow: "0 4px 12px rgba(0,0,0,0.06)",
  },
  container: {
    maxWidth: 700,
    margin: "0 auto",
    padding: "24px",
    fontFamily: "Arial, sans-serif",
  },
  title: {
    fontSize: "22px",
    marginBottom: "16px",
    color: "#1f2937",
  },
  subtitle: {
    fontSize: "18px",
    marginTop: "32px",
    marginBottom: "12px",
    color: "#1f2937",
  },
  form: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    background: "#f9fafb",
    padding: "16px",
    borderRadius: "8px",
    border: "1px solid #e5e7eb",
  },
  formGroup: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  label: {
    fontSize: "14px",
    fontWeight: 600,
    color: "#374151",
  },
  input: {
    padding: "8px 10px",
    border: "1px solid #d1d5db",
    borderRadius: "6px",
    fontSize: "14px",
  },
  buttonRow: {
    display: "flex",
    gap: "8px",
    marginTop: "4px",
  },
  buttonPrimary: {
    padding: "10px 16px",
    background: "#2563eb",
    color: "#fff",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    fontSize: "14px",
  },
  buttonSecondary: {
    padding: "10px 16px",
    background: "#e5e7eb",
    color: "#111827",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    fontSize: "14px",
  },
  pesan: {
    marginTop: "12px",
    fontSize: "14px",
    color: "#b45309",
  },
  importBox: {
    marginTop: "24px",
    padding: "16px",
    background: "#f0fdf4",
    border: "1px solid #bbf7d0",
    borderRadius: "8px",
  },
  importText: {
    fontSize: "13px",
    color: "#4b5563",
    marginBottom: "10px",
  },
  table: {
    width: "100%",
    borderCollapse: "collapse",
  },
  th: {
    textAlign: "left",
    padding: "8px",
    borderBottom: "2px solid #e5e7eb",
    fontSize: "14px",
    color: "#374151",
  },
  td: {
    padding: "8px",
    borderBottom: "1px solid #f3f4f6",
    fontSize: "14px",
  },
  buttonEdit: {
    marginRight: "6px",
    padding: "4px 10px",
    background: "#facc15",
    border: "none",
    borderRadius: "4px",
    cursor: "pointer",
    fontSize: "12px",
  },
  filterRow: {
    display: "flex",
    gap: "8px",
    marginBottom: "12px",
    flexWrap: "wrap",
  },
  buttonDeleteAll: {
    padding: "10px 16px",
    background: "#b91c1c",
    color: "#fff",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    fontSize: "14px",
    marginTop: "16px",
  },
  buttonDelete: {
    padding: "4px 10px",
    background: "#ef4444",
    color: "#fff",
    border: "none",
    borderRadius: "4px",
    cursor: "pointer",
    fontSize: "12px",
  },
};
