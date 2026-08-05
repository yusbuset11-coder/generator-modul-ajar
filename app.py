from io import BytesIO
import json
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
import google.generativeai as genai
import streamlit as st

st.set_page_config(
    page_title="GENERATOR: MODUL AJAR PEMBELAJARAN MENDALAM",
    page_icon="📚",
    layout="wide",
)


def check_password():
  """Mengembalikan True jika pengguna memasukkan password yang benar."""

  def password_entered():
    if st.session_state["password"] == "modulpm230371":
      st.session_state["password_correct"] = True
      del st.session_state["password"]
    else:
      st.session_state["password_correct"] = False

  if "password_correct" not in st.session_state:
    st.text_input(
        "🔑 Masukkan Password Akses Aplikasi:",
        type="password",
        on_change=password_entered,
        key="password",
    )
    return False
  elif not st.session_state["password_correct"]:
    st.text_input(
        "🔑 Masukkan Password Akses Aplikasi:",
        type="password",
        on_change=password_entered,
        key="password",
    )
    st.error("😕 Maaf, password yang Anda masukkan salah.")
    return False
  else:
    return True


if not check_password():
  st.stop()
# ===================================

# Custom CSS untuk tampilan UI yang modern, keren, dan profesional
st.markdown(
    """
    <style>
    /* Styling Global & Background */
    .stApp {
        background-color: #0e1117;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Header Card Banner Modern */
    .header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 16px 20px;
        border-radius: 12px;
        border: 1px solid #334155;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    }
    .header-title {
        color: #f8fafc;
        font-size: 15px;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.3px;
    }
    
    /* Animasi Berkedip (Blinking) untuk Subtitle agar kontras & menarik */
    @keyframes blink-animation {
        0% { opacity: 1; color: #facc15; } /* Kuning Cerah */
        50% { opacity: 0.35; color: #38bdf8; } /* Biru Cyan Cerah */
        100% { opacity: 1; color: #facc15; }
    }
    
    .header-subtitle {
        font-size: 11.5px;
        margin-top: 6px;
        margin-bottom: 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        animation: blink-animation 1.6s infinite ease-in-out;
        font-weight: 600;
    }
    
    /* Tombol Utama (Generate Button) */
    .stButton > button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%);
        color: white;
        border: none;
        padding: 0.65rem 1rem;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.35);
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #4338ca 0%, #2563eb 100%);
        box-shadow: 0 6px 18px rgba(59, 130, 246, 0.5);
        transform: translateY(-2px);
    }
    
    /* Tombol Unduh (Download Button) */
    .stDownloadButton > button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        background: linear-gradient(135deg, #059669 0%, #10b981 100%);
        color: white;
        border: none;
        padding: 0.65rem 1rem;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.35);
        transition: all 0.3s ease;
    }
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #047857 0%, #059669 100%);
        box-shadow: 0 6px 18px rgba(16, 185, 129, 0.5);
        transform: translateY(-2px);
    }
    
    /* Styling Sidebar */
    [data-testid="stSidebar"] {
        background-color: #111827;
        border-right: 1px solid #1f2937;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Tampilan Header Modern dalam Card
st.markdown(
    """
    <div class="header-card">
        <h2 class="header-title">
            <marquee behavior="scroll" direction="left" scrollamount="7" style="color: #38bdf8; text-shadow: 0 0 12px rgba(56, 189, 248, 0.5);">📚 GENERATOR: MODUL AJAR PEMBELAJARAN MENDALAM</marquee>
        </h2>
        <div class="header-subtitle">
            <b>Pengembang:</b> Yustinus Budi Setyanta - PS Cabdin Bangkalan &nbsp;|&nbsp; 
            <em>Aplikasi Otomatisasi Perancangan Pembelajaran Deep Learning</em>
        </div>
    </div>
""",
    unsafe_allow_html=True,
)

# Input Pengguna di Sidebar
with st.sidebar:
  st.header("⚙️ Parameter Pembelajaran")
  api_key = st.text_input("Masukkan Google Gemini API Key", type="password")

  jenjang_pendidikan = st.selectbox(
      "Pilih Jenjang Pendidikan",
      ["SD / MI", "SMP / MTs", "SMA / MA", "SMK / MAK"],
  )

  if jenjang_pendidikan == "SD / MI":
    default_mapel = "Tematik / Kelas"
    fase_options = [
        "Fase A / Kelas 1 SD",
        "Fase A / Kelas 2 SD",
        "Fase B / Kelas 3 SD",
        "Fase B / Kelas 4 SD",
        "Fase C / Kelas 5 SD",
        "Fase C / Kelas 6 SD",
    ]
  elif jenjang_pendidikan == "SMP / MTs":
    default_mapel = "Matematika / IPA / IPS"
    fase_options = [
        "Fase D / Kelas 7 SMP",
        "Fase D / Kelas 8 SMP",
        "Fase D / Kelas 9 SMP",
    ]
  elif jenjang_pendidikan == "SMA / MA":
    default_mapel = "Bahasa Indonesia / Matematika"
    fase_options = [
        "Fase E / Kelas X SMA",
        "Fase F / Kelas XI SMA",
        "Fase F / Kelas XII SMA",
    ]
  else:
    default_mapel = (
        "Dasar-dasar Teknik Otomotif / Produk Kreatif dan Kewirausahaan"
    )
    fase_options = [
        "Fase E / Kelas X SMK (Program Dasar Keahlian)",
        "Fase F / Kelas XI SMK (Konsentrasi Keahlian)",
        "Fase F / Kelas XII SMK (Konsentrasi Keahlian)",
    ]

  mata_pelajaran = st.text_input(
      "Mata Pelajaran / Program Kejuruan", default_mapel
  )
  fase_kelas = st.selectbox("Fase / Kelas", fase_options)

  topik = st.text_input(
      "Topik / Materi Pokok / Elemen",
      (
          "Contoh: Pemeliharaan Sistem Rem Kendaraan Ringan"
          if jenjang_pendidikan == "SMK / MAK"
          else "Contoh: Menyimak Teks Laporan Observasi Secara Kritis"
      ),
  )
  alokasi_waktu = st.text_input("Alokasi Waktu", "2 Pertemuan (4 x 45 Menit)")

  st.markdown("---")
  st.header("🏫 Identitas Satuan Pendidikan")
  nama_sekolah = st.text_input("Nama Sekolah", "SD Negeri Balongsari 2 Surabaya")
  semester = st.selectbox("Semester", ["Ganjil", "Genap"])
  tahun_pelajaran = st.text_input("Tahun Pelajaran", "2026/2027")

  st.markdown("---")
  st.header("✍️ Identitas Pengesahan Dokumen")
  nama_kota = st.text_input("Nama Kota", "Bangkalan")
  tanggal_pembuatan = st.text_input(
      "Tanggal / Bulan / Tahun", "5 Agustus 2026"
  )
  nama_penulis = st.text_input(
      "Nama Penulis Modul", "Yustinus Budi Setyanta, S.Pd., M.Pd."
  )
  nip_penulis = st.text_input("NIP Penulis", "196908302005011003")


# Fungsi pembantu warna latar belakang sel tabel Word
def set_cell_background(cell, fill_color):
  shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
  cell._tc.get_or_add_tcPr().append(shading_elm)


# Fungsi pembentuk dokumen Word lengkap berbasis Tabel Matriks dengan Estetika Eye-Catching & Format Rata Kiri
def generate_docx(
    data_ai,
    nama_sekolah,
    semester,
    tahun_pelajaran,
    mata_pelajaran,
    fase_kelas,
    topik,
    alokasi_waktu,
    nama_penulis,
    nama_kota,
    tanggal_pembuatan,
    nip_penulis,
):
  doc = docx.Document()

  # Pengaturan margin halaman
  for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

  style = doc.styles["Normal"]
  font = style.font
  font.name = "Arial"
  font.size = Pt(10)
  font.color.rgb = RGBColor(51, 51, 51)

  # Judul Utama Dokumen
  p_title = doc.add_paragraph()
  p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_title.paragraph_format.space_before = Pt(0)
  p_title.paragraph_format.space_after = Pt(12)
  run_title = p_title.add_run("MODUL AJAR PEMBELAJARAN MENDALAM")
  run_title.font.name = "Arial"
  run_title.font.size = Pt(13)
  run_title.font.bold = True

  # Helper untuk menambah tabel section dengan header kuning, kolom kiri berwarna (eye-catching), dan format teks sesuai permintaan
  def add_section_table(title_text, rows_data):
    table = doc.add_table(rows=len(rows_data) + 1, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header Row
    hdr_cells = table.rows[0].cells
    hdr_cells[0].merge(hdr_cells[1])
    hdr_cells[0].text = title_text
    set_cell_background(hdr_cells[0], "FFE599")
    for p in hdr_cells[0].paragraphs:
      p.alignment = WD_ALIGN_PARAGRAPH.LEFT
      p.paragraph_format.space_before = Pt(4)
      p.paragraph_format.space_after = Pt(4)
      for run in p.runs:
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(51, 51, 51)

    # Content Rows
    for idx, (label, val) in enumerate(rows_data):
      row_cells = table.rows[idx + 1].cells
      row_cells[0].text = label
      row_cells[1].text = str(val)
      row_cells[0].width = Inches(2.2)
      row_cells[1].width = Inches(4.3)

      # Memberikan warna latar belakang lembut pada kolom kiri agar Eye-Catching
      set_cell_background(row_cells[0], "F2F5F9")

      # Kolom Kiri (Label) -> Bold & Rata Kiri
      for p in row_cells[0].paragraphs:
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
          run.font.size = Pt(10)
          run.font.bold = True
          run.font.color.rgb = RGBColor(51, 51, 51)

      # Kolom Kanan (Isi)
      for p in row_cells[1].paragraphs:
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        text_p = p.text.strip()

        # KHUSUS TABEL IDENTITAS: Kolom kanan TIDAK BOLEH BOLD (Normal)
        if title_text == "IDENTITAS":
          for run in p.runs:
            run.font.size = Pt(10)
            run.font.bold = False
            run.font.color.rgb = RGBColor(51, 51, 51)
        else:
          # TABEL KONTEN LAINNYA: Sub-bagian/subjudul dicetak Bold, isi teks normal
          is_subheader = (
              text_p.startswith("PERTEMUAN")
              or text_p.startswith("Kegiatan")
              or text_p.startswith("Asesmen")
              or text_p.startswith("1.")
              or text_p.startswith("2.")
              or text_p.startswith("3.")
              or text_p.startswith("4.")
              or text_p.startswith("a.")
              or text_p.startswith("b.")
              or text_p.startswith("c.")
              or text_p.startswith("d.")
              or text_p.startswith("☐")
              or text_p.endswith(":")
              or (len(text_p) < 85 and not text_p.endswith("."))
          )
          for run in p.runs:
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(51, 51, 51)
            if is_subheader:
              run.font.bold = True
            else:
              run.font.bold = False

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

  # 1. Tabel Identitas
  tabel_identitas = [
      ("Penyusun", nama_penulis),
      ("Sekolah", nama_sekolah),
      ("Tahun Pelajaran", tahun_pelajaran),
      ("Semester", semester),
      ("Mata Pelajaran", mata_pelajaran),
      ("Kelas / Fase Capaian", fase_kelas),
      ("Topik / Elemen", topik),
      ("Alokasi Waktu", alokasi_waktu),
  ]
  add_section_table("IDENTITAS", tabel_identitas)

  # 2. Tabel A. Identifikasi
  tabel_identifikasi = [
      (
          "Peserta Didik",
          data_ai.get(
              "peserta_didik",
              "Analisis kesiapan awal dan karakteristik peserta didik.",
          ),
      ),
      (
          "Materi Pelajaran",
          data_ai.get("materi_esensial", "Analisis materi esensial."),
      ),
      (
          "Dimensi Profil Lulusan",
          data_ai.get(
              "dimensi_profil_lulusan",
              "Dimensi Profil Lulusan yang dipilih:\n"
              "☐ DPL3 Penalaran Kritis\n☐ DPL5 Kolaborasi",
          ),
      ),
  ]
  add_section_table("A. Identifikasi Kompetensi & Karakteristik", tabel_identifikasi)

  # 3. Tabel B. Desain Pembelajaran
  tabel_desain = [
      (
          "Desain Pembelajaran Mendalam",
          data_ai.get(
              "desain_pembelajaran",
              "Praktik pedagogik, kemitraan, lingkungan, dan digital.",
          ),
      ),
  ]
  add_section_table("B. Desain Pembelajaran (Deep Learning)", tabel_desain)

  # 4. Tabel C. Pengalaman Belajar
  tabel_pengalaman = [
      (
          "Pertemuan 1",
          data_ai.get(
              "pengalaman_belajar_1",
              "Kegiatan Awal, Inti, dan Penutup Pertemuan 1.",
          ),
      ),
      (
          "Pertemuan 2",
          data_ai.get(
              "pengalaman_belajar_2",
              "Kegiatan Awal, Inti, dan Penutup Pertemuan 2.",
          ),
      ),
  ]
  add_section_table("C. Pengalaman Belajar (Berkesan & Bermakna)", tabel_pengalaman)

  # 5. Tabel D. Asesmen & Rubrik
  tabel_asesmen = [
      (
          "Asesmen Pembelajaran & Rubrik",
          data_ai.get(
              "asesmen_rubrik",
              "Asesmen diagnostik, formatif, sumatif, dan rubrik unjuk kerja.",
          ),
      ),
  ]
  add_section_table("D. Asesmen Pembelajaran Vokasi / Akademik", tabel_asesmen)

  # Tanda Tangan / Pengesahan di Bagian Bawah
  p_sign = doc.add_paragraph()
  p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  p_sign.paragraph_format.space_before = Pt(14)
  p_sign.paragraph_format.space_after = Pt(4)
  p_sign.add_run(f"{nama_kota}, {tanggal_pembuatan}\nDisusun,\n\n\n")
  run_name = p_sign.add_run(f"{nama_penulis}")
  run_name.font.bold = True
  p_sign.add_run(f"\nNIP. {nip_penulis}")

  bio = BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


if st.button("🚀 Buat Modul Ajar Lengkap Format Tabel"):
  if not api_key:
    st.error("Mohon masukkan Google Gemini API Key terlebih dahulu.")
  elif not topik:
    st.warning("Mohon isi topik pembelajaran.")
  else:
    with st.spinner(
        "Gemini sedang menyusun Modul Ajar Komprehensif Format Tabel..."
    ):
      genai.configure(api_key=api_key)
      model = genai.GenerativeModel("gemini-3.5-flash")

      prompt = f"""
            Bertindaklah sebagai pakar kurikulum dan praktisi pendidikan profesional. Buatkan konten Modul Ajar Berbasis Pembelajaran Mendalam (Deep Learning) yang **SANGAT LENGKAP, DETAIL, DAN KOMPREHENSIF** untuk:
            - Jenjang: {jenjang_pendidikan} ({fase_kelas})
            - Mata Pelajaran: {mata_pelajaran}
            - Topik / Materi Pokok: {topik}
            - Alokasi Waktu: {alokasi_waktu}

            Catatan Penting:
            - Gunakan istilah "Dimensi Profil Lulusan (DPL)" secara konsisten (jangan gunakan istilah lama Profil Pelajar Pancasila).
            - Pilih 2 sampai 4 Dimensi Profil Lulusan (DPL) yang paling relevan dan kontekstual dengan topik "{topik}" dari daftar 8 DPL berikut (DPL1: Keimanan & Ketakwaan, DPL2: Kewargaan, DPL3: Penalaran Kritis, DPL4: Kreativitas, DPL5: Kolaborasi, DPL6: Kemandirian, DPL7: Kesehatan, DPL8: Komunikasi). Berikan format kotak centang (☐) hanya pada DPL yang dipilih tersebut.
            - Pastikan setiap sub-bagian, subjudul, atau poin penomoran (seperti "Analisis Karakteristik...", "1. Kompetensi Awal & Prasyarat:", dll.) diletakkan pada baris baru dan diakhiri dengan tanda titik dua (:) agar aplikasi dapat memformatnya menjadi tebal (bold) dengan rapi.

            Berikan output HANYA dalam format JSON valid yang memuat kunci-kunci berikut:
            {{
              "peserta_didik": "Uraian sangat detail mengenai kompetensi awal, prasyarat, gaya belajar, dan kesiapan peserta didik. Buat subjudul terstruktur dengan diakhiri tanda titik dua (:).",
              "materi_esensial": "Uraian sangat detail mengenai materi esensial, konsep utama, dan relevansi dunia nyata/industri.",
              "dimensi_profil_lulusan": "Sebutkan pilihan DPL yang relevan saja (misal 2 atau 3 DPL) yang paling sesuai dengan topik ini, lengkap dengan kotak centang ☐ pada masing-masing pilihan terpilih serta deskripsi singkat penerapannya.",
              "desain_pembelajaran": "Uraian rinci desain pembelajaran mendalam mencakup praktik pedagogik, kemitraan, lingkungan belajar, dan integrasi digital.",
              "pengalaman_belajar_1": "Uraian rinci Pertemuan 1 yang mencakup Kegiatan Awal, Kegiatan Inti, dan Kegiatan Penutup/Refleksi.",
              "pengalaman_belajar_2": "Uraian rinci Pertemuan 2 yang mencakup Kegiatan Awal, Kegiatan Inti, dan Kegiatan Penutup/Refleksi.",
              "asesmen_rubrik": "Uraian rinci asesmen diagnostik, formatif, sumatif, beserta matriks rubrik penilaian unjuk kerja yang operasional dan lengkap."
            }}
            """

      response = model.generate_content(prompt)
      text_resp = response.text.strip()

      if text_resp.startswith("```json"):
        text_resp = text_resp[7:]
      if text_resp.startswith("```"):
        text_resp = text_resp[3:]
      if text_resp.endswith("```"):
        text_resp = text_resp[:-3]
      text_resp = text_resp.strip()

      try:
        data_ai = json.loads(text_resp)
      except Exception:
        data_ai = {
            "peserta_didik": (
                "Peserta didik telah memahami konsep prasyarat dan siap"
                " mengikuti pembelajaran."
            ),
            "materi_esensial": (
                f"Materi esensial terkait {topik} dan relevansinya."
            ),
            "dimensi_profil_lulusan": (
                "Dimensi Profil Lulusan yang relevan:\n☐ DPL3 Penalaran"
                " Kritis\n☐ DPL5 Kolaborasi"
            ),
            "desain_pembelajaran": (
                "Pendekatan pembelajaran mendalam (Deep Learning) terintegrasi."
            ),
            "pengalaman_belajar_1": (
                "Kegiatan awal, inti, dan penutup Pertemuan 1."
            ),
            "pengalaman_belajar_2": (
                "Kegiatan awal, inti, dan penutup Pertemuan 2."
            ),
            "asesmen_rubrik": (
                "Asesmen diagnostik, formatif, sumatif, dan rubrik penilaian."
            ),
        }

      st.success(
          "Modul Ajar Komprehensif Berbasis Format Tabel Berhasil Disusun!"
      )
      st.info(
          "Dokumen Word (.docx) siap diunduh dengan kolom kiri berwarna,"
          " identitas kanan normal, serta sub-bagian tercetak tebal."
      )

      docx_file = generate_docx(
          data_ai,
          nama_sekolah,
          semester,
          tahun_pelajaran,
          mata_pelajaran,
          fase_kelas,
          topik,
          alokasi_waktu,
          nama_penulis,
          nama_kota,
          tanggal_pembuatan,
          nip_penulis,
      )

      st.download_button(
          label=(
              "📥 Unduh Modul Ajar Lengkap (Format Tabel Matriks .docx)"
          ),
          data=docx_file,
          file_name=f"Modul_Ajar_Lengkap_{topik.replace(' ', '_')}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )