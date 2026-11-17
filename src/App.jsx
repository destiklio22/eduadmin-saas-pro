import React, { useState, useEffect, useRef } from 'react';
import {
  BookOpen,
  LayoutDashboard,
  FileText,
  Calendar,
  BookMarked,
  Users,
  CheckCircle2,
  Award,
  Settings,
  Plus,
  FileDown,
  Trash2,
  Save,
  RotateCcw,
  ChevronRight,
  GraduationCap,
  Clock,
  Check,
  X,
  Upload,
  FileSpreadsheet,
  ArrowRight,
  ArrowLeft,
  LogOut,
  ExternalLink
} from 'lucide-react';

// GANTI TAUTAN INI DENGAN LINK PRODUK LYNK.ID ANDA
const LYNK_ID_URL = 'https://lynk.id';

// --- FUNGSI EKSPOR DOKUMEN MICROSOFT WORD (.DOC) ---
const downloadWordDocument = (filename, title, profile, activeClass, bodyHtml) => {
  const fullHtml = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' 
          xmlns:w='urn:schemas-microsoft-com:office:word' 
          xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <meta charset='utf-8'>
      <title>${title}</title>
      <style>
        @page { size: A4; margin: 2cm; }
        body { font-family: 'Times New Roman', Times, serif; font-size: 11pt; line-height: 1.5; color: #000; }
        .header-title { text-align: center; font-size: 14pt; font-weight: bold; text-transform: uppercase; margin-bottom: 4px; }
        .header-sub { text-align: center; font-size: 12pt; font-weight: bold; margin-bottom: 16px; border-bottom: 2px solid #000; padding-bottom: 8px; }
        .meta-table { width: 100%; border-collapse: collapse; margin-bottom: 16px; font-size: 11pt; }
        .meta-table td { padding: 3px 4px; vertical-align: top; }
        .data-table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 10.5pt; }
        .data-table th, .data-table td { border: 1px solid #000; padding: 6px 8px; text-align: left; }
        .data-table th { background-color: #f0f4f8; font-weight: bold; text-align: center; }
        .box-cp { border: 1px solid #000; padding: 10px; background-color: #fafafa; margin-bottom: 14px; text-align: justify; }
        .ttd-table { width: 100%; border-collapse: collapse; margin-top: 36px; font-size: 11pt; }
        .ttd-table td { width: 50%; text-align: center; vertical-align: top; }
      </style>
    </head>
    <body>
      <div class="header-title">${title}</div>
      <div class="header-sub">${profile.sekolah.toUpperCase()}</div>
      
      <table class="meta-table">
        <tr>
          <td style="width: 22%;"><strong>Mata Pelajaran</strong></td>
          <td style="width: 28%;">: ${activeClass.mapel}</td>
          <td style="width: 22%;"><strong>Tahun Ajaran</strong></td>
          <td style="width: 28%;">: ${profile.tahunAjaran}</td>
        </tr>
        <tr>
          <td><strong>Kelas / Rombel</strong></td>
          <td>: ${activeClass.nama}</td>
          <td><strong>Semester</strong></td>
          <td>: ${profile.semester}</td>
        </tr>
        <tr>
          <td><strong>Fase Kurikulum</strong></td>
          <td>: ${activeClass.fase}</td>
          <td><strong>Guru Pengampu</strong></td>
          <td>: ${profile.nama}</td>
        </tr>
      </table>

      ${bodyHtml}

      <table class="ttd-table">
        <tr>
          <td>
            Mengetahui,<br/>
            Kepala ${profile.sekolah}<br/><br/><br/><br/>
            <strong><u>...................................................</u></strong><br/>
            NIP. ........................................
          </td>
          <td>
            ${new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}<br/>
            Guru Mata Pelajaran<br/><br/><br/><br/>
            <strong><u>${profile.nama}</u></strong><br/>
            NIP. ${profile.nip}
          </td>
        </tr>
      </table>
    </body>
    </html>
  `;

  const blob = new Blob(['\ufeff', fullHtml], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${filename}.doc`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

// --- BANK DATA KURIKULUM MERDEKA LENGKAP (FASE A - F & MULTI MAPEL) ---
const SUBJECT_BANK = {
  'Bahasa Indonesia': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik memiliki kemampuan berbahasa untuk berkomunikasi dan bernalar sesuai dengan tujuan, konteks sosial, akademis, dan dunia kerja. Peserta didik mampu memahami, mengolah, menginterpretasi, dan mengevaluasi informasi dari berbagai tipe teks fiksi dan informasional serta menyajikannya secara lisan dan tulis dengan kaidah kebahasaan yang tepat.',
    tps: [
      { kode: 'TP 2.1', judul: 'Menganalisis Informasi, Ide Pokok, dan Struktur Teks Eksplanasi dan Berita', jp: 12, status: 'SELESAI', elemen: 'Membaca dan Memirsa' },
      { kode: 'TP 2.2', judul: 'Menyimak dan Mengevaluasi Gagasan dari Teks Pidato / Diskusi Formal', jp: 10, status: 'SELESAI', elemen: 'Menyimak' },
      { kode: 'TP 2.3', judul: 'Menulis Teks Argumentasi dan Teks Tanggapan Kritis Berdasarkan Data Faktual', jp: 14, status: 'DRAFT', elemen: 'Menulis' },
      { kode: 'TP 2.4', judul: 'Mempresentasikan Karya Tulis dan Mengapresiasi Puisi / Cerpen Secara Lisan', jp: 10, status: 'DRAFT', elemen: 'Berbicara dan Mempresentasikan' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Mengkritisi Informasi pada Teks Berita & Eksplanasi', tpKode: 'TP 2.1', pertemuan: '4 Pertemuan (12 JP)', model: 'Discovery Learning', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Menulis Teks Argumentasi Berbasis Fakta', tpKode: 'TP 2.3', pertemuan: '4 Pertemuan (14 JP)', model: 'Project-Based Learning (PjBL)', status: 'Siap Unduh' }
    ]
  },
  'Matematika': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik menunjukkan pemahaman dan intuisi bilangan, melakukan operasi aljabar, fungsi, persamaan dan pertidaksamaan linear, menerapkan teorema geometri dan pengukuran bangun ruang, serta menganalisis data statistik dan peluang dalam pemecahan masalah kehidupan nyata.',
    tps: [
      { kode: 'TP 2.1', judul: 'Menyelesaikan Sistem Persamaan Linear dan Pemodelan Masalah Sehari-hari', jp: 14, status: 'SELESAI', elemen: 'Aljabar' },
      { kode: 'TP 2.2', judul: 'Menerapkan Teorema Pythagoras dan Hubungan Antar Sudut pada Geometri', jp: 12, status: 'SELESAI', elemen: 'Geometri' },
      { kode: 'TP 2.3', judul: 'Menghitung Luas Permukaan dan Volume Bangun Ruang Sisi Datar & Lengkung', jp: 14, status: 'DRAFT', elemen: 'Pengukuran' },
      { kode: 'TP 2.4', judul: 'Menganalisis Ukuran Pemusatan Data (Mean, Median, Modus) dan Peluang Empirik', jp: 12, status: 'DRAFT', elemen: 'Analisis Data dan Peluang' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Pemodelan Sistem Persamaan Linear', tpKode: 'TP 2.1', pertemuan: '4 Pertemuan (14 JP)', model: 'Problem-Based Learning (PBL)', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Eksplorasi Bangun Ruang & Pengukuran', tpKode: 'TP 2.3', pertemuan: '4 Pertemuan (14 JP)', model: 'Project-Based Learning (PjBL)', status: 'Siap Unduh' }
    ]
  },
  'Bahasa Inggris': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik menggunakan teks lisan, tulisan, dan visual dalam bahasa Inggris untuk berinteraksi dan berkomunikasi dalam konteks yang beragam. Peserta didik memahami ide utama dan detail relevan dari berbagai jenis teks (recount, narrative, descriptive, procedure, exposition) serta menyusun teks terstruktur.',
    tps: [
      { kode: 'TP 2.1', judul: 'Memahami Struktur dan Unsur Kebahasaan Recount Text (Pengalaman Pribadi)', jp: 12, status: 'SELESAI', elemen: 'Membaca - Memirsa' },
      { kode: 'TP 2.2', judul: 'Menyusun dan Menceritakan Pengalaman Masa Lampau Menggunakan Simple Past Tense', jp: 10, status: 'SELESAI', elemen: 'Menulis - Mempresentasikan' },
      { kode: 'TP 2.3', judul: 'Menganalisis Nilai Moral dan Alur Cerita pada Narrative Text (Fable & Folktale)', jp: 14, status: 'DRAFT', elemen: 'Menyimak - Berbicara' },
      { kode: 'TP 2.4', judul: 'Membuat Pesan Singkat, Pengumuman (Announcement), dan Pemberitahuan (Notice)', jp: 8, status: 'DRAFT', elemen: 'Menulis - Mempresentasikan' }
    ],
    moduls: [
      { judul: 'Modul Ajar: My Unforgettable Holiday (Recount Text)', tpKode: 'TP 2.1', pertemuan: '4 Pertemuan (12 JP)', model: 'Genre-Based Approach (GBA)', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Telling Past Events with Simple Past Tense', tpKode: 'TP 2.2', pertemuan: '3 Pertemuan (10 JP)', model: 'Discovery Learning', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Indonesian Folktales & Moral Values', tpKode: 'TP 2.3', pertemuan: '4 Pertemuan (14 JP)', model: 'Project-Based Learning (PjBL)', status: 'Siap Unduh' }
    ]
  },
  'IPAS (Ilmu Pengetahuan Alam dan Sosial)': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik mengidentifikasi keterkaitan antara pengetahuan alam dan sosial di lingkungan sekitar, memahami siklus makhluk hidup, wujud zat dan perubahannya, energi, keragaman budaya lokal, serta interaksi manusia dengan lingkungan melalui pengamatan ilmiah.',
    tps: [
      { kode: 'TP 2.1', judul: 'Menganalisis Bagian Tubuh Tumbuhan dan Proses Fotosintesis bagi Kehidupan', jp: 12, status: 'SELESAI', elemen: 'Pemahaman IPAS' },
      { kode: 'TP 2.2', judul: 'Mengidentifikasi Wujud Zat dan Perubahan Wujud Benda di Lingkungan Sekitar', jp: 14, status: 'SELESAI', elemen: 'Keterampilan Proses' },
      { kode: 'TP 2.3', judul: 'Mengenal Keragaman Budaya, Kearifan Lokal, dan Bentang Alam Daerah Setempat', jp: 12, status: 'DRAFT', elemen: 'Pemahaman IPAS' },
      { kode: 'TP 2.4', judul: 'Memahami Gaya di Sekitar Kita dan Pengaruhnya Terhadap Gerak Benda', jp: 10, status: 'DRAFT', elemen: 'Keterampilan Proses' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Tumbuhan Sumber Kehidupan di Bumi', tpKode: 'TP 2.1', pertemuan: '3 Pertemuan (12 JP)', model: 'Inquiry-Based Learning', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Eksperimen Perubahan Wujud Zat', tpKode: 'TP 2.2', pertemuan: '4 Pertemuan (14 JP)', model: 'Experimental Learning', status: 'Siap Unduh' }
    ]
  },
  'IPA (Ilmu Pengetahuan Alam)': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik mampu melakukan klasifikasi makhluk hidup dan benda, memahami struktur sel dan sistem organ manusia, menganalisis gaya, gerak, energi, getaran dan gelombang, serta menerapkan keterampilan proses inkuiri dalam penyelidikan ilmiah.',
    tps: [
      { kode: 'TP 2.1', judul: 'Menganalisis Sistem Pencernaan, Pernapasan, dan Peredaran Darah Manusia', jp: 14, status: 'SELESAI', elemen: 'Pemahaman IPA' },
      { kode: 'TP 2.2', judul: 'Menyelidiki Konsep Usaha, Energi, dan Pesawat Sederhana dalam Kehidupan', jp: 12, status: 'SELESAI', elemen: 'Keterampilan Proses' },
      { kode: 'TP 2.3', judul: 'Menganalisis Getaran, Gelombang, Cahaya, dan Alat Optik Sederhana', jp: 14, status: 'DRAFT', elemen: 'Pemahaman IPA' },
      { kode: 'TP 2.4', judul: 'Mengevaluasi Unsur, Senyawa, Campuran, dan Pemisahan Campuran Sederhana', jp: 10, status: 'DRAFT', elemen: 'Keterampilan Proses' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Sistem Organ Manusia & Pola Hidup Sehat', tpKode: 'TP 2.1', pertemuan: '4 Pertemuan (14 JP)', model: 'Inquiry-Based Learning', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Praktikum Getaran, Gelombang, dan Cahaya', tpKode: 'TP 2.3', pertemuan: '4 Pertemuan (14 JP)', model: 'Problem-Based Learning (PBL)', status: 'Siap Unduh' }
    ]
  },
  'IPS (Ilmu Pengetahuan Sosial)': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik memahami keberagaman kondisi geografis Indonesia, konektivitas antarruang, interaksi sosial dan dinamika kependudukan, kegiatan ekonomi masyarakat, serta perkembangan sejarah peradaban nusantara hingga masa modern.',
    tps: [
      { kode: 'TP 2.1', judul: 'Menganalisis Kondisi Geografis, Sumber Daya Alam, dan Pelestarian Lingkungan', jp: 12, status: 'SELESAI', elemen: 'Pemahaman Konsep' },
      { kode: 'TP 2.2', judul: 'Mengevaluasi Interaksi Sosial, Lembaga Sosial, dan Mobilitas Masyarakat', jp: 12, status: 'SELESAI', elemen: 'Pemahaman Konsep' },
      { kode: 'TP 2.3', judul: 'Menganalisis Kegiatan Ekonomi (Produksi, Distribusi, Konsumsi) dan Pasar', jp: 12, status: 'DRAFT', elemen: 'Keterampilan Proses' },
      { kode: 'TP 2.4', judul: 'Menelusuri Perkembangan Masyarakat Nusantara pada Masa Kerajaan hingga Kolonial', jp: 12, status: 'DRAFT', elemen: 'Keterampilan Proses' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Potensi Sumber Daya Alam & Kemaritiman Indonesia', tpKode: 'TP 2.1', pertemuan: '4 Pertemuan (12 JP)', model: 'Problem-Based Learning (PBL)', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Dinamika Kegiatan Ekonomi & Kewirausahaan', tpKode: 'TP 2.3', pertemuan: '4 Pertemuan (12 JP)', model: 'Project-Based Learning (PjBL)', status: 'Siap Unduh' }
    ]
  },
  'Pendidikan Pancasila': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik memahami kronologi sejarah kelahiran Pancasila, meneladani sikap para perumus Pancasila, menerapkan nilai-nilai Pancasila dalam kehidupan bermasyarakat, mematuhi norma dan UUD NRI Tahun 1945, serta menjaga keutuhan NKRI dalam bingkai Bhinneka Tunggal Ika.',
    tps: [
      { kode: 'TP 2.1', judul: 'Menganalisis Kedudukan dan Fungsi Pancasila sebagai Dasar Negara dan Pandangan Hidup', jp: 10, status: 'SELESAI', elemen: 'Pancasila' },
      { kode: 'TP 2.2', judul: 'Menerapkan Tata Urutan Peraturan Perundang-undangan dalam Kehidupan Sehari-hari', jp: 10, status: 'SELESAI', elemen: 'UUD NRI Tahun 1945' },
      { kode: 'TP 2.3', judul: 'Menghargai Keberagaman Suku, Agama, Ras, dan Antargolongan dalam Bhinneka Tunggal Ika', jp: 10, status: 'DRAFT', elemen: 'Bhinneka Tunggal Ika' },
      { kode: 'TP 2.4', judul: 'Menunjukkan Komitmen Menjaga Keutuhan Wilayah Negara Kesatuan Republik Indonesia', jp: 8, status: 'DRAFT', elemen: 'NKRI' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Mengamalkan Nilai Pancasila di Lingkungan Sekolah', tpKode: 'TP 2.1', pertemuan: '3 Pertemuan (10 JP)', model: 'Value Clarification Technique', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Harmoni dalam Keberagaman Bhinneka Tunggal Ika', tpKode: 'TP 2.3', pertemuan: '3 Pertemuan (10 JP)', model: 'Project-Based Learning (PjBL)', status: 'Siap Unduh' }
    ]
  },
  'Informatika': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik menerapkan berpikir komputasional untuk menyelesaikan persoalan kompleks, memahami sistem komputer dan jaringan internet, melakukan pengolahan serta analisis data digital, dan mengembangkan algoritma pemrograman terstruktur secara etis dan aman.',
    tps: [
      { kode: 'TP 2.1', judul: 'Menerapkan Strategi Algoritmik Standar (Searching & Sorting) pada Pemecahan Masalah', jp: 12, status: 'SELESAI', elemen: 'Berpikir Komputasional' },
      { kode: 'TP 2.2', judul: 'Menganalisis Jaringan Komputer Lokal, Internet, dan Keamanan Data Digital', jp: 10, status: 'SELESAI', elemen: 'Jaringan Komputer & Internet' },
      { kode: 'TP 2.3', judul: 'Melakukan Pengolahan, Visualisasi, dan Analisis Data Menggunakan Lembar Kerja', jp: 12, status: 'DRAFT', elemen: 'Analisis Data' },
      { kode: 'TP 2.4', judul: 'Mengembangkan Program Sederhana Menggunakan Bahasa Pemrograman Prosedural', jp: 14, status: 'DRAFT', elemen: 'Algoritma dan Pemrograman' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Berpikir Komputasional (Searching & Sorting)', tpKode: 'TP 2.1', pertemuan: '4 Pertemuan (12 JP)', model: 'Problem-Based Learning (PBL)', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Algoritma Pemrograman Dasar', tpKode: 'TP 2.4', pertemuan: '5 Pertemuan (14 JP)', model: 'Project-Based Learning (PjBL)', status: 'Siap Unduh' }
    ]
  },
  'PAI & Budi Pekerti': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik memahami ayat Al-Qur’an dan hadis pilihan tentang toleransi dan menuntut ilmu, memperkuat keimanan kepada kitab dan rasul Allah, membiasakan akhlak terpuji, melaksanakan ketentuan ibadah muamalah dengan benar, serta meneladani sejarah peradaban Islam.',
    tps: [
      { kode: 'TP 2.1', judul: 'Membaca, Mengartikan, dan Mengamalkan Kandungan Ayat tentang Sikap Toleransi', jp: 12, status: 'SELESAI', elemen: 'Al-Qur’an dan Hadis' },
      { kode: 'TP 2.2', judul: 'Meyakini Keberadaan Rasul-Rasul Allah dan Meneladani Sifat-Sifat Utamanya', jp: 10, status: 'SELESAI', elemen: 'Akidah' },
      { kode: 'TP 2.3', judul: 'Membiasakan Perilaku Jujur, Amanah, dan Menghindari Perilaku Tercela', jp: 10, status: 'DRAFT', elemen: 'Akhlak' },
      { kode: 'TP 2.4', judul: 'Memahami Ketentuan Ibadah Puasa, Zakat, dan Sejarah Peradaban Islam Nusantara', jp: 12, status: 'DRAFT', elemen: 'Fikih & Sejarah Peradaban Islam' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Indahnya Toleransi & Moderasi Beragama', tpKode: 'TP 2.1', pertemuan: '4 Pertemuan (12 JP)', model: 'Contextual Teaching & Learning', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Meneladani Sifat Amanah & Kejujuran', tpKode: 'TP 2.3', pertemuan: '3 Pertemuan (10 JP)', model: 'Discovery Learning', status: 'Siap Unduh' }
    ]
  },
  'PJOK': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik mempraktikkan variasi dan kombinasi keterampilan gerak spesifik dalam berbagai permainan olahraga, senam, dan aktivitas ritmik, merancang program latihan kebugaran jasmani personal, serta menerapkan pola hidup sehat dan keselamatan diri.',
    tps: [
      { kode: 'TP 2.1', judul: 'Mempraktikkan Variasi Gerak Spesifik Permainan Bola Besar dan Bola Kecil', jp: 12, status: 'SELESAI', elemen: 'Terampil Bergerak' },
      { kode: 'TP 2.2', judul: 'Mempraktikkan Keterampilan Gerak Dasar Atletik (Lari, Lompat, Lempar)', jp: 10, status: 'SELESAI', elemen: 'Terampil Bergerak' },
      { kode: 'TP 2.3', judul: 'Merancang dan Mengukur Latihan Kebugaran Jasmani Terkait Kesehatan', jp: 12, status: 'DRAFT', elemen: 'Belajar Melalui Gerak' },
      { kode: 'TP 2.4', judul: 'Menerapkan Budaya Hidup Sehat, Gizi Seimbang, dan Pencegahan Cedera', jp: 8, status: 'DRAFT', elemen: 'Memilih Hidup yang Menyehatkan' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Teknik Dasar & Strategi Permainan Bola Basket/Voli', tpKode: 'TP 2.1', pertemuan: '4 Pertemuan (12 JP)', model: 'Teaching Games for Understanding', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Tes & Pengukuran Kebugaran Jasmani Siswa', tpKode: 'TP 2.3', pertemuan: '4 Pertemuan (12 JP)', model: 'Practice Style', status: 'Siap Unduh' }
    ]
  },
  'Ekonomi': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik mampu memahami dan menganalisis permasalahan ekonomi nasional dan global, mengevaluasi kebijakan moneter dan fiskal, menyusun siklus akuntansi perusahaan jasa dan dagang sederhana, serta merancang strategi kewirausahaan.',
    tps: [
      { kode: 'TP 2.1', judul: 'Menganalisis Pendapatan Nasional, Pertumbuhan Ekonomi, dan Kesenjangan Distribusi', jp: 12, status: 'SELESAI', elemen: 'Pemahaman Konsep' },
      { kode: 'TP 2.2', judul: 'Mengevaluasi Dampak Inflasi, Kebijakan Moneter Bank Sentral, dan Kebijakan Fiskal', jp: 12, status: 'SELESAI', elemen: 'Keterampilan Proses' },
      { kode: 'TP 2.3', judul: 'Menganalisis Perdagangan Internasional, Neraca Pembayaran, dan Kurs Valuta Asing', jp: 10, status: 'DRAFT', elemen: 'Pemahaman Konsep' },
      { kode: 'TP 2.4', judul: 'Menyusun Siklus Akuntansi Perusahaan Jasa dan Laporan Laba Rugi Sederhana', jp: 16, status: 'DRAFT', elemen: 'Keterampilan Proses' }
    ],
    moduls: [
      { judul: 'Modul Ajar: Analisis Pendapatan Nasional & Pertumbuhan', tpKode: 'TP 2.1', pertemuan: '3 Pertemuan (12 JP)', model: 'Case-Based Learning', status: 'Siap Unduh' },
      { judul: 'Modul Ajar: Praktik Siklus Akuntansi Perusahaan Jasa', tpKode: 'TP 2.4', pertemuan: '4 Pertemuan (16 JP)', model: 'Project-Based Learning (PjBL)', status: 'Siap Unduh' }
    ]
  }
};

const FASE_LIST = [
  { label: 'Fase A (Kelas 1-2 SD)', defaultKelas: '1', mapels: ['Bahasa Indonesia', 'Matematika', 'Pendidikan Pancasila', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase B (Kelas 3-4 SD)', defaultKelas: '4', mapels: ['IPAS (Ilmu Pengetahuan Alam dan Sosial)', 'Bahasa Indonesia', 'Matematika', 'Pendidikan Pancasila', 'Bahasa Inggris', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase C (Kelas 5-6 SD)', defaultKelas: '5', mapels: ['Matematika', 'IPAS (Ilmu Pengetahuan Alam dan Sosial)', 'Bahasa Indonesia', 'Pendidikan Pancasila', 'Bahasa Inggris', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase D (Kelas 7-9 SMP)', defaultKelas: '8', mapels: ['Bahasa Inggris', 'Bahasa Indonesia', 'Matematika', 'IPA (Ilmu Pengetahuan Alam)', 'IPS (Ilmu Pengetahuan Sosial)', 'Pendidikan Pancasila', 'Informatika', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase E (Kelas 10 SMA/SMK)', defaultKelas: '10', mapels: ['Informatika', 'Bahasa Indonesia', 'Bahasa Inggris', 'Matematika', 'IPA (Ilmu Pengetahuan Alam)', 'IPS (Ilmu Pengetahuan Sosial)', 'Pendidikan Pancasila', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase F (Kelas 11-12 SMA/SMK)', defaultKelas: '11', mapels: ['Ekonomi', 'Bahasa Inggris', 'Bahasa Indonesia', 'Matematika', 'Informatika', 'Pendidikan Pancasila', 'PAI & Budi Pekerti', 'PJOK'] }
];

const DEFAULT_STUDENTS = [
  { id: 1, nisn: '0123456701', nama: 'Ahmad Budi Santoso', jk: 'L', statusHadir: 'H', f1: 85, f2: 88, sumatif: 86 },
  { id: 2, nisn: '0123456702', nama: 'Citra Lestari', jk: 'P', statusHadir: 'H', f1: 92, f2: 90, sumatif: 94 },
  { id: 3, nisn: '0123456703', nama: 'Dian Wardhana', jk: 'L', statusHadir: 'S', f1: 78, f2: 80, sumatif: 79 },
  { id: 4, nisn: '0123456704', nama: 'Eka Putri Setiawati', jk: 'P', statusHadir: 'H', f1: 88, f2: 91, sumatif: 89 },
  { id: 5, nisn: '0123456705', nama: 'Fajar Nugroho', jk: 'L', statusHadir: 'A', f1: 72, f2: 75, sumatif: 74 }
];

const createClassObject = (id, nama, fase, mapel, kelas, customStudents = null) => {
  const bank = SUBJECT_BANK[mapel];
  const cpText = bank
    ? `[${fase}] ${bank.cp}`
    : `Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025 pada ${fase}, peserta didik mampu memahami, menerapkan, dan menganalisis konsep-konsep esensial pada mata pelajaran ${mapel} secara kritis, kreatif, dan kontekstual sesuai perkembangan jenjang Kelas ${kelas}.`;

  const tps = bank
    ? bank.tps.map((t, idx) => ({ ...t, id: `tp-${id}-${idx + 1}` }))
    : [
        { id: `tp-${id}-1`, kode: 'TP 2.1', judul: `Memahami Konsep Dasar dan Ruang Lingkup ${mapel} Kelas ${kelas}`, jp: 12, status: 'SELESAI', elemen: 'Pemahaman Konsep' },
        { id: `tp-${id}-2`, kode: 'TP 2.2', judul: `Menganalisis Penerapan Materi ${mapel} dalam Kehidupan Sehari-hari`, jp: 12, status: 'DRAFT', elemen: 'Keterampilan Proses' },
        { id: `tp-${id}-3`, kode: 'TP 2.3', judul: `Menyusun Proyek dan Evaluasi Pembelajaran ${mapel}`, jp: 12, status: 'DRAFT', elemen: 'Aplikasi & Proyek' }
      ];

  const modulAjar = bank
    ? bank.moduls.map((m, idx) => ({ ...m, id: `ma-${id}-${idx + 1}` }))
    : [
        { id: `ma-${id}-1`, judul: `Modul Ajar: Konsep & Aplikasi ${mapel} (Kelas ${kelas})`, tpKode: 'TP 2.1', pertemuan: '4 Pertemuan (12 JP)', model: 'Problem-Based Learning (PBL)', status: 'Siap Unduh' }
      ];

  const studentsList = customStudents && customStudents.length > 0 ? customStudents : JSON.parse(JSON.stringify(DEFAULT_STUDENTS));

  return {
    id,
    nama,
    fase,
    mapel,
    kelas,
    cpText,
    tps,
    modulAjar,
    students: studentsList,
    journals: [
      {
        id: `j-${id}-1`,
        tanggal: new Date().toISOString().split('T')[0],
        jamKe: '1-2',
        tpKode: `${tps[0].kode} - ${tps[0].judul}`,
        catatan: 'Siswa aktif mengikuti kegiatan inti dan diskusi kelompok.',
        hadirCount: studentsList.length,
        tidakHadirCount: 0
      }
    ]
  };
};

// --- PARSER EXCEL (.XLSX / .XLS / .CSV) DAPODIK & EMIS ---
const loadSheetJSLibrary = () => {
  return new Promise((resolve, reject) => {
    if (window.XLSX) {
      resolve(window.XLSX);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
    script.onload = () => resolve(window.XLSX);
    script.onerror = () => reject(new Error('Gagal memuat pembaca Excel'));
    document.head.appendChild(script);
  });
};

const extractStudentsFromRows = (rows) => {
  if (!Array.isArray(rows) || rows.length === 0) return [];

  let headerIdx = -1;
  let nameCol = -1;
  let nisnCol = -1;
  let jkCol = -1;

  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i];
    if (!Array.isArray(row)) continue;
    const lowerCells = row.map((c) => String(c || '').toLowerCase().trim());
    const foundName = lowerCells.findIndex((c) => c === 'nama' || c.includes('nama peserta') || c.includes('nama siswa') || c.includes('nama lengkap'));
    if (foundName !== -1) {
      headerIdx = i;
      nameCol = foundName;
      nisnCol = lowerCells.findIndex((c) => c.includes('nisn') || c === 'nipd' || c === 'nis');
      jkCol = lowerCells.findIndex((c) => c === 'jk' || c === 'l/p' || c.includes('jenis kelamin') || c.includes('gender'));
      break;
    }
  }

  const startRow = headerIdx !== -1 ? headerIdx + 1 : 1;
  if (nameCol === -1) nameCol = 1;
  if (nisnCol === -1) nisnCol = 2;
  if (jkCol === -1) jkCol = 3;

  const parsed = [];
  for (let r = startRow; r < rows.length; r++) {
    const row = rows[r];
    if (!Array.isArray(row)) continue;
    const rawName = String(row[nameCol] || '').trim();
    if (!rawName || rawName.length < 2 || rawName.toLowerCase().includes('total') || rawName.toLowerCase().includes('jumlah')) {
      continue;
    }
    const rawNisn = String(row[nisnCol] || '').replace(/[^0-9]/g, '').trim();
    const rawJk = String(row[jkCol] || 'L').trim().toUpperCase();
    const jk = rawJk.startsWith('P') || rawJk === 'PEREMPUAN' || rawJk === 'FEMALE' ? 'P' : 'L';

    parsed.push({
      id: parsed.length + 1,
      nisn: rawNisn || `012345${String(parsed.length + 1).padStart(4, '0')}`,
      nama: rawName,
      jk,
      statusHadir: 'H',
      f1: 80,
      f2: 80,
      sumatif: 80
    });
  }
  return parsed;
};

const INITIAL_STATE = {
  isLoggedIn: false,
  wizardCompleted: false,
  profile: {
    nama: 'Budi Santoso, S.Pd',
    nip: '19850212 201001 1 015',
    sekolah: 'Sekolah Nusantara Merdeka',
    tahunAjaran: '2026/2027',
    semester: 'Smt Genap'
  },
  activeClassId: 'cls-smp-8b',
  classes: [
    createClassObject('cls-smp-8b', 'Kelas 8 B - Matematika', 'Fase D (Kelas 7-9 SMP)', 'Matematika', '8')
  ]
};

const STORAGE_KEY = 'eduadmin_saas_word_v6';

export default function App() {
  const [data, setData] = useState(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      return saved ? JSON.parse(saved) : INITIAL_STATE;
    } catch {
      return INITIAL_STATE;
    }
  });

  // --- STATE LOGIN ---
  const [loginEmail, setLoginEmail] = useState('budi.guru@sekolah.id');
  const [loginPass, setLoginPass] = useState('••••••••••••');

  // --- STATE WIZARD 3 LANGKAH ---
  const [wizardStep, setWizardStep] = useState(1);
  const [wizForm, setWizForm] = useState({
    fase: 'Fase D (Kelas 7-9 SMP)',
    kelas: '8',
    rombel: 'B',
    mapel: 'Matematika'
  });
  const [importFormat, setImportFormat] = useState(null);
  const [uploadedFile, setUploadedFile] = useState(null);
  const [parsedStudents, setParsedStudents] = useState(null);
  const [isParsingFile, setIsParsingFile] = useState(false);
  const fileInputRef = useRef(null);

  // --- STATE DASHBOARD ---
  const [activeMenu, setActiveMenu] = useState('cp-tp');
  const [toast, setToast] = useState(null);
  const [editingTp, setEditingTp] = useState(null);
  const [showAddTp, setShowAddTp] = useState(false);
  const [newTp, setNewTp] = useState({ kode: '', judul: '', jp: 8, status: 'DRAFT', elemen: 'Pemahaman Konsep' });

  const [showAddRombel, setShowAddRombel] = useState(false);
  const [rombelForm, setRombelForm] = useState({
    fase: 'Fase D (Kelas 7-9 SMP)',
    kelas: '7',
    rombel: 'A',
    mapel: 'Bahasa Inggris'
  });

  const [newStudentName, setNewStudentName] = useState('');
  const [newStudentNisn, setNewStudentNisn] = useState('');
  const [newStudentJk, setNewStudentJk] = useState('L');

  const activeClass = data.classes.find((c) => c.id === data.activeClassId) || data.classes[0];
  const [editingCp, setEditingCp] = useState(false);
  const [cpDraft, setCpDraft] = useState(activeClass.cpText);

  const [journalForm, setJournalForm] = useState({
    tanggal: new Date().toISOString().split('T')[0],
    jamKe: '1-2',
    tpKode: activeClass.tps[0] ? `${activeClass.tps[0].kode} - ${activeClass.tps[0].judul}` : '',
    catatan: ''
  });

  useEffect(() => {
    setEditingCp(false);
    setCpDraft(activeClass.cpText);
    setJournalForm((prev) => ({
      ...prev,
      tpKode: activeClass.tps[0] ? `${activeClass.tps[0].kode} - ${activeClass.tps[0].judul}` : ''
    }));
  }, [activeClass.id]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    } catch (e) {
      console.error('Gagal menyimpan:', e);
    }
  }, [data]);

  const showNotification = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 3200);
  };

  // --- HANDLER EKSPOR WORD (.DOC) ---
  const handleExportAtpWord = () => {
    const totalJp = activeClass.tps.reduce((acc, item) => acc + (Number(item.jp) || 0), 0);
    const rowsHtml = activeClass.tps
      .map(
        (tp, idx) => `
        <tr>
          <td style="text-align:center;">${idx + 1}</td>
          <td style="text-align:center;"><strong>${tp.kode}</strong></td>
          <td>${tp.elemen}</td>
          <td>${tp.judul}</td>
          <td style="text-align:center;">${tp.jp} JP</td>
        </tr>`
      )
      .join('');

    const body = `
      <p><strong>A. CAPAIAN PEMBELAJARAN (CP)</strong></p>
      <div class="box-cp">${activeClass.cpText}</div>
      <p><strong>B. ALUR TUJUAN PEMBELAJARAN (ATP)</strong></p>
      <table class="data-table">
        <thead>
          <tr>
            <th style="width:6%;">No</th>
            <th style="width:12%;">Kode TP</th>
            <th style="width:22%;">Elemen CP</th>
            <th style="width:48%;">Tujuan Pembelajaran</th>
            <th style="width:12%;">Alokasi</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
          <tr>
            <td colspan="4" style="text-align:right;"><strong>Total Alokasi Waktu</strong></td>
            <td style="text-align:center;"><strong>${totalJp} JP</strong></td>
          </tr>
        </tbody>
      </table>
    `;
    downloadWordDocument(`CP_ATP_${activeClass.nama.replace(/\s+/g, '_')}`, 'CAPAIAN & ALUR TUJUAN PEMBELAJARAN (ATP)', data.profile, activeClass, body);
    showNotification('File Word CP & ATP (.doc) berhasil diunduh!');
  };

  const handleExportPromesWord = () => {
    const rowsHtml = activeClass.tps
      .map(
        (tp, idx) => `
        <tr>
          <td style="text-align:center;"><strong>${tp.kode}</strong></td>
          <td>${tp.judul}</td>
          <td style="text-align:center;">${tp.jp} JP</td>
          <td style="text-align:center;">${idx === 0 ? 'V' : ''}</td>
          <td style="text-align:center;">${idx === 0 || idx === 1 ? 'V' : ''}</td>
          <td style="text-align:center;">${idx === 1 || idx === 2 ? 'V' : ''}</td>
          <td style="text-align:center;">${idx === 2 || idx === 3 ? 'V' : ''}</td>
          <td style="text-align:center;">${idx >= 3 ? 'V' : ''}</td>
        </tr>`
      )
      .join('');

    const body = `
      <p><strong>DISTRIBUSI ALOKASI WAKTU PROGRAM SEMESTER (PROMES)</strong></p>
      <table class="data-table">
        <thead>
          <tr>
            <th style="width:12%;">Kode</th>
            <th style="width:48%;">Tujuan Pembelajaran / Materi Pokok</th>
            <th style="width:10%;">JP</th>
            <th style="width:6%;">Jan</th>
            <th style="width:6%;">Feb</th>
            <th style="width:6%;">Mar</th>
            <th style="width:6%;">Apr</th>
            <th style="width:6%;">Mei</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    `;
    downloadWordDocument(`Promes_${activeClass.nama.replace(/\s+/g, '_')}`, 'PROGRAM SEMESTER (PROMES)', data.profile, activeClass, body);
    showNotification('File Word Program Semester (.doc) berhasil diunduh!');
  };

  const handleExportModulWord = (modul) => {
    const relatedTp = activeClass.tps.find((t) => t.kode === modul.tpKode) || activeClass.tps[0];
    const body = `
      <p><strong>I. INFORMASI UMUM</strong></p>
      <table class="data-table">
        <tr><td style="width:30%;"><strong>Judul Modul Ajar</strong></td><td>${modul.judul}</td></tr>
        <tr><td><strong>Alokasi Waktu</strong></td><td>${modul.pertemuan}</td></tr>
        <tr><td><strong>Model Pembelajaran</strong></td><td>${modul.model}</td></tr>
        <tr><td><strong>Profil Pelajar Pancasila</strong></td><td>Beriman & Bertakwa, Bernalar Kritis, Mandiri, Gotong Royong</td></tr>
      </table>

      <p style="margin-top:14px;"><strong>II. KOMPONEN INTI</strong></p>
      <table class="data-table">
        <tr><td style="width:30%;"><strong>Tujuan Pembelajaran (${modul.tpKode})</strong></td><td>${relatedTp ? relatedTp.judul : modul.judul}</td></tr>
        <tr><td><strong>Pemahaman Bermakna</strong></td><td>Peserta didik mampu menghubungkan konsep ${activeClass.mapel} dengan pemecahan masalah nyata dalam kehidupan sehari-hari.</td></tr>
        <tr><td><strong>Kegiatan Pendahuluan (15 Menit)</strong></td><td>Guru membuka pelajaran dengan salam, doa, presensi, apersepsi kontekstual, dan menyampaikan tujuan pembelajaran.</td></tr>
        <tr><td><strong>Kegiatan Inti (60 Menit)</strong></td><td>Peserta didik mengamati stimulus masalah, berdiskusi dalam kelompok kecil menggunakan LKPD, mengolah data/informasi, dan mempresentasikan hasilnya dengan model <em>${modul.model}</em>.</td></tr>
        <tr><td><strong>Kegiatan Penutup (15 Menit)</strong></td><td>Guru dan peserta didik menyimpulkan materi, melakukan refleksi pembelajaran, serta menyampaikan rencana pertemuan berikutnya.</td></tr>
        <tr><td><strong>Asesmen / Penilaian</strong></td><td>Asesmen Diagnostik (Awal), Asesmen Formatif (Observasi & LKPD), dan Asesmen Sumatif Lingkup Materi.</td></tr>
      </table>
    `;
    downloadWordDocument(`Modul_Ajar_${modul.tpKode.replace(/\s+/g, '_')}`, 'MODUL AJAR KURIKULUM MERDEKA', data.profile, activeClass, body);
    showNotification(`File Word "${modul.judul}" (.doc) berhasil diunduh!`);
  };

  const handleExportPresensiWord = () => {
    const rowsHtml = activeClass.students
      .map(
        (s, idx) => `
        <tr>
          <td style="text-align:center;">${idx + 1}</td>
          <td style="text-align:center;">${s.nisn}</td>
          <td>${s.nama}</td>
          <td style="text-align:center;">${s.jk}</td>
          <td style="text-align:center;"><strong>${s.statusHadir}</strong></td>
          <td>${s.statusHadir === 'H' ? 'Hadir' : s.statusHadir === 'S' ? 'Sakit' : s.statusHadir === 'I' ? 'Izin' : 'Alpa'}</td>
        </tr>`
      )
      .join('');

    const body = `
      <p><strong>REKAPITULASI KEHADIRAN PESERTA DIDIK</strong></p>
      <table class="data-table">
        <thead>
          <tr>
            <th style="width:6%;">No</th>
            <th style="width:18%;">NISN</th>
            <th style="width:38%;">Nama Lengkap Siswa</th>
            <th style="width:8%;">L/P</th>
            <th style="width:10%;">Kode</th>
            <th style="width:20%;">Keterangan</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    `;
    downloadWordDocument(`Presensi_${activeClass.nama.replace(/\s+/g, '_')}`, 'DAFTAR HADIR PESERTA DIDIK', data.profile, activeClass, body);
    showNotification('File Word Daftar Hadir (.doc) berhasil diunduh!');
  };

  const handleExportNilaiWord = () => {
    const rowsHtml = activeClass.students
      .map((s, idx) => {
        const akhir = Math.round((Number(s.f1) + Number(s.f2) + Number(s.sumatif)) / 3);
        const predikat = akhir >= 90 ? 'A (Sangat Baik)' : akhir >= 80 ? 'B (Baik)' : 'C (Cukup)';
        return `
        <tr>
          <td style="text-align:center;">${idx + 1}</td>
          <td style="text-align:center;">${s.nisn}</td>
          <td>${s.nama}</td>
          <td style="text-align:center;">${s.f1}</td>
          <td style="text-align:center;">${s.f2}</td>
          <td style="text-align:center;">${s.sumatif}</td>
          <td style="text-align:center;"><strong>${akhir}</strong></td>
          <td style="text-align:center;">${predikat}</td>
        </tr>`;
      })
      .join('');

    const body = `
      <p><strong>LEGER PENILAIAN FORMATIF & SUMATIF (e-RAPOR)</strong></p>
      <table class="data-table">
        <thead>
          <tr>
            <th style="width:5%;">No</th>
            <th style="width:15%;">NISN</th>
            <th style="width:30%;">Nama Siswa</th>
            <th style="width:9%;">Form. 1</th>
            <th style="width:9%;">Form. 2</th>
            <th style="width:10%;">Sumatif</th>
            <th style="width:10%;">Rapor</th>
            <th style="width:12%;">Predikat</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    `;
    downloadWordDocument(`Leger_Nilai_${activeClass.nama.replace(/\s+/g, '_')}`, 'DAFTAR NILAI PESERTA DIDIK', data.profile, activeClass, body);
    showNotification('File Word Leger Nilai (.doc) berhasil diunduh!');
  };

  // --- HANDLER LOGIN ---
  const handleLogin = (e) => {
    e.preventDefault();
    setData((prev) => ({ ...prev, isLoggedIn: true }));
  };

  // --- HANDLER WIZARD IMPOR DAPODIK & EMIS ---
  const handleSelectFormat = (format) => {
    setImportFormat(format);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
      fileInputRef.current.click();
    }
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const activeFmt = importFormat || 'dapodik';
    const formatLabel = activeFmt === 'emis' ? 'Format Emis' : 'Format Dapodik';
    setIsParsingFile(true);

    try {
      const XLSX = await loadSheetJSLibrary();
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
      const studentsExtracted = extractStudentsFromRows(rows);

      setParsedStudents(studentsExtracted.length > 0 ? studentsExtracted : null);
      setUploadedFile({
        name: file.name,
        size: `${(file.size / 1024).toFixed(1)} KB`,
        format: formatLabel,
        count: studentsExtracted.length
      });
    } catch (err) {
      console.error('Error membaca file:', err);
      setParsedStudents(null);
      setUploadedFile({
        name: file.name,
        size: `${(file.size / 1024).toFixed(1)} KB`,
        format: formatLabel,
        count: 0
      });
    } finally {
      setIsParsingFile(false);
    }
  };

  const handleFinishWizard = () => {
    setWizardStep(3);
    setTimeout(() => {
      const namaKelas = `Kelas ${wizForm.kelas} ${wizForm.rombel} - ${wizForm.mapel}`;
      const newId = `cls-${Date.now()}`;
      const created = createClassObject(
        newId,
        namaKelas,
        wizForm.fase,
        wizForm.mapel,
        wizForm.kelas,
        parsedStudents
      );
      setData((prev) => ({
        ...prev,
        wizardCompleted: true,
        activeClassId: newId,
        classes: [created]
      }));
      showNotification(`Perangkat ${namaKelas} (${wizForm.fase}) berhasil disiapkan!`);
    }, 800);
  };

  // --- HELPER DASHBOARD ---
  const updateActiveClass = (updater) => {
    setData((prev) => ({
      ...prev,
      classes: prev.classes.map((cls) => (cls.id === prev.activeClassId ? updater(cls) : cls))
    }));
  };

  const handleAddTp = (e) => {
    e.preventDefault();
    if (!newTp.kode.trim() || !newTp.judul.trim()) return;
    const item = {
      id: `tp-${Date.now()}`,
      kode: newTp.kode.trim(),
      judul: newTp.judul.trim(),
      jp: Number(newTp.jp) || 8,
      status: newTp.status,
      elemen: newTp.elemen
    };
    updateActiveClass((cls) => ({ ...cls, tps: [...cls.tps, item] }));
    setNewTp({ kode: '', judul: '', jp: 8, status: 'DRAFT', elemen: 'Pemahaman Konsep' });
    setShowAddTp(false);
    showNotification('Tujuan Pembelajaran baru berhasil disimpan!');
  };

  const handleUpdateTp = (e) => {
    e.preventDefault();
    if (!editingTp) return;
    updateActiveClass((cls) => ({
      ...cls,
      tps: cls.tps.map((t) => (t.id === editingTp.id ? editingTp : t))
    }));
    setEditingTp(null);
    showNotification('Perubahan Tujuan Pembelajaran disimpan!');
  };

  const handleDeleteTp = (id) => {
    updateActiveClass((cls) => ({
      ...cls,
      tps: cls.tps.filter((t) => t.id !== id)
    }));
    setEditingTp(null);
    showNotification('Tujuan Pembelajaran dihapus.');
  };

  const handleStatusHadir = (studentId, status) => {
    updateActiveClass((cls) => ({
      ...cls,
      students: cls.students.map((s) => (s.id === studentId ? { ...s, statusHadir: status } : s))
    }));
  };

  const handleSaveJournal = (e) => {
    e.preventDefault();
    const hadirCount = activeClass.students.filter((s) => s.statusHadir === 'H').length;
    const tidakHadirCount = activeClass.students.length - hadirCount;
    const newEntry = {
      id: `j-${Date.now()}`,
      tanggal: journalForm.tanggal,
      jamKe: journalForm.jamKe || '1-2',
      tpKode: journalForm.tpKode || (activeClass.tps[0] ? `${activeClass.tps[0].kode} - ${activeClass.tps[0].judul}` : 'Materi Umum'),
      catatan: journalForm.catatan || 'Kegiatan pembelajaran berjalan lancar.',
      hadirCount,
      tidakHadirCount
    };
    updateActiveClass((cls) => ({
      ...cls,
      journals: [newEntry, ...cls.journals]
    }));
    setJournalForm((prev) => ({ ...prev, catatan: '' }));
    showNotification('Jurnal mengajar & rekap presensi berhasil disimpan!');
  };

  const handleAddStudent = (e) => {
    e.preventDefault();
    if (!newStudentName.trim()) return;
    const nextId = activeClass.students.length > 0 ? Math.max(...activeClass.students.map((s) => s.id)) + 1 : 1;
    const student = {
      id: nextId,
      nisn: newStudentNisn.trim() || `012345670${nextId}`,
      nama: newStudentName.trim(),
      jk: newStudentJk,
      statusHadir: 'H',
      f1: 80,
      f2: 80,
      sumatif: 80
    };
    updateActiveClass((cls) => ({ ...cls, students: [...cls.students, student] }));
    setNewStudentName('');
    setNewStudentNisn('');
    showNotification('Siswa baru berhasil ditambahkan!');
  };

  const handleGradeChange = (studentId, field, val) => {
    const num = Math.max(0, Math.min(100, Number(val) || 0));
    updateActiveClass((cls) => ({
      ...cls,
      students: cls.students.map((s) => (s.id === studentId ? { ...s, [field]: num } : s))
    }));
  };

  const handleAddRombel = (e) => {
    e.preventDefault();
    const namaLengkap = `Kelas ${rombelForm.kelas} ${rombelForm.rombel} - ${rombelForm.mapel}`;
    const id = `cls-${Date.now()}`;
    const newCls = createClassObject(id, namaLengkap, rombelForm.fase, rombelForm.mapel, rombelForm.kelas);
    setData((prev) => ({
      ...prev,
      classes: [...prev.classes, newCls],
      activeClassId: id
    }));
    setShowAddRombel(false);
    showNotification(`Rombel "${namaLengkap}" (${rombelForm.fase}) berhasil dibuat!`);
  };

  const handleResetAll = () => {
    localStorage.removeItem(STORAGE_KEY);
    setWizardStep(1);
    setImportFormat(null);
    setUploadedFile(null);
    setParsedStudents(null);
    setData(INITIAL_STATE);
  };

  // ============================================================================
  // TAMPILAN 1: HALAMAN LOGIN
  // ============================================================================
  if (!data.isLoggedIn) {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4 font-sans">
        <div className="bg-white rounded-3xl shadow-xl max-w-md w-full overflow-hidden border border-slate-200/80">
          <div className="bg-blue-600 p-8 text-center text-white">
            <div className="w-14 h-14 bg-white/15 rounded-2xl flex items-center justify-center mx-auto mb-4">
              <BookOpen className="w-7 h-7 text-white" />
            </div>
            <h1 className="text-2xl font-bold tracking-tight">EduAdmin SaaS</h1>
            <p className="text-blue-100 text-sm mt-1">Platform Administrasi Guru Terpadu</p>
          </div>

          <form onSubmit={handleLogin} className="p-8 space-y-5">
            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1.5">Email atau NIP</label>
              <input
                type="text"
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
                className="w-full rounded-xl border border-slate-200 px-4 py-3 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500"
                required
              />
            </div>

            <div>
              <div className="flex items-center justify-between mb-1.5">
                <label className="block text-xs font-bold text-slate-700">Kata Sandi</label>
                <span className="text-xs font-semibold text-blue-600 cursor-pointer">Lupa Sandi?</span>
              </div>
              <input
                type="password"
                value={loginPass}
                onChange={(e) => setLoginPass(e.target.value)}
                className="w-full rounded-xl border border-slate-200 px-4 py-3 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500"
                required
              />
            </div>

            <button
              type="submit"
              className="w-full py-3.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm transition shadow-sm cursor-pointer"
            >
              Masuk ke Aplikasi
            </button>

            <div className="text-center pt-2 text-xs text-slate-500">
              Belum memiliki lisensi?{' '}
              <a
                href={LYNK_ID_URL}
                target="_blank"
                rel="noopener noreferrer"
                className="font-bold text-blue-600 hover:underline inline-flex items-center gap-1"
              >
                Beli Akses di Lynk.id <ExternalLink className="w-3 h-3" />
              </a>
            </div>
          </form>
        </div>
      </div>
    );
  }

  // ============================================================================
  // TAMPILAN 2: WIZARD 3 LANGKAH
  // ============================================================================
  if (!data.wizardCompleted) {
    const selectedFaseObj = FASE_LIST.find((f) => f.label === wizForm.fase) || FASE_LIST[3];

    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4 font-sans">
        <div className="bg-white rounded-3xl shadow-xl max-w-2xl w-full overflow-hidden border border-slate-200">
          <div className="px-8 py-5 border-b border-slate-100 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-blue-600 text-white flex items-center justify-center">
                <BookOpen className="w-5 h-5" />
              </div>
              <span className="font-bold text-lg text-slate-900">Sistem Administrasi Guru</span>
            </div>
            <span className="px-3 py-1 rounded-full bg-blue-50 text-blue-600 text-xs font-bold">
              Tahun Ajaran {data.profile.tahunAjaran}
            </span>
          </div>

          <div className="px-8 py-5 bg-slate-50/70 border-b border-slate-100 flex items-center justify-around">
            {[
              { num: 1, label: 'Data Kelas' },
              { num: 2, label: 'Impor Siswa' },
              { num: 3, label: 'Proses' }
            ].map((s) => (
              <div key={s.num} className="flex flex-col items-center gap-1.5">
                <div
                  className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold transition ${
                    wizardStep >= s.num
                      ? 'bg-blue-600 text-white shadow-sm'
                      : 'bg-white border border-slate-300 text-slate-400'
                  }`}
                >
                  {s.num}
                </div>
                <span className={`text-xs font-semibold ${wizardStep >= s.num ? 'text-blue-600' : 'text-slate-400'}`}>
                  {s.label}
                </span>
              </div>
            ))}
          </div>

          {wizardStep === 1 && (
            <div className="p-8 space-y-5">
              <div>
                <h2 className="text-xl font-bold text-slate-900">Pengaturan Kelas & Mata Pelajaran</h2>
                <p className="text-sm text-slate-500 mt-1">
                  Pilih jenjang dan mata pelajaran Anda. Perangkat CP, ATP, Prota/Promes, dan Modul Ajar otomatis tersedia.
                </p>
              </div>

              <div className="space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">Jenjang / Fase Kurikulum Merdeka</label>
                  <select
                    value={wizForm.fase}
                    onChange={(e) => {
                      const found = FASE_LIST.find((f) => f.label === e.target.value) || FASE_LIST[3];
                      setWizForm({
                        ...wizForm,
                        fase: found.label,
                        kelas: found.defaultKelas,
                        mapel: found.mapels[0]
                      });
                    }}
                    className="w-full rounded-xl border border-slate-200 px-4 py-2.5 text-sm bg-white"
                  >
                    {FASE_LIST.map((f) => (
                      <option key={f.label} value={f.label}>
                        {f.label}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1.5">Tingkat Kelas</label>
                    <select
                      value={wizForm.kelas}
                      onChange={(e) => setWizForm({ ...wizForm, kelas: e.target.value })}
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2.5 text-sm bg-white"
                    >
                      {['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'].map((k) => (
                        <option key={k} value={k}>
                          Kelas {k}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1.5">Nama Rombel</label>
                    <input
                      type="text"
                      value={wizForm.rombel}
                      onChange={(e) => setWizForm({ ...wizForm, rombel: e.target.value })}
                      placeholder="Contoh: A / B / IPA-1"
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2.5 text-sm"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">Mata Pelajaran</label>
                  <select
                    value={wizForm.mapel}
                    onChange={(e) => setWizForm({ ...wizForm, mapel: e.target.value })}
                    className="w-full rounded-xl border border-slate-200 px-3.5 py-2.5 text-sm bg-white"
                  >
                    {selectedFaseObj.mapels.map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="pt-4 flex justify-end">
                <button
                  type="button"
                  onClick={() => setWizardStep(2)}
                  className="px-6 py-3 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-semibold text-sm flex items-center gap-2 cursor-pointer"
                >
                  Lanjut Impor Siswa <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}

          {wizardStep === 2 && (
            <div className="p-8 space-y-6">
              <div>
                <h2 className="text-xl font-bold text-slate-900">Impor Data Peserta Didik</h2>
                <p className="text-sm text-slate-500 mt-1">
                  Sistem kami terintegrasi dengan format unduhan standar Dapodik dan Emis Kementerian. Tidak perlu input manual.
                </p>
              </div>

              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleFileUpload}
                className="hidden"
              />

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                <button
                  type="button"
                  onClick={() => handleSelectFormat('dapodik')}
                  className={`py-3.5 px-4 rounded-xl font-semibold text-sm border transition cursor-pointer flex items-center justify-center gap-2 ${
                    importFormat === 'dapodik'
                      ? 'border-blue-600 text-blue-600 bg-blue-50/40 ring-2 ring-blue-100'
                      : 'border-slate-200 text-slate-700 bg-white hover:border-blue-400 hover:bg-slate-50'
                  }`}
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  <span>Format Dapodik</span>
                </button>

                <button
                  type="button"
                  onClick={() => handleSelectFormat('emis')}
                  className={`py-3.5 px-4 rounded-xl font-semibold text-sm border transition cursor-pointer flex items-center justify-center gap-2 ${
                    importFormat === 'emis'
                      ? 'border-blue-600 text-blue-600 bg-blue-50/40 ring-2 ring-blue-100'
                      : 'border-slate-200 text-slate-700 bg-white hover:border-blue-400 hover:bg-slate-50'
                  }`}
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  <span>Format Emis</span>
                </button>
              </div>

              {isParsingFile ? (
                <div className="p-6 rounded-xl border border-blue-200 bg-blue-50/50 text-center text-sm text-blue-700 font-medium">
                  Membaca data peserta didik dari file Excel...
                </div>
              ) : uploadedFile ? (
                <div className="p-4 rounded-xl bg-emerald-50 border border-emerald-200 flex items-center justify-between gap-3">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center">
                      <FileSpreadsheet className="w-5 h-5" />
                    </div>
                    <div>
                      <div className="text-sm font-bold text-slate-800">{uploadedFile.name}</div>
                      <div className="text-xs text-emerald-700 font-medium">
                        Berhasil diunggah ({uploadedFile.format}) • Ukuran {uploadedFile.size}
                        {uploadedFile.count > 0 ? ` • ${uploadedFile.count} Siswa terbaca` : ''}
                      </div>
                    </div>
                  </div>
                  <div className="flex items-center gap-3">
                    <button
                      type="button"
                      onClick={() => fileInputRef.current?.click()}
                      className="text-xs font-semibold text-slate-600 underline hover:text-blue-600 cursor-pointer"
                    >
                      Ganti File
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setUploadedFile(null);
                        setParsedStudents(null);
                        if (fileInputRef.current) fileInputRef.current.value = '';
                      }}
                      className="text-xs font-semibold text-rose-600 hover:underline cursor-pointer"
                    >
                      Hapus
                    </button>
                  </div>
                </div>
              ) : (
                <div
                  onClick={() => handleSelectFormat(importFormat || 'dapodik')}
                  className="p-6 rounded-xl border-2 border-dashed border-slate-200 text-center hover:border-blue-400 hover:bg-slate-50/70 transition cursor-pointer"
                >
                  <Upload className="w-6 h-6 text-slate-400 mx-auto mb-2" />
                  <p className="text-sm font-semibold text-slate-700">
                    Klik tombol Format Dapodik atau Format Emis di atas untuk memilih file dari laptop/HP
                  </p>
                  <p className="text-xs text-slate-400 mt-1">
                    Mendukung file unduhan asli .xlsx, .xls, atau .csv
                  </p>
                </div>
              )}

              <div className="pt-2 flex items-center justify-between">
                <button
                  type="button"
                  onClick={() => setWizardStep(1)}
                  className="px-5 py-2.5 rounded-xl border border-slate-200 text-sm font-semibold text-slate-700 hover:bg-slate-50 flex items-center gap-1.5 cursor-pointer"
                >
                  <ArrowLeft className="w-4 h-4" /> Kembali
                </button>
                <button
                  type="button"
                  onClick={handleFinishWizard}
                  className="px-6 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold shadow-sm cursor-pointer"
                >
                  Proses Administrasi
                </button>
              </div>
            </div>
          )}

          {wizardStep === 3 && (
            <div className="p-12 text-center space-y-3">
              <div className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full animate-spin mx-auto"></div>
              <h3 className="text-lg font-bold text-slate-900">Menyiapkan Perangkat Administrasi...</h3>
              <p className="text-sm text-slate-500">
                Menyusun CP, ATP, Modul Ajar, dan Daftar Siswa untuk {wizForm.mapel} Kelas {wizForm.kelas} ({wizForm.fase})
              </p>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ============================================================================
  // TAMPILAN 3: DASHBOARD UTAMA (DENGAN EKSPOR MICROSOFT WORD .DOC)
  // ============================================================================
  const menuItems = [
    { id: 'beranda', label: 'Beranda', icon: LayoutDashboard, group: 'UTAMA' },
    { id: 'cp-tp', label: 'CP, TP, & ATP', icon: FileText, group: 'PERANGKAT AJAR' },
    { id: 'prota-promes', label: 'Prota & Promes', icon: Calendar, group: 'PERANGKAT AJAR' },
    { id: 'modul-ajar', label: 'Modul Ajar / RPP', icon: BookMarked, group: 'PERANGKAT AJAR', badge: activeClass.modulAjar.length },
    { id: 'data-siswa', label: 'Data Siswa', icon: Users, group: 'KBM & PENILAIAN' },
    { id: 'jurnal-presensi', label: 'Jurnal & Presensi', icon: CheckCircle2, group: 'KBM & PENILAIAN' },
    { id: 'daftar-nilai', label: 'Daftar Nilai (e-Rapor)', icon: Award, group: 'KBM & PENILAIAN' },
    { id: 'pengaturan', label: 'Pengaturan & Profil', icon: Settings, group: 'SISTEM' }
  ];

  const totalJp = activeClass.tps.reduce((acc, item) => acc + (Number(item.jp) || 0), 0);
  const selectedRombelFaseObj = FASE_LIST.find((f) => f.label === rombelForm.fase) || FASE_LIST[3];

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col md:flex-row font-sans text-slate-800">
      {toast && (
        <div className="fixed bottom-5 right-5 z-50 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-lg flex items-center gap-2.5 text-sm">
          <Check className="w-4 h-4 text-emerald-400" />
          <span>{toast}</span>
        </div>
      )}

      {/* SIDEBAR KIRI */}
      <aside className="w-full md:w-64 bg-white border-r border-slate-200 flex-shrink-0 flex flex-col">
        <div className="p-4 border-b border-slate-100 flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-sm">
            <BookOpen className="w-5 h-5" />
          </div>
          <div>
            <h1 className="font-bold text-lg tracking-tight text-slate-900 leading-none">EduAdmin</h1>
            <span className="text-[11px] text-slate-400 font-medium">Semua Jenjang (Fase A - F)</span>
          </div>
        </div>

        {/* KARTU KELAS AKTIF */}
        <div className="p-4 border-b border-slate-100">
          <div className="flex items-center justify-between text-[11px] font-semibold text-slate-400 uppercase tracking-wider mb-1">
            <span>KELAS AKTIF • {activeClass.fase.split(' ')[0]} {activeClass.fase.split(' ')[1]}</span>
            <button
              onClick={() => setActiveMenu('pengaturan')}
              className="text-slate-400 hover:text-blue-600 transition"
              title="Pengaturan Kelas"
            >
              <Settings className="w-3.5 h-3.5" />
            </button>
          </div>
          <div className="font-bold text-sm text-slate-800 truncate">{activeClass.nama}</div>
          <div className="text-xs text-slate-500 mt-0.5">
            {data.profile.semester} • {data.profile.tahunAjaran}
          </div>

          <div className="mt-3 flex flex-wrap gap-1.5">
            {data.classes.map((cls) => (
              <button
                key={cls.id}
                onClick={() => {
                  setData((prev) => ({ ...prev, activeClassId: cls.id }));
                  showNotification(`Beralih ke ${cls.nama} (${cls.fase})`);
                }}
                className={`px-2.5 py-1.5 rounded-lg text-xs font-semibold transition ${
                  cls.id === data.activeClassId
                    ? 'bg-blue-600 text-white shadow-sm'
                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                }`}
              >
                Kls {cls.kelas}
              </button>
            ))}
            <button
              onClick={() => setShowAddRombel(true)}
              className="px-2.5 py-1.5 rounded-lg text-xs font-medium border border-slate-200 text-slate-600 hover:bg-slate-50 flex items-center gap-1"
            >
              <Plus className="w-3.5 h-3.5" /> Rombel
            </button>
          </div>
        </div>

        {/* NAVIGASI MENU */}
        <nav className="p-3 flex-1 space-y-5 overflow-y-auto">
          {['UTAMA', 'PERANGKAT AJAR', 'KBM & PENILAIAN', 'SISTEM'].map((group) => (
            <div key={group}>
              <div className="px-3 mb-1.5 text-[10px] font-bold text-slate-400 tracking-wider uppercase">
                {group}
              </div>
              <div className="space-y-1">
                {menuItems
                  .filter((m) => m.group === group)
                  .map((item) => {
                    const Icon = item.icon;
                    const isActive = activeMenu === item.id;
                    return (
                      <button
                        key={item.id}
                        onClick={() => setActiveMenu(item.id)}
                        className={`w-full flex items-center justify-between px-3 py-2 rounded-xl text-sm font-medium transition ${
                          isActive
                            ? 'bg-blue-50 text-blue-600 font-semibold'
                            : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
                        }`}
                      >
                        <div className="flex items-center gap-2.5">
                          <Icon className={`w-4 h-4 ${isActive ? 'text-blue-600' : 'text-slate-400'}`} />
                          <span>{item.label}</span>
                        </div>
                        {item.badge && (
                          <span className="bg-rose-100 text-rose-600 text-xs font-bold px-2 py-0.5 rounded-full">
                            {item.badge}
                          </span>
                        )}
                      </button>
                    );
                  })}
              </div>
            </div>
          ))}
        </nav>

        <div className="p-3 border-t border-slate-100">
          <button
            onClick={() => {
              setWizardStep(1);
              setUploadedFile(null);
              setParsedStudents(null);
              setData((prev) => ({ ...prev, wizardCompleted: false }));
            }}
            className="w-full py-2 px-3 rounded-xl text-xs font-semibold text-slate-600 hover:bg-slate-100 flex items-center justify-center gap-1.5"
          >
            <Upload className="w-3.5 h-3.5" /> Impor Ulang Dapodik / Emis
          </button>
        </div>
      </aside>

      {/* AREA KONTEN UTAMA */}
      <div className="flex-1 flex flex-col min-w-0">
        <header className="bg-white border-b border-slate-200 px-6 py-3.5 flex items-center justify-between sticky top-0 z-10">
          <div className="flex items-center gap-2 text-sm text-slate-500">
            <span>EduAdmin</span>
            <ChevronRight className="w-4 h-4 text-slate-300" />
            <span className="font-semibold text-slate-800">
              {menuItems.find((m) => m.id === activeMenu)?.label}
            </span>
          </div>
          <div className="flex items-center gap-3 text-right">
            <div>
              <div className="text-sm font-bold text-slate-800">{data.profile.nama}</div>
              <div className="text-xs text-slate-400">NIP. {data.profile.nip}</div>
            </div>
            <button
              onClick={() => setData((prev) => ({ ...prev, isLoggedIn: false }))}
              title="Keluar"
              className="w-9 h-9 rounded-full bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-600 hover:text-rose-600"
            >
              <LogOut className="w-4 h-4" />
            </button>
          </div>
        </header>

        <main className="p-6 max-w-6xl w-full mx-auto space-y-6">
          {/* 1. BERANDA */}
          {activeMenu === 'beranda' && (
            <div className="space-y-6">
              <div className="bg-gradient-to-r from-blue-600 to-indigo-600 rounded-2xl p-6 text-white shadow-sm">
                <h2 className="text-xl font-bold">Selamat Datang, {data.profile.nama}!</h2>
                <p className="text-blue-100 text-sm mt-1">
                  {data.profile.sekolah} • {activeClass.nama} • {activeClass.fase}
                </p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                  <div className="text-xs font-semibold text-slate-400 uppercase">Tujuan Pembelajaran</div>
                  <div className="text-2xl font-bold text-slate-800 mt-1">{activeClass.tps.length} TP</div>
                  <div className="text-xs text-emerald-600 mt-1 font-medium">
                    {activeClass.tps.filter((t) => t.status === 'SELESAI').length} Terlaksana • Total {totalJp} JP
                  </div>
                </div>
                <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                  <div className="text-xs font-semibold text-slate-400 uppercase">Modul Ajar</div>
                  <div className="text-2xl font-bold text-slate-800 mt-1">{activeClass.modulAjar.length} Modul</div>
                  <div className="text-xs text-blue-600 mt-1 font-medium">Siap unduh format Word (.doc)</div>
                </div>
                <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                  <div className="text-xs font-semibold text-slate-400 uppercase">Jumlah Siswa</div>
                  <div className="text-2xl font-bold text-slate-800 mt-1">{activeClass.students.length} Siswa</div>
                  <div className="text-xs text-slate-500 mt-1">Aktif di {activeClass.nama}</div>
                </div>
                <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                  <div className="text-xs font-semibold text-slate-400 uppercase">Jurnal Mengajar</div>
                  <div className="text-2xl font-bold text-slate-800 mt-1">{activeClass.journals.length} Pertemuan</div>
                  <div className="text-xs text-emerald-600 mt-1 font-medium">Tersimpan otomatis</div>
                </div>
              </div>
            </div>
          )}

          {/* 2. CP, TP, & ATP */}
          {activeMenu === 'cp-tp' && (
            <div className="space-y-6">
              <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                <div className="flex items-start justify-between gap-4">
                  <div>
                    <h2 className="text-lg font-bold text-slate-900">Capaian Pembelajaran (CP)</h2>
                    <p className="text-sm text-slate-500 mt-0.5">
                      Mata Pelajaran: <strong className="text-slate-700">{activeClass.mapel}</strong> • Kelas{' '}
                      {activeClass.kelas} ({activeClass.fase})
                    </p>
                  </div>
                  <button
                    onClick={() => {
                      if (editingCp) {
                        updateActiveClass((cls) => ({ ...cls, cpText: cpDraft }));
                        setEditingCp(false);
                        showNotification('Teks Capaian Pembelajaran diperbarui!');
                      } else {
                        setCpDraft(activeClass.cpText);
                        setEditingCp(true);
                      }
                    }}
                    className="px-3 py-1.5 text-xs font-semibold rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50"
                  >
                    {editingCp ? 'Simpan Teks CP' : 'Edit Teks CP'}
                  </button>
                </div>

                {editingCp ? (
                  <textarea
                    rows={4}
                    value={cpDraft}
                    onChange={(e) => setCpDraft(e.target.value)}
                    className="mt-4 w-full rounded-xl border border-blue-300 p-3.5 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                ) : (
                  <div className="mt-4 bg-slate-50 border border-slate-200/80 rounded-xl p-4 text-sm text-slate-700 leading-relaxed">
                    {activeClass.cpText}
                  </div>
                )}
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h3 className="text-base font-bold text-slate-900">
                      Alur Tujuan Pembelajaran (ATP) - {data.profile.semester}
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">Total Alokasi Waktu: {totalJp} Jam Pelajaran (JP)</p>
                  </div>
                  <button
                    onClick={handleExportAtpWord}
                    className="px-3.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 shadow-xs cursor-pointer"
                  >
                    <FileDown className="w-3.5 h-3.5" /> Unduh Word (.doc)
                  </button>
                </div>

                <div className="space-y-3">
                  {activeClass.tps.map((tp) => (
                    <div
                      key={tp.id}
                      className="border border-slate-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4 hover:border-blue-200 transition"
                    >
                      <div className="flex items-start gap-3.5">
                        <div className="px-2.5 py-1.5 rounded-lg bg-blue-50 text-blue-600 font-bold text-xs text-center min-w-[56px]">
                          {tp.kode}
                        </div>
                        <div>
                          <div className="font-semibold text-sm text-slate-900">{tp.judul}</div>
                          <div className="text-xs text-slate-500 mt-1 flex flex-wrap items-center gap-2">
                            <span>Alokasi Waktu: {tp.jp} Jam Pelajaran (JP)</span>
                            <span>•</span>
                            <span className="text-blue-600 font-medium">Elemen: {tp.elemen}</span>
                          </div>
                        </div>
                      </div>

                      <div className="flex items-center gap-2.5 self-end sm:self-center">
                        <button
                          onClick={() =>
                            updateActiveClass((cls) => ({
                              ...cls,
                              tps: cls.tps.map((item) =>
                                item.id === tp.id
                                  ? { ...item, status: item.status === 'SELESAI' ? 'DRAFT' : 'SELESAI' }
                                  : item
                              )
                            }))
                          }
                          className={`px-2.5 py-1 rounded-md text-[11px] font-bold tracking-wide cursor-pointer ${
                            tp.status === 'SELESAI'
                              ? 'bg-emerald-100 text-emerald-700'
                              : 'bg-slate-100 text-slate-600'
                          }`}
                        >
                          {tp.status}
                        </button>
                        <button
                          onClick={() => setEditingTp(tp)}
                          className="p-1.5 rounded-lg text-slate-400 hover:text-blue-600 hover:bg-blue-50 transition"
                        >
                          <Settings className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>

                {!showAddTp ? (
                  <button
                    onClick={() => setShowAddTp(true)}
                    className="mt-4 w-full py-3 rounded-xl border border-dashed border-blue-300 text-blue-600 font-semibold text-sm hover:bg-blue-50/60 transition flex items-center justify-center gap-1.5"
                  >
                    <Plus className="w-4 h-4" /> Tambah Tujuan Pembelajaran Baru
                  </button>
                ) : (
                  <form onSubmit={handleAddTp} className="mt-4 p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
                    <div className="font-semibold text-sm text-slate-800">Tambah Tujuan Pembelajaran Baru</div>
                    <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                      <input
                        type="text"
                        placeholder="Kode (mis. TP 2.5)"
                        value={newTp.kode}
                        onChange={(e) => setNewTp({ ...newTp, kode: e.target.value })}
                        className="rounded-lg border border-slate-300 px-3 py-2 text-sm bg-white"
                        required
                      />
                      <input
                        type="text"
                        placeholder="Deskripsi Tujuan Pembelajaran..."
                        value={newTp.judul}
                        onChange={(e) => setNewTp({ ...newTp, judul: e.target.value })}
                        className="sm:col-span-2 rounded-lg border border-slate-300 px-3 py-2 text-sm bg-white"
                        required
                      />
                      <input
                        type="number"
                        placeholder="Alokasi JP"
                        value={newTp.jp}
                        onChange={(e) => setNewTp({ ...newTp, jp: e.target.value })}
                        className="rounded-lg border border-slate-300 px-3 py-2 text-sm bg-white"
                        min={1}
                      />
                    </div>
                    <div className="flex items-center justify-between gap-3">
                      <input
                        type="text"
                        placeholder="Elemen CP..."
                        value={newTp.elemen}
                        onChange={(e) => setNewTp({ ...newTp, elemen: e.target.value })}
                        className="rounded-lg border border-slate-300 px-3 py-2 text-sm bg-white flex-1"
                      />
                      <div className="flex gap-2">
                        <button
                          type="button"
                          onClick={() => setShowAddTp(false)}
                          className="px-3 py-1.5 rounded-lg text-sm text-slate-600 hover:bg-slate-200"
                        >
                          Batal
                        </button>
                        <button
                          type="submit"
                          className="px-4 py-1.5 rounded-lg bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700"
                        >
                          Simpan TP
                        </button>
                      </div>
                    </div>
                  </form>
                )}
              </div>
            </div>
          )}

          {/* 3. PROTA & PROMES */}
          {activeMenu === 'prota-promes' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Program Tahunan & Program Semester (Promes)</h2>
                  <p className="text-xs text-slate-500">
                    {activeClass.nama} • {activeClass.fase} • {data.profile.semester} {data.profile.tahunAjaran}
                  </p>
                </div>
                <button
                  onClick={handleExportPromesWord}
                  className="px-3.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 shadow-xs cursor-pointer"
                >
                  <FileDown className="w-3.5 h-3.5" /> Unduh Word (.doc)
                </button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-y border-slate-200 text-xs uppercase text-slate-500">
                      <th className="py-3 px-4">Kode</th>
                      <th className="py-3 px-4">Tujuan Pembelajaran / Materi Pokok</th>
                      <th className="py-3 px-4 text-center">JP</th>
                      <th className="py-3 px-4 text-center">Jan</th>
                      <th className="py-3 px-4 text-center">Feb</th>
                      <th className="py-3 px-4 text-center">Mar</th>
                      <th className="py-3 px-4 text-center">Apr</th>
                      <th className="py-3 px-4 text-center">Mei</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {activeClass.tps.map((tp, idx) => (
                      <tr key={tp.id} className="hover:bg-slate-50/70">
                        <td className="py-3 px-4 font-bold text-blue-600">{tp.kode}</td>
                        <td className="py-3 px-4 font-medium text-slate-800">{tp.judul}</td>
                        <td className="py-3 px-4 text-center font-semibold">{tp.jp} JP</td>
                        <td className="py-3 px-4 text-center">{idx === 0 ? '✅' : '-'}</td>
                        <td className="py-3 px-4 text-center">{idx === 0 || idx === 1 ? '✅' : '-'}</td>
                        <td className="py-3 px-4 text-center">{idx === 1 || idx === 2 ? '✅' : '-'}</td>
                        <td className="py-3 px-4 text-center">{idx === 2 || idx === 3 ? '✅' : '-'}</td>
                        <td className="py-3 px-4 text-center">{idx >= 3 ? '✅' : '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 4. MODUL AJAR / RPP */}
          {activeMenu === 'modul-ajar' && (
            <div className="space-y-4">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Modul Ajar Kurikulum Merdeka ({activeClass.fase})</h2>
                <p className="text-xs text-slate-500">Klik tombol di bawah kartu untuk mengunduh dokumen Microsoft Word (.doc) siap edit.</p>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {activeClass.modulAjar.map((modul) => (
                  <div key={modul.id} className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col justify-between">
                    <div>
                      <div className="flex items-center justify-between">
                        <span className="px-2.5 py-1 rounded-md bg-blue-50 text-blue-600 text-xs font-bold">
                          {modul.tpKode}
                        </span>
                        <span className="text-xs font-medium text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded">
                          {modul.status}
                        </span>
                      </div>
                      <h3 className="font-bold text-slate-900 mt-3 text-base">{modul.judul}</h3>
                      <div className="mt-3 space-y-1 text-xs text-slate-500">
                        <div className="flex items-center gap-1.5">
                          <Clock className="w-3.5 h-3.5" /> {modul.pertemuan}
                        </div>
                        <div>Model: {modul.model}</div>
                      </div>
                    </div>
                    <button
                      onClick={() => handleExportModulWord(modul)}
                      className="mt-5 w-full py-2.5 rounded-xl bg-blue-600 text-xs font-semibold text-white hover:bg-blue-700 flex items-center justify-center gap-1.5 cursor-pointer"
                    >
                      <FileDown className="w-3.5 h-3.5" /> Unduh Word (.doc)
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* 5. DATA SISWA */}
          {activeMenu === 'data-siswa' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-5">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Daftar Peserta Didik ({activeClass.nama})</h2>
                <p className="text-xs text-slate-500">Total: {activeClass.students.length} Siswa Aktif</p>
              </div>

              <form onSubmit={handleAddStudent} className="grid grid-cols-1 sm:grid-cols-4 gap-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                <input
                  type="text"
                  placeholder="NISN (opsional)"
                  value={newStudentNisn}
                  onChange={(e) => setNewStudentNisn(e.target.value)}
                  className="rounded-lg border border-slate-300 px-3 py-2 text-sm bg-white"
                />
                <input
                  type="text"
                  placeholder="Nama Lengkap Siswa..."
                  value={newStudentName}
                  onChange={(e) => setNewStudentName(e.target.value)}
                  className="rounded-lg border border-slate-300 px-3 py-2 text-sm bg-white"
                  required
                />
                <select
                  value={newStudentJk}
                  onChange={(e) => setNewStudentJk(e.target.value)}
                  className="rounded-lg border border-slate-300 px-3 py-2 text-sm bg-white"
                >
                  <option value="L">Laki-laki (L)</option>
                  <option value="P">Perempuan (P)</option>
                </select>
                <button
                  type="submit"
                  className="bg-blue-600 text-white font-semibold text-sm rounded-lg px-4 py-2 hover:bg-blue-700 flex items-center justify-center gap-1.5"
                >
                  <Plus className="w-4 h-4" /> Tambah Siswa
                </button>
              </form>

              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left">
                  <thead>
                    <tr className="border-b border-slate-200 text-xs text-slate-400 uppercase">
                      <th className="py-3 px-3">No</th>
                      <th className="py-3 px-3">NISN</th>
                      <th className="py-3 px-3">Nama Siswa</th>
                      <th className="py-3 px-3 text-center">L/P</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {activeClass.students.map((s, index) => (
                      <tr key={s.id} className="hover:bg-slate-50">
                        <td className="py-3 px-3 text-slate-500">{index + 1}</td>
                        <td className="py-3 px-3 font-mono text-xs text-slate-600">{s.nisn}</td>
                        <td className="py-3 px-3 font-semibold text-slate-800">{s.nama}</td>
                        <td className="py-3 px-3 text-center">{s.jk}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 6. JURNAL & PRESENSI */}
          {activeMenu === 'jurnal-presensi' && (
            <div className="space-y-6">
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                <form
                  onSubmit={handleSaveJournal}
                  className="lg:col-span-5 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4 h-fit"
                >
                  <h2 className="text-base font-bold text-slate-900">Jurnal Mengajar ({activeClass.nama})</h2>
                  <div>
                    <label className="block text-xs font-semibold text-slate-600 mb-1">Tanggal Pertemuan</label>
                    <input
                      type="date"
                      value={journalForm.tanggal}
                      onChange={(e) => setJournalForm({ ...journalForm, tanggal: e.target.value })}
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-600 mb-1">Jam Pelajaran Ke-</label>
                    <input
                      type="text"
                      placeholder="Contoh: 1-2"
                      value={journalForm.jamKe}
                      onChange={(e) => setJournalForm({ ...journalForm, jamKe: e.target.value })}
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-600 mb-1">Tujuan Pembelajaran (TP)</label>
                    <select
                      value={journalForm.tpKode}
                      onChange={(e) => setJournalForm({ ...journalForm, tpKode: e.target.value })}
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm bg-white"
                    >
                      {activeClass.tps.map((tp) => (
                        <option key={tp.id} value={`${tp.kode} - ${tp.judul}`}>
                          {tp.kode} - {tp.judul}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-600 mb-1">Catatan / Kendala (Opsional)</label>
                    <textarea
                      rows={3}
                      placeholder="Tambahkan catatan khusus..."
                      value={journalForm.catatan}
                      onChange={(e) => setJournalForm({ ...journalForm, catatan: e.target.value })}
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                    />
                  </div>
                  <button
                    type="submit"
                    className="w-full py-2.5 rounded-xl bg-blue-600 text-white font-semibold text-sm hover:bg-blue-700 transition shadow-sm"
                  >
                    Simpan Jurnal
                  </button>
                </form>

                <div className="lg:col-span-7 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                  <div className="flex items-center justify-between mb-4">
                    <div>
                      <h2 className="text-base font-bold text-slate-900">Daftar Hadir Siswa</h2>
                      <p className="text-xs text-slate-500">Tandai siswa yang tidak hadir (H/S/I/A).</p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={handleExportPresensiWord}
                        className="px-3 py-1.5 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1 cursor-pointer"
                      >
                        <FileDown className="w-3.5 h-3.5" /> Unduh Word
                      </button>
                      <span className="px-3 py-1.5 rounded-lg bg-slate-100 text-xs font-bold text-slate-700">
                        Total: {activeClass.students.length}
                      </span>
                    </div>
                  </div>

                  <div className="divide-y divide-slate-100">
                    {activeClass.students.map((s, idx) => (
                      <div key={s.id} className="py-3 flex items-center justify-between gap-2">
                        <span className="text-sm font-semibold text-slate-800">
                          {idx + 1}. {s.nama}
                        </span>
                        <div className="flex gap-1.5">
                          {[
                            { code: 'H', activeClass: 'bg-emerald-100 text-emerald-700 border-emerald-300' },
                            { code: 'S', activeClass: 'bg-blue-100 text-blue-700 border-blue-300' },
                            { code: 'I', activeClass: 'bg-amber-100 text-amber-700 border-amber-300' },
                            { code: 'A', activeClass: 'bg-rose-100 text-rose-700 border-rose-300' }
                          ].map((btn) => (
                            <button
                              key={btn.code}
                              type="button"
                              onClick={() => handleStatusHadir(s.id, btn.code)}
                              className={`w-8 h-8 rounded-lg text-xs font-bold border transition ${
                                s.statusHadir === btn.code
                                  ? btn.activeClass
                                  : 'bg-white text-slate-400 border-slate-200 hover:bg-slate-50'
                              }`}
                            >
                              {btn.code}
                            </button>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                <h3 className="text-base font-bold text-slate-900 mb-3">Riwayat Jurnal Mengajar ({activeClass.nama})</h3>
                <div className="space-y-3">
                  {activeClass.journals.map((j) => (
                    <div key={j.id} className="p-4 rounded-xl bg-slate-50 border border-slate-200/80 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                      <div>
                        <div className="text-xs font-bold text-blue-600">
                          {j.tanggal} • Jam ke-{j.jamKe}
                        </div>
                        <div className="font-semibold text-sm text-slate-800 mt-0.5">{j.tpKode}</div>
                        <div className="text-xs text-slate-500 mt-1">{j.catatan}</div>
                      </div>
                      <div className="text-xs font-semibold text-emerald-700 bg-emerald-100 px-3 py-1 rounded-full self-start sm:self-center">
                        Hadir: {j.hadirCount} | Tidak Hadir: {j.tidakHadirCount}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* 7. DAFTAR NILAI (e-RAPOR) */}
          {activeMenu === 'daftar-nilai' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Daftar Nilai Formatif & Sumatif ({activeClass.nama})</h2>
                  <p className="text-xs text-slate-500">Klik angka nilai untuk mengedit langsung; Nilai Akhir dihitung otomatis.</p>
                </div>
                <button
                  onClick={handleExportNilaiWord}
                  className="px-3.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 shadow-xs cursor-pointer"
                >
                  <FileDown className="w-3.5 h-3.5" /> Unduh Word (.doc)
                </button>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left">
                  <thead>
                    <tr className="bg-slate-50 border-y border-slate-200 text-xs uppercase text-slate-500">
                      <th className="py-3 px-3">No</th>
                      <th className="py-3 px-3">Nama Siswa</th>
                      <th className="py-3 px-3 text-center">Formatif 1</th>
                      <th className="py-3 px-3 text-center">Formatif 2</th>
                      <th className="py-3 px-3 text-center">Sumatif Lingkup</th>
                      <th className="py-3 px-3 text-center">Nilai Rapor</th>
                      <th className="py-3 px-3 text-center">Predikat</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {activeClass.students.map((s, idx) => {
                      const akhir = Math.round((Number(s.f1) + Number(s.f2) + Number(s.sumatif)) / 3);
                      const predikat = akhir >= 90 ? 'A (Sangat Baik)' : akhir >= 80 ? 'B (Baik)' : 'C (Cukup)';
                      return (
                        <tr key={s.id} className="hover:bg-slate-50/70">
                          <td className="py-2.5 px-3 text-slate-500">{idx + 1}</td>
                          <td className="py-2.5 px-3 font-semibold text-slate-800">{s.nama}</td>
                          <td className="py-2.5 px-3 text-center">
                            <input
                              type="number"
                              value={s.f1}
                              onChange={(e) => handleGradeChange(s.id, 'f1', e.target.value)}
                              className="w-16 text-center rounded-lg border border-slate-200 py-1 text-sm"
                            />
                          </td>
                          <td className="py-2.5 px-3 text-center">
                            <input
                              type="number"
                              value={s.f2}
                              onChange={(e) => handleGradeChange(s.id, 'f2', e.target.value)}
                              className="w-16 text-center rounded-lg border border-slate-200 py-1 text-sm"
                            />
                          </td>
                          <td className="py-2.5 px-3 text-center">
                            <input
                              type="number"
                              value={s.sumatif}
                              onChange={(e) => handleGradeChange(s.id, 'sumatif', e.target.value)}
                              className="w-16 text-center rounded-lg border border-slate-200 py-1 text-sm"
                            />
                          </td>
                          <td className="py-2.5 px-3 text-center font-bold text-blue-600">{akhir}</td>
                          <td className="py-2.5 px-3 text-center">
                            <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700">
                              {predikat}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 8. PENGATURAN & PROFIL */}
          {activeMenu === 'pengaturan' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm max-w-2xl space-y-5">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Pengaturan Identitas Guru & Sistem</h2>
                <p className="text-xs text-slate-500">Identitas ini otomatis muncul pada kop dan tanda tangan file Word (.doc).</p>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Nama Lengkap & Gelar</label>
                  <input
                    type="text"
                    value={data.profile.nama}
                    onChange={(e) =>
                      setData((prev) => ({ ...prev, profile: { ...prev.profile, nama: e.target.value } }))
                    }
                    className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">NIP / NUPTK</label>
                  <input
                    type="text"
                    value={data.profile.nip}
                    onChange={(e) =>
                      setData((prev) => ({ ...prev, profile: { ...prev.profile, nip: e.target.value } }))
                    }
                    className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Nama Satuan Pendidikan</label>
                  <input
                    type="text"
                    value={data.profile.sekolah}
                    onChange={(e) =>
                      setData((prev) => ({ ...prev, profile: { ...prev.profile, sekolah: e.target.value } }))
                    }
                    className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Tahun Ajaran</label>
                  <input
                    type="text"
                    value={data.profile.tahunAjaran}
                    onChange={(e) =>
                      setData((prev) => ({ ...prev, profile: { ...prev.profile, tahunAjaran: e.target.value } }))
                    }
                    className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                  />
                </div>
              </div>

              <div className="pt-4 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
                <button
                  type="button"
                  onClick={() => showNotification('Identitas berhasil disimpan!')}
                  className="px-4 py-2 rounded-xl bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700 flex items-center gap-1.5"
                >
                  <Save className="w-4 h-4" /> Simpan Pengaturan
                </button>
                <button
                  type="button"
                  onClick={handleResetAll}
                  className="px-3 py-2 rounded-xl border border-rose-200 text-rose-600 text-xs font-semibold hover:bg-rose-50 flex items-center gap-1.5"
                >
                  <RotateCcw className="w-4 h-4" /> Reset Awal & Ulangi Wizard
                </button>
              </div>
            </div>
          )}
        </main>
      </div>

      {/* MODAL EDIT TP */}
      {editingTp && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <form onSubmit={handleUpdateTp} className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="font-bold text-slate-900">Edit Tujuan Pembelajaran</h3>
              <button type="button" onClick={() => setEditingTp(null)} className="text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">Kode TP</label>
              <input
                type="text"
                value={editingTp.kode}
                onChange={(e) => setEditingTp({ ...editingTp, kode: e.target.value })}
                className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm"
                required
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">Deskripsi Tujuan Pembelajaran</label>
              <textarea
                rows={3}
                value={editingTp.judul}
                onChange={(e) => setEditingTp({ ...editingTp, judul: e.target.value })}
                className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm"
                required
              />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Alokasi JP</label>
                <input
                  type="number"
                  value={editingTp.jp}
                  onChange={(e) => setEditingTp({ ...editingTp, jp: Number(e.target.value) })}
                  className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm"
                  min={1}
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Status</label>
                <select
                  value={editingTp.status}
                  onChange={(e) => setEditingTp({ ...editingTp, status: e.target.value })}
                  className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm bg-white"
                >
                  <option value="DRAFT">DRAFT</option>
                  <option value="SELESAI">SELESAI</option>
                </select>
              </div>
            </div>
            <div className="pt-2 flex items-center justify-between">
              <button
                type="button"
                onClick={() => handleDeleteTp(editingTp.id)}
                className="px-3 py-2 rounded-xl text-rose-600 text-xs font-semibold hover:bg-rose-50 flex items-center gap-1"
              >
                <Trash2 className="w-4 h-4" /> Hapus TP
              </button>
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => setEditingTp(null)}
                  className="px-4 py-2 rounded-xl text-sm text-slate-600 hover:bg-slate-100"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 rounded-xl bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700"
                >
                  Simpan Perubahan
                </button>
              </div>
            </div>
          </form>
        </div>
      )}

      {/* MODAL TAMBAH ROMBEL SEMUA TINGKATAN & MAPEL */}
      {showAddRombel && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <form onSubmit={handleAddRombel} className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="font-bold text-slate-900">Tambah Rombel / Tingkatan Baru</h3>
              <button type="button" onClick={() => setShowAddRombel(false)} className="text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">Pilih Jenjang / Fase Kurikulum</label>
              <select
                value={rombelForm.fase}
                onChange={(e) => {
                  const found = FASE_LIST.find((f) => f.label === e.target.value) || FASE_LIST[3];
                  setRombelForm({
                    ...rombelForm,
                    fase: found.label,
                    kelas: found.defaultKelas,
                    mapel: found.mapels[0]
                  });
                }}
                className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm bg-white"
              >
                {FASE_LIST.map((f) => (
                  <option key={f.label} value={f.label}>
                    {f.label}
                  </option>
                ))}
              </select>
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Tingkat Kelas (1-12)</label>
                <select
                  value={rombelForm.kelas}
                  onChange={(e) => setRombelForm({ ...rombelForm, kelas: e.target.value })}
                  className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm bg-white"
                >
                  {['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'].map((k) => (
                    <option key={k} value={k}>
                      Kelas {k}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Nama Rombel</label>
                <input
                  type="text"
                  placeholder="Contoh: A / B / IPA-1"
                  value={rombelForm.rombel}
                  onChange={(e) => setRombelForm({ ...rombelForm, rombel: e.target.value })}
                  className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                  required
                />
              </div>
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">Mata Pelajaran</label>
              <select
                value={rombelForm.mapel}
                onChange={(e) => setRombelForm({ ...rombelForm, mapel: e.target.value })}
                className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm bg-white"
              >
                {selectedRombelFaseObj.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setShowAddRombel(false)}
                className="px-4 py-2 rounded-xl text-sm text-slate-600 hover:bg-slate-100"
              >
                Batal
              </button>
              <button
                type="submit"
                className="px-4 py-2 rounded-xl bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700"
              >
                Buat Rombel
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
}
