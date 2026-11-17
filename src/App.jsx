import React, { useState, useEffect, useRef } from 'react';
import {
  BookOpen,
  LayoutDashboard,
  FileText,
  Calendar,
  CalendarRange,
  ListChecks,
  BookMarked,
  ClipboardList,
  HelpCircle,
  FileCheck2,
  GraduationCap,
  Users,
  CheckCircle2,
  Award,
  FileSpreadsheet,
  FolderCheck,
  Settings,
  Plus,
  FileDown,
  Trash2,
  Save,
  RotateCcw,
  ChevronRight,
  Clock,
  Check,
  X,
  Upload,
  ArrowRight,
  ArrowLeft,
  LogOut,
  ExternalLink,
  Layers
} from 'lucide-react';

// GANTI TAUTAN INI DENGAN LINK PRODUK LYNK.ID ANDA
const LYNK_ID_URL = 'https://lynk.id';

// --- FUNGSI EKSPOR DOKUMEN MICROSOFT WORD (.DOC) ---
const downloadWordDocument = (filename, title, subtitle, profile, activeClass, bodyHtml, isModulAjar = false) => {
  const kopSekolah = profile.headerKop
    ? `<div style="text-align:center;font-size:10pt;color:#334155;margin-bottom:2px;">${profile.headerKop}</div>`
    : '';

  const headerBlock = isModulAjar
    ? `
      ${kopSekolah}
      <div class="doc-main-title">${title}</div>
      <div class="doc-topic-title">${subtitle}</div>
    `
    : `
      ${kopSekolah}
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
    `;

  const fullHtml = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' 
          xmlns:w='urn:schemas-microsoft-com:office:word' 
          xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <meta charset='utf-8'>
      <title>${title}</title>
      <!--[if gte mso 9]>
      <xml>
        <w:WordDocument>
          <w:View>Print</w:View>
          <w:Zoom>100</w:Zoom>
          <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
      </xml>
      <![endif]-->
      <style>
        @page { size: A4; margin: 2cm; }
        body { font-family: 'Calibri', 'Arial', sans-serif; font-size: 11pt; line-height: 1.45; color: #111827; }
        .doc-main-title { font-size: 16pt; font-weight: bold; color: #1e3a8a; margin-bottom: 2px; }
        .doc-topic-title { font-size: 13pt; font-weight: bold; color: #1d4ed8; margin-bottom: 8px; border-bottom: 2px solid #1e3a8a; padding-bottom: 6px; }
        .header-title { text-align: center; font-size: 14pt; font-weight: bold; text-transform: uppercase; margin-bottom: 4px; }
        .header-sub { text-align: center; font-size: 12pt; font-weight: bold; margin-bottom: 16px; border-bottom: 2px solid #000; padding-bottom: 8px; }
        h2.sec-head { font-size: 12pt; font-weight: bold; color: #1e3a8a; margin-top: 16px; margin-bottom: 6px; border-bottom: 1px solid #cbd5e1; padding-bottom: 2px; }
        p { margin: 4px 0 8px 0; text-align: justify; }
        ul, ol { margin: 4px 0 10px 20px; padding: 0; }
        li { margin-bottom: 4px; }
        .meta-table { width: 100%; border-collapse: collapse; margin-bottom: 16px; font-size: 10.5pt; }
        .meta-table td { padding: 3px 4px; vertical-align: top; }
        .data-table { width: 100%; border-collapse: collapse; margin-top: 6px; margin-bottom: 12px; font-size: 10pt; }
        .data-table th, .data-table td { border: 1px solid #64748b; padding: 6px 8px; text-align: left; vertical-align: top; }
        .data-table th { background-color: #e2e8f0; font-weight: bold; color: #0f172a; }
        .box-highlight { border-left: 4px solid #2563eb; background-color: #f8fafc; padding: 8px 12px; margin: 8px 0 12px 0; }
        .ttd-table { width: 100%; border-collapse: collapse; margin-top: 32px; font-size: 10.5pt; page-break-inside: avoid; }
        .ttd-table td { width: 50%; text-align: center; vertical-align: top; }
      </style>
    </head>
    <body>
      ${headerBlock}
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

  const blob = new Blob(['\ufeff', fullHtml], { type: 'application/msword;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${filename}.doc`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

// --- GENERATOR MODUL AJAR MENDALAM BERBASIS TPACK (BAGIAN A - H) ---
const buildDeepLearningModulHtml = (tp, profile, activeClass) => {
  const mapel = activeClass.mapel;
  const kelas = activeClass.kelas;
  const faseLabel = activeClass.fase.split(' ')[1] || 'D';
  const elemen = tp.elemen || 'Pemahaman Konsep & Keterampilan Proses';
  const tpJudul = tp.judul;

  return `
    <p><em>Peserta didik mengeksplorasi konsep esensial <strong>${tpJudul}</strong> melalui aktivitas penemuan terbimbing berbantuan media digital interaktif, lalu mengaplikasikannya untuk memecahkan masalah kontekstual kehidupan sehari-hari.</em></p>

    <h2 class="sec-head">A. Identitas dan Capaian</h2>
    <table class="data-table">
      <thead>
        <tr>
          <th style="width:32%;">Komponen</th>
          <th style="width:68%;">Keterangan</th>
        </tr>
      </thead>
      <tbody>
        <tr><td><strong>Satuan Pendidikan</strong></td><td>${profile.sekolah}</td></tr>
        <tr><td><strong>Mata Pelajaran / Fase / Kelas</strong></td><td>${mapel} / Fase ${faseLabel} / Kelas ${kelas}</td></tr>
        <tr><td><strong>Elemen</strong></td><td>${elemen}</td></tr>
        <tr><td><strong>Alokasi Waktu</strong></td><td>${Math.max(1, Math.round(tp.jp / 3))} Pertemuan (${tp.jp} JP)</td></tr>
        <tr><td><strong>Capaian Pembelajaran</strong></td><td>${activeClass.cpText}</td></tr>
        <tr><td><strong>Dimensi Profil Lulusan</strong></td><td>Penalaran Kritis, Kolaborasi, Komunikasi, Kemandirian</td></tr>
        <tr><td><strong>Model dan Metode</strong></td><td>Problem-Based Learning (PBL), diskusi kelompok 4 orang, eksplorasi digital</td></tr>
        <tr><td><strong>Pengetahuan Prasyarat</strong></td><td>Pemahaman konsep dasar ${mapel} pada pertemuan sebelumnya serta literasi numerasi/bacaan dasar.</td></tr>
      </tbody>
    </table>

    <h2 class="sec-head">B. Tujuan Pembelajaran (${tp.kode})</h2>
    <p>Setelah mengikuti rangkaian pembelajaran, peserta didik mampu:</p>
    <ol>
      <li>Menjelaskan konsep dasar dan keterkaitan unsur utama pada <strong>${tpJudul}</strong> berdasarkan hasil eksplorasi.</li>
      <li>Menerapkan prinsip dan langkah penyelesaian masalah pada latihan terstruktur secara runtut dan tepat.</li>
      <li>Menganalisis serta menguji kebenaran solusi pada berbagai variasi kasus.</li>
      <li>Menyelesaikan masalah kontekstual kehidupan sehari-hari yang berkaitan dengan ${mapel}.</li>
    </ol>
    <div class="box-highlight">
      <p><strong>Pemahaman Bermakna:</strong> Penguasaan konsep ${mapel} menjadi alat berpikir sistematis untuk menganalisis dan memecahkan berbagai persoalan nyata di lingkungan sekitar secara akurat.</p>
      <p style="margin-bottom:2px;"><strong>Pertanyaan Pemantik:</strong></p>
      <ul style="margin-bottom:0;">
        <li>Mengapa pemahaman tentang ${tpJudul.toLowerCase()} sangat penting diterapkan dalam kehidupan nyata?</li>
        <li>Masalah apa di sekitar kita yang dapat diselesaikan secara lebih efektif setelah menguasai konsep ini?</li>
      </ul>
    </div>

    <h2 class="sec-head">C. Analisis TPACK</h2>
    <table class="data-table">
      <thead>
        <tr>
          <th style="width:18%;">Komponen</th>
          <th style="width:82%;">Penerapan dalam Modul Ini</th>
        </tr>
      </thead>
      <tbody>
        <tr><td><strong>CK</strong><br/><span style="font-size:9pt;color:#475569;">(Content Knowledge)</span></td><td>${tpJudul}. Miskonsepsi umum: menghafal rumus/definisi tanpa memahami syarat penerapannya pada kasus kontekstual.</td></tr>
        <tr><td><strong>PK</strong><br/><span style="font-size:9pt;color:#475569;">(Pedagogical Knowledge)</span></td><td>Problem-Based Learning (PBL) dengan masalah nyata, penemuan terbimbing, diskusi kelompok heterogen, umpan balik sejawat, dan asesmen formatif berkelanjutan.</td></tr>
        <tr><td><strong>TK</strong><br/><span style="font-size:9pt;color:#475569;">(Technological Knowledge)</span></td><td>Media simulasi interaktif (GeoGebra/PhET/teks digital), Padlet, Quizizz, proyektor; cadangan luring: kertas kerja cetak & alat peraga.</td></tr>
        <tr><td><strong>PCK</strong><br/><span style="font-size:9pt;color:#475569;">(Pedagogical Content)</span></td><td>Representasi visual/konkret sebelum simbol abstrak, tabel pola pengamatan untuk memunculkan konjektur, serta klarifikasi miskonsepsi.</td></tr>
        <tr><td><strong>TCK</strong><br/><span style="font-size:9pt;color:#475569;">(Technological Content)</span></td><td>Media digital interaktif menampilkan perubahan visual dan data seketika saat variabel diubah sehingga pola terlihat jelas.</td></tr>
        <tr><td><strong>TPK</strong><br/><span style="font-size:9pt;color:#475569;">(Technological Pedagogical)</span></td><td>Padlet untuk galeri solusi dan komentar sejawat; Quizizz untuk diagnosis awal dan kuis keluar cepat.</td></tr>
        <tr><td><strong>TPACK</strong><br/><span style="font-size:9pt;color:#475569;">(Integrasi Utuh)</span></td><td>Eksplorasi dinamis (TCK) dipandu penemuan terbimbing (PK) agar konsep ditemukan secara bermakna, bukan dihafal; teknologi dipakai karena memungkinkan banyak pengujian contoh dalam waktu singkat.</td></tr>
      </tbody>
    </table>

    <h2 class="sec-head">D. Persiapan</h2>
    <ul>
      <li><strong>Guru:</strong> Menyiapkan perangkat eksplorasi digital interaktif, tautan Padlet dan Quizizz, LKPD terstruktur, media demonstrasi konkret, serta lembar rubrik observasi.</li>
      <li><strong>Peserta Didik:</strong> Alat tulis, buku catatan, perlengkapan aktivitas kelompok, serta ponsel atau laptop kelompok (minimal 1 perangkat per kelompok).</li>
      <li><strong>Pengelompokan:</strong> Kelompok heterogen beranggotakan 4 orang (kemampuan campuran) untuk mendorong tutor sebaya.</li>
    </ul>

    <h2 class="sec-head">E. Langkah Pembelajaran (Pembelajaran Mendalam / Deep Learning)</h2>
    <table class="data-table">
      <thead>
        <tr>
          <th style="width:18%;">Tahap</th>
          <th style="width:12%;">Waktu</th>
          <th style="width:48%;">Kegiatan Pembelajaran</th>
          <th style="width:22%;">Peran Teknologi</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><strong>Pendahuluan</strong></td>
          <td>10 menit</td>
          <td>Salam dan doa. Guru melakukan demonstrasi pemantik kontekstual dan apersepsi cepat terkait pengetahuan prasyarat. Menyampaikan tujuan dan alur pelajaran.</td>
          <td>Kuis diagnostik 3 soal di Quizizz atau kertas</td>
        </tr>
        <tr>
          <td><strong>Memahami: Eksplorasi</strong></td>
          <td>20 menit</td>
          <td>Setiap kelompok mengeksplorasi media interaktif, mencatat hasil pengamatan ke dalam tabel LKPD, dan merumuskan pola/konjektur. Guru memberi pertanyaan penuntun sebelum memformalkan konsep inti.</td>
          <td>Media eksplorasi digital; cadangan: LKPD & alat peraga cetak</td>
        </tr>
        <tr>
          <td><strong>Mengaplikasi: Masalah PBL</strong></td>
          <td>30 menit</td>
          <td>Tiga stasiun tantangan kontekstual sesuai topik ${tp.kode}; kelompok memilih urutan pengerjaan sendiri. Guru berkeliling dengan pertanyaan penuntun, bukan jawaban langsung.</td>
          <td>Padlet: unggah foto penyelesaian kelompok</td>
        </tr>
        <tr>
          <td><strong>Merefleksi: Galeri</strong></td>
          <td>10 menit</td>
          <td>Galeri jalan: kelompok menilai dua solusi kelompok lain dengan satu bintang dan satu saran. Guru mengklarifikasi miskonsepsi yang muncul.</td>
          <td>Komentar sejawat di Padlet</td>
        </tr>
        <tr>
          <td><strong>Penutup</strong></td>
          <td>10 menit</td>
          <td>Refleksi individu (apa yang sudah dipahami, apa yang masih membingungkan, di mana penerapannya). Kuis keluar (exit ticket) 3 soal, tugas pengayaan, dan salam.</td>
          <td>Quizizz atau lembar refleksi</td>
        </tr>
      </tbody>
    </table>

    <h2 class="sec-head">F. Asesmen</h2>
    <p><strong>1. Bentuk Asesmen:</strong> Diagnostik (3 soal prasyarat di pendahuluan), Formatif (observasi diskusi, LKPD, galeri solusi), dan Kuis Keluar (<em>Exit Ticket</em>).</p>
    <p><strong>2. Rubrik Penilaian Proses Kelompok dan LKPD:</strong></p>
    <table class="data-table">
      <thead>
        <tr>
          <th style="width:20%;">Kriteria</th>
          <th style="width:20%;">Perlu Bimbingan</th>
          <th style="width:20%;">Berkembang</th>
          <th style="width:20%;">Cakap</th>
          <th style="width:20%;">Mahir</th>
        </tr>
      </thead>
      <tbody>
        <tr><td><strong>Penalaran dan Konjektur</strong></td><td>Belum merumuskan pola</td><td>Pola benar dengan bantuan guru</td><td>Merumuskan pola dari data sendiri</td><td>Merumuskan, menguji pada kasus berbeda, dan menjelaskan batasnya</td></tr>
        <tr><td><strong>Ketepatan Konsep</strong></td><td>Banyak kesalahan konsep</td><td>Konsep benar, kurang teliti</td><td>Benar dan runtut</td><td>Benar, runtut, dan memeriksa ulang jawaban</td></tr>
        <tr><td><strong>Penerapan Konteks</strong></td><td>Belum dapat memodelkan</td><td>Memodelkan dengan bantuan</td><td>Memodelkan mandiri</td><td>Memodelkan dan membuat soal konteks baru</td></tr>
        <tr><td><strong>Kolaborasi & Komunikasi</strong></td><td>Pasif</td><td>Berkontribusi bila diminta</td><td>Aktif dan menghargai pendapat</td><td>Memfasilitasi kelompok dan menjelaskan dengan jelas</td></tr>
      </tbody>
    </table>

    <h2 class="sec-head">G. Diferensiasi dan Pengayaan</h2>
    <ul>
      <li><strong>Dukungan:</strong> Kartu konsep dengan ilustrasi visual, soal bertahap dari tingkat dasar, serta pasangan sebaya.</li>
      <li><strong>Tantangan:</strong> Studi kasus multi-langkah tingkat lanjut dan pembuktian/analisis variasi kasus kompleks.</li>
      <li><strong>Produk Pilihan:</strong> Poster infografis, video penjelasan 1 menit, atau soal cerita buatan sendiri.</li>
      <li><strong>Tanpa Internet:</strong> Seluruh eksplorasi dapat dilakukan menggunakan LKPD cetak dan alat peraga manual.</li>
    </ul>

    <h2 class="sec-head">H. Refleksi dan Tindak Lanjut</h2>
    <ul>
      <li><strong>Guru:</strong> Apakah peserta didik menemukan pola sendiri atau menunggu penjelasan? Bagian mana yang paling lama? Apa yang diubah untuk kelas berikutnya?</li>
      <li><strong>Peserta Didik:</strong> Apa yang paling membantu saya paham hari ini? Bagian mana yang masih membingungkan?</li>
      <li><strong>Tindak Lanjut:</strong> Remedial terstruktur untuk yang belum tuntas; pengayaan studi kasus tingkat lanjut pada pertemuan berikutnya.</li>
    </ul>
  `;
};

// --- BANK DATA KURIKULUM MERDEKA LENGKAP (FASE A - F & MULTI MAPEL) ---
const SUBJECT_BANK = {
  'Matematika': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik dapat menjelaskan dan menggunakan Teorema Pythagoras untuk menyelesaikan masalah (termasuk jarak dua titik pada koordinat Kartesius), menyelesaikan sistem persamaan linear, menghitung luas permukaan dan volume bangun ruang, serta menganalisis data statistik dan peluang.',
    tps: [
      { kode: 'TP 8.1', judul: 'Teorema Pythagoras dan Penerapannya pada Masalah Kontekstual', jp: 12, status: 'SELESAI', elemen: 'Geometri', minggu: [1, 2, 3, 4] },
      { kode: 'TP 8.2', judul: 'Sistem Persamaan Linear Dua Variabel (SPLDV)', jp: 10, status: 'SELESAI', elemen: 'Aljabar', minggu: [5, 6, 7, 8] },
      { kode: 'TP 8.3', judul: 'Luas Permukaan dan Volume Bangun Ruang Sisi Datar', jp: 8, status: 'DRAFT', elemen: 'Pengukuran', minggu: [9, 10, 11, 12] },
      { kode: 'TP 8.4', judul: 'Statistika (Ukuran Pemusatan Data) dan Peluang Empirik', jp: 14, status: 'DRAFT', elemen: 'Analisis Data & Peluang', minggu: [13, 14, 15, 16] }
    ]
  },
  'Bahasa Inggris': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik menggunakan teks lisan, tulisan, dan visual dalam bahasa Inggris untuk berinteraksi dan berkomunikasi dalam konteks yang beragam (Recount, Narrative, Descriptive, Procedure) serta menyusun teks terstruktur.',
    tps: [
      { kode: 'TP 8.1', judul: 'Recount Text: Menceritakan Pengalaman Masa Lampau (Simple Past Tense)', jp: 12, status: 'SELESAI', elemen: 'Membaca - Memirsa', minggu: [1, 2, 3, 4] },
      { kode: 'TP 8.2', judul: 'Narrative Text: Menganalisis Alur & Nilai Moral Cerita Rakyat', jp: 10, status: 'SELESAI', elemen: 'Menyimak - Berbicara', minggu: [5, 6, 7, 8] },
      { kode: 'TP 8.3', judul: 'Short Functional Text: Notice, Caution, and Announcement', jp: 8, status: 'DRAFT', elemen: 'Menulis - Mempresentasikan', minggu: [9, 10, 11, 12] },
      { kode: 'TP 8.4', judul: 'Descriptive Text: Mendeskripsikan Tempat Wisata dan Tokoh', jp: 10, status: 'DRAFT', elemen: 'Menulis - Mempresentasikan', minggu: [13, 14, 15, 16] }
    ]
  },
  'Bahasa Indonesia': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik mampu memahami, mengolah, menginterpretasi, dan mengevaluasi informasi dari berbagai tipe teks fiksi dan informasional serta menyajikannya secara lisan dan tulis dengan kaidah kebahasaan yang tepat.',
    tps: [
      { kode: 'TP 8.1', judul: 'Menganalisis Informasi dan Struktur Teks Eksplanasi & Berita', jp: 12, status: 'SELESAI', elemen: 'Membaca dan Memirsa', minggu: [1, 2, 3, 4] },
      { kode: 'TP 8.2', judul: 'Menyimak dan Mengevaluasi Gagasan dari Teks Pidato Persuasif', jp: 10, status: 'SELESAI', elemen: 'Menyimak', minggu: [5, 6, 7, 8] },
      { kode: 'TP 8.3', judul: 'Menulis Teks Argumentasi dan Ulasan Karya Sastra', jp: 12, status: 'DRAFT', elemen: 'Menulis', minggu: [9, 10, 11, 12] },
      { kode: 'TP 8.4', judul: 'Mengapresiasi dan Mendeklamasikan Puisi Modern', jp: 10, status: 'DRAFT', elemen: 'Berbicara', minggu: [13, 14, 15, 16] }
    ]
  },
  'IPA (Ilmu Pengetahuan Alam)': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik mampu melakukan penyelidikan ilmiah, memahami sistem organ manusia, menganalisis gaya, gerak, energi, getaran dan gelombang, serta struktur bumi dan tata surya.',
    tps: [
      { kode: 'TP 8.1', judul: 'Sistem Pencernaan, Pernapasan, dan Peredaran Darah Manusia', jp: 12, status: 'SELESAI', elemen: 'Pemahaman IPA', minggu: [1, 2, 3, 4] },
      { kode: 'TP 8.2', judul: 'Usaha, Energi, dan Pesawat Sederhana dalam Kehidupan Sehari-hari', jp: 12, status: 'SELESAI', elemen: 'Keterampilan Proses', minggu: [5, 6, 7, 8] },
      { kode: 'TP 8.3', judul: 'Getaran, Gelombang, Bunyi, dan Cahaya / Alat Optik', jp: 12, status: 'DRAFT', elemen: 'Pemahaman IPA', minggu: [9, 10, 11, 12] },
      { kode: 'TP 8.4', judul: 'Unsur, Senyawa, Campuran, dan Struktur Lapisan Bumi', jp: 10, status: 'DRAFT', elemen: 'Keterampilan Proses', minggu: [13, 14, 15, 16] }
    ]
  },
  'IPAS (Ilmu Pengetahuan Alam dan Sosial)': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik mengidentifikasi keterkaitan antara pengetahuan alam dan sosial di lingkungan sekitar, siklus makhluk hidup, wujud zat, energi, keragaman budaya lokal, serta interaksi manusia dengan lingkungan.',
    tps: [
      { kode: 'TP 4.1', judul: 'Siklus Hidup Makhluk Hidup dan Upaya Pelestariannya', jp: 12, status: 'SELESAI', elemen: 'Pemahaman IPAS', minggu: [1, 2, 3, 4] },
      { kode: 'TP 4.2', judul: 'Wujud Zat dan Perubahan Wujud Benda di Lingkungan Sekitar', jp: 12, status: 'SELESAI', elemen: 'Keterampilan Proses', minggu: [5, 6, 7, 8] },
      { kode: 'TP 4.3', judul: 'Gaya di Sekitar Kita dan Pengaruhnya Terhadap Gerak Benda', jp: 10, status: 'DRAFT', elemen: 'Pemahaman IPAS', minggu: [9, 10, 11, 12] },
      { kode: 'TP 4.4', judul: 'Keragaman Budaya, Kearifan Lokal, dan Kegiatan Ekonomi Daerah', jp: 10, status: 'DRAFT', elemen: 'Pemahaman IPAS', minggu: [13, 14, 15, 16] }
    ]
  },
  'IPS (Ilmu Pengetahuan Sosial)': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik memahami kondisi geografis Indonesia, konektivitas antarruang, interaksi sosial, kegiatan ekonomi masyarakat, serta sejarah peradaban nusantara.',
    tps: [
      { kode: 'TP 8.1', judul: 'Kondisi Geografis, Potensi Sumber Daya Alam, dan Pelestarian Lingkungan', jp: 12, status: 'SELESAI', elemen: 'Pemahaman Konsep', minggu: [1, 2, 3, 4] },
      { kode: 'TP 8.2', judul: 'Interaksi Sosial, Lembaga Sosial, dan Mobilitas Masyarakat', jp: 10, status: 'SELESAI', elemen: 'Pemahaman Konsep', minggu: [5, 6, 7, 8] },
      { kode: 'TP 8.3', judul: 'Kegiatan Ekonomi (Produksi, Distribusi, Konsumsi) & Perdagangan', jp: 12, status: 'DRAFT', elemen: 'Keterampilan Proses', minggu: [9, 10, 11, 12] },
      { kode: 'TP 8.4', judul: 'Perkembangan Masyarakat Nusantara pada Masa Kolonial & Pergerakan', jp: 10, status: 'DRAFT', elemen: 'Keterampilan Proses', minggu: [13, 14, 15, 16] }
    ]
  },
  'Pendidikan Pancasila': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik memahami kedudukan Pancasila, tata urutan peraturan perundang-undangan UUD NRI Tahun 1945, keberagaman dalam Bhinneka Tunggal Ika, serta komitmen menjaga keutuhan NKRI.',
    tps: [
      { kode: 'TP 8.1', judul: 'Kedudukan dan Fungsi Pancasila sebagai Dasar Negara & Pandangan Hidup', jp: 10, status: 'SELESAI', elemen: 'Pancasila', minggu: [1, 2, 3, 4] },
      { kode: 'TP 8.2', judul: 'Tata Urutan Peraturan Perundang-undangan Nasional', jp: 10, status: 'SELESAI', elemen: 'UUD NRI 1945', minggu: [5, 6, 7, 8] },
      { kode: 'TP 8.3', judul: 'Kebangkitan Nasional dan Sumpah Pemuda dalam Bingkai Bhinneka Tunggal Ika', jp: 10, status: 'DRAFT', elemen: 'Bhinneka Tunggal Ika', minggu: [9, 10, 11, 12] },
      { kode: 'TP 8.4', judul: 'Memperkuat Komitmen Kebangsaan dan Keutuhan NKRI', jp: 8, status: 'DRAFT', elemen: 'NKRI', minggu: [13, 14, 15, 16] }
    ]
  },
  'Informatika': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik menerapkan berpikir komputasional, memahami sistem komputer dan jaringan internet, menganalisis data digital, serta mengembangkan algoritma pemrograman secara etis.',
    tps: [
      { kode: 'TP 8.1', judul: 'Berpikir Komputasional: Algoritma Searching, Sorting, dan Struktur Data', jp: 12, status: 'SELESAI', elemen: 'Berpikir Komputasional', minggu: [1, 2, 3, 4] },
      { kode: 'TP 8.2', judul: 'Sistem Komputer, Jaringan Lokal, Internet, dan Keamanan Siber', jp: 10, status: 'SELESAI', elemen: 'Jaringan Komputer', minggu: [5, 6, 7, 8] },
      { kode: 'TP 8.3', judul: 'Pengolahan, Visualisasi, dan Analisis Data Digital', jp: 12, status: 'DRAFT', elemen: 'Analisis Data', minggu: [9, 10, 11, 12] },
      { kode: 'TP 8.4', judul: 'Algoritma Pemrograman Prosedural & Praktik Lintas Bidang', jp: 14, status: 'DRAFT', elemen: 'Pemrograman', minggu: [13, 14, 15, 16] }
    ]
  },
  'PAI & Budi Pekerti': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik memahami ayat Al-Qur’an dan hadis tentang toleransi dan menuntut ilmu, memperkuat keimanan, membiasakan akhlak terpuji, melaksanakan ketentuan ibadah muamalah, serta meneladani sejarah peradaban Islam.',
    tps: [
      { kode: 'TP 10.1', judul: 'Kontrol Diri dan Prasangka Baik (Husnuzan) dalam Al-Qur’an & Hadis', jp: 12, status: 'SELESAI', elemen: 'Al-Qur’an & Hadis', minggu: [1, 2, 3, 4] },
      { kode: 'TP 10.2', judul: 'Semangat Menuntut Ilmu dan Menyebarkannya dalam Islam', jp: 10, status: 'SELESAI', elemen: 'Akidah & Akhlak', minggu: [5, 6, 7, 8] },
      { kode: 'TP 10.3', judul: 'Fikih Muamalah: Prinsip Ekonomi Syariah dan Larangan Riba', jp: 12, status: 'DRAFT', elemen: 'Fikih', minggu: [9, 10, 11, 12] },
      { kode: 'TP 10.4', judul: 'Sejarah Peradaban Islam dan Masuknya Islam ke Nusantara', jp: 10, status: 'DRAFT', elemen: 'Sejarah Peradaban Islam', minggu: [13, 14, 15, 16] }
    ]
  },
  'PJOK': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik mempraktikkan variasi keterampilan gerak spesifik permainan olahraga, atletik, senam, merancang latihan kebugaran jasmani, serta menerapkan pola hidup sehat.',
    tps: [
      { kode: 'TP 8.1', judul: 'Variasi Gerak Spesifik Permainan Bola Besar & Bola Kecil', jp: 12, status: 'SELESAI', elemen: 'Terampil Bergerak', minggu: [1, 2, 3, 4] },
      { kode: 'TP 8.2', judul: 'Keterampilan Gerak Dasar Atletik (Lari, Lompat, Lempar)', jp: 10, status: 'SELESAI', elemen: 'Terampil Bergerak', minggu: [5, 6, 7, 8] },
      { kode: 'TP 8.3', judul: 'Latihan dan Pengukuran Kebugaran Jasmani Terkait Kesehatan', jp: 12, status: 'DRAFT', elemen: 'Belajar Melalui Gerak', minggu: [9, 10, 11, 12] },
      { kode: 'TP 8.4', judul: 'Budaya Hidup Sehat, Gizi Seimbang, dan Keselamatan Diri', jp: 8, status: 'DRAFT', elemen: 'Hidup Menyehatkan', minggu: [13, 14, 15, 16] }
    ]
  },
  'Biologi': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik menganalisis struktur dan fungsi sel, jaringan tumbuhan dan hewan, sistem organ manusia (gerak, peredaran darah, pencernaan), serta keterkaitannya dengan kelainan/penyakit.',
    tps: [
      { kode: 'TP 11.1', judul: 'Struktur Sel dan Organel Sel pada Makhluk Hidup', jp: 12, status: 'SELESAI', elemen: 'Pemahaman Biologi', minggu: [1, 2, 3, 4] },
      { kode: 'TP 11.2', judul: 'Struktur dan Fungsi Jaringan Tumbuhan', jp: 10, status: 'SELESAI', elemen: 'Pemahaman Biologi', minggu: [5, 6, 7, 8] },
      { kode: 'TP 11.3', judul: 'Sistem Gerak Manusia (Rangka, Otot, Sendi)', jp: 12, status: 'DRAFT', elemen: 'Keterampilan Proses', minggu: [9, 10, 11, 12] },
      { kode: 'TP 11.4', judul: 'Sistem Peredaran Darah Manusia dan Kelainannya', jp: 12, status: 'DRAFT', elemen: 'Keterampilan Proses', minggu: [13, 14, 15, 16] }
    ]
  },
  'Ekonomi': {
    cp: 'Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025, peserta didik menganalisis pendapatan nasional, pertumbuhan ekonomi, inflasi, kebijakan moneter dan fiskal, perdagangan internasional, serta siklus akuntansi.',
    tps: [
      { kode: 'TP 11.1', judul: 'Pendapatan Nasional, Pertumbuhan Ekonomi, dan Distribusi Pendapatan', jp: 12, status: 'SELESAI', elemen: 'Pemahaman Konsep', minggu: [1, 2, 3, 4] },
      { kode: 'TP 11.2', judul: 'Inflasi, Kebijakan Moneter Bank Sentral, dan Kebijakan Fiskal', jp: 12, status: 'SELESAI', elemen: 'Keterampilan Proses', minggu: [5, 6, 7, 8] },
      { kode: 'TP 11.3', judul: 'Perdagangan Internasional, Neraca Pembayaran, dan Kurs Valuta Asing', jp: 10, status: 'DRAFT', elemen: 'Pemahaman Konsep', minggu: [9, 10, 11, 12] },
      { kode: 'TP 11.4', judul: 'Siklus Akuntansi Perusahaan Jasa dan Laporan Keuangan', jp: 14, status: 'DRAFT', elemen: 'Keterampilan Proses', minggu: [13, 14, 15, 16] }
    ]
  }
};

const FASE_LIST = [
  { label: 'Fase A (Kelas 1-2 SD)', defaultKelas: '1', mapels: ['Bahasa Indonesia', 'Matematika', 'Pendidikan Pancasila', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase B (Kelas 3-4 SD)', defaultKelas: '4', mapels: ['IPAS (Ilmu Pengetahuan Alam dan Sosial)', 'Bahasa Indonesia', 'Matematika', 'Pendidikan Pancasila', 'Bahasa Inggris', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase C (Kelas 5-6 SD)', defaultKelas: '5', mapels: ['Matematika', 'IPAS (Ilmu Pengetahuan Alam dan Sosial)', 'Bahasa Indonesia', 'Pendidikan Pancasila', 'Bahasa Inggris', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase D (Kelas 7-9 SMP)', defaultKelas: '8', mapels: ['Matematika', 'Bahasa Inggris', 'Bahasa Indonesia', 'IPA (Ilmu Pengetahuan Alam)', 'IPS (Ilmu Pengetahuan Sosial)', 'Pendidikan Pancasila', 'Informatika', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase E (Kelas 10 SMA/SMK)', defaultKelas: '10', mapels: ['Informatika', 'Bahasa Indonesia', 'Bahasa Inggris', 'Matematika', 'Biologi', 'IPA (Ilmu Pengetahuan Alam)', 'IPS (Ilmu Pengetahuan Sosial)', 'Pendidikan Pancasila', 'PAI & Budi Pekerti', 'PJOK'] },
  { label: 'Fase F (Kelas 11-12 SMA/SMK)', defaultKelas: '11', mapels: ['Biologi', 'Ekonomi', 'Bahasa Inggris', 'Bahasa Indonesia', 'Matematika', 'Informatika', 'Pendidikan Pancasila', 'PAI & Budi Pekerti', 'PJOK'] }
];

const DEFAULT_STUDENTS = [
  { id: 1, nisn: '0123456701', nama: 'Andi Pratama', jk: 'L', statusHadir: 'H', tp1: 88, tp2: 86, tp3: 89, sts: 87, sas: 90 },
  { id: 2, nisn: '0123456702', nama: 'Bunga Lestari', jk: 'P', statusHadir: 'H', tp1: 74, tp2: 76, tp3: 72, sts: 75, sas: 74 },
  { id: 3, nisn: '0123456703', nama: 'Candra Wibowo', jk: 'L', statusHadir: 'H', tp1: 93, tp2: 91, tp3: 94, sts: 92, sas: 95 },
  { id: 4, nisn: '0123456704', nama: 'Dian Wardhana', jk: 'L', statusHadir: 'S', tp1: 82, tp2: 85, tp3: 80, sts: 83, sas: 84 },
  { id: 5, nisn: '0123456705', nama: 'Eka Putri Setiawati', jk: 'P', statusHadir: 'H', tp1: 90, tp2: 88, tp3: 91, sts: 89, sas: 92 }
];

const createClassObject = (id, nama, fase, mapel, kelas, customStudents = null) => {
  const bank = SUBJECT_BANK[mapel];
  const cpText = bank
    ? `[${fase}] ${bank.cp}`
    : `Berdasarkan Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025 pada ${fase}, peserta didik mampu memahami, menerapkan, dan menganalisis konsep-konsep esensial pada mata pelajaran ${mapel} secara kritis, kreatif, dan kontekstual sesuai jenjang Kelas ${kelas}.`;

  const tps = bank
    ? bank.tps.map((t, idx) => ({ ...t, id: `tp-${id}-${idx + 1}` }))
    : [
        { id: `tp-${id}-1`, kode: `TP ${kelas}.1`, judul: `Memahami Konsep Dasar dan Ruang Lingkup ${mapel}`, jp: 12, status: 'SELESAI', elemen: 'Pemahaman Konsep', minggu: [1, 2, 3, 4] },
        { id: `tp-${id}-2`, kode: `TP ${kelas}.2`, judul: `Menganalisis Penerapan Materi ${mapel} dalam Kehidupan`, jp: 10, status: 'SELESAI', elemen: 'Keterampilan Proses', minggu: [5, 6, 7, 8] },
        { id: `tp-${id}-3`, kode: `TP ${kelas}.3`, judul: `Pemecahan Masalah Kontekstual ${mapel}`, jp: 10, status: 'DRAFT', elemen: 'Aplikasi & Proyek', minggu: [9, 10, 11, 12] },
        { id: `tp-${id}-4`, judul: `Evaluasi dan Proyek Akhir Semester ${mapel}`, kode: `TP ${kelas}.4`, jp: 12, status: 'DRAFT', elemen: 'Evaluasi', minggu: [13, 14, 15, 16] }
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
    students: studentsList,
    journals: [
      {
        id: `j-${id}-1`,
        tanggal: new Date().toISOString().split('T')[0],
        jamKe: '1-2',
        tpKode: `${tps[0].kode} - ${tps[0].judul}`,
        kegiatan: `Pembelajaran ${tps[0].judul} dengan diskusi kelompok dan eksplorasi LKPD.`,
        catatan: 'Siswa aktif mengikuti kegiatan pembelajaran.',
        hadirCount: studentsList.filter((s) => s.statusHadir === 'H').length,
        tidakHadirCount: studentsList.filter((s) => s.statusHadir !== 'H').length
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
      tp1: 82,
      tp2: 84,
      tp3: 85,
      sts: 83,
      sas: 85
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
    sekolah: 'SMP Nusantara Merdeka',
    headerKop: 'PEMERINTAH KABUPATEN NUSANTARA • DINAS PENDIDIKAN',
    tahunAjaran: '2026/2027',
    semester: 'Semester 1 (Ganjil)'
  },
  activeClassId: 'cls-smp-8a',
  classes: [
    createClassObject('cls-smp-8a', 'Kelas 8 A - Matematika', 'Fase D (Kelas 7-9 SMP)', 'Matematika', '8')
  ]
};

const STORAGE_KEY = 'eduadmin_kapurbiru_full_v8';

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
    sekolah: INITIAL_STATE.profile.sekolah,
    fase: 'Fase D (Kelas 7-9 SMP)',
    kelas: '8',
    rombel: 'A',
    mapel: 'Matematika',
    tahunAjaran: '2026/2027'
  });
  const [importFormat, setImportFormat] = useState(null);
  const [uploadedFile, setUploadedFile] = useState(null);
  const [parsedStudents, setParsedStudents] = useState(null);
  const [isParsingFile, setIsParsingFile] = useState(false);
  const fileInputRef = useRef(null);

  // --- STATE DASHBOARD & MODUL ---
  const [activeMenu, setActiveMenu] = useState('cp-tp');
  const [toast, setToast] = useState(null);
  const [selectedTpIdx, setSelectedTpIdx] = useState(0);
  const [examType, setExamType] = useState('STS (Sumatif Tengah Semester)');

  // --- MODAL TAMBAH ROMBEL PARALEL (3 KETUKAN) ---
  const [showAddRombel, setShowAddRombel] = useState(false);
  const [rombelParalelForm, setRombelParalelForm] = useState({
    fase: 'Fase D (Kelas 7-9 SMP)',
    kelas: '8',
    mapel: 'Matematika',
    selectedLetters: ['B', 'C']
  });

  // --- MODAL EDIT TP ---
  const [editingTp, setEditingTp] = useState(null);
  const [showAddTp, setShowAddTp] = useState(false);
  const [newTp, setNewTp] = useState({ kode: '', judul: '', jp: 8, status: 'DRAFT', elemen: 'Pemahaman Konsep' });

  const [newStudentName, setNewStudentName] = useState('');
  const [newStudentNisn, setNewStudentNisn] = useState('');
  const [newStudentJk, setNewStudentJk] = useState('L');

  const activeClass = data.classes.find((c) => c.id === data.activeClassId) || data.classes[0];
  const currentTp = activeClass.tps[selectedTpIdx] || activeClass.tps[0];
  const [editingCp, setEditingCp] = useState(false);
  const [cpDraft, setCpDraft] = useState(activeClass.cpText);

  const [journalForm, setJournalForm] = useState({
    tanggal: new Date().toISOString().split('T')[0],
    jamKe: '1-2',
    tpKode: activeClass.tps[0] ? `${activeClass.tps[0].kode} - ${activeClass.tps[0].judul}` : '',
    kegiatan: activeClass.tps[0] ? `Pembelajaran ${activeClass.tps[0].judul}` : '',
    catatan: ''
  });

  useEffect(() => {
    setSelectedTpIdx(0);
    setEditingCp(false);
    setCpDraft(activeClass.cpText);
    setJournalForm((prev) => ({
      ...prev,
      tpKode: activeClass.tps[0] ? `${activeClass.tps[0].kode} - ${activeClass.tps[0].judul}` : '',
      kegiatan: activeClass.tps[0] ? `Pembelajaran ${activeClass.tps[0].judul}` : ''
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

  // --- HELPER DESKRIPSI RAPOR OTOMATIS DARI NILAI ---
  const computeStudentSummary = (s) => {
    const nilaiAkhir = Math.round((Number(s.tp1) + Number(s.tp2) + Number(s.tp3) + Number(s.sts) + Number(s.sas)) / 5);
    const tpScores = [
      { tp: activeClass.tps[0]?.judul || 'Materi 1', val: Number(s.tp1) },
      { tp: activeClass.tps[1]?.judul || 'Materi 2', val: Number(s.tp2) },
      { tp: activeClass.tps[2]?.judul || 'Materi 3', val: Number(s.tp3) }
    ].sort((a, b) => b.val - a.val);

    const best = tpScores[0];
    const lowest = tpScores[tpScores.length - 1];
    let deskripsi = '';
    if (nilaiAkhir >= 90) {
      deskripsi = `Ananda sangat baik dalam ${best.tp.toLowerCase()}, serta menunjukkan penguasaan konsep yang sangat konsisten pada seluruh tujuan pembelajaran.`;
    } else if (nilaiAkhir >= 80) {
      deskripsi = `Ananda menunjukkan penguasaan yang baik dalam ${best.tp.toLowerCase()}, dan terus berkembang dalam ${lowest.tp.toLowerCase()}.`;
    } else {
      deskripsi = `Ananda perlu bimbingan dan pendampingan lebih lanjut dalam ${lowest.tp.toLowerCase()}, meskipun sudah mulai memahami ${best.tp.toLowerCase()}.`;
    }
    return { nilaiAkhir, deskripsi };
  };

  // --- HANDLER EKSPOR WORD (.DOC) UNTUK SEMUA 15 MODUL ---
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
      <h2 class="sec-head">A. Capaian Pembelajaran (CP)</h2>
      <div class="box-highlight">${activeClass.cpText}</div>
      <h2 class="sec-head">B. Alur Tujuan Pembelajaran (ATP)</h2>
      <table class="data-table">
        <thead><tr><th style="width:6%;">No</th><th style="width:12%;">Kode TP</th><th style="width:22%;">Elemen</th><th style="width:48%;">Tujuan Pembelajaran</th><th style="width:12%;">Alokasi</th></tr></thead>
        <tbody>${rowsHtml}<tr><td colspan="4" style="text-align:right;"><strong>Total Alokasi</strong></td><td style="text-align:center;"><strong>${totalJp} JP</strong></td></tr></tbody>
      </table>`;
    downloadWordDocument(`ATP_${activeClass.nama.replace(/\s+/g, '_')}`, 'ALUR TUJUAN PEMBELAJARAN (ATP)', '', data.profile, activeClass, body);
    showNotification('File Word ATP (.doc) berhasil diunduh!');
  };

  const handleExportProtaWord = () => {
    const totalJp = activeClass.tps.reduce((acc, item) => acc + (Number(item.jp) || 0), 0);
    const rowsHtml = activeClass.tps
      .map(
        (tp, idx) => `
        <tr>
          <td style="text-align:center;">${idx < 2 ? 'Semester 1 (Ganjil)' : 'Semester 2 (Genap)'}</td>
          <td style="text-align:center;"><strong>${tp.kode}</strong></td>
          <td>${tp.judul}</td>
          <td style="text-align:center;">${tp.jp} JP</td>
          <td>${idx < 2 ? 'Juli – Desember' : 'Januari – Juni'}</td>
        </tr>`
      )
      .join('');
    const body = `
      <h2 class="sec-head">Distribusi Alokasi Waktu Program Tahunan (Prota)</h2>
      <table class="data-table">
        <thead><tr><th style="width:18%;">Semester</th><th style="width:12%;">Kode TP</th><th style="width:42%;">Tujuan Pembelajaran / Lingkup Materi</th><th style="width:12%;">Alokasi JP</th><th style="width:16%;">Keterangan</th></tr></thead>
        <tbody>${rowsHtml}<tr><td colspan="3" style="text-align:right;"><strong>Total Jam Pelajaran 1 Tahun</strong></td><td style="text-align:center;"><strong>${totalJp * 2} JP</strong></td><td>36 Minggu Efektif</td></tr></tbody>
      </table>`;
    downloadWordDocument(`Prota_${activeClass.nama.replace(/\s+/g, '_')}`, 'PROGRAM TAHUNAN (PROTA)', '', data.profile, activeClass, body);
    showNotification('File Word Program Tahunan (Prota) berhasil diunduh!');
  };

  const handleExportPromesWord = () => {
    const rowsHtml = activeClass.tps
      .map(
        (tp) => `
        <tr>
          <td style="text-align:center;"><strong>${tp.kode}</strong></td>
          <td>${tp.judul}</td>
          <td style="text-align:center;">${tp.jp}</td>
          ${Array.from({ length: 16 }, (_, i) => i + 1)
            .map((m) => `<td style="text-align:center;background:${(tp.minggu || []).includes(m) ? '#dbeafe' : '#fff'};">${(tp.minggu || []).includes(m) ? 'V' : ''}</td>`)
            .join('')}
        </tr>`
      )
      .join('');
    const body = `
      <h2 class="sec-head">Distribusi Jam ke 16 Minggu Efektif (Program Semester)</h2>
      <table class="data-table">
        <thead>
          <tr>
            <th>Kode</th><th>Tujuan Pembelajaran</th><th>JP</th>
            ${Array.from({ length: 16 }, (_, i) => `<th style="text-align:center;padding:3px;">M${i + 1}</th>`).join('')}
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
    downloadWordDocument(`Promes_${activeClass.nama.replace(/\s+/g, '_')}`, 'PROGRAM SEMESTER (PROMES)', '', data.profile, activeClass, body);
    showNotification('File Word Program Semester (Promes) berhasil diunduh!');
  };

  const handleExportSilabusWord = () => {
    const rowsHtml = activeClass.tps
      .map(
        (tp) => `
        <tr>
          <td><strong>${tp.kode}</strong><br/>${tp.elemen}</td>
          <td>${tp.judul}</td>
          <td>Eksplorasi konsep, diskusi kelompok, pemecahan masalah kontekstual, dan presentasi hasil.</td>
          <td>Diagnostik, Formatif (LKPD & Observasi), Sumatif Lingkup Materi</td>
          <td style="text-align:center;">${tp.jp} JP</td>
          <td>Buku Teks Kemendikdasmen, LKPD, Media Interaktif</td>
        </tr>`
      )
      .join('');
    const body = `
      <h2 class="sec-head">Silabus Pembelajaran Kurikulum Merdeka</h2>
      <table class="data-table">
        <thead>
          <tr>
            <th style="width:14%;">Kode & Elemen</th>
            <th style="width:24%;">Tujuan Pembelajaran</th>
            <th style="width:24%;">Kegiatan Pembelajaran</th>
            <th style="width:16%;">Penilaian</th>
            <th style="width:8%;">JP</th>
            <th style="width:14%;">Sumber Belajar</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
    downloadWordDocument(`Silabus_${activeClass.nama.replace(/\s+/g, '_')}`, 'SILABUS PEMBELAJARAN', '', data.profile, activeClass, body);
    showNotification('File Word Silabus berhasil diunduh!');
  };

  const handleExportModulWord = (tp) => {
    const body = buildDeepLearningModulHtml(tp, data.profile, activeClass);
    downloadWordDocument(
      `Modul_Ajar_${activeClass.mapel.replace(/\s+/g, '_')}_${tp.kode.replace(/\s+/g, '_')}`,
      `Modul Ajar ${activeClass.mapel} Kelas ${activeClass.kelas}`,
      `${tp.kode}: ${tp.judul}`,
      data.profile,
      activeClass,
      body,
      true
    );
    showNotification(`Modul Ajar TPACK "${tp.kode}" (.doc) berhasil diunduh!`);
  };

  const handleExportLkpdWord = (tp) => {
    const body = `
      <h2 class="sec-head">Identitas Kelompok Peserta Didik</h2>
      <table class="data-table">
        <tr><td style="width:25%;"><strong>Nama Kelompok</strong></td><td style="width:75%;">....................................................................................</td></tr>
        <tr><td><strong>Anggota (4 Siswa)</strong></td><td>1. ........................................ 2. ........................................<br/>3. ........................................ 4. ........................................</td></tr>
        <tr><td><strong>Tujuan Pembelajaran</strong></td><td><strong>${tp.kode}</strong> — ${tp.judul}</td></tr>
      </table>

      <h2 class="sec-head">A. Petunjuk Pengerjaan LKPD</h2>
      <ol>
        <li>Berdoalah sebelum memulai aktivitas eksplorasi bersama kelompokmu.</li>
        <li>Amati stimulus permasalahan kontekstual yang diberikan oleh guru pada bagian B.</li>
        <li>Diskusikan langkah penyelesaian secara kolaboratif dan tuliskan hasil pengamatan pada tabel.</li>
      </ol>

      <h2 class="sec-head">B. Stimulus & Eksplorasi Masalah (${tp.elemen})</h2>
      <div class="box-highlight">
        <strong>Studi Kasus Kontekstual:</strong> Analisislah penerapan konsep <em>${tp.judul}</em> pada permasalahan nyata di lingkungan sekolah/rumah. Catat minimal 3 temuan atau variasi data pada tabel pengamatan berikut!
      </div>
      <table class="data-table">
        <thead><tr><th style="width:10%;">No</th><th style="width:35%;">Data / Kasus yang Diamati</th><th style="width:35%;">Langkah Analisis / Perhitungan</th><th style="width:20%;">Kesimpulan</th></tr></thead>
        <tbody>
          <tr><td style="height:36px;text-align:center;">1</td><td></td><td></td><td></td></tr>
          <tr><td style="height:36px;text-align:center;">2</td><td></td><td></td><td></td></tr>
          <tr><td style="height:36px;text-align:center;">3</td><td></td><td></td><td></td></tr>
        </tbody>
      </table>

      <h2 class="sec-head">C. Pertanyaan Analisis & Kesimpulan Kelompok</h2>
      <p>1. Berdasarkan tabel pengamatan di atas, pola atau hubungan apa yang dapat kelompokmu simpulkan?</p>
      <p style="border-bottom:1px dotted #64748b;height:24px;"></p>
      <p>2. Bagaimana penerapan kesimpulan tersebut untuk menyelesaikan masalah baru dalam kehidupan sehari-hari?</p>
      <p style="border-bottom:1px dotted #64748b;height:24px;"></p>
    `;
    downloadWordDocument(`LKPD_${activeClass.mapel.replace(/\s+/g, '_')}_${tp.kode.replace(/\s+/g, '_')}`, 'LEMBAR KERJA PESERTA DIDIK (LKPD)', '', data.profile, activeClass, body);
    showNotification(`File Word LKPD (${tp.kode}) berhasil diunduh!`);
  };

  const handleExportSoalHarianWord = (tp) => {
    const body = `
      <div class="box-highlight"><strong>Topik / Tujuan Pembelajaran:</strong> ${tp.kode} — ${tp.judul} (${tp.elemen})</div>
      <h2 class="sec-head">A. Soal Latihan Harian & Kuis (Lembar Siswa)</h2>
      <ol>
        <li>Jelaskan konsep utama dari <strong>${tp.judul}</strong> beserta contoh penerapannya dalam kehidupan sehari-hari! (Skor: 20)</li>
        <li>Sebuah permasalahan kontekstual berkaitan langsung dengan <em>${tp.elemen}</em> pada materi ${tp.judul}. Uraikan langkah sistematis untuk menyelesaikannya! (Skor: 25)</li>
        <li>Identifikasi kesalahan konsep yang sering terjadi saat menyelesaikan persoalan terkait ${tp.judul} dan berikan pembenarannya! (Skor: 25)</li>
        <li>Buatlah satu studi kasus sederhana beserta penyelesaian lengkapnya berdasarkan materi ${tp.kode}! (Skor: 30)</li>
      </ol>
      <h2 class="sec-head">B. Kunci Jawaban & Pedoman Penskoran</h2>
      <table class="data-table">
        <thead><tr><th style="width:10%;">No</th><th style="width:70%;">Kriteria Jawaban Benar</th><th style="width:20%;">Skor Maks</th></tr></thead>
        <tbody>
          <tr><td style="text-align:center;">1</td><td>Menyebutkan definisi/prinsip utama dengan tepat dan memberikan 1 contoh kontekstual yang relevan.</td><td style="text-align:center;">20</td></tr>
          <tr><td style="text-align:center;">2</td><td>Menuliskan tahapan penyelesaian secara runtut, akurat, dan menyertakan kesimpulan akhir.</td><td style="text-align:center;">25</td></tr>
          <tr><td style="text-align:center;">3</td><td>Menjelaskan miskonsepsi umum dan menunjukkan prosedur perbaikan yang benar.</td><td style="text-align:center;">25</td></tr>
          <tr><td style="text-align:center;">4</td><td>Menyusun soal kontekstual orisinal lengkap dengan langkah penyelesaian yang tepat.</td><td style="text-align:center;">30</td></tr>
        </tbody>
      </table>
    `;
    downloadWordDocument(`Soal_Kuis_${activeClass.mapel.replace(/\s+/g, '_')}_${tp.kode.replace(/\s+/g, '_')}`, 'SOAL LATIHAN HARIAN & KUIS', '', data.profile, activeClass, body);
    showNotification(`File Word Soal Harian & Kuis (${tp.kode}) berhasil diunduh!`);
  };

  const handleExportUjianPaketWord = (judulUjian, filePrefix) => {
    const kisiRows = activeClass.tps
      .map(
        (tp, idx) => `
        <tr>
          <td style="text-align:center;">${idx + 1}</td>
          <td><strong>${tp.kode}</strong> — ${tp.judul}</td>
          <td>Disajikan permasalahan kontekstual terkait ${tp.elemen}, peserta didik dapat menganalisis dan menentukan penyelesaian yang tepat.</td>
          <td style="text-align:center;">${idx % 2 === 0 ? 'L2 (Aplikasi)' : 'L3 (Penalaran)'}</td>
          <td style="text-align:center;">PG / Uraian</td>
          <td style="text-align:center;">No. ${idx * 2 + 1}–${idx * 2 + 2}</td>
        </tr>`
      )
      .join('');

    const body = `
      <h2 class="sec-head">BAGIAN 1: KISI-KISI PENULISAN SOAL</h2>
      <table class="data-table">
        <thead><tr><th>No</th><th>Tujuan Pembelajaran</th><th>Indikator Soal</th><th>Level</th><th>Bentuk</th><th>No. Soal</th></tr></thead>
        <tbody>${kisiRows}</tbody>
      </table>

      <h2 class="sec-head">BAGIAN 2: LEMBAR SOAL PESERTA DIDIK</h2>
      <p><strong>Petunjuk:</strong> Kerjakan soal berikut dengan jujur, teliti, dan sistematis pada lembar jawab yang tersedia.</p>
      <ol>
        ${activeClass.tps
          .map(
            (tp) => `
          <li><strong>[Indikator ${tp.kode}]</strong> Jelaskan analisis dan langkah penyelesaian permasalahan kontekstual yang berkaitan dengan <em>${tp.judul}</em>!</li>
          <li><strong>[Studi Kasus ${tp.kode}]</strong> Berikan contoh penerapan nyata materi <em>${tp.elemen}</em> dalam kehidupan sehari-hari beserta alasannya!</li>`
          )
          .join('')}
      </ol>

      <h2 class="sec-head">BAGIAN 3: KUNCI JAWABAN & RUBRIK PENILAIAN</h2>
      <table class="data-table">
        <thead><tr><th style="width:15%;">Kode TP</th><th style="width:65%;">Pokok Kunci Jawaban</th><th style="width:20%;">Bobot Nilai</th></tr></thead>
        <tbody>
          ${activeClass.tps
            .map(
              (tp) => `
            <tr>
              <td style="text-align:center;"><strong>${tp.kode}</strong></td>
              <td>Memuat konsep inti ${tp.judul}, langkah analisis yang runtut, dan simpulan tepat.</td>
              <td style="text-align:center;">25 Poin</td>
            </tr>`
            )
            .join('')}
        </tbody>
      </table>
    `;
    downloadWordDocument(`${filePrefix}_${activeClass.nama.replace(/\s+/g, '_')}`, judulUjian, '', data.profile, activeClass, body);
    showNotification(`Paket Lengkap ${judulUjian} (Kisi-kisi + Soal + Kunci + Lembar Siswa) diunduh!`);
  };

  const handleExportDeskripsiRaporWord = () => {
    const rowsHtml = activeClass.students
      .map((s, idx) => {
        const { nilaiAkhir, deskripsi } = computeStudentSummary(s);
        return `
        <tr>
          <td style="text-align:center;">${idx + 1}</td>
          <td style="text-align:center;">${s.nisn}</td>
          <td><strong>${s.nama}</strong></td>
          <td style="text-align:center;"><strong>${nilaiAkhir}</strong></td>
          <td>${deskripsi}</td>
        </tr>`;
      })
      .join('');
    const body = `
      <h2 class="sec-head">Rekapitulasi Nilai Akhir & Deskripsi Capaian Kompetensi Rapor</h2>
      <table class="data-table">
        <thead>
          <tr>
            <th style="width:5%;">No</th>
            <th style="width:15%;">NISN</th>
            <th style="width:22%;">Nama Peserta Didik</th>
            <th style="width:10%;">Nilai Akhir</th>
            <th style="width:48%;">Deskripsi Capaian Kompetensi (e-Rapor)</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
    downloadWordDocument(`Deskripsi_Rapor_${activeClass.nama.replace(/\s+/g, '_')}`, 'DESKRIPSI CAPAIAN KOMPETENSI RAPOR', '', data.profile, activeClass, body);
    showNotification('File Word Deskripsi Rapor (.doc) berhasil diunduh!');
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
      <h2 class="sec-head">Rekapitulasi Kehadiran Peserta Didik</h2>
      <table class="data-table">
        <thead><tr><th>No</th><th>NISN</th><th>Nama Lengkap Siswa</th><th>L/P</th><th>Kode</th><th>Keterangan</th></tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
    downloadWordDocument(`Presensi_${activeClass.nama.replace(/\s+/g, '_')}`, 'DAFTAR HADIR PESERTA DIDIK', '', data.profile, activeClass, body);
    showNotification('File Word Daftar Hadir (.doc) berhasil diunduh!');
  };

  // --- HANDLER LOGIN & WIZARD ---
  const handleLogin = (e) => {
    e.preventDefault();
    setData((prev) => ({ ...prev, isLoggedIn: true }));
  };

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
    } catch {
      setParsedStudents(null);
      setUploadedFile({ name: file.name, size: `${(file.size / 1024).toFixed(1)} KB`, format: formatLabel, count: 0 });
    } finally {
      setIsParsingFile(false);
    }
  };

  const handleFinishWizard = () => {
    setWizardStep(3);
    setTimeout(() => {
      const namaKelas = `Kelas ${wizForm.kelas} ${wizForm.rombel} - ${wizForm.mapel}`;
      const newId = `cls-${Date.now()}`;
      const created = createClassObject(newId, namaKelas, wizForm.fase, wizForm.mapel, wizForm.kelas, parsedStudents);
      setData((prev) => ({
        ...prev,
        wizardCompleted: true,
        profile: {
          ...prev.profile,
          sekolah: wizForm.sekolah || prev.profile.sekolah,
          tahunAjaran: wizForm.tahunAjaran || prev.profile.tahunAjaran
        },
        activeClassId: newId,
        classes: [created]
      }));
      showNotification(`15 Perangkat Administrasi untuk ${namaKelas} siap digunakan!`);
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
      elemen: newTp.elemen,
      minggu: [13, 14, 15, 16]
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
    updateActiveClass((cls) => ({ ...cls, tps: cls.tps.filter((t) => t.id !== id) }));
    setEditingTp(null);
    showNotification('Tujuan Pembelajaran dihapus.');
  };

  const toggleWeekInPromes = (tpId, weekNum) => {
    updateActiveClass((cls) => ({
      ...cls,
      tps: cls.tps.map((t) => {
        if (t.id !== tpId) return t;
        const currentWeeks = t.minggu || [];
        const updatedWeeks = currentWeeks.includes(weekNum)
          ? currentWeeks.filter((w) => w !== weekNum)
          : [...currentWeeks, weekNum].sort((a, b) => a - b);
        return { ...t, minggu: updatedWeeks };
      })
    }));
  };

  const handleStatusHadir = (studentId, status) => {
    updateActiveClass((cls) => ({
      ...cls,
      students: cls.students.map((s) => (s.id === studentId ? { ...s, statusHadir: status } : s))
    }));
  };

  const handleAllHadir = () => {
    updateActiveClass((cls) => ({
      ...cls,
      students: cls.students.map((s) => ({ ...s, statusHadir: 'H' }))
    }));
    showNotification('Semua siswa ditandai Hadir (H).');
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
      kegiatan: journalForm.kegiatan || 'Kegiatan pembelajaran dan diskusi kelas.',
      catatan: journalForm.catatan || 'Berjalan lancar.',
      hadirCount,
      tidakHadirCount
    };
    updateActiveClass((cls) => ({ ...cls, journals: [newEntry, ...cls.journals] }));
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
      tp1: 82,
      tp2: 84,
      tp3: 85,
      sts: 83,
      sas: 85
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

  // Tambah Banyak Rombel Paralel dalam 3 Ketukan
  const handleAddRombelParalel = (e) => {
    e.preventDefault();
    if (rombelParalelForm.selectedLetters.length === 0) return;
    const createdClasses = rombelParalelForm.selectedLetters.map((letter, idx) => {
      const id = `cls-${Date.now()}-${idx}`;
      const namaLengkap = `Kelas ${rombelParalelForm.kelas} ${letter} - ${rombelParalelForm.mapel}`;
      return createClassObject(id, namaLengkap, rombelParalelForm.fase, rombelParalelForm.mapel, rombelParalelForm.kelas);
    });
    setData((prev) => ({
      ...prev,
      classes: [...prev.classes, ...createdClasses],
      activeClassId: createdClasses[0].id
    }));
    setShowAddRombel(false);
    showNotification(`${createdClasses.length} rombel paralel berhasil ditambahkan sekaligus!`);
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
            <p className="text-blue-100 text-sm mt-1">Satu Aplikasi, Administrasi Guru Satu Semester</p>
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
  // TAMPILAN 2: WIZARD 3 LANGKAH (ISI DATA KELAS SEKALI AJA)
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
              <span className="font-bold text-lg text-slate-900">Data Kelas Saya (Cukup Sekali)</span>
            </div>
            <span className="px-3 py-1 rounded-full bg-blue-50 text-blue-600 text-xs font-bold">
              {wizForm.tahunAjaran}
            </span>
          </div>

          <div className="px-8 py-5 bg-slate-50/70 border-b border-slate-100 flex items-center justify-around">
            {[
              { num: 1, label: 'Data Kelas' },
              { num: 2, label: 'Impor Siswa' },
              { num: 3, label: 'Proses 15 Berkas' }
            ].map((s) => (
              <div key={s.num} className="flex flex-col items-center gap-1.5">
                <div
                  className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold transition ${
                    wizardStep >= s.num ? 'bg-blue-600 text-white shadow-sm' : 'bg-white border border-slate-300 text-slate-400'
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
                <h2 className="text-xl font-bold text-slate-900">Isi Data Kelas Sekali Aja</h2>
                <p className="text-sm text-slate-500 mt-1">
                  Dari sini CP, TP, ATP, Prota, Promes 16 minggu efektif, Silabus, Modul Ajar, LKPD, Soal STS/SAS, hingga Deskripsi Rapor langsung tersusun.
                </p>
              </div>

              <div className="space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">Satuan Pendidikan</label>
                  <input
                    type="text"
                    value={wizForm.sekolah}
                    onChange={(e) => setWizForm({ ...wizForm, sekolah: e.target.value })}
                    className="w-full rounded-xl border border-slate-200 px-4 py-2.5 text-sm bg-white"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">Jenjang / Fase Kurikulum Merdeka</label>
                  <select
                    value={wizForm.fase}
                    onChange={(e) => {
                      const found = FASE_LIST.find((f) => f.label === e.target.value) || FASE_LIST[3];
                      setWizForm({ ...wizForm, fase: found.label, kelas: found.defaultKelas, mapel: found.mapels[0] });
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
                      placeholder="Contoh: A / B / XI-2"
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2.5 text-sm"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
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
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1.5">Tahun Ajaran</label>
                    <input
                      type="text"
                      value={wizForm.tahunAjaran}
                      onChange={(e) => setWizForm({ ...wizForm, tahunAjaran: e.target.value })}
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2.5 text-sm"
                    />
                  </div>
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
                  Terintegrasi dengan format unduhan standar Dapodik dan Emis Kementerian (.xlsx / .xls / .csv).
                </p>
              </div>

              <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} className="hidden" />

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
                        Berhasil diunggah ({uploadedFile.format}) • {uploadedFile.size}
                        {uploadedFile.count > 0 ? ` • ${uploadedFile.count} Siswa terbaca` : ''}
                      </div>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      setUploadedFile(null);
                      setParsedStudents(null);
                    }}
                    className="text-xs font-semibold text-rose-600 hover:underline cursor-pointer"
                  >
                    Hapus
                  </button>
                </div>
              ) : (
                <div
                  onClick={() => handleSelectFormat(importFormat || 'dapodik')}
                  className="p-6 rounded-xl border-2 border-dashed border-slate-200 text-center hover:border-blue-400 hover:bg-slate-50/70 transition cursor-pointer"
                >
                  <Upload className="w-6 h-6 text-slate-400 mx-auto mb-2" />
                  <p className="text-sm font-semibold text-slate-700">
                    Klik Format Dapodik atau Format Emis untuk memilih file dari laptop/HP
                  </p>
                  <p className="text-xs text-slate-400 mt-1">Atau langsung klik "Proses Administrasi" untuk memakai data simulasi</p>
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
              <h3 className="text-lg font-bold text-slate-900">Menyusun 15 Perangkat Administrasi...</h3>
              <p className="text-sm text-slate-500">
                ATP, Prota, Promes 16 minggu efektif, Silabus, Modul Ajar TPACK, LKPD, Soal STS/SAS, dan Deskripsi Rapor siap!
              </p>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ============================================================================
  // TAMPILAN 3: DASHBOARD LENGKAP (15 MODUL ADMINISTRASI GURU)
  // ============================================================================
  const menuItems = [
    { id: 'beranda', label: 'Beranda & 15 Modul', icon: LayoutDashboard, group: 'UTAMA' },
    { id: 'cp-tp', label: 'CP, TP, & ATP', icon: FileText, group: 'PERSIAPAN SEMESTER' },
    { id: 'prota', label: 'Program Tahunan (Prota)', icon: CalendarRange, group: 'PERSIAPAN SEMESTER' },
    { id: 'promes', label: 'Program Semester (Promes)', icon: Calendar, group: 'PERSIAPAN SEMESTER' },
    { id: 'silabus', label: 'Silabus', icon: ListChecks, group: 'PERSIAPAN SEMESTER' },
    { id: 'modul-ajar', label: 'Modul Ajar / RPP', icon: BookMarked, group: 'MATERI BARU (PILIH TP)' },
    { id: 'lkpd', label: 'LKPD Siswa', icon: ClipboardList, group: 'MATERI BARU (PILIH TP)' },
    { id: 'soal-harian', label: 'Soal Harian & Kuis', icon: HelpCircle, group: 'MATERI BARU (PILIH TP)' },
    { id: 'asesmen-sumatif', label: 'Asesmen Sumatif', icon: FileCheck2, group: 'UJIAN & EVALUASI' },
    { id: 'soal-sts-sas', label: 'Soal STS dan SAS', icon: GraduationCap, group: 'UJIAN & EVALUASI' },
    { id: 'data-siswa', label: 'Data Siswa', icon: Users, group: 'KBM HARIAN & RAPOR' },
    { id: 'jurnal-presensi', label: 'Presensi & Jurnal', icon: CheckCircle2, group: 'KBM HARIAN & RAPOR' },
    { id: 'daftar-nilai', label: 'Daftar Nilai', icon: Award, group: 'KBM HARIAN & RAPOR' },
    { id: 'deskripsi-rapor', label: 'Deskripsi Rapor', icon: FileSpreadsheet, group: 'KBM HARIAN & RAPOR' },
    { id: 'paket-supervisi', label: 'Paket Siap Supervisi', icon: FolderCheck, group: 'SUPERVISI & FORMAT' },
    { id: 'pengaturan', label: 'Format Sekolah & Profil', icon: Settings, group: 'SUPERVISI & FORMAT' }
  ];

  const totalJp = activeClass.tps.reduce((acc, item) => acc + (Number(item.jp) || 0), 0);
  const selectedRombelFaseObj = FASE_LIST.find((f) => f.label === rombelParalelForm.fase) || FASE_LIST[3];

  // Komponen Pemilih Tujuan Pembelajaran (untuk Modul Ajar, LKPD, Soal Harian, Asesmen Sumatif)
  const TpSelectorBar = () => (
    <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-sm">
      <div className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-2.5">
        Pilih Tujuan Pembelajaran (Otomatis Menyusun Dokumen Word):
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
        {activeClass.tps.map((tp, idx) => (
          <button
            key={tp.id}
            type="button"
            onClick={() => setSelectedTpIdx(idx)}
            className={`p-3 rounded-xl border text-left transition cursor-pointer ${
              selectedTpIdx === idx
                ? 'border-blue-600 bg-blue-50/70 ring-2 ring-blue-100'
                : 'border-slate-200 bg-white hover:border-blue-300'
            }`}
          >
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-blue-600">{tp.kode}</span>
              <span className="text-[11px] text-slate-400">{tp.jp} JP</span>
            </div>
            <div className="text-xs font-semibold text-slate-800 mt-1 line-clamp-2">{tp.judul}</div>
          </button>
        ))}
      </div>
    </div>
  );

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
            <span className="text-[11px] text-slate-400 font-medium">Administrasi 1 Semester</span>
          </div>
        </div>

        {/* KARTU KELAS AKTIF & ROMBEL PARALEL */}
        <div className="p-4 border-b border-slate-100">
          <div className="flex items-center justify-between text-[11px] font-semibold text-slate-400 uppercase tracking-wider mb-1">
            <span>KELAS AKTIF • {activeClass.fase.split(' ')[0]} {activeClass.fase.split(' ')[1]}</span>
            <button onClick={() => setActiveMenu('pengaturan')} className="text-slate-400 hover:text-blue-600">
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
                  showNotification(`Beralih ke ${cls.nama}`);
                }}
                className={`px-2.5 py-1.5 rounded-lg text-xs font-semibold transition cursor-pointer ${
                  cls.id === data.activeClassId
                    ? 'bg-blue-600 text-white shadow-sm'
                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                }`}
              >
                {cls.nama.split(' - ')[0].replace('Kelas ', '')}
              </button>
            ))}
            <button
              onClick={() => setShowAddRombel(true)}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold border border-blue-200 bg-blue-50/50 text-blue-600 hover:bg-blue-100 flex items-center gap-1 cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" /> Rombel Paralel
            </button>
          </div>
        </div>

        {/* NAVIGASI 15 MODUL */}
        <nav className="p-3 flex-1 space-y-4 overflow-y-auto">
          {[
            'UTAMA',
            'PERSIAPAN SEMESTER',
            'MATERI BARU (PILIH TP)',
            'UJIAN & EVALUASI',
            'KBM HARIAN & RAPOR',
            'SUPERVISI & FORMAT'
          ].map((group) => (
            <div key={group}>
              <div className="px-3 mb-1 text-[10px] font-bold text-slate-400 tracking-wider uppercase">{group}</div>
              <div className="space-y-0.5">
                {menuItems
                  .filter((m) => m.group === group)
                  .map((item) => {
                    const Icon = item.icon;
                    const isActive = activeMenu === item.id;
                    return (
                      <button
                        key={item.id}
                        onClick={() => setActiveMenu(item.id)}
                        className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition cursor-pointer ${
                          isActive
                            ? 'bg-blue-50 text-blue-600 font-bold'
                            : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
                        }`}
                      >
                        <Icon className={`w-4 h-4 flex-shrink-0 ${isActive ? 'text-blue-600' : 'text-slate-400'}`} />
                        <span className="truncate">{item.label}</span>
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
              setData((prev) => ({ ...prev, wizardCompleted: false }));
            }}
            className="w-full py-2 px-3 rounded-xl text-xs font-semibold text-slate-600 hover:bg-slate-100 flex items-center justify-center gap-1.5 cursor-pointer"
          >
            <Upload className="w-3.5 h-3.5" /> Ubah Data Kelas / Impor
          </button>
        </div>
      </aside>

      {/* AREA KONTEN UTAMA */}
      <div className="flex-1 flex flex-col min-w-0">
        <header className="bg-white border-b border-slate-200 px-6 py-3.5 flex items-center justify-between sticky top-0 z-10">
          <div className="flex items-center gap-2 text-sm text-slate-500">
            <span>EduAdmin</span>
            <ChevronRight className="w-4 h-4 text-slate-300" />
            <span className="font-semibold text-slate-800">{menuItems.find((m) => m.id === activeMenu)?.label}</span>
          </div>
          <div className="flex items-center gap-3 text-right">
            <div>
              <div className="text-sm font-bold text-slate-800">{data.profile.nama}</div>
              <div className="text-xs text-slate-400">{data.profile.sekolah}</div>
            </div>
            <button
              onClick={() => setData((prev) => ({ ...prev, isLoggedIn: false }))}
              title="Keluar"
              className="w-9 h-9 rounded-full bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-600 hover:text-rose-600 cursor-pointer"
            >
              <LogOut className="w-4 h-4" />
            </button>
          </div>
        </header>

        <main className="p-6 max-w-6xl w-full mx-auto space-y-6">
          {/* 1. BERANDA & PETA 15 MODUL */}
          {activeMenu === 'beranda' && (
            <div className="space-y-6">
              <div className="bg-gradient-to-r from-blue-600 to-indigo-600 rounded-2xl p-6 text-white shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                  <h2 className="text-xl font-bold">Satu Aplikasi, Administrasi Guru Satu Semester</h2>
                  <p className="text-blue-100 text-sm mt-1">
                    {data.profile.sekolah} • {activeClass.nama} • {activeClass.fase}
                  </p>
                </div>
                <button
                  onClick={() => setActiveMenu('paket-supervisi')}
                  className="px-4 py-2.5 rounded-xl bg-white text-blue-600 font-bold text-xs shadow-sm hover:bg-blue-50 cursor-pointer self-start sm:self-center"
                >
                  Lihat Paket Siap Supervisi
                </button>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
                {[
                  { id: 'cp-tp', title: 'CP & TP', desc: `${activeClass.tps.length} TP Tersusun` },
                  { id: 'cp-tp', title: 'ATP', desc: `Total ${totalJp} JP` },
                  { id: 'prota', title: 'Prota', desc: 'Program Tahunan' },
                  { id: 'promes', title: 'Promes', desc: '16 Minggu Efektif' },
                  { id: 'silabus', title: 'Silabus', desc: 'Siap Unduh Word' },
                  { id: 'modul-ajar', title: 'Modul Ajar / RPP', desc: 'TPACK 8 Bagian' },
                  { id: 'lkpd', title: 'LKPD', desc: 'Lembar Kerja Siswa' },
                  { id: 'soal-harian', title: 'Soal Harian & Kuis', desc: 'Soal + Kunci' },
                  { id: 'asesmen-sumatif', title: 'Asesmen Sumatif', desc: 'Kisi-kisi + Kunci' },
                  { id: 'soal-sts-sas', title: 'Soal STS dan SAS', desc: 'Lengkap Lembar Siswa' },
                  { id: 'data-siswa', title: 'Data Siswa', desc: `${activeClass.students.length} Siswa Aktif` },
                  { id: 'jurnal-presensi', title: 'Presensi', desc: 'Absen H/S/I/A' },
                  { id: 'jurnal-presensi', title: 'Jurnal Mengajar', desc: `${activeClass.journals.length} Catatan KBM` },
                  { id: 'daftar-nilai', title: 'Rekap Nilai Rapor', desc: 'Otomatis Terhitung' },
                  { id: 'deskripsi-rapor', title: 'Deskripsi Rapor', desc: 'Narasi Per Anak' }
                ].map((card, i) => (
                  <button
                    key={i}
                    onClick={() => setActiveMenu(card.id)}
                    className="bg-white p-4 rounded-2xl border border-slate-200 hover:border-blue-400 hover:shadow-sm transition text-left cursor-pointer"
                  >
                    <div className="font-bold text-sm text-slate-900">{card.title}</div>
                    <div className="text-xs text-blue-600 font-medium mt-1">{card.desc}</div>
                  </button>
                ))}
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
                      {activeClass.mapel} • Kelas {activeClass.kelas} ({activeClass.fase})
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
                    className="px-3 py-1.5 text-xs font-semibold rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50 cursor-pointer"
                  >
                    {editingCp ? 'Simpan Teks CP' : 'Edit Teks CP'}
                  </button>
                </div>

                {editingCp ? (
                  <textarea
                    rows={4}
                    value={cpDraft}
                    onChange={(e) => setCpDraft(e.target.value)}
                    className="mt-4 w-full rounded-xl border border-blue-300 p-3.5 text-sm text-slate-700"
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
                    <h3 className="text-base font-bold text-slate-900">Alur Tujuan Pembelajaran (ATP)</h3>
                    <p className="text-xs text-slate-500 mt-0.5">Total Alokasi Waktu: {totalJp} Jam Pelajaran (JP)</p>
                  </div>
                  <button
                    onClick={handleExportAtpWord}
                    className="px-3.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer"
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
                        <div className="px-2.5 py-1.5 rounded-lg bg-blue-50 text-blue-600 font-bold text-xs text-center min-w-[60px]">
                          {tp.kode}
                        </div>
                        <div>
                          <div className="font-semibold text-sm text-slate-900">{tp.judul}</div>
                          <div className="text-xs text-slate-500 mt-1 flex flex-wrap items-center gap-2">
                            <span>Alokasi: {tp.jp} JP</span>
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
                            tp.status === 'SELESAI' ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-600'
                          }`}
                        >
                          {tp.status}
                        </button>
                        <button
                          onClick={() => setEditingTp(tp)}
                          className="p-1.5 rounded-lg text-slate-400 hover:text-blue-600 hover:bg-blue-50"
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
                    className="mt-4 w-full py-3 rounded-xl border border-dashed border-blue-300 text-blue-600 font-semibold text-sm hover:bg-blue-50/60 transition flex items-center justify-center gap-1.5 cursor-pointer"
                  >
                    <Plus className="w-4 h-4" /> Tambah Tujuan Pembelajaran Baru
                  </button>
                ) : (
                  <form onSubmit={handleAddTp} className="mt-4 p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
                    <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                      <input
                        type="text"
                        placeholder="Kode (mis. TP 8.5)"
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
                        placeholder="JP"
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
                        <button type="button" onClick={() => setShowAddTp(false)} className="px-3 py-1.5 rounded-lg text-sm text-slate-600">
                          Batal
                        </button>
                        <button type="submit" className="px-4 py-1.5 rounded-lg bg-blue-600 text-white text-sm font-semibold">
                          Simpan TP
                        </button>
                      </div>
                    </div>
                  </form>
                )}
              </div>
            </div>
          )}

          {/* 3. PROGRAM TAHUNAN (PROTA) */}
          {activeMenu === 'prota' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Program Tahunan (Prota)</h2>
                  <p className="text-xs text-slate-500">
                    {activeClass.nama} • {activeClass.fase} • Tahun Ajaran {data.profile.tahunAjaran}
                  </p>
                </div>
                <button
                  onClick={handleExportProtaWord}
                  className="px-3.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer"
                >
                  <FileDown className="w-3.5 h-3.5" /> Unduh Word Prota (.doc)
                </button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-y border-slate-200 text-xs uppercase text-slate-500">
                      <th className="py-3 px-4">Semester</th>
                      <th className="py-3 px-4">Kode TP</th>
                      <th className="py-3 px-4">Lingkup Materi / Tujuan Pembelajaran</th>
                      <th className="py-3 px-4 text-center">Alokasi JP</th>
                      <th className="py-3 px-4">Periode Bulan</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {activeClass.tps.map((tp, idx) => (
                      <tr key={tp.id} className="hover:bg-slate-50/70">
                        <td className="py-3 px-4 font-semibold text-slate-700">
                          {idx < 2 ? 'Semester 1 (Ganjil)' : 'Semester 2 (Genap)'}
                        </td>
                        <td className="py-3 px-4 font-bold text-blue-600">{tp.kode}</td>
                        <td className="py-3 px-4 text-slate-800">{tp.judul}</td>
                        <td className="py-3 px-4 text-center font-bold">{tp.jp} JP</td>
                        <td className="py-3 px-4 text-xs text-slate-500">{idx < 2 ? 'Juli – Desember' : 'Januari – Juni'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 4. PROGRAM SEMESTER (PROMES — 16 MINGGU EFEKTIF SEPERTI VIDEO) */}
          {activeMenu === 'promes' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Program Semester (Promes) — 16 Minggu Efektif</h2>
                  <p className="text-xs text-slate-500">
                    {activeClass.nama} • Klik kotak minggu (1–16) untuk menyesuaikan distribusi jam efektif.
                  </p>
                </div>
                <button
                  onClick={handleExportPromesWord}
                  className="px-3.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer"
                >
                  <FileDown className="w-3.5 h-3.5" /> Unduh Word Promes (.doc)
                </button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-y border-slate-200 text-slate-500">
                      <th className="py-2.5 px-3">Kode & TP</th>
                      <th className="py-2.5 px-2 text-center">JP</th>
                      {Array.from({ length: 16 }, (_, i) => i + 1).map((m) => (
                        <th key={m} className="py-2.5 px-1 text-center w-7">
                          {m}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {activeClass.tps.map((tp) => (
                      <tr key={tp.id} className="hover:bg-slate-50/70">
                        <td className="py-3 px-3 min-w-[200px]">
                          <div className="font-bold text-blue-600">{tp.kode}</div>
                          <div className="text-slate-700 font-medium line-clamp-1">{tp.judul}</div>
                        </td>
                        <td className="py-3 px-2 text-center font-bold text-slate-700">{tp.jp} JP</td>
                        {Array.from({ length: 16 }, (_, i) => i + 1).map((m) => {
                          const active = (tp.minggu || []).includes(m);
                          return (
                            <td key={m} className="py-2 px-0.5 text-center">
                              <button
                                type="button"
                                onClick={() => toggleWeekInPromes(tp.id, m)}
                                className={`w-6 h-7 rounded-md transition cursor-pointer ${
                                  active ? 'bg-blue-600 shadow-xs' : 'bg-slate-100 hover:bg-slate-200'
                                }`}
                                title={`Minggu ke-${m}`}
                              />
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 5. SILABUS */}
          {activeMenu === 'silabus' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Silabus Pembelajaran ({activeClass.nama})</h2>
                  <p className="text-xs text-slate-500">Terhubung otomatis dengan CP, TP, Kegiatan Pembelajaran, dan Penilaian.</p>
                </div>
                <button
                  onClick={handleExportSilabusWord}
                  className="px-3.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer"
                >
                  <FileDown className="w-3.5 h-3.5" /> Unduh Word Silabus (.doc)
                </button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-y border-slate-200 uppercase text-slate-500">
                      <th className="py-3 px-3">Kode & Elemen</th>
                      <th className="py-3 px-3">Tujuan Pembelajaran</th>
                      <th className="py-3 px-3">Kegiatan Pembelajaran</th>
                      <th className="py-3 px-3">Asesmen</th>
                      <th className="py-3 px-3 text-center">JP</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {activeClass.tps.map((tp) => (
                      <tr key={tp.id} className="hover:bg-slate-50/70">
                        <td className="py-3 px-3">
                          <div className="font-bold text-blue-600">{tp.kode}</div>
                          <div className="text-slate-500">{tp.elemen}</div>
                        </td>
                        <td className="py-3 px-3 font-medium text-slate-800">{tp.judul}</td>
                        <td className="py-3 px-3 text-slate-600">Eksplorasi konsep, diskusi kelompok, pemecahan masalah kontekstual, dan refleksi.</td>
                        <td className="py-3 px-3 text-slate-600">Diagnostik, Formatif (LKPD), Sumatif</td>
                        <td className="py-3 px-3 text-center font-bold">{tp.jp} JP</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 6. MODUL AJAR / RPP (PILIH TP -> UNDUH WORD TPACK) */}
          {activeMenu === 'modul-ajar' && (
            <div className="space-y-4">
              <TpSelectorBar />
              <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="px-2.5 py-1 rounded-md bg-blue-50 text-blue-600 text-xs font-bold">{currentTp.kode}</span>
                    <h2 className="text-lg font-bold text-slate-900 mt-1.5">Modul Ajar Mendalam Berbasis TPACK: {currentTp.judul}</h2>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Alokasi: {currentTp.jp} JP • Elemen: {currentTp.elemen} • Format Word 8 Bagian Lengkap (A–H)
                    </p>
                  </div>
                  <button
                    onClick={() => handleExportModulWord(currentTp)}
                    className="px-4 py-2.5 rounded-xl bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer self-start"
                  >
                    <FileDown className="w-4 h-4" /> Unduh Word Modul Ajar (.doc)
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs bg-slate-50 p-4 rounded-xl border border-slate-200/80">
                  <div><strong>A. Identitas & Capaian:</strong> {activeClass.mapel} • {activeClass.fase} • Kelas {activeClass.kelas}</div>
                  <div><strong>B. Tujuan Pembelajaran:</strong> Indikator terukur, pemahaman bermakna & pertanyaan pemantik</div>
                  <div><strong>C. Analisis TPACK:</strong> CK, PK, TK, PCK, TCK, TPK, hingga integrasi TPACK utuh</div>
                  <div><strong>D–E. Langkah Deep Learning:</strong> Pendahuluan, Memahami, Mengaplikasi, Merefleksi, Penutup</div>
                  <div><strong>F. Asesmen & Rubrik:</strong> Diagnostik, Formatif, Exit Ticket, dan Rubrik 4 Level</div>
                  <div><strong>G–H. Diferensiasi & Refleksi:</strong> Dukungan, Tantangan, Opsi Luring, dan Tindak Lanjut</div>
                </div>
              </div>
            </div>
          )}

          {/* 7. LKPD SISWA (PILIH TP -> UNDUH WORD LKPD) */}
          {activeMenu === 'lkpd' && (
            <div className="space-y-4">
              <TpSelectorBar />
              <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="px-2.5 py-1 rounded-md bg-emerald-50 text-emerald-700 text-xs font-bold">LKPD • {currentTp.kode}</span>
                    <h2 className="text-lg font-bold text-slate-900 mt-1.5">Lembar Kerja Peserta Didik: {currentTp.judul}</h2>
                    <p className="text-xs text-slate-500 mt-0.5">Berisi kolom identitas kelompok, petunjuk kerja, tabel pengamatan, dan pertanyaan analisis.</p>
                  </div>
                  <button
                    onClick={() => handleExportLkpdWord(currentTp)}
                    className="px-4 py-2.5 rounded-xl bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer self-start"
                  >
                    <FileDown className="w-4 h-4" /> Unduh Word LKPD (.doc)
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* 8. SOAL HARIAN & KUIS (PILIH TP -> UNDUH WORD SOAL + KUNCI) */}
          {activeMenu === 'soal-harian' && (
            <div className="space-y-4">
              <TpSelectorBar />
              <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="px-2.5 py-1 rounded-md bg-amber-50 text-amber-700 text-xs font-bold">Latihan & Kuis • {currentTp.kode}</span>
                    <h2 className="text-lg font-bold text-slate-900 mt-1.5">Soal Latihan Harian & Kuis: {currentTp.judul}</h2>
                    <p className="text-xs text-slate-500 mt-0.5">Lengkap dengan lembar soal siswa, kunci jawaban, dan pedoman penskoran.</p>
                  </div>
                  <button
                    onClick={() => handleExportSoalHarianWord(currentTp)}
                    className="px-4 py-2.5 rounded-xl bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer self-start"
                  >
                    <FileDown className="w-4 h-4" /> Unduh Word Soal & Kunci (.doc)
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* 9. ASESMEN SUMATIF (KISI-KISI + SOAL & KUNCI + LEMBAR SISWA) */}
          {activeMenu === 'asesmen-sumatif' && (
            <div className="space-y-4">
              <TpSelectorBar />
              <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h2 className="text-lg font-bold text-slate-900">Asesmen Sumatif Lingkup Materi ({currentTp.kode})</h2>
                    <p className="text-xs text-slate-500 mt-0.5">Paket lengkap: Kisi-kisi soal, Lembar Peserta Didik, Kunci Jawaban, dan Rubrik.</p>
                  </div>
                  <button
                    onClick={() => handleExportUjianPaketWord(`ASESMEN SUMATIF (${currentTp.kode})`, `Asesmen_Sumatif_${currentTp.kode}`)}
                    className="px-4 py-2.5 rounded-xl bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer self-start"
                  >
                    <FileDown className="w-4 h-4" /> Unduh Paket Asesmen Word (.doc)
                  </button>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
                    <div className="font-bold text-sm text-slate-800">1. Kisi-Kisi Penulisan Soal</div>
                    <p className="text-xs text-slate-500 mt-1">Pemetaan TP, indikator soal, level kognitif (L2/L3), dan nomor soal.</p>
                  </div>
                  <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
                    <div className="font-bold text-sm text-slate-800">2. Soal dan Kunci Jawaban</div>
                    <p className="text-xs text-slate-500 mt-1">Naskah soal terukur lengkap dengan pembahasan dan bobot nilai.</p>
                  </div>
                  <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
                    <div className="font-bold text-sm text-slate-800">3. Lembar Siswa Siap Cetak</div>
                    <p className="text-xs text-slate-500 mt-1">Format naskah ujian bersih tanpa kunci jawaban untuk dibagikan ke siswa.</p>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* 10. SOAL STS DAN SAS */}
          {activeMenu === 'soal-sts-sas' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-5">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Soal STS dan SAS ({activeClass.nama})</h2>
                  <p className="text-xs text-slate-500 mt-0.5">
                    Sumatif Tengah Semester (STS) & Sumatif Akhir Semester (SAS) lengkap Kisi-Kisi, Kunci Jawaban, dan Lembar Siswa.
                  </p>
                </div>
                <div className="flex gap-2">
                  {['STS (Sumatif Tengah Semester)', 'SAS (Sumatif Akhir Semester)'].map((type) => (
                    <button
                      key={type}
                      onClick={() => setExamType(type)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold border cursor-pointer ${
                        examType === type ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-200'
                      }`}
                    >
                      {type.split(' ')[0]}
                    </button>
                  ))}
                </div>
              </div>

              <div className="p-5 rounded-2xl bg-slate-50 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                  <div className="font-bold text-base text-slate-900">{examType}</div>
                  <div className="text-xs text-slate-500 mt-1">
                    Mencakup seluruh {activeClass.tps.length} Tujuan Pembelajaran ({activeClass.tps.map((t) => t.kode).join(', ')})
                  </div>
                </div>
                <button
                  onClick={() => handleExportUjianPaketWord(examType.toUpperCase(), examType.split(' ')[0])}
                  className="px-4 py-2.5 rounded-xl bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer self-start"
                >
                  <FileDown className="w-4 h-4" /> Unduh Paket {examType.split(' ')[0]} Word (.doc)
                </button>
              </div>
            </div>
          )}

          {/* 11. DATA SISWA */}
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
                  className="bg-blue-600 text-white font-semibold text-sm rounded-lg px-4 py-2 hover:bg-blue-700 flex items-center justify-center gap-1.5 cursor-pointer"
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

          {/* 12. PRESENSI & JURNAL MENGAJAR */}
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
                    <label className="block text-xs font-semibold text-slate-600 mb-1">Dari Promes / Tujuan Pembelajaran</label>
                    <select
                      value={journalForm.tpKode}
                      onChange={(e) =>
                        setJournalForm({
                          ...journalForm,
                          tpKode: e.target.value,
                          kegiatan: `Pembelajaran ${e.target.value}`
                        })
                      }
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
                    <label className="block text-xs font-semibold text-slate-600 mb-1">Kegiatan Pembelajaran</label>
                    <input
                      type="text"
                      value={journalForm.kegiatan}
                      onChange={(e) => setJournalForm({ ...journalForm, kegiatan: e.target.value })}
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-600 mb-1">Catatan (Opsional)</label>
                    <textarea
                      rows={2}
                      placeholder="Tambahkan catatan khusus..."
                      value={journalForm.catatan}
                      onChange={(e) => setJournalForm({ ...journalForm, catatan: e.target.value })}
                      className="w-full rounded-xl border border-slate-200 px-3.5 py-2 text-sm"
                    />
                  </div>
                  <button
                    type="submit"
                    className="w-full py-2.5 rounded-xl bg-blue-600 text-white font-semibold text-sm hover:bg-blue-700 transition shadow-sm cursor-pointer"
                  >
                    Simpan Jurnal
                  </button>
                </form>

                <div className="lg:col-span-7 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                  <div className="flex flex-wrap items-center justify-between gap-2 mb-4">
                    <div>
                      <h2 className="text-base font-bold text-slate-900">Presensi Harian</h2>
                      <p className="text-xs text-slate-500">
                        {activeClass.students.filter((s) => s.statusHadir === 'H').length} Hadir dari {activeClass.students.length} Siswa
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={handleAllHadir}
                        className="px-3 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-600 hover:bg-slate-50 cursor-pointer"
                      >
                        Kembalikan semua jadi Hadir
                      </button>
                      <button
                        type="button"
                        onClick={handleExportPresensiWord}
                        className="px-3 py-1.5 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1 cursor-pointer"
                      >
                        <FileDown className="w-3.5 h-3.5" /> Unduh Word
                      </button>
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
                              className={`w-8 h-8 rounded-lg text-xs font-bold border transition cursor-pointer ${
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
            </div>
          )}

          {/* 13. DAFTAR NILAI */}
          {activeMenu === 'daftar-nilai' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Daftar Nilai per Tujuan Pembelajaran ({activeClass.nama})</h2>
                  <p className="text-xs text-slate-500">Nilai akhir otomatis terhitung dan langsung menyusun Deskripsi Rapor tiap anak.</p>
                </div>
                <button
                  onClick={handleExportDeskripsiRaporWord}
                  className="px-3.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer"
                >
                  <FileDown className="w-3.5 h-3.5" /> Ekspor Rekap Nilai Rapor (.doc)
                </button>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left">
                  <thead>
                    <tr className="bg-slate-50 border-y border-slate-200 text-xs uppercase text-slate-500">
                      <th className="py-3 px-3">No</th>
                      <th className="py-3 px-3">Nama Siswa</th>
                      <th className="py-3 px-2 text-center">TP 1</th>
                      <th className="py-3 px-2 text-center">TP 2</th>
                      <th className="py-3 px-2 text-center">TP 3</th>
                      <th className="py-3 px-2 text-center">STS</th>
                      <th className="py-3 px-2 text-center">SAS</th>
                      <th className="py-3 px-3 text-center">Nilai Akhir</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {activeClass.students.map((s, idx) => {
                      const { nilaiAkhir } = computeStudentSummary(s);
                      return (
                        <tr key={s.id} className="hover:bg-slate-50/70">
                          <td className="py-2.5 px-3 text-slate-500">{idx + 1}</td>
                          <td className="py-2.5 px-3 font-semibold text-slate-800">{s.nama}</td>
                          {['tp1', 'tp2', 'tp3', 'sts', 'sas'].map((field) => (
                            <td key={field} className="py-2.5 px-2 text-center">
                              <input
                                type="number"
                                value={s[field]}
                                onChange={(e) => handleGradeChange(s.id, field, e.target.value)}
                                className="w-14 text-center rounded-lg border border-slate-200 py-1 text-sm"
                              />
                            </td>
                          ))}
                          <td className="py-2.5 px-3 text-center font-bold text-blue-600">{nilaiAkhir}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 14. DESKRIPSI RAPOR OTOMATIS */}
          {activeMenu === 'deskripsi-rapor' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Deskripsi Rapor • {activeClass.nama}</h2>
                  <p className="text-xs text-slate-500">Deskripsi rapor tiap anak otomatis tersusun dari nilai per Tujuan Pembelajaran.</p>
                </div>
                <button
                  onClick={handleExportDeskripsiRaporWord}
                  className="px-4 py-2 rounded-xl bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer"
                >
                  <FileDown className="w-4 h-4" /> Ekspor: Rekap Nilai & Deskripsi Rapor (.doc)
                </button>
              </div>

              <div className="space-y-3">
                {activeClass.students.map((s) => {
                  const { nilaiAkhir, deskripsi } = computeStudentSummary(s);
                  return (
                    <div key={s.id} className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 flex items-start justify-between gap-4">
                      <div>
                        <div className="font-bold text-sm text-slate-900">{s.nama}</div>
                        <p className="text-xs text-slate-600 mt-1 leading-relaxed">{deskripsi}</p>
                      </div>
                      <span className="px-3 py-1 rounded-lg bg-blue-100 text-blue-700 font-bold text-sm flex-shrink-0">
                        {nilaiAkhir}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* 15. PAKET SIAP SUPERVISI */}
          {activeMenu === 'paket-supervisi' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Paket Siap Supervisi ({activeClass.nama})</h2>
                <p className="text-xs text-slate-500">Seluruh berkas administrasi terhubung dan siap diunduh dalam format Microsoft Word (.doc).</p>
              </div>
              <div className="divide-y divide-slate-100 border border-slate-200 rounded-xl overflow-hidden">
                {[
                  { name: 'Alur Tujuan Pembelajaran (ATP)', action: handleExportAtpWord },
                  { name: 'Program Tahunan (Prota)', action: handleExportProtaWord },
                  { name: 'Program Semester (Promes 16 Minggu)', action: handleExportPromesWord },
                  { name: 'Silabus Pembelajaran', action: handleExportSilabusWord },
                  { name: `Modul Ajar / RPP (${currentTp.kode})`, action: () => handleExportModulWord(currentTp) },
                  { name: `LKPD & Soal Harian (${currentTp.kode})`, action: () => handleExportLkpdWord(currentTp) },
                  { name: 'Paket Asesmen Sumatif / STS / SAS', action: () => handleExportUjianPaketWord('PAKET ASESMEN SUMATIF', 'Asesmen_Sumatif') },
                  { name: 'Rekap Nilai & Deskripsi Rapor', action: handleExportDeskripsiRaporWord }
                ].map((doc, idx) => (
                  <div key={idx} className="p-4 flex items-center justify-between bg-white hover:bg-slate-50">
                    <div className="flex items-center gap-3">
                      <FileText className="w-4 h-4 text-blue-600" />
                      <span className="text-sm font-semibold text-slate-800">{doc.name}</span>
                    </div>
                    <div className="flex items-center gap-3">
                      <span className="text-xs font-bold text-emerald-600 bg-emerald-50 px-2.5 py-1 rounded-full">Sudah ada</span>
                      <button
                        onClick={doc.action}
                        className="px-3 py-1.5 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 cursor-pointer"
                      >
                        Unduh Word
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* 16. FORMAT SEKOLAH & PROFIL */}
          {activeMenu === 'pengaturan' && (
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm max-w-2xl space-y-5">
              <div>
                <h2 className="text-lg font-bold text-slate-900">Template Format Sekolah & Identitas Guru</h2>
                <p className="text-xs text-slate-500">Sekolah punya format kop sendiri? Sesuaikan di sini dan otomatis dipakai di seluruh file Word.</p>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="sm:col-span-2">
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Baris Atas Kop Surat Sekolah</label>
                  <input
                    type="text"
                    value={data.profile.headerKop || ''}
                    onChange={(e) =>
                      setData((prev) => ({ ...prev, profile: { ...prev.profile, headerKop: e.target.value } }))
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
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Nama Lengkap & Gelar Guru</label>
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
              </div>

              <div className="pt-4 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
                <button
                  type="button"
                  onClick={() => showNotification('Format sekolah & identitas berhasil disimpan!')}
                  className="px-4 py-2 rounded-xl bg-blue-600 text-white text-sm font-semibold hover:bg-blue-700 flex items-center gap-1.5 cursor-pointer"
                >
                  <Save className="w-4 h-4" /> Pakai Format Sekolah
                </button>
                <button
                  type="button"
                  onClick={handleResetAll}
                  className="px-3 py-2 rounded-xl border border-rose-200 text-rose-600 text-xs font-semibold hover:bg-rose-50 flex items-center gap-1.5 cursor-pointer"
                >
                  <RotateCcw className="w-4 h-4" /> Reset Awal
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
                <button type="button" onClick={() => setEditingTp(null)} className="px-4 py-2 rounded-xl text-sm text-slate-600">
                  Batal
                </button>
                <button type="submit" className="px-4 py-2 rounded-xl bg-blue-600 text-white text-sm font-semibold">
                  Simpan
                </button>
              </div>
            </div>
          </form>
        </div>
      )}

      {/* MODAL TAMBAH ROMBEL PARALEL (3 KETUKAN SEPERTI VIDEO) */}
      {showAddRombel && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <form onSubmit={handleAddRombelParalel} className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl space-y-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Layers className="w-5 h-5 text-blue-600" />
                <h3 className="font-bold text-slate-900">Tambah Rombel Paralel (3 Ketukan)</h3>
              </div>
              <button type="button" onClick={() => setShowAddRombel(false)} className="text-slate-400 hover:text-slate-600">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Jenjang / Fase</label>
                <select
                  value={rombelParalelForm.fase}
                  onChange={(e) => {
                    const found = FASE_LIST.find((f) => f.label === e.target.value) || FASE_LIST[3];
                    setRombelParalelForm({
                      ...rombelParalelForm,
                      fase: found.label,
                      kelas: found.defaultKelas,
                      mapel: found.mapels[0]
                    });
                  }}
                  className="w-full rounded-xl border border-slate-200 px-3 py-2 text-xs bg-white"
                >
                  {FASE_LIST.map((f) => (
                    <option key={f.label} value={f.label}>
                      {f.label}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Tingkat Kelas</label>
                <select
                  value={rombelParalelForm.kelas}
                  onChange={(e) => setRombelParalelForm({ ...rombelParalelForm, kelas: e.target.value })}
                  className="w-full rounded-xl border border-slate-200 px-3 py-2 text-xs bg-white"
                >
                  {['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'].map((k) => (
                    <option key={k} value={k}>
                      Kelas {k}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">Mata Pelajaran</label>
              <select
                value={rombelParalelForm.mapel}
                onChange={(e) => setRombelParalelForm({ ...rombelParalelForm, mapel: e.target.value })}
                className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm bg-white"
              >
                {selectedRombelFaseObj.mapels.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-2">Ketuk Huruf Rombel Paralel yang Diajar:</label>
              <div className="grid grid-cols-3 gap-2">
                {['A', 'B', 'C', 'D', 'E', 'F'].map((letter) => {
                  const selected = rombelParalelForm.selectedLetters.includes(letter);
                  return (
                    <button
                      key={letter}
                      type="button"
                      onClick={() => {
                        const updated = selected
                          ? rombelParalelForm.selectedLetters.filter((l) => l !== letter)
                          : [...rombelParalelForm.selectedLetters, letter].sort();
                        setRombelParalelForm({ ...rombelParalelForm, selectedLetters: updated });
                      }}
                      className={`py-2.5 rounded-xl text-xs font-bold border transition cursor-pointer ${
                        selected
                          ? 'bg-blue-600 text-white border-blue-600 shadow-xs'
                          : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                      }`}
                    >
                      Kelas {rombelParalelForm.kelas} {letter}
                    </button>
                  );
                })}
              </div>
            </div>

            <button
              type="submit"
              className="w-full py-2.5 rounded-xl bg-blue-600 text-white text-sm font-bold hover:bg-blue-700 cursor-pointer"
            >
              Tambah {rombelParalelForm.selectedLetters.length} Rombel Paralel
            </button>
          </form>
        </div>
      )}
    </div>
  );
}
