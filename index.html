<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <title>CBT Promax Assessment - SMP Negeri 6 Muara Kaman</title>
  <!-- PWA Manifest Data URI -->
  <link rel="manifest" href='data:application/manifest+json;charset=utf-8,%7B%22name%22%3A%22CBT%20Promax%20Assessment%22%2C%22short_name%22%3A%22CBT%20Promax%22%2C%22start_url%22%3A%22.%22%2C%22display%22%3A%22standalone%22%2C%22background_color%22%3A%22%23ffffff%22%2C%22theme_color%22%3A%22%235d7359%22%7D' />
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Google Fonts: Plus Jakarta Sans -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <!-- SheetJS (Excel), Mammoth (Word .docx Import) & JSZip (Template .docx Generator) -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>

  <style>
    /* Full Page Blogger Widget Overlay */
    #cbt-promax-root {
      position: fixed;
      inset: 0;
      z-index: 999999;
      background: linear-gradient(135deg, #f3faf6 0%, #f4f8fc 50%, #fdfaf3 100%);
      font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
      color: #1e293b;
      overflow-y: auto;
      overflow-x: hidden;
      -webkit-tap-highlight-color: transparent;
    }

    /* Glassmorphism Utilities */
    .glass-panel {
      background: rgba(255, 255, 255, 0.92);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border: 1px solid rgba(203, 224, 214, 0.85);
      box-shadow: 0 4px 20px -2px rgba(15, 23, 42, 0.04);
    }

    .glass-sidebar {
      background: linear-gradient(180deg, rgba(242, 250, 246, 0.98) 0%, rgba(243, 248, 253, 0.98) 55%, rgba(253, 249, 241, 0.98) 100%);
      backdrop-filter: blur(18px);
      -webkit-backdrop-filter: blur(18px);
      border-right: 1px solid rgba(203, 224, 214, 0.9);
      box-shadow: 4px 0 24px rgba(15, 23, 42, 0.06);
      transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1);
    }

    /* Custom Scrollbar */
    #cbt-promax-root ::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    #cbt-promax-root ::-webkit-scrollbar-track {
      background: transparent;
    }
    #cbt-promax-root ::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 9999px;
    }

    /* Responsive Automatic Left Sidebar (Desktop & HP/Tablet) */
    .app-sidebar {
      display: flex !important;
      transform: translateX(-100%);
    }
    .app-sidebar.is-open {
      transform: translateX(0);
    }
    .main-with-sidebar {
      margin-left: 0;
      padding-bottom: 2.5rem;
      transition: margin-left 0.25s cubic-bezier(0.4, 0, 0.2, 1);
    }
    @media (min-width: 1024px) {
      .main-with-sidebar.sidebar-docked {
        margin-left: 264px;
      }
    }

    /* Anti-Cheat & Anti-Swipe / Pull Lock During Active Exam */
    html.exam-active-lock,
    body.exam-active-lock,
    #cbt-promax-root.exam-active-lock {
      overscroll-behavior: none !important;
      overscroll-behavior-y: none !important;
      overscroll-behavior-x: none !important;
      touch-action: pan-y !important;
    }

    .exam-protected {
      -webkit-touch-callout: none;
      overscroll-behavior: none;
      touch-action: pan-y;
    }

    /* Izinkan siswa mengetik & memilih teks hanya di dalam kotak jawaban Isian Singkat & Uraian */
    .exam-protected .exam-question-read {
      -webkit-user-select: none;
      -moz-user-select: none;
      user-select: none;
    }

    .exam-protected input,
    .exam-protected textarea,
    .exam-protected select {
      -webkit-user-select: text !important;
      user-select: text !important;
      touch-action: manipulation !important;
    }

    .exam-watermark-layer {
      position: fixed;
      inset: 0;
      pointer-events: none;
      z-index: 40;
      overflow: hidden;
      opacity: 0.045;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-around;
      align-content: space-around;
      transform: rotate(-22deg) scale(1.3);
    }

    /* Security Curtain when Screenshot / 3-Finger / Volume+Power / Locked detected */
    #security-curtain {
      position: fixed;
      inset: 0;
      z-index: 9999999;
      background: #090d16;
      color: #ffffff;
      display: none;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      text-align: center;
      padding: 1.5rem;
      overscroll-behavior: none;
      touch-action: none;
    }

    @media print {
      body * { visibility: hidden !important; }
      #cbt-promax-root { position: static !important; overflow: visible !important; background: #fff !important; }
      #print-area, #print-area * { visibility: visible !important; }
      #print-area { position: absolute !important; left: 0 !important; top: 0 !important; width: 100% !important; display: block !important; }
      .no-print { display: none !important; }
    }
  </style>
</head>
<body>
<div id="cbt-promax-root"></div>

<!-- Tirai Hitam Anti-Screenshot / 3 Jari / Power+Volume / Terkunci -->
<div id="security-curtain">
  <div class="max-w-md w-full bg-red-950/95 border-2 border-red-500 rounded-2xl p-6 shadow-2xl">
    <div id="curtain-icon" class="w-16 h-16 mx-auto mb-4 rounded-full bg-red-600 flex items-center justify-center text-white text-3xl font-extrabold animate-pulse">!</div>
    <h2 id="curtain-title" class="text-xl font-extrabold tracking-tight text-white mb-2">PELANGGARAN KEAMANAN TERDETEKSI</h2>
    <p id="curtain-reason" class="text-sm text-red-200 mb-4 leading-relaxed">
      Aktivitas terlarang terdeteksi. Layar soal dikunci sementara dan pelanggaran telah dicatat ke server pengawas.
    </p>
    <div id="curtain-counter" class="inline-block px-4 py-1.5 rounded-full bg-red-900 text-red-100 text-xs font-bold mb-5">
      Pelanggaran ke-1
    </div>
    <div id="curtain-actions" class="space-y-2.5">
      <button id="btn-dismiss-curtain" type="button" class="w-full py-3 px-4 rounded-xl bg-white text-red-900 font-bold text-sm hover:bg-red-50 transition">
        Kembali ke Halaman Ujian (Layar Penuh)
      </button>
      <button id="btn-check-unlock" type="button" class="hidden w-full py-3 px-4 rounded-xl bg-emerald-500 text-slate-950 font-extrabold text-sm hover:bg-emerald-400 transition">
        ↻ Cek Status Buka Kunci &amp; Lanjutkan Ujian
      </button>
    </div>
  </div>
</div>

<script type="module">
  import { createClient } from 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm';

  // =========================================================================
  // KONFIGURASI DEPLOYMENT APLIKASI (UBAH IS_PREVIEW = false SAAT PRODUKSI)
  // =========================================================================
  const IS_PREVIEW = false; // Ubah menjadi false SAAT DEPLOY agar terhubung 100% ke Supabase
  const SUPABASE_URL = "https://hmcwbwitmlacyitcwdbs.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_TE1XQijD61mHvlC5jt5qsQ_FLu6Z-j1";
  const ADMIN_AUTH_EMAIL = "smpn6muarakaman@gmail.com"; // Email Supabase Auth untuk login username: admin

  const isSupabaseConfigured = !IS_PREVIEW && /^https?:\/\//i.test(SUPABASE_URL) && SUPABASE_ANON_KEY.length > 10;
  const supabase = isSupabaseConfigured ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

  // =========================================================================
  // REGISTRASI PWA SERVICE WORKER DINAMIS (1 FILE - AMAN DI IFRAME/HTTPS)
  // =========================================================================
  if ('serviceWorker' in navigator && (window.location.protocol === 'https:' || window.location.hostname === 'localhost')) {
    try {
      const swCode = `
        const CACHE_NAME = 'cbt-promax-v3.1';
        self.addEventListener('install', () => { self.skipWaiting(); });
        self.addEventListener('activate', (e) => { e.waitUntil(self.clients.claim()); });
        self.addEventListener('fetch', (e) => { e.respondWith(fetch(e.request).catch(() => new Response('Offline'))); });
      `;
      const blob = new Blob([swCode], { type: 'application/javascript' });
      const swUrl = URL.createObjectURL(blob);
      navigator.serviceWorker.register(swUrl).catch(() => {});
    } catch (err) {}
  }

  // =========================================================================
  // DATA DUMMY MINIMALIS (HANYA AKTIF JIKA IS_PREVIEW === true)
  // =========================================================================
  const PREVIEW_INITIAL_DB = IS_PREVIEW ? {
    pengaturan_sekolah: [{
      id: 'cfg-1',
      nama_sekolah: 'SMP NEGERI 6 MUARA KAMAN',
      npsn: '30407337',
      alamat: 'Jln. Batu Onggos Rt. 006 Desa Teratak',
      tahun_pelajaran: '2026/2027',
      semester: 'Ganjil',
      nama_kepsek: 'Sukriansyah, S.Pd',
      nip_kepsek: '19701210 200801 1 019',
      nama_proktor: 'Galih Rakasiwi, S.Pd',
      nip_proktor: '19830523 201410 1 002',
      nama_pengawas: 'Esdiansyah, S.Pd',
      nip_pengawas: '19820706 201410 1 003',
      max_pelanggaran: 3,
      logo_dinas_url: '',
      logo_sekolah_url: '',
      logo_aplikasi_url: ''
    }],
    users_cbt: [
      {
        id: 'u-guru-1',
        nomor_peserta: '-',
        username: 'guruips',
        password: '123',
        nama_lengkap: 'Ani Fitriani, S.Pd',
        role: 'guru',
        kelas: '-',
        agama: 'Islam',
        ruang_sesi: '-',
        nip_nisn: '198705062022212003',
        is_logged_in: false
      },
      {
        id: 'u-siswa-1',
        nomor_peserta: '26-2697-0001-9',
        username: '*2026',
        password: '*1001',
        nama_lengkap: 'Algates',
        role: 'siswa',
        kelas: '7a',
        agama: 'Islam',
        ruang_sesi: 'Ruang 01 / Lab Komputer 1 / Sesi 1',
        nip_nisn: '0143042697',
        is_logged_in: false
      }
    ],
    mata_pelajaran: [
      {
        id: 'm-1',
        kode_mapel: 'IPS-07',
        nama_mapel: 'IPS',
        kelas_ampu: '7a, 7b',
        kategori_agama: '',
        is_english: false,
        guru_id: 'u-guru-1'
      }
    ],
    bank_soal: [
      {
        id: 's-1',
        mapel_id: 'm-1',
        kelas_target: '7a, 7b',
        nomor_urut: 1,
        tipe_soal: 'PILIHAN GANDA',
        pertanyaan: 'Setiap pagi, Pak Rahmat membeli sayuran segar langsung dari para petani di desa, lalu mengangkutnya menggunakan mobil bak terbuka untuk dijual kembali kepada para pedagang dan warga di pasar kota. Kegiatan ekonomi yang dilakukan oleh Pak Rahmat disebut ....',
        gambar_url: '',
        opsi_a: 'Retribusi',
        opsi_b: 'Konsumsi',
        opsi_c: 'Distribusi',
        opsi_d: 'Produksi',
        pasangan_json: [],
        kunci_jawaban: 'C',
        bobot: 20,
        tingkat: 'Sedang',
        topik: 'Kegiatan Ekonomi',
        pembahasan: ''
      },
      {
        id: 's-2',
        mapel_id: 'm-1',
        kelas_target: '7a, 7b',
        nomor_urut: 2,
        tipe_soal: 'PILIHAN GANDA KOMPLEKS',
        pertanyaan: 'Manakah di antara pernyataan berikut yang merupakan ciri-ciri pelaku kegiatan distribusi? (Pilih lebih dari satu jawaban yang benar)',
        gambar_url: '',
        opsi_a: 'Menyalurkan barang dari produsen ke konsumen',
        opsi_b: 'Menghabiskan nilai guna suatu barang secara langsung',
        opsi_c: 'Memperlancar arus ketersediaan barang di masyarakat',
        opsi_d: 'Mengolah bahan mentah menjadi barang setengah jadi',
        pasangan_json: [],
        kunci_jawaban: 'A, C',
        bobot: 20,
        tingkat: 'Sedang',
        topik: 'Kegiatan Ekonomi',
        pembahasan: ''
      },
      {
        id: 's-3',
        mapel_id: 'm-1',
        kelas_target: '7a, 7b',
        nomor_urut: 3,
        tipe_soal: 'MENJODOHKAN',
        pertanyaan: 'Pasangkan jenis kegiatan ekonomi di sebelah kiri dengan pelaku atau contoh kegiatannya yang paling tepat di sebelah kanan!',
        gambar_url: '',
        opsi_a: '',
        opsi_b: '',
        opsi_c: '',
        opsi_d: '',
        pasangan_json: [
          { kiri: 'Produksi', kanan: 'Pabrik roti mengolah tepung terigu menjadi roti siap jual' },
          { kiri: 'Distribusi', kanan: 'Agen ekspedisi mengirim paket barang ke berbagai kota' },
          { kiri: 'Konsumsi', kanan: 'Siswa membeli dan memakan sarapan di kantin sekolah' }
        ],
        kunci_jawaban: 'SESUAI PASANGAN',
        bobot: 30,
        tingkat: 'Sedang',
        topik: 'Kegiatan Ekonomi',
        pembahasan: ''
      },
      {
        id: 's-4',
        mapel_id: 'm-1',
        kelas_target: '7a, 7b',
        nomor_urut: 4,
        tipe_soal: 'URAIAN',
        pertanyaan: 'Jelaskan mengapa kegiatan distribusi memegang peranan penting dalam menjaga kestabilan harga kebutuhan pokok di daerah terpencil!',
        gambar_url: '',
        opsi_a: '',
        opsi_b: '',
        opsi_c: '',
        opsi_d: '',
        pasangan_json: [],
        kunci_jawaban: 'Karena distribusi memastikan pasokan barang merata sehingga mencegah kelangkaan dan lonjakan harga.',
        bobot: 30,
        tingkat: 'Sulit',
        topik: 'Kegiatan Ekonomi',
        pembahasan: ''
      }
    ],
    jadwal_ujian: [
      {
        id: 'j-1',
        mapel_id: 'm-1',
        judul_ujian: 'Ujian Simulasi',
        kelas_target: '7a, 7b',
        waktu_mulai: new Date(Date.now() - 3600000).toISOString(),
        waktu_selesai: new Date(Date.now() + 86400000).toISOString(),
        durasi_menit: 90,
        jumlah_soal: 4,
        soal_ids_json: ['s-1', 's-2', 's-3', 's-4'],
        token_ujian: 'SIM7A1',
        acak_soal: false,
        acak_opsi: false,
        status: 'Aktif'
      }
    ],
    sesi_ujian_siswa: []
  } : {
    pengaturan_sekolah: [],
    users_cbt: [],
    mata_pelajaran: [],
    bank_soal: [],
    jadwal_ujian: [],
    sesi_ujian_siswa: []
  };

  // =========================================================================
  // STATE MANAGEMENT GLOBAL
  // =========================================================================
  const state = {
    currentUser: null,
    activeMenu: 'dashboard',
    sidebarOpen: typeof window !== 'undefined' ? window.innerWidth >= 1024 : true,
    loading: false,
    toast: null,
    toastTimer: null,
    confirmDialog: null,
    db: JSON.parse(JSON.stringify(PREVIEW_INITIAL_DB)),

    selectedIds: {
      soal: new Set(),
      siswa: new Set(),
      guru: new Set(),
      mapel: new Set(),
      jadwal: new Set(),
      nilai: new Set()
    },
    filterSoalMapel: '',
    filterSoalKelas: '',
    filterMonitoringJadwal: '',
    filterMonitoringKelas: '',
    filterMonitoringSearch: '',
    filterUserTab: 'siswa',
    filterUserKelas: '',
    filterNilaiJadwal: '',
    filterNilaiKelas: '',
    docType: 'daftar_hadir',
    docKelas: 'SEMUA',
    docJadwalId: '',
    editingBeritaAcara: false,
    beritaAcaraForm: null,

    modal: null,
    modalSoalDraft: null,
    modalSoalImagePreview: '',
    modalSoalSelectedKelas: new Set(),
    modalMapelSelectedKelas: new Set(),
    modalImportMapelId: '',
    modalImportSelectedKelas: new Set(),
    modalJadwalMapelId: '',
    modalJadwalSelectedSoal: new Set(),
    modalLogos: { dinas: undefined, sekolah: undefined, aplikasi: undefined },

    exam: {
      active: false,
      locked: false,
      needsFullscreenTap: false,
      isSubmitting: false,
      finishedHandled: false,
      wasFullscreen: false,
      graceUntil: 0,
      lastViolationAt: 0,
      jadwal: null,
      mapel: null,
      soalList: [],
      optionOrders: {},
      currentIndex: 0,
      jawaban: {},
      ragu: {},
      sesiRecord: null,
      remainingSeconds: 0,
      fontSize: 'base',
      showDaftarSoalMobile: false,
      timerInterval: null,
      wakeLock: null,
      watchdogInterval: null,
      domObserver: null,
      navLockActive: false
    }
  };

  // =========================================================================
  // FUNGSI UTILITAS INTI, NORMALISASI TIPE SOAL, PASANGAN & JAWABAN
  // =========================================================================
  const SCHOOL_CFG_CACHE_KEY = 'cbt_promax_school_cfg_v1';

  function saveSchoolConfigCache(cfgObj) {
    try { localStorage.setItem(SCHOOL_CFG_CACHE_KEY, JSON.stringify(cfgObj)); } catch (e) {}
  }

  function loadSchoolConfigCache() {
    try {
      const raw = localStorage.getItem(SCHOOL_CFG_CACHE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  }

  function generateLocalId(prefix = 'id') {
    return `${prefix}-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`;
  }

  const syncChannel = typeof BroadcastChannel !== 'undefined' ? new BroadcastChannel('cbt_promax_sync_v1') : null;
  let isFastSyncing = false;
  let lastFastSyncHash = '';

  function broadcastLocalSync(table) {
    try {
      if (syncChannel) syncChannel.postMessage({ table, ts: Date.now() });
    } catch (e) {}
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&#8226;?|&bull;|\u2022/g, '|')
      .replace(/&#8212;?|&mdash;|\u2014/g, '-')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // Normalisasi nama tipe soal agar tidak pernah gagal baca meski ada variasi ejaan/kapitalisasi
  function normalizeTipeSoal(rawTipe) {
    const up = String(rawTipe || 'PILIHAN GANDA').trim().toUpperCase().replace(/[_-]+/g, ' ');
    if (up.includes('KOMPLEK')) return 'PILIHAN GANDA KOMPLEKS';
    if (up.includes('BENAR') && up.includes('SALAH')) return 'BENAR SALAH';
    if (up.includes('JODOH')) return 'MENJODOHKAN';
    if (up.includes('ISIAN') || up.includes('SINGKAT')) return 'ISIAN SINGKAT';
    if (up.includes('URAIAN') || up.includes('ESSAY') || up.includes('ESAI')) return 'URAIAN';
    return 'PILIHAN GANDA';
  }

  // Normalisasi data pasangan menjodohkan (mendukung array, JSON string, atau teks fallback "Kiri = Kanan")
  function normalizePasanganJson(raw, fallbackSoal = null) {
    let list = [];
    if (Array.isArray(raw)) {
      list = raw;
    } else if (typeof raw === 'string' && raw.trim() !== '') {
      try {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) list = parsed;
      } catch (e) {
        // Fallback jika tersimpan sebagai teks baris "A = B"
        list = raw.split('\n').map(line => {
          const [k, ...v] = line.split('=');
          return { kiri: (k || '').trim(), kanan: v.join('=').trim() };
        });
      }
    }

    const cleaned = list
      .map(item => {
        if (!item || typeof item !== 'object') return null;
        const kiri = String(item.kiri ?? item.left ?? item.pertanyaan ?? item.soal ?? '').trim();
        const kanan = String(item.kanan ?? item.right ?? item.jawaban ?? item.pasangan ?? '').trim();
        return (kiri || kanan) ? { kiri, kanan } : null;
      })
      .filter(Boolean);

    // Jika pasangan_json kosong namun opsi_a..d berisi format "Kiri = Kanan", pulihkan otomatis
    if (cleaned.length === 0 && fallbackSoal) {
      ['opsi_a', 'opsi_b', 'opsi_c', 'opsi_d'].forEach(k => {
        const val = String(fallbackSoal[k] || '').trim();
        if (val.includes('=')) {
          const [left, ...right] = val.split('=');
          cleaned.push({ kiri: left.trim(), kanan: right.join('=').trim() });
        }
      });
    }

    return cleaned;
  }

  // Pastikan jawaban Pilihan Ganda Kompleks selalu terbaca sebagai Array huruf ['A', 'C', ...]
  function normalizePgkAnswerArray(rawAns) {
    if (Array.isArray(rawAns)) {
      return [...new Set(rawAns.map(x => String(x || '').trim().toUpperCase()).filter(Boolean))];
    }
    if (typeof rawAns === 'string' && rawAns.trim() !== '') {
      try {
        const parsed = JSON.parse(rawAns);
        if (Array.isArray(parsed)) return normalizePgkAnswerArray(parsed);
      } catch (e) {}
      return [...new Set(rawAns.split(',').map(x => x.trim().toUpperCase()).filter(Boolean))];
    }
    return [];
  }

  // Pastikan jawaban Menjodohkan selalu terbaca sebagai Object { "0": "Teks Kanan", ... }
  function normalizeMatchAnswerObj(rawAns) {
    if (!rawAns) return {};
    if (typeof rawAns === 'object' && !Array.isArray(rawAns)) {
      return { ...rawAns };
    }
    if (typeof rawAns === 'string' && rawAns.trim() !== '') {
      try {
        const parsed = JSON.parse(rawAns);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) return parsed;
      } catch (e) {}
    }
    return {};
  }

  function getSchoolConfig() {
    const dbRow = (state.db.pengaturan_sekolah && state.db.pengaturan_sekolah[0]) || null;
    const cached = loadSchoolConfigCache();
    const row = dbRow || cached || {};
    return {
      id: row.id || 'cfg-1',
      nama_sekolah: row.nama_sekolah || 'SMP NEGERI 6 MUARA KAMAN',
      npsn: row.npsn || '30407337',
      alamat: row.alamat || 'Jln. Batu Onggos Rt. 006 Desa Teratak',
      tahun_pelajaran: row.tahun_pelajaran || '2026/2027',
      semester: row.semester || 'Ganjil',
      nama_kepsek: row.nama_kepsek || 'Sukriansyah, S.Pd',
      nip_kepsek: row.nip_kepsek || '19701210 200801 1 019',
      nama_proktor: row.nama_proktor || 'Galih Rakasiwi, S.Pd',
      nip_proktor: row.nip_proktor || '19830523 201410 1 002',
      nama_pengawas: row.nama_pengawas || 'Esdiansyah, S.Pd',
      nip_pengawas: row.nip_pengawas || '19820706 201410 1 003',
      max_pelanggaran: Number(row.max_pelanggaran) || 3,
      logo_dinas_url: row.logo_dinas_url || (cached?.logo_dinas_url || ''),
      logo_sekolah_url: row.logo_sekolah_url || (cached?.logo_sekolah_url || ''),
      logo_aplikasi_url: row.logo_aplikasi_url || (cached?.logo_aplikasi_url || '')
    };
  }

  function getWitaTimeString() {
    const now = new Date();
    const tgl = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Makassar',
      weekday: 'short',
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    }).format(now);
    const jam = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Makassar',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    }).format(now).replace(/\./g, ':');
    return `${tgl} | ${jam} WITA`;
  }

  function formatDateTimeWita(isoStr) {
    if (!isoStr) return '-';
    const d = new Date(isoStr);
    if (Number.isNaN(d.getTime())) return '-';
    const tgl = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Makassar',
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    }).format(d);
    const jam = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Makassar',
      hour: '2-digit',
      minute: '2-digit',
      hour12: false
    }).format(d).replace(/\./g, ':');
    return `${tgl}, ${jam} WITA`;
  }

  function toDatetimeLocalWita(isoStr) {
    const d = isoStr ? new Date(isoStr) : new Date();
    if (Number.isNaN(d.getTime())) return '';
    const witaMs = d.getTime() + 8 * 3600 * 1000;
    const witaDate = new Date(witaMs);
    return witaDate.toISOString().slice(0, 16);
  }

  function parseWitaDatetimeLocal(localVal) {
    if (!localVal) return new Date().toISOString();
    const normalized = localVal.length === 16 ? `${localVal}:00+08:00` : `${localVal}+08:00`;
    const parsed = new Date(normalized);
    return Number.isNaN(parsed.getTime()) ? new Date().toISOString() : parsed.toISOString();
  }

  function formatRemainingTime(sec) {
    const s = Math.max(0, Math.floor(Number(sec) || 0));
    const hrs = Math.floor(s / 3600);
    const mins = Math.floor((s % 3600) / 60);
    const secs = s % 60;
    const pad = (n) => String(n).padStart(2, '0');
    return `${pad(hrs)}:${pad(mins)}:${pad(secs)}`;
  }

  function isJadwalExpired(jadwal) {
    if (!jadwal || !jadwal.waktu_selesai) return false;
    return Date.now() > new Date(jadwal.waktu_selesai).getTime();
  }

  function getEffectiveJadwalStatus(jadwal) {
    if (!jadwal) return 'Nonaktif';
    if (jadwal.status === 'Selesai' || isJadwalExpired(jadwal)) return 'Selesai';
    return jadwal.status || 'Aktif';
  }

  function parseKelasString(str) {
    if (!str) return [];
    return String(str)
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);
  }

  function buildInitialSoalDraft(existing, defaultMapelId) {
    const mapelId = existing ? existing.mapel_id : defaultMapelId;
    const tipe = existing ? normalizeTipeSoal(existing.tipe_soal) : 'PILIHAN GANDA';
    const rawPasangan = existing ? normalizePasanganJson(existing.pasangan_json, existing) : [];
    const pasangan = rawPasangan.length > 0
      ? rawPasangan
      : [{ kiri: '', kanan: '' }, { kiri: '', kanan: '' }, { kiri: '', kanan: '' }];

    let kunci = existing ? String(existing.kunci_jawaban || '').trim() : '';
    if (tipe === 'BENAR SALAH') {
      const up = kunci.toUpperCase();
      if (up === 'A' || up === 'BENAR') kunci = 'BENAR';
      else if (up === 'B' || up === 'SALAH') kunci = 'SALAH';
      else kunci = 'BENAR';
    } else if (tipe === 'PILIHAN GANDA' && !kunci) {
      kunci = 'A';
    } else if (tipe === 'PILIHAN GANDA KOMPLEKS' && !kunci) {
      kunci = 'A, B';
    }

    return {
      mapel_id: mapelId,
      tipe_soal: tipe,
      nomor_urut: existing ? Number(existing.nomor_urut) || 1 : (state.db.bank_soal.filter(s => s.mapel_id === mapelId).length + 1),
      bobot: existing ? Number(existing.bobot) || 10 : 10,
      pertanyaan: existing ? String(existing.pertanyaan || '') : '',
      opsi_a: existing ? String(existing.opsi_a || '') : '',
      opsi_b: existing ? String(existing.opsi_b || '') : '',
      opsi_c: existing ? String(existing.opsi_c || '') : '',
      opsi_d: existing ? String(existing.opsi_d || '') : '',
      pasangan,
      kunci_jawaban: kunci,
      topik: existing ? String(existing.topik || 'Umum') : 'Umum'
    };
  }

  function syncModalSoalDraftFromDom() {
    const form = document.getElementById('form-save-soal');
    if (!form || !state.modalSoalDraft) return;
    const d = state.modalSoalDraft;
    const prevTipe = normalizeTipeSoal(d.tipe_soal);

    const mapelEl = form.querySelector('[name="mapel_id"]');
    const tipeEl = form.querySelector('[name="tipe_soal"]');
    const nomorEl = form.querySelector('[name="nomor_urut"]');
    const bobotEl = form.querySelector('[name="bobot"]');
    const tanyaEl = form.querySelector('[name="pertanyaan"]');
    const topikEl = form.querySelector('[name="topik"]');

    if (mapelEl) d.mapel_id = mapelEl.value;
    if (tipeEl) d.tipe_soal = normalizeTipeSoal(tipeEl.value);
    if (nomorEl) d.nomor_urut = Number(nomorEl.value) || 1;
    if (bobotEl) d.bobot = Number(bobotEl.value) || 10;
    if (tanyaEl) d.pertanyaan = tanyaEl.value;
    if (topikEl) d.topik = topikEl.value;

    const checkedKelas = [...form.querySelectorAll('.chk-soal-kelas:checked')].map(c => c.value);
    if (form.querySelectorAll('.chk-soal-kelas').length > 0) {
      state.modalSoalSelectedKelas = new Set(checkedKelas);
    }

    if (prevTipe === 'PILIHAN GANDA' || prevTipe === 'PILIHAN GANDA KOMPLEKS') {
      const oa = form.querySelector('[name="opsi_a"]');
      const ob = form.querySelector('[name="opsi_b"]');
      const oc = form.querySelector('[name="opsi_c"]');
      const od = form.querySelector('[name="opsi_d"]');
      if (oa) d.opsi_a = oa.value;
      if (ob) d.opsi_b = ob.value;
      if (oc) d.opsi_c = oc.value;
      if (od) d.opsi_d = od.value;

      if (prevTipe === 'PILIHAN GANDA') {
        const kRadio = form.querySelector('input[name="kunci_pg_radio"]:checked');
        if (kRadio) d.kunci_jawaban = kRadio.value;
      } else {
        const kChecks = [...form.querySelectorAll('.chk-kunci-pgk:checked')].map(c => c.value);
        d.kunci_jawaban = kChecks.join(', ');
      }
    } else if (prevTipe === 'BENAR SALAH') {
      const kBs = form.querySelector('input[name="kunci_bs_radio"]:checked');
      if (kBs) d.kunci_jawaban = kBs.value;
    } else if (prevTipe === 'MENJODOHKAN') {
      const kiriInputs = [...form.querySelectorAll('[data-pasangan-kiri]')];
      const kananInputs = [...form.querySelectorAll('[data-pasangan-kanan]')];
      if (kiriInputs.length > 0) {
        d.pasangan = kiriInputs.map((inp, idx) => ({
          kiri: inp.value,
          kanan: kananInputs[idx] ? kananInputs[idx].value : ''
        }));
      }
    } else if (prevTipe === 'ISIAN SINGKAT' || prevTipe === 'URAIAN') {
      const kText = form.querySelector('[name="kunci_jawaban"]');
      if (kText) d.kunci_jawaban = kText.value;
    }

    // Inisialisasi default kunci saat user baru saja mengganti tipe soal
    if (d.tipe_soal !== prevTipe) {
      if (d.tipe_soal === 'PILIHAN GANDA') {
        if (!['A', 'B', 'C', 'D'].includes(d.kunci_jawaban.toUpperCase())) d.kunci_jawaban = 'A';
      } else if (d.tipe_soal === 'PILIHAN GANDA KOMPLEKS') {
        const valid = d.kunci_jawaban.split(',').map(x => x.trim().toUpperCase()).filter(x => ['A', 'B', 'C', 'D'].includes(x));
        d.kunci_jawaban = valid.length > 0 ? valid.join(', ') : 'A, B';
      } else if (d.tipe_soal === 'BENAR SALAH') {
        d.kunci_jawaban = d.kunci_jawaban.toUpperCase() === 'SALAH' ? 'SALAH' : 'BENAR';
      } else if (d.tipe_soal === 'MENJODOHKAN') {
        if (!Array.isArray(d.pasangan) || d.pasangan.length === 0) {
          d.pasangan = [{ kiri: '', kanan: '' }, { kiri: '', kanan: '' }, { kiri: '', kanan: '' }];
        }
        d.kunci_jawaban = 'SESUAI PASANGAN';
      } else if (d.tipe_soal === 'ISIAN SINGKAT' || d.tipe_soal === 'URAIAN') {
        if (['A', 'B', 'C', 'D', 'BENAR', 'SALAH', 'SESUAI PASANGAN'].includes(d.kunci_jawaban.toUpperCase())) {
          d.kunci_jawaban = '';
        }
      }
    }
  }

  function getAllMasterClasses() {
    const set = new Set(['7a', '7b', '8a', '8b', '9a', '9b']);
    state.db.users_cbt.forEach(u => {
      if (u.role === 'siswa' && u.kelas && u.kelas !== '-') {
        set.add(u.kelas.trim());
      }
    });
    state.db.mata_pelajaran.forEach(m => {
      parseKelasString(m.kelas_ampu).forEach(k => set.add(k));
    });
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }

  function getAuthorizedMapels() {
    if (!state.currentUser) return [];
    if (state.currentUser.role === 'admin') return state.db.mata_pelajaran;
    return state.db.mata_pelajaran.filter(m => m.guru_id === state.currentUser.id);
  }

  function getAuthorizedMapelIds() {
    return new Set(getAuthorizedMapels().map(m => m.id));
  }

  function getAuthorizedKelasList(mapelId) {
    const mp = state.db.mata_pelajaran.find(m => m.id === mapelId);
    const ampu = mp ? parseKelasString(mp.kelas_ampu) : [];
    return ampu.length > 0 ? ampu : getAllMasterClasses();
  }

  function shuffleArray(arr) {
    const copy = [...arr];
    for (let i = copy.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [copy[i], copy[j]] = [copy[j], copy[i]];
    }
    return copy;
  }

  function showToast(message, type = 'success') {
    state.toast = { message, type };
    if (state.toastTimer) clearTimeout(state.toastTimer);
    render();
    state.toastTimer = setTimeout(() => {
      state.toast = null;
      render();
    }, 3400);
  }

  function openConfirmDialog({ title, message, confirmLabel = 'Ya, Lanjutkan', variant = 'primary', onConfirm }) {
    state.confirmDialog = {
      title: title || 'Konfirmasi',
      message: message || 'Apakah Anda yakin?',
      confirmLabel,
      variant,
      onConfirm
    };
    render();
  }

  function compressImageToDataUrl(file, maxDim = 320, quality = 0.82) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (ev) => {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement('canvas');
          let w = img.width;
          let h = img.height;
          if (w > maxDim || h > maxDim) {
            if (w > h) {
              h = Math.round((h * maxDim) / w);
              w = maxDim;
            } else {
              w = Math.round((w * maxDim) / h);
              h = maxDim;
            }
          }
          canvas.width = w;
          canvas.height = h;
          const ctx = canvas.getContext('2d');
          ctx.clearRect(0, 0, w, h);
          ctx.drawImage(img, 0, 0, w, h);
          let dataUrl = '';
          try {
            dataUrl = canvas.toDataURL('image/webp', quality);
          } catch (err) {}
          if (!dataUrl || !dataUrl.startsWith('data:image/webp')) {
            dataUrl = canvas.toDataURL('image/png');
          }
          resolve(dataUrl);
        };
        img.onerror = () => resolve(ev.target.result);
        img.src = ev.target.result;
      };
      reader.readAsDataURL(file);
    });
  }

  function setAntiTranslate(lock) {
    try {
      const htmlEl = document.documentElement;
      if (lock) {
        htmlEl.setAttribute('translate', 'no');
        htmlEl.classList.add('notranslate');
      } else {
        htmlEl.removeAttribute('translate');
        htmlEl.classList.remove('notranslate');
      }
    } catch (e) {}
  }

  let sharedAudioCtx = null;
  function unlockAudioContext() {
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      if (!sharedAudioCtx) sharedAudioCtx = new AudioCtx();
      if (sharedAudioCtx.state === 'suspended') {
        sharedAudioCtx.resume().catch(() => {});
      }
    } catch (e) {}
  }

  function playViolationSound() {
    try {
      unlockAudioContext();
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      const ctx = sharedAudioCtx || new AudioCtx();
      const now = ctx.currentTime;
      const freqs = [880, 620, 960, 620, 1040];
      freqs.forEach((f, idx) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sawtooth';
        osc.frequency.setValueAtTime(f, now + idx * 0.14);
        gain.gain.setValueAtTime(0.24, now + idx * 0.14);
        gain.gain.exponentialRampToValueAtTime(0.001, now + idx * 0.14 + 0.13);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(now + idx * 0.14);
        osc.stop(now + idx * 0.14 + 0.135);
      });
    } catch (e) {}
  }

  function svgLogoPromax(size = 44) {
    return `
      <svg width="${size}" height="${size}" viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect width="64" height="64" rx="16" fill="#4f6d4a"/>
        <path d="M18 33L27 42L46 22" stroke="#ffffff" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round"/>
        <circle cx="48" cy="46" r="6" fill="#fbbf24"/>
      </svg>
    `;
  }

  function svgLogoKukar(size = 54) {
    return `
      <svg width="${size}" height="${size}" viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M32 4L54 14V32C54 46 44 55 32 60C20 55 10 46 10 32V14L32 4Z" fill="#15803d" stroke="#0f172a" stroke-width="2.5"/>
        <circle cx="32" cy="30" r="11" fill="#facc15"/>
        <path d="M22 44H42" stroke="#ffffff" stroke-width="3" stroke-linecap="round"/>
      </svg>
    `;
  }

  function svgLogoSmp(size = 54) {
    return `
      <svg width="${size}" height="${size}" viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
        <circle cx="32" cy="32" r="28" fill="#1e3a8a" stroke="#0f172a" stroke-width="2.5"/>
        <path d="M18 26L32 18L46 26L32 34L18 26Z" fill="#facc15"/>
        <path d="M22 31V41C22 41 27 45 32 45C37 45 42 41 42 41V31" stroke="#ffffff" stroke-width="2.8" stroke-linecap="round"/>
      </svg>
    `;
  }

  function renderKopSurat(cfg) {
    const leftLogo = cfg.logo_dinas_url
      ? `<img src="${cfg.logo_dinas_url}" alt="Logo Dinas" style="width:58px; height:58px; object-fit:contain; display:block; margin:0 auto;" />`
      : svgLogoKukar(56);
    const rightLogo = cfg.logo_sekolah_url
      ? `<img src="${cfg.logo_sekolah_url}" alt="Logo Sekolah" style="width:58px; height:58px; object-fit:contain; display:block; margin:0 auto;" />`
      : svgLogoSmp(56);

    return `
      <table style="width:100%; border-collapse:collapse; border-bottom:3px double #000; margin-bottom:14px;">
        <tr>
          <td style="width:74px; text-align:center; vertical-align:middle; padding-bottom:8px;">${leftLogo}</td>
          <td style="text-align:center; vertical-align:middle; padding:0 6px 8px 6px; line-height:1.28;">
            <div style="font-size:10.5pt; font-weight:700; letter-spacing:0.2px;">PEMERINTAH KABUPATEN KUTAI KARTANEGARA</div>
            <div style="font-size:10.5pt; font-weight:700; letter-spacing:0.2px;">DINAS PENDIDIKAN DAN KEBUDAYAAN</div>
            <div style="font-size:13.5pt; font-weight:800; letter-spacing:0.4px; margin:1px 0;">${escapeHtml(cfg.nama_sekolah)}</div>
            <div style="font-size:8.5pt; font-weight:400;">Alamat: ${escapeHtml(cfg.alamat)} | NPSN: ${escapeHtml(cfg.npsn)}</div>
          </td>
          <td style="width:74px; text-align:center; vertical-align:middle; padding-bottom:8px;">${rightLogo}</td>
        </tr>
      </table>
    `;
  }

  async function saveSchoolConfigRecord(payload) {
    if (IS_PREVIEW || !supabase) {
      let saved;
      if (state.db.pengaturan_sekolah.length > 0) {
        saved = await dbUpdate('pengaturan_sekolah', state.db.pengaturan_sekolah[0].id, payload);
      } else {
        saved = await dbInsert('pengaturan_sekolah', payload);
      }
      saveSchoolConfigCache(saved || payload);
      return saved;
    }

    const { data: existingRows, error: fetchErr } = await supabase
      .from('pengaturan_sekolah')
      .select('id')
      .limit(1);
    if (fetchErr) throw fetchErr;

    let savedRow;
    if (existingRows && existingRows.length > 0) {
      const targetId = existingRows[0].id;
      const { data, error } = await supabase
        .from('pengaturan_sekolah')
        .update(payload)
        .eq('id', targetId)
        .select()
        .single();
      if (error) throw error;
      savedRow = data;
    } else {
      const { data, error } = await supabase
        .from('pengaturan_sekolah')
        .insert([payload])
        .select()
        .single();
      if (error) throw error;
      savedRow = data;
    }

    state.db.pengaturan_sekolah = [savedRow];
    saveSchoolConfigCache(savedRow);
    broadcastLocalSync('pengaturan_sekolah');
    return savedRow;
  }

  function isUserInteractingWithForm() {
    if (state.modal || state.confirmDialog) return true;
    const activeEl = document.activeElement;
    if (!activeEl) return false;
    const tag = (activeEl.tagName || '').toUpperCase();
    return tag === 'SELECT' || tag === 'INPUT' || tag === 'TEXTAREA';
  }

  function handleRemoteSessionUnlocked(updatedSesi) {
    Object.assign(state.exam.sesiRecord, updatedSesi);
    if (state.exam.locked && updatedSesi.status !== 'terkunci') {
      state.exam.locked = false;
      state.exam.graceUntil = Date.now() + 4500;
      state.exam.lastViolationAt = 0;
      const curtain = document.getElementById('security-curtain');
      if (curtain) curtain.style.display = 'none';
      safeRequestFullscreen();
      setTimeout(() => {
        if (state.exam.active && !state.exam.locked && !document.fullscreenElement) {
          state.exam.needsFullscreenTap = true;
          render();
        }
      }, 350);
      showToast('Kunci ujian telah dibuka oleh pengawas! Silakan lanjutkan.');
      render();
    }
  }

  function applyRealtimeRowChange(table, eventType, newRow, oldRow) {
    if (!state.db[table]) return;
    if (eventType === 'INSERT' && newRow) {
      const exists = state.db[table].some(item => item.id === newRow.id);
      if (!exists) state.db[table].push(newRow);
      else state.db[table] = state.db[table].map(item => item.id === newRow.id ? newRow : item);
    } else if (eventType === 'UPDATE' && newRow) {
      const idx = state.db[table].findIndex(item => item.id === newRow.id);
      if (idx !== -1) state.db[table][idx] = newRow;
      else state.db[table].push(newRow);
    } else if (eventType === 'DELETE' && oldRow) {
      state.db[table] = state.db[table].filter(item => item.id !== oldRow.id);
    }

    if (state.exam.active && state.exam.sesiRecord && table === 'sesi_ujian_siswa') {
      const updatedSesi = state.db.sesi_ujian_siswa.find(s => s.id === state.exam.sesiRecord.id);
      if (updatedSesi) {
        handleRemoteSessionUnlocked(updatedSesi);
        return;
      }
    }

    if (state.exam.active && state.exam.jadwal && table === 'jadwal_ujian' && newRow?.id === state.exam.jadwal.id) {
      state.exam.jadwal = { ...state.exam.jadwal, ...newRow };
    }

    if (!state.exam.active && !isUserInteractingWithForm()) {
      render();
    }
  }

  function setupRealtimeSync() {
    if (syncChannel) {
      syncChannel.onmessage = () => {
        if (!IS_PREVIEW && supabase) {
          syncFastOperationalData(true);
        }
      };
    }

    if (IS_PREVIEW || !supabase) return;
    try {
      const tablesToWatch = ['jadwal_ujian', 'sesi_ujian_siswa', 'users_cbt', 'mata_pelajaran', 'pengaturan_sekolah', 'bank_soal'];
      const channel = supabase.channel('cbt-promax-realtime-fast');
      tablesToWatch.forEach(tbl => {
        channel.on('postgres_changes', { event: '*', schema: 'public', table: tbl }, (payload) => {
          applyRealtimeRowChange(tbl, payload.eventType, payload.new, payload.old);
        });
      });
      channel.subscribe();
    } catch (e) {}
  }

  async function syncFastOperationalData(forceRender = false) {
    if (IS_PREVIEW || !supabase || isFastSyncing) return;
    isFastSyncing = true;
    try {
      const [usersRes, mapelRes, soalRes, jadwalRes, sesiRes] = await Promise.all([
        supabase.from('users_cbt').select('*').order('nama_lengkap', { ascending: true }),
        supabase.from('mata_pelajaran').select('*').order('nama_mapel', { ascending: true }),
        supabase.from('bank_soal').select('*').order('nomor_urut', { ascending: true }),
        supabase.from('jadwal_ujian').select('*').order('created_at', { ascending: false }),
        supabase.from('sesi_ujian_siswa').select('*').order('waktu_mulai', { ascending: false })
      ]);

      if (!usersRes.error && usersRes.data) state.db.users_cbt = usersRes.data;
      if (!mapelRes.error && mapelRes.data) state.db.mata_pelajaran = mapelRes.data;
      if (!soalRes.error && soalRes.data) state.db.bank_soal = soalRes.data;
      if (!jadwalRes.error && jadwalRes.data) state.db.jadwal_ujian = jadwalRes.data;
      if (!sesiRes.error && sesiRes.data) state.db.sesi_ujian_siswa = sesiRes.data;

      if (state.exam.active && state.exam.locked && state.exam.sesiRecord) {
        const latestSesi = state.db.sesi_ujian_siswa.find(s => s.id === state.exam.sesiRecord.id);
        if (latestSesi && latestSesi.status !== 'terkunci') {
          handleRemoteSessionUnlocked(latestSesi);
          return;
        }
      }

      const nextHash = JSON.stringify({
        j: state.db.jadwal_ujian.map(x => [x.id, x.status, x.waktu_mulai, x.waktu_selesai, x.token_ujian, x.acak_soal, x.acak_opsi]),
        s: state.db.sesi_ujian_siswa.map(x => [x.id, x.status, x.jumlah_pelanggaran, x.nilai_akhir, Object.keys(x.jawaban_json || {}).length]),
        u: state.db.users_cbt.map(x => [x.id, x.is_logged_in, x.kelas])
      });

      if ((forceRender || nextHash !== lastFastSyncHash) && !state.exam.active && !isUserInteractingWithForm()) {
        lastFastSyncHash = nextHash;
        render();
      } else {
        lastFastSyncHash = nextHash;
      }
    } catch (e) {
    } finally {
      isFastSyncing = false;
    }
  }

  async function loadAllData() {
    if (IS_PREVIEW || !supabase) return;
    try {
      const [cfgRes, usersRes, mapelRes, soalRes, jadwalRes, sesiRes] = await Promise.all([
        supabase.from('pengaturan_sekolah').select('*').limit(1),
        supabase.from('users_cbt').select('*').order('nama_lengkap', { ascending: true }),
        supabase.from('mata_pelajaran').select('*').order('nama_mapel', { ascending: true }),
        supabase.from('bank_soal').select('*').order('nomor_urut', { ascending: true }),
        supabase.from('jadwal_ujian').select('*').order('created_at', { ascending: false }),
        supabase.from('sesi_ujian_siswa').select('*').order('waktu_mulai', { ascending: false })
      ]);

      if (!cfgRes.error && cfgRes.data) state.db.pengaturan_sekolah = cfgRes.data;
      if (!usersRes.error && usersRes.data) state.db.users_cbt = usersRes.data;
      if (!mapelRes.error && mapelRes.data) state.db.mata_pelajaran = mapelRes.data;
      if (!soalRes.error && soalRes.data) state.db.bank_soal = soalRes.data;
      if (!jadwalRes.error && jadwalRes.data) state.db.jadwal_ujian = jadwalRes.data;
      if (!sesiRes.error && sesiRes.data) state.db.sesi_ujian_siswa = sesiRes.data;
    } catch (err) {
      console.error('Gagal memuat data Supabase:', err);
    }
  }

  async function dbInsert(table, payload) {
    if (IS_PREVIEW || !supabase) {
      const row = { id: generateLocalId(table.slice(0, 3)), created_at: new Date().toISOString(), ...payload };
      state.db[table].push(row);
      return row;
    }
    const { data, error } = await supabase.from(table).insert([payload]).select().single();
    if (error) throw error;
    state.db[table].push(data);
    broadcastLocalSync(table);
    return data;
  }

  async function dbUpdate(table, id, payload) {
    if (IS_PREVIEW || !supabase) {
      const idx = state.db[table].findIndex(item => item.id === id);
      if (idx !== -1) {
        state.db[table][idx] = { ...state.db[table][idx], ...payload };
        return state.db[table][idx];
      }
      return null;
    }
    const { data, error } = await supabase.from(table).update(payload).eq('id', id).select().single();
    if (error) throw error;
    const idx = state.db[table].findIndex(item => item.id === id);
    if (idx !== -1) state.db[table][idx] = data;
    broadcastLocalSync(table);
    return data;
  }

  async function dbDeleteMany(table, ids) {
    if (!Array.isArray(ids) || ids.length === 0) return;
    const idSet = new Set(ids);
    if (IS_PREVIEW || !supabase) {
      state.db[table] = state.db[table].filter(item => !idSet.has(item.id));
      if (table === 'mata_pelajaran') {
        state.db.bank_soal = state.db.bank_soal.filter(s => !idSet.has(s.mapel_id));
        state.db.jadwal_ujian = state.db.jadwal_ujian.filter(j => !idSet.has(j.mapel_id));
      } else if (table === 'jadwal_ujian') {
        state.db.sesi_ujian_siswa = state.db.sesi_ujian_siswa.filter(s => !idSet.has(s.jadwal_id));
      }
      return;
    }
    const { error } = await supabase.from(table).delete().in('id', ids);
    if (error) throw error;
    state.db[table] = state.db[table].filter(item => !idSet.has(item.id));
    broadcastLocalSync(table);
  }

  // =========================================================================
  // AUTENTIKASI LOGIN & LOGOUT (HYBRID ADMIN AUTH + GURU/SISWA DB)
  // =========================================================================
  async function handleLogin(rawUsername, rawPassword) {
    const username = (rawUsername || '').trim();
    const password = (rawPassword || '').trim();
    if (!username || !password) {
      showToast('Masukkan username dan password Anda.', 'error');
      return;
    }

    state.loading = true;
    render();

    try {
      if (!IS_PREVIEW && !isSupabaseConfigured) {
        throw new Error('Konfigurasi SUPABASE_URL dan SUPABASE_ANON_KEY belum diisi dengan benar.');
      }

      if (username.toLowerCase() === 'admin') {
        if (IS_PREVIEW || !supabase) {
          if (password === 'admincbt') {
            state.currentUser = {
              id: 'admin-root',
              username: 'admin',
              nama_lengkap: 'Administrator CBT Promax',
              role: 'admin',
              kelas: '-',
              agama: 'Islam',
              nomor_peserta: '-'
            };
            state.activeMenu = 'dashboard';
            showToast('Selamat datang, Administrator!');
          } else {
            showToast('Password Administrator salah!', 'error');
          }
          state.loading = false;
          render();
          return;
        } else {
          const { data: authData, error: authErr } = await supabase.auth.signInWithPassword({
            email: ADMIN_AUTH_EMAIL,
            password
          });
          if (authErr || !authData?.user) {
            showToast('Login Admin gagal! Pastikan akun admin@cbtpromax.sch.id terdaftar di Supabase Auth.', 'error');
            state.loading = false;
            render();
            return;
          }
          state.currentUser = {
            id: authData.user.id,
            username: 'admin',
            nama_lengkap: 'Administrator CBT Promax',
            role: 'admin',
            kelas: '-',
            agama: 'Islam',
            nomor_peserta: '-'
          };
          await loadAllData();
          state.activeMenu = 'dashboard';
          showToast('Berhasil masuk sebagai Administrator!');
          state.loading = false;
          render();
          return;
        }
      }

      let foundUser = null;
      if (IS_PREVIEW || !supabase) {
        foundUser = state.db.users_cbt.find(u => u.username === username && u.password === password);
      } else {
        const { data, error } = await supabase
          .from('users_cbt')
          .select('*')
          .eq('username', username)
          .eq('password', password)
          .limit(1)
          .maybeSingle();
        if (error) throw error;
        foundUser = data;
      }

      if (!foundUser) {
        showToast('Username atau Password tidak ditemukan!', 'error');
        state.loading = false;
        render();
        return;
      }

      if (foundUser.role === 'siswa' && foundUser.is_logged_in) {
        showToast('Akun sedang aktif di perangkat lain! Hubungi Proktor/Pengawas untuk Reset Login.', 'error');
        state.loading = false;
        render();
        return;
      }

      if (foundUser.role === 'siswa') {
        await dbUpdate('users_cbt', foundUser.id, { is_logged_in: true });
        foundUser.is_logged_in = true;
      }

      if (!IS_PREVIEW && supabase) {
        await loadAllData();
      }

      state.currentUser = foundUser;
      state.activeMenu = foundUser.role === 'siswa' ? 'dashboard_siswa' : 'dashboard';
      showToast(`Selamat datang, ${foundUser.nama_lengkap}!`);
    } catch (err) {
      showToast('Terjadi kesalahan saat masuk: ' + (err.message || err), 'error');
    } finally {
      state.loading = false;
      render();
    }
  }

  async function handleLogout() {
    cleanupExamEnvironment();
    if (state.currentUser) {
      if (state.currentUser.role === 'siswa') {
        try {
          await dbUpdate('users_cbt', state.currentUser.id, { is_logged_in: false });
        } catch (e) {}
      } else if (state.currentUser.role === 'admin' && !IS_PREVIEW && supabase) {
        try { await supabase.auth.signOut(); } catch (e) {}
      }
    }
    state.currentUser = null;
    state.activeMenu = 'dashboard';
    render();
  }

  // =========================================================================
  // KEAMANAN ANTI-CURANG (ANTI-AI, ANTI-FLOATING, ANTI-EMULATOR, ANTI-SWIPE)
  // =========================================================================
  function detectVirtualAndroidOrEmulator() {
    try {
      const ua = (navigator.userAgent || '').toLowerCase();
      const emuSignatures = [
        'bluestacks', 'nox', 'ldplayer', 'memu', 'genymotion', 'andy',
        'droid4x', 'koplayer', 'remixos', 'phoenixos', 'vmos',
        'parallel space', 'dual space', 'virtualxposed', 'x86_64', 'i686'
      ];
      if (navigator.webdriver) return 'Browser Otomatis / WebDriver terdeteksi';
      for (const sig of emuSignatures) {
        if (ua.includes(sig)) return `Lingkungan Virtual / Emulator (${sig}) terdeteksi`;
      }
      const canvas = document.createElement('canvas');
      const gl = canvas.getContext('webgl') || canvas.getContext('experimental-webgl');
      if (gl) {
        const dbg = gl.getExtension('WEBGL_debug_renderer_info');
        if (dbg) {
          const renderer = String(gl.getParameter(dbg.UNMASKED_RENDERER_WEBGL) || '').toLowerCase();
          if (renderer.includes('swiftshader') || renderer.includes('llvmpipe') || renderer.includes('virtualbox') || renderer.includes('vmware')) {
            return `Renderer Virtual (${renderer}) terdeteksi`;
          }
        }
      }
    } catch (e) {}
    return null;
  }

  function scanAndNeutralizeAiExtensions() {
    if (!state.exam.active || state.exam.isSubmitting || state.exam.finishedHandled) return;
    const aiSelectors = [
      '[id*="monica"]', '[class*="monica"]',
      '[id*="sider"]', '[class*="sider"]',
      '[id*="merlin"]', '[class*="merlin"]',
      '[id*="harpa"]', '[class*="harpa"]',
      '[id*="maxai"]', '[class*="maxai"]',
      '[id*="chatgpt"]', '[id*="copilot"]', '[id*="perplexity"]', '[id*="cici-ai"]'
    ];
    aiSelectors.forEach(sel => {
      document.querySelectorAll(sel).forEach(el => {
        if (!el.closest('#cbt-promax-root') && !el.closest('#security-curtain')) {
          el.remove();
        }
      });
    });
  }

  function checkFloatingAppsAndViewport() {
    if (!state.exam.active || state.exam.locked || state.exam.isSubmitting || state.exam.finishedHandled) return;
    if (Date.now() < (state.exam.graceUntil || 0)) return;
    if (state.confirmDialog) return;

    if (document.pictureInPictureElement) {
      recordViolation('Mode Picture-in-Picture / Floating Video terdeteksi');
      return;
    }

    const sw = window.screen?.width || 0;
    const sh = window.screen?.height || 0;
    const iw = window.innerWidth || 0;
    const ih = window.innerHeight || 0;

    if (sw > 0 && sh > 0 && iw > 0 && ih > 0) {
      const activeTag = (document.activeElement?.tagName || '').toUpperCase();
      const isTyping = activeTag === 'INPUT' || activeTag === 'TEXTAREA' || activeTag === 'SELECT';
      if (!isTyping && (iw < sw * 0.72 || ih < sh * 0.62)) {
        recordViolation('Layar Belah (Split Screen) atau Floating App terdeteksi');
      }
    }
  }

  function lockBrowserNavigation() {
    try {
      window.history.pushState({ cbtExamLock: true }, '', window.location.href);
      state.exam.navLockActive = true;
    } catch (e) {}
  }

  function setExamScrollLock(locked) {
    const htmlEl = document.documentElement;
    const bodyEl = document.body;
    const rootEl = document.getElementById('cbt-promax-root');
    if (locked) {
      htmlEl?.classList.add('exam-active-lock');
      bodyEl?.classList.add('exam-active-lock');
      rootEl?.classList.add('exam-active-lock');
    } else {
      htmlEl?.classList.remove('exam-active-lock');
      bodyEl?.classList.remove('exam-active-lock');
      rootEl?.classList.remove('exam-active-lock');
    }
  }

  function safeRequestFullscreen() {
    try {
      const el = document.documentElement;
      if (!document.fullscreenElement && el.requestFullscreen) {
        const p = el.requestFullscreen({ navigationUI: 'hide' });
        if (p && typeof p.then === 'function') {
          p.then(() => {
            state.exam.wasFullscreen = true;
            state.exam.needsFullscreenTap = false;
            if (state.exam.active && !state.exam.locked) render();
          }).catch(() => {
            if (state.exam.active && !state.exam.locked) {
              state.exam.needsFullscreenTap = true;
            }
          });
        }
      } else if (document.fullscreenElement) {
        state.exam.wasFullscreen = true;
        state.exam.needsFullscreenTap = false;
      }
    } catch (e) {}
  }

  function safeExitFullscreen() {
    try {
      if (document.fullscreenElement && document.exitFullscreen) {
        document.exitFullscreen().catch(() => {});
      }
    } catch (e) {}
  }

  function cleanupExamEnvironment() {
    state.exam.isSubmitting = true;
    state.exam.finishedHandled = true;
    state.exam.active = false;
    state.exam.locked = false;
    state.exam.needsFullscreenTap = false;
    state.exam.wasFullscreen = false;
    state.exam.navLockActive = false;

    if (state.exam.timerInterval) {
      clearInterval(state.exam.timerInterval);
      state.exam.timerInterval = null;
    }
    if (state.exam.watchdogInterval) {
      clearInterval(state.exam.watchdogInterval);
      state.exam.watchdogInterval = null;
    }
    if (state.exam.domObserver) {
      try { state.exam.domObserver.disconnect(); } catch (e) {}
      state.exam.domObserver = null;
    }
    if (state.exam.wakeLock) {
      try { state.exam.wakeLock.release(); } catch (e) {}
      state.exam.wakeLock = null;
    }

    setExamScrollLock(false);
    setAntiTranslate(false);

    const curtain = document.getElementById('security-curtain');
    if (curtain) curtain.style.display = 'none';

    safeExitFullscreen();
  }

  function showLockedCurtain(reason, count, maxViolation) {
    const curtain = document.getElementById('security-curtain');
    if (!curtain) return;
    const isHardLocked = count >= maxViolation || state.exam.locked;

    document.getElementById('curtain-title').textContent = isHardLocked
      ? 'UJIAN DIKUNCI OTOMATIS OLEH SISTEM'
      : 'PELANGGARAN KEAMANAN TERDETEKSI';
    document.getElementById('curtain-reason').textContent = isHardLocked
      ? `${reason}. Batas maksimal pelanggaran (${maxViolation}x) telah tercapai. Hubungi Pengawas / Proktor untuk membuka kunci ujian Anda.`
      : `${reason}. Aktivitas ini dicatat ke server pengawas. Jangan mengulangi tindakan ini!`;
    document.getElementById('curtain-counter').textContent = `Pelanggaran ke-${count} dari maksimal ${maxViolation}`;

    const btnDismiss = document.getElementById('btn-dismiss-curtain');
    const btnCheck = document.getElementById('btn-check-unlock');
    if (isHardLocked) {
      btnDismiss?.classList.add('hidden');
      btnCheck?.classList.remove('hidden');
    } else {
      btnDismiss?.classList.remove('hidden');
      btnCheck?.classList.add('hidden');
    }
    curtain.style.display = 'flex';
  }

  async function recordViolation(reason) {
    if (!state.exam.active || state.exam.isSubmitting || state.exam.finishedHandled) return;
    if (Date.now() < (state.exam.graceUntil || 0)) return;
    if (Date.now() - (state.exam.lastViolationAt || 0) < 2200) return;

    state.exam.lastViolationAt = Date.now();
    playViolationSound();
    if (navigator.vibrate) {
      try { navigator.vibrate([250, 100, 250, 100, 350]); } catch (e) {}
    }

    const cfg = getSchoolConfig();
    const maxV = cfg.max_pelanggaran || 3;
    const sesi = state.exam.sesiRecord;
    if (!sesi) return;

    const nextCount = (Number(sesi.jumlah_pelanggaran) || 0) + 1;
    const nextLogs = Array.isArray(sesi.log_pelanggaran) ? [...sesi.log_pelanggaran] : [];
    nextLogs.push({
      waktu: getWitaTimeString(),
      alasan: reason
    });

    const isHardLocked = nextCount >= maxV;
    if (isHardLocked) {
      state.exam.locked = true;
    }

    const patch = {
      jumlah_pelanggaran: nextCount,
      log_pelanggaran: nextLogs,
      status: isHardLocked ? 'terkunci' : 'mengerjakan'
    };
    Object.assign(sesi, patch);
    showLockedCurtain(reason, nextCount, maxV);

    try {
      await dbUpdate('sesi_ujian_siswa', sesi.id, patch);
    } catch (e) {}
  }

  async function unlockStudentSession(sesiId) {
    const sesi = state.db.sesi_ujian_siswa.find(s => s.id === sesiId);
    if (!sesi) return;
    const patch = {
      status: 'mengerjakan',
      jumlah_pelanggaran: 0
    };
    try {
      await dbUpdate('sesi_ujian_siswa', sesiId, patch);
      if (state.exam.active && state.exam.sesiRecord?.id === sesiId) {
        Object.assign(state.exam.sesiRecord, patch);
        state.exam.locked = false;
        state.exam.graceUntil = Date.now() + 4500;
        state.exam.lastViolationAt = 0;
        const curtain = document.getElementById('security-curtain');
        if (curtain) curtain.style.display = 'none';
        safeRequestFullscreen();
      }
      showToast('Kunci sesi peserta berhasil dibuka! Pelanggaran direset ke 0.');
      render();
    } catch (err) {
      showToast('Gagal membuka kunci sesi: ' + (err.message || err), 'error');
    }
  }

  function setupAntiCheatListeners() {
    document.getElementById('btn-dismiss-curtain')?.addEventListener('click', () => {
      const curtain = document.getElementById('security-curtain');
      if (curtain) curtain.style.display = 'none';
      state.exam.graceUntil = Date.now() + 3000;
      state.exam.needsFullscreenTap = false;
      safeRequestFullscreen();
      render();
    });

    document.getElementById('btn-check-unlock')?.addEventListener('click', async () => {
      await syncFastOperationalData(false);
      if (state.exam.sesiRecord) {
        const latest = state.db.sesi_ujian_siswa.find(s => s.id === state.exam.sesiRecord.id);
        if (latest && latest.status !== 'terkunci') {
          Object.assign(state.exam.sesiRecord, latest);
          state.exam.locked = false;
          state.exam.needsFullscreenTap = false;
          state.exam.graceUntil = Date.now() + 4500;
          state.exam.lastViolationAt = 0;
          const curtain = document.getElementById('security-curtain');
          if (curtain) curtain.style.display = 'none';
          safeRequestFullscreen();
          showToast('Sesi telah dibuka oleh pengawas! Silakan lanjutkan ujian.');
          render();
          return;
        }
      }
      showToast('Sesi masih terkunci. Tunggu pengawas membuka kunci di Monitoring Live.', 'error');
    });

    document.addEventListener('pointerdown', () => {
      if (state.exam.active && !state.exam.locked && !state.exam.isSubmitting && !state.exam.finishedHandled) {
        if (!document.fullscreenElement) {
          state.exam.graceUntil = Math.max(state.exam.graceUntil || 0, Date.now() + 2500);
          safeRequestFullscreen();
        }
      }
    }, true);

    document.addEventListener('visibilitychange', () => {
      if (document.hidden && state.exam.active && !state.exam.isSubmitting && !state.exam.finishedHandled) {
        recordViolation('Berpindah tab atau meminimalkan browser saat ujian');
      }
    });

    window.addEventListener('blur', () => {
      if (state.exam.active && !state.exam.isSubmitting && !state.exam.finishedHandled && !state.confirmDialog) {
        if (Date.now() < (state.exam.graceUntil || 0)) return;
        const activeTag = (document.activeElement?.tagName || '').toUpperCase();
        if (activeTag === 'SELECT') return;
        recordViolation('Keluar dari fokus jendela ujian (Split Screen / Notifikasi / Berpindah Aplikasi)');
      }
    });

    let touchStartX = 0;
    let touchStartY = 0;
    document.addEventListener('touchstart', (e) => {
      if (!state.exam.active || state.exam.isSubmitting || state.exam.finishedHandled) return;
      if (e.touches.length >= 3) {
        e.preventDefault();
        recordViolation('Gestur 3 jari (Screenshot) terdeteksi di layar ujian');
        return;
      }
      if (e.touches.length >= 2) {
        e.preventDefault();
        return;
      }
      if (e.touches.length === 1) {
        touchStartX = e.touches[0].clientX;
        touchStartY = e.touches[0].clientY;
      }
    }, { passive: false });

    document.addEventListener('touchmove', (e) => {
      if (!state.exam.active || state.exam.isSubmitting || state.exam.finishedHandled) return;
      if (e.touches.length >= 2) {
        e.preventDefault();
        return;
      }
      if (e.touches.length !== 1) return;

      const curX = e.touches[0].clientX;
      const curY = e.touches[0].clientY;
      const dx = curX - touchStartX;
      const dy = curY - touchStartY;
      const rootEl = document.getElementById('cbt-promax-root');

      if (dy > 10 && (!rootEl || rootEl.scrollTop <= 0)) {
        e.preventDefault();
      }
      if (touchStartY <= 18 && dy > 26) {
        e.preventDefault();
        recordViolation('Menarik bilah notifikasi dari tepi atas layar');
        return;
      }
      if (touchStartY >= window.innerHeight - 18 && dy < -26) {
        e.preventDefault();
        recordViolation('Menarik bilah navigasi dari tepi bawah layar');
        return;
      }
      if (Math.abs(dx) > Math.abs(dy) * 1.3 && Math.abs(dx) > 22) {
        e.preventDefault();
        if ((touchStartX <= 16 || touchStartX >= window.innerWidth - 16) && Math.abs(dx) > 38) {
          recordViolation('Gestur usap navigasi (Back/Swipe) dari tepi layar');
        }
      }
    }, { passive: false });

    window.addEventListener('popstate', () => {
      if (state.exam.active && state.exam.navLockActive && !state.exam.isSubmitting && !state.exam.finishedHandled) {
        lockBrowserNavigation();
        recordViolation('Menekan tombol Kembali (Back) saat ujian berlangsung');
      }
    });

    window.addEventListener('resize', () => {
      if (state.exam.active && !state.exam.isSubmitting && !state.exam.finishedHandled) {
        if (Date.now() < (state.exam.graceUntil || 0)) return;
        checkFloatingAppsAndViewport();
      }
    });

    document.addEventListener('fullscreenchange', () => {
      if (document.fullscreenElement) {
        state.exam.wasFullscreen = true;
        state.exam.needsFullscreenTap = false;
        if (state.exam.active) {
          state.exam.graceUntil = Math.max(state.exam.graceUntil || 0, Date.now() + 2500);
          render();
        }
        return;
      }
      if (state.exam.wasFullscreen && state.exam.active && !state.exam.isSubmitting && !state.exam.finishedHandled) {
        if (Date.now() < (state.exam.graceUntil || 0)) return;
        recordViolation('Keluar dari mode layar penuh (Fullscreen)');
      }
    });

    document.addEventListener('keydown', (e) => {
      if (!state.exam.active || state.exam.isSubmitting || state.exam.finishedHandled) return;
      const activeTag = (document.activeElement?.tagName || '').toUpperCase();
      const isInputFocused = activeTag === 'INPUT' || activeTag === 'TEXTAREA';
      const key = (e.key || '').toLowerCase();

      if (key === 'printscreen' || key === 'escape' || key === 'meta' || (e.keyCode >= 112 && e.keyCode <= 123)) {
        e.preventDefault();
        recordViolation(`Tombol sistem terlarang (${e.key}) ditekan`);
        return;
      }
      if ((e.ctrlKey || e.metaKey || e.altKey) && ['c', 'v', 'x', 'a', 'p', 's', 'u', 'f', 'j', 'k', 'm', 'tab', ' '].includes(key)) {
        if (isInputFocused && ['a', 'c', 'v', 'x'].includes(key)) return;
        e.preventDefault();
        recordViolation(`Pintasan keyboard terlarang (${e.key}) terdeteksi`);
      }
    });

    ['contextmenu', 'copy', 'cut', 'paste', 'dragstart'].forEach(evtName => {
      document.addEventListener(evtName, (e) => {
        if (state.exam.active && !state.exam.isSubmitting && !state.exam.finishedHandled) {
          const tag = (e.target?.tagName || '').toUpperCase();
          if ((tag === 'INPUT' || tag === 'TEXTAREA') && ['copy', 'cut', 'paste'].includes(evtName)) {
            return;
          }
          e.preventDefault();
        }
      });
    });
  }

  // =========================================================================
  // EKSEKUSI UJIAN SISWA (MULAI, SIMPAN JAWABAN, SKOR OTOMATIS, SELESAI)
  // =========================================================================
  async function startExamForJadwal(jadwal, tokenInput) {
    if (!state.currentUser || state.currentUser.role !== 'siswa') return;
    if (getEffectiveJadwalStatus(jadwal) !== 'Aktif') {
      showToast('Jadwal ujian ini sudah berakhir atau tidak aktif.', 'error');
      return;
    }
    if ((tokenInput || '').trim().toUpperCase() !== (jadwal.token_ujian || '').trim().toUpperCase()) {
      showToast('Token Ujian tidak valid! Periksa kembali token dari pengawas.', 'error');
      return;
    }

    const emuError = detectVirtualAndroidOrEmulator();
    if (emuError) {
      showToast(`Akses Ujian Ditolak: ${emuError}. Gunakan perangkat asli!`, 'error');
      return;
    }

    const mapel = state.db.mata_pelajaran.find(m => m.id === jadwal.mapel_id) || {
      id: jadwal.mapel_id,
      nama_mapel: 'Mata Pelajaran CBT',
      is_english: false
    };

    let soalPool = state.db.bank_soal.filter(s => s.mapel_id === jadwal.mapel_id);
    if (Array.isArray(jadwal.soal_ids_json) && jadwal.soal_ids_json.length > 0) {
      const idSet = new Set(jadwal.soal_ids_json);
      const filtered = soalPool.filter(s => idSet.has(s.id));
      if (filtered.length > 0) soalPool = filtered;
    } else {
      const studentKelas = (state.currentUser.kelas || '').toLowerCase();
      const classFiltered = soalPool.filter(s => {
        const targets = parseKelasString(s.kelas_target).map(k => k.toLowerCase());
        return targets.length === 0 || targets.includes(studentKelas);
      });
      if (classFiltered.length > 0) soalPool = classFiltered;
    }

    soalPool.sort((a, b) => (Number(a.nomor_urut) || 0) - (Number(b.nomor_urut) || 0));
    if (jadwal.acak_soal) {
      soalPool = shuffleArray(soalPool);
    }

    if (soalPool.length === 0) {
      showToast('Belum ada butir soal tersedia untuk jadwal ujian ini.', 'error');
      return;
    }

    const optionOrders = {};
    soalPool.forEach(s => {
      const tipe = normalizeTipeSoal(s.tipe_soal);
      if (tipe === 'PILIHAN GANDA' || tipe === 'PILIHAN GANDA KOMPLEKS') {
        const availableKeys = ['A', 'B', 'C', 'D'].filter(k => (s[`opsi_${k.toLowerCase()}`] || '').trim() !== '');
        optionOrders[s.id] = jadwal.acak_opsi ? shuffleArray(availableKeys) : availableKeys;
      } else if (tipe === 'MENJODOHKAN') {
        const pairs = normalizePasanganJson(s.pasangan_json, s);
        const indices = pairs.map((_, idx) => idx);
        const rightTexts = pairs.map(p => p.kanan);
        const shouldShuffleMatch = Boolean(jadwal.acak_opsi || jadwal.acak_soal);
        optionOrders[`${s.id}_left`] = shouldShuffleMatch ? shuffleArray(indices) : indices;
        optionOrders[`${s.id}_right`] = shouldShuffleMatch ? shuffleArray(rightTexts) : rightTexts;
      }
    });

    let sesi = state.db.sesi_ujian_siswa.find(
      s => String(s.jadwal_id) === String(jadwal.id) && String(s.siswa_id) === String(state.currentUser.id)
    );
    if (sesi && sesi.status === 'selesai') {
      showToast('Anda sudah menyelesaikan ujian ini.', 'error');
      return;
    }
    if (!sesi) {
      try {
        sesi = await dbInsert('sesi_ujian_siswa', {
          jadwal_id: jadwal.id,
          siswa_id: state.currentUser.id,
          jawaban_json: {},
          ragu_json: {},
          jumlah_pelanggaran: 0,
          log_pelanggaran: [],
          skor_pg: 0,
          skor_essay: 0,
          nilai_akhir: 0,
          status: 'mengerjakan',
          waktu_mulai: new Date().toISOString()
        });
      } catch (err) {
        showToast('Gagal memulai sesi ujian di server: ' + (err.message || err), 'error');
        return;
      }
    }

    unlockAudioContext();
    const durationSec = (Number(jadwal.durasi_menit) || 90) * 60;
    const startMs = sesi.waktu_mulai ? new Date(sesi.waktu_mulai).getTime() : Date.now();
    const elapsedSec = Math.max(0, Math.floor((Date.now() - startMs) / 1000));
    const untilEndSec = jadwal.waktu_selesai
      ? Math.max(0, Math.floor((new Date(jadwal.waktu_selesai).getTime() - Date.now()) / 1000))
      : durationSec;
    const remainingSeconds = Math.max(0, Math.min(durationSec - elapsedSec, untilEndSec));

    if (remainingSeconds <= 0) {
      state.exam.sesiRecord = sesi;
      state.exam.jadwal = jadwal;
      state.exam.soalList = soalPool;
      state.exam.jawaban = sesi.jawaban_json || {};
      state.exam.ragu = sesi.ragu_json || {};
      await finishExamAndSubmit(true);
      return;
    }

    if (state.exam.timerInterval) clearInterval(state.exam.timerInterval);
    if (state.exam.watchdogInterval) clearInterval(state.exam.watchdogInterval);
    if (state.exam.domObserver) {
      try { state.exam.domObserver.disconnect(); } catch (e) {}
    }

    state.exam.active = true;
    state.exam.locked = sesi.status === 'terkunci';
    state.exam.needsFullscreenTap = false;
    state.exam.isSubmitting = false;
    state.exam.finishedHandled = false;
    state.exam.wasFullscreen = false;
    state.exam.graceUntil = Date.now() + 4500;
    state.exam.lastViolationAt = 0;
    state.exam.jadwal = jadwal;
    state.exam.mapel = mapel;
    state.exam.soalList = soalPool;
    state.exam.optionOrders = optionOrders;
    state.exam.currentIndex = 0;
    state.exam.jawaban = sesi.jawaban_json && typeof sesi.jawaban_json === 'object' ? { ...sesi.jawaban_json } : {};
    state.exam.ragu = sesi.ragu_json && typeof sesi.ragu_json === 'object' ? { ...sesi.ragu_json } : {};
    state.exam.sesiRecord = sesi;
    state.exam.remainingSeconds = remainingSeconds;
    state.exam.showDaftarSoalMobile = false;

    state.modal = null;
    setExamScrollLock(true);
    lockBrowserNavigation();
    setAntiTranslate(Boolean(mapel?.is_english));
    safeRequestFullscreen();

    if (navigator.wakeLock && navigator.wakeLock.request) {
      navigator.wakeLock.request('screen').then(wl => {
        state.exam.wakeLock = wl;
      }).catch(() => {});
    }

    if (sesi.status === 'terkunci') {
      const maxV = getSchoolConfig().max_pelanggaran || 3;
      showLockedCurtain('Sesi ujian Anda sedang dalam status terkunci', Number(sesi.jumlah_pelanggaran) || maxV, maxV);
    }

    state.exam.timerInterval = setInterval(() => {
      if (!state.exam.active || state.exam.isSubmitting || state.exam.finishedHandled) return;
      state.exam.remainingSeconds = Math.max(0, state.exam.remainingSeconds - 1);
      const timerEl = document.getElementById('exam-countdown-timer');
      if (timerEl) timerEl.textContent = formatRemainingTime(state.exam.remainingSeconds);
      if (state.exam.remainingSeconds <= 0 || isJadwalExpired(state.exam.jadwal)) {
        finishExamAndSubmit(true);
      }
    }, 1000);

    state.exam.watchdogInterval = setInterval(() => {
      scanAndNeutralizeAiExtensions();
      checkFloatingAppsAndViewport();
    }, 2500);

    if (typeof MutationObserver !== 'undefined') {
      state.exam.domObserver = new MutationObserver(() => {
        scanAndNeutralizeAiExtensions();
      });
      state.exam.domObserver.observe(document.body, { childList: true, subtree: false });
    }

    render();
  }

  let saveProgressDebounce = null;
  function saveCurrentExamProgress() {
    if (!state.exam.active || !state.exam.sesiRecord) return;
    state.exam.sesiRecord.jawaban_json = { ...state.exam.jawaban };
    state.exam.sesiRecord.ragu_json = { ...state.exam.ragu };

    if (saveProgressDebounce) clearTimeout(saveProgressDebounce);
    saveProgressDebounce = setTimeout(async () => {
      if (!state.exam.sesiRecord) return;
      try {
        await dbUpdate('sesi_ujian_siswa', state.exam.sesiRecord.id, {
          jawaban_json: state.exam.jawaban,
          ragu_json: state.exam.ragu
        });
      } catch (e) {}
    }, 500);
  }

  function calculateObjectiveScore(soalList, jawaban) {
    let totalBobot = 0;
    let earnedBobot = 0;

    soalList.forEach(soal => {
      const tipe = normalizeTipeSoal(soal.tipe_soal);
      if (tipe === 'URAIAN') return;

      const bobot = Number(soal.bobot) || 10;
      totalBobot += bobot;

      const ans = jawaban ? jawaban[soal.id] : undefined;
      if (ans === undefined || ans === null) return;

      if (tipe === 'PILIHAN GANDA') {
        const kunci = String(soal.kunci_jawaban || '').trim().toUpperCase();
        const ansStr = String(ans).trim().toUpperCase();
        if (kunci && ansStr === kunci) {
          earnedBobot += bobot;
        }
      } else if (tipe === 'BENAR SALAH') {
        let kunci = String(soal.kunci_jawaban || '').trim().toUpperCase();
        if (kunci === 'A') kunci = 'BENAR';
        if (kunci === 'B') kunci = 'SALAH';
        let ansStr = String(ans).trim().toUpperCase();
        if (ansStr === 'A') ansStr = 'BENAR';
        if (ansStr === 'B') ansStr = 'SALAH';
        if (kunci && ansStr === kunci) {
          earnedBobot += bobot;
        }
      } else if (tipe === 'PILIHAN GANDA KOMPLEKS') {
        const kunciArr = String(soal.kunci_jawaban || '')
          .split(',')
          .map(s => s.trim().toUpperCase())
          .filter(Boolean);
        const ansArr = normalizePgkAnswerArray(ans);
        if (kunciArr.length > 0 && ansArr.length > 0) {
          const allSelectedAreValid = ansArr.every(a => kunciArr.includes(a));
          if (allSelectedAreValid) {
            const matched = ansArr.filter(a => kunciArr.includes(a)).length;
            earnedBobot += (matched / kunciArr.length) * bobot;
          }
        }
      } else if (tipe === 'MENJODOHKAN') {
        const pairs = normalizePasanganJson(soal.pasangan_json, soal);
        const ansObj = normalizeMatchAnswerObj(ans);
        if (pairs.length > 0) {
          let correctPairs = 0;
          pairs.forEach((p, idx) => {
            const userVal = String(ansObj[idx] || ansObj[String(idx)] || '').trim().toLowerCase();
            const keyVal = String(p.kanan || '').trim().toLowerCase();
            if (userVal && userVal === keyVal) correctPairs++;
          });
          earnedBobot += (correctPairs / pairs.length) * bobot;
        }
      } else if (tipe === 'ISIAN SINGKAT') {
        const validKeys = String(soal.kunci_jawaban || '')
          .split('|')
          .map(s => s.trim().toLowerCase())
          .filter(Boolean);
        const ansStr = String(ans).trim().toLowerCase();
        if (ansStr && validKeys.includes(ansStr)) {
          earnedBobot += bobot;
        }
      }
    });

    if (totalBobot <= 0) return 0;
    return Math.round((earnedBobot / totalBobot) * 1000) / 10;
  }

  async function finishExamAndSubmit(isAutoTimeout = false) {
    if (state.exam.finishedHandled && !isAutoTimeout) return;
    const sesi = state.exam.sesiRecord;
    const jadwal = state.exam.jadwal;
    const soalList = [...state.exam.soalList];
    const jawaban = { ...state.exam.jawaban };
    const ragu = { ...state.exam.ragu };

    cleanupExamEnvironment();

    const skorPg = calculateObjectiveScore(soalList, jawaban);
    const skorEssay = Number(sesi?.skor_essay) || 0;
    const hasEssay = soalList.some(s => normalizeTipeSoal(s.tipe_soal) === 'URAIAN');
    const nilaiAkhir = hasEssay ? Math.round(((skorPg + skorEssay) / 2) * 10) / 10 : skorPg;

    const patch = {
      jawaban_json: jawaban,
      ragu_json: ragu,
      skor_pg: skorPg,
      skor_essay: skorEssay,
      nilai_akhir: nilaiAkhir,
      status: 'selesai',
      waktu_selesai: new Date().toISOString()
    };

    if (sesi) {
      Object.assign(sesi, patch);
      try {
        if (!IS_PREVIEW && supabase) {
          const { data: updatedRows, error: upErr } = await supabase
            .from('sesi_ujian_siswa')
            .update(patch)
            .eq('id', sesi.id)
            .select();
          if (upErr || !updatedRows || updatedRows.length === 0) {
            const insertPayload = {
              jadwal_id: sesi.jadwal_id || jadwal?.id,
              siswa_id: sesi.siswa_id || state.currentUser?.id,
              jumlah_pelanggaran: Number(sesi.jumlah_pelanggaran) || 0,
              log_pelanggaran: Array.isArray(sesi.log_pelanggaran) ? sesi.log_pelanggaran : [],
              waktu_mulai: sesi.waktu_mulai || new Date().toISOString(),
              ...patch
            };
            const { data: insertedRow, error: insErr } = await supabase
              .from('sesi_ujian_siswa')
              .insert([insertPayload])
              .select()
              .single();
            if (insErr) throw insErr;
            if (insertedRow) {
              state.db.sesi_ujian_siswa = state.db.sesi_ujian_siswa.filter(x => x.id !== sesi.id);
              state.db.sesi_ujian_siswa.push(insertedRow);
            }
          } else {
            const idx = state.db.sesi_ujian_siswa.findIndex(x => x.id === sesi.id);
            if (idx !== -1) state.db.sesi_ujian_siswa[idx] = updatedRows[0];
          }
          broadcastLocalSync('sesi_ujian_siswa');
        } else {
          await dbUpdate('sesi_ujian_siswa', sesi.id, patch);
        }
      } catch (err) {
        showToast('Peringatan sinkronisasi nilai ke server: ' + (err.message || err), 'error');
      }
    }

    state.activeMenu = 'dashboard_siswa';
    showToast(isAutoTimeout
      ? 'Waktu ujian habis! Jawaban Anda otomatis dikumpulkan.'
      : 'Ujian selesai! Jawaban Anda berhasil dikumpulkan dan layar kembali normal.');
    render();
  }

  function buildDaftarHadirHtml(kelasFilter, jadwalId) {
    const cfg = getSchoolConfig();
    const selJadwal = jadwalId ? state.db.jadwal_ujian.find(j => String(j.id) === String(jadwalId)) : null;
    const selMapel = selJadwal ? state.db.mata_pelajaran.find(m => String(m.id) === String(selJadwal.mapel_id)) : null;

    let siswaList = state.db.users_cbt.filter(u => u.role === 'siswa');
    if (selJadwal) {
      const targets = parseKelasString(selJadwal.kelas_target).map(k => k.toLowerCase());
      if (targets.length > 0 && !targets.includes('semua')) {
        siswaList = siswaList.filter(s => targets.includes((s.kelas || '').trim().toLowerCase()));
      }
      if (selMapel && selMapel.kategori_agama) {
        siswaList = siswaList.filter(s => (s.agama || '').trim().toLowerCase() === selMapel.kategori_agama.trim().toLowerCase());
      }
    }
    if (kelasFilter && kelasFilter !== 'SEMUA') {
      siswaList = siswaList.filter(s => (s.kelas || '').trim().toLowerCase() === kelasFilter.trim().toLowerCase());
    }
    siswaList.sort((a, b) => (a.nomor_peserta || '').localeCompare(b.nomor_peserta || '', undefined, { numeric: true }));

    const tanggalObj = selJadwal && selJadwal.waktu_mulai ? new Date(selJadwal.waktu_mulai) : new Date();
    const tanggalStr = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Makassar',
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    }).format(tanggalObj);

    const tglTtdStr = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Makassar',
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    }).format(new Date());

    const waktuStr = selJadwal
      ? `${new Intl.DateTimeFormat('id-ID', { timeZone: 'Asia/Makassar', hour: '2-digit', minute: '2-digit' }).format(new Date(selJadwal.waktu_mulai)).replace(/\./g, ':')} - ${new Intl.DateTimeFormat('id-ID', { timeZone: 'Asia/Makassar', hour: '2-digit', minute: '2-digit' }).format(new Date(selJadwal.waktu_selesai)).replace(/\./g, ':')} WITA`
      : '07.30 - 09.30 WITA (Sesi 1)';

    const rowsHtml = siswaList.length === 0
      ? `<tr><td colspan="6" style="border:1px solid #000; padding:14px; text-align:center;">Belum ada data peserta didik pada filter kelas/jadwal terpilih.</td></tr>`
      : siswaList.map((s, i) => {
          const no = i + 1;
          const isOdd = no % 2 === 1;
          const ttdCell = isOdd
            ? `<div style="text-align:left; padding-left:4px;">${no}. ......................</div>`
            : `<div style="text-align:right; padding-right:4px;">${no}. ......................</div>`;
          return `
            <tr class="no-break">
              <td style="border:1px solid #000; padding:5px 4px; text-align:center; vertical-align:middle;">${no}</td>
              <td style="border:1px solid #000; padding:5px 6px; text-align:center; vertical-align:middle; font-family:monospace; font-size:9pt;">${escapeHtml(s.nomor_peserta || '-')}</td>
              <td style="border:1px solid #000; padding:5px 6px; text-align:center; vertical-align:middle; font-size:9pt;">${escapeHtml(s.nip_nisn || '-')}</td>
              <td style="border:1px solid #000; padding:5px 8px; vertical-align:middle;">${escapeHtml(s.nama_lengkap)}</td>
              <td style="border:1px solid #000; padding:5px 4px; text-align:center; vertical-align:middle; font-weight:700;">${escapeHtml(s.kelas || '-')}</td>
              <td style="border:1px solid #000; padding:6px 8px; width:155px; height:26px; vertical-align:middle; font-size:9pt;">${ttdCell}</td>
            </tr>
          `;
        }).join('');

    return `
      ${renderKopSurat(cfg)}
      <div style="text-align:center; margin-bottom:12px;">
        <div style="font-size:11.5pt; font-weight:800; text-decoration:underline; letter-spacing:0.3px;">DAFTAR HADIR PESERTA ASESMEN / UJIAN BERBASIS KOMPUTER (CBT)</div>
        <div style="font-size:9.5pt; margin-top:2px;">Tahun Pelajaran ${escapeHtml(cfg.tahun_pelajaran)} | Semester ${escapeHtml(cfg.semester)}</div>
      </div>

      <table style="width:100%; font-size:9.5pt; margin-bottom:12px; border-collapse:collapse;">
        <tr>
          <td style="width:18%; padding:2px 0; vertical-align:top;">Satuan Pendidikan</td>
          <td style="width:2%; padding:2px 0; vertical-align:top;">:</td>
          <td style="width:32%; padding:2px 8px 2px 0; vertical-align:top; font-weight:700;">${escapeHtml(cfg.nama_sekolah)}</td>
          <td style="width:15%; padding:2px 0; vertical-align:top;">Hari / Tanggal</td>
          <td style="width:2%; padding:2px 0; vertical-align:top;">:</td>
          <td style="width:31%; padding:2px 0; vertical-align:top;">${escapeHtml(tanggalStr)}</td>
        </tr>
        <tr>
          <td style="padding:2px 0; vertical-align:top;">Mata Pelajaran</td>
          <td style="padding:2px 0; vertical-align:top;">:</td>
          <td style="padding:2px 8px 2px 0; vertical-align:top; font-weight:700;">${escapeHtml(selMapel ? selMapel.nama_mapel : 'Semua Mata Pelajaran')}</td>
          <td style="padding:2px 0; vertical-align:top;">Waktu Ujian</td>
          <td style="padding:2px 0; vertical-align:top;">:</td>
          <td style="padding:2px 0; vertical-align:top;">${escapeHtml(waktuStr)}</td>
        </tr>
        <tr>
          <td style="padding:2px 0; vertical-align:top;">Kelas / Rombel</td>
          <td style="padding:2px 0; vertical-align:top;">:</td>
          <td style="padding:2px 8px 2px 0; vertical-align:top; font-weight:700;">${escapeHtml(kelasFilter === 'SEMUA' ? (selJadwal ? selJadwal.kelas_target : 'Semua Kelas') : kelasFilter)}</td>
          <td style="padding:2px 0; vertical-align:top;">Ruang Ujian</td>
          <td style="padding:2px 0; vertical-align:top;">:</td>
          <td style="padding:2px 0; vertical-align:top;">Ruang 01 / Lab Komputer 1</td>
        </tr>
      </table>

      <table style="width:100%; font-size:9.5pt; margin-bottom:12px; border-collapse:collapse;">
        <thead>
          <tr style="background:#f1f5f9;">
            <th style="border:1px solid #000; padding:6px 4px; width:34px; text-align:center;">NO</th>
            <th style="border:1px solid #000; padding:6px; width:125px; text-align:center;">NOMOR PESERTA</th>
            <th style="border:1px solid #000; padding:6px; width:95px; text-align:center;">NISN</th>
            <th style="border:1px solid #000; padding:6px 8px; text-align:left;">NAMA LENGKAP PESERTA</th>
            <th style="border:1px solid #000; padding:6px 4px; width:55px; text-align:center;">KELAS</th>
            <th style="border:1px solid #000; padding:6px; width:155px; text-align:center;">TANDA TANGAN</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>

      <div class="no-break" style="border:1px solid #000; padding:7px 10px; font-size:9.5pt; margin-bottom:18px; width:62%;">
        <table style="width:100%; border-collapse:collapse;">
          <tr>
            <td style="width:55%; padding:1px 0;">Jumlah Peserta Seharusnya</td>
            <td style="width:5%;">:</td>
            <td style="font-weight:700;">${siswaList.length} Orang</td>
          </tr>
          <tr>
            <td style="padding:1px 0;">Jumlah Peserta Hadir</td>
            <td>:</td>
            <td>............. Orang</td>
          </tr>
          <tr>
            <td style="padding:1px 0;">Jumlah Peserta Tidak Hadir</td>
            <td>:</td>
            <td>............. Orang</td>
          </tr>
        </table>
      </div>

      <table class="no-break" style="width:100%; font-size:9.5pt; text-align:center; border-collapse:collapse;">
        <tr>
          <td style="width:33%; vertical-align:top; padding:0 4px;">
            &nbsp;<br/>
            Pengawas Ruang,<br/><br/><br/><br/>
            <strong><u>${escapeHtml(cfg.nama_pengawas)}</u></strong><br/>
            NIP. ${escapeHtml(cfg.nip_pengawas)}
          </td>
          <td style="width:33%; vertical-align:top; padding:0 4px;">
            &nbsp;<br/>
            Proktor CBT,<br/><br/><br/><br/>
            <strong><u>${escapeHtml(cfg.nama_proktor)}</u></strong><br/>
            NIP. ${escapeHtml(cfg.nip_proktor)}
          </td>
          <td style="width:34%; vertical-align:top; padding:0 4px;">
            Muara Kaman, ${escapeHtml(tglTtdStr)}<br/>
            Kepala Satuan Pendidikan,<br/><br/><br/><br/>
            <strong><u>${escapeHtml(cfg.nama_kepsek)}</u></strong><br/>
            NIP. ${escapeHtml(cfg.nip_kepsek)}
          </td>
        </tr>
      </table>
    `;
  }

  function buildKartuPesertaHtml(kelasFilter) {
    const cfg = getSchoolConfig();
    let siswaList = state.db.users_cbt.filter(u => u.role === 'siswa');
    if (kelasFilter && kelasFilter !== 'SEMUA') {
      siswaList = siswaList.filter(s => (s.kelas || '').toLowerCase() === kelasFilter.toLowerCase());
    }
    siswaList.sort((a, b) => (a.nomor_peserta || '').localeCompare(b.nomor_peserta || '', undefined, { numeric: true }));

    if (siswaList.length === 0) {
      return `<div style="padding:24px; text-align:center;">Belum ada data siswa untuk kelas terpilih.</div>`;
    }

    const schoolLogo = cfg.logo_sekolah_url
      ? `<img src="${cfg.logo_sekolah_url}" style="width:34px; height:34px; object-fit:contain; display:block;" />`
      : svgLogoSmp(34);

    const renderSingleCard = (s) => {
      if (!s) return `<td style="width:50%; padding:3px;"></td>`;
      return `
        <td style="width:50%; padding:3px; vertical-align:top;">
          <div class="no-break" style="border:1.5px solid #000; border-radius:6px; padding:6px 8px; font-size:7.6pt; background:#fff;">
            <table style="width:100%; border-collapse:collapse; border-bottom:1.5px solid #000; margin-bottom:4px;">
              <tr>
                <td style="width:32px; padding-bottom:3px; vertical-align:middle;">${schoolLogo}</td>
                <td style="text-align:center; padding-bottom:3px; vertical-align:middle; line-height:1.15;">
                  <div style="font-weight:800; font-size:8.2pt;">${escapeHtml(cfg.nama_sekolah)}</div>
                  <div style="font-weight:700; font-size:7.3pt;">KARTU PESERTA UJIAN / ASESMEN CBT</div>
                  <div style="font-size:6.8pt;">TP ${escapeHtml(cfg.tahun_pelajaran)} | NPSN: ${escapeHtml(cfg.npsn)}</div>
                </td>
              </tr>
            </table>

            <table style="width:100%; border-collapse:collapse;">
              <tr>
                <td style="vertical-align:top; padding-right:5px;">
                  <table style="width:100%; font-size:7.3pt; line-height:1.25; border-collapse:collapse;">
                    <tr><td style="width:68px;">No. Peserta</td><td style="width:7px;">:</td><td style="font-weight:700;">${escapeHtml(s.nomor_peserta || '-')}</td></tr>
                    <tr><td>Nama Lengkap</td><td>:</td><td style="font-weight:700;">${escapeHtml(s.nama_lengkap)}</td></tr>
                    <tr><td>NISN / NIS</td><td>:</td><td>${escapeHtml(s.nip_nisn || '-')}</td></tr>
                    <tr><td>Kelas / Agama</td><td>:</td><td><strong>${escapeHtml(s.kelas || '-')}</strong> / ${escapeHtml(s.agama || 'Islam')}</td></tr>
                    <tr><td style="vertical-align:top;">Ruang &amp; Sesi</td><td style="vertical-align:top;">:</td><td style="font-size:6.9pt;">${escapeHtml(s.ruang_sesi || 'Ruang 01 / Sesi 1')}</td></tr>
                  </table>
                </td>
                <td style="width:48px; vertical-align:top;">
                  <div style="width:46px; height:60px; border:1px dashed #334155; display:flex; align-items:center; justify-content:center; text-align:center; font-size:6.4pt; color:#475569;">
                    FOTO<br/>3 x 4
                  </div>
                </td>
              </tr>
            </table>

            <div style="margin-top:4px; padding:3px 7px; background:#f1f5f9; border:1px solid #64748b; border-radius:4px; font-family:monospace; font-size:7.8pt;">
              <table style="width:100%; border-collapse:collapse;">
                <tr>
                  <td style="text-align:left;">USER: <strong>${escapeHtml(s.username)}</strong></td>
                  <td style="text-align:right;">PASS: <strong>${escapeHtml(s.password)}</strong></td>
                </tr>
              </table>
            </div>

            <div style="font-size:6.4pt; line-height:1.18; border-top:1px solid #94a3b8; padding-top:3px; margin-top:4px;">
              <strong>TATA TERTIB PESERTA UJIAN:</strong><br/>
              1. Kartu ini wajib dibawa saat pelaksanaan ujian CBT.<br/>
              2. Jagalah kerahasiaan Username &amp; Password akun Anda.<br/>
              3. Dilarang membuka tab lain, layar belah, atau aplikasi bantuan AI.
            </div>
          </div>
        </td>
      `;
    };

    const pages = [];
    for (let i = 0; i < siswaList.length; i += 8) {
      const chunk = siswaList.slice(i, i + 8);
      const rows = [];
      for (let r = 0; r < chunk.length; r += 2) {
        rows.push(`<tr>${renderSingleCard(chunk[r])}${renderSingleCard(chunk[r + 1])}</tr>`);
      }
      pages.push(`
        <div class="${i + 8 < siswaList.length ? 'page-break' : ''}">
          <table style="width:100%; border-collapse:collapse; table-layout:fixed;">
            ${rows.join('')}
          </table>
        </div>
      `);
    }
    return pages.join('');
  }

  function getDefaultBeritaAcaraData() {
    const totalSiswa = state.db.users_cbt.filter(u => u.role === 'siswa').length;
    const tglStr = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Makassar',
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    }).format(new Date());
    const tglTtd = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Makassar',
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    }).format(new Date());

    return {
      hari_tanggal: tglStr,
      mapel_ujian: 'Semua Mata Pelajaran CBT',
      ruang_sesi: 'Ruang 01 / Lab Komputer 1 - Sesi 1',
      waktu_pelaksanaan: '07.30 - 09.30 WITA',
      jumlah_seharusnya: `${totalSiswa} Orang`,
      jumlah_hadir: `${totalSiswa} Orang`,
      jumlah_tidak_hadir: '0 Orang',
      rincian_tidak_hadir: 'Nihil (Semua peserta hadir lengkap)',
      catatan_khusus: 'Pelaksanaan asesmen CBT berjalan tertib, aman, lancar dan terkendali. Tidak ada kendala teknis server maupun pemadaman listrik.',
      tempat_tanggal_ttd: `Muara Kaman, ${tglTtd}`
    };
  }

  function buildBeritaAcaraHtml() {
    const cfg = getSchoolConfig();
    const ba = state.beritaAcaraForm || getDefaultBeritaAcaraData();
    return `
      ${renderKopSurat(cfg)}
      <div style="text-align:center; margin:12px 0 16px;">
        <div style="font-size:12pt; font-weight:800; text-decoration:underline; letter-spacing:0.3px;">BERITA ACARA PELAKSANAAN ASESMEN / UJIAN BERBASIS KOMPUTER</div>
        <div style="font-size:9.8pt; margin-top:2px;">Tahun Pelajaran ${escapeHtml(cfg.tahun_pelajaran)} | Semester ${escapeHtml(cfg.semester)}</div>
      </div>

      <p style="text-align:justify; line-height:1.55; margin-bottom:14px; font-size:10.2pt;">
        Pada hari ini <strong>${escapeHtml(ba.hari_tanggal)}</strong>, bertempat di <strong>${escapeHtml(cfg.nama_sekolah)}</strong> telah diselenggarakan Asesmen / Ujian Berbasis Komputer (CBT) untuk mata pelajaran <strong>${escapeHtml(ba.mapel_ujian)}</strong> pada pukul <strong>${escapeHtml(ba.waktu_pelaksanaan)}</strong> dengan rincian pelaksanaan sebagai berikut:
      </p>

      <table style="width:100%; font-size:10pt; margin-bottom:14px; border-collapse:collapse;">
        <tr>
          <td style="border:1px solid #000; padding:6px 10px; width:38%;">1. Satuan Pendidikan</td>
          <td style="border:1px solid #000; padding:6px 10px; width:3%; text-align:center;">:</td>
          <td style="border:1px solid #000; padding:6px 10px; font-weight:700;">${escapeHtml(cfg.nama_sekolah)} (NPSN: ${escapeHtml(cfg.npsn)})</td>
        </tr>
        <tr>
          <td style="border:1px solid #000; padding:6px 10px;">2. Mata Pelajaran / Ujian</td>
          <td style="border:1px solid #000; padding:6px 10px; text-align:center;">:</td>
          <td style="border:1px solid #000; padding:6px 10px; font-weight:700;">${escapeHtml(ba.mapel_ujian)}</td>
        </tr>
        <tr>
          <td style="border:1px solid #000; padding:6px 10px;">3. Ruang Ujian &amp; Sesi</td>
          <td style="border:1px solid #000; padding:6px 10px; text-align:center;">:</td>
          <td style="border:1px solid #000; padding:6px 10px;">${escapeHtml(ba.ruang_sesi)}</td>
        </tr>
        <tr>
          <td style="border:1px solid #000; padding:6px 10px;">4. Waktu Pelaksanaan</td>
          <td style="border:1px solid #000; padding:6px 10px; text-align:center;">:</td>
          <td style="border:1px solid #000; padding:6px 10px;">${escapeHtml(ba.waktu_pelaksanaan)}</td>
        </tr>
        <tr>
          <td style="border:1px solid #000; padding:6px 10px;">5. Jumlah Peserta Seharusnya</td>
          <td style="border:1px solid #000; padding:6px 10px; text-align:center;">:</td>
          <td style="border:1px solid #000; padding:6px 10px; font-weight:700;">${escapeHtml(ba.jumlah_seharusnya)}</td>
        </tr>
        <tr>
          <td style="border:1px solid #000; padding:6px 10px;">6. Jumlah Peserta Hadir</td>
          <td style="border:1px solid #000; padding:6px 10px; text-align:center;">:</td>
          <td style="border:1px solid #000; padding:6px 10px; font-weight:700;">${escapeHtml(ba.jumlah_hadir)}</td>
        </tr>
        <tr>
          <td style="border:1px solid #000; padding:6px 10px;">7. Jumlah Peserta Tidak Hadir</td>
          <td style="border:1px solid #000; padding:6px 10px; text-align:center;">:</td>
          <td style="border:1px solid #000; padding:6px 10px;">${escapeHtml(ba.jumlah_tidak_hadir)}</td>
        </tr>
        <tr>
          <td style="border:1px solid #000; padding:6px 10px; vertical-align:top;">8. Rincian Peserta Tidak Hadir</td>
          <td style="border:1px solid #000; padding:6px 10px; text-align:center; vertical-align:top;">:</td>
          <td style="border:1px solid #000; padding:6px 10px;">${escapeHtml(ba.rincian_tidak_hadir)}</td>
        </tr>
      </table>

      <div class="no-break" style="margin-bottom:14px;">
        <div style="font-weight:700; font-size:10pt; margin-bottom:4px;">Catatan Kejadian Khusus Selama Pelaksanaan Ujian:</div>
        <div style="border:1px solid #000; padding:8px 10px; min-height:60px; font-size:10pt; line-height:1.45;">${escapeHtml(ba.catatan_khusus)}</div>
      </div>

      <p style="margin-bottom:20px; font-size:10.2pt;">Demikian Berita Acara ini dibuat dengan sesungguhnya untuk dapat dipergunakan sebagaimana mestinya.</p>

      <table class="no-break" style="width:100%; font-size:10pt; text-align:center; border-collapse:collapse;">
        <tr>
          <td style="width:33%; vertical-align:top; padding:0 4px;">
            &nbsp;<br/>
            Pengawas Ruang,<br/><br/><br/><br/>
            <strong><u>${escapeHtml(cfg.nama_pengawas)}</u></strong><br/>
            NIP. ${escapeHtml(cfg.nip_pengawas)}
          </td>
          <td style="width:33%; vertical-align:top; padding:0 4px;">
            &nbsp;<br/>
            Proktor CBT,<br/><br/><br/><br/>
            <strong><u>${escapeHtml(cfg.nama_proktor)}</u></strong><br/>
            NIP. ${escapeHtml(cfg.nip_proktor)}
          </td>
          <td style="width:34%; vertical-align:top; padding:0 4px;">
            ${escapeHtml(ba.tempat_tanggal_ttd)}<br/>
            Kepala Satuan Pendidikan,<br/><br/><br/><br/>
            <strong><u>${escapeHtml(cfg.nama_kepsek)}</u></strong><br/>
            NIP. ${escapeHtml(cfg.nip_kepsek)}
          </td>
        </tr>
      </table>
    `;
  }

  function buildDaftarNilaiHtml(jadwal, rowsData, kelasFilter) {
    const cfg = getSchoolConfig();
    const mapel = state.db.mata_pelajaran.find(m => m.id === jadwal.mapel_id) || { nama_mapel: '-' };
    const guru = state.db.users_cbt.find(u => u.id === mapel.guru_id) || { nama_lengkap: '-', nip_nisn: '-' };

    const tglTtdStr = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Makassar',
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    }).format(new Date());

    const tbody = rowsData.length === 0
      ? `<tr><td colspan="7" style="border:1px solid #000; padding:12px; text-align:center;">Belum ada data nilai peserta.</td></tr>`
      : rowsData.map((r, idx) => `
          <tr class="no-break">
            <td style="border:1px solid #000; padding:5px 4px; text-align:center;">${idx + 1}</td>
            <td style="border:1px solid #000; padding:5px 8px;">${escapeHtml(r.siswa.nama_lengkap)}</td>
            <td style="border:1px solid #000; padding:5px 6px; text-align:center; font-size:9pt;">${escapeHtml(r.siswa.nip_nisn || '-')}</td>
            <td style="border:1px solid #000; padding:5px 4px; text-align:center; font-weight:700;">${escapeHtml(r.siswa.kelas || '-')}</td>
            <td style="border:1px solid #000; padding:5px; text-align:center;">${r.sesi.skor_pg ?? 0}</td>
            <td style="border:1px solid #000; padding:5px; text-align:center;">${r.sesi.skor_essay ?? 0}</td>
            <td style="border:1px solid #000; padding:5px; text-align:center; font-weight:700;">${r.sesi.nilai_akhir ?? 0}</td>
          </tr>
        `).join('');

    return `
      ${renderKopSurat(cfg)}
      <div style="text-align:center; margin-bottom:12px;">
        <div style="font-size:11.5pt; font-weight:800; text-decoration:underline;">DAFTAR NILAI &amp; HASIL UJIAN BERBASIS KOMPUTER (CBT)</div>
        <div style="font-size:9.5pt; margin-top:2px;">Tahun Pelajaran ${escapeHtml(cfg.tahun_pelajaran)} | Semester ${escapeHtml(cfg.semester)}</div>
      </div>

      <table style="width:100%; font-size:9.5pt; margin-bottom:12px; border-collapse:collapse;">
        <tr>
          <td style="width:20%; padding:2px 0;">Satuan Pendidikan</td>
          <td style="width:2%;">:</td>
          <td style="width:32%; font-weight:700;">${escapeHtml(cfg.nama_sekolah)}</td>
          <td style="width:16%;">Kelas / Rombel</td>
          <td style="width:2%;">:</td>
          <td style="font-weight:700;">${escapeHtml(kelasFilter || jadwal.kelas_target)}</td>
        </tr>
        <tr>
          <td style="padding:2px 0;">Mata Pelajaran</td>
          <td>:</td>
          <td style="font-weight:700;">${escapeHtml(mapel.nama_mapel)}</td>
          <td>Judul Ujian</td>
          <td>:</td>
          <td>${escapeHtml(jadwal.judul_ujian)}</td>
        </tr>
        <tr>
          <td style="padding:2px 0;">Guru Pengampu</td>
          <td>:</td>
          <td>${escapeHtml(guru.nama_lengkap)}</td>
          <td>Total Peserta</td>
          <td>:</td>
          <td style="font-weight:700;">${rowsData.length} Siswa</td>
        </tr>
      </table>

      <table style="width:100%; font-size:9.5pt; margin-bottom:18px; border-collapse:collapse;">
        <thead>
          <tr style="background:#f1f5f9;">
            <th style="border:1px solid #000; padding:6px 4px; width:34px; text-align:center;">NO</th>
            <th style="border:1px solid #000; padding:6px 8px; text-align:left;">NAMA LENGKAP SISWA</th>
            <th style="border:1px solid #000; padding:6px; width:105px; text-align:center;">NISN</th>
            <th style="border:1px solid #000; padding:6px 4px; width:55px; text-align:center;">KELAS</th>
            <th style="border:1px solid #000; padding:6px; width:72px; text-align:center;">SKOR PG</th>
            <th style="border:1px solid #000; padding:6px; width:80px; text-align:center;">SKOR ESSAY</th>
            <th style="border:1px solid #000; padding:6px; width:82px; text-align:center;">NILAI AKHIR</th>
          </tr>
        </thead>
        <tbody>${tbody}</tbody>
      </table>

      <table class="no-break" style="width:100%; font-size:9.5pt; text-align:center; border-collapse:collapse;">
        <tr>
          <td style="width:50%; vertical-align:top;">
            Mengetahui,<br/>
            Kepala Satuan Pendidikan,<br/><br/><br/><br/>
            <strong><u>${escapeHtml(cfg.nama_kepsek)}</u></strong><br/>
            NIP. ${escapeHtml(cfg.nip_kepsek)}
          </td>
          <td style="width:50%; vertical-align:top;">
            Muara Kaman, ${escapeHtml(tglTtdStr)}<br/>
            Guru Mata Pelajaran,<br/><br/><br/><br/>
            <strong><u>${escapeHtml(guru.nama_lengkap)}</u></strong><br/>
            NIP. ${escapeHtml(guru.nip_nisn || '-')}
          </td>
        </tr>
      </table>
    `;
  }

  function openA4PrintWindow(title, innerHtml) {
    const docHtml = `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <title>${escapeHtml(title)}</title>
  <style>
    @page { size: A4 portrait; margin: 12mm; }
    body { font-family: 'Times New Roman', Times, serif; color: #000; background: #fff; margin: 0; padding: 0; }
    .page-break { page-break-after: always; }
    .no-break { break-inside: avoid; page-break-inside: avoid; }
    table { border-collapse: collapse; }
  </style>
</head>
<body>${innerHtml}</body>
</html>`;

    const win = window.open('', '_blank', 'width=900,height=700');
    if (win) {
      win.document.open();
      win.document.write(docHtml);
      win.document.close();
      win.focus();
      setTimeout(() => {
        try { win.print(); } catch (e) {}
      }, 350);
    } else {
      window.print();
    }
  }

  async function downloadWordSoalTemplate() {
    if (typeof window.JSZip === 'undefined') {
      showToast('Pustaka pembuat Word (.docx) sedang dimuat, silakan coba lagi.', 'error');
      return;
    }

    const paragraphs = [
      'TEMPLATE BANK SOAL CBT PROMAX ASSESSMENT',
      'Petunjuk: Pertahankan format penomoran soal (1., 2., dst), tag TIPE:, opsi A./B./C./D., dan KUNCI:. Gambar yang disisipkan di bawah soal otomatis ikut terimpor.',
      '',
      '1. Kegiatan menyalurkan barang dan jasa dari produsen ke konsumen disebut ....',
      'TIPE: PILIHAN GANDA',
      'A. Produksi',
      'B. Konsumsi',
      'C. Distribusi',
      'D. Retribusi',
      'KUNCI: C',
      'BOBOT: 10',
      '',
      '2. Manakah yang termasuk sumber daya alam yang dapat diperbarui? (Pilih lebih dari satu)',
      'TIPE: PILIHAN GANDA KOMPLEKS',
      'A. Hutan dan hasil kayu',
      'B. Batu bara',
      'C. Air dan energi surya',
      'D. Minyak bumi',
      'KUNCI: A, C',
      'BOBOT: 15',
      '',
      '3. Matahari terbit dari arah barat.',
      'TIPE: BENAR SALAH',
      'A. BENAR',
      'B. SALAH',
      'KUNCI: SALAH',
      'BOBOT: 10',
      '',
      '4. Jodohkan organel sel berikut dengan fungsinya yang tepat!',
      'TIPE: MENJODOHKAN',
      'Mitokondria = Pusat penghasil energi sel',
      'Ribosom = Tempat sintesis protein',
      'Kloroplas = Tempat berlangsungnya fotosintesis',
      'KUNCI: SESUAI PASANGAN',
      'BOBOT: 20',
      '',
      '5. Ibu kota Provinsi Kalimantan Timur saat ini adalah ....',
      'TIPE: ISIAN SINGKAT',
      'KUNCI: Samarinda',
      'BOBOT: 15',
      '',
      '6. Jelaskan 3 upaya pelestarian lingkungan hidup di sekitar sekolah!',
      'TIPE: URAIAN',
      'KUNCI: Membuang sampah pada tempatnya, menanam pohon, dan menghemat penggunaan air serta listrik.',
      'BOBOT: 30'
    ];

    const bodyXml = paragraphs.map(line =>
      `<w:p><w:r><w:t xml:space="preserve">${escapeHtml(line)}</w:t></w:r></w:p>`
    ).join('');

    const zip = new window.JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
    zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
    zip.folder('word').file('document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${bodyXml}</w:body>
</w:document>`);

    const blob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Template_Soal_CBT_Promax.docx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    showToast('Template Word (.docx) berhasil diunduh!');
  }

  function parseWordHtmlToQuestions(html, mapelId, kelasTargetStr, startNo = 1) {
    const container = document.createElement('div');
    container.innerHTML = html;

    const blocks = [];
    container.querySelectorAll('p, li, tr').forEach(el => {
      const text = (el.textContent || '').replace(/\u00a0/g, ' ').trim();
      const imgs = [...el.querySelectorAll('img')].map(img => img.getAttribute('src')).filter(Boolean);
      if (text || imgs.length > 0) {
        blocks.push({ text, imgs });
      }
    });

    const questions = [];
    let current = null;

    const finalizeCurrent = () => {
      if (!current || !current.pertanyaan.trim()) return;
      const tipe = normalizeTipeSoal(current.tipe_soal);
      if (tipe === 'BENAR SALAH') {
        if (!current.opsi_a) current.opsi_a = 'BENAR';
        if (!current.opsi_b) current.opsi_b = 'SALAH';
      }
      if (tipe === 'MENJODOHKAN' && !current.kunci_jawaban) {
        current.kunci_jawaban = 'SESUAI PASANGAN';
      }
      questions.push({
        mapel_id: mapelId,
        kelas_target: kelasTargetStr,
        nomor_urut: startNo + questions.length,
        tipe_soal: tipe,
        pertanyaan: current.pertanyaan.trim(),
        gambar_url: current.gambar_url || '',
        opsi_a: current.opsi_a,
        opsi_b: current.opsi_b,
        opsi_c: current.opsi_c,
        opsi_d: current.opsi_d,
        pasangan_json: current.pasangan_json,
        kunci_jawaban: current.kunci_jawaban || '',
        bobot: current.bobot || 10,
        topik: 'Import Word'
      });
    };

    for (const blk of blocks) {
      const line = blk.text;
      const startMatch = line.match(/^(\d+)[\.\)]\s*(.+)$/);
      if (startMatch) {
        finalizeCurrent();
        current = {
          pertanyaan: startMatch[2],
          gambar_url: blk.imgs[0] || '',
          tipe_soal: 'PILIHAN GANDA',
          opsi_a: '',
          opsi_b: '',
          opsi_c: '',
          opsi_d: '',
          pasangan_json: [],
          kunci_jawaban: '',
          bobot: 10
        };
        continue;
      }

      if (!current) continue;
      if (!current.gambar_url && blk.imgs.length > 0) {
        current.gambar_url = blk.imgs[0];
      }
      if (!line) continue;

      if (/^TIPE\s*:/i.test(line)) {
        const rawTipe = line.replace(/^TIPE\s*:/i, '').trim();
        current.tipe_soal = normalizeTipeSoal(rawTipe);
      } else if (/^KUNCI\s*:/i.test(line)) {
        current.kunci_jawaban = line.replace(/^KUNCI\s*:/i, '').trim();
      } else if (/^BOBOT\s*:/i.test(line)) {
        current.bobot = Number(line.replace(/^BOBOT\s*:/i, '').trim()) || 10;
      } else if (/^A[\.\)]\s*/i.test(line)) {
        current.opsi_a = line.replace(/^A[\.\)]\s*/i, '').trim();
      } else if (/^B[\.\)]\s*/i.test(line)) {
        current.opsi_b = line.replace(/^B[\.\)]\s*/i, '').trim();
      } else if (/^C[\.\)]\s*/i.test(line)) {
        current.opsi_c = line.replace(/^C[\.\)]\s*/i, '').trim();
      } else if (/^D[\.\)]\s*/i.test(line)) {
        current.opsi_d = line.replace(/^D[\.\)]\s*/i, '').trim();
      } else if (normalizeTipeSoal(current.tipe_soal) === 'MENJODOHKAN' && line.includes('=')) {
        const [k, ...v] = line.split('=');
        const kiri = (k || '').trim();
        const kanan = v.join('=').trim();
        if (kiri && kanan) current.pasangan_json.push({ kiri, kanan });
      } else if (!current.opsi_a && !current.kunci_jawaban && current.pasangan_json.length === 0) {
        current.pertanyaan += `\n${line}`;
      }
    }

    finalizeCurrent();
    return questions;
  }

  function downloadExcelTemplate(type) {
    if (typeof window.XLSX === 'undefined') {
      showToast('Pustaka Excel sedang dimuat, silakan coba sesaat lagi.', 'error');
      return;
    }

    let rows = [];
    let fileName = 'Template_CBT_Promax.xlsx';
    if (type === 'siswa') {
      fileName = 'Template_Import_Siswa_CBT.xlsx';
      rows = [
        {
          No_Peserta: '26-2697-0001-9',
          NISN: '0143042697',
          Nama_Lengkap: 'Contoh Nama Siswa',
          Kelas: '7a',
          Agama: 'Islam',
          Ruang_Sesi: 'Ruang 01 / Lab Komputer 1 / Sesi 1',
          Username: '*2026',
          Password: '*1001'
        }
      ];
    } else if (type === 'guru') {
      fileName = 'Template_Import_Guru_CBT.xlsx';
      rows = [
        {
          NIP: '198705062022212003',
          Nama_Lengkap: 'Contoh Nama Guru, S.Pd',
          Username: 'guru01',
          Password: '123'
        }
      ];
    } else {
      fileName = 'Template_Import_Mapel_CBT.xlsx';
      rows = [
        {
          Kode_Mapel: 'MTK-07',
          Nama_Mapel: 'Matematika',
          Kelas_Ampu: '7a, 7b',
          Username_Guru: 'guruips',
          Khusus_Agama: '',
          Bahasa_Inggris: 'TIDAK'
        }
      ];
    }

    const ws = window.XLSX.utils.json_to_sheet(rows);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, 'Template');
    window.XLSX.writeFile(wb, fileName);
    showToast(`Template ${fileName} berhasil diunduh!`);
  }

  function exportRekapNilaiExcel(jadwal, rowsData, kelasFilter) {
    if (typeof window.XLSX === 'undefined') {
      showToast('Pustaka Excel sedang dimuat, silakan coba sesaat lagi.', 'error');
      return;
    }
    const mapel = state.db.mata_pelajaran.find(m => m.id === jadwal.mapel_id) || { nama_mapel: 'Mapel' };
    const exportRows = rowsData.map((r, idx) => ({
      No: idx + 1,
      No_Peserta: r.siswa.nomor_peserta || '-',
      NISN: r.siswa.nip_nisn || '-',
      Nama_Lengkap: r.siswa.nama_lengkap,
      Kelas: r.siswa.kelas || '-',
      Status: r.sesi.status === 'selesai' ? 'Selesai' : 'Mengerjakan',
      Skor_Objektif: r.sesi.skor_pg ?? 0,
      Skor_Essay: r.sesi.skor_essay ?? 0,
      Nilai_Akhir: r.sesi.nilai_akhir ?? 0
    }));

    const ws = window.XLSX.utils.json_to_sheet(exportRows);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, 'Rekap Nilai');
    const suffix = kelasFilter ? `_Kelas_${kelasFilter}` : '';
    window.XLSX.writeFile(wb, `Rekap_Nilai_${mapel.nama_mapel}_${jadwal.judul_ujian}${suffix}.xlsx`.replace(/\s+/g, '_'));
    showToast('Rekapitulasi nilai berhasil diekspor ke Excel!');
  }

  // =========================================================================
  // RENDER LOGIN VIEW
  // =========================================================================
  function renderLoginView() {
    const cfg = getSchoolConfig();
    const appLogo = cfg.logo_aplikasi_url
      ? `<img src="${cfg.logo_aplikasi_url}" alt="Logo Aplikasi" class="w-16 h-16 rounded-2xl object-contain bg-white p-1.5 shadow-md ring-4 ring-emerald-200/70 mx-auto" />`
      : `<div class="inline-flex p-1.5 rounded-2xl bg-gradient-to-br from-emerald-100 via-teal-100 to-sky-100 shadow-sm ring-2 ring-white mx-auto">${svgLogoPromax(52)}</div>`;

    return `
      <div class="relative min-h-[100dvh] w-full flex flex-col items-center justify-center px-3.5 py-6 overflow-hidden bg-gradient-to-br from-emerald-50 via-sky-50 to-amber-50">
        <!-- Ornamen Cahaya Lembut Berwarna-warni -->
        <div class="pointer-events-none absolute -top-24 -left-24 w-72 h-72 rounded-full bg-gradient-to-br from-emerald-300/35 to-teal-300/30 blur-3xl"></div>
        <div class="pointer-events-none absolute top-1/4 -right-24 w-80 h-80 rounded-full bg-gradient-to-br from-sky-300/35 to-indigo-300/30 blur-3xl"></div>
        <div class="pointer-events-none absolute -bottom-24 left-1/4 w-72 h-72 rounded-full bg-gradient-to-tr from-amber-200/40 via-rose-200/25 to-emerald-200/30 blur-3xl"></div>

        <div class="relative z-10 w-full max-w-[430px] rounded-3xl bg-white/90 backdrop-blur-xl border border-white/90 shadow-[0_20px_50px_-12px_rgba(15,23,42,0.12)] overflow-hidden">
          <!-- Pita Warna Pelangi Edukatif di Bagian Atas Kartu -->
          <div class="h-2 w-full bg-gradient-to-r from-emerald-500 via-teal-500 via-sky-500 to-amber-400"></div>

          <div class="p-6 sm:p-8 space-y-5">
            <div class="text-center space-y-2.5">
              ${appLogo}
              <div>
                <span class="inline-flex items-center gap-1.5 px-3.5 py-1 rounded-full bg-gradient-to-r from-emerald-500/15 via-teal-500/15 to-sky-500/15 border border-emerald-300/70 text-[11px] font-extrabold tracking-wide text-emerald-900 shadow-2xs">
                  <span class="w-2 h-2 rounded-full bg-gradient-to-r from-emerald-500 to-sky-500 animate-pulse"></span>
                  CBT PROMAX ASSESSMENT
                </span>
              </div>
              <h1 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-slate-900 via-emerald-900 to-sky-900 bg-clip-text text-transparent tracking-tight leading-snug">
                ${escapeHtml(cfg.nama_sekolah)}
              </h1>
              <div class="inline-flex flex-wrap items-center justify-center gap-1.5 text-[11px] font-bold">
                <span class="px-2.5 py-0.5 rounded-lg bg-sky-50 text-sky-800 border border-sky-200/80">
                  TP ${escapeHtml(cfg.tahun_pelajaran)}
                </span>
                <span class="px-2.5 py-0.5 rounded-lg bg-amber-50 text-amber-800 border border-amber-200/80">
                  Semester ${escapeHtml(cfg.semester)}
                </span>
              </div>
            </div>

            <form id="login-form" class="space-y-4">
              <div>
                <label for="login-username" class="flex items-center gap-1.5 text-xs font-extrabold text-emerald-950 mb-1.5">
                  <span class="w-5 h-5 rounded-lg bg-emerald-100 text-emerald-700 inline-flex items-center justify-center text-[11px]">👤</span>
                  <span>Username / ID Peserta</span>
                </label>
                <input id="login-username" type="text" required autocomplete="username" placeholder="Masukkan username atau ID peserta"
                  class="w-full px-4 py-3 rounded-2xl border-2 border-emerald-200/80 bg-emerald-50/25 text-base sm:text-sm font-semibold text-slate-900 placeholder:text-slate-400 focus:bg-white focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 focus:outline-none transition" />
              </div>

              <div>
                <label for="login-password" class="flex items-center gap-1.5 text-xs font-extrabold text-sky-950 mb-1.5">
                  <span class="w-5 h-5 rounded-lg bg-sky-100 text-sky-700 inline-flex items-center justify-center text-[11px]">🔑</span>
                  <span>Password / Kata Sandi</span>
                </label>
                <div class="relative flex items-center">
                  <input id="login-password" type="password" required autocomplete="current-password" placeholder="Masukkan password"
                    class="w-full pl-4 pr-12 py-3 rounded-2xl border-2 border-sky-200/80 bg-sky-50/25 text-base sm:text-sm font-semibold text-slate-900 placeholder:text-slate-400 focus:bg-white focus:border-sky-500 focus:ring-4 focus:ring-sky-500/10 focus:outline-none transition" />
                  <button id="toggle-pwd-btn" type="button" aria-label="Tampilkan Password"
                    class="absolute right-1.5 w-10 h-10 rounded-xl flex items-center justify-center text-sky-700 hover:bg-sky-100/80 text-sm transition">
                    👁
                  </button>
                </div>
              </div>

              <button type="submit" ${state.loading ? 'disabled' : ''}
                class="w-full py-3.5 px-4 rounded-2xl bg-gradient-to-r from-emerald-600 via-teal-600 to-sky-600 hover:from-emerald-700 hover:via-teal-700 hover:to-sky-700 text-white font-extrabold text-sm shadow-lg shadow-emerald-600/25 transition-all active:scale-[0.99] disabled:opacity-60">
                ${state.loading ? 'Memverifikasi Akun...' : 'Masuk ke Ruang Ujian →'}
              </button>
            </form>

            <div class="pt-3.5 border-t border-slate-200/80 text-center space-y-1.5">
              <div class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-gradient-to-r from-amber-50 via-orange-50 to-rose-50 border border-amber-200/80">
                <span class="text-xs">🕒</span>
                <span id="login-clock" class="text-[11px] font-extrabold text-amber-900">${getWitaTimeString()}</span>
              </div>
              <div class="text-[10.5px] font-semibold bg-gradient-to-r from-slate-500 via-teal-700 to-indigo-600 bg-clip-text text-transparent">
                Sistem Ujian Berbasis Komputer &amp; Anti-Kecurangan
              </div>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  // =========================================================================
  // RENDER HEADER & SIDEBAR OTOMATIS KIRI ATAS
  // =========================================================================
  function renderHeaderAndSidebar(contentHtml) {
    const cfg = getSchoolConfig();
    const u = state.currentUser;
    const isAdmin = u.role === 'admin';
    const isSiswa = u.role === 'siswa';

    const menuItems = isSiswa ? [
      { id: 'dashboard_siswa', icon: '🗓', label: 'Jadwal & Mulai Ujian', accent: 'from-emerald-600 to-teal-600' },
      { id: 'nilai_siswa', icon: '📊', label: 'Riwayat & Nilai Ujian', accent: 'from-sky-600 to-indigo-600' }
    ] : [
      { id: 'dashboard', icon: '🏠', label: 'Dashboard Utama', accent: 'from-emerald-600 to-teal-600' },
      { id: 'bank_soal', icon: '📚', label: 'Bank Butir Soal', accent: 'from-indigo-600 to-blue-600' },
      { id: 'jadwal', icon: '🗓', label: 'Jadwal & Token Ujian', accent: 'from-amber-600 to-orange-500' },
      { id: 'monitoring', icon: '📡', label: 'Monitoring Live', accent: 'from-cyan-600 to-sky-600' },
      { id: 'rekap_nilai', icon: '📊', label: 'Rekapitulasi Nilai', accent: 'from-violet-600 to-purple-600' },
      { id: 'cetak_dokumen', icon: '🖨', label: 'Cetak Dokumen A4', accent: 'from-teal-600 to-emerald-600' },
      { id: 'mapel', icon: '📘', label: 'Mata Pelajaran', accent: 'from-blue-600 to-indigo-600' },
      ...(isAdmin ? [
        { id: 'users', icon: '👥', label: 'Data Guru & Siswa', accent: 'from-rose-600 to-pink-600' },
        { id: 'pengaturan', icon: '⚙️', label: 'Pengaturan Sekolah', accent: 'from-slate-700 to-slate-800' }
      ] : [])
    ];

    const appLogo = cfg.logo_aplikasi_url
      ? `<img src="${cfg.logo_aplikasi_url}" alt="Logo" class="w-9 h-9 rounded-xl object-contain bg-white p-0.5 ring-2 ring-emerald-200/70" />`
      : svgLogoPromax(36);

    return `
      ${state.sidebarOpen ? `<div id="sidebar-backdrop" class="fixed inset-0 z-30 bg-slate-900/40 backdrop-blur-xs lg:hidden"></div>` : ''}

      <aside class="app-sidebar glass-sidebar fixed top-0 left-0 bottom-0 z-40 w-[264px] flex-col justify-between p-4 ${state.sidebarOpen ? 'is-open' : ''}">
        <div class="space-y-4 overflow-y-auto">
          <div class="flex items-center gap-3 pb-3 border-b border-emerald-200/70">
            ${appLogo}
            <div class="min-w-0">
              <div class="text-xs font-extrabold bg-gradient-to-r from-emerald-800 via-teal-700 to-sky-800 bg-clip-text text-transparent truncate">CBT PROMAX</div>
              <div class="text-[10px] font-semibold text-slate-500 truncate">${escapeHtml(cfg.nama_sekolah)}</div>
            </div>
          </div>

          <div class="p-3 rounded-2xl bg-gradient-to-br from-white via-emerald-50/70 to-sky-50/70 border border-emerald-200/80 shadow-2xs space-y-0.5">
            <div class="text-xs font-extrabold text-slate-900 truncate">${escapeHtml(u.nama_lengkap)}</div>
            <div class="inline-flex items-center gap-1.5 text-[11px] font-bold text-emerald-700 capitalize">
              <span class="w-2 h-2 rounded-full bg-emerald-500"></span>
              <span>${u.role === 'siswa' ? `Peserta - Kelas ${escapeHtml(u.kelas)}` : (u.role === 'guru' ? 'Guru Mata Pelajaran' : 'Administrator / Proktor')}</span>
            </div>
          </div>

          <nav class="space-y-1">
            ${menuItems.map(item => `
              <button type="button" data-menu="${item.id}"
                class="nav-menu-btn w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs font-bold transition-all text-left ${
                  state.activeMenu === item.id
                    ? `bg-gradient-to-r ${item.accent} text-white shadow-sm`
                    : 'text-slate-700 hover:bg-gradient-to-r hover:from-emerald-50/80 hover:to-sky-50/80 hover:text-slate-950'
                }">
                <span class="text-base">${item.icon}</span>
                <span>${item.label}</span>
              </button>
            `).join('')}
          </nav>
        </div>

        <div class="pt-3 border-t border-slate-200/80">
          <button id="btn-logout" type="button"
            class="w-full py-2.5 px-4 rounded-xl bg-gradient-to-r from-rose-50 to-red-50 hover:from-rose-100 hover:to-red-100 border border-rose-200/80 text-rose-700 font-extrabold text-xs transition flex items-center justify-center gap-2">
            <span>🚪</span>
            <span>Keluar Akun</span>
          </button>
        </div>
      </aside>

      <div class="main-with-sidebar ${state.sidebarOpen ? 'sidebar-docked' : ''}">
        <header class="sticky top-0 z-20 bg-gradient-to-r from-white/95 via-emerald-50/60 to-sky-50/70 backdrop-blur-md border-b border-emerald-200/70 px-3.5 sm:px-6 py-3 flex items-center justify-between gap-2 shadow-2xs">
          <div class="flex items-center gap-2.5 min-w-0">
            <button id="btn-toggle-sidebar" type="button" aria-label="Buka/Tutup Menu Navigasi"
              class="w-10 h-10 rounded-xl bg-white hover:bg-emerald-50 border border-emerald-200/80 text-emerald-900 font-extrabold text-base flex items-center justify-center shrink-0 transition shadow-2xs">
              ${state.sidebarOpen ? '✕' : '☰'}
            </button>
            <div class="min-w-0">
              <div class="text-xs sm:text-sm font-extrabold bg-gradient-to-r from-slate-900 via-emerald-900 to-teal-800 bg-clip-text text-transparent truncate">${escapeHtml(cfg.nama_sekolah)}</div>
              <div id="header-clock" class="text-[10px] sm:text-[11px] font-bold text-teal-700/90 truncate">${getWitaTimeString()}</div>
            </div>
          </div>
          <span class="px-3 py-1 rounded-full bg-gradient-to-r from-emerald-100/90 to-teal-100/90 border border-emerald-300/70 text-emerald-900 text-[10px] sm:text-xs font-extrabold shrink-0 shadow-2xs">
            <span class="hidden sm:inline">TP </span>${escapeHtml(cfg.tahun_pelajaran)}
          </span>
        </header>

        <main class="max-w-6xl mx-auto px-3.5 sm:px-6 pt-4 sm:pt-6">
          ${contentHtml}
        </main>
      </div>
    `;
  }

  // =========================================================================
  // DASHBOARD ADMIN & GURU
  // =========================================================================
  function renderDashboardView() {
    const authMapels = getAuthorizedMapels();
    const authMapelIds = getAuthorizedMapelIds();
    const totalSoal = state.db.bank_soal.filter(s => authMapelIds.has(s.mapel_id)).length;
    const totalSiswa = state.db.users_cbt.filter(u => u.role === 'siswa').length;
    const activeJadwals = state.db.jadwal_ujian.filter(j => authMapelIds.has(j.mapel_id) && getEffectiveJadwalStatus(j) === 'Aktif');
    const activeSesi = state.db.sesi_ujian_siswa.filter(s => {
      const jd = state.db.jadwal_ujian.find(j => j.id === s.jadwal_id);
      return jd && authMapelIds.has(jd.mapel_id);
    });
    const lockedCount = activeSesi.filter(s => s.status === 'terkunci').length;

    return `
      <div class="space-y-4">
        <div class="rounded-2xl p-5 bg-gradient-to-r from-white via-emerald-50/75 to-sky-50/75 border border-emerald-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-slate-900 via-emerald-900 to-sky-900 bg-clip-text text-transparent">Dashboard Pelaksanaan CBT</h2>
            <p class="text-xs text-slate-600 mt-0.5">Ringkasan bank soal, jadwal aktif, dan status pengawasan peserta secara langsung.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button type="button" data-menu="jadwal" class="nav-menu-btn px-3.5 py-2 rounded-xl bg-gradient-to-r from-[#4f6d4a] to-emerald-700 hover:opacity-95 text-white font-bold text-xs shadow-xs transition">+ Buat Jadwal Ujian</button>
            <button type="button" data-menu="monitoring" class="nav-menu-btn px-3.5 py-2 rounded-xl bg-gradient-to-r from-sky-50 to-cyan-50 hover:from-sky-100 hover:to-cyan-100 text-sky-900 border border-sky-200 font-bold text-xs transition">📡 Pantau Live</button>
          </div>
        </div>

        <div class="grid grid-cols-2 lg:grid-cols-4 gap-3">
          <div class="rounded-2xl p-4 bg-gradient-to-br from-white to-indigo-50/80 border border-indigo-200/80 shadow-2xs">
            <div class="flex items-center justify-between">
              <span class="text-[11px] font-extrabold text-indigo-800">Mata Pelajaran</span>
              <span class="w-7 h-7 rounded-xl bg-indigo-100 text-indigo-700 flex items-center justify-center text-xs">📘</span>
            </div>
            <div class="text-2xl font-extrabold text-indigo-950 mt-1.5">${authMapels.length}</div>
          </div>
          <div class="rounded-2xl p-4 bg-gradient-to-br from-white to-emerald-50/80 border border-emerald-200/80 shadow-2xs">
            <div class="flex items-center justify-between">
              <span class="text-[11px] font-extrabold text-emerald-800">Total Butir Soal</span>
              <span class="w-7 h-7 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center text-xs">📚</span>
            </div>
            <div class="text-2xl font-extrabold text-emerald-900 mt-1.5">${totalSoal}</div>
          </div>
          <div class="rounded-2xl p-4 bg-gradient-to-br from-white to-sky-50/80 border border-sky-200/80 shadow-2xs">
            <div class="flex items-center justify-between">
              <span class="text-[11px] font-extrabold text-sky-800">Jadwal Ujian Aktif</span>
              <span class="w-7 h-7 rounded-xl bg-sky-100 text-sky-700 flex items-center justify-center text-xs">🗓</span>
            </div>
            <div class="text-2xl font-extrabold text-sky-900 mt-1.5">${activeJadwals.length}</div>
          </div>
          <div class="rounded-2xl p-4 bg-gradient-to-br from-white to-amber-50/80 border border-amber-200/80 shadow-2xs">
            <div class="flex items-center justify-between">
              <span class="text-[11px] font-extrabold text-amber-800">Peserta Terdaftar</span>
              <span class="w-7 h-7 rounded-xl bg-amber-100 text-amber-700 flex items-center justify-center text-xs">🎓</span>
            </div>
            <div class="text-2xl font-extrabold text-amber-950 mt-1.5">${totalSiswa}</div>
            ${lockedCount > 0 ? `<div class="text-[10px] font-extrabold text-rose-600 mt-1">🔒 ${lockedCount} Peserta Terkunci</div>` : ''}
          </div>
        </div>

        <div class="glass-panel rounded-2xl p-4 sm:p-5 space-y-3 border-emerald-200/70">
          <div class="font-extrabold text-sm text-emerald-950 flex items-center gap-2">
            <span class="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
            <span>🗓 Jadwal Ujian Aktif &amp; Token Hari Ini</span>
          </div>
          ${activeJadwals.length === 0 ? `
            <div class="text-xs text-slate-500 py-4 text-center">Belum ada jadwal ujian yang sedang aktif.</div>
          ` : `
            <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
              ${activeJadwals.map((j, idx) => {
                const mp = state.db.mata_pelajaran.find(m => m.id === j.mapel_id) || { nama_mapel: '-' };
                const cardTint = idx % 2 === 0
                  ? 'from-emerald-50/60 via-white to-teal-50/60 border-emerald-200/80'
                  : 'from-sky-50/60 via-white to-indigo-50/60 border-sky-200/80';
                return `
                  <div class="p-4 rounded-2xl bg-gradient-to-br ${cardTint} border flex items-center justify-between gap-3 shadow-2xs">
                    <div class="min-w-0">
                      <div class="text-xs font-extrabold text-slate-900 truncate">${escapeHtml(mp.nama_mapel)} - <span class="text-emerald-800">${escapeHtml(j.judul_ujian)}</span></div>
                      <div class="text-[11px] text-slate-600 mt-0.5">Kelas: <strong class="text-indigo-800">${escapeHtml(j.kelas_target)}</strong> | Durasi: <strong class="text-sky-800">${j.durasi_menit} Menit</strong></div>
                      <div class="text-[10px] text-slate-500 mt-0.5">Selesai: ${formatDateTimeWita(j.waktu_selesai)}</div>
                    </div>
                    <div class="text-right shrink-0">
                      <div class="text-[10px] font-extrabold text-amber-800">TOKEN</div>
                      <div class="px-3 py-1 rounded-xl bg-gradient-to-r from-amber-100 to-yellow-100 border border-amber-300 font-mono text-sm font-extrabold text-amber-950 shadow-2xs">${escapeHtml(j.token_ujian)}</div>
                    </div>
                  </div>
                `;
              }).join('')}
            </div>
          `}
        </div>
      </div>
    `;
  }

  // =========================================================================
  // BANK BUTIR SOAL
  // =========================================================================
  function renderBankSoalView() {
    const authMapels = getAuthorizedMapels();
    const activeMapel = authMapels.find(m => m.id === state.filterSoalMapel) || null;
    const kelasList = activeMapel ? getAuthorizedKelasList(activeMapel.id) : [];

    let soalList = activeMapel
      ? state.db.bank_soal.filter(s => s.mapel_id === activeMapel.id)
      : [];

    if (state.filterSoalKelas) {
      soalList = soalList.filter(s => {
        const targets = parseKelasString(s.kelas_target).map(k => k.toLowerCase());
        return targets.length === 0 || targets.includes(state.filterSoalKelas.toLowerCase());
      });
    }
    soalList.sort((a, b) => (Number(a.nomor_urut) || 0) - (Number(b.nomor_urut) || 0));

    const mapelSoftThemes = [
      { bg: 'from-white to-emerald-50/75 border-emerald-200/85 hover:border-emerald-400', badge: 'bg-emerald-100 text-emerald-900', code: 'bg-emerald-800 text-white' },
      { bg: 'from-white to-sky-50/75 border-sky-200/85 hover:border-sky-400', badge: 'bg-sky-100 text-sky-900', code: 'bg-sky-800 text-white' },
      { bg: 'from-white to-indigo-50/75 border-indigo-200/85 hover:border-indigo-400', badge: 'bg-indigo-100 text-indigo-900', code: 'bg-indigo-800 text-white' },
      { bg: 'from-white to-amber-50/75 border-amber-200/85 hover:border-amber-400', badge: 'bg-amber-100 text-amber-950', code: 'bg-amber-800 text-white' },
      { bg: 'from-white to-teal-50/75 border-teal-200/85 hover:border-teal-400', badge: 'bg-teal-100 text-teal-900', code: 'bg-teal-800 text-white' },
      { bg: 'from-white to-purple-50/75 border-purple-200/85 hover:border-purple-400', badge: 'bg-purple-100 text-purple-900', code: 'bg-purple-800 text-white' }
    ];

    return `
      <div class="space-y-4">
        <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-r from-white via-indigo-50/60 to-emerald-50/60 border border-indigo-200/75 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-indigo-950 via-slate-900 to-emerald-900 bg-clip-text text-transparent">Bank Butir Soal</h2>
            <p class="text-xs text-slate-600">Ketuk kartu mata pelajaran sekali untuk membuka daftar soal, ketuk sekali lagi untuk menutupnya.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button id="btn-download-word-tpl" type="button" class="px-3 py-2 rounded-xl bg-white hover:bg-slate-100 border border-slate-200 text-slate-800 font-bold text-xs transition shadow-2xs">
              ⬇ Template Word (.docx)
            </button>
            <button id="btn-open-import-word" type="button" class="px-3 py-2 rounded-xl bg-gradient-to-r from-sky-50 to-indigo-50 hover:from-sky-100 hover:to-indigo-100 text-indigo-900 border border-indigo-200 font-bold text-xs transition">
              📥 Import Soal (.docx)
            </button>
            <button id="btn-open-add-soal" type="button" class="px-3.5 py-2 rounded-xl bg-gradient-to-r from-[#4f6d4a] to-emerald-700 hover:opacity-95 text-white font-extrabold text-xs shadow-xs transition">
              + Tambah Soal
            </button>
          </div>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
          ${authMapels.map((m, idx) => {
            const count = state.db.bank_soal.filter(s => s.mapel_id === m.id).length;
            const isSelected = state.filterSoalMapel === m.id;
            const theme = mapelSoftThemes[idx % mapelSoftThemes.length];
            return `
              <button type="button" data-select-soal-mapel="${m.id}"
                class="p-4 rounded-2xl border-2 text-left transition-all bg-gradient-to-br ${
                  isSelected
                    ? 'from-emerald-100/85 via-emerald-50 to-teal-50 border-[#4f6d4a] shadow-sm'
                    : `${theme.bg} shadow-2xs`
                }">
                <div class="flex items-center justify-between gap-2">
                  <span class="px-2.5 py-0.5 rounded-lg font-mono text-[11px] font-extrabold ${isSelected ? 'bg-[#4f6d4a] text-white' : theme.code}">${escapeHtml(m.kode_mapel)}</span>
                  <span class="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold ${isSelected ? 'bg-emerald-200 text-emerald-950' : theme.badge}">${count} Soal</span>
                </div>
                <div class="font-extrabold text-sm text-slate-900 mt-2">${escapeHtml(m.nama_mapel)}</div>
                <div class="flex items-center justify-between text-[11px] text-slate-600 mt-1">
                  <span>Kelas Ampu: <strong class="text-slate-800">${escapeHtml(m.kelas_ampu || 'Semua Kelas')}</strong></span>
                  <span class="font-extrabold ${isSelected ? 'text-emerald-800' : 'text-indigo-700'}">${isSelected ? '▾ Terbuka' : '▸ Buka'}</span>
                </div>
              </button>
            `;
          }).join('')}
        </div>

        ${!activeMapel ? `
          <div class="rounded-2xl p-8 text-center space-y-1.5 bg-gradient-to-br from-white via-slate-50 to-indigo-50/40 border border-slate-200/80">
            <div class="text-2xl">📂</div>
            <div class="font-extrabold text-sm text-slate-800">Ketuk Salah Satu Mata Pelajaran di Atas</div>
            <p class="text-xs text-slate-500">Ketuk sekali untuk membuka daftar butir soal, dan ketuk sekali lagi pada kartu yang sama untuk menutupnya.</p>
          </div>
        ` : `
          <div class="glass-panel rounded-2xl p-4 sm:p-5 space-y-4 border-emerald-200/70">
            <div class="flex flex-wrap items-center justify-between gap-2 pb-3 border-b border-slate-200">
              <div class="flex flex-wrap items-center gap-2">
                <label class="inline-flex items-center gap-1.5 text-xs font-bold text-slate-700 cursor-pointer">
                  <input id="check-all-soal" type="checkbox" ${soalList.length > 0 && soalList.every(s => state.selectedIds.soal.has(s.id)) ? 'checked' : ''} class="rounded border-slate-300" />
                  <span>Pilih Semua (${soalList.length})</span>
                </label>
                <select id="filter-soal-kelas" class="px-3 py-1.5 rounded-xl border border-indigo-200 bg-indigo-50/40 text-xs font-bold text-indigo-950">
                  <option value="">Semua Kelas</option>
                  ${kelasList.map(k => `<option value="${k}" ${state.filterSoalKelas === k ? 'selected' : ''}>Kelas ${k}</option>`).join('')}
                </select>
              </div>
              ${state.selectedIds.soal.size > 0 ? `
                <button id="btn-bulk-delete-soal" type="button" class="px-3 py-1.5 rounded-xl bg-rose-600 text-white font-bold text-xs">
                  🗑 Hapus Terpilih (${state.selectedIds.soal.size})
                </button>
              ` : ''}
            </div>

            ${soalList.length === 0 ? `
              <div class="py-8 text-center text-xs text-slate-500">Belum ada butir soal untuk filter ini. Klik <strong>+ Tambah Soal</strong> atau <strong>Import Soal (.docx)</strong>.</div>
            ` : `
              <div class="space-y-3">
                ${soalList.map(s => {
                  const sTipe = normalizeTipeSoal(s.tipe_soal);
                  const sPairs = normalizePasanganJson(s.pasangan_json, s);
                  const tipeChip = {
                    'PILIHAN GANDA': 'bg-emerald-50 text-emerald-800 border-emerald-200',
                    'PILIHAN GANDA KOMPLEKS': 'bg-amber-50 text-amber-900 border-amber-200',
                    'BENAR SALAH': 'bg-indigo-50 text-indigo-800 border-indigo-200',
                    'MENJODOHKAN': 'bg-sky-50 text-sky-900 border-sky-200',
                    'ISIAN SINGKAT': 'bg-teal-50 text-teal-800 border-teal-200',
                    'URAIAN': 'bg-purple-50 text-purple-800 border-purple-200'
                  }[sTipe] || 'bg-sky-50 text-sky-800 border-sky-200';
                  return `
                    <div class="p-4 rounded-2xl bg-gradient-to-br from-white to-slate-50/70 border border-slate-200/90 space-y-2.5 shadow-2xs">
                      <div class="flex items-start justify-between gap-2">
                        <div class="flex flex-wrap items-center gap-2">
                          <input type="checkbox" data-id="${s.id}" class="row-check-soal rounded border-slate-300" ${state.selectedIds.soal.has(s.id) ? 'checked' : ''} />
                          <span class="px-2.5 py-0.5 rounded-lg bg-gradient-to-r from-slate-800 to-slate-900 text-white text-xs font-extrabold">No. ${s.nomor_urut}</span>
                          <span class="px-2.5 py-0.5 rounded-lg border text-[11px] font-extrabold ${tipeChip}">${escapeHtml(sTipe)}</span>
                          <span class="px-2.5 py-0.5 rounded-lg bg-indigo-50/70 text-indigo-800 text-[11px] font-bold">Kelas: ${escapeHtml(s.kelas_target || 'Semua')}</span>
                          <span class="text-[11px] font-bold text-amber-800">Bobot: ${s.bobot}</span>
                        </div>
                        <div class="flex items-center gap-1.5 shrink-0">
                          <button type="button" data-edit-soal="${s.id}" class="px-2.5 py-1 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 text-xs font-bold">✎ Edit</button>
                          <button type="button" data-del-soal="${s.id}" class="px-2.5 py-1 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-bold">🗑</button>
                        </div>
                      </div>

                      <div class="text-xs sm:text-sm text-slate-800 whitespace-pre-line leading-relaxed">${escapeHtml(s.pertanyaan)}</div>
                      ${s.gambar_url ? `<img src="${s.gambar_url}" alt="Gambar Soal" class="max-h-48 rounded-xl border border-slate-200 object-contain" />` : ''}

                      ${(sTipe === 'PILIHAN GANDA' || sTipe === 'PILIHAN GANDA KOMPLEKS') ? `
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-1.5 text-xs pt-1">
                          ${['A', 'B', 'C', 'D'].filter(k => s[`opsi_${k.toLowerCase()}`]).map(k => {
                            const kunciList = String(s.kunci_jawaban || '').split(',').map(x => x.trim().toUpperCase());
                            const isKey = kunciList.includes(k);
                            return `
                              <div class="px-3 py-1.5 rounded-lg border ${isKey ? 'bg-emerald-50 border-emerald-300 text-emerald-900 font-bold' : 'bg-slate-50 border-slate-200 text-slate-700'}">
                                <strong>${k}.</strong> ${escapeHtml(s[`opsi_${k.toLowerCase()}`])}
                              </div>
                            `;
                          }).join('')}
                        </div>
                      ` : ''}

                      ${sTipe === 'BENAR SALAH' ? `
                        <div class="flex flex-wrap gap-2 text-xs pt-1">
                          ${['BENAR', 'SALAH'].map(opt => {
                            let kNorm = String(s.kunci_jawaban || '').trim().toUpperCase();
                            if (kNorm === 'A') kNorm = 'BENAR';
                            if (kNorm === 'B') kNorm = 'SALAH';
                            const isKey = kNorm === opt;
                            return `
                              <span class="px-3 py-1 rounded-lg border ${isKey ? 'bg-emerald-50 border-emerald-300 text-emerald-900 font-extrabold' : 'bg-slate-50 border-slate-200 text-slate-600'}">
                                ${isKey ? '✓ ' : ''}${opt}
                              </span>
                            `;
                          }).join('')}
                        </div>
                      ` : ''}

                      ${sTipe === 'MENJODOHKAN' ? `
                        <div class="space-y-1 text-xs bg-sky-50/50 p-2.5 rounded-xl border border-sky-200/80">
                          ${sPairs.map((p, idx) => `<div>${idx + 1}. <strong>${escapeHtml(p.kiri)}</strong> ➔ <span class="text-emerald-800 font-semibold">${escapeHtml(p.kanan)}</span></div>`).join('')}
                        </div>
                      ` : ''}

                      <div class="text-[11px] font-bold text-emerald-700 pt-1">
                        ${sTipe === 'URAIAN' ? 'Pedoman / Rubrik' : 'Kunci Jawaban'}: <span class="underline">${escapeHtml(s.kunci_jawaban || '-')}</span>
                      </div>
                    </div>
                  `;
                }).join('')}
              </div>
            `}
          </div>
        `}
      </div>
    `;
  }

  // =========================================================================
  // MONITORING PESERTA UJIAN (LIVE)
  // =========================================================================
  function renderMonitoringView() {
    const authMapelIds = getAuthorizedMapelIds();
    const scheduledJadwals = state.db.jadwal_ujian.filter(j => authMapelIds.has(j.mapel_id));
    const activeJadwal = scheduledJadwals.find(j => String(j.id) === String(state.filterMonitoringJadwal)) || null;

    let pesertaRows = [];
    let kelasTargets = [];
    if (activeJadwal) {
      const mp = state.db.mata_pelajaran.find(m => String(m.id) === String(activeJadwal.mapel_id));
      kelasTargets = parseKelasString(activeJadwal.kelas_target);
      const kelasTargetsLower = kelasTargets.map(t => t.toLowerCase());
      const jadwalSesiList = state.db.sesi_ujian_siswa.filter(s => String(s.jadwal_id) === String(activeJadwal.id));
      const sesiBySiswaId = new Map(jadwalSesiList.map(s => [String(s.siswa_id), s]));

      let siswaCandidates = state.db.users_cbt.filter(u => {
        if (u.role !== 'siswa') return false;
        if (sesiBySiswaId.has(String(u.id))) return true;
        if (kelasTargetsLower.length > 0 && !kelasTargetsLower.includes('semua')) {
          if (!kelasTargetsLower.includes((u.kelas || '').trim().toLowerCase())) return false;
        }
        if (mp && mp.kategori_agama) {
          if ((u.agama || '').trim().toLowerCase() !== mp.kategori_agama.trim().toLowerCase()) return false;
        }
        return true;
      });

      if (state.filterMonitoringKelas) {
        const fKelas = state.filterMonitoringKelas.trim().toLowerCase();
        siswaCandidates = siswaCandidates.filter(s => (s.kelas || '').trim().toLowerCase() === fKelas);
      }
      if (state.filterMonitoringSearch) {
        const q = state.filterMonitoringSearch.toLowerCase();
        siswaCandidates = siswaCandidates.filter(s =>
          (s.nama_lengkap || '').toLowerCase().includes(q) ||
          (s.nomor_peserta || '').toLowerCase().includes(q)
        );
      }

      pesertaRows = siswaCandidates.map(siswa => {
        const sesi = sesiBySiswaId.get(String(siswa.id)) || null;
        return { siswa, sesi };
      });
    }

    return `
      <div class="space-y-4">
        <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-r from-white via-sky-50/75 to-cyan-50/75 border border-sky-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-sky-950 via-cyan-900 to-teal-900 bg-clip-text text-transparent">Monitoring Peserta Ujian (Live)</h2>
            <p class="text-xs text-slate-600">Ketuk kartu mapel terjadwal sekali untuk membuka pemantauan peserta, ketuk sekali lagi untuk menutupnya.</p>
          </div>
          <button id="btn-refresh-monitoring" type="button" class="px-3.5 py-2 rounded-xl bg-gradient-to-r from-cyan-600 to-sky-600 hover:opacity-95 text-white font-bold text-xs transition shadow-2xs self-start sm:self-auto">
            ↻ Segarkan Live
          </button>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
          ${scheduledJadwals.length === 0 ? `
            <div class="col-span-full glass-panel rounded-2xl p-6 text-center text-xs text-slate-500">Belum ada jadwal ujian yang dibuat.</div>
          ` : scheduledJadwals.map(j => {
            const mp = state.db.mata_pelajaran.find(m => m.id === j.mapel_id) || { kode_mapel: '-', nama_mapel: '-' };
            const isSelected = state.filterMonitoringJadwal === j.id;
            const effStatus = getEffectiveJadwalStatus(j);
            const sesiList = state.db.sesi_ujian_siswa.filter(s => s.jadwal_id === j.id);
            const workingCount = sesiList.filter(s => s.status === 'mengerjakan').length;
            const doneCount = sesiList.filter(s => s.status === 'selesai').length;
            const lockedCount = sesiList.filter(s => s.status === 'terkunci').length;

            return `
              <button type="button" data-select-mon-jadwal="${j.id}"
                class="p-4 rounded-2xl border-2 text-left transition-all bg-gradient-to-br ${
                  isSelected
                    ? 'from-sky-100/85 via-cyan-50 to-emerald-50 border-sky-600 shadow-sm'
                    : 'from-white to-sky-50/45 border-slate-200 hover:border-sky-300 shadow-2xs'
                }">
                <div class="flex items-center justify-between gap-2">
                  <span class="px-2 py-0.5 rounded-lg bg-sky-900 text-white font-mono text-[11px] font-extrabold">${escapeHtml(mp.kode_mapel)}</span>
                  <span class="px-2.5 py-0.5 rounded-full text-[10px] font-extrabold ${
                    effStatus === 'Aktif' ? 'bg-emerald-100 text-emerald-800 border border-emerald-200' : 'bg-slate-200 text-slate-700'
                  }">${effStatus}</span>
                </div>
                <div class="font-extrabold text-sm text-slate-900 mt-2">${escapeHtml(mp.nama_mapel)} - <span class="text-sky-900">${escapeHtml(j.judul_ujian)}</span></div>
                <div class="text-[11px] text-slate-600 mt-1">Kelas: <strong class="text-indigo-800">${escapeHtml(j.kelas_target)}</strong> | Token: <strong class="font-mono px-1.5 py-0.5 rounded bg-amber-100 text-amber-950">${escapeHtml(j.token_ujian)}</strong></div>
                <div class="flex flex-wrap items-center justify-between gap-1.5 mt-2.5 text-[10px] font-bold">
                  <div class="flex flex-wrap items-center gap-1.5">
                    <span class="px-2 py-0.5 rounded bg-sky-100/80 text-sky-900">Mengerjakan: ${workingCount}</span>
                    <span class="px-2 py-0.5 rounded bg-emerald-100/80 text-emerald-900">Selesai: ${doneCount}</span>
                    ${lockedCount > 0 ? `<span class="px-2 py-0.5 rounded bg-rose-600 text-white animate-pulse">🔒 Terkunci: ${lockedCount}</span>` : ''}
                  </div>
                  <span class="${isSelected ? 'text-sky-800 font-extrabold' : 'text-slate-400'}">${isSelected ? '▾ Terbuka' : '▸ Buka'}</span>
                </div>
              </button>
            `;
          }).join('')}
        </div>

        ${!activeJadwal ? `
          <div class="rounded-2xl p-8 text-center space-y-1.5 bg-gradient-to-br from-white via-sky-50/40 to-cyan-50/40 border border-sky-200/70">
            <div class="text-2xl">📡</div>
            <div class="font-extrabold text-sm text-slate-800">Ketuk Salah Satu Jadwal Mapel di Atas</div>
            <p class="text-xs text-slate-500">Ketuk sekali untuk membuka pemantauan peserta, dan ketuk sekali lagi pada kartu yang sama untuk menutupnya.</p>
          </div>
        ` : `
          <div class="glass-panel rounded-2xl p-4 sm:p-5 space-y-4 border-sky-200/70">
            <div class="flex flex-wrap items-center justify-between gap-2 pb-3 border-b border-slate-200">
              <div class="flex flex-wrap items-center gap-2">
                <select id="filter-mon-kelas" class="px-3 py-2 rounded-xl border border-sky-200 bg-sky-50/40 text-xs font-bold text-sky-950">
                  <option value="">Semua Kelas (${escapeHtml(activeJadwal.kelas_target)})</option>
                  ${kelasTargets.map(k => `<option value="${k}" ${state.filterMonitoringKelas === k ? 'selected' : ''}>Kelas ${k}</option>`).join('')}
                </select>
                <input id="filter-mon-search" type="text" value="${escapeHtml(state.filterMonitoringSearch)}" placeholder="Cari nama / no. peserta..."
                  class="px-3 py-2 rounded-xl border border-slate-300 bg-white text-xs font-semibold text-slate-800" />
              </div>
              <div class="px-3 py-1 rounded-full bg-indigo-50 text-indigo-900 border border-indigo-200 text-xs font-extrabold">Total Peserta: ${pesertaRows.length}</div>
            </div>

            <div class="overflow-x-auto">
              <table class="w-full text-xs text-left border-collapse">
                <thead>
                  <tr class="bg-gradient-to-r from-sky-50 to-indigo-50 text-slate-800 border-b border-sky-200/80">
                    <th class="p-2.5">Peserta</th>
                    <th class="p-2.5 text-center">Kelas</th>
                    <th class="p-2.5 text-center">Status Ujian</th>
                    <th class="p-2.5 text-center">Dijawab</th>
                    <th class="p-2.5 text-center">Pelanggaran</th>
                    <th class="p-2.5 text-right">Tindakan</th>
                  </tr>
                </thead>
                <tbody class="divide-y divide-slate-200">
                  ${pesertaRows.length === 0 ? `
                    <tr><td colspan="6" class="p-6 text-center text-slate-500">Tidak ada peserta didik pada filter ini.</td></tr>
                  ` : pesertaRows.map(({ siswa, sesi }) => {
                    const ansCount = sesi && sesi.jawaban_json ? Object.keys(sesi.jawaban_json).length : 0;
                    const violCount = sesi ? (Number(sesi.jumlah_pelanggaran) || 0) : 0;
                    const maxViolation = getSchoolConfig().max_pelanggaran || 3;
                    const isSessionLocked = sesi && (sesi.status === 'terkunci' || violCount >= maxViolation);
                    const statusBadge = !sesi
                      ? `<span class="px-2.5 py-0.5 rounded-full bg-slate-100 text-slate-600 font-bold">Belum Mulai</span>`
                      : isSessionLocked
                        ? `<span class="px-2.5 py-0.5 rounded-full bg-rose-600 text-white font-extrabold animate-pulse">🔒 TERKUNCI</span>`
                        : sesi.status === 'selesai'
                          ? `<span class="px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 font-extrabold">✓ Selesai (${sesi.nilai_akhir ?? 0})</span>`
                          : `<span class="px-2.5 py-0.5 rounded-full bg-sky-100 text-sky-800 font-extrabold">⏳ Mengerjakan</span>`;

                    return `
                      <tr class="hover:bg-sky-50/40">
                        <td class="p-2.5">
                          <div class="font-extrabold text-slate-900">${escapeHtml(siswa.nama_lengkap)}</div>
                          <div class="text-[11px] text-indigo-700/80 font-mono">${escapeHtml(siswa.nomor_peserta || '-')}</div>
                        </td>
                        <td class="p-2.5 text-center font-extrabold text-indigo-900">${escapeHtml(siswa.kelas || '-')}</td>
                        <td class="p-2.5 text-center">${statusBadge}</td>
                        <td class="p-2.5 text-center font-bold text-slate-700">${sesi ? `${ansCount} Soal` : '-'}</td>
                        <td class="p-2.5 text-center">
                          ${violCount > 0
                            ? `<span class="px-2 py-0.5 rounded-lg bg-rose-50 text-rose-700 border border-rose-200 font-extrabold">${violCount}x</span>`
                            : `<span class="text-slate-400">0</span>`}
                        </td>
                        <td class="p-2.5 text-right space-x-1">
                          ${isSessionLocked ? `
                            <button type="button" data-unlock-sesi="${sesi.id}" class="px-2.5 py-1 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-[11px]">
                              🔓 Buka Kunci
                            </button>
                          ` : ''}
                          ${siswa.is_logged_in ? `
                            <button type="button" data-reset-login="${siswa.id}" class="px-2.5 py-1 rounded-lg bg-amber-100 hover:bg-amber-200 text-amber-900 font-bold text-[11px]">
                              Reset Login
                            </button>
                          ` : ''}
                        </td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          </div>
        `}
      </div>
    `;
  }

  // =========================================================================
  // CETAK DOKUMEN A4
  // =========================================================================
  function renderCetakDokumenView() {
    const classes = getAllMasterClasses();
    const jadwals = state.db.jadwal_ujian;
    if (!state.beritaAcaraForm) {
      state.beritaAcaraForm = getDefaultBeritaAcaraData();
    }

    let previewHtml = '';
    if (state.docType === 'daftar_hadir') previewHtml = buildDaftarHadirHtml(state.docKelas, state.docJadwalId);
    else if (state.docType === 'kartu_peserta') previewHtml = buildKartuPesertaHtml(state.docKelas);
    else previewHtml = buildBeritaAcaraHtml();

    return `
      <div class="space-y-4">
        <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-r from-white via-teal-50/75 to-emerald-50/75 border border-teal-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-teal-950 via-emerald-900 to-sky-900 bg-clip-text text-transparent">Cetak Dokumen Resmi CBT / ANBK</h2>
            <p class="text-xs text-slate-600">Pilih jenis dokumen resmi A4 (Daftar Hadir, Kartu Peserta, atau Berita Acara Pelaksanaan) lalu cetak / simpan ke PDF.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            ${state.docType === 'berita_acara' ? `
              <button id="btn-toggle-edit-ba" type="button" class="px-3.5 py-2 rounded-xl bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 font-bold text-xs transition">
                ${state.editingBeritaAcara ? '✕ Tutup Form Edit' : '✎ Edit Berita Acara'}
              </button>
            ` : ''}
            <button id="btn-print-doc-a4" type="button" class="px-4 py-2 rounded-xl bg-gradient-to-r from-[#4f6d4a] to-teal-700 hover:opacity-95 text-white font-bold text-xs shadow-xs transition">
              🖨 Cetak / Simpan PDF (A4)
            </button>
          </div>
        </div>

        <div class="rounded-2xl p-4 bg-gradient-to-r from-white via-sky-50/50 to-amber-50/50 border border-teal-200/70 flex flex-wrap items-center gap-3 shadow-2xs">
          <div class="flex flex-wrap gap-1.5">
            <button type="button" data-doc-type="daftar_hadir" class="px-3 py-2 rounded-xl text-xs font-bold transition ${state.docType === 'daftar_hadir' ? 'bg-gradient-to-r from-[#4f6d4a] to-emerald-700 text-white shadow-2xs' : 'bg-emerald-50/80 text-emerald-900 border border-emerald-200/70 hover:bg-emerald-100/70'}">
              📋 Daftar Hadir Peserta
            </button>
            <button type="button" data-doc-type="kartu_peserta" class="px-3 py-2 rounded-xl text-xs font-bold transition ${state.docType === 'kartu_peserta' ? 'bg-gradient-to-r from-sky-600 to-indigo-600 text-white shadow-2xs' : 'bg-sky-50/80 text-sky-900 border border-sky-200/70 hover:bg-sky-100/70'}">
              🪪 Kartu Peserta (8/Hal)
            </button>
            <button type="button" data-doc-type="berita_acara" class="px-3 py-2 rounded-xl text-xs font-bold transition ${state.docType === 'berita_acara' ? 'bg-gradient-to-r from-amber-600 to-orange-500 text-white shadow-2xs' : 'bg-amber-50/80 text-amber-900 border border-amber-200/70 hover:bg-amber-100/70'}">
              📄 Berita Acara Ujian
            </button>
          </div>

          ${state.docType === 'daftar_hadir' ? `
            <select id="select-doc-jadwal" class="px-3 py-2 rounded-xl border border-teal-200 bg-teal-50/40 text-xs font-bold text-teal-950">
              <option value="" ${!state.docJadwalId ? 'selected' : ''}>Semua Jadwal / Umum</option>
              ${jadwals.map(j => {
                const mp = state.db.mata_pelajaran.find(m => m.id === j.mapel_id);
                return `<option value="${j.id}" ${state.docJadwalId === j.id ? 'selected' : ''}>${escapeHtml((mp ? mp.nama_mapel + ' - ' : '') + j.judul_ujian)}</option>`;
              }).join('')}
            </select>
          ` : ''}

          ${state.docType !== 'berita_acara' ? `
            <select id="select-doc-kelas" class="px-3 py-2 rounded-xl border border-indigo-200 bg-indigo-50/40 text-xs font-bold text-indigo-950">
              <option value="SEMUA" ${state.docKelas === 'SEMUA' ? 'selected' : ''}>Semua Kelas</option>
              ${classes.map(k => `<option value="${k}" ${state.docKelas === k ? 'selected' : ''}>Kelas ${k}</option>`).join('')}
            </select>
          ` : ''}
        </div>

        ${state.docType === 'berita_acara' && state.editingBeritaAcara ? `
          <form id="form-berita-acara" class="glass-panel rounded-2xl p-4 sm:p-5 space-y-3 text-xs">
            <div class="font-extrabold text-sm text-slate-900">Edit Isi Berita Acara Pelaksanaan Ujian</div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold mb-1">Hari / Tanggal</label>
                <input name="hari_tanggal" value="${escapeHtml(state.beritaAcaraForm.hari_tanggal)}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
              <div>
                <label class="block font-bold mb-1">Mata Pelajaran / Ujian</label>
                <input name="mapel_ujian" value="${escapeHtml(state.beritaAcaraForm.mapel_ujian)}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
              <div>
                <label class="block font-bold mb-1">Ruang Ujian &amp; Sesi</label>
                <input name="ruang_sesi" value="${escapeHtml(state.beritaAcaraForm.ruang_sesi)}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
              <div>
                <label class="block font-bold mb-1">Waktu Pelaksanaan</label>
                <input name="waktu_pelaksanaan" value="${escapeHtml(state.beritaAcaraForm.waktu_pelaksanaan)}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
              <div>
                <label class="block font-bold mb-1">Jumlah Peserta Seharusnya</label>
                <input name="jumlah_seharusnya" value="${escapeHtml(state.beritaAcaraForm.jumlah_seharusnya)}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
              <div>
                <label class="block font-bold mb-1">Jumlah Peserta Hadir</label>
                <input name="jumlah_hadir" value="${escapeHtml(state.beritaAcaraForm.jumlah_hadir)}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
              <div>
                <label class="block font-bold mb-1">Jumlah Peserta Tidak Hadir</label>
                <input name="jumlah_tidak_hadir" value="${escapeHtml(state.beritaAcaraForm.jumlah_tidak_hadir)}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
              <div>
                <label class="block font-bold mb-1">Tempat &amp; Tanggal Tanda Tangan</label>
                <input name="tempat_tanggal_ttd" value="${escapeHtml(state.beritaAcaraForm.tempat_tanggal_ttd)}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
              <div class="sm:col-span-2">
                <label class="block font-bold mb-1">Rincian Peserta Tidak Hadir</label>
                <input name="rincian_tidak_hadir" value="${escapeHtml(state.beritaAcaraForm.rincian_tidak_hadir)}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
              <div class="sm:col-span-2">
                <label class="block font-bold mb-1">Catatan Kejadian Khusus Selama Pelaksanaan</label>
                <textarea name="catatan_khusus" rows="2" class="w-full p-2.5 rounded-xl border border-slate-300">${escapeHtml(state.beritaAcaraForm.catatan_khusus)}</textarea>
              </div>
            </div>
            <div class="flex justify-end pt-1">
              <button type="submit" class="px-4 py-2 rounded-xl bg-[#4f6d4a] text-white font-extrabold">💾 Terapkan ke Pratinjau A4</button>
            </div>
          </form>
        ` : ''}

        <div class="rounded-2xl p-4 sm:p-6 overflow-x-auto bg-gradient-to-br from-slate-100/80 via-teal-50/50 to-sky-50/50 border border-slate-200/80">
          <div id="print-area" class="max-w-[210mm] min-w-[680px] mx-auto bg-white text-black p-6 sm:p-10 shadow-md border border-slate-300" style="font-family: 'Times New Roman', Times, serif;">
            ${previewHtml}
          </div>
        </div>
      </div>
    `;
  }

  // =========================================================================
  // DATA GURU & SISWA (ADMIN)
  // =========================================================================
  function renderUsersView() {
    const isSiswa = state.filterUserTab === 'siswa';
    const classes = getAllMasterClasses();
    let list = state.db.users_cbt.filter(u => u.role === state.filterUserTab);
    if (isSiswa && state.filterUserKelas) {
      list = list.filter(s => (s.kelas || '').toLowerCase() === state.filterUserKelas.toLowerCase());
    }

    return `
      <div class="space-y-4">
        <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-r from-white via-rose-50/65 to-amber-50/65 border border-rose-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-rose-950 via-slate-900 to-amber-900 bg-clip-text text-transparent">Data Guru &amp; Peserta Didik</h2>
            <p class="text-xs text-slate-600">Kelola akun login guru pengampu dan kartu peserta siswa secara individu maupun impor massal Excel.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button id="btn-tpl-user-excel" type="button" class="px-3 py-2 rounded-xl bg-white hover:bg-amber-50 border border-amber-200 text-amber-900 font-bold text-xs shadow-2xs">
              ⬇ Template Excel
            </button>
            <label class="px-3 py-2 rounded-xl bg-gradient-to-r from-sky-50 to-indigo-50 hover:from-sky-100 hover:to-indigo-100 text-indigo-900 border border-indigo-200 font-bold text-xs cursor-pointer">
              📥 Import Excel
              <input id="input-import-user-excel" type="file" accept=".xlsx,.xls" class="hidden" />
            </label>
            <button id="btn-open-add-user" type="button" class="px-3.5 py-2 rounded-xl bg-gradient-to-r from-rose-600 to-pink-600 hover:opacity-95 text-white font-extrabold text-xs shadow-xs">
              + Tambah ${isSiswa ? 'Siswa' : 'Guru'}
            </button>
          </div>
        </div>

        <div class="glass-panel rounded-2xl p-4 space-y-3 border-rose-200/60">
          <div class="flex flex-wrap items-center justify-between gap-2 pb-3 border-b border-slate-200">
            <div class="flex gap-1.5">
              <button type="button" data-user-tab="siswa" class="px-3.5 py-2 rounded-xl text-xs font-extrabold transition ${isSiswa ? 'bg-gradient-to-r from-emerald-600 to-teal-600 text-white shadow-2xs' : 'bg-emerald-50/70 text-emerald-900 border border-emerald-200/70'}">
                🎓 Peserta Didik (${state.db.users_cbt.filter(u => u.role === 'siswa').length})
              </button>
              <button type="button" data-user-tab="guru" class="px-3.5 py-2 rounded-xl text-xs font-extrabold transition ${!isSiswa ? 'bg-gradient-to-r from-indigo-600 to-blue-600 text-white shadow-2xs' : 'bg-indigo-50/70 text-indigo-900 border border-indigo-200/70'}">
                👨‍🏫 Guru Pengampu (${state.db.users_cbt.filter(u => u.role === 'guru').length})
              </button>
            </div>
            ${isSiswa ? `
              <select id="filter-user-kelas" class="px-3 py-2 rounded-xl border border-emerald-200 bg-emerald-50/40 text-xs font-bold text-emerald-950">
                <option value="">Semua Kelas</option>
                ${classes.map(k => `<option value="${k}" ${state.filterUserKelas === k ? 'selected' : ''}>Kelas ${k}</option>`).join('')}
              </select>
            ` : ''}
          </div>

          <div class="overflow-x-auto">
            <table class="w-full text-xs text-left border-collapse">
              <thead>
                <tr class="bg-gradient-to-r from-rose-50/80 via-amber-50/80 to-sky-50/80 text-slate-800 border-b border-rose-200/70">
                  <th class="p-2.5">Nama Lengkap</th>
                  <th class="p-2.5">${isSiswa ? 'No. Peserta / NISN' : 'NIP'}</th>
                  ${isSiswa ? '<th class="p-2.5 text-center">Kelas</th><th class="p-2.5 text-center">Agama</th>' : ''}
                  <th class="p-2.5">Username</th>
                  <th class="p-2.5">Password</th>
                  <th class="p-2.5 text-right">Aksi</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-200">
                ${list.length === 0 ? `
                  <tr><td colspan="7" class="p-6 text-center text-slate-500">Belum ada data ${state.filterUserTab}.</td></tr>
                ` : list.map(u => `
                  <tr class="hover:bg-rose-50/30">
                    <td class="p-2.5 font-extrabold text-slate-900">${escapeHtml(u.nama_lengkap)}</td>
                    <td class="p-2.5 font-mono text-[11px] text-indigo-800/80">${isSiswa ? `${escapeHtml(u.nomor_peserta || '-')} / ${escapeHtml(u.nip_nisn || '-')}` : escapeHtml(u.nip_nisn || '-')}</td>
                    ${isSiswa ? `<td class="p-2.5 text-center font-extrabold text-emerald-800">${escapeHtml(u.kelas || '-')}</td><td class="p-2.5 text-center text-slate-700">${escapeHtml(u.agama || 'Islam')}</td>` : ''}
                    <td class="p-2.5 font-mono font-bold text-sky-900">${escapeHtml(u.username)}</td>
                    <td class="p-2.5 font-mono text-amber-900">${escapeHtml(u.password)}</td>
                    <td class="p-2.5 text-right space-x-1">
                      <button type="button" data-edit-user="${u.id}" class="px-2.5 py-1 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 font-bold">✎</button>
                      <button type="button" data-del-user="${u.id}" class="px-2.5 py-1 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 font-bold">🗑</button>
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;
  }

  // =========================================================================
  // MATA PELAJARAN
  // =========================================================================
  function renderMapelView() {
    const mapels = getAuthorizedMapels();
    const mapelCardPalette = [
      'from-white via-blue-50/55 to-indigo-50/55 border-indigo-200/80',
      'from-white via-emerald-50/55 to-teal-50/55 border-emerald-200/80',
      'from-white via-amber-50/55 to-orange-50/55 border-amber-200/80',
      'from-white via-purple-50/55 to-pink-50/55 border-purple-200/80'
    ];
    return `
      <div class="space-y-4">
        <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-r from-white via-blue-50/75 to-indigo-50/75 border border-blue-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-blue-950 via-indigo-900 to-teal-900 bg-clip-text text-transparent">Mata Pelajaran &amp; Kelas Ampu</h2>
            <p class="text-xs text-slate-600">Atur mata pelajaran, guru pengampu, dan multi-kelas yang diampu oleh setiap guru.</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button id="btn-tpl-mapel-excel" type="button" class="px-3 py-2 rounded-xl bg-white hover:bg-blue-50 border border-blue-200 text-blue-900 font-bold text-xs shadow-2xs">
              ⬇ Template Excel
            </button>
            <label class="px-3 py-2 rounded-xl bg-gradient-to-r from-sky-50 to-indigo-50 hover:from-sky-100 hover:to-indigo-100 text-indigo-900 border border-indigo-200 font-bold text-xs cursor-pointer">
              📥 Import Excel
              <input id="input-import-mapel-excel" type="file" accept=".xlsx,.xls" class="hidden" />
            </label>
            <button id="btn-open-add-mapel" type="button" class="px-3.5 py-2 rounded-xl bg-gradient-to-r from-blue-600 to-indigo-600 hover:opacity-95 text-white font-extrabold text-xs shadow-xs">
              + Tambah Mapel
            </button>
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
          ${mapels.length === 0 ? `
            <div class="col-span-full glass-panel rounded-2xl p-8 text-center text-xs text-slate-500">Belum ada mata pelajaran terdaftar.</div>
          ` : mapels.map((m, idx) => {
            const guru = state.db.users_cbt.find(u => u.id === m.guru_id);
            const soalCount = state.db.bank_soal.filter(s => s.mapel_id === m.id).length;
            const palette = mapelCardPalette[idx % mapelCardPalette.length];
            return `
              <div class="rounded-2xl p-4 space-y-2.5 bg-gradient-to-br ${palette} border shadow-2xs">
                <div class="flex items-start justify-between gap-2">
                  <div>
                    <div class="flex flex-wrap items-center gap-1.5">
                      <span class="px-2.5 py-0.5 rounded-lg bg-indigo-900 text-white font-mono text-[11px] font-extrabold">${escapeHtml(m.kode_mapel)}</span>
                      ${m.kategori_agama ? `<span class="px-2 py-0.5 rounded-lg bg-amber-100/90 text-amber-900 border border-amber-200 text-[10px] font-bold">Agama: ${escapeHtml(m.kategori_agama)}</span>` : ''}
                      ${m.is_english ? `<span class="px-2 py-0.5 rounded-lg bg-sky-100/90 text-sky-900 border border-sky-200 text-[10px] font-bold">Anti-Translate</span>` : ''}
                    </div>
                    <h3 class="font-extrabold text-base text-slate-900 mt-1.5">${escapeHtml(m.nama_mapel)}</h3>
                  </div>
                  <div class="flex gap-1">
                    <button type="button" data-edit-mapel="${m.id}" class="px-2.5 py-1 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 text-xs font-bold">✎</button>
                    <button type="button" data-del-mapel="${m.id}" class="px-2.5 py-1 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-bold">🗑</button>
                  </div>
                </div>
                <div class="text-xs text-slate-600 space-y-0.5">
                  <div>Guru Pengampu: <strong class="text-slate-900">${escapeHtml(guru ? guru.nama_lengkap : 'Belum ditentukan')}</strong></div>
                  <div>Kelas Ampu: <strong class="text-indigo-800">${escapeHtml(m.kelas_ampu || 'Semua Kelas')}</strong> | Bank Soal: <strong class="text-emerald-800">${soalCount} Butir</strong></div>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      </div>
    `;
  }

  // =========================================================================
  // JADWAL & TOKEN UJIAN
  // =========================================================================
  function renderJadwalView() {
    const authMapelIds = getAuthorizedMapelIds();
    const jadwals = state.db.jadwal_ujian.filter(j => authMapelIds.has(j.mapel_id));

    return `
      <div class="space-y-4">
        <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-r from-white via-amber-50/70 to-emerald-50/60 border border-amber-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-amber-950 via-slate-900 to-emerald-900 bg-clip-text text-transparent">Jadwal &amp; Token Ujian</h2>
            <p class="text-xs text-slate-600">Atur waktu pelaksanaan, acak soal, acak opsi, serta token masuk. Jadwal otomatis tertutup jika batas waktu berakhir.</p>
          </div>
          <button id="btn-open-add-jadwal" type="button" class="px-4 py-2.5 rounded-xl bg-gradient-to-r from-[#4f6d4a] to-emerald-700 hover:opacity-95 text-white font-extrabold text-xs shadow-xs transition">
            + Buat Jadwal &amp; Token
          </button>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
          ${jadwals.length === 0 ? `
            <div class="col-span-full glass-panel rounded-2xl p-8 text-center text-xs text-slate-500">Belum ada jadwal ujian dibuat.</div>
          ` : jadwals.map(j => {
            const mp = state.db.mata_pelajaran.find(m => m.id === j.mapel_id) || { nama_mapel: '-' };
            const effStatus = getEffectiveJadwalStatus(j);
            const isAktif = effStatus === 'Aktif';
            return `
              <div class="rounded-2xl p-4 space-y-3 border shadow-2xs bg-gradient-to-br ${
                isAktif
                  ? 'from-white via-emerald-50/45 to-amber-50/45 border-emerald-200/80'
                  : 'from-white to-slate-50 border-slate-200'
              }">
                <div class="flex items-start justify-between gap-2">
                  <div>
                    <div class="flex flex-wrap items-center gap-1.5">
                      <span class="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold ${
                        isAktif ? 'bg-emerald-100 text-emerald-800 border border-emerald-200' : 'bg-rose-100 text-rose-800'
                      }">${isAktif ? '● Aktif' : '■ Selesai / Tutup'}</span>
                      ${j.acak_soal ? `<span class="px-2 py-0.5 rounded-lg bg-sky-50 text-sky-800 border border-sky-200 text-[10px] font-bold">🔀 Acak Soal</span>` : ''}
                      ${j.acak_opsi ? `<span class="px-2 py-0.5 rounded-lg bg-purple-50 text-purple-800 border border-purple-200 text-[10px] font-bold">🎲 Acak Opsi</span>` : ''}
                    </div>
                    <h3 class="font-extrabold text-base text-slate-900 mt-1.5">${escapeHtml(mp.nama_mapel)} - <span class="text-emerald-800">${escapeHtml(j.judul_ujian)}</span></h3>
                  </div>
                  <div class="text-right shrink-0">
                    <div class="text-[10px] font-extrabold text-amber-800">TOKEN</div>
                    <div class="px-3 py-1 rounded-xl bg-gradient-to-r from-amber-100 to-yellow-100 border border-amber-300 font-mono text-sm font-extrabold text-amber-950 shadow-2xs">${escapeHtml(j.token_ujian)}</div>
                  </div>
                </div>

                <div class="text-xs text-slate-600 space-y-1 bg-white/80 p-3 rounded-xl border border-slate-200/80">
                  <div>Target Kelas: <strong class="text-indigo-800">${escapeHtml(j.kelas_target)}</strong> | Durasi: <strong class="text-sky-800">${j.durasi_menit} Menit</strong></div>
                  <div>Mulai: <strong class="text-slate-800">${formatDateTimeWita(j.waktu_mulai)}</strong></div>
                  <div>Batas Selesai: <strong class="text-rose-800">${formatDateTimeWita(j.waktu_selesai)}</strong></div>
                </div>

                <div class="flex items-center justify-end gap-1.5">
                  <button type="button" data-edit-jadwal="${j.id}" class="px-3 py-1.5 rounded-xl bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-200 text-xs font-bold">
                    ✎ Edit / Perpanjang Waktu
                  </button>
                  <button type="button" data-del-jadwal="${j.id}" class="px-2.5 py-1.5 rounded-xl bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-bold">
                    🗑
                  </button>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      </div>
    `;
  }

  // =========================================================================
  // REKAPITULASI NILAI & KOREKSI URAIAN
  // =========================================================================
  function renderRekapNilaiView() {
    const authMapelIds = getAuthorizedMapelIds();
    const jadwals = state.db.jadwal_ujian.filter(j => authMapelIds.has(j.mapel_id));
    const hasValidSelected = jadwals.some(j => String(j.id) === String(state.filterNilaiJadwal));
    if (!hasValidSelected && jadwals.length > 0) {
      const jadwalWithSesi = jadwals.find(j =>
        state.db.sesi_ujian_siswa.some(s => String(s.jadwal_id) === String(j.id))
      );
      state.filterNilaiJadwal = (jadwalWithSesi || jadwals[0]).id;
      state.filterNilaiKelas = '';
    }
    const activeJadwal = jadwals.find(j => String(j.id) === String(state.filterNilaiJadwal)) || null;
    const classes = activeJadwal ? parseKelasString(activeJadwal.kelas_target) : getAllMasterClasses();
    if (state.filterNilaiKelas && !classes.some(k => k.toLowerCase() === state.filterNilaiKelas.toLowerCase())) {
      state.filterNilaiKelas = '';
    }

    let rowsData = [];
    if (activeJadwal) {
      const sesiList = state.db.sesi_ujian_siswa.filter(s => String(s.jadwal_id) === String(activeJadwal.id));
      rowsData = sesiList
        .map(sesi => {
          const siswa = state.db.users_cbt.find(u => String(u.id) === String(sesi.siswa_id));
          return siswa ? { siswa, sesi } : null;
        })
        .filter(Boolean);

      if (state.filterNilaiKelas) {
        rowsData = rowsData.filter(r => (r.siswa.kelas || '').trim().toLowerCase() === state.filterNilaiKelas.trim().toLowerCase());
      }
      rowsData.sort((a, b) => (a.siswa.nama_lengkap || '').localeCompare(b.siswa.nama_lengkap || ''));
    }

    return `
      <div class="space-y-4">
        <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-r from-white via-violet-50/70 to-purple-50/70 border border-violet-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-violet-950 via-purple-900 to-indigo-900 bg-clip-text text-transparent">Rekapitulasi Nilai &amp; Koreksi Essay</h2>
            <p class="text-xs text-slate-600">Lihat hasil skor otomatis, input nilai soal uraian, reset ujian ulang, dan cetak/ekspor daftar nilai.</p>
          </div>
          ${activeJadwal ? `
            <div class="flex flex-wrap gap-2">
              <button id="btn-refresh-nilai" type="button" class="px-3 py-2 rounded-xl bg-white hover:bg-violet-50 border border-violet-200 text-violet-900 font-bold text-xs shadow-2xs">
                ↻ Segarkan
              </button>
              <button id="btn-export-nilai-excel" type="button" class="px-3 py-2 rounded-xl bg-gradient-to-r from-emerald-50 to-teal-50 hover:from-emerald-100 hover:to-teal-100 text-emerald-900 border border-emerald-200 font-bold text-xs">
                📊 Export Excel
              </button>
              <button id="btn-print-nilai-pdf" type="button" class="px-3.5 py-2 rounded-xl bg-gradient-to-r from-violet-600 to-purple-600 hover:opacity-95 text-white font-bold text-xs shadow-xs">
                🖨 Cetak Daftar Nilai A4
              </button>
            </div>
          ` : ''}
        </div>

        <div class="glass-panel rounded-2xl p-4 sm:p-5 space-y-4 border-violet-200/60">
          <div class="flex flex-wrap items-center gap-2.5 pb-3 border-b border-slate-200">
            <select id="filter-nilai-jadwal" class="px-3 py-2 rounded-xl border border-violet-200 bg-violet-50/40 text-xs font-bold text-violet-950">
              ${jadwals.length === 0 ? '<option value="">Belum ada jadwal ujian</option>' : jadwals.map(j => {
                const mp = state.db.mata_pelajaran.find(m => String(m.id) === String(j.mapel_id));
                const pesertaCount = state.db.sesi_ujian_siswa.filter(s => String(s.jadwal_id) === String(j.id)).length;
                return `<option value="${j.id}" ${String(state.filterNilaiJadwal) === String(j.id) ? 'selected' : ''}>${escapeHtml((mp ? mp.nama_mapel + ' - ' : '') + j.judul_ujian + ' (' + pesertaCount + ' Peserta)')}</option>`;
              }).join('')}
            </select>

            <select id="filter-nilai-kelas" class="px-3 py-2 rounded-xl border border-indigo-200 bg-indigo-50/40 text-xs font-bold text-indigo-950">
              <option value="">Semua Kelas</option>
              ${classes.map(k => `<option value="${k}" ${state.filterNilaiKelas === k ? 'selected' : ''}>Kelas ${k}</option>`).join('')}
            </select>
          </div>

          <div class="overflow-x-auto">
            <table class="w-full text-xs text-left border-collapse">
              <thead>
                <tr class="bg-gradient-to-r from-violet-50 via-purple-50 to-indigo-50 text-slate-800 border-b border-violet-200/80">
                  <th class="p-2.5">Nama Peserta</th>
                  <th class="p-2.5 text-center">Kelas</th>
                  <th class="p-2.5 text-center">Status</th>
                  <th class="p-2.5 text-center">Skor Objektif</th>
                  <th class="p-2.5 text-center">Skor Essay</th>
                  <th class="p-2.5 text-center">Nilai Akhir</th>
                  <th class="p-2.5 text-right">Aksi</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-200">
                ${rowsData.length === 0 ? `
                  <tr><td colspan="7" class="p-6 text-center text-slate-500">Belum ada peserta yang mengikuti ujian pada jadwal/kelas ini.</td></tr>
                ` : rowsData.map(({ siswa, sesi }) => `
                  <tr class="hover:bg-violet-50/35">
                    <td class="p-2.5">
                      <div class="font-extrabold text-slate-900">${escapeHtml(siswa.nama_lengkap)}</div>
                      <div class="text-[11px] text-indigo-700/80 font-mono">${escapeHtml(siswa.nip_nisn || '-')}</div>
                    </td>
                    <td class="p-2.5 text-center font-extrabold text-indigo-900">${escapeHtml(siswa.kelas || '-')}</td>
                    <td class="p-2.5 text-center">
                      <span class="px-2.5 py-0.5 rounded-full text-[11px] font-bold ${
                        sesi.status === 'selesai' ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'
                      }">${sesi.status === 'selesai' ? 'Selesai' : 'Mengerjakan'}</span>
                    </td>
                    <td class="p-2.5 text-center font-bold text-sky-900">${sesi.skor_pg ?? 0}</td>
                    <td class="p-2.5 text-center font-bold text-purple-900">${sesi.skor_essay ?? 0}</td>
                    <td class="p-2.5 text-center font-extrabold text-sm text-emerald-800">${sesi.nilai_akhir ?? 0}</td>
                    <td class="p-2.5 text-right space-x-1">
                      <button type="button" data-grade-essay="${sesi.id}" class="px-2.5 py-1 rounded-lg bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-200 font-bold">
                        ✎ Skor Essay
                      </button>
                      <button type="button" data-reset-ujian="${sesi.id}" data-siswa-id="${siswa.id}" class="px-2.5 py-1 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 font-bold">
                        ↻ Ujian Ulang
                      </button>
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;
  }

  // =========================================================================
  // PENGATURAN SEKOLAH & LOGO PERMANEN
  // =========================================================================
  function renderPengaturanView() {
    const cfg = getSchoolConfig();
    const dinasPreview = state.modalLogos.dinas !== undefined ? state.modalLogos.dinas : cfg.logo_dinas_url;
    const sekolahPreview = state.modalLogos.sekolah !== undefined ? state.modalLogos.sekolah : cfg.logo_sekolah_url;
    const aplikasiPreview = state.modalLogos.aplikasi !== undefined ? state.modalLogos.aplikasi : cfg.logo_aplikasi_url;

    return `
      <form id="form-pengaturan-sekolah" class="space-y-4">
        <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-r from-white via-emerald-50/70 to-amber-50/70 border border-emerald-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-lg sm:text-xl font-extrabold bg-gradient-to-r from-slate-900 via-emerald-900 to-amber-900 bg-clip-text text-transparent">Pengaturan Identitas Sekolah &amp; Logo</h2>
            <p class="text-xs text-slate-600">Data ini digunakan otomatis pada Kop Surat Cetak A4, Kartu Peserta, dan batas toleransi pelanggaran ujian.</p>
          </div>
          <button type="submit" class="px-4 py-2.5 rounded-xl bg-gradient-to-r from-[#4f6d4a] to-emerald-700 hover:opacity-95 text-white font-extrabold text-xs shadow-xs">
            💾 Simpan Pengaturan
          </button>
        </div>

        <div class="glass-panel rounded-2xl p-4 sm:p-5 grid grid-cols-1 sm:grid-cols-2 gap-3.5 text-xs">
          <div>
            <label class="block font-bold mb-1">Nama Satuan Pendidikan</label>
            <input name="nama_sekolah" value="${escapeHtml(cfg.nama_sekolah)}" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 font-semibold" />
          </div>
          <div>
            <label class="block font-bold mb-1">NPSN</label>
            <input name="npsn" value="${escapeHtml(cfg.npsn)}" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 font-semibold" />
          </div>
          <div class="sm:col-span-2">
            <label class="block font-bold mb-1">Alamat Lengkap Sekolah</label>
            <input name="alamat" value="${escapeHtml(cfg.alamat)}" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300" />
          </div>
          <div>
            <label class="block font-bold mb-1">Tahun Pelajaran</label>
            <input name="tahun_pelajaran" value="${escapeHtml(cfg.tahun_pelajaran)}" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 font-semibold" />
          </div>
          <div>
            <label class="block font-bold mb-1">Semester</label>
            <select name="semester" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 bg-white font-semibold">
              <option value="Ganjil" ${cfg.semester === 'Ganjil' ? 'selected' : ''}>Ganjil</option>
              <option value="Genap" ${cfg.semester === 'Genap' ? 'selected' : ''}>Genap</option>
            </select>
          </div>
          <div>
            <label class="block font-bold mb-1">Nama Kepala Sekolah</label>
            <input name="nama_kepsek" value="${escapeHtml(cfg.nama_kepsek)}" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300" />
          </div>
          <div>
            <label class="block font-bold mb-1">NIP Kepala Sekolah</label>
            <input name="nip_kepsek" value="${escapeHtml(cfg.nip_kepsek)}" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300" />
          </div>
          <div>
            <label class="block font-bold mb-1">Nama Proktor CBT</label>
            <input name="nama_proktor" value="${escapeHtml(cfg.nama_proktor)}" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300" />
          </div>
          <div>
            <label class="block font-bold mb-1">NIP Proktor CBT</label>
            <input name="nip_proktor" value="${escapeHtml(cfg.nip_proktor)}" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300" />
          </div>
          <div>
            <label class="block font-bold mb-1">Nama Pengawas Ruang</label>
            <input name="nama_pengawas" value="${escapeHtml(cfg.nama_pengawas)}" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300" />
          </div>
          <div>
            <label class="block font-bold mb-1">NIP Pengawas Ruang</label>
            <input name="nip_pengawas" value="${escapeHtml(cfg.nip_pengawas)}" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300" />
          </div>
          <div>
            <label class="block font-bold mb-1">Batas Maksimal Pelanggaran (Kunci Otomatis)</label>
            <input type="number" min="1" max="10" name="max_pelanggaran" value="${cfg.max_pelanggaran}" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 font-bold" />
          </div>
        </div>

        <div class="glass-panel rounded-2xl p-4 sm:p-5 space-y-3 text-xs border-amber-200/60">
          <div class="font-extrabold text-sm text-amber-950">Logo Kop Surat &amp; Identitas Aplikasi</div>
          <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
            ${[
              { key: 'dinas', label: 'Logo Kabupaten / Dinas (Kiri)', preview: dinasPreview, fallback: svgLogoKukar(44), bg: 'from-emerald-50/60 to-teal-50/60 border-emerald-200/80' },
              { key: 'sekolah', label: 'Logo Sekolah (Kanan & Kartu)', preview: sekolahPreview, fallback: svgLogoSmp(44), bg: 'from-sky-50/60 to-indigo-50/60 border-sky-200/80' },
              { key: 'aplikasi', label: 'Logo Aplikasi (Login & Header)', preview: aplikasiPreview, fallback: svgLogoPromax(44), bg: 'from-amber-50/60 to-orange-50/60 border-amber-200/80' }
            ].map(item => `
              <div class="p-3.5 rounded-2xl bg-gradient-to-br ${item.bg} border flex flex-col items-center text-center gap-2">
                <div class="font-bold text-slate-800">${item.label}</div>
                <div class="w-16 h-16 rounded-2xl bg-white border border-slate-200 flex items-center justify-center overflow-hidden p-1 shadow-2xs">
                  ${item.preview ? `<img src="${item.preview}" class="w-full h-full object-contain" />` : item.fallback}
                </div>
                <div class="flex gap-1.5">
                  <label class="px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 font-bold cursor-pointer hover:bg-slate-100">
                    Pilih Gambar
                    <input type="file" accept="image/*" data-upload-logo="${item.key}" class="hidden" />
                  </label>
                  ${item.preview ? `<button type="button" data-clear-logo="${item.key}" class="px-2 py-1.5 rounded-lg bg-rose-50 text-rose-700 font-bold">Hapus</button>` : ''}
                </div>
              </div>
            `).join('')}
          </div>
        </div>
      </form>
    `;
  }

  // =========================================================================
  // DASHBOARD & RIWAYAT NILAI SISWA
  // =========================================================================
  function renderStudentDashboardView() {
    const siswa = state.currentUser;
    const studentKelas = (siswa.kelas || '').toLowerCase();
    const studentAgama = (siswa.agama || 'Islam').toLowerCase();

    const relevantJadwals = state.db.jadwal_ujian.filter(j => {
      const mp = state.db.mata_pelajaran.find(m => m.id === j.mapel_id);
      if (mp && mp.kategori_agama && mp.kategori_agama.toLowerCase() !== studentAgama) {
        return false;
      }
      const targets = parseKelasString(j.kelas_target).map(k => k.toLowerCase());
      return targets.length === 0 || targets.includes('semua') || targets.includes(studentKelas);
    });

    return `
      <div class="space-y-4">
        <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-r from-white via-emerald-50/75 to-teal-50/75 border border-emerald-200/80 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 class="text-base sm:text-lg font-extrabold bg-gradient-to-r from-emerald-950 via-teal-900 to-sky-900 bg-clip-text text-transparent">Jadwal Ujian &amp; Asesmen Tersedia</h2>
            <p class="text-xs text-slate-600">Pilih jadwal ujian yang aktif lalu masukkan Token Ujian dari pengawas untuk mulai mengerjakan.</p>
          </div>
          <button id="btn-refresh-siswa-jadwal" type="button" class="px-3.5 py-2 rounded-xl bg-white hover:bg-emerald-50 border border-emerald-200 text-emerald-900 font-bold text-xs shadow-2xs self-start sm:self-auto">
            ↻ Segarkan Jadwal
          </button>
        </div>

        ${relevantJadwals.length === 0 ? `
          <div class="glass-panel rounded-2xl p-8 text-center text-xs text-slate-500">Belum ada jadwal ujian untuk kelas Anda saat ini.</div>
        ` : `
          <div class="grid grid-cols-1 md:grid-cols-2 gap-3.5">
            ${relevantJadwals.map(j => {
              const mp = state.db.mata_pelajaran.find(m => m.id === j.mapel_id) || { nama_mapel: 'Mata Pelajaran' };
              const sesi = state.db.sesi_ujian_siswa.find(s => String(s.jadwal_id) === String(j.id) && String(s.siswa_id) === String(siswa.id));
              const effStatus = getEffectiveJadwalStatus(j);
              const isSelesai = sesi && sesi.status === 'selesai';
              const isMengerjakan = sesi && sesi.status !== 'selesai';
              const isAktif = effStatus === 'Aktif';

              return `
                <div class="rounded-2xl p-4 sm:p-5 bg-gradient-to-br from-white via-emerald-50/40 to-sky-50/40 border border-emerald-200/80 shadow-2xs flex flex-col justify-between gap-3">
                  <div class="space-y-1.5">
                    <div class="flex items-center justify-between gap-2">
                      <span class="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold ${
                        isSelesai
                          ? 'bg-emerald-100 text-emerald-800'
                          : (isAktif ? 'bg-sky-100 text-sky-900' : 'bg-slate-200 text-slate-700')
                      }">
                        ${isSelesai ? '✓ Sudah Selesai' : (isAktif ? '● Ujian Aktif' : '■ Ditutup')}
                      </span>
                      <span class="text-[11px] font-bold text-slate-500">Durasi: ${j.durasi_menit} Menit</span>
                    </div>
                    <h3 class="font-extrabold text-base text-slate-900">${escapeHtml(mp.nama_mapel)}</h3>
                    <div class="text-xs font-bold text-emerald-800">${escapeHtml(j.judul_ujian)}</div>
                    <div class="text-[11px] text-slate-600 pt-1">
                      <div>Mulai: <strong>${formatDateTimeWita(j.waktu_mulai)}</strong></div>
                      <div>Batas Akhir: <strong>${formatDateTimeWita(j.waktu_selesai)}</strong></div>
                    </div>
                  </div>

                  <div class="pt-2 border-t border-slate-200/80 flex items-center justify-between gap-2">
                    ${isSelesai ? `
                      <div class="text-xs font-extrabold text-emerald-800">Nilai Akhir: <span class="text-base">${sesi.nilai_akhir ?? 0}</span></div>
                      <span class="px-3 py-1.5 rounded-xl bg-emerald-50 text-emerald-700 border border-emerald-200 text-xs font-bold">Terkumpul</span>
                    ` : (isAktif ? `
                      <span class="text-[11px] font-bold text-amber-800">${isMengerjakan ? '⏳ Sesi sedang berjalan' : 'Siap dikerjakan'}</span>
                      <button type="button" data-open-token="${j.id}" class="px-4 py-2 rounded-xl bg-gradient-to-r from-[#4f6d4a] to-emerald-700 hover:opacity-95 text-white font-extrabold text-xs shadow-xs transition">
                        ${isMengerjakan ? '▶ Lanjutkan Ujian' : '▶ Kerjakan Ujian'}
                      </button>
                    ` : `
                      <span class="text-xs font-bold text-slate-400">Waktu ujian telah berakhir</span>
                    `)}
                  </div>
                </div>
              `;
            }).join('')}
          </div>
        `}
      </div>
    `;
  }

  function renderStudentNilaiView() {
    const siswa = state.currentUser;
    const mySesi = state.db.sesi_ujian_siswa.filter(s => s.siswa_id === siswa.id);

    return `
      <div class="space-y-4">
        <div class="glass-panel rounded-2xl p-4 sm:p-5">
          <h2 class="text-base sm:text-lg font-extrabold text-slate-900">Riwayat &amp; Hasil Ujian Saya</h2>
          <p class="text-xs text-slate-500">Daftar ujian yang telah Anda ikuti beserta status pengumpulan dan nilai akhir.</p>
        </div>

        ${mySesi.length === 0 ? `
          <div class="glass-panel rounded-2xl p-8 text-center text-xs text-slate-500">Belum ada riwayat pengerjaan ujian.</div>
        ` : `
          <div class="space-y-2.5">
            ${mySesi.map(s => {
              const jd = state.db.jadwal_ujian.find(j => j.id === s.jadwal_id) || { judul_ujian: 'Ujian CBT', mapel_id: '' };
              const mp = state.db.mata_pelajaran.find(m => m.id === jd.mapel_id) || { nama_mapel: '-' };
              return `
                <div class="glass-panel rounded-2xl p-4 flex items-center justify-between gap-3">
                  <div>
                    <div class="text-xs font-extrabold text-emerald-800">${escapeHtml(mp.nama_mapel)}</div>
                    <div class="font-extrabold text-sm text-slate-900">${escapeHtml(jd.judul_ujian)}</div>
                    <div class="text-[11px] text-slate-500 mt-0.5">Status: <strong>${s.status === 'selesai' ? 'Selesai Dikumpulkan' : 'Sedang Berjalan'}</strong></div>
                  </div>
                  <div class="text-right">
                    <div class="text-[10px] font-bold text-slate-500">NILAI AKHIR</div>
                    <div class="text-xl font-extrabold text-slate-900">${s.status === 'selesai' ? (s.nilai_akhir ?? 0) : '-'}</div>
                  </div>
                </div>
              `;
            }).join('')}
          </div>
        `}
      </div>
    `;
  }

  // =========================================================================
  // RUANG UJIAN FULLSCREEN SISWA (TIPOGRAFI PREMIUM & INTERAKSI 6 TIPE SOAL)
  // =========================================================================
  function isSoalAnswered(s, val) {
    const sTipe = normalizeTipeSoal(s.tipe_soal);
    if (sTipe === 'MENJODOHKAN') {
      const totalPairs = normalizePasanganJson(s.pasangan_json, s).length;
      const ansObj = normalizeMatchAnswerObj(val);
      const filledPairs = Object.values(ansObj).filter(v => String(v || '').trim() !== '').length;
      return totalPairs > 0 ? filledPairs >= totalPairs : filledPairs > 0;
    }
    if (sTipe === 'PILIHAN GANDA KOMPLEKS') {
      return normalizePgkAnswerArray(val).length > 0;
    }
    if (Array.isArray(val)) return val.length > 0;
    if (val && typeof val === 'object') return Object.values(val).some(v => String(v || '').trim() !== '');
    return val !== undefined && val !== null && String(val).trim() !== '';
  }

  function renderStudentExamView() {
    const ex = state.exam;
    const soal = ex.soalList[ex.currentIndex];
    if (!soal) return '';

    const isEnglish = Boolean(ex.mapel?.is_english);
    const noTransClass = isEnglish ? 'notranslate' : '';
    const noTransAttr = isEnglish ? 'translate="no"' : '';

    // Skala tipografi yang tertata rapi, seimbang & mudah dibaca
    const questionTextClass = ex.fontSize === 'lg'
      ? 'text-[17px] sm:text-[18.5px] leading-[1.8]'
      : (ex.fontSize === 'xl'
          ? 'text-[19px] sm:text-[21px] leading-[1.85]'
          : 'text-[15.5px] sm:text-[16.5px] leading-[1.75]');

    const optionTextClass = ex.fontSize === 'lg'
      ? 'text-[15.5px] sm:text-[17px] leading-[1.65]'
      : (ex.fontSize === 'xl'
          ? 'text-[17px] sm:text-[18.5px] leading-[1.7]'
          : 'text-[14px] sm:text-[15px] leading-[1.6]');

    const currentAns = ex.jawaban[soal.id];
    const tipe = normalizeTipeSoal(soal.tipe_soal);
    const displayOptionKeys = ex.optionOrders[soal.id] || ['A', 'B', 'C', 'D'].filter(k => (soal[`opsi_${k.toLowerCase()}`] || '').trim() !== '');
    const displayLabels = ['A', 'B', 'C', 'D'];

    const tipeBadgeStyle = {
      'PILIHAN GANDA': 'bg-emerald-50 text-emerald-800 border-emerald-200',
      'PILIHAN GANDA KOMPLEKS': 'bg-amber-50 text-amber-900 border-amber-300',
      'BENAR SALAH': 'bg-indigo-50 text-indigo-800 border-indigo-200',
      'MENJODOHKAN': 'bg-sky-50 text-sky-900 border-sky-200',
      'ISIAN SINGKAT': 'bg-teal-50 text-teal-800 border-teal-200',
      'URAIAN': 'bg-purple-50 text-purple-900 border-purple-200'
    }[tipe] || 'bg-slate-100 text-slate-700 border-slate-200';

    let answerInputHtml = '';
    if (tipe === 'PILIHAN GANDA') {
      answerInputHtml = `
        <div class="space-y-3 ${noTransClass}" ${noTransAttr}>
          <div class="flex items-center gap-2 text-xs font-bold text-slate-500">
            <span class="w-1.5 h-1.5 rounded-full bg-[#4f6d4a]"></span>
            <span>Pilih salah satu jawaban yang paling tepat:</span>
          </div>
          <div class="space-y-2.5">
            ${displayOptionKeys.map((origKey, dispIdx) => {
              const optText = soal[`opsi_${origKey.toLowerCase()}`];
              const dispLetter = displayLabels[dispIdx];
              const isSelected = String(currentAns || '').trim().toUpperCase() === origKey;
              return `
                <button type="button" data-answer-pg="${origKey}"
                  class="group w-full p-4 sm:p-4.5 rounded-2xl border-2 text-left flex items-start gap-3.5 transition-all duration-150 ${
                    isSelected
                      ? 'border-[#4f6d4a] bg-emerald-50/90 text-slate-950 font-semibold shadow-sm'
                      : 'border-slate-200/90 bg-white hover:border-[#4f6d4a]/50 hover:bg-slate-50/70 text-slate-800'
                  }">
                  <span class="w-8 h-8 rounded-full flex items-center justify-center text-xs font-extrabold shrink-0 mt-0.5 transition-colors ${
                    isSelected
                      ? 'bg-[#4f6d4a] text-white shadow-xs'
                      : 'bg-slate-100 text-slate-600 group-hover:bg-slate-200'
                  }">${dispLetter}</span>
                  <span class="${optionTextClass} flex-1 break-words tracking-[-0.005em]">${escapeHtml(optText)}</span>
                </button>
              `;
            }).join('')}
          </div>
        </div>
      `;
    } else if (tipe === 'BENAR SALAH') {
      let curBs = String(currentAns || '').trim().toUpperCase();
      if (curBs === 'A') curBs = 'BENAR';
      if (curBs === 'B') curBs = 'SALAH';
      answerInputHtml = `
        <div class="space-y-3 ${noTransClass}" ${noTransAttr}>
          <div class="flex items-center gap-2 text-xs font-bold text-slate-500">
            <span class="w-1.5 h-1.5 rounded-full bg-indigo-500"></span>
            <span>Tentukan apakah pernyataan di atas Benar atau Salah:</span>
          </div>
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
            ${[
              { val: 'BENAR', label: 'BENAR', icon: '✓', activeCls: 'border-emerald-600 bg-emerald-50/95 text-emerald-950 shadow-sm', badgeCls: 'bg-emerald-600 text-white' },
              { val: 'SALAH', label: 'SALAH', icon: '✕', activeCls: 'border-rose-600 bg-rose-50/95 text-rose-950 shadow-sm', badgeCls: 'bg-rose-600 text-white' }
            ].map(item => {
              const isSelected = curBs === item.val;
              return `
                <button type="button" data-answer-bs="${item.val}"
                  class="p-4 sm:p-5 rounded-2xl border-2 text-left flex items-center gap-3.5 transition-all duration-150 ${
                    isSelected
                      ? item.activeCls
                      : 'border-slate-200 bg-white hover:border-slate-300 hover:bg-slate-50 text-slate-800'
                  }">
                  <span class="w-9 h-9 rounded-xl flex items-center justify-center text-sm font-extrabold shrink-0 ${
                    isSelected ? item.badgeCls : 'bg-slate-100 text-slate-600'
                  }">${item.icon}</span>
                  <span class="text-sm sm:text-base font-extrabold tracking-wide">${item.label}</span>
                </button>
              `;
            }).join('')}
          </div>
        </div>
      `;
    } else if (tipe === 'PILIHAN GANDA KOMPLEKS') {
      const selectedArr = normalizePgkAnswerArray(currentAns);
      answerInputHtml = `
        <div class="space-y-3 ${noTransClass}" ${noTransAttr}>
          <div class="px-4 py-2.5 rounded-2xl bg-amber-50 border border-amber-200/90 flex items-center justify-between gap-2">
            <div class="flex items-center gap-2 text-xs font-bold text-amber-950">
              <span class="text-sm">☑</span>
              <span>Pilih lebih dari satu jawaban yang benar (centang semua opsi yang sesuai):</span>
            </div>
            <span class="px-2.5 py-0.5 rounded-full bg-amber-200/80 text-amber-950 text-[11px] font-extrabold shrink-0">
              ${selectedArr.length} Dipilih
            </span>
          </div>
          <div class="space-y-2.5">
            ${displayOptionKeys.map((origKey, dispIdx) => {
              const optText = soal[`opsi_${origKey.toLowerCase()}`];
              const dispLetter = displayLabels[dispIdx];
              const isChecked = selectedArr.includes(origKey);
              return `
                <button type="button" data-answer-pgk="${origKey}"
                  class="group w-full p-4 sm:p-4.5 rounded-2xl border-2 text-left flex items-start gap-3.5 transition-all duration-150 ${
                    isChecked
                      ? 'border-[#4f6d4a] bg-emerald-50/90 text-slate-950 font-semibold shadow-sm'
                      : 'border-slate-200/90 bg-white hover:border-[#4f6d4a]/50 hover:bg-slate-50/70 text-slate-800'
                  }">
                  <span class="w-8 h-8 rounded-xl flex items-center justify-center text-xs font-extrabold shrink-0 mt-0.5 border-2 transition-colors ${
                    isChecked
                      ? 'bg-[#4f6d4a] border-[#4f6d4a] text-white shadow-xs'
                      : 'bg-white border-slate-300 text-slate-600 group-hover:border-slate-400'
                  }">${isChecked ? '✓' : dispLetter}</span>
                  <span class="${optionTextClass} flex-1 break-words tracking-[-0.005em]">${escapeHtml(optText)}</span>
                </button>
              `;
            }).join('')}
          </div>
        </div>
      `;
    } else if (tipe === 'MENJODOHKAN') {
      const pairs = normalizePasanganJson(soal.pasangan_json, soal);
      const leftOrder = Array.isArray(ex.optionOrders[`${soal.id}_left`]) && ex.optionOrders[`${soal.id}_left`].length === pairs.length
        ? ex.optionOrders[`${soal.id}_left`]
        : pairs.map((_, idx) => idx);
      const rightOptions = Array.isArray(ex.optionOrders[`${soal.id}_right`]) && ex.optionOrders[`${soal.id}_right`].length === pairs.length
        ? ex.optionOrders[`${soal.id}_right`]
        : pairs.map(p => p.kanan);
      const ansObj = normalizeMatchAnswerObj(currentAns);
      const matchedCount = pairs.filter((_, idx) => String(ansObj[idx] || ansObj[String(idx)] || '').trim() !== '').length;

      answerInputHtml = `
        <div class="space-y-3.5 ${noTransClass}" ${noTransAttr}>
          <div class="px-4 py-2.5 rounded-2xl bg-sky-50 border border-sky-200 flex items-center justify-between gap-2">
            <div class="flex items-center gap-2 text-xs font-bold text-sky-950">
              <span>🔗</span>
              <span>Pasangkan setiap pernyataan di kolom kiri dengan jawaban yang sesuai di kolom kanan:</span>
            </div>
            <span class="px-2.5 py-0.5 rounded-full bg-sky-200/80 text-sky-950 text-[11px] font-extrabold shrink-0">
              ${matchedCount}/${pairs.length} Terpasang
            </span>
          </div>

          ${pairs.length === 0 ? `
            <div class="p-4 rounded-2xl bg-rose-50 border border-rose-200 text-xs text-rose-700 font-bold">
              Data pasangan menjodohkan belum tersedia pada butir soal ini.
            </div>
          ` : `
            <div class="space-y-3">
              <div class="hidden md:grid md:grid-cols-12 gap-3 px-3 text-[11px] font-extrabold uppercase tracking-wider text-slate-400">
                <div class="col-span-6">Kolom Pernyataan / Premis</div>
                <div class="col-span-1 text-center">Hubung</div>
                <div class="col-span-5">Kolom Pilihan Pasangan</div>
              </div>
              ${leftOrder.map((origIdx, dispIdx) => {
                const p = pairs[origIdx];
                if (!p) return '';
                const chosenVal = String(ansObj[origIdx] ?? ansObj[String(origIdx)] ?? '');
                const isMatched = Boolean(chosenVal.trim());
                return `
                  <div class="p-3.5 sm:p-4 rounded-2xl border-2 transition-all ${
                    isMatched
                      ? 'bg-emerald-50/70 border-emerald-400 shadow-xs'
                      : 'bg-slate-50/90 border-slate-200'
                  } grid grid-cols-1 md:grid-cols-12 gap-3 md:items-center">
                    <div class="md:col-span-6 flex items-start gap-3">
                      <span class="w-7 h-7 rounded-xl ${isMatched ? 'bg-[#4f6d4a]' : 'bg-slate-800'} text-white text-xs font-extrabold flex items-center justify-center shrink-0 mt-0.5">${dispIdx + 1}</span>
                      <span class="${optionTextClass} font-semibold text-slate-900 leading-relaxed break-words">${escapeHtml(p.kiri)}</span>
                    </div>
                    <div class="hidden md:flex md:col-span-1 justify-center text-slate-400 font-extrabold text-base">➔</div>
                    <div class="md:col-span-5">
                      <label class="block md:hidden text-[11px] font-extrabold text-slate-500 mb-1">Pilih Pasangan Jawaban:</label>
                      <select data-answer-match="${origIdx}"
                        class="w-full px-3.5 py-3 rounded-xl border-2 ${
                          isMatched
                            ? 'border-[#4f6d4a] bg-white text-slate-950 font-bold'
                            : 'border-slate-300 bg-white text-slate-700 font-semibold'
                        } text-xs sm:text-sm focus:border-[#4f6d4a] focus:outline-none">
                        <option value="">-- Pilih Pasangan Jawaban --</option>
                        ${rightOptions.map(opt => `<option value="${escapeHtml(opt)}" ${chosenVal === opt ? 'selected' : ''}>${escapeHtml(opt)}</option>`).join('')}
                      </select>
                    </div>
                  </div>
                `;
              }).join('')}
            </div>
          `}
        </div>
      `;
    } else if (tipe === 'ISIAN SINGKAT') {
      answerInputHtml = `
        <div class="p-4 sm:p-5 rounded-2xl bg-teal-50/50 border-2 border-teal-200/80 space-y-2.5 ${noTransClass}" ${noTransAttr}>
          <label for="input-isian-singkat" class="flex items-center gap-2 text-xs sm:text-sm font-extrabold text-teal-950">
            <span>✏️</span>
            <span>Kolom Jawaban Singkat (Ketik jawaban Anda di bawah ini):</span>
          </label>
          <input id="input-isian-singkat" type="text" value="${escapeHtml(currentAns || '')}" placeholder="Ketik jawaban singkat dengan ejaan yang benar..."
            class="w-full px-4 py-3.5 rounded-xl border-2 border-slate-300 bg-white ${optionTextClass} font-bold text-slate-950 focus:border-[#4f6d4a] focus:outline-none shadow-2xs" />
          <div class="text-[11px] text-slate-500 font-medium">Jawaban otomatis tersimpan saat Anda mengetik.</div>
        </div>
      `;
    } else {
      const charCount = String(currentAns || '').length;
      answerInputHtml = `
        <div class="p-4 sm:p-5 rounded-2xl bg-purple-50/40 border-2 border-purple-200/80 space-y-2.5 ${noTransClass}" ${noTransAttr}>
          <div class="flex items-center justify-between gap-2">
            <label for="input-uraian-jawab" class="flex items-center gap-2 text-xs sm:text-sm font-extrabold text-purple-950">
              <span>📝</span>
              <span>Kolom Jawaban Uraian / Esai:</span>
            </label>
            <span id="uraian-char-count" class="text-[11px] font-bold text-slate-500">${charCount} karakter</span>
          </div>
          <textarea id="input-uraian-jawab" rows="6" placeholder="Tuliskan penjelasan jawaban Anda secara lengkap, jelas, dan terstruktur pada kolom ini..."
            class="w-full p-4 rounded-xl border-2 border-slate-300 bg-white ${optionTextClass} text-slate-950 leading-relaxed focus:border-[#4f6d4a] focus:outline-none shadow-2xs">${escapeHtml(currentAns || '')}</textarea>
          <div class="text-[11px] text-slate-500 font-medium">Jawaban uraian otomatis tersimpan dan akan dikoreksi oleh Guru Mata Pelajaran.</div>
        </div>
      `;
    }

    const totalSoal = ex.soalList.length;
    const answeredCount = ex.soalList.filter(s => isSoalAnswered(s, ex.jawaban[s.id])).length;
    const isRagu = Boolean(ex.ragu[soal.id]);

    const navGridHtml = ex.soalList.map((s, idx) => {
      const answered = isSoalAnswered(s, ex.jawaban[s.id]);
      const ragu = Boolean(ex.ragu[s.id]);
      const isCurrent = idx === ex.currentIndex;
      let cls = 'bg-slate-100 text-slate-700 border-slate-300';
      if (ragu) cls = 'bg-amber-400 text-slate-950 border-amber-500';
      else if (answered) cls = 'bg-[#4f6d4a] text-white border-[#3f593b]';
      if (isCurrent) cls += ' ring-2 ring-offset-1 ring-slate-900 scale-105';
      return `<button type="button" data-jump-soal="${idx}" class="h-10 rounded-xl border font-extrabold text-xs transition ${cls}">${idx + 1}</button>`;
    }).join('');

    const watermarkItems = Array.from({ length: 14 })
      .map(() => `<span class="text-xs font-extrabold text-slate-900 tracking-widest p-6">${escapeHtml(state.currentUser.nama_lengkap)} • ${escapeHtml(state.currentUser.nomor_peserta)}</span>`)
      .join('');

    return `
      <div class="exam-protected min-h-[100dvh] flex flex-col justify-between bg-[#f4f6f3] ${noTransClass}" ${noTransAttr}>
        <div class="exam-watermark-layer">${watermarkItems}</div>

        <!-- Sticky Top Exam Bar -->
        <header class="sticky top-0 z-30 bg-white/95 backdrop-blur-md border-b border-slate-200/90 px-3.5 sm:px-6 py-3 flex items-center justify-between gap-2 shadow-2xs">
          <div class="min-w-0">
            <div class="text-xs sm:text-sm font-extrabold text-slate-900 truncate">${escapeHtml(ex.mapel?.nama_mapel || 'Ujian CBT')}</div>
            <div class="text-[11px] font-medium text-slate-500 truncate">${escapeHtml(state.currentUser.nama_lengkap)} (${escapeHtml(state.currentUser.kelas)})</div>
          </div>

          <div class="flex items-center gap-1.5 sm:gap-2.5 shrink-0">
            <div class="px-3 py-1.5 rounded-xl bg-slate-900 text-white font-mono text-xs sm:text-sm font-extrabold tracking-wider shadow-xs">
              ⏱ <span id="exam-countdown-timer">${formatRemainingTime(ex.remainingSeconds)}</span>
            </div>

            <button id="btn-toggle-daftar-soal" type="button" class="lg:hidden px-3 py-1.5 rounded-xl bg-emerald-50 text-emerald-900 border border-emerald-200 text-xs font-extrabold">
              Daftar (${answeredCount}/${totalSoal})
            </button>
          </div>
        </header>

        <!-- Main Exam Body -->
        <div class="flex-1 max-w-6xl w-full mx-auto px-3.5 sm:px-6 py-4 sm:py-6 grid grid-cols-1 lg:grid-cols-12 gap-5 items-start">
          <!-- Kolom Soal Utama -->
          <div class="lg:col-span-8 glass-panel rounded-3xl p-5 sm:p-7 space-y-5 shadow-md">
            <div class="flex flex-col gap-2.5 pb-3.5 border-b border-slate-200/80">
              <div class="flex flex-wrap items-center justify-between gap-2">
                <div class="flex flex-wrap items-center gap-2">
                  <span class="px-3.5 py-1.5 rounded-xl bg-[#4f6d4a] text-white text-xs sm:text-sm font-extrabold tracking-wide shadow-2xs">
                    SOAL NO. ${ex.currentIndex + 1} <span class="font-normal opacity-80">/ ${totalSoal}</span>
                  </span>
                  <span class="px-3 py-1 rounded-xl border text-[11px] font-extrabold tracking-wide ${tipeBadgeStyle}">
                    ${escapeHtml(tipe)}
                  </span>
                </div>
                <span class="text-xs font-bold text-slate-400">
                  Terjawab: <strong class="text-slate-700">${answeredCount}</strong> dari ${totalSoal}
                </span>
              </div>

              <!-- Kontrol Ukuran Huruf (Font Size) di bawah Tipe Soal -->
              <div class="flex items-center justify-between gap-2 pt-0.5">
                <div class="inline-flex items-center gap-1.5 bg-slate-100/90 px-2.5 py-1 rounded-xl border border-slate-200/80">
                  <span class="text-[11px] font-bold text-slate-500">Ukuran Teks:</span>
                  <button type="button" data-exam-font="base" aria-label="Ukuran Teks Normal"
                    class="px-2.5 py-1 rounded-lg text-[11px] font-extrabold transition ${ex.fontSize === 'base' ? 'bg-[#4f6d4a] text-white shadow-2xs' : 'bg-white/70 text-slate-600 hover:bg-white'}">
                    A
                  </button>
                  <button type="button" data-exam-font="lg" aria-label="Ukuran Teks Sedang"
                    class="px-2.5 py-1 rounded-lg text-xs font-extrabold transition ${ex.fontSize === 'lg' ? 'bg-[#4f6d4a] text-white shadow-2xs' : 'bg-white/70 text-slate-600 hover:bg-white'}">
                    A+
                  </button>
                  <button type="button" data-exam-font="xl" aria-label="Ukuran Teks Besar"
                    class="px-2.5 py-1 rounded-lg text-sm font-extrabold transition ${ex.fontSize === 'xl' ? 'bg-[#4f6d4a] text-white shadow-2xs' : 'bg-white/70 text-slate-600 hover:bg-white'}">
                    A++
                  </button>
                </div>
              </div>
            </div>

            <!-- Box Teks Pertanyaan Tertata Rapi -->
            <div class="exam-question-read p-4 sm:p-5 rounded-2xl bg-slate-50/90 border border-slate-200/80 space-y-4">
              <div class="${questionTextClass} text-slate-900 font-medium whitespace-pre-line break-words tracking-[-0.005em]">${escapeHtml(soal.pertanyaan)}</div>
              ${soal.gambar_url ? `
                <div class="pt-1 flex justify-center">
                  <img src="${soal.gambar_url}" alt="Stimulus Gambar Soal" class="max-h-72 rounded-2xl border border-slate-200 bg-white p-1.5 object-contain shadow-2xs" />
                </div>
              ` : ''}
            </div>

            <!-- Area Jawaban Sesuai Tipe Soal -->
            <div class="pt-0.5">
              ${answerInputHtml}
            </div>
          </div>

          <!-- Navigasi Nomor Soal (Desktop Sidebar & Mobile Drawer) -->
          <div class="lg:col-span-4 ${ex.showDaftarSoalMobile ? 'block' : 'hidden lg:block'}">
            <div class="glass-panel rounded-3xl p-5 space-y-4 shadow-md sticky top-20">
              <div class="flex items-center justify-between">
                <span class="font-extrabold text-sm text-slate-900">Navigasi Nomor Soal</span>
                <span class="px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200 text-xs font-extrabold">${answeredCount}/${totalSoal} Terjawab</span>
              </div>
              <div class="grid grid-cols-5 gap-2">
                ${navGridHtml}
              </div>
              <div class="flex flex-wrap gap-3 text-[11px] font-bold text-slate-600 pt-3 border-t border-slate-200">
                <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-[#4f6d4a] inline-block"></span> Dijawab</span>
                <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-amber-400 inline-block"></span> Ragu-ragu</span>
                <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-slate-200 border border-slate-300 inline-block"></span> Kosong</span>
              </div>
            </div>
          </div>
        </div>

        <!-- Bottom Action Bar -->
        <footer class="sticky bottom-0 z-30 bg-white/95 backdrop-blur-md border-t border-slate-200/90 px-3.5 sm:px-6 py-3 shadow-sm">
          <div class="max-w-6xl mx-auto flex items-center justify-between gap-2">
            <button id="btn-prev-soal" type="button" ${ex.currentIndex === 0 ? 'disabled' : ''}
              class="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-800 font-extrabold text-xs disabled:opacity-40 transition">
              ← Sebelumnya
            </button>

            <label class="px-3.5 py-2.5 rounded-xl bg-amber-50 hover:bg-amber-100 border border-amber-300 text-amber-950 font-extrabold text-xs flex items-center gap-2 cursor-pointer transition">
              <input id="chk-ragu-soal" type="checkbox" ${isRagu ? 'checked' : ''} class="rounded border-amber-500 text-amber-600 focus:ring-0" />
              <span>Ragu-ragu</span>
            </label>

            ${ex.currentIndex < totalSoal - 1 ? `
              <button id="btn-next-soal" type="button"
                class="px-4 py-2.5 rounded-xl bg-[#4f6d4a] hover:bg-[#3f593b] text-white font-extrabold text-xs shadow-xs transition">
                Berikutnya →
              </button>
            ` : `
              <button id="btn-finish-exam" type="button"
                class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-xs shadow-sm transition">
                ✓ Selesai &amp; Kumpulkan
              </button>
            `}
          </div>
        </footer>
      </div>
    `;
  }

  // =========================================================================
  // MODAL CRUD (BANK SOAL LENGKAP 6 TIPE, IMPORT WORD, MAPEL, USER, JADWAL, TOKEN)
  // =========================================================================
  function renderModalOverlay() {
    if (!state.modal) return '';
    const m = state.modal;
    let title = '';
    let body = '';

    if (m.type === 'soal') {
      const existing = m.data;
      title = existing ? `Edit Butir Soal (No. ${existing.nomor_urut})` : 'Tambah Butir Soal Baru';
      const authMapels = getAuthorizedMapels();
      const defaultMapelId = existing ? existing.mapel_id : (state.filterSoalMapel || authMapels[0]?.id || '');

      if (!state.modalSoalDraft) {
        state.modalSoalDraft = buildInitialSoalDraft(existing, defaultMapelId);
      }
      const d = state.modalSoalDraft;
      const curTipe = normalizeTipeSoal(d.tipe_soal);
      const authKelas = getAuthorizedKelasList(d.mapel_id || defaultMapelId);
      const selectedKelas = state.modalSoalSelectedKelas.size > 0
        ? state.modalSoalSelectedKelas
        : new Set(existing ? parseKelasString(existing.kelas_target) : authKelas);
      const imgPrev = state.modalSoalImagePreview || (existing ? existing.gambar_url : '');

      let dynamicAnswerEditorHtml = '';
      if (curTipe === 'PILIHAN GANDA') {
        const curKey = String(d.kunci_jawaban || 'A').trim().toUpperCase();
        dynamicAnswerEditorHtml = `
          <div class="p-3.5 rounded-2xl bg-slate-50 border border-slate-200 space-y-2.5">
            <div class="flex items-center justify-between">
              <span class="font-extrabold text-slate-800">Opsi Pilihan Ganda &amp; Kunci Jawaban (Pilih 1)</span>
              <span class="text-[11px] font-bold text-emerald-700">Kunci Terpilih: <strong>${escapeHtml(curKey)}</strong></span>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
              ${['A', 'B', 'C', 'D'].map(opt => `
                <div class="p-2.5 rounded-xl bg-white border ${curKey === opt ? 'border-emerald-500 ring-1 ring-emerald-400' : 'border-slate-300'} space-y-1.5">
                  <div class="flex items-center justify-between">
                    <span class="font-extrabold text-xs text-slate-800">Opsi ${opt}</span>
                    <label class="inline-flex items-center gap-1.5 text-[11px] font-extrabold text-emerald-700 cursor-pointer">
                      <input type="radio" name="kunci_pg_radio" value="${opt}" ${curKey === opt ? 'checked' : ''} />
                      <span>Kunci ${opt}</span>
                    </label>
                  </div>
                  <input name="opsi_${opt.toLowerCase()}" value="${escapeHtml(d[`opsi_${opt.toLowerCase()}`] || '')}" placeholder="Teks jawaban opsi ${opt}..."
                    class="w-full px-3 py-2 rounded-lg border border-slate-200 text-xs focus:border-[#4f6d4a] focus:outline-none" />
                </div>
              `).join('')}
            </div>
          </div>
        `;
      } else if (curTipe === 'PILIHAN GANDA KOMPLEKS') {
        const selectedKeys = new Set(String(d.kunci_jawaban || '').split(',').map(x => x.trim().toUpperCase()).filter(Boolean));
        dynamicAnswerEditorHtml = `
          <div class="p-3.5 rounded-2xl bg-amber-50/60 border border-amber-200 space-y-2.5">
            <div class="flex items-center justify-between">
              <span class="font-extrabold text-amber-950">Opsi Pilihan Ganda Kompleks (Centang Lebih Dari 1 Kunci)</span>
              <span class="text-[11px] font-bold text-amber-800">Kunci: <strong>${escapeHtml([...selectedKeys].join(', ') || 'Belum dipilih')}</strong></span>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
              ${['A', 'B', 'C', 'D'].map(opt => {
                const isChecked = selectedKeys.has(opt);
                return `
                  <div class="p-2.5 rounded-xl bg-white border ${isChecked ? 'border-emerald-500 ring-1 ring-emerald-400' : 'border-slate-300'} space-y-1.5">
                    <div class="flex items-center justify-between">
                      <span class="font-extrabold text-xs text-slate-800">Opsi ${opt}</span>
                      <label class="inline-flex items-center gap-1.5 text-[11px] font-extrabold text-emerald-700 cursor-pointer">
                        <input type="checkbox" value="${opt}" class="chk-kunci-pgk rounded border-slate-300" ${isChecked ? 'checked' : ''} />
                        <span>Kunci Benar</span>
                      </label>
                    </div>
                    <input name="opsi_${opt.toLowerCase()}" value="${escapeHtml(d[`opsi_${opt.toLowerCase()}`] || '')}" placeholder="Teks pernyataan opsi ${opt}..."
                      class="w-full px-3 py-2 rounded-lg border border-slate-200 text-xs focus:border-[#4f6d4a] focus:outline-none" />
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        `;
      } else if (curTipe === 'BENAR SALAH') {
        const curBs = String(d.kunci_jawaban || 'BENAR').trim().toUpperCase() === 'SALAH' ? 'SALAH' : 'BENAR';
        dynamicAnswerEditorHtml = `
          <div class="p-3.5 rounded-2xl bg-indigo-50/60 border border-indigo-200 space-y-2">
            <div class="font-extrabold text-indigo-950">Tentukan Kunci Jawaban Pernyataan (Benar / Salah)</div>
            <div class="grid grid-cols-2 gap-2.5">
              ${['BENAR', 'SALAH'].map(opt => `
                <label class="p-2.5 rounded-xl bg-white border ${curBs === opt ? 'border-emerald-500 ring-1 ring-emerald-400 font-extrabold text-emerald-900' : 'border-slate-300 font-bold text-slate-700'} flex items-center gap-2 cursor-pointer">
                  <input type="radio" name="kunci_bs_radio" value="${opt}" ${curBs === opt ? 'checked' : ''} />
                  <span>${opt}</span>
                </label>
              `).join('')}
            </div>
          </div>
        `;
      } else if (curTipe === 'MENJODOHKAN') {
        const pairs = Array.isArray(d.pasangan) && d.pasangan.length > 0 ? d.pasangan : [{ kiri: '', kanan: '' }];
        dynamicAnswerEditorHtml = `
          <div class="p-3.5 rounded-2xl bg-sky-50/70 border border-sky-200 space-y-2.5">
            <div class="flex items-center justify-between">
              <span class="font-extrabold text-sky-950">Daftar Pasangan Menjodohkan (Premis Kiri ➔ Pasangan Kanan)</span>
              <button id="btn-add-pasangan-row" type="button" class="px-2.5 py-1 rounded-lg bg-sky-700 text-white font-bold text-[11px]">+ Tambah Baris</button>
            </div>
            <div class="space-y-2">
              ${pairs.map((p, idx) => `
                <div class="grid grid-cols-12 gap-2 items-center bg-white p-2 rounded-xl border border-slate-200">
                  <div class="col-span-5">
                    <input data-pasangan-kiri="${idx}" value="${escapeHtml(p.kiri)}" placeholder="Premis Kiri ${idx + 1}..."
                      class="w-full px-2.5 py-2 rounded-lg border border-slate-300 text-xs font-semibold" />
                  </div>
                  <div class="col-span-1 text-center font-extrabold text-slate-400">➔</div>
                  <div class="col-span-5">
                    <input data-pasangan-kanan="${idx}" value="${escapeHtml(p.kanan)}" placeholder="Pasangan Jawaban ${idx + 1}..."
                      class="w-full px-2.5 py-2 rounded-lg border border-slate-300 text-xs font-semibold text-emerald-900" />
                  </div>
                  <div class="col-span-1 text-right">
                    ${pairs.length > 1 ? `<button type="button" data-del-pasangan-row="${idx}" class="px-2 py-1.5 rounded-lg bg-rose-50 text-rose-700 font-bold text-xs">✕</button>` : ''}
                  </div>
                </div>
              `).join('')}
            </div>
            <div class="text-[11px] text-slate-600">Sistem akan otomatis mengacak posisi pilihan pasangan di layar ujian siswa dan menilai berdasarkan kecocokan pasangan di atas.</div>
          </div>
        `;
      } else if (curTipe === 'ISIAN SINGKAT') {
        dynamicAnswerEditorHtml = `
          <div class="p-3.5 rounded-2xl bg-teal-50/70 border border-teal-200 space-y-2">
            <label class="block font-extrabold text-teal-950">Kunci Jawaban Isian Singkat</label>
            <input name="kunci_jawaban" value="${escapeHtml(d.kunci_jawaban || '')}" required placeholder="Contoh: Samarinda (Gunakan tanda | jika ada variasi jawaban, misal: 17 Agustus 1945 | 17-08-1945)"
              class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 bg-white font-bold text-slate-900" />
            <div class="text-[11px] text-slate-600">Pencocokan otomatis bersifat tidak peka huruf besar/kecil (case-insensitive).</div>
          </div>
        `;
      } else {
        dynamicAnswerEditorHtml = `
          <div class="p-3.5 rounded-2xl bg-purple-50/70 border border-purple-200 space-y-2">
            <label class="block font-extrabold text-purple-950">Pedoman Penskoran / Kunci Jawaban Uraian (Esai)</label>
            <textarea name="kunci_jawaban" rows="3" placeholder="Tuliskan poin-poin jawaban benar sebagai pedoman guru saat mengoreksi jawaban uraian siswa..."
              class="w-full p-3 rounded-xl border border-slate-300 bg-white text-slate-900">${escapeHtml(d.kunci_jawaban || '')}</textarea>
          </div>
        `;
      }

      body = `
        <form id="form-save-soal" class="space-y-3.5 text-xs">
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label class="block font-bold mb-1">Mata Pelajaran</label>
              <select id="modal-soal-mapel" name="mapel_id" required class="w-full px-3 py-2 rounded-xl border border-slate-300 bg-white font-semibold">
                ${authMapels.map(mp => `<option value="${mp.id}" ${d.mapel_id === mp.id ? 'selected' : ''}>${escapeHtml(mp.nama_mapel)}</option>`).join('')}
              </select>
            </div>
            <div>
              <label class="block font-bold mb-1">Tipe Butir Soal (6 Bentuk Soal)</label>
              <select id="modal-soal-tipe" name="tipe_soal" class="w-full px-3 py-2 rounded-xl border-2 border-[#4f6d4a]/40 bg-emerald-50/40 font-extrabold text-slate-900">
                ${['PILIHAN GANDA', 'PILIHAN GANDA KOMPLEKS', 'BENAR SALAH', 'MENJODOHKAN', 'ISIAN SINGKAT', 'URAIAN'].map(t => `<option value="${t}" ${curTipe === t ? 'selected' : ''}>${t}</option>`).join('')}
              </select>
            </div>
          </div>

          <div>
            <label class="block font-bold mb-1">Target Kelas Soal</label>
            <div class="flex flex-wrap gap-2 p-2.5 rounded-xl bg-slate-50 border border-slate-200">
              ${authKelas.map(k => `
                <label class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white border border-slate-200 cursor-pointer font-bold">
                  <input type="checkbox" value="${k}" class="chk-soal-kelas" ${selectedKelas.has(k) ? 'checked' : ''} />
                  <span>${k}</span>
                </label>
              `).join('')}
            </div>
          </div>

          <div class="grid grid-cols-2 gap-3">
            <div>
              <label class="block font-bold mb-1">Nomor Urut</label>
              <input type="number" name="nomor_urut" value="${d.nomor_urut}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-bold" />
            </div>
            <div>
              <label class="block font-bold mb-1">Bobot Nilai</label>
              <input type="number" name="bobot" value="${d.bobot}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-bold" />
            </div>
          </div>

          <div>
            <label class="block font-bold mb-1">Pertanyaan / Stimulus Soal</label>
            <textarea name="pertanyaan" rows="3" required placeholder="Tuliskan pertanyaan atau stimulus soal secara jelas..." class="w-full p-3 rounded-xl border border-slate-300 leading-relaxed">${escapeHtml(d.pertanyaan)}</textarea>
          </div>

          <div>
            <label class="block font-bold mb-1">Gambar Stimulus Soal (Opsional)</label>
            <div class="flex flex-wrap items-center gap-3">
              <input id="input-soal-image" type="file" accept="image/*" class="text-xs" />
              ${imgPrev ? `
                <img src="${imgPrev}" class="h-16 rounded-lg border border-slate-300 object-contain" />
                <button id="btn-clear-soal-img" type="button" class="px-2.5 py-1 rounded-lg bg-rose-50 text-rose-700 font-bold">Hapus Gambar</button>
              ` : ''}
            </div>
          </div>

          ${dynamicAnswerEditorHtml}

          <div class="flex justify-end gap-2 pt-2 border-t border-slate-200">
            <button type="button" data-close-modal class="px-4 py-2 rounded-xl bg-slate-100 font-bold">Batal</button>
            <button type="submit" class="px-4 py-2 rounded-xl bg-[#4f6d4a] text-white font-extrabold">💾 Simpan Butir Soal</button>
          </div>
        </form>
      `;
    } else if (m.type === 'import_word') {
      title = 'Import Massal Butir Soal dari Word (.docx)';
      const authMapels = getAuthorizedMapels();
      const targetMapelId = state.modalImportMapelId || state.filterSoalMapel || authMapels[0]?.id || '';
      const authKelas = getAuthorizedKelasList(targetMapelId);
      const selectedKelas = state.modalImportSelectedKelas.size > 0 ? state.modalImportSelectedKelas : new Set(authKelas);

      body = `
        <form id="form-import-word-soal" class="space-y-3.5 text-xs">
          <div>
            <label class="block font-bold mb-1">Pilih Mata Pelajaran Tujuan</label>
            <select id="import-word-mapel" name="mapel_id" required class="w-full px-3 py-2.5 rounded-xl border border-slate-300 bg-white font-bold">
              ${authMapels.map(mp => `<option value="${mp.id}" ${targetMapelId === mp.id ? 'selected' : ''}>${escapeHtml(mp.nama_mapel)}</option>`).join('')}
            </select>
          </div>
          <div>
            <label class="block font-bold mb-1">Target Kelas untuk Soal yang Diimpor</label>
            <div class="flex flex-wrap gap-2 p-2.5 rounded-xl bg-slate-50 border border-slate-200">
              ${authKelas.map(k => `
                <label class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white border border-slate-200 cursor-pointer font-bold">
                  <input type="checkbox" value="${k}" class="chk-import-kelas" ${selectedKelas.has(k) ? 'checked' : ''} />
                  <span>${k}</span>
                </label>
              `).join('')}
            </div>
          </div>
          <div>
            <label class="block font-bold mb-1">Pilih File Microsoft Word (.docx)</label>
            <input id="input-file-word-docx" type="file" accept=".docx" required class="w-full p-2.5 rounded-xl border border-slate-300 bg-slate-50" />
            <p class="text-[11px] text-slate-500 mt-1">Mendukung teks soal, 6 tipe soal, dan gambar tertanam di dalam dokumen Word.</p>
          </div>
          <div class="flex justify-end gap-2 pt-2">
            <button type="button" data-close-modal class="px-4 py-2 rounded-xl bg-slate-100 font-bold">Batal</button>
            <button type="submit" class="px-4 py-2 rounded-xl bg-[#4f6d4a] text-white font-extrabold">📥 Proses &amp; Simpan ke Bank Soal</button>
          </div>
        </form>
      `;
    } else if (m.type === 'mapel') {
      const existing = m.data;
      title = existing ? 'Edit Mata Pelajaran' : 'Tambah Mata Pelajaran';
      const gurus = state.db.users_cbt.filter(u => u.role === 'guru');
      const allClasses = getAllMasterClasses();
      const selectedKelas = state.modalMapelSelectedKelas.size > 0
        ? state.modalMapelSelectedKelas
        : new Set(existing ? parseKelasString(existing.kelas_ampu) : []);

      body = `
        <form id="form-save-mapel" class="space-y-3 text-xs">
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label class="block font-bold mb-1">Kode Mapel</label>
              <input name="kode_mapel" value="${escapeHtml(existing?.kode_mapel || '')}" required placeholder="Contoh: IPS-07" class="w-full px-3 py-2 rounded-xl border border-slate-300 font-mono font-bold" />
            </div>
            <div>
              <label class="block font-bold mb-1">Nama Mata Pelajaran</label>
              <input name="nama_mapel" value="${escapeHtml(existing?.nama_mapel || '')}" required placeholder="Contoh: Ilmu Pengetahuan Sosial" class="w-full px-3 py-2 rounded-xl border border-slate-300 font-semibold" />
            </div>
          </div>

          <div>
            <label class="block font-bold mb-1">Guru Pengampu</label>
            <select name="guru_id" class="w-full px-3 py-2 rounded-xl border border-slate-300 bg-white font-semibold">
              <option value="">-- Pilih Guru Pengampu --</option>
              ${gurus.map(g => `<option value="${g.id}" ${(existing?.guru_id || (state.currentUser.role === 'guru' ? state.currentUser.id : '')) === g.id ? 'selected' : ''}>${escapeHtml(g.nama_lengkap)}</option>`).join('')}
            </select>
          </div>

          <div>
            <label class="block font-bold mb-1">Multi-Kelas yang Diampu Guru Ini</label>
            <div class="flex flex-wrap gap-2 p-2.5 rounded-xl bg-slate-50 border border-slate-200">
              ${allClasses.map(k => `
                <label class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white border border-slate-200 cursor-pointer font-bold">
                  <input type="checkbox" value="${k}" class="chk-mapel-kelas" ${selectedKelas.has(k) ? 'checked' : ''} />
                  <span>${k}</span>
                </label>
              `).join('')}
            </div>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label class="block font-bold mb-1">Khusus Agama (Opsional)</label>
              <select name="kategori_agama" class="w-full px-3 py-2 rounded-xl border border-slate-300 bg-white">
                <option value="" ${!existing?.kategori_agama ? 'selected' : ''}>Umum (Semua Agama)</option>
                ${['Islam', 'Kristen', 'Katolik', 'Hindu', 'Buddha', 'Konghucu'].map(ag => `<option value="${ag}" ${existing?.kategori_agama === ag ? 'selected' : ''}>${ag}</option>`).join('')}
              </select>
            </div>
            <div class="flex items-center pt-5">
              <label class="inline-flex items-center gap-2 font-bold cursor-pointer">
                <input type="checkbox" name="is_english" ${existing?.is_english ? 'checked' : ''} class="rounded border-slate-300" />
                <span>Mapel Bahasa Inggris (Anti-Translate)</span>
              </label>
            </div>
          </div>

          <div class="flex justify-end gap-2 pt-2">
            <button type="button" data-close-modal class="px-4 py-2 rounded-xl bg-slate-100 font-bold">Batal</button>
            <button type="submit" class="px-4 py-2 rounded-xl bg-[#4f6d4a] text-white font-extrabold">💾 Simpan Mapel</button>
          </div>
        </form>
      `;
    } else if (m.type === 'user') {
      const existing = m.data;
      const role = existing ? existing.role : state.filterUserTab;
      const isSiswa = role === 'siswa';
      title = existing ? `Edit Data ${isSiswa ? 'Siswa' : 'Guru'}` : `Tambah ${isSiswa ? 'Siswa' : 'Guru'} Baru`;

      body = `
        <form id="form-save-user" class="space-y-3 text-xs">
          <input type="hidden" name="role" value="${role}" />
          <div>
            <label class="block font-bold mb-1">Nama Lengkap</label>
            <input name="nama_lengkap" value="${escapeHtml(existing?.nama_lengkap || '')}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-semibold" />
          </div>
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label class="block font-bold mb-1">${isSiswa ? 'NISN' : 'NIP'}</label>
              <input name="nip_nisn" value="${escapeHtml(existing?.nip_nisn || '')}" class="w-full px-3 py-2 rounded-xl border border-slate-300 font-mono" />
            </div>
            ${isSiswa ? `
              <div>
                <label class="block font-bold mb-1">Nomor Peserta Ujian</label>
                <input name="nomor_peserta" value="${escapeHtml(existing?.nomor_peserta || '')}" placeholder="26-2697-0001-9" class="w-full px-3 py-2 rounded-xl border border-slate-300 font-mono" />
              </div>
            ` : ''}
          </div>
          ${isSiswa ? `
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label class="block font-bold mb-1">Kelas</label>
                <input name="kelas" value="${escapeHtml(existing?.kelas || '7a')}" required placeholder="Contoh: 7a" class="w-full px-3 py-2 rounded-xl border border-slate-300 font-bold" />
              </div>
              <div>
                <label class="block font-bold mb-1">Agama</label>
                <select name="agama" class="w-full px-3 py-2 rounded-xl border border-slate-300 bg-white">
                  ${['Islam', 'Kristen', 'Katolik', 'Hindu', 'Buddha', 'Konghucu'].map(ag => `<option value="${ag}" ${(existing?.agama || 'Islam') === ag ? 'selected' : ''}>${ag}</option>`).join('')}
                </select>
              </div>
              <div class="sm:col-span-2">
                <label class="block font-bold mb-1">Ruang &amp; Sesi Ujian</label>
                <input name="ruang_sesi" value="${escapeHtml(existing?.ruang_sesi || 'Ruang 01 / Lab Komputer 1 / Sesi 1')}" class="w-full px-3 py-2 rounded-xl border border-slate-300" />
              </div>
            </div>
          ` : ''}
          <div class="grid grid-cols-2 gap-3">
            <div>
              <label class="block font-bold mb-1">Username Login</label>
              <input name="username" value="${escapeHtml(existing?.username || '')}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-mono font-bold" />
            </div>
            <div>
              <label class="block font-bold mb-1">Password Login</label>
              <input name="password" value="${escapeHtml(existing?.password || '')}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-mono font-bold" />
            </div>
          </div>
          <div class="flex justify-end gap-2 pt-2">
            <button type="button" data-close-modal class="px-4 py-2 rounded-xl bg-slate-100 font-bold">Batal</button>
            <button type="submit" class="px-4 py-2 rounded-xl bg-[#4f6d4a] text-white font-extrabold">💾 Simpan Akun</button>
          </div>
        </form>
      `;
    } else if (m.type === 'jadwal') {
      const existing = m.data;
      title = existing ? 'Edit Jadwal & Token Ujian' : 'Buat Jadwal & Token Ujian Baru';
      const authMapels = getAuthorizedMapels();
      const mapelId = state.modalJadwalMapelId || existing?.mapel_id || authMapels[0]?.id || '';
      const mapelSoal = state.db.bank_soal.filter(s => s.mapel_id === mapelId).sort((a, b) => (a.nomor_urut || 0) - (b.nomor_urut || 0));
      const selectedSoal = state.modalJadwalSelectedSoal.size > 0
        ? state.modalJadwalSelectedSoal
        : new Set(existing && Array.isArray(existing.soal_ids_json) && existing.soal_ids_json.length > 0 ? existing.soal_ids_json : mapelSoal.map(s => s.id));

      const defaultMulai = toDatetimeLocalWita(existing?.waktu_mulai || new Date().toISOString());
      const defaultSelesai = toDatetimeLocalWita(existing?.waktu_selesai || new Date(Date.now() + 2 * 3600000).toISOString());
      const randomToken = existing?.token_ujian || Math.random().toString(36).substring(2, 8).toUpperCase();
      const authKelas = getAuthorizedKelasList(mapelId);

      body = `
        <form id="form-save-jadwal" class="space-y-3 text-xs">
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label class="block font-bold mb-1">Mata Pelajaran</label>
              <select id="modal-jadwal-mapel" name="mapel_id" required class="w-full px-3 py-2 rounded-xl border border-slate-300 bg-white font-bold">
                ${authMapels.map(mp => `<option value="${mp.id}" ${mapelId === mp.id ? 'selected' : ''}>${escapeHtml(mp.nama_mapel)}</option>`).join('')}
              </select>
            </div>
            <div>
              <label class="block font-bold mb-1">Judul Ujian / Asesmen</label>
              <input name="judul_ujian" value="${escapeHtml(existing?.judul_ujian || 'Asesmen Sumatif')}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-semibold" />
            </div>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div>
              <label class="block font-bold mb-1">Target Kelas</label>
              <input name="kelas_target" value="${escapeHtml(existing?.kelas_target || authKelas.join(', '))}" required placeholder="7a, 7b" class="w-full px-3 py-2 rounded-xl border border-slate-300 font-bold" />
            </div>
            <div>
              <label class="block font-bold mb-1">Durasi (Menit)</label>
              <input type="number" min="5" name="durasi_menit" value="${existing?.durasi_menit || 90}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-bold" />
            </div>
            <div>
              <label class="block font-bold mb-1">Token Masuk Ujian</label>
              <div class="flex gap-1">
                <input id="input-jadwal-token" name="token_ujian" value="${escapeHtml(randomToken)}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-mono font-extrabold uppercase" />
                <button id="btn-gen-token" type="button" class="px-2.5 py-2 rounded-xl bg-amber-100 text-amber-900 font-bold">🎲</button>
              </div>
            </div>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label class="block font-bold mb-1">Waktu Mulai (WITA)</label>
              <input type="datetime-local" name="waktu_mulai" value="${defaultMulai}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-semibold" />
            </div>
            <div>
              <label class="block font-bold mb-1">Batas Akhir Selesai (WITA)</label>
              <input type="datetime-local" name="waktu_selesai" value="${defaultSelesai}" required class="w-full px-3 py-2 rounded-xl border border-slate-300 font-semibold" />
            </div>
          </div>

          <div class="flex flex-wrap gap-4 p-3 rounded-xl bg-slate-50 border border-slate-200">
            <label class="inline-flex items-center gap-2 font-bold cursor-pointer">
              <input type="checkbox" name="acak_soal" ${existing?.acak_soal ? 'checked' : ''} class="rounded border-slate-300" />
              <span>🔀 Acak Urutan Soal</span>
            </label>
            <label class="inline-flex items-center gap-2 font-bold cursor-pointer">
              <input type="checkbox" name="acak_opsi" ${existing?.acak_opsi ? 'checked' : ''} class="rounded border-slate-300" />
              <span>🎲 Acak Pilihan Jawaban (A/B/C/D &amp; Pasangan)</span>
            </label>
            <div class="ml-auto flex items-center gap-2">
              <span class="font-bold">Status:</span>
              <select name="status" class="px-2.5 py-1 rounded-lg border border-slate-300 bg-white font-bold">
                <option value="Aktif" ${(existing?.status || 'Aktif') === 'Aktif' ? 'selected' : ''}>Aktif</option>
                <option value="Selesai" ${existing?.status === 'Selesai' ? 'selected' : ''}>Selesai / Tutup</option>
              </select>
            </div>
          </div>

          <div>
            <label class="block font-bold mb-1">Pilih Butir Soal yang Diujikan (${selectedSoal.size} dari ${mapelSoal.length} soal)</label>
            <div class="max-h-40 overflow-y-auto p-2.5 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
              ${mapelSoal.length === 0 ? '<div class="text-slate-400">Belum ada soal pada mapel ini.</div>' : mapelSoal.map(s => `
                <label class="flex items-start gap-2 p-1.5 rounded-lg bg-white border border-slate-200 cursor-pointer">
                  <input type="checkbox" value="${s.id}" class="chk-jadwal-soal mt-0.5" ${selectedSoal.has(s.id) ? 'checked' : ''} />
                  <span class="truncate"><strong>No. ${s.nomor_urut}</strong> [${escapeHtml(normalizeTipeSoal(s.tipe_soal))}]: ${escapeHtml(s.pertanyaan)}</span>
                </label>
              `).join('')}
            </div>
          </div>

          <div class="flex justify-end gap-2 pt-2">
            <button type="button" data-close-modal class="px-4 py-2 rounded-xl bg-slate-100 font-bold">Batal</button>
            <button type="submit" class="px-4 py-2 rounded-xl bg-[#4f6d4a] text-white font-extrabold">💾 Simpan Jadwal</button>
          </div>
        </form>
      `;
    } else if (m.type === 'token_exam') {
      const jadwal = m.data;
      const mp = state.db.mata_pelajaran.find(x => x.id === jadwal.mapel_id) || { nama_mapel: 'Ujian CBT' };
      title = 'Konfirmasi Masuk Ruang Ujian';
      body = `
        <form id="form-enter-exam-token" class="space-y-4 text-xs">
          <div class="p-3.5 rounded-2xl bg-emerald-50 border border-emerald-200 space-y-1">
            <div class="font-extrabold text-sm text-emerald-950">${escapeHtml(mp.nama_mapel)} - ${escapeHtml(jadwal.judul_ujian)}</div>
            <div class="text-emerald-800">Durasi: <strong>${jadwal.durasi_menit} Menit</strong> | Batas Selesai: <strong>${formatDateTimeWita(jadwal.waktu_selesai)}</strong></div>
          </div>

          <div>
            <label class="block font-extrabold text-slate-800 mb-1.5">Masukkan Token Ujian dari Pengawas:</label>
            <input id="input-student-token" type="text" required autocomplete="off" placeholder="KETIK TOKEN UJIAN..."
              class="w-full px-4 py-3 rounded-2xl border-2 border-slate-300 font-mono text-center text-lg font-extrabold uppercase tracking-widest focus:border-[#4f6d4a] focus:outline-none" />
          </div>

          <div class="p-3 rounded-xl bg-amber-50 border border-amber-200 text-[11px] text-amber-950 space-y-1">
            <div class="font-extrabold">⚠️ PERHATIAN KEAMANAN UJIAN:</div>
            <div>• Ujian berjalan dalam mode Layar Penuh (Fullscreen) dan mengunci tombol Kembali.</div>
            <div>• Dilarang membuka tab lain, layar belah, screenshot, atau asisten AI.</div>
          </div>

          <div class="flex justify-end gap-2">
            <button type="button" data-close-modal class="px-4 py-2.5 rounded-xl bg-slate-100 font-bold">Batal</button>
            <button type="submit" class="px-5 py-2.5 rounded-xl bg-[#4f6d4a] text-white font-extrabold">▶ Mulai Kerjakan Sekarang</button>
          </div>
        </form>
      `;
    } else if (m.type === 'grade_essay') {
      const sesi = m.data;
      const siswa = state.db.users_cbt.find(u => u.id === sesi.siswa_id) || { nama_lengkap: 'Siswa' };
      const jd = state.db.jadwal_ujian.find(j => j.id === sesi.jadwal_id);
      const essaySoals = jd
        ? state.db.bank_soal.filter(s => s.mapel_id === jd.mapel_id && normalizeTipeSoal(s.tipe_soal) === 'URAIAN')
        : [];

      title = `Input Skor Essay - ${siswa.nama_lengkap}`;
      body = `
        <form id="form-grade-essay" class="space-y-3 text-xs">
          ${essaySoals.length === 0 ? `
            <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 text-slate-600">Tidak ada butir soal bertipe Uraian pada jadwal ini.</div>
          ` : `
            <div class="max-h-60 overflow-y-auto space-y-2.5">
              ${essaySoals.map((s, idx) => `
                <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
                  <div class="font-extrabold text-slate-800">Soal Uraian #${idx + 1}: ${escapeHtml(s.pertanyaan)}</div>
                  <div class="p-2 rounded-lg bg-white border border-slate-200 text-slate-900 whitespace-pre-line">
                    <strong>Jawaban Siswa:</strong><br/>${escapeHtml(sesi.jawaban_json?.[s.id] || '(Tidak dijawab)')}
                  </div>
                  <div class="text-[11px] text-emerald-700 font-bold">Pedoman Kunci: ${escapeHtml(s.kunci_jawaban || '-')}</div>
                </div>
              `).join('')}
            </div>
          `}
          <div>
            <label class="block font-extrabold mb-1">Nilai Total Bagian Essay / Uraian (Skala 0 - 100)</label>
            <input type="number" min="0" max="100" step="0.5" name="skor_essay" value="${sesi.skor_essay ?? 0}" required
              class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 font-extrabold text-sm" />
          </div>
          <div class="flex justify-end gap-2 pt-2">
            <button type="button" data-close-modal class="px-4 py-2 rounded-xl bg-slate-100 font-bold">Batal</button>
            <button type="submit" class="px-4 py-2 rounded-xl bg-[#4f6d4a] text-white font-extrabold">💾 Simpan Nilai Akhir</button>
          </div>
        </form>
      `;
    }

    return `
      <div class="fixed inset-0 z-50 bg-slate-900/55 backdrop-blur-xs flex items-center justify-center p-3.5 overflow-y-auto">
        <div class="bg-gradient-to-b from-white via-[#fcfefd] to-[#f6faf8] rounded-3xl max-w-2xl w-full p-5 sm:p-6 shadow-2xl border border-emerald-200/80 space-y-4 my-auto max-h-[92dvh] overflow-y-auto">
          <div class="flex items-center justify-between pb-3 border-b border-emerald-200/70">
            <h3 class="font-extrabold text-base bg-gradient-to-r from-slate-900 via-emerald-900 to-teal-800 bg-clip-text text-transparent">${escapeHtml(title)}</h3>
            <button type="button" data-close-modal class="w-8 h-8 rounded-xl bg-rose-50 hover:bg-rose-100 border border-rose-200 font-bold text-rose-700">✕</button>
          </div>
          ${body}
        </div>
      </div>
    `;
  }

  function renderConfirmDialogOverlay() {
    if (!state.confirmDialog) return '';
    const d = state.confirmDialog;
    const isDanger = d.variant === 'danger';
    return `
      <div class="fixed inset-0 z-[60] bg-slate-900/65 backdrop-blur-xs flex items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 shadow-2xl border border-slate-200 space-y-4">
          <h3 class="font-extrabold text-base text-slate-900">${escapeHtml(d.title)}</h3>
          <p class="text-xs text-slate-600 leading-relaxed">${escapeHtml(d.message)}</p>
          <div class="flex justify-end gap-2 pt-2">
            <button id="btn-confirm-cancel" type="button" class="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs">Batal</button>
            <button id="btn-confirm-ok" type="button" class="px-4 py-2.5 rounded-xl ${isDanger ? 'bg-rose-600 hover:bg-rose-700' : 'bg-[#4f6d4a] hover:bg-[#3f593b]'} text-white font-extrabold text-xs">${escapeHtml(d.confirmLabel)}</button>
          </div>
        </div>
      </div>
    `;
  }

  function renderToastOverlay() {
    if (!state.toast) return '';
    const isErr = state.toast.type === 'error';
    return `
      <div class="fixed bottom-4 right-4 left-4 sm:left-auto sm:max-w-md z-[70] px-4 py-3 rounded-2xl shadow-xl border text-xs font-extrabold flex items-center gap-2.5 ${
        isErr ? 'bg-rose-950 text-rose-100 border-rose-600' : 'bg-slate-900 text-white border-emerald-500'
      }">
        <span>${isErr ? '⚠️' : '✅'}</span>
        <span class="flex-1">${escapeHtml(state.toast.message)}</span>
      </div>
    `;
  }

  // =========================================================================
  // MAIN RENDER & EVENT BINDING
  // =========================================================================
  function render() {
    const root = document.getElementById('cbt-promax-root');
    if (!root) return;

    let mainHtml = '';
    if (!state.currentUser) {
      mainHtml = renderLoginView();
    } else if (state.exam.active) {
      mainHtml = renderStudentExamView();
    } else if (state.currentUser.role === 'siswa') {
      const content = state.activeMenu === 'nilai_siswa' ? renderStudentNilaiView() : renderStudentDashboardView();
      mainHtml = renderHeaderAndSidebar(content);
    } else {
      let content = '';
      switch (state.activeMenu) {
        case 'bank_soal': content = renderBankSoalView(); break;
        case 'jadwal': content = renderJadwalView(); break;
        case 'monitoring': content = renderMonitoringView(); break;
        case 'rekap_nilai': content = renderRekapNilaiView(); break;
        case 'cetak_dokumen': content = renderCetakDokumenView(); break;
        case 'mapel': content = renderMapelView(); break;
        case 'users': content = renderUsersView(); break;
        case 'pengaturan': content = renderPengaturanView(); break;
        default: content = renderDashboardView(); break;
      }
      mainHtml = renderHeaderAndSidebar(content);
    }

    root.innerHTML = `
      ${mainHtml}
      ${renderModalOverlay()}
      ${renderConfirmDialogOverlay()}
      ${renderToastOverlay()}
    `;

    attachEventListeners();
  }

  function attachEventListeners() {
    // Login
    document.getElementById('login-form')?.addEventListener('submit', (e) => {
      e.preventDefault();
      const u = document.getElementById('login-username')?.value;
      const p = document.getElementById('login-password')?.value;
      handleLogin(u, p);
    });

    document.getElementById('toggle-pwd-btn')?.addEventListener('click', () => {
      const inp = document.getElementById('login-password');
      if (inp) inp.type = inp.type === 'password' ? 'text' : 'password';
    });

    // Sidebar & Navigation
    document.getElementById('btn-toggle-sidebar')?.addEventListener('click', () => {
      state.sidebarOpen = !state.sidebarOpen;
      render();
    });

    document.getElementById('sidebar-backdrop')?.addEventListener('click', () => {
      state.sidebarOpen = false;
      render();
    });

    document.querySelectorAll('.nav-menu-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        state.activeMenu = btn.getAttribute('data-menu');
        if (window.innerWidth < 1024) state.sidebarOpen = false;
        render();
      });
    });

    document.getElementById('btn-logout')?.addEventListener('click', () => {
      openConfirmDialog({
        title: 'Konfirmasi Keluar Akun',
        message: 'Apakah Anda yakin ingin keluar dari aplikasi CBT Promax?',
        confirmLabel: 'Ya, Keluar',
        variant: 'danger',
        onConfirm: handleLogout
      });
    });

    // Confirm Dialog
    document.getElementById('btn-confirm-cancel')?.addEventListener('click', () => {
      state.confirmDialog = null;
      render();
    });
    document.getElementById('btn-confirm-ok')?.addEventListener('click', async () => {
      const fn = state.confirmDialog?.onConfirm;
      state.confirmDialog = null;
      render();
      if (typeof fn === 'function') await fn();
    });

    // Close Modal
    document.querySelectorAll('[data-close-modal]').forEach(btn => {
      btn.addEventListener('click', () => {
        state.modal = null;
        state.modalSoalDraft = null;
        state.modalSoalImagePreview = '';
        state.modalSoalSelectedKelas = new Set();
        state.modalMapelSelectedKelas = new Set();
        state.modalImportSelectedKelas = new Set();
        state.modalJadwalSelectedSoal = new Set();
        render();
      });
    });

    // Bank Soal: Ketuk Mapel Sekali Buka, Ketuk Sekali Lagi Tutup
    document.querySelectorAll('[data-select-soal-mapel]').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-select-soal-mapel');
        state.filterSoalMapel = state.filterSoalMapel === id ? '' : id;
        state.filterSoalKelas = '';
        state.selectedIds.soal.clear();
        render();
      });
    });

    document.getElementById('filter-soal-kelas')?.addEventListener('change', (e) => {
      state.filterSoalKelas = e.target.value;
      render();
    });

    document.getElementById('check-all-soal')?.addEventListener('change', (e) => {
      const checked = e.target.checked;
      document.querySelectorAll('.row-check-soal').forEach(cb => {
        const id = cb.getAttribute('data-id');
        if (checked) state.selectedIds.soal.add(id);
        else state.selectedIds.soal.delete(id);
      });
      render();
    });

    document.querySelectorAll('.row-check-soal').forEach(cb => {
      cb.addEventListener('change', (e) => {
        const id = cb.getAttribute('data-id');
        if (e.target.checked) state.selectedIds.soal.add(id);
        else state.selectedIds.soal.delete(id);
        render();
      });
    });

    document.getElementById('btn-bulk-delete-soal')?.addEventListener('click', () => {
      const ids = [...state.selectedIds.soal];
      openConfirmDialog({
        title: 'Hapus Butir Soal Terpilih',
        message: `Hapus permanen ${ids.length} butir soal yang dipilih?`,
        confirmLabel: 'Hapus Permanen',
        variant: 'danger',
        onConfirm: async () => {
          await dbDeleteMany('bank_soal', ids);
          state.selectedIds.soal.clear();
          showToast(`${ids.length} butir soal berhasil dihapus.`);
          render();
        }
      });
    });

    document.getElementById('btn-open-add-soal')?.addEventListener('click', () => {
      const authMapels = getAuthorizedMapels();
      if (authMapels.length === 0) {
        showToast('Tambahkan Mata Pelajaran terlebih dahulu.', 'error');
        return;
      }
      const defaultMapelId = state.filterSoalMapel || authMapels[0]?.id || '';
      state.modalSoalDraft = buildInitialSoalDraft(null, defaultMapelId);
      state.modalSoalImagePreview = '';
      state.modalSoalSelectedKelas = new Set(getAuthorizedKelasList(defaultMapelId));
      state.modal = { type: 'soal', data: null };
      render();
    });

    document.querySelectorAll('[data-edit-soal]').forEach(btn => {
      btn.addEventListener('click', () => {
        const s = state.db.bank_soal.find(x => x.id === btn.getAttribute('data-edit-soal'));
        if (!s) return;
        state.modalSoalDraft = buildInitialSoalDraft(s, s.mapel_id);
        state.modalSoalImagePreview = s.gambar_url || '';
        state.modalSoalSelectedKelas = new Set(parseKelasString(s.kelas_target));
        state.modal = { type: 'soal', data: s };
        render();
      });
    });

    document.querySelectorAll('[data-del-soal]').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-del-soal');
        openConfirmDialog({
          title: 'Hapus Butir Soal',
          message: 'Apakah Anda yakin ingin menghapus butir soal ini?',
          confirmLabel: 'Hapus',
          variant: 'danger',
          onConfirm: async () => {
            await dbDeleteMany('bank_soal', [id]);
            showToast('Butir soal dihapus.');
            render();
          }
        });
      });
    });

    document.getElementById('modal-soal-mapel')?.addEventListener('change', (e) => {
      syncModalSoalDraftFromDom();
      const mapelId = e.target.value;
      if (state.modalSoalDraft) state.modalSoalDraft.mapel_id = mapelId;
      state.modalSoalSelectedKelas = new Set(getAuthorizedKelasList(mapelId));
      render();
    });

    document.getElementById('modal-soal-tipe')?.addEventListener('change', () => {
      syncModalSoalDraftFromDom();
      render();
    });

    document.getElementById('btn-add-pasangan-row')?.addEventListener('click', () => {
      syncModalSoalDraftFromDom();
      if (state.modalSoalDraft) {
        state.modalSoalDraft.pasangan.push({ kiri: '', kanan: '' });
      }
      render();
    });

    document.querySelectorAll('[data-del-pasangan-row]').forEach(btn => {
      btn.addEventListener('click', () => {
        syncModalSoalDraftFromDom();
        const idx = Number(btn.getAttribute('data-del-pasangan-row'));
        if (state.modalSoalDraft && Array.isArray(state.modalSoalDraft.pasangan) && state.modalSoalDraft.pasangan.length > 1) {
          state.modalSoalDraft.pasangan.splice(idx, 1);
        }
        render();
      });
    });

    document.getElementById('input-soal-image')?.addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      syncModalSoalDraftFromDom();
      state.modalSoalImagePreview = await compressImageToDataUrl(file, 720, 0.85);
      render();
    });

    document.getElementById('btn-clear-soal-img')?.addEventListener('click', () => {
      syncModalSoalDraftFromDom();
      state.modalSoalImagePreview = '';
      if (state.modal?.data) state.modal.data.gambar_url = '';
      render();
    });

    document.getElementById('form-save-soal')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      syncModalSoalDraftFromDom();
      const d = state.modalSoalDraft;
      const form = e.target;
      const fd = new FormData(form);
      const mapelId = String(fd.get('mapel_id') || d?.mapel_id || '');
      const tipe = normalizeTipeSoal(fd.get('tipe_soal') || d?.tipe_soal || 'PILIHAN GANDA');
      const checkedKelas = [...form.querySelectorAll('.chk-soal-kelas:checked')].map(c => c.value);

      let opsiA = '';
      let opsiB = '';
      let opsiC = '';
      let opsiD = '';
      let pasanganJson = [];
      let kunciJawaban = '';

      if (tipe === 'PILIHAN GANDA') {
        opsiA = String(fd.get('opsi_a') || '').trim();
        opsiB = String(fd.get('opsi_b') || '').trim();
        opsiC = String(fd.get('opsi_c') || '').trim();
        opsiD = String(fd.get('opsi_d') || '').trim();
        kunciJawaban = String(fd.get('kunci_pg_radio') || 'A').trim().toUpperCase();
        if (!opsiA && !opsiB) {
          showToast('Isi minimal opsi A dan B untuk soal Pilihan Ganda.', 'error');
          return;
        }
      } else if (tipe === 'PILIHAN GANDA KOMPLEKS') {
        opsiA = String(fd.get('opsi_a') || '').trim();
        opsiB = String(fd.get('opsi_b') || '').trim();
        opsiC = String(fd.get('opsi_c') || '').trim();
        opsiD = String(fd.get('opsi_d') || '').trim();
        const kChecks = [...form.querySelectorAll('.chk-kunci-pgk:checked')].map(c => c.value.toUpperCase());
        if (kChecks.length === 0) {
          showToast('Centang minimal satu kunci jawaban benar pada Pilihan Ganda Kompleks.', 'error');
          return;
        }
        kunciJawaban = kChecks.join(', ');
      } else if (tipe === 'BENAR SALAH') {
        opsiA = 'BENAR';
        opsiB = 'SALAH';
        kunciJawaban = String(fd.get('kunci_bs_radio') || 'BENAR').trim().toUpperCase();
      } else if (tipe === 'MENJODOHKAN') {
        const cleanPairs = (d?.pasangan || [])
          .map(p => ({ kiri: String(p.kiri || '').trim(), kanan: String(p.kanan || '').trim() }))
          .filter(p => p.kiri && p.kanan);
        if (cleanPairs.length < 2) {
          showToast('Isi minimal 2 baris pasangan kiri dan kanan yang lengkap untuk soal Menjodohkan.', 'error');
          return;
        }
        pasanganJson = cleanPairs;
        kunciJawaban = 'SESUAI PASANGAN';
      } else if (tipe === 'ISIAN SINGKAT') {
        kunciJawaban = String(fd.get('kunci_jawaban') || '').trim();
        if (!kunciJawaban) {
          showToast('Isi kunci jawaban untuk soal Isian Singkat.', 'error');
          return;
        }
      } else {
        kunciJawaban = String(fd.get('kunci_jawaban') || '').trim();
      }

      const payload = {
        mapel_id: mapelId,
        kelas_target: checkedKelas.join(', '),
        nomor_urut: Number(fd.get('nomor_urut')) || 1,
        tipe_soal: tipe,
        pertanyaan: String(fd.get('pertanyaan') || '').trim(),
        gambar_url: state.modalSoalImagePreview || '',
        opsi_a: opsiA,
        opsi_b: opsiB,
        opsi_c: opsiC,
        opsi_d: opsiD,
        pasangan_json: pasanganJson,
        kunci_jawaban: kunciJawaban,
        bobot: Number(fd.get('bobot')) || 10,
        topik: d?.topik || 'Umum'
      };

      try {
        if (state.modal.data) {
          await dbUpdate('bank_soal', state.modal.data.id, payload);
          showToast('Perubahan butir soal berhasil disimpan!');
        } else {
          await dbInsert('bank_soal', payload);
          showToast('Butir soal baru berhasil ditambahkan!');
        }
        state.filterSoalMapel = payload.mapel_id;
        state.modal = null;
        state.modalSoalDraft = null;
        state.modalSoalImagePreview = '';
        render();
      } catch (err) {
        showToast('Gagal menyimpan soal: ' + (err.message || err), 'error');
      }
    });

    // Template & Import Word (.docx)
    document.getElementById('btn-download-word-tpl')?.addEventListener('click', downloadWordSoalTemplate);

    document.getElementById('btn-open-import-word')?.addEventListener('click', () => {
      const authMapels = getAuthorizedMapels();
      if (authMapels.length === 0) {
        showToast('Tambahkan Mata Pelajaran terlebih dahulu.', 'error');
        return;
      }
      state.modalImportMapelId = state.filterSoalMapel || authMapels[0].id;
      state.modalImportSelectedKelas = new Set(getAuthorizedKelasList(state.modalImportMapelId));
      state.modal = { type: 'import_word' };
      render();
    });

    document.getElementById('import-word-mapel')?.addEventListener('change', (e) => {
      state.modalImportMapelId = e.target.value;
      state.modalImportSelectedKelas = new Set(getAuthorizedKelasList(e.target.value));
      render();
    });

    document.getElementById('form-import-word-soal')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      if (typeof window.mammoth === 'undefined') {
        showToast('Pustaka pembaca Word (.docx) sedang dimuat, silakan coba lagi.', 'error');
        return;
      }
      const form = e.target;
      const mapelId = form.querySelector('#import-word-mapel')?.value;
      const checkedKelas = [...form.querySelectorAll('.chk-import-kelas:checked')].map(c => c.value).join(', ');
      const fileInput = document.getElementById('input-file-word-docx');
      const file = fileInput?.files?.[0];
      if (!file) return;

      try {
        const arrayBuffer = await file.arrayBuffer();
        const result = await window.mammoth.convertToHtml({ arrayBuffer });
        const startNo = state.db.bank_soal.filter(s => s.mapel_id === mapelId).length + 1;
        const parsedQuestions = parseWordHtmlToQuestions(result.value, mapelId, checkedKelas, startNo);

        if (parsedQuestions.length === 0) {
          showToast('Tidak ditemukan format soal bernomor (1., 2., dst) pada dokumen Word.', 'error');
          return;
        }

        for (const q of parsedQuestions) {
          await dbInsert('bank_soal', q);
        }
        state.filterSoalMapel = mapelId;
        state.modal = null;
        showToast(`${parsedQuestions.length} butir soal berhasil diimpor dari Word!`);
        render();
      } catch (err) {
        showToast('Gagal mengimpor file Word: ' + (err.message || err), 'error');
      }
    });

    // Monitoring Live: Ketuk Sekali Buka / Ketuk Lagi Tutup
    document.querySelectorAll('[data-select-mon-jadwal]').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-select-mon-jadwal');
        state.filterMonitoringJadwal = state.filterMonitoringJadwal === id ? '' : id;
        state.filterMonitoringKelas = '';
        state.filterMonitoringSearch = '';
        render();
      });
    });

    document.getElementById('btn-refresh-monitoring')?.addEventListener('click', async () => {
      await loadAllData();
      showToast('Data monitoring peserta berhasil diperbarui.');
      render();
    });

    document.getElementById('filter-mon-kelas')?.addEventListener('change', (e) => {
      state.filterMonitoringKelas = e.target.value;
      render();
    });

    document.getElementById('filter-mon-search')?.addEventListener('input', (e) => {
      state.filterMonitoringSearch = e.target.value;
      render();
      const inp = document.getElementById('filter-mon-search');
      if (inp) {
        inp.focus();
        inp.setSelectionRange(inp.value.length, inp.value.length);
      }
    });

    document.querySelectorAll('[data-unlock-sesi]').forEach(btn => {
      btn.addEventListener('click', () => unlockStudentSession(btn.getAttribute('data-unlock-sesi')));
    });

    document.querySelectorAll('[data-reset-login]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const uid = btn.getAttribute('data-reset-login');
        await dbUpdate('users_cbt', uid, { is_logged_in: false });
        showToast('Status login peserta berhasil direset. Siswa dapat login kembali.');
        render();
      });
    });

    // Cetak Dokumen A4
    document.querySelectorAll('[data-doc-type]').forEach(btn => {
      btn.addEventListener('click', () => {
        state.docType = btn.getAttribute('data-doc-type');
        render();
      });
    });

    document.getElementById('select-doc-jadwal')?.addEventListener('change', (e) => {
      state.docJadwalId = e.target.value;
      render();
    });

    document.getElementById('select-doc-kelas')?.addEventListener('change', (e) => {
      state.docKelas = e.target.value;
      render();
    });

    document.getElementById('btn-toggle-edit-ba')?.addEventListener('click', () => {
      state.editingBeritaAcara = !state.editingBeritaAcara;
      render();
    });

    document.getElementById('form-berita-acara')?.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      state.beritaAcaraForm = Object.fromEntries(fd.entries());
      state.editingBeritaAcara = false;
      showToast('Perubahan isi Berita Acara berhasil diterapkan!');
      render();
    });

    document.getElementById('btn-print-doc-a4')?.addEventListener('click', () => {
      const printEl = document.getElementById('print-area');
      if (printEl) {
        openA4PrintWindow('Dokumen Resmi CBT Promax', printEl.innerHTML);
      }
    });

    // Mata Pelajaran CRUD & Excel
    document.getElementById('btn-tpl-mapel-excel')?.addEventListener('click', () => downloadExcelTemplate('mapel'));
    document.getElementById('input-import-mapel-excel')?.addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file || typeof window.XLSX === 'undefined') return;
      const buf = await file.arrayBuffer();
      const wb = window.XLSX.read(buf);
      const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      let count = 0;
      for (const r of rows) {
        const nama = String(r.Nama_Mapel || r.nama_mapel || '').trim();
        if (!nama) continue;
        const gUser = String(r.Username_Guru || '').trim();
        const guru = state.db.users_cbt.find(u => u.role === 'guru' && u.username === gUser);
        await dbInsert('mata_pelajaran', {
          kode_mapel: String(r.Kode_Mapel || `MP-${count + 1}`).trim(),
          nama_mapel: nama,
          kelas_ampu: String(r.Kelas_Ampu || '').trim(),
          guru_id: guru ? guru.id : (state.currentUser.role === 'guru' ? state.currentUser.id : null),
          kategori_agama: String(r.Khusus_Agama || '').trim(),
          is_english: String(r.Bahasa_Inggris || '').trim().toUpperCase() === 'YA'
        });
        count++;
      }
      showToast(`${count} mata pelajaran berhasil diimpor!`);
      render();
    });

    document.getElementById('btn-open-add-mapel')?.addEventListener('click', () => {
      state.modalMapelSelectedKelas = new Set();
      state.modal = { type: 'mapel', data: null };
      render();
    });

    document.querySelectorAll('[data-edit-mapel]').forEach(btn => {
      btn.addEventListener('click', () => {
        const mp = state.db.mata_pelajaran.find(x => x.id === btn.getAttribute('data-edit-mapel'));
        if (!mp) return;
        state.modalMapelSelectedKelas = new Set(parseKelasString(mp.kelas_ampu));
        state.modal = { type: 'mapel', data: mp };
        render();
      });
    });

    document.querySelectorAll('[data-del-mapel]').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-del-mapel');
        openConfirmDialog({
          title: 'Hapus Mata Pelajaran',
          message: 'Menghapus mata pelajaran juga akan menghapus bank soal dan jadwal ujian terkait. Lanjutkan?',
          confirmLabel: 'Hapus Permanen',
          variant: 'danger',
          onConfirm: async () => {
            await dbDeleteMany('mata_pelajaran', [id]);
            showToast('Mata pelajaran dihapus.');
            render();
          }
        });
      });
    });

    document.getElementById('form-save-mapel')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const form = e.target;
      const fd = new FormData(form);
      const checkedKelas = [...form.querySelectorAll('.chk-mapel-kelas:checked')].map(c => c.value);
      const payload = {
        kode_mapel: String(fd.get('kode_mapel') || '').trim(),
        nama_mapel: String(fd.get('nama_mapel') || '').trim(),
        guru_id: fd.get('guru_id') || null,
        kelas_ampu: checkedKelas.join(', '),
        kategori_agama: String(fd.get('kategori_agama') || '').trim(),
        is_english: Boolean(fd.get('is_english'))
      };
      if (state.modal.data) {
        await dbUpdate('mata_pelajaran', state.modal.data.id, payload);
        showToast('Mata pelajaran diperbarui!');
      } else {
        await dbInsert('mata_pelajaran', payload);
        showToast('Mata pelajaran ditambahkan!');
      }
      state.modal = null;
      render();
    });

    // Users Guru & Siswa CRUD + Excel
    document.querySelectorAll('[data-user-tab]').forEach(btn => {
      btn.addEventListener('click', () => {
        state.filterUserTab = btn.getAttribute('data-user-tab');
        render();
      });
    });

    document.getElementById('filter-user-kelas')?.addEventListener('change', (e) => {
      state.filterUserKelas = e.target.value;
      render();
    });

    document.getElementById('btn-tpl-user-excel')?.addEventListener('click', () => downloadExcelTemplate(state.filterUserTab));

    document.getElementById('input-import-user-excel')?.addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file || typeof window.XLSX === 'undefined') return;
      const buf = await file.arrayBuffer();
      const wb = window.XLSX.read(buf);
      const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      const isSiswa = state.filterUserTab === 'siswa';
      let count = 0;
      for (const r of rows) {
        const nama = String(r.Nama_Lengkap || r.nama_lengkap || '').trim();
        const username = String(r.Username || r.username || '').trim();
        const password = String(r.Password || r.password || '').trim();
        if (!nama || !username || !password) continue;
        await dbInsert('users_cbt', {
          nomor_peserta: isSiswa ? String(r.No_Peserta || '-').trim() : '-',
          nip_nisn: String(isSiswa ? (r.NISN || '-') : (r.NIP || '-')).trim(),
          nama_lengkap: nama,
          role: state.filterUserTab,
          kelas: isSiswa ? String(r.Kelas || '7a').trim() : '-',
          agama: isSiswa ? String(r.Agama || 'Islam').trim() : 'Islam',
          ruang_sesi: isSiswa ? String(r.Ruang_Sesi || 'Ruang 01 / Lab Komputer 1 / Sesi 1').trim() : '-',
          username,
          password,
          is_logged_in: false
        });
        count++;
      }
      showToast(`${count} akun ${state.filterUserTab} berhasil diimpor!`);
      render();
    });

    document.getElementById('btn-open-add-user')?.addEventListener('click', () => {
      state.modal = { type: 'user', data: null };
      render();
    });

    document.querySelectorAll('[data-edit-user]').forEach(btn => {
      btn.addEventListener('click', () => {
        const u = state.db.users_cbt.find(x => x.id === btn.getAttribute('data-edit-user'));
        if (u) {
          state.modal = { type: 'user', data: u };
          render();
        }
      });
    });

    document.querySelectorAll('[data-del-user]').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-del-user');
        openConfirmDialog({
          title: 'Hapus Akun Pengguna',
          message: 'Apakah Anda yakin ingin menghapus akun pengguna ini?',
          confirmLabel: 'Hapus',
          variant: 'danger',
          onConfirm: async () => {
            await dbDeleteMany('users_cbt', [id]);
            showToast('Akun pengguna dihapus.');
            render();
          }
        });
      });
    });

    document.getElementById('form-save-user')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const role = fd.get('role');
      const isSiswa = role === 'siswa';
      const payload = {
        role,
        nama_lengkap: String(fd.get('nama_lengkap') || '').trim(),
        nip_nisn: String(fd.get('nip_nisn') || '-').trim(),
        nomor_peserta: isSiswa ? String(fd.get('nomor_peserta') || '-').trim() : '-',
        kelas: isSiswa ? String(fd.get('kelas') || '7a').trim() : '-',
        agama: isSiswa ? String(fd.get('agama') || 'Islam').trim() : 'Islam',
        ruang_sesi: isSiswa ? String(fd.get('ruang_sesi') || 'Ruang 01 / Lab Komputer 1 / Sesi 1').trim() : '-',
        username: String(fd.get('username') || '').trim(),
        password: String(fd.get('password') || '').trim()
      };
      try {
        if (state.modal.data) {
          await dbUpdate('users_cbt', state.modal.data.id, payload);
          showToast('Data akun diperbarui!');
        } else {
          await dbInsert('users_cbt', { ...payload, is_logged_in: false });
          showToast('Akun baru berhasil ditambahkan!');
        }
        state.modal = null;
        render();
      } catch (err) {
        showToast('Gagal menyimpan akun (pastikan username unik): ' + (err.message || err), 'error');
      }
    });

    // Jadwal & Token CRUD
    document.getElementById('btn-open-add-jadwal')?.addEventListener('click', () => {
      const authMapels = getAuthorizedMapels();
      if (authMapels.length === 0) {
        showToast('Tambahkan Mata Pelajaran terlebih dahulu.', 'error');
        return;
      }
      state.modalJadwalMapelId = authMapels[0].id;
      state.modalJadwalSelectedSoal = new Set(state.db.bank_soal.filter(s => s.mapel_id === authMapels[0].id).map(s => s.id));
      state.modal = { type: 'jadwal', data: null };
      render();
    });

    document.querySelectorAll('[data-edit-jadwal]').forEach(btn => {
      btn.addEventListener('click', () => {
        const jd = state.db.jadwal_ujian.find(x => x.id === btn.getAttribute('data-edit-jadwal'));
        if (!jd) return;
        state.modalJadwalMapelId = jd.mapel_id;
        state.modalJadwalSelectedSoal = new Set(Array.isArray(jd.soal_ids_json) ? jd.soal_ids_json : []);
        state.modal = { type: 'jadwal', data: jd };
        render();
      });
    });

    document.querySelectorAll('[data-del-jadwal]').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-del-jadwal');
        openConfirmDialog({
          title: 'Hapus Jadwal Ujian',
          message: 'Apakah Anda yakin ingin menghapus jadwal ujian ini?',
          confirmLabel: 'Hapus',
          variant: 'danger',
          onConfirm: async () => {
            await dbDeleteMany('jadwal_ujian', [id]);
            showToast('Jadwal ujian dihapus.');
            render();
          }
        });
      });
    });

    document.getElementById('modal-jadwal-mapel')?.addEventListener('change', (e) => {
      state.modalJadwalMapelId = e.target.value;
      state.modalJadwalSelectedSoal = new Set(state.db.bank_soal.filter(s => s.mapel_id === e.target.value).map(s => s.id));
      render();
    });

    document.getElementById('btn-gen-token')?.addEventListener('click', () => {
      const inp = document.getElementById('input-jadwal-token');
      if (inp) inp.value = Math.random().toString(36).substring(2, 8).toUpperCase();
    });

    document.getElementById('form-save-jadwal')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const form = e.target;
      const fd = new FormData(form);
      const selectedSoals = [...form.querySelectorAll('.chk-jadwal-soal:checked')].map(c => c.value);
      const payload = {
        mapel_id: fd.get('mapel_id'),
        judul_ujian: String(fd.get('judul_ujian') || '').trim(),
        kelas_target: String(fd.get('kelas_target') || '7a').trim(),
        durasi_menit: Number(fd.get('durasi_menit')) || 90,
        token_ujian: String(fd.get('token_ujian') || '').trim().toUpperCase(),
        waktu_mulai: parseWitaDatetimeLocal(fd.get('waktu_mulai')),
        waktu_selesai: parseWitaDatetimeLocal(fd.get('waktu_selesai')),
        acak_soal: Boolean(fd.get('acak_soal')),
        acak_opsi: Boolean(fd.get('acak_opsi')),
        status: fd.get('status') || 'Aktif',
        jumlah_soal: selectedSoals.length,
        soal_ids_json: selectedSoals
      };
      try {
        if (state.modal.data) {
          await dbUpdate('jadwal_ujian', state.modal.data.id, payload);
          showToast('Jadwal & token ujian diperbarui!');
        } else {
          await dbInsert('jadwal_ujian', payload);
          showToast('Jadwal & token ujian berhasil diterbitkan!');
        }
        state.modal = null;
        render();
      } catch (err) {
        showToast('Gagal menyimpan jadwal: ' + (err.message || err), 'error');
      }
    });

    // Rekap Nilai & Koreksi Essay
    document.getElementById('btn-refresh-nilai')?.addEventListener('click', async () => {
      await loadAllData();
      showToast('Data nilai berhasil diperbarui dari server.');
      render();
    });

    document.getElementById('filter-nilai-jadwal')?.addEventListener('change', (e) => {
      state.filterNilaiJadwal = e.target.value;
      state.filterNilaiKelas = '';
      render();
    });

    document.getElementById('filter-nilai-kelas')?.addEventListener('change', (e) => {
      state.filterNilaiKelas = e.target.value;
      render();
    });

    document.querySelectorAll('[data-grade-essay]').forEach(btn => {
      btn.addEventListener('click', () => {
        const sesi = state.db.sesi_ujian_siswa.find(s => s.id === btn.getAttribute('data-grade-essay'));
        if (sesi) {
          state.modal = { type: 'grade_essay', data: sesi };
          render();
        }
      });
    });

    document.getElementById('form-grade-essay')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const sesi = state.modal.data;
      const skorEssay = Number(fd.get('skor_essay')) || 0;
      const skorPg = Number(sesi.skor_pg) || 0;
      const nilaiAkhir = Math.round(((skorPg + skorEssay) / 2) * 10) / 10;
      await dbUpdate('sesi_ujian_siswa', sesi.id, { skor_essay: skorEssay, nilai_akhir: nilaiAkhir });
      state.modal = null;
      showToast('Nilai essay dan nilai akhir berhasil diperbarui!');
      render();
    });

    document.querySelectorAll('[data-reset-ujian]').forEach(btn => {
      btn.addEventListener('click', () => {
        const sesiId = btn.getAttribute('data-reset-ujian');
        const siswaId = btn.getAttribute('data-siswa-id');
        openConfirmDialog({
          title: 'Reset Ujian Ulang Peserta',
          message: 'Menghapus sesi ini akan mengizinkan peserta mengerjakan ujian dari awal kembali. Lanjutkan?',
          confirmLabel: 'Ya, Reset Ujian',
          variant: 'danger',
          onConfirm: async () => {
            await dbDeleteMany('sesi_ujian_siswa', [sesiId]);
            if (siswaId) await dbUpdate('users_cbt', siswaId, { is_logged_in: false });
            showToast('Sesi ujian berhasil direset. Peserta dapat mengikuti ujian ulang.');
            render();
          }
        });
      });
    });

    document.getElementById('btn-export-nilai-excel')?.addEventListener('click', () => {
      const jd = state.db.jadwal_ujian.find(j => String(j.id) === String(state.filterNilaiJadwal));
      if (!jd) return;
      const rows = state.db.sesi_ujian_siswa
        .filter(s => String(s.jadwal_id) === String(jd.id))
        .map(sesi => ({ sesi, siswa: state.db.users_cbt.find(u => String(u.id) === String(sesi.siswa_id)) }))
        .filter(r => r.siswa && (!state.filterNilaiKelas || (r.siswa.kelas || '').trim().toLowerCase() === state.filterNilaiKelas.trim().toLowerCase()));
      exportRekapNilaiExcel(jd, rows, state.filterNilaiKelas);
    });

    document.getElementById('btn-print-nilai-pdf')?.addEventListener('click', () => {
      const jd = state.db.jadwal_ujian.find(j => String(j.id) === String(state.filterNilaiJadwal));
      if (!jd) return;
      const rows = state.db.sesi_ujian_siswa
        .filter(s => String(s.jadwal_id) === String(jd.id))
        .map(sesi => ({ sesi, siswa: state.db.users_cbt.find(u => String(u.id) === String(sesi.siswa_id)) }))
        .filter(r => r.siswa && (!state.filterNilaiKelas || (r.siswa.kelas || '').trim().toLowerCase() === state.filterNilaiKelas.trim().toLowerCase()));
      const html = buildDaftarNilaiHtml(jd, rows, state.filterNilaiKelas);
      openA4PrintWindow('Daftar Nilai CBT', html);
    });

    // Pengaturan Sekolah & Upload Logo Permanen
    document.querySelectorAll('[data-upload-logo]').forEach(inp => {
      inp.addEventListener('change', async (e) => {
        const key = inp.getAttribute('data-upload-logo');
        const file = e.target.files?.[0];
        if (!file) return;
        state.modalLogos[key] = await compressImageToDataUrl(file, 260, 0.85);
        render();
      });
    });

    document.querySelectorAll('[data-clear-logo]').forEach(btn => {
      btn.addEventListener('click', () => {
        const key = btn.getAttribute('data-clear-logo');
        state.modalLogos[key] = '';
        render();
      });
    });

    document.getElementById('form-pengaturan-sekolah')?.addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const cfg = getSchoolConfig();
      const payload = {
        nama_sekolah: String(fd.get('nama_sekolah') || '').trim(),
        npsn: String(fd.get('npsn') || '').trim(),
        alamat: String(fd.get('alamat') || '').trim(),
        tahun_pelajaran: String(fd.get('tahun_pelajaran') || '').trim(),
        semester: String(fd.get('semester') || 'Ganjil'),
        nama_kepsek: String(fd.get('nama_kepsek') || '').trim(),
        nip_kepsek: String(fd.get('nip_kepsek') || '').trim(),
        nama_proktor: String(fd.get('nama_proktor') || '').trim(),
        nip_proktor: String(fd.get('nip_proktor') || '').trim(),
        nama_pengawas: String(fd.get('nama_pengawas') || '').trim(),
        nip_pengawas: String(fd.get('nip_pengawas') || '').trim(),
        max_pelanggaran: Number(fd.get('max_pelanggaran')) || 3,
        logo_dinas_url: state.modalLogos.dinas !== undefined ? state.modalLogos.dinas : cfg.logo_dinas_url,
        logo_sekolah_url: state.modalLogos.sekolah !== undefined ? state.modalLogos.sekolah : cfg.logo_sekolah_url,
        logo_aplikasi_url: state.modalLogos.aplikasi !== undefined ? state.modalLogos.aplikasi : cfg.logo_aplikasi_url
      };
      try {
        await saveSchoolConfigRecord(payload);
        state.modalLogos = { dinas: undefined, sekolah: undefined, aplikasi: undefined };
        showToast('Pengaturan identitas sekolah & logo berhasil disimpan permanen!');
        render();
      } catch (err) {
        showToast('Gagal menyimpan pengaturan: ' + (err.message || err), 'error');
      }
    });

    // Dashboard Siswa: Refresh Jadwal & Masuk Token Ujian
    document.getElementById('btn-refresh-siswa-jadwal')?.addEventListener('click', async () => {
      await loadAllData();
      showToast('Daftar jadwal ujian diperbarui.');
      render();
    });

    document.querySelectorAll('[data-open-token]').forEach(btn => {
      btn.addEventListener('click', () => {
        const jd = state.db.jadwal_ujian.find(j => j.id === btn.getAttribute('data-open-token'));
        if (jd) {
          state.modal = { type: 'token_exam', data: jd };
          render();
        }
      });
    });

    document.getElementById('form-enter-exam-token')?.addEventListener('submit', (e) => {
      e.preventDefault();
      const tokenVal = document.getElementById('input-student-token')?.value;
      startExamForJadwal(state.modal.data, tokenVal);
    });

    // =========================================================================
    // INTERAKSI RUANG UJIAN AKTIF SISWA (6 TIPE SOAL & NAVIGASI)
    // =========================================================================
    if (state.exam.active) {
      const curSoal = state.exam.soalList[state.exam.currentIndex];

      document.getElementById('btn-restore-fullscreen')?.addEventListener('click', () => {
        state.exam.graceUntil = Date.now() + 3000;
        state.exam.needsFullscreenTap = false;
        safeRequestFullscreen();
        render();
      });

      document.querySelectorAll('[data-exam-font]').forEach(btn => {
        btn.addEventListener('click', () => {
          state.exam.fontSize = btn.getAttribute('data-exam-font');
          render();
        });
      });

      document.getElementById('btn-toggle-daftar-soal')?.addEventListener('click', () => {
        state.exam.showDaftarSoalMobile = !state.exam.showDaftarSoalMobile;
        render();
      });

      document.querySelectorAll('[data-jump-soal]').forEach(btn => {
        btn.addEventListener('click', () => {
          state.exam.currentIndex = Number(btn.getAttribute('data-jump-soal')) || 0;
          state.exam.showDaftarSoalMobile = false;
          render();
        });
      });

      // 1. Jawaban Pilihan Ganda
      document.querySelectorAll('[data-answer-pg]').forEach(btn => {
        btn.addEventListener('click', () => {
          state.exam.jawaban[curSoal.id] = btn.getAttribute('data-answer-pg');
          saveCurrentExamProgress();
          render();
        });
      });

      // 2. Jawaban Benar / Salah
      document.querySelectorAll('[data-answer-bs]').forEach(btn => {
        btn.addEventListener('click', () => {
          state.exam.jawaban[curSoal.id] = btn.getAttribute('data-answer-bs');
          saveCurrentExamProgress();
          render();
        });
      });

      // 3. Jawaban Pilihan Ganda Kompleks (Bisa pilih banyak opsi sekaligus)
      document.querySelectorAll('[data-answer-pgk]').forEach(btn => {
        btn.addEventListener('click', () => {
          const opt = String(btn.getAttribute('data-answer-pgk') || '').trim().toUpperCase();
          const currentList = normalizePgkAnswerArray(state.exam.jawaban[curSoal.id]);
          const nextSet = new Set(currentList);
          if (nextSet.has(opt)) nextSet.delete(opt);
          else nextSet.add(opt);
          state.exam.jawaban[curSoal.id] = [...nextSet];
          saveCurrentExamProgress();
          render();
        });
      });

      // 4. Jawaban Menjodohkan
      document.querySelectorAll('[data-answer-match]').forEach(sel => {
        sel.addEventListener('change', (e) => {
          const idx = String(sel.getAttribute('data-answer-match'));
          const existingObj = normalizeMatchAnswerObj(state.exam.jawaban[curSoal.id]);
          existingObj[idx] = e.target.value;
          state.exam.jawaban[curSoal.id] = existingObj;
          saveCurrentExamProgress();
          render();
        });
      });

      // 5. Jawaban Isian Singkat
      document.getElementById('input-isian-singkat')?.addEventListener('input', (e) => {
        state.exam.jawaban[curSoal.id] = e.target.value;
        saveCurrentExamProgress();
      });

      // 6. Jawaban Uraian / Esai
      document.getElementById('input-uraian-jawab')?.addEventListener('input', (e) => {
        state.exam.jawaban[curSoal.id] = e.target.value;
        const counterEl = document.getElementById('uraian-char-count');
        if (counterEl) counterEl.textContent = `${e.target.value.length} karakter`;
        saveCurrentExamProgress();
      });

      document.getElementById('chk-ragu-soal')?.addEventListener('change', (e) => {
        state.exam.ragu[curSoal.id] = e.target.checked;
        saveCurrentExamProgress();
        render();
      });

      document.getElementById('btn-prev-soal')?.addEventListener('click', () => {
        if (state.exam.currentIndex > 0) {
          state.exam.currentIndex--;
          render();
        }
      });

      document.getElementById('btn-next-soal')?.addEventListener('click', () => {
        if (state.exam.currentIndex < state.exam.soalList.length - 1) {
          state.exam.currentIndex++;
          render();
        }
      });

      document.getElementById('btn-finish-exam')?.addEventListener('click', () => {
        const total = state.exam.soalList.length;
        const answered = state.exam.soalList.filter(s => isSoalAnswered(s, state.exam.jawaban[s.id])).length;
        const unanswered = total - answered;
        openConfirmDialog({
          title: 'Konfirmasi Selesai Ujian',
          message: unanswered > 0
            ? `Masih ada ${unanswered} soal yang belum terjawab lengkap. Apakah Anda yakin ingin mengakhiri ujian dan mengumpulkan semua jawaban sekarang?`
            : `Semua ${total} soal telah terjawab. Apakah Anda yakin ingin menyelesaikan ujian dan mengumpulkan jawaban sekarang?`,
          confirmLabel: 'Ya, Kumpulkan Jawaban',
          variant: 'primary',
          onConfirm: () => finishExamAndSubmit(false)
        });
      });
    }
  }

  // =========================================================================
  // INISIALISASI APLIKASI & PEMBARUAN JAM WITA REALTIME
  // =========================================================================
  async function initApp() {
    setupAntiCheatListeners();
    if (!IS_PREVIEW && isSupabaseConfigured) {
      await loadAllData();
      setupRealtimeSync();
      setInterval(() => {
        syncFastOperationalData(false);
      }, 3000);
    }

    setInterval(() => {
      const nowStr = getWitaTimeString();
      const loginClock = document.getElementById('login-clock');
      const headerClock = document.getElementById('header-clock');
      if (loginClock) loginClock.textContent = nowStr;
      if (headerClock) headerClock.textContent = nowStr;
    }, 1000);

    render();
  }

  initApp();
</script>
</body>
</html>
