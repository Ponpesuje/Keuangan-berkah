import React, { useState, useEffect, useMemo } from 'react';
import {
  LayoutDashboard, Receipt, BookOpen, Settings, LogOut, 
  Plus, Printer, Download, UserCircle, Building2, Wallet, TrendingUp, 
  Lock, Trash2, Edit2, RefreshCw, ArrowLeft, FileText, Landmark, X, Check,
  AlertCircle, Filter, Shield, HelpCircle, FileSpreadsheet
} from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from 'recharts';
import ExcelJS from 'exceljs';

// URL Google Sheets Database (Sesuai dengan database Anda tanpa mengubah struktur)
const GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbwc5hKNMisIqk5Jp2fmLISH-bAkkjh6d23ZBptF35hizz4RJnh9oBdUD5PUDPnr9433/exec";

const INITIAL_COMPANY = {
  name: "PT Berkah Mulia Kacida",
  address: "Jl. Raya Buah Segar No. 88, Gedung Berkah Lantai 5, Jakarta Selatan",
  phone: "021-555-0199",
  email: "corporate@berkahmuliakacida.com",
  npwp: "01.234.567.8-901.000"
};

const INITIAL_ACCOUNTS = [
  { id: '101', name: 'Kas & Bank (Operasional)', type: 'asset' },
  { id: '102', name: 'Piutang Usaha Dagang', type: 'asset' },
  { id: '103', name: 'Persediaan Buah & Logistik', type: 'asset' },
  { id: '201', name: 'Hutang Usaha Suplier', type: 'liability' },
  { id: '301', name: 'Modal Saham Ditempatkan', type: 'equity' },
  { id: '401', name: 'Pendapatan Penjualan Komoditas', type: 'revenue' },
  { id: '501', name: 'Beban Pokok Penjualan (BPP)', type: 'expense' },
  { id: '502', name: 'Beban Umum & Operasional', type: 'expense' },
  { id: '503', name: 'Beban Gaji & Tunjangan Karyawan', type: 'expense' }
];

const formatCurrency = (amount) => {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 }).format(amount);
};

// Fungsi Universal Eksport Excel Bergaris Kotak Rapi (Grid)
const exportToExcelStyle = async (sheetName, columns, data, filename, companyName) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);

  // Judul Laporan Atas
  ws.mergeCells('A1:G1');
  ws.getCell('A1').value = companyName.toUpperCase();
  ws.getCell('A1').font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF1E293B' } };
  
  ws.mergeCells('A2:G2');
  ws.getCell('A2').value = `LAPORAN ${sheetName.toUpperCase()}`;
  ws.getCell('A2').font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF475569' } };
  ws.addRow([]); // Baris kosong

  // Set Kolom Header
  ws.getRow(4).values = columns.map(c => c.header);
  ws.columns = columns.map(c => ({ key: c.key, width: c.width || 15 }));
  
  // Masukkan Data
  data.forEach(item => ws.addRow(item));

  // Styling Header Tabel (Row 4)
  const headerRow = ws.getRow(4);
  headerRow.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } }; // Dark Slate
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

  // Grid / Kotak border tipis untuk seluruh sel data
  ws.eachRow((row, rowNumber) => {
    if (rowNumber >= 4) {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
        };
        cell.font = { name: 'Arial', size: 10 };
        if (typeof cell.value === 'number') {
          cell.numFmt = '#,##0';
          cell.alignment = { horizontal: 'right' };
        }
      });
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  window.URL.revokeObjectURL(url);
};

export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [currentView, setCurrentView] = useState('dashboard');
  const [isLoading, setIsLoading] = useState(false);
  const [syncStatus, setSyncStatus] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [selectedRole, setSelectedRole] = useState('admin');

  const [company, setCompany] = useState(INITIAL_COMPANY);
  const [accounts, setAccounts] = useState(INITIAL_ACCOUNTS);
  const [transactions, setTransactions] = useState([]);

  useEffect(() => {
    const loadData = async () => {
      setIsLoading(true);
      try {
        const res = await fetch(GOOGLE_SHEET_URL);
        const text = await res.text();
        if (text) {
          const data = JSON.parse(text);
          if (data.transactions && data.transactions.length > 0) setTransactions(data.transactions);
          if (data.accounts && data.accounts.length > 0) setAccounts(data.accounts);
          if (data.company && data.company.name) setCompany(data.company);
        }
      } catch (e) { console.log("Gagal memuat data cloud. Menggunakan data lokal."); }
      setIsLoading(false);
    };
    loadData();
  }, []);

  const saveToCloud = async (newTransactions, newAccounts, newCompany) => {
    if (currentUser?.role !== 'admin') return; // Proteksi ganda level kode
    setSyncStatus('Sinkronisasi Cloud...');
    try {
      await fetch(GOOGLE_SHEET_URL, {
        method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          transactions: newTransactions || transactions,
          accounts: newAccounts || accounts,
          company: newCompany || company
        })
      });
      setSyncStatus('Sistem Terbuka & Sinkron ✓');
      setTimeout(() => setSyncStatus(''), 3000);
    } catch (e) { setSyncStatus('Gagal sinkronisasi data'); }
  };

  const handleLogin = (e) => {
    e.preventDefault();
    setLoginError('');
    if (selectedRole === 'admin') {
      if (loginPassword === 'admin123') {
        setCurrentUser({ username: 'Direksi Utama / Admin', role: 'admin' });
      } else {
        setLoginError('Sandi Rahasia Admin Salah!');
      }
    } else {
      setCurrentUser({ username: 'Dewan Komisaris / RUPS / Pajak', role: 'viewer' });
    }
  };

  const balances = useMemo(() => {
    const bals = {};
    accounts.forEach(a => bals[a.id] = 0);
    transactions.forEach(trx => {
      const debitAcc = accounts.find(a => a.id === trx.debitAccountId);
      const creditAcc = accounts.find(a => a.id === trx.creditAccountId);
      if (debitAcc) {
        if (debitAcc.type === 'asset' || debitAcc.type === 'expense') bals[trx.debitAccountId] += trx.amount;
        else bals[trx.debitAccountId] -= trx.amount;
      }
      if (creditAcc) {
        if (creditAcc.type === 'liability' || creditAcc.type === 'equity' || creditAcc.type === 'revenue') bals[trx.creditAccountId] += trx.amount;
        else bals[trx.creditAccountId] -= trx.amount;
      }
    });
    return bals;
  }, [transactions, accounts]);

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-950 flex items-center justify-center p-4">
        <div className="w-full max-w-lg bg-slate-900 border-2 border-slate-800 rounded-3xl shadow-2xl overflow-hidden backdrop-blur-md">
          <div className="p-8 text-center bg-gradient-to-b from-slate-800/80 to-slate-900 border-b border-slate-800">
            <Building2 size={44} className="mx-auto text-amber-500 mb-3" />
            <h1 className="text-2xl font-black text-white tracking-wide uppercase">{company.name}</h1>
            <p className="text-slate-400 text-xs tracking-widest uppercase mt-1">Sistem Laporan Keuangan Eksekutif RUPS</p>
          </div>
          <form onSubmit={handleLogin} className="p-8 space-y-6">
            <div>
              <label className="block text-xs font-bold uppercase text-slate-400 mb-2 tracking-wider">Pilih Hak Akses Jabatan</label>
              <div className="grid grid-cols-2 gap-3">
                <button type="button" onClick={() => { setSelectedRole('admin'); setLoginError(''); }} className={`p-4 rounded-xl font-bold text-sm transition-all border flex flex-col items-center gap-2 ${selectedRole === 'admin' ? 'bg-emerald-600 border-emerald-500 text-white shadow-lg' : 'bg-slate-950 border-slate-800 text-slate-400'}`}>
                  <Shield size={18}/> Dewan Direksi (Admin)
                </button>
                <button type="button" onClick={() => { setSelectedRole('viewer'); setLoginError(''); }} className={`p-4 rounded-xl font-bold text-sm transition-all border flex flex-col items-center gap-2 ${selectedRole === 'viewer' ? 'bg-amber-600 border-amber-500 text-white shadow-lg' : 'bg-slate-950 border-slate-800 text-slate-400'}`}>
                  <UserCircle size={18}/> Komite RUPS / Viewer
                </button>
              </div>
            </div>

            {selectedRole === 'admin' && (
              <div>
                <label className="block text-xs font-bold uppercase text-slate-400 mb-2 tracking-wider">Kunci Sandi Otentikasi</label>
                <input type="password" placeholder="Masukkan sandi khusus admin..." value={loginPassword} onChange={e => setLoginPassword(e.target.value)} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3.5 text-white placeholder-slate-600 focus:outline-none focus:border-emerald-500 font-mono tracking-widest text-center" required />
              </div>
            )}

            {loginError && <div className="p-3 bg-rose-500/10 border border-rose-500/30 text-rose-400 text-xs font-semibold rounded-lg text-center">{loginError}</div>}

            <button type="submit" className="w-full bg-gradient-to-r from-slate-100 to-white text-slate-950 p-4 rounded-xl font-black uppercase tracking-wider hover:opacity-90 active:scale-[0.99] transition-all shadow-xl">
              Buka Portal Keuangan
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-950 text-slate-200 flex print:bg-white print:text-black">
      {/* SIDEBAR */}
      <aside className="w-64 bg-slate-900 border-r border-slate-800 flex flex-col print:hidden">
        <div className="p-6 border-b border-slate-800 bg-slate-950/40">
          <div className="flex items-center gap-2 text-amber-500 font-black tracking-wider uppercase text-sm">
            <Building2 size={18} /> <span>{company.name}</span>
          </div>
          <span className="text-[10px] bg-slate-800 px-2 py-0.5 rounded-full mt-2 inline-block font-mono text-slate-400 capitalize border border-slate-700">Akses: {currentUser.role}</span>
        </div>
        <nav className="flex-1 p-4 space-y-1.5 overflow-y-auto">
          <NavItem icon={<LayoutDashboard size={18}/>} label="Dashboard Korporat" active={currentView === 'dashboard'} onClick={() => setCurrentView('dashboard')} />
          <NavItem icon={<Receipt size={18}/>} label="Input Kas Harian" active={currentView === 'transactions'} onClick={() => setCurrentView('transactions')} />
          <NavItem icon={<Wallet size={18}/>} label="Buku Kas Besar" active={currentView === 'buku-kas'} onClick={() => setCurrentView('buku-kas')} />
          <NavItem icon={<BookOpen size={18}/>} label="Buku Jurnal Umum" active={currentView === 'journal'} onClick={() => setCurrentView('journal')} />
          <NavItem icon={<TrendingUp size={18}/>} label="Laporan Laba Rugi" active={currentView === 'laba-rugi'} onClick={() => setCurrentView('laba-rugi')} />
          <NavItem icon={<FileText size={18}/>} label="Laporan Neraca" active={currentView === 'neraca'} onClick={() => setCurrentView('neraca')} />
          <NavItem icon={<HelpCircle size={18}/>} label="Catatan CALK" active={currentView === 'calk'} onClick={() => setCurrentView('calk')} />
          <NavItem icon={<Landmark size={18}/>} label="Kepatuhan Pajak" active={currentView === 'laporan-pajak'} onClick={() => setCurrentView('laporan-pajak')} />
          {currentUser.role === 'admin' && <NavItem icon={<Settings size={18}/>} label="Konfigurasi PT" active={currentView === 'settings'} onClick={() => setCurrentView('settings')} />}
        </nav>
        <div className="p-4 border-t border-slate-800 bg-slate-950/20 space-y-3">
          {syncStatus && <div className="text-[11px] text-emerald-400 flex items-center gap-2 font-medium bg-emerald-500/5 p-2 rounded border border-emerald-500/10"><RefreshCw size={12} className="animate-spin" /> {syncStatus}</div>}
          <button onClick={() => { setCurrentUser(null); setLoginPassword(''); }} className="w-full flex items-center justify-center gap-2 p-2.5 text-xs font-bold text-rose-400 hover:bg-rose-500/10 rounded-lg transition-colors border border-rose-500/10">
            <LogOut size={14} /> Keluar Sistem
          </button>
        </div>
      </aside>

      {/* KONTEN UTAMA */}
      <main className="flex-1 flex flex-col overflow-hidden">
        <header className="h-16 bg-slate-900 border-b border-slate-800 flex items-center justify-between px-8 print:hidden">
          <h2 className="text-md font-bold uppercase tracking-wider text-white flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-emerald-500"></span> {currentView.replace('-', ' ')}</h2>
          <div className="text-xs text-slate-400 font-medium">Sidang Keuangan Aktual PT Berkah Mulia Kacida</div>
        </header>
        <div className="flex-1 overflow-y-auto p-4 md:p-8 bg-slate-950 print:p-0 print:bg-white">
          {currentView === 'dashboard' && <DashboardView transactions={transactions} accounts={accounts} balances={balances} />}
          {currentView === 'transactions' && <TransactionsView transactions={transactions} accounts={accounts} setTransactions={setTransactions} company={company} saveToCloud={saveToCloud} user={currentUser} />}
          {currentView === 'buku-kas' && <BukuKasView transactions={transactions} accounts={accounts} company={company} />}
          {currentView === 'journal' && <JournalView transactions={transactions} accounts={accounts} company={company} />}
          {currentView === 'laba-rugi' && <LabaRugiView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'neraca' && <NeracaView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'calk' && <CalkView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'laporan-pajak' && <LaporanPajakView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'settings' && <SettingsView company={company} setCompany={setCompany} accounts={accounts} saveToCloud={saveToCloud} user={currentUser} />}
        </div>
      </main>
    </div>
  );
}

function NavItem({ icon, label, active, onClick }) {
  return (
    <button onClick={onClick} className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all text-xs font-bold uppercase tracking-wider ${active ? 'bg-gradient-to-r from-emerald-600 to-teal-600 text-white shadow-lg shadow-emerald-950/50' : 'text-slate-400 hover:bg-slate-800/60 hover:text-slate-200'}`}>
      {icon} {label}
    </button>
  );
}

/* =======================================================
   1. DASHBOARD VIEW (PREMIUM TYPOGRAPHY & INTERACTIVE CHARTS)
   ======================================================= */
function DashboardView({ transactions, accounts, balances }) {
  const cashAccounts = accounts.filter(a => a.name.toLowerCase().includes('kas') || a.name.toLowerCase().includes('bank'));
  const totalCash = cashAccounts.reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const totalRevenue = accounts.filter(a => a.type === 'revenue').reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const totalExpense = accounts.filter(a => a.type === 'expense').reduce((sum, a) => sum + (balances[a.id] || 0), 0);

  const chartData = useMemo(() => {
    const daily = {};
    transactions.forEach(t => {
      if (!daily[t.date]) daily[t.date] = { date: t.date, Pemasukan: 0, Pengeluaran: 0 };
      if (t.debitAccountId === '101') daily[t.date].Pemasukan += t.amount;
      if (t.creditAccountId === '101') daily[t.date].Pengeluaran += t.amount;
    });
    return Object.values(daily).sort((a,b) => new Date(a.date) - new Date(b.date)).slice(-10);
  }, [transactions]);

  return (
    <div className="space-y-8 print:hidden animate-fade-in">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="bg-slate-900 border-l-4 border-amber-500 rounded-2xl p-6 shadow-xl">
          <p className="text-xs font-bold uppercase text-slate-400 tracking-widest">Likuiditas Konsolidasi (Kas & Bank)</p>
          <p className="text-2xl md:text-3xl font-black text-white mt-2 font-mono">{formatCurrency(totalCash)}</p>
        </div>
        <div className="bg-slate-900 border-l-4 border-emerald-500 rounded-2xl p-6 shadow-xl">
          <p className="text-xs font-bold uppercase text-slate-400 tracking-widest">Akumulasi Pendapatan Usaha</p>
          <p className="text-2xl md:text-3xl font-black text-white mt-2 font-mono">{formatCurrency(totalRevenue)}</p>
        </div>
        <div className="bg-slate-900 border-l-4 border-rose-500 rounded-2xl p-6 shadow-xl">
          <p className="text-xs font-bold uppercase text-slate-400 tracking-widest">Total Alokasi Beban Usaha</p>
          <p className="text-2xl md:text-3xl font-black text-white mt-2 font-mono">{formatCurrency(totalExpense)}</p>
        </div>
      </div>

      <div className="bg-slate-900 p-6 rounded-3xl border border-slate-800 shadow-2xl">
        <h3 className="text-sm font-bold uppercase text-slate-400 tracking-widest mb-6">Tren Aliran Finansial Perusahaan (10 Hari Terakhir)</h3>
        <div className="h-64 md:h-80 w-full">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
              <XAxis dataKey="date" stroke="#64748b" fontSize={11} tickLine={false} />
              <YAxis stroke="#64748b" fontSize={11} tickLine={false} tickFormatter={(v) => `Rp${v/1000000}M`} />
              <Tooltip cursor={{fill: '#1e293b'}} contentStyle={{backgroundColor: '#0f172a', borderRadius: '12px', border: '1px solid #334155', color: '#fff'}} />
              <Legend verticalAlign="top" height={36} iconType="circle" />
              <Bar name="Arus Masuk (Pemasukan)" dataKey="Pemasukan" fill="#10b981" radius={[6, 6, 0, 0]} />
              <Bar name="Arus Keluar (Pengeluaran)" dataKey="Pengeluaran" fill="#f43f5e" radius={[6, 6, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* =======================================================
   2. INPUT TRANSAKSI (SISTEM MEMBEDAKAN ROLE SECARA TOTAL)
   ======================================================= */
function TransactionsView({ transactions, accounts, setTransactions, company, saveToCloud, user }) {
  const [showForm, setShowForm] = useState(false);
  const [trxType, setTrxType] = useState('pemasukan'); 
  const [formData, setFormData] = useState({ date: new Date().toISOString().split('T')[0], description: '', amount: '', accountId: '' });

  const executeDelete = (id) => {
    if (user.role !== 'admin') return;
    if (window.confirm("Apakah Anda yakin ingin membatalkan/menghapus entri transaksi ini secara permanen?")) {
      const updated = transactions.filter(t => t.id !== id);
      setTransactions(updated);
      saveToCloud(updated, null, null);
    }
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    if (user.role !== 'admin') return alert("Akses Ditolak! Anda berada dalam mode Viewer.");
    const amountNum = parseFloat(formData.amount);
    if (!formData.accountId || isNaN(amountNum) || amountNum <= 0) return alert("Mohon lengkapi data dengan benar.");

    const newTrx = {
      id: `TRX-${Date.now()}`,
      date: formData.date,
      description: formData.description,
      amount: amountNum,
      debitAccountId: trxType === 'pemasukan' ? '101' : formData.accountId,
      creditAccountId: trxType === 'pemasukan' ? formData.accountId : '101'
    };

    const updated = [...transactions, newTrx];
    setTransactions(updated);
    saveToCloud(updated, null, null);
    setShowForm(false);
    setFormData({ date: new Date().toISOString().split('T')[0], description: '', amount: '', accountId: '' });
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h3 className="text-xl font-bold text-white uppercase tracking-wider">Lembar Pencatatan Transaksi</h3>
          <p className="text-slate-400 text-xs mt-1">Registrasi seluruh transaksi kas masuk dan keluar korporasi</p>
        </div>
        {user.role === 'admin' ? (
          <button onClick={() => setShowForm(!showForm)} className="w-full md:w-auto bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 px-5 py-2.5 rounded-xl text-white font-bold text-xs uppercase tracking-wider flex items-center justify-center gap-2 shadow-lg shadow-emerald-950">
            <Plus size={16} /> {showForm ? 'Sembunyikan Form' : 'Entri Transaksi Baru'}
          </button>
        ) : (
          <div className="flex items-center gap-2 bg-amber-500/10 border border-amber-500/20 text-amber-400 px-4 py-2 rounded-xl text-xs font-bold uppercase tracking-wider">
            <Lock size={14} /> Hak Akses: Lihat Saja (Viewer)
          </div>
        )}
      </div>

      {showForm && user.role === 'admin' && (
        <div className="bg-slate-900 border border-slate-800 p-6 md:p-8 rounded-3xl shadow-2xl max-w-2xl animate-fade-in">
          <div className="flex gap-4 mb-6">
            <button type="button" onClick={() => setTrxType('pemasukan')} className={`flex-1 py-3.5 rounded-xl font-black text-xs uppercase tracking-widest border transition-all ${trxType === 'pemasukan' ? 'bg-emerald-600 border-emerald-500 text-white shadow-lg' : 'bg-slate-950 border-slate-800 text-slate-500'}`}>Pemasukan Kas (+)</button>
            <button type="button" onClick={() => setTrxType('pengeluaran')} className={`flex-1 py-3.5 rounded-xl font-black text-xs uppercase tracking-widest border transition-all ${trxType === 'pengeluaran' ? 'bg-rose-600 border-rose-500 text-white shadow-lg' : 'bg-slate-950 border-slate-800 text-slate-500'}`}>Pengeluaran Kas (-)</button>
          </div>
          <form onSubmit={handleSubmit} className="space-y-5">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-xs text-slate-400 font-bold uppercase tracking-wider mb-1">Tanggal Transaksi</label>
                <input type="date" value={formData.date} onChange={e => setFormData({...formData, date: e.target.value})} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-white focus:outline-none focus:border-emerald-500" required />
              </div>
              <div>
                <label className="block text-xs text-slate-400 font-bold uppercase tracking-wider mb-1">Nominal (Rupiah)</label>
                <input type="number" placeholder="Rp 0" value={formData.amount} onChange={e => setFormData({...formData, amount: e.target.value})} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-white focus:outline-none focus:border-emerald-500 font-mono" required />
              </div>
            </div>
            <div>
              <label className="block text-xs text-slate-400 font-bold uppercase tracking-wider mb-1">Uraian / Deskripsi Lengkap</label>
              <input type="text" placeholder="Tulis rincian deskripsi transaksi formal PT..." value={formData.description} onChange={e => setFormData({...formData, description: e.target.value})} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-white focus:outline-none focus:border-emerald-500" required />
            </div>
            <div>
              <label className="block text-xs text-slate-400 font-bold uppercase tracking-wider mb-1">Klasifikasi Akun Lawan</label>
              <select value={formData.accountId} onChange={e => setFormData({...formData, accountId: e.target.value})} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-white focus:outline-none focus:border-emerald-500 text-sm" required>
                <option value="">-- Hubungkan Ke Bagan Akun Lawan --</option>
                {accounts.filter(a => a.id !== '101').map(a => <option key={a.id} value={a.id}>{a.id} - {a.name} ({a.type})</option>)}
              </select>
            </div>
            <button type="submit" className="w-full bg-gradient-to-r from-emerald-600 to-teal-600 text-white py-3.5 rounded-xl font-black uppercase tracking-wider shadow-lg">Simpan Ke Buku Besar Cloud</button>
          </form>
        </div>
      )}

      <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-x-auto shadow-2xl">
        <table className="w-full text-left text-xs border-collapse min-w-[700px]">
          <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
            <tr><th className="p-4 uppercase font-bold tracking-wider">Tanggal</th><th className="p-4 uppercase font-bold tracking-wider">No. Bukti</th><th className="p-4 uppercase font-bold tracking-wider">Keterangan</th><th className="p-4 uppercase font-bold tracking-wider">Akun Debit</th><th className="p-4 uppercase font-bold tracking-wider">Akun Kredit</th><th className="p-4 uppercase font-bold tracking-wider text-right">Nominal</th>{user.role === 'admin' && <th className="p-4 uppercase font-bold tracking-wider text-center">Aksi</th>}</tr>
          </thead>
          <tbody className="divide-y divide-slate-800/40">
            {[...transactions].reverse().map(t => (
              <tr key={t.id} className="hover:bg-slate-800/40 transition-colors">
                <td className="p-4 text-slate-300 font-medium whitespace-nowrap">{t.date}</td>
                <td className="p-4 font-mono text-slate-500 text-[10px]">{t.id}</td>
                <td className="p-4 text-white font-medium">{t.description}</td>
                <td className="p-4 text-emerald-400 font-semibold">{accounts.find(a=>a.id===t.debitAccountId)?.name}</td>
                <td className="p-4 text-rose-400 font-semibold">{accounts.find(a=>a.id===t.creditAccountId)?.name}</td>
                <td className="p-4 text-right font-bold text-white font-mono whitespace-nowrap">{formatCurrency(t.amount)}</td>
                {user.role === 'admin' && (
                  <td className="p-4 text-center">
                    <button onClick={() => executeDelete(t.id)} className="text-rose-500 hover:bg-rose-500/10 p-2 rounded-lg transition-colors"><Trash2 size={15}/></button>
                  </td>
                )}
              </tr>
            ))}
            {transactions.length === 0 && <tr><td colSpan="7" className="p-8 text-center text-slate-600 font-medium italic">Belum ada historis transaksi keuangan korporasi.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* =======================================================
   3. BUKU KAS BESAR (DENGAN FILTER BULAN & EXCELJS GRID)
   ======================================================= */
function BukuKasView({ transactions, accounts, company }) {
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');

  let balance = 0;
  const ledger = useMemo(() => {
    const list = [];
    [...transactions].sort((a,b)=>new Date(a.date)-new Date(b.date)).forEach(t => {
      if (t.debitAccountId === '101') {
        balance += t.amount;
        list.push({ date: t.date, id: t.id, desc: t.description, ref: accounts.find(a=>a.id===t.creditAccountId)?.name, masuk: t.amount, keluar: 0, saldo: balance });
      }
      if (t.creditAccountId === '101') {
        balance -= t.amount;
        list.push({ date: t.date, id: t.id, desc: t.description, ref: accounts.find(a=>a.id===t.debitAccountId)?.name, masuk: 0, keluar: t.amount, saldo: balance });
      }
    });
    return list;
  }, [transactions, accounts]);

  const filtered = ledger.filter(e => {
    if (startDate && e.date < startDate) return false;
    if (endDate && e.date > endDate) return false;
    return true;
  });

  const downloadExcel = () => {
    const cols = [
      { header: 'Tanggal', key: 'date', width: 14 },
      { header: 'ID Jurnal', key: 'id', width: 22 },
      { header: 'Uraian Deskripsi', key: 'desc', width: 35 },
      { header: 'Ref Akun Lawan', key: 'ref', width: 26 },
      { header: 'Debet (Masuk)', key: 'masuk', width: 18 },
      { header: 'Kredit (Keluar)', key: 'keluar', width: 18 },
      { header: 'Saldo Kumulatif', key: 'saldo', width: 20 }
    ];
    exportToExcelStyle('Buku Kas Besar', cols, filtered, `BUKU_KAS_BESAR_${startDate||'SEMUA'}_SD_${endDate||'AKHIR'}.xlsx`, company.name);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row gap-4 items-start md:items-center bg-slate-900 p-5 rounded-2xl border border-slate-800 print:hidden shadow-xl">
        <div className="flex items-center gap-2">
          <Filter size={18} className="text-amber-500" />
          <span className="text-xs font-bold uppercase text-slate-400">Rentang Buku Kas:</span>
        </div>
        <div className="flex items-center gap-2">
          <input type="date" className="bg-slate-950 border border-slate-800 p-2 rounded-xl text-xs text-white" value={startDate} onChange={e => setStartDate(e.target.value)} />
          <span className="text-slate-500 text-xs">s.d.</span>
          <input type="date" className="bg-slate-950 border border-slate-800 p-2 rounded-xl text-xs text-white" value={endDate} onChange={e => setEndDate(e.target.value)} />
        </div>
        <button onClick={downloadExcel} className="w-full md:w-auto md:ml-auto bg-blue-600 hover:bg-blue-500 px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider text-white flex items-center justify-center gap-2 shadow-lg shadow-blue-950">
          <FileSpreadsheet size={15}/> Export Excel Grid Rapi
        </button>
      </div>

      <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-x-auto shadow-2xl">
        <table className="w-full text-left text-xs min-w-[700px]">
          <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
            <tr><th className="p-4">Tanggal</th><th className="p-4">Keterangan / Ref Akun</th><th className="p-4 text-right">Debet (Masuk)</th><th className="p-4 text-right">Kredit (Keluar)</th><th className="p-4 text-right">Saldo Kas</th></tr>
          </thead>
          <tbody className="divide-y divide-slate-800/30">
            {[...filtered].reverse().map((e, idx) => (
              <tr key={idx} className="hover:bg-slate-800/30">
                <td className="p-4 font-medium text-slate-300">{e.date}</td>
                <td className="p-4 text-white font-medium">{e.desc} <span className="text-[10px] text-slate-500 block font-normal mt-0.5">Lawan: {e.ref}</span></td>
                <td className="p-4 text-right text-emerald-400 font-mono font-bold whitespace-nowrap">{e.masuk > 0 ? formatCurrency(e.masuk) : '-'}</td>
                <td className="p-4 text-right text-rose-400 font-mono font-bold whitespace-nowrap">{e.keluar > 0 ? formatCurrency(e.keluar) : '-'}</td>
                <td className="p-4 text-right text-white font-mono font-black whitespace-nowrap">{formatCurrency(e.saldo)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* =======================================================
   4. BUKU JURNAL UMUM (DOUBLE ENTRY SYSTEM WITH FILTER)
   ======================================================= */
function JournalView({ transactions, accounts, company }) {
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');

  const filtered = transactions.filter(t => {
    if (startDate && t.date < startDate) return false;
    if (endDate && t.date > endDate) return false;
    return true;
  });

  const downloadExcel = () => {
    const rows = [];
    [...filtered].sort((a,b)=>new Date(a.date)-new Date(b.date)).forEach(t => {
      rows.push({ date: t.date, id: t.id, desc: t.description, account: accounts.find(a=>a.id===t.debitAccountId)?.name, debit: t.amount, kredit: 0 });
      rows.push({ date: '', id: '', desc: '', account: `   ${accounts.find(a=>a.id===t.creditAccountId)?.name}`, debit: 0, kredit: t.amount });
    });

    const cols = [
      { header: 'Tanggal', key: 'date', width: 14 },
      { header: 'No Jurnal', key: 'id', width: 20 },
      { header: 'Keterangan Siklus', key: 'desc', width: 30 },
      { header: 'Akun Perkiraan', key: 'account', width: 30 },
      { header: 'Debit (Rp)', key: 'debit', width: 18 },
      { header: 'Kredit (Rp)', key: 'kredit', width: 18 }
    ];
    exportToExcelStyle('Jurnal Umum', cols, rows, `JURNAL_UMUM_${startDate||'SEMUA'}_SD_${endDate||'AKHIR'}.xlsx`, company.name);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row gap-4 items-start md:items-center bg-slate-900 p-5 rounded-2xl border border-slate-800 print:hidden shadow-xl">
        <div className="flex items-center gap-2">
          <Filter size={18} className="text-emerald-500" />
          <span className="text-xs font-bold uppercase text-slate-400">Rentang Jurnal:</span>
        </div>
        <div className="flex items-center gap-2">
          <input type="date" className="bg-slate-950 border border-slate-800 p-2 rounded-xl text-xs text-white" value={startDate} onChange={e => setStartDate(e.target.value)} />
          <span className="text-slate-500 text-xs">s.d.</span>
          <input type="date" className="bg-slate-950 border border-slate-800 p-2 rounded-xl text-xs text-white" value={endDate} onChange={e => setEndDate(e.target.value)} />
        </div>
        <button onClick={downloadExcel} className="w-full md:w-auto md:ml-auto bg-blue-600 hover:bg-blue-500 px-4 py-2 rounded-xl text-xs font-black uppercase tracking-wider text-white flex items-center justify-center gap-2 shadow-lg shadow-blue-950">
          <FileSpreadsheet size={15}/> Export Jurnal Grid
        </button>
      </div>

      <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-x-auto shadow-2xl">
        <table className="w-full text-left text-xs border-collapse min-w-[600px]">
          <thead className="bg-slate-950 text-slate-400">
            <tr><th className="p-4">Tanggal Jurnal</th><th className="p-4">Kode Perkiraan & Akun Akuntansi</th><th className="p-4 text-right">Debit</th><th className="p-4 text-right">Kredit</th></tr>
          </thead>
          <tbody className="divide-y divide-slate-800/40">
            {[...filtered].reverse().map((t, idx) => (
              <React.Fragment key={idx}>
                <tr className="bg-slate-900/60">
                  <td className="p-4 text-slate-400 font-medium align-top whitespace-nowrap" rowSpan="2">{t.date}<br/><span className="text-[9px] font-mono block text-slate-600 mt-1">{t.id}</span></td>
                  <td className="p-4 text-white font-bold tracking-wide">{t.debitAccountId} - {accounts.find(a=>a.id===t.debitAccountId)?.name}<span className="block text-[10px] text-slate-400 font-normal italic mt-0.5">Memo: {t.description}</span></td>
                  <td className="p-4 text-right text-emerald-400 font-mono font-bold text-sm whitespace-nowrap">{formatCurrency(t.amount)}</td>
                  <td className="p-4 text-right text-slate-600 font-mono">-</td>
                </tr>
                <tr className="border-b border-slate-800/40 bg-slate-900/10">
                  <td className="p-4 text-slate-300 font-medium italic pl-12">{t.creditAccountId} - {accounts.find(a=>a.id===t.creditAccountId)?.name}</td>
                  <td className="p-4 text-right text-slate-600 font-mono">-</td>
                  <td className="p-4 text-right text-rose-400 font-mono font-bold text-sm whitespace-nowrap">{formatCurrency(t.amount)}</td>
                </tr>
              </React.Fragment>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* =======================================================
   5. LAPORAN LABA RUGI FISKAL & KOMERSIAL (RUPS READY)
   ======================================================= */
function LabaRugiView({ accounts, balances, company }) {
  const revAccounts = accounts.filter(a => a.type === 'revenue');
  const expAccounts = accounts.filter(a => a.type === 'expense');

  const totalRev = revAccounts.reduce((s, a) => s + (balances[a.id] || 0), 0);
  const totalExp = expAccounts.reduce((s, a) => s + (balances[a.id] || 0), 0);
  const netProfit = totalRev - totalExp;

  return (
    <div className="space-y-6">
      <div className="flex justify-end print:hidden"><button onClick={() => window.print()} className="bg-slate-800 border border-slate-700 px-4 py-2 rounded-xl text-xs font-bold uppercase tracking-wider text-slate-200 flex items-center gap-2"><Printer size={14}/> Cetak Dokumen RUPS</button></div>
      <div className="bg-white text-slate-900 p-8 md:p-12 rounded-3xl max-w-4xl mx-auto shadow-2xl border border-slate-200 min-h-[750px] font-sans">
        <div className="text-center border-b-2 border-slate-900 pb-6 mb-8">
          <h2 className="text-xl font-black tracking-wide text-slate-900 uppercase">{company.name}</h2>
          <h3 className="text-md font-bold mt-1 text-slate-600 uppercase tracking-widest">Laporan Laba Rugi Komersial</h3>
          <p className="text-xs text-slate-400 font-mono mt-2">Periode Berakhir Per: {new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
        </div>

        <div className="space-y-6 text-sm">
          <div>
            <h4 className="font-bold text-slate-900 border-b border-slate-300 pb-1.5 uppercase text-xs tracking-wider">1. PEREDARAN USAHA (PENDAPATAN)</h4>
            <div className="space-y-2 mt-3 pl-4">
              {revAccounts.map(a => (
                <div key={a.id} className="flex justify-between text-slate-700 font-medium"><span>{a.name}</span><span className="font-mono whitespace-nowrap">{formatCurrency(balances[a.id] || 0)}</span></div>
              ))}
            </div>
            <div className="flex justify-between font-bold mt-4 pt-2 border-t border-slate-900 text-slate-900"><span>TOTAL PENDAPATAN OPERASIONAL</span><span className="font-mono whitespace-nowrap">{formatCurrency(totalRev)}</span></div>
          </div>

          <div>
            <h4 className="font-bold text-slate-900 border-b border-slate-300 pb-1.5 uppercase text-xs tracking-wider">2. BEBAN OPERASIONAL & BPP</h4>
            <div className="space-y-2 mt-3 pl-4">
              {expAccounts.map(a => (
                <div key={a.id} className="flex justify-between text-slate-600 font-medium"><span>{a.name}</span><span className="font-mono whitespace-nowrap">({formatCurrency(balances[a.id] || 0)})</span></div>
              ))}
            </div>
            <div className="flex justify-between font-bold mt-4 pt-2 border-t border-slate-900 text-slate-900"><span>TOTAL BEBAN KONSOLIDASI</span><span className="font-mono whitespace-nowrap">({formatCurrency(totalExp)})</span></div>
          </div>

          <div className="pt-6 border-t-4 border-double border-slate-900 flex justify-between font-black text-sm md:text-lg text-slate-900 bg-slate-50 p-3 rounded-xl">
            <span>LABA BERSIH TAHUN BERJALAN</span>
            <span className={`font-mono whitespace-nowrap ${netProfit < 0 ? 'text-rose-600' : 'text-emerald-700'}`}>{formatCurrency(netProfit)}</span>
          </div>
        </div>

        <div className="mt-20 grid grid-cols-2 gap-4 md:gap-8 text-center text-[10px] md:text-xs font-medium">
          <div><p className="mb-16 md:mb-20 text-slate-500">Dibuat & Dipertanggungjawabkan,</p><p className="font-bold border-b border-slate-900 inline-block px-4 md:px-10 pb-1 text-slate-900">DIREKSI UTAMA</p></div>
          <div><p className="mb-16 md:mb-20 text-slate-500">Disetujui & Disahkan,</p><p className="font-bold border-b border-slate-900 inline-block px-4 md:px-10 pb-1 text-slate-900">DEWAN KOMISARIS / RUPS</p></div>
        </div>
      </div>
    </div>
  );
}

/* =======================================================
   6. LAPORAN NERACA POSISI KEUANGAN FORMAL
   ======================================================= */
function NeracaView({ accounts, balances, company }) {
  const revTotal = accounts.filter(a => a.type === 'revenue').reduce((s, a) => s + (balances[a.id] || 0), 0);
  const expTotal = accounts.filter(a => a.type === 'expense').reduce((s, a) => s + (balances[a.id] || 0), 0);
  const netIncome = revTotal - expTotal;

  const assets = accounts.filter(a => a.type === 'asset');
  const liabilities = accounts.filter(a => a.type === 'liability');
  const equities = accounts.filter(a => a.type === 'equity');

  const totalAsset = assets.reduce((s, a) => s + (balances[a.id] || 0), 0);
  const totalLiab = liabilities.reduce((s, a) => s + (balances[a.id] || 0), 0);
  const totalEqBase = equities.reduce((s, a) => s + (balances[a.id] || 0), 0);
  const totalPasiva = totalLiab + totalEqBase + netIncome;

  return (
    <div className="space-y-6">
      <div className="flex justify-end print:hidden"><button onClick={() => window.print()} className="bg-slate-800 border border-slate-700 px-4 py-2 rounded-xl text-xs font-bold uppercase tracking-wider text-slate-200 flex items-center gap-2"><Printer size={14}/> Cetak Dokumen RUPS</button></div>
      <div className="bg-white text-slate-900 p-8 md:p-12 rounded-3xl max-w-5xl mx-auto shadow-2xl border border-slate-200 min-h-[750px] font-sans">
        <div className="text-center border-b-2 border-slate-900 pb-6 mb-10">
          <h2 className="text-xl font-black text-slate-900 uppercase">{company.name}</h2>
          <h3 className="text-md font-bold mt-1 text-slate-600 uppercase tracking-widest">Laporan Neraca Posisi Keuangan</h3>
          <p className="text-xs text-slate-500 font-mono mt-1">Per Tanggal Aktual: {new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 text-sm">
          {/* SEKTOR AKTIVA */}
          <div className="space-y-4">
            <h4 className="font-bold bg-slate-900 text-white p-2 text-center uppercase text-xs tracking-wider rounded-md">AKTIVA (ASET PERUSAHAAN)</h4>
            <div className="space-y-2.5 pl-2">
              {assets.map(a => (
                <div key={a.id} className="flex justify-between border-b border-slate-100 pb-1 text-slate-700"><span>{a.name}</span><span className="font-mono whitespace-nowrap">{formatCurrency(balances[a.id] || 0)}</span></div>
              ))}
            </div>
            <div className="flex justify-between font-black pt-4 border-t-2 border-slate-900 text-slate-900 bg-slate-50 p-2 rounded-lg mt-6"><span>TOTAL AKTIVA KONSOLIDASI</span><span className="font-mono whitespace-nowrap">{formatCurrency(totalAsset)}</span></div>
          </div>

          {/* SEKTOR PASIVA */}
          <div className="space-y-6">
            <div>
              <h4 className="font-bold bg-slate-900 text-white p-2 text-center uppercase text-xs tracking-wider rounded-md mb-3">PASIVA (KEWAJIBAN & MODAL)</h4>
              <h5 className="font-bold text-slate-800 italic text-xs uppercase tracking-wider border-b mb-2">A. Kewajiban Jangka Pendek</h5>
              <div className="space-y-2 pl-2">
                {liabilities.map(a => (
                  <div key={a.id} className="flex justify-between border-b border-slate-100 pb-1 text-slate-700"><span>{a.name}</span><span className="font-mono whitespace-nowrap">{formatCurrency(balances[a.id] || 0)}</span></div>
                ))}
              </div>
            </div>
className
            <div>
              <h5 className="font-bold text-slate-800 italic text-xs uppercase tracking-wider border-b mb-2">B. Ekuitas (Modal PT)</h5>
              <div claclassNamessName="space-y-2 pl-2">
                {equities.map(a => (
                  <div key={a.id} className="flex justify-between text-slate-700"><span>{a.name}</span><span className="font-mono whitespace-nowrap">{formatCurrency(balances[a.id] || 0)}</span></div>
                ))}
                <div className="flex justify-between text-emerald-800 font-bold bg-emerald-50 px-2 py-0.5 rounded mt-2"><span>Laba Bersih Tahun Berjalan</span><span className="font-mono whitespace-nowrap">{formatCurrency(netIncome)}</span></div>
              </div>
            </div>

            <div className="flex justify-between font-black pt-4 border-t-2 border-slate-900 text-slate-900 bg-slate-50 p-2 rounded-lg mt-6"><span>TOTAL PASIVA KONSOLIDASI</span><span className="font-mono whitespace-nowrap">{formatCurrency(totalPasiva)}</span></div>
          </div>
        </div>

        <div className="mt-24 grid grid-cols-2 gap-4 md:gap-8 text-center text-[10px] md:text-xs font-medium">
          <div><p className="mb-16 md:mb-20 text-slate-500">Dilaporkan Oleh Manajemen,</p><p className="font-bold border-b border-slate-900 inline-block px-4 md:px-10 pb-1 text-slate-900">DIREKTUR KEUANGAN</p></div>
          <div><p className="mb-16 md:mb-20 text-slate-500">Mengesahkan Sidang RUPS,</p><p className="font-bold border-b border-slate-900 inline-block px-4 md:px-10 pb-1 text-slate-900">KOMISARIS UTAMA</p></div>
        </div>
      </div>
    </div>
  );
}

/* =======================================================
   7. CATATAN ATAS LAPORAN KEUANGAN (CALK - NEW REQUEST)
   ======================================================= */
function CalkView({ accounts, balances, company }) {
  return (
    <div className="space-y-6">
      <div className="flex justify-end print:hidden"><button onClick={() => window.print()} className="bg-slate-800 border border-slate-700 px-4 py-2 rounded-xl text-xs font-bold uppercase tracking-wider text-slate-200 flex items-center gap-2"><Printer size={14}/> Cetak CALK RUPS</button></div>
      <div className="bg-white text-slate-900 p-8 md:p-12 rounded-3xl max-w-4xl mx-auto shadow-2xl border border-slate-200 min-h-[750px] font-sans text-sm space-y-6">
        <div className="text-center border-b-2 border-slate-900 pb-4 mb-6">
          <h2 className="text-xl font-black text-slate-900 uppercase">{company.name}</h2>
          <h3 className="text-xs font-bold mt-1 text-slate-500 uppercase tracking-widest">Catatan Atas Laporan Keuangan (CALK)</h3>
          <p className="text-xs font-mono text-slate-400 mt-1">Tahun Buku Konstitusi Aktif</p>
        </div>

        <div>
          <h4 className="font-bold text-slate-900 uppercase text-xs tracking-wider mb-2">1. Gambaran Umum Perusahaan</h4>
          <p className="text-slate-600 text-justify leading-relaxed pl-4">
            {company.name} didirikan secara sah berdasarkan hukum perseroan terbatas di Republik Indonesia. Perusahaan bergerak dalam bidang perdagangan komoditas utama buah-buahan segar serta logistik rantai pasok pangan. Segala operasional berpusat di alamat {company.address}.
          </p>
        </div>

        <div>
          <h4 className="font-bold text-slate-900 uppercase text-xs tracking-wider mb-2">2. Ikhtisar Kebijakan Akuntansi Penting</h4>
          <p className="text-slate-600 text-justify leading-relaxed pl-4 mb-2">
            <strong>A. Dasar Penyusunan Laporan:</strong> Laporan Keuangan disusun berdasarkan Standar Akuntansi Keuangan Entitas Tanpa Akuntabilitas Publik (SAK ETAP). Sistem pencatatan menggunakan basis akrual penuh (*Full Accrual Basis*).
          </p>
          <p className="text-slate-600 text-justify leading-relaxed pl-4">
            <strong>B. Pengakuan Pendapatan:</strong> Pendapatan diakui pada saat penyerahan komoditas buah secara fisik kepada pembeli (klien) dan invoice telah tervalidasi.
          </p>
        </div>

        <div>
          <h4 className="font-bold text-slate-900 uppercase text-xs tracking-wider mb-3">3. Rincian Penjelasan Pos Laporan Keuangan Aktual</h4>
          <div className="space-y-2 pl-4 overflow-x-auto">
            {accounts.map(a => (
              <div key={a.id} className="flex justify-between border-b border-slate-100 py-1 text-xs font-medium text-slate-700 min-w-[300px]">
                <span>Catatan Perkiraan {a.id} - {a.name} ({a.type})</span>
                <span className="font-mono font-bold text-slate-900">{formatCurrency(balances[a.id] || 0)}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* =======================================================
   8. LAPORAN PAJAK UMKM TAHUNAN (PP NO. 55 TAHUN 2022)
   ======================================================= */
function LaporanPajakView({ accounts, balances, company }) {
  const totalRev = accounts.filter(a => a.type === 'revenue').reduce((s, a) => s + (balances[a.id] || 0), 0);
  const pphFinal = totalRev * 0.005; // 0.5% Berdasarkan PP 55 / 2022

  return (
    <div className="space-y-6">
      <div className="flex justify-end print:hidden"><button onClick={() => window.print()} className="bg-slate-800 border border-slate-700 px-4 py-2 rounded-xl text-xs font-bold uppercase tracking-wider text-slate-200 flex items-center gap-2"><Printer size={14}/> Cetak Salinan Pajak</button></div>
      <div className="bg-white text-slate-900 p-8 md:p-12 rounded-3xl max-w-4xl mx-auto shadow-2xl border border-slate-200 min-h-[700px]">
        <div className="flex flex-col md:flex-row justify-between items-start border-b-4 border-slate-900 pb-6 mb-8 gap-4">
          <div>
            <h2 className="text-lg font-black uppercase text-slate-900">{company.name}</h2>
            <p className="text-xs text-slate-600 mt-1 max-w-xs">{company.address}</p>
            <p className="text-xs font-mono font-bold text-slate-900 mt-2">NPWP BADAN: {company.npwp}</p>
          </div>
          <div className="text-left md:text-right">
            <span className="text-xs font-black bg-slate-100 border border-slate-300 p-2 rounded-md uppercase tracking-wider inline-block">Kepatuhan SPT Tahunan</span>
            <p className="text-xs font-mono mt-2 text-slate-500">Tahun Pajak Fiskal Aktif</p>
          </div>
        </div>

        <div className="bg-slate-50 p-4 border rounded-xl border-slate-200 text-center mb-8">
          <h4 className="font-bold text-slate-900 text-sm uppercase tracking-wide">Transkrip Perhitungan PPh Final Pajak Korporasi</h4>
          <p className="text-slate-500 text-xs mt-1">Sesuai Peraturan Pemerintah (PP) No. 55 Tahun 2022 mengenai Insentif Final PPh UMKM Badan Mandiri</p>
        </div>

        <div className="space-y-4 text-sm mt-8">
          <div className="flex justify-between border-b py-2 text-slate-700"><span>Dasar Pengenaan Pajak (Total Omzet Kotor Peredaran Usaha)</span><span className="font-mono font-bold whitespace-nowrap">{formatCurrency(totalRev)}</span></div>
          <div className="flex justify-between border-b py-2 text-slate-700"><span>Tarif Pajak Insentif PP55 Terikat</span><span className="font-mono font-bold text-emerald-700">0.5 % (Final)</span></div>
          <div className="flex justify-between font-black text-md bg-slate-100 p-4 rounded-xl text-slate-900 pt-3"><span>ESTIMASI PPh FINAL TERUTANG</span><span className="font-mono text-rose-600 text-lg whitespace-nowrap">{formatCurrency(pphFinal)}</span></div>
        </div>

        <div className="mt-28 text-right text-xs no-print print:block">
          <div className="w-full md:w-56 ml-auto text-center md:text-right">
            <p className="mb-16 md:text-center">Wajib Pajak Badan / Kuasa Direksi,</p>
            <p className="font-bold border-b border-slate-900 inline-block px-6 pb-1 uppercase text-slate-900">{company.name}</p>
            <p className="text-[10px] text-slate-400 mt-1 md:text-center">Cap Resmi Perusahaan</p>
          </div>
        </div>
      </div>
    </div>
  );
}

/* =======================================================
   9. KONFIGURASI PROFIL PT (DIKUNCI JIKA VIEW ER)
   ======================================================= */
function SettingsView({ company, setCompany, accounts, saveToCloud, user }) {
  const [name, setName] = useState(company.name);
  const [address, setAddress] = useState(company.address);
  const [phone, setPhone] = useState(company.phone);
  const [email, setEmail] = useState(company.email);
  const [npwp, setNpwp] = useState(company.npwp);

  const handleSave = (e) => {
    e.preventDefault();
    if(user.role !== 'admin') return alert("Akses ditolak!");
    const updated = { name, address, phone, email, npwp };
    setCompany(updated);
    saveToCloud(null, null, updated);
    alert("Profil korporasi PT berhasil disinkronkan ke Cloud!");
  };

  return (
    <div className="max-w-2xl bg-slate-900 p-6 md:p-8 rounded-3xl border border-slate-800 shadow-2xl animate-fade-in">
      <h3 className="text-lg font-bold text-white uppercase tracking-wider mb-6 border-b border-slate-800 pb-3">Profil Yuridis Perusahaan</h3>
      <form onSubmit={handleSave} className="space-y-4">
        <div><label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Nama Resmi PT</label><input type="text" className="w-full bg-slate-950 border border-slate-800 p-3 rounded-xl text-white text-sm" value={name} onChange={e=>setName(e.target.value)} required /></div>
        <div><label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Alamat Kantor Pusat</label><textarea className="w-full bg-slate-950 border border-slate-800 p-3 rounded-xl text-white text-sm h-20" value={address} onChange={e=>setAddress(e.target.value)} required /></div>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div><label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Telepon</label><input type="text" className="w-full bg-slate-950 border border-slate-800 p-3 rounded-xl text-white text-sm" value={phone} onChange={e=>setPhone(e.target.value)} /></div>
          <div><label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Email Resmi</label><input type="email" className="w-full bg-slate-950 border border-slate-800 p-3 rounded-xl text-white text-sm" value={email} onChange={e=>setEmail(e.target.value)} /></div>
        </div>
        <div><label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">NPWP Badan Korporasi</label><input type="text" className="w-full bg-slate-950 border border-slate-800 p-3 rounded-xl text-white text-sm font-mono" value={npwp} onChange={e=>setNpwp(e.target.value)} /></div>
        <button type="submit" className="w-full md:w-auto bg-emerald-600 hover:bg-emerald-500 px-6 py-3 rounded-xl font-black text-xs uppercase tracking-wider text-white shadow-lg mt-2">Simpan Perubahan Struktur</button>
      </form>
    </div>
  );
}