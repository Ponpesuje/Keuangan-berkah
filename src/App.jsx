import React, { useState, useEffect, useMemo } from 'react';
import {
  LayoutDashboard, Receipt, BookOpen, Settings, LogOut,
  Plus, Printer,FileSignature, Download, UserCircle, Building2, Wallet, TrendingUp,
  Lock, Trash2, FileText, Landmark, Filter, Shield, HelpCircle, FileSpreadsheet,
  ChevronRight, AlertCircle, RefreshCw, X
} from 'lucide-react';
import { AreaChart, Area, CartesianGrid, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend, PieChart, Pie, Cell, BarChart, Bar } from 'recharts';

/* ========================================================================
   KONFIGURASI UTAMA & DATA AWAL
   ======================================================================== */
const GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbwc5hKNMisIqk5Jp2fmLISH-bAkkjh6d23ZBptF35hizz4RJnh9oBdUD5PUDPnr9433/exec";

const INITIAL_COMPANY = {
  name: "PT BERKAH MULIA KACIDA",
  address: "Gedung Cyber, Lantai 5, Jl. Raya Buah Segar No. 88, Jakarta Selatan",
  phone: "021-555-0199",
  email: "corporate@berkahmuliakacida.com",
  npwp: "01.234.567.8-901.000",
  logo: ""
};

const INITIAL_ACCOUNTS = [
  { id: '101', name: 'Kas & Bank', type: 'asset' },
  { id: '102', name: 'Piutang Usaha', type: 'asset' },
  { id: '103', name: 'Persediaan Buah & Logistik', type: 'asset' },
  { id: '201', name: 'Hutang Usaha', type: 'liability' },
  { id: '301', name: 'Modal Saham Ditempatkan', type: 'equity' },
  { id: '401', name: 'Pendapatan Penjualan', type: 'revenue' },
  { id: '501', name: 'Beban Pokok Penjualan (BPP)', type: 'expense' },
  { id: '502', name: 'Beban Operasional', type: 'expense' },
  { id: '503', name: 'Beban Gaji & Tunjangan', type: 'expense' }
];

const formatCurrency = (amount) => {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 }).format(amount);
};

/* ========================================================================
   FUNGSI EXPORT DATA (CSV Fallback)
   ======================================================================== */
const exportGridToExcel = async (sheetName, columns, data, filename, companyName) => {
  let csvContent = `"${companyName.toUpperCase()}"\n`;
  csvContent += `"LAPORAN ${sheetName.toUpperCase()}"\n\n`;
  
  const headerRow = columns.map(c => `"${c.header}"`);
  csvContent += headerRow.join(",") + "\n";
  
  data.forEach(item => {
    const row = columns.map(c => {
      let val = item[c.key];
      return `"${val !== null && val !== undefined ? String(val).replace(/"/g, '""') : ''}"`;
    });
    csvContent += row.join(",") + "\n";
  });

  triggerDownload(csvContent, filename.replace('.xlsx', '.csv'), 'text/csv;charset=utf-8;');
};

const exportFinancialToExcel = async (title, companyName, rowsData, filename) => {
  let csvContent = `"${companyName.toUpperCase()}"\n`;
  csvContent += `"LAPORAN ${title.toUpperCase()}"\n\n`;
  
  rowsData.forEach(r => {
    let label = r.label || '';
    let amount = r.amount !== null ? r.amount : '';
    csvContent += `"${String(label).replace(/"/g, '""')}","${String(amount).replace(/"/g, '""')}"\n`;
  });

  triggerDownload(csvContent, filename.replace('.xlsx', '.csv'), 'text/csv;charset=utf-8;');
};

const triggerDownload = (content, filename, mimeType) => {
  const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), content], { type: mimeType });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  window.URL.revokeObjectURL(url);
};

/* ========================================================================
   KOMPONEN KOP SURAT (PRINT HEADER)
   ======================================================================== */
const PrintLetterhead = ({ company }) => (
  <div className="border-b-[5px] border-double border-slate-900 pb-6 mb-8 flex items-center justify-between">
    <div className="w-32 flex justify-start">
      {company.logo && <img src={company.logo} alt="Logo Perusahaan" className="max-h-24 max-w-full object-contain" />}
    </div>
    <div className="flex-1 text-center px-4">
      <h1 className="text-2xl md:text-3xl font-black text-slate-900 tracking-widest uppercase">{company.name}</h1>
      <p className="text-xs md:text-sm text-slate-800 mt-1">{company.address}</p>
      <p className="text-xs md:text-sm text-slate-800">Telepon: {company.phone} | Email: {company.email}</p>
      <p className="text-xs md:text-sm font-bold text-slate-900 mt-1">NPWP: {company.npwp}</p>
    </div>
    <div className="w-32"></div> {/* Spacer untuk keseimbangan center */}
  </div>
);

/* ========================================================================
   KOMPONEN UTAMA (APP & AUTHENTICATION)
   ======================================================================== */
export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [currentView, setCurrentView] = useState('dashboard');
  const [syncStatus, setSyncStatus] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [selectedRole, setSelectedRole] = useState('admin');

  const [company, setCompany] = useState(INITIAL_COMPANY);
  const [accounts, setAccounts] = useState(INITIAL_ACCOUNTS);
  const [transactions, setTransactions] = useState([]);

  useEffect(() => {
    const loadData = async () => {
      try {
        const res = await fetch(GOOGLE_SHEET_URL);
        const data = await res.json();
        if (data.transactions) setTransactions(data.transactions);
        if (data.accounts) setAccounts(data.accounts);
        if (data.company) setCompany(data.company);
      } catch (e) { console.log("Menggunakan data lokal."); }
    };
    loadData();
  }, []);

  const saveToCloud = async (newTransactions, newAccounts, newCompany) => {
    if (currentUser?.role !== 'admin') return;
    setSyncStatus('Sinkronisasi...');
    try {
      await fetch(GOOGLE_SHEET_URL, {
        method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          transactions: newTransactions || transactions,
          accounts: newAccounts || accounts,
          company: newCompany || company
        })
      });
      setSyncStatus('Cloud Tersinkron ✓');
      setTimeout(() => setSyncStatus(''), 3000);
    } catch (e) { setSyncStatus('Gagal Sync'); }
  };

  const handleLogin = (e) => {
    e.preventDefault();
    setLoginError('');
    if (selectedRole === 'admin') {
      if (loginPassword === 'admin123') {
        setCurrentUser({ username: 'Direktur Utama', role: 'admin' });
      } else {
        setLoginError('Sandi Rahasia Salah!');
      }
    } else {
      setCurrentUser({ username: 'Dewan Komisaris / RUPS', role: 'viewer' });
    }
  };

  const balances = useMemo(() => {
    const bals = {};
    accounts.forEach(a => bals[a.id] = 0);
    transactions.forEach(trx => {
      const dbAcc = accounts.find(a => a.id === trx.debitAccountId);
      const crAcc = accounts.find(a => a.id === trx.creditAccountId);
      if (dbAcc) {
        if (dbAcc.type === 'asset' || dbAcc.type === 'expense') bals[trx.debitAccountId] += trx.amount;
        else bals[trx.debitAccountId] -= trx.amount;
      }
      if (crAcc) {
        if (crAcc.type === 'liability' || crAcc.type === 'equity' || crAcc.type === 'revenue') bals[trx.creditAccountId] += trx.amount;
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
            <p className="text-slate-400 text-xs tracking-widest uppercase mt-1">Enterprise Financial Portal</p>
          </div>
          <form onSubmit={handleLogin} className="p-8 space-y-6">
            <div>
              <label className="block text-[10px] font-bold uppercase text-slate-400 mb-2 tracking-wider pl-1">Pilih Hak Akses Otoritas</label>
              <div className="grid grid-cols-2 gap-3">
                <button type="button" onClick={() => { setSelectedRole('admin'); setLoginError(''); }} className={`p-4 rounded-xl font-bold text-xs uppercase tracking-wider transition-all border flex flex-col items-center gap-3 ${selectedRole === 'admin' ? 'bg-emerald-600 border-emerald-500 text-white shadow-lg shadow-emerald-900/50' : 'bg-slate-950/50 border-slate-800 text-slate-400 hover:bg-slate-800/50'}`}>
                  <Shield size={20}/> Akses Direksi
                </button>
                <button type="button" onClick={() => { setSelectedRole('viewer'); setLoginError(''); }} className={`p-4 rounded-xl font-bold text-xs uppercase tracking-wider transition-all border flex flex-col items-center gap-3 ${selectedRole === 'viewer' ? 'bg-blue-600 border-blue-500 text-white shadow-lg shadow-blue-900/50' : 'bg-slate-950/50 border-slate-800 text-slate-400 hover:bg-slate-800/50'}`}>
                  <UserCircle size={20}/> Akses RUPS
                </button>
              </div>
            </div>

            {selectedRole === 'admin' && (
              <div className="space-y-2 animate-fade-in">
                <label className="block text-[10px] font-bold uppercase text-slate-400 tracking-widest pl-1">Security Key (Password)</label>
                <div className="relative">
                  <Lock size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-500" />
                  <input type="password" placeholder="••••••••" value={loginPassword} onChange={e => setLoginPassword(e.target.value)} className="w-full bg-slate-950/80 border border-slate-700 rounded-xl pl-12 pr-4 py-4 text-white focus:outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 font-mono tracking-widest" required />
                </div>
              </div>
            )}

            {loginError && <div className="p-4 bg-rose-500/10 border border-rose-500/30 text-rose-400 text-xs font-semibold rounded-xl text-center flex items-center justify-center gap-2"><AlertCircle size={14}/> {loginError}</div>}

            <button type="submit" className="w-full bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-400 hover:to-teal-500 text-white p-4 rounded-xl font-black uppercase tracking-widest transition-all shadow-lg shadow-emerald-900/50 hover:shadow-emerald-900/80 active:scale-[0.98]">
              Authenticate & Login
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-950 text-slate-200 flex font-sans print:bg-white print:text-black">
      {/* SIDEBAR */}
      <aside className="w-72 bg-slate-900 border-r border-slate-800/60 flex flex-col z-20 print:hidden relative">
        <div className="absolute top-0 right-0 w-full h-full bg-gradient-to-b from-emerald-900/10 to-transparent opacity-50 pointer-events-none"></div>
        <div className="p-8 border-b border-slate-800/60 bg-slate-900/80 backdrop-blur-md relative z-10">
          <div className="flex items-center gap-3 text-emerald-400 mb-2">
            <Building2 size={24} className="text-emerald-500 drop-shadow-[0_0_8px_rgba(16,185,129,0.5)]" />
            <h1 className="font-black text-lg tracking-widest uppercase truncate">{company.name.split(' ')[0]} {company.name.split(' ')[1]}</h1>
          </div>
          <span className="text-[10px] bg-slate-950/50 px-3 py-1 rounded-full inline-block font-mono text-emerald-500 uppercase tracking-widest border border-emerald-900/50 shadow-inner">Portal Aktif: {currentUser.role}</span>
        </div>
        
        <nav className="flex-1 p-6 space-y-2 overflow-y-auto relative z-10 scrollbar-thin scrollbar-thumb-slate-800 scrollbar-track-transparent">
          <NavItem icon={<LayoutDashboard size={18}/>} label="Executive Dashboard" active={currentView === 'dashboard'} onClick={() => setCurrentView('dashboard')} />
          
          <div className="pt-4 pb-2">
            <p className="text-[10px] font-black text-slate-600 uppercase tracking-widest px-4">Aktivitas & Siklus</p>
          </div>
          <NavItem icon={<Receipt size={18}/>} label="Input Transaksi" active={currentView === 'transactions'} onClick={() => setCurrentView('transactions')} />
          <NavItem icon={<Wallet size={18}/>} label="Buku Kas Besar" active={currentView === 'buku-kas'} onClick={() => setCurrentView('buku-kas')} />
          <NavItem icon={<BookOpen size={18}/>} label="Jurnal Umum" active={currentView === 'journal'} onClick={() => setCurrentView('journal')} />
          
          <div className="pt-4 pb-2">
            <p className="text-[10px] font-black text-slate-600 uppercase tracking-widest px-4">Laporan RUPS</p>
          </div>
          <NavItem icon={<TrendingUp size={18}/>} label="Laba Rugi Komprehensif" active={currentView === 'laba-rugi'} onClick={() => setCurrentView('laba-rugi')} />
          <NavItem icon={<RefreshCw size={18}/>} label="Perubahan Ekuitas" active={currentView === 'perubahan-modal'} onClick={() => setCurrentView('perubahan-modal')} />
          <NavItem icon={<FileSpreadsheet size={18}/>} label="Arus Kas (Cash Flow)" active={currentView === 'arus-kas'} onClick={() => setCurrentView('arus-kas')} />
          <NavItem icon={<FileText size={18}/>} label="Neraca Keuangan" active={currentView === 'neraca'} onClick={() => setCurrentView('neraca')} />
          <NavItem icon={<HelpCircle size={18}/>} label="Catatan CALK" active={currentView === 'calk'} onClick={() => setCurrentView('calk')} />
          <NavItem icon={<Landmark size={18}/>} label="Pajak Badan" active={currentView === 'laporan-pajak'} onClick={() => setCurrentView('laporan-pajak')} />
          <NavItem icon={<FileSignature size={18}/>} label="Berita Acara RUPS" active={currentView === 'berita-acara'} onClick={() => setCurrentView('berita-acara')} />
          {currentUser.role === 'admin' && (
            <>
              <div className="pt-4 pb-2">
                <p className="text-[10px] font-black text-slate-600 uppercase tracking-widest px-4">Konfigurasi</p>
              </div>
              <NavItem icon={<Settings size={18}/>} label="Profil & Akun" active={currentView === 'settings'} onClick={() => setCurrentView('settings')} />
            </>
          )}
        </nav>

        <div className="p-6 border-t border-slate-800/60 bg-slate-900/80 backdrop-blur-md relative z-10 space-y-4">
          {syncStatus && <div className="text-[10px] text-emerald-400 flex items-center justify-center gap-2 font-bold uppercase tracking-widest bg-emerald-950/30 p-2.5 rounded-lg border border-emerald-900/50"><RefreshCw size={12} className="animate-spin" /> {syncStatus}</div>}
          <button onClick={() => { setCurrentUser(null); setLoginPassword(''); }} className="w-full flex items-center justify-center gap-2 p-3 text-xs font-bold text-rose-400 hover:bg-rose-500/10 rounded-xl transition-all border border-transparent hover:border-rose-500/20">
            <LogOut size={16} /> KELUAR SISTEM
          </button>
        </div>
      </aside>

      {/* KONTEN UTAMA */}
      <main className="flex-1 flex flex-col overflow-hidden relative">
        <div className="absolute top-0 left-0 w-full h-96 bg-emerald-900/10 blur-[100px] pointer-events-none -z-10"></div>
        <header className="h-20 border-b border-slate-800/60 flex items-center justify-between px-10 bg-slate-900/20 backdrop-blur-md print:hidden z-10">
          <div className="flex items-center gap-3">
            <span className="text-slate-500 font-medium text-sm">Dashboard</span>
            <ChevronRight size={14} className="text-slate-600" />
            <h2 className="text-sm font-bold uppercase tracking-widest text-emerald-400">{currentView.replace('-', ' ')}</h2>
          </div>
          <div className="flex items-center gap-4 bg-slate-900/80 px-4 py-2 rounded-full border border-slate-800">
            <div className="w-8 h-8 rounded-full bg-slate-800 flex items-center justify-center text-slate-400"><UserCircle size={18} /></div>
            <div className="text-right pr-2">
              <p className="text-xs font-bold text-white">{currentUser.username}</p>
              <p className="text-[9px] font-mono text-emerald-400 uppercase tracking-widest">Active Session</p>
            </div>
          </div>
        </header>
        
        <div className="flex-1 overflow-y-auto p-6 md:p-10 z-10 print:p-0 print:bg-white print:text-black">
          {currentView === 'dashboard' && <DashboardView transactions={transactions} accounts={accounts} balances={balances} />}
          {currentView === 'transactions' && <TransactionsView transactions={transactions} accounts={accounts} setTransactions={setTransactions} company={company} saveToCloud={saveToCloud} user={currentUser} />}
          {currentView === 'buku-kas' && <BukuKasView transactions={transactions} accounts={accounts} company={company} />}
          {currentView === 'journal' && <JournalView transactions={transactions} accounts={accounts} company={company} />}
          {currentView === 'laba-rugi' && <LabaRugiView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'perubahan-modal' && <PerubahanModalView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'arus-kas' && <ArusKasView transactions={transactions} accounts={accounts} balances={balances} company={company} />}
          {currentView === 'neraca' && <NeracaView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'calk' && <CalkView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'laporan-pajak' && <LaporanPajakView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'settings' && <SettingsView company={company} setCompany={setCompany} accounts={accounts} saveToCloud={saveToCloud} user={currentUser} />}
          {currentView === 'berita-acara' && <BeritaAcaraView company={company} />}
        </div>
      </main>
    </div>
  );
}

function NavItem({ icon, label, active, onClick }) {
  return (
    <button onClick={onClick} className={`w-full flex items-center gap-3 px-4 py-3.5 rounded-xl transition-all duration-300 text-xs font-bold uppercase tracking-widest ${active ? 'bg-emerald-500/10 text-emerald-400 shadow-[inset_2px_0_0_0_#10b981]' : 'text-slate-400 hover:bg-slate-800/50 hover:text-slate-200'}`}>
      {icon} {label}
    </button>
  );
}

/* ========================================================================
   1. DASHBOARD VIEW (ULTRA PREMIUM CHARTS)
   ======================================================================== */
function DashboardView({ transactions, accounts, balances }) {
  const cashAccounts = accounts.filter(a => a.name.toLowerCase().includes('kas') || a.name.toLowerCase().includes('bank'));
  const totalCash = cashAccounts.reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const totalRevenue = accounts.filter(a => a.type === 'revenue').reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const totalExpense = accounts.filter(a => a.type === 'expense').reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const netProfit = totalRevenue - totalExpense;

  const assetsTotal = accounts.filter(a => a.type === 'asset').reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const liabilitiesTotal = accounts.filter(a => a.type === 'liability').reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const equityTotal = accounts.filter(a => a.type === 'equity').reduce((sum, a) => sum + (balances[a.id] || 0), 0);

  const pieData = [
    { name: 'Aset', value: assetsTotal, color: '#10b981' },
    { name: 'Kewajiban', value: liabilitiesTotal, color: '#f43f5e' },
    { name: 'Ekuitas', value: equityTotal + netProfit, color: '#3b82f6' }
  ].filter(d => d.value > 0);

  const barData = [
    { name: 'Kinerja', Pendapatan: totalRevenue, Beban: totalExpense }
  ];

  const chartData = useMemo(() => {
    const daily = {};
    transactions.forEach(t => {
      if (!daily[t.date]) daily[t.date] = { date: t.date, Pendapatan: 0, Beban: 0 };
      const isRev = accounts.find(a => a.id === t.creditAccountId)?.type === 'revenue' || accounts.find(a => a.id === t.debitAccountId)?.type === 'revenue';
      const isExp = accounts.find(a => a.id === t.debitAccountId)?.type === 'expense' || accounts.find(a => a.id === t.creditAccountId)?.type === 'expense';
      if (isRev) daily[t.date].Pendapatan += t.amount;
      if (isExp) daily[t.date].Beban += t.amount;
    });
    return Object.values(daily).sort((a,b) => new Date(a.date) - new Date(b.date)).slice(-14);
  }, [transactions, accounts]);

  return (
    <div className="space-y-8 animate-fade-in print:hidden">
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        <StatCard title="Total Kas & Bank" value={totalCash} icon={<Wallet size={24}/>} color="emerald" />
        <StatCard title="Pendapatan Usaha" value={totalRevenue} icon={<TrendingUp size={24}/>} color="blue" />
        <StatCard title="Beban Operasional" value={totalExpense} icon={<Receipt size={24}/>} color="rose" />
        <StatCard title="Laba Bersih" value={netProfit} icon={<Landmark size={24}/>} color="amber" />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        <div className="lg:col-span-2 bg-slate-900/50 border border-slate-800 p-8 rounded-3xl shadow-xl backdrop-blur-sm">
          <h3 className="text-xs font-bold uppercase text-slate-400 tracking-widest mb-8">Tren Kinerja Finansial (14 Hari)</h3>
          <div className="h-80 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                <defs>
                  <linearGradient id="colorRev" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#10b981" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                  </linearGradient>
                  <linearGradient id="colorExp" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#f43f5e" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#f43f5e" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                <XAxis dataKey="date" stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} />
                <YAxis stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} tickFormatter={(v) => `Rp${v/1000000}M`} />
                <Tooltip contentStyle={{backgroundColor: '#0f172a', borderRadius: '12px', border: '1px solid #334155', color: '#fff'}} itemStyle={{fontSize: '12px'}} labelStyle={{fontSize: '10px', color: '#94a3b8'}} />
                <Legend verticalAlign="top" height={36} iconType="circle" wrapperStyle={{fontSize: '11px', fontWeight: 'bold'}}/>
                <Area type="monotone" dataKey="Pendapatan" stroke="#10b981" strokeWidth={3} fillOpacity={1} fill="url(#colorRev)" />
                <Area type="monotone" dataKey="Beban" stroke="#f43f5e" strokeWidth={3} fillOpacity={1} fill="url(#colorExp)" />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="space-y-8">
          <div className="bg-slate-900/50 border border-slate-800 p-6 rounded-3xl shadow-xl backdrop-blur-sm transition-all hover:scale-[1.02]">
            <h3 className="text-[10px] font-bold uppercase text-slate-400 tracking-widest mb-4">Struktur Modal & Kewajiban (Neraca)</h3>
            <div className="h-44 w-full">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={pieData} cx="50%" cy="50%" innerRadius={50} outerRadius={70} paddingAngle={5} dataKey="value" stroke="none">
                    {pieData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip contentStyle={{backgroundColor: '#0f172a', borderRadius: '12px', border: '1px solid #334155', color: '#fff'}} itemStyle={{fontSize: '12px'}} formatter={(val) => formatCurrency(val)} />
                  <Legend verticalAlign="bottom" height={20} iconType="circle" wrapperStyle={{fontSize: '10px', fontWeight: 'bold'}}/>
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="bg-slate-900/50 border border-slate-800 p-6 rounded-3xl shadow-xl backdrop-blur-sm transition-all hover:scale-[1.02]">
            <h3 className="text-[10px] font-bold uppercase text-slate-400 tracking-widest mb-4">Komparasi Pendapatan vs Beban</h3>
            <div className="h-44 w-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={barData} layout="vertical" margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                  <XAxis type="number" hide />
                  <YAxis dataKey="name" type="category" hide />
                  <Tooltip cursor={{fill: 'transparent'}} contentStyle={{backgroundColor: '#0f172a', borderRadius: '12px', border: '1px solid #334155', color: '#fff'}} itemStyle={{fontSize: '12px'}} formatter={(val) => formatCurrency(val)} />
                  <Bar dataKey="Pendapatan" fill="#10b981" radius={[0, 4, 4, 0]} barSize={24} animationDuration={1500} />
                  <Bar dataKey="Beban" fill="#f43f5e" radius={[0, 4, 4, 0]} barSize={24} animationDuration={1500} />
                  <Legend verticalAlign="bottom" height={20} iconType="circle" wrapperStyle={{fontSize: '10px', fontWeight: 'bold'}}/>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-8 mt-8">
        <div className="bg-gradient-to-br from-slate-900 to-slate-950 border border-slate-800 p-8 rounded-3xl shadow-xl flex flex-col md:flex-row justify-between items-center gap-6">
          <div className="flex-1">
            <h3 className="text-xs font-bold uppercase text-slate-400 tracking-widest mb-6">Ringkasan Eksekutif Terkini</h3>
            <div className="grid grid-cols-2 gap-6">
              <div className="border-l-2 border-emerald-500 pl-4">
                <p className="text-[10px] text-slate-500 uppercase tracking-widest font-bold">Margin Laba Kotor</p>
                <p className="text-2xl font-bold text-white mt-1">{totalRevenue > 0 ? ((netProfit/totalRevenue)*100).toFixed(1) : 0}%</p>
              </div>
              <div className="border-l-2 border-blue-500 pl-4">
                <p className="text-[10px] text-slate-500 uppercase tracking-widest font-bold">Total Transaksi Tercatat</p>
                <p className="text-2xl font-bold text-white mt-1">{transactions.length} <span className="text-sm font-normal text-slate-500">Entri</span></p>
              </div>
            </div>
          </div>
          <div className="w-full md:w-1/3 pt-6 md:pt-0 border-t md:border-t-0 md:border-l border-slate-800/60 md:pl-6">
            <p className="text-[10px] text-slate-500 italic text-justify leading-relaxed">Seluruh data yang disajikan dilindungi oleh enkripsi end-to-end dan disinkronisasi secara real-time dengan server cloud korporasi. Kinerja perusahaan saat ini berada pada tren <strong>{netProfit >= 0 ? 'SEHAT (PROFITABLE)' : 'DEFISIT (PERLU PERHATIAN)'}</strong>.</p>
          </div>
        </div>
      </div>
    </div>
  );
}

function StatCard({ title, value, icon, color }) {
  const colors = {
    emerald: 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20',
    blue: 'bg-blue-500/10 text-blue-400 border-blue-500/20',
    rose: 'bg-rose-500/10 text-rose-400 border-rose-500/20',
    amber: 'bg-amber-500/10 text-amber-400 border-amber-500/20'
  };
  return (
    <div className="bg-slate-900/50 border border-slate-800 p-6 rounded-3xl shadow-xl relative overflow-hidden group hover:border-slate-700 transition-colors backdrop-blur-sm">
      <div className={`absolute -right-6 -top-6 w-24 h-24 rounded-full blur-2xl transition-all group-hover:scale-150 ${colors[color].split(' ')[0]}`}></div>
      <div className="relative z-10 flex flex-col h-full justify-between gap-6">
        <div className="flex justify-between items-start">
          <div className={`p-2.5 rounded-xl border ${colors[color]}`}>{icon}</div>
        </div>
        <div>
          <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-1">{title}</p>
          <h4 className="text-2xl font-black text-white font-mono tracking-tight">{formatCurrency(value)}</h4>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   2. INPUT TRANSAKSI (Form Mewah & Cerdas)
   ======================================================================== */
function TransactionsView({ transactions, accounts, setTransactions, company, saveToCloud, user }) {
  const [showForm, setShowForm] = useState(false);
  const [trxType, setTrxType] = useState('pemasukan');
  const [formData, setFormData] = useState({ date: new Date().toISOString().split('T')[0], description: '', amount: '', accountId: '' });

  const executeDelete = (id) => {
    if (user.role !== 'admin') return;
    if (window.confirm("Hapus transaksi permanen?")) {
      const updated = transactions.filter(t => t.id !== id);
      setTransactions(updated);
      saveToCloud(updated, null, null);
    }
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    if (user.role !== 'admin') return;
    const amountNum = parseFloat(formData.amount);
    if (!formData.accountId || isNaN(amountNum) || amountNum <= 0) return;

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
    <div className="space-y-8 animate-fade-in">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-2xl font-black text-white uppercase tracking-wider">Pencatatan Transaksi</h2>
          <p className="text-slate-400 text-xs font-medium mt-1">Registrasi aliran kas operasional harian perusahaan.</p>
        </div>
        {user.role === 'admin' ? (
          <button onClick={() => setShowForm(!showForm)} className="bg-emerald-600 hover:bg-emerald-500 px-6 py-3.5 rounded-xl text-white font-bold text-xs uppercase tracking-widest flex items-center gap-2 shadow-lg shadow-emerald-900/40 transition-all active:scale-95">
            {showForm ? <X size={16}/> : <Plus size={16}/>} {showForm ? 'Tutup Formulir' : 'Entri Kas Baru'}
          </button>
        ) : (
          <div className="flex items-center gap-2 bg-amber-500/10 border border-amber-500/20 text-amber-400 px-4 py-2.5 rounded-xl text-xs font-bold uppercase tracking-wider">
            <Lock size={14} /> Akses Terkunci (Viewer)
          </div>
        )}
      </div>

      {showForm && user.role === 'admin' && (
        <div className="bg-slate-900 border border-slate-800 p-8 rounded-3xl shadow-2xl max-w-3xl">
          <div className="flex bg-slate-950 p-1.5 rounded-xl mb-8 border border-slate-800">
            <button type="button" onClick={() => setTrxType('pemasukan')} className={`flex-1 py-3.5 rounded-lg font-black text-xs uppercase tracking-widest transition-all ${trxType === 'pemasukan' ? 'bg-emerald-600 text-white shadow-md' : 'text-slate-500 hover:text-slate-300'}`}>Kas Masuk (+)</button>
            <button type="button" onClick={() => setTrxType('pengeluaran')} className={`flex-1 py-3.5 rounded-lg font-black text-xs uppercase tracking-widest transition-all ${trxType === 'pengeluaran' ? 'bg-rose-600 text-white shadow-md' : 'text-slate-500 hover:text-slate-300'}`}>Kas Keluar (-)</button>
          </div>
          <form onSubmit={handleSubmit} className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Tanggal Transaksi</label>
                <input type="date" value={formData.date} onChange={e => setFormData({...formData, date: e.target.value})} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3.5 text-sm text-white focus:outline-none focus:border-emerald-500 transition-colors" required />
              </div>
              <div>
                <label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Nominal Uang (Rp)</label>
                <input type="number" placeholder="0" value={formData.amount} onChange={e => setFormData({...formData, amount: e.target.value})} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3.5 text-sm text-white focus:outline-none focus:border-emerald-500 font-mono transition-colors" required />
              </div>
            </div>
            <div>
              <label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Uraian / Keterangan Lengkap</label>
              <input type="text" placeholder="Contoh: Pembayaran invoice klien A..." value={formData.description} onChange={e => setFormData({...formData, description: e.target.value})} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3.5 text-sm text-white focus:outline-none focus:border-emerald-500 transition-colors" required />
            </div>
            <div>
              <label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Sumber / Tujuan (Akun Lawan)</label>
              <select value={formData.accountId} onChange={e => setFormData({...formData, accountId: e.target.value})} className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3.5 text-sm text-white focus:outline-none focus:border-emerald-500 transition-colors" required>
                <option value="">-- Pilih Kategori Akuntansi --</option>
                {accounts.filter(a => a.id !== '101').map(a => <option key={a.id} value={a.id}>{a.id} - {a.name} ({a.type})</option>)}
              </select>
            </div>
            <button type="submit" className="w-full bg-emerald-600 hover:bg-emerald-500 text-white py-4 rounded-xl font-black text-xs uppercase tracking-widest shadow-lg shadow-emerald-900/40 transition-all mt-4">Simpan ke Buku Besar</button>
          </form>
        </div>
      )}

      <div className="bg-slate-900 border border-slate-800 rounded-3xl overflow-hidden shadow-xl">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm whitespace-nowrap">
            <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
              <tr><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Tanggal</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Keterangan</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Debet</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Kredit</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-right">Nominal</th>{user.role === 'admin' && <th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-center">Aksi</th>}</tr>
            </thead>
            <tbody className="divide-y divide-slate-800/50">
              {[...transactions].reverse().map(t => (
                <tr key={t.id} className="hover:bg-slate-800/30 transition-colors group">
                  <td className="px-6 py-4 text-slate-300">{t.date} <br/><span className="text-[9px] font-mono text-slate-600">{t.id}</span></td>
                  <td className="px-6 py-4 text-white font-medium">{t.description}</td>
                  <td className="px-6 py-4 text-emerald-400 text-xs font-semibold">{accounts.find(a=>a.id===t.debitAccountId)?.name}</td>
                  <td className="px-6 py-4 text-rose-400 text-xs font-semibold">{accounts.find(a=>a.id===t.creditAccountId)?.name}</td>
                  <td className="px-6 py-4 text-right font-bold text-white font-mono">{formatCurrency(t.amount)}</td>
                  {user.role === 'admin' && (
                    <td className="px-6 py-4 text-center">
                      <button onClick={() => executeDelete(t.id)} className="text-slate-600 hover:text-rose-500 hover:bg-rose-500/10 p-2 rounded-lg transition-all opacity-0 group-hover:opacity-100"><Trash2 size={16}/></button>
                    </td>
                  )}
                </tr>
              ))}
              {transactions.length === 0 && <tr><td colSpan="7" className="p-10 text-center text-slate-500 font-medium">Belum ada transaksi tercatat.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   3. BUKU KAS BESAR (Filter & Download Rapi)
   ======================================================================== */
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
      { header: 'Tanggal', key: 'date', width: 15 },
      { header: 'No. Bukti', key: 'id', width: 25 },
      { header: 'Uraian Deskripsi', key: 'desc', width: 40 },
      { header: 'Ref Akun Lawan', key: 'ref', width: 30 },
      { header: 'Debet (Masuk)', key: 'masuk', width: 20 },
      { header: 'Kredit (Keluar)', key: 'keluar', width: 20 },
      { header: 'Saldo Kumulatif', key: 'saldo', width: 25 }
    ];
    exportGridToExcel('Buku Kas', cols, filtered, `Buku_Kas_${company.name}_${startDate||'All'}.xlsx`, company.name);
  };

  return (
    <div className="space-y-8 animate-fade-in">
      <div>
        <h2 className="text-2xl font-black text-white uppercase tracking-wider">Buku Kas & Bank</h2>
        <p className="text-slate-400 text-xs font-medium mt-1">Rincian mutasi kas masuk dan keluar secara kronologis.</p>
      </div>

      <div className="bg-slate-900 border border-slate-800 p-4 rounded-2xl shadow-lg flex flex-wrap gap-4 items-center print:hidden">
        <Filter size={20} className="text-emerald-500" />
        <span className="text-xs font-bold uppercase text-slate-400 tracking-widest">Filter Periode:</span>
        <input type="date" className="bg-slate-950 border border-slate-700 p-2.5 rounded-xl text-sm text-white focus:border-emerald-500 focus:outline-none" value={startDate} onChange={e => setStartDate(e.target.value)} />
        <span className="text-slate-500 text-sm">s.d.</span>
        <input type="date" className="bg-slate-950 border border-slate-700 p-2.5 rounded-xl text-sm text-white focus:border-emerald-500 focus:outline-none" value={endDate} onChange={e => setEndDate(e.target.value)} />
        <button onClick={downloadExcel} className="ml-auto bg-blue-600 hover:bg-blue-500 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg shadow-blue-900/40 transition-all">
          <Download size={16}/> Unduh Excel Rapi
        </button>
      </div>

      <div className="bg-slate-900 border border-slate-800 rounded-3xl overflow-hidden shadow-xl">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm whitespace-nowrap">
            <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
              <tr><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Tanggal</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Keterangan & Ref</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-right">Debet (Masuk)</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-right">Kredit (Keluar)</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-right">Saldo Kas</th></tr>
            </thead>
            <tbody className="divide-y divide-slate-800/40">
              {[...filtered].reverse().map((e, idx) => (
                <tr key={idx} className="hover:bg-slate-800/30">
                  <td className="px-6 py-4 font-medium text-slate-300">{e.date}</td>
                  <td className="px-6 py-4 text-white font-medium">{e.desc} <span className="text-[10px] text-slate-500 block font-normal mt-1">Ref: {e.ref}</span></td>
                  <td className="px-6 py-4 text-right text-emerald-400 font-mono font-bold">{e.masuk > 0 ? formatCurrency(e.masuk) : '-'}</td>
                  <td className="px-6 py-4 text-right text-rose-400 font-mono font-bold">{e.keluar > 0 ? formatCurrency(e.keluar) : '-'}</td>
                  <td className="px-6 py-4 text-right text-white font-mono font-black bg-slate-950/30">{formatCurrency(e.saldo)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   4. JURNAL UMUM (Filter & Download Rapi)
   ======================================================================== */
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
      rows.push({ date: '', id: '', desc: '', account: `    ${accounts.find(a=>a.id===t.creditAccountId)?.name}`, debit: 0, kredit: t.amount });
    });

    const cols = [
      { header: 'Tanggal', key: 'date', width: 15 },
      { header: 'No Jurnal', key: 'id', width: 25 },
      { header: 'Keterangan Transaksi', key: 'desc', width: 40 },
      { header: 'Akun Perkiraan', key: 'account', width: 35 },
      { header: 'Debit (Rp)', key: 'debit', width: 20 },
      { header: 'Kredit (Rp)', key: 'kredit', width: 20 }
    ];
    exportGridToExcel('Jurnal Umum', cols, rows, `Jurnal_Umum_${company.name}_${startDate||'All'}.xlsx`, company.name);
  };

  return (
    <div className="space-y-8 animate-fade-in">
      <div>
        <h2 className="text-2xl font-black text-white uppercase tracking-wider">Jurnal Umum</h2>
        <p className="text-slate-400 text-xs font-medium mt-1">Pencatatan double-entry (Debit & Kredit) standar akuntansi.</p>
      </div>

      <div className="bg-slate-900 border border-slate-800 p-4 rounded-2xl shadow-lg flex flex-wrap gap-4 items-center print:hidden">
        <Filter size={20} className="text-emerald-500" />
        <span className="text-xs font-bold uppercase text-slate-400 tracking-widest">Filter Periode:</span>
        <input type="date" className="bg-slate-950 border border-slate-700 p-2.5 rounded-xl text-sm text-white" value={startDate} onChange={e => setStartDate(e.target.value)} />
        <span className="text-slate-500 text-sm">s.d.</span>
        <input type="date" className="bg-slate-950 border border-slate-700 p-2.5 rounded-xl text-sm text-white" value={endDate} onChange={e => setEndDate(e.target.value)} />
        <button onClick={downloadExcel} className="ml-auto bg-blue-600 hover:bg-blue-500 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg shadow-blue-900/40 transition-all">
          <Download size={16}/> Unduh Excel Rapi
        </button>
      </div>

      <div className="bg-slate-900 border border-slate-800 rounded-3xl overflow-hidden shadow-xl">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm whitespace-nowrap">
            <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
              <tr><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Tanggal / Ref</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Nama Akun & Keterangan</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-right">Debit</th><th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-right">Kredit</th></tr>
            </thead>
            <tbody className="divide-y divide-slate-800/40">
              {[...filtered].reverse().map((t, idx) => (
                <React.Fragment key={idx}>
                  <tr className="bg-slate-900/40 hover:bg-slate-800/20">
                    <td className="px-6 py-4 text-slate-400 font-medium align-top" rowSpan="2">{t.date}<br/><span className="text-[10px] font-mono block text-slate-600 mt-1">{t.id}</span></td>
                    <td className="px-6 py-3 text-white font-bold">{t.debitAccountId} - {accounts.find(a=>a.id===t.debitAccountId)?.name}<span className="block text-xs text-slate-500 font-normal mt-1">{t.description}</span></td>
                    <td className="px-6 py-3 text-right text-emerald-400 font-mono font-bold text-sm bg-slate-950/20">{formatCurrency(t.amount)}</td>
                    <td className="px-6 py-3 text-right text-slate-600 font-mono bg-slate-950/20">-</td>
                  </tr>
                  <tr className="border-b-4 border-slate-950 hover:bg-slate-800/20">
                    <td className="px-6 py-3 text-slate-300 font-medium italic pl-14">{t.creditAccountId} - {accounts.find(a=>a.id===t.creditAccountId)?.name}</td>
                    <td className="px-6 py-3 text-right text-slate-600 font-mono bg-slate-950/20">-</td>
                    <td className="px-6 py-3 text-right text-rose-400 font-mono font-bold text-sm bg-slate-950/20">{formatCurrency(t.amount)}</td>
                  </tr>
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   5. LAPORAN LABA RUGI (Print Ready + Excel)
   ======================================================================== */
function LabaRugiView({ accounts, balances, company }) {
  const revAccounts = accounts.filter(a => a.type === 'revenue');
  const expAccounts = accounts.filter(a => a.type === 'expense');

  const totalRev = revAccounts.reduce((s, a) => s + (balances[a.id] || 0), 0);
  const totalExp = expAccounts.reduce((s, a) => s + (balances[a.id] || 0), 0);
  const netProfit = totalRev - totalExp;

  const handleDownloadExcel = () => {
    const data = [
      { label: 'PENDAPATAN OPERASIONAL', amount: null, isHeader: true, isBold: true },
      ...revAccounts.map(a => ({ label: `   ${a.name}`, amount: balances[a.id] || 0, isBold: false })),
      { label: 'TOTAL PENDAPATAN', amount: totalRev, isBold: true, isTotal: false },
      { label: '', amount: null },
      { label: 'BEBAN OPERASIONAL', amount: null, isHeader: true, isBold: true },
      ...expAccounts.map(a => ({ label: `   ${a.name}`, amount: balances[a.id] || 0, isBold: false })),
      { label: 'TOTAL BEBAN', amount: totalExp, isBold: true, isTotal: false },
      { label: '', amount: null },
      { label: 'LABA BERSIH TAHUN BERJALAN', amount: netProfit, isBold: true, isTotal: true }
    ];
    exportFinancialToExcel('Laba Rugi', company.name, data, `Laba_Rugi_${company.name}.xlsx`);
  };

  return (
    <div className="space-y-8 animate-fade-in">
      <div className="flex justify-end gap-4 print:hidden">
        <button onClick={handleDownloadExcel} className="bg-blue-600 hover:bg-blue-500 border border-blue-500 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Download size={14}/> Unduh Excel Rapi</button>
        <button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 border border-slate-700 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Printer size={14}/> Cetak RUPS</button>
      </div>
      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-4xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[800px] font-sans border border-slate-200">
        <PrintLetterhead company={company} />
        <div className="text-center pb-6 mb-10">
          <h3 className="text-lg font-bold mt-2 text-slate-700 uppercase tracking-widest">Laporan Laba Rugi Komprehensif</h3>
          <p className="text-sm text-slate-500 font-medium mt-1">Periode Berakhir Per: {new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
        </div>

        <div className="space-y-8 text-sm md:text-base">
          <div>
            <h4 className="font-bold text-slate-900 bg-slate-100 p-2 uppercase tracking-widest border-l-4 border-slate-900">1. PEREDARAN USAHA (PENDAPATAN)</h4>
            <div className="space-y-3 mt-4 pl-4 pr-2">
              {revAccounts.map(a => (
                <div key={a.id} className="flex justify-between text-slate-700"><span>{a.name}</span><span className="font-mono">{formatCurrency(balances[a.id] || 0)}</span></div>
              ))}
            </div>
            <div className="flex justify-between font-black mt-4 pt-3 border-t-2 border-slate-900 text-slate-900 pl-4 pr-2"><span>TOTAL PENDAPATAN</span><span className="font-mono">{formatCurrency(totalRev)}</span></div>
          </div>

          <div>
            <h4 className="font-bold text-slate-900 bg-slate-100 p-2 uppercase tracking-widest border-l-4 border-slate-900">2. BEBAN OPERASIONAL & BPP</h4>
            <div className="space-y-3 mt-4 pl-4 pr-2">
              {expAccounts.map(a => (
                <div key={a.id} className="flex justify-between text-slate-700"><span>{a.name}</span><span className="font-mono">({formatCurrency(balances[a.id] || 0)})</span></div>
              ))}
            </div>
            <div className="flex justify-between font-black mt-4 pt-3 border-t-2 border-slate-900 text-slate-900 pl-4 pr-2"><span>TOTAL BEBAN</span><span className="font-mono">({formatCurrency(totalExp)})</span></div>
          </div>

          <div className="pt-6 border-t-[6px] border-double border-slate-900 flex justify-between font-black text-xl text-slate-900 p-2">
            <span>LABA BERSIH TAHUN BERJALAN</span>
            <span className={`font-mono ${netProfit < 0 ? 'text-rose-600' : 'text-slate-900'}`}>{formatCurrency(netProfit)}</span>
          </div>
        </div>

        <div className="mt-32 grid grid-cols-2 gap-8 text-center text-sm font-medium print:mt-40">
          <div><p className="mb-24 text-slate-500">Dibuat Oleh,</p><p className="font-bold border-b border-slate-900 inline-block px-10 pb-1 text-slate-900">DIREKSI KEUANGAN</p></div>
          <div><p className="mb-24 text-slate-500">Disahkan Oleh,</p><p className="font-bold border-b border-slate-900 inline-block px-10 pb-1 text-slate-900">DIREKTUR UTAMA</p></div>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   6. NERACA POSISI KEUANGAN (Print Ready + Excel)
   ======================================================================== */
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

  const handleDownloadExcel = () => {
    const data = [
      { label: 'AKTIVA (ASET)', amount: null, isHeader: true, isBold: true },
      ...assets.map(a => ({ label: `   ${a.name}`, amount: balances[a.id] || 0, isBold: false })),
      { label: 'TOTAL AKTIVA', amount: totalAsset, isBold: true, isTotal: true },
      { label: '', amount: null },
      { label: 'PASIVA (KEWAJIBAN & EKUITAS)', amount: null, isHeader: true, isBold: true },
      { label: 'Kewajiban', amount: null, isBold: true },
      ...liabilities.map(a => ({ label: `   ${a.name}`, amount: balances[a.id] || 0, isBold: false })),
      { label: 'Ekuitas', amount: null, isBold: true },
      ...equities.map(a => ({ label: `   ${a.name}`, amount: balances[a.id] || 0, isBold: false })),
      { label: '   Laba Bersih Berjalan', amount: netIncome, isBold: false },
      { label: 'TOTAL PASIVA', amount: totalPasiva, isBold: true, isTotal: true }
    ];
    exportFinancialToExcel('Neraca', company.name, data, `Neraca_${company.name}.xlsx`);
  };

  return (
    <div className="space-y-8 animate-fade-in">
      <div className="flex justify-end gap-4 print:hidden">
        <button onClick={handleDownloadExcel} className="bg-blue-600 hover:bg-blue-500 border border-blue-500 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Download size={14}/> Unduh Excel Rapi</button>
        <button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 border border-slate-700 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Printer size={14}/> Cetak RUPS</button>
      </div>
      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-5xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[800px] font-sans border border-slate-200">
        <PrintLetterhead company={company} />
        <div className="text-center pb-6 mb-10">
          <h3 className="text-lg font-bold mt-2 text-slate-700 uppercase tracking-widest">Laporan Posisi Keuangan (Neraca)</h3>
          <p className="text-sm text-slate-500 font-medium mt-1">Per Tanggal: {new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-12 text-sm md:text-base">
          {/* AKTIVA */}
          <div className="space-y-4">
            <h4 className="font-bold text-slate-900 bg-slate-100 p-2 text-center uppercase tracking-widest border border-slate-300">AKTIVA (ASET)</h4>
            <div className="space-y-3 pl-2 pr-2">
              {assets.map(a => (
                <div key={a.id} className="flex justify-between border-b border-slate-100 pb-1 text-slate-700"><span>{a.name}</span><span className="font-mono">{formatCurrency(balances[a.id] || 0)}</span></div>
              ))}
            </div>
            <div className="flex justify-between font-black pt-4 border-t-2 border-slate-900 text-slate-900 mt-6 px-2"><span>TOTAL AKTIVA</span><span className="font-mono border-b-4 border-double border-slate-900 pb-1">{formatCurrency(totalAsset)}</span></div>
          </div>

          {/* PASIVA */}
          <div className="space-y-6">
            <h4 className="font-bold text-slate-900 bg-slate-100 p-2 text-center uppercase tracking-widest border border-slate-300">PASIVA (KEWAJIBAN & MODAL)</h4>
            <div>
              <h5 className="font-bold text-slate-800 uppercase tracking-widest border-b-2 border-slate-200 mb-3">Kewajiban</h5>
              <div className="space-y-3 pl-2 pr-2">
                {liabilities.map(a => (
                  <div key={a.id} className="flex justify-between border-b border-slate-100 pb-1 text-slate-700"><span>{a.name}</span><span className="font-mono">{formatCurrency(balances[a.id] || 0)}</span></div>
                ))}
              </div>
            </div>
            <div>
              <h5 className="font-bold text-slate-800 uppercase tracking-widest border-b-2 border-slate-200 mb-3 mt-6">Ekuitas</h5>
              <div className="space-y-3 pl-2 pr-2">
                {equities.map(a => (
                  <div key={a.id} className="flex justify-between text-slate-700"><span>{a.name}</span><span className="font-mono">{formatCurrency(balances[a.id] || 0)}</span></div>
                ))}
                <div className="flex justify-between text-slate-900 font-bold bg-slate-50 p-2 rounded border border-slate-200"><span>Laba Bersih Berjalan</span><span className="font-mono">{formatCurrency(netIncome)}</span></div>
              </div>
            </div>
            <div className="flex justify-between font-black pt-4 border-t-2 border-slate-900 text-slate-900 mt-6 px-2"><span>TOTAL PASIVA</span><span className="font-mono border-b-4 border-double border-slate-900 pb-1">{formatCurrency(totalPasiva)}</span></div>
          </div>
        </div>

        <div className="mt-32 grid grid-cols-2 gap-8 text-center text-sm font-medium print:mt-40">
          <div><p className="mb-24 text-slate-500">Dibuat Oleh,</p><p className="font-bold border-b border-slate-900 inline-block px-10 pb-1 text-slate-900">DIREKSI KEUANGAN</p></div>
          <div><p className="mb-24 text-slate-500">Disahkan Oleh,</p><p className="font-bold border-b border-slate-900 inline-block px-10 pb-1 text-slate-900">DIREKTUR UTAMA</p></div>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   7 & 8. CALK & PAJAK (Print Ready)
   ======================================================================== */
function CalkView({ accounts, balances, company }) {
  return (
    <div className="space-y-8 animate-fade-in">
      <div className="flex justify-end print:hidden"><button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 border border-slate-700 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg"><Printer size={14}/> Cetak Dokumen RUPS</button></div>
      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-4xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[800px] font-sans border border-slate-200">
        <PrintLetterhead company={company} />
        <div className="text-center pb-6 mb-10">
          <h3 className="text-lg font-bold mt-2 text-slate-700 uppercase tracking-widest">Catatan Atas Laporan Keuangan (CALK)</h3>
          <p className="text-sm text-slate-500 font-medium mt-1">Tahun Buku Konstitusi Aktif</p>
        </div>

        <div className="space-y-8 text-sm md:text-base">
          <section>
            <h4 className="font-bold text-slate-900 uppercase tracking-widest mb-3 bg-slate-100 p-2 border-l-4 border-slate-900">1. Gambaran Umum Perusahaan</h4>
            <p className="text-slate-700 text-justify leading-relaxed pl-4">
              {company.name} didirikan secara sah berdasarkan hukum di Republik Indonesia. Perusahaan bergerak dalam bidang perdagangan komoditas utama serta logistik. Segala operasional berpusat di alamat {company.address}.
            </p>
          </section>

          <section>
            <h4 className="font-bold text-slate-900 uppercase tracking-widest mb-3 bg-slate-100 p-2 border-l-4 border-slate-900">2. Kebijakan Akuntansi Penting</h4>
            <p className="text-slate-700 text-justify leading-relaxed pl-4 mb-3">
              <strong>A. Dasar Penyusunan:</strong> Disusun berdasarkan Standar Akuntansi Keuangan Entitas Tanpa Akuntabilitas Publik (SAK ETAP). Menggunakan basis akrual penuh (*Full Accrual Basis*).
            </p>
            <p className="text-slate-700 text-justify leading-relaxed pl-4">
              <strong>B. Pengakuan Pendapatan:</strong> Diakui pada saat penyerahan komoditas secara fisik kepada pembeli dan invoice telah tervalidasi sah.
            </p>
          </section>

          <section>
            <h4 className="font-bold text-slate-900 uppercase tracking-widest mb-4 bg-slate-100 p-2 border-l-4 border-slate-900">3. Rincian Pos Keuangan</h4>
            <div className="space-y-3 pl-4 pr-4">
              {accounts.map(a => (
                <div key={a.id} className="flex justify-between border-b border-slate-200 py-2 text-slate-700">
                  <span>Saldo Akun {a.id} - {a.name} ({a.type})</span>
                  <span className="font-mono font-bold text-slate-900">{formatCurrency(balances[a.id] || 0)}</span>
                </div>
              ))}
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}

function LaporanPajakView({ accounts, balances, company }) {
  const totalRev = accounts.filter(a => a.type === 'revenue').reduce((s, a) => s + (balances[a.id] || 0), 0);
  const pphFinal = totalRev * 0.005; // PP 55 / 2022

  return (
    <div className="space-y-8 animate-fade-in">
      <div className="flex justify-end print:hidden"><button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg"><Printer size={14}/> Cetak SPT Masa</button></div>
      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-4xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[800px] font-sans border border-slate-200">
        <PrintLetterhead company={company} />
        <div className="flex justify-between items-start pb-6 mb-10">
          <div>
            <h3 className="text-lg font-bold mt-2 text-slate-700 uppercase tracking-widest">Lampiran SPT PPh Final</h3>
          </div>
          <div className="text-right">
            <span className="text-sm font-black bg-slate-100 border border-slate-300 p-3 rounded-md uppercase tracking-widest inline-block">Sesuai PP 55 / 2022</span>
            <p className="text-xs font-mono mt-3 text-slate-500">Tahun Fiskal Berjalan</p>
          </div>
        </div>

        <div className="bg-slate-50 p-6 border-2 border-slate-200 rounded-2xl text-center mb-10">
          <h4 className="font-black text-slate-900 text-lg uppercase tracking-widest mb-2">Transkrip Perhitungan Pajak Badan</h4>
          <p className="text-slate-600 text-sm">Sesuai Peraturan Pemerintah (PP) No. 55 Tahun 2022 (Insentif UMKM 0.5%)</p>
        </div>

        <div className="space-y-6 text-base mt-8 px-4">
          <div className="flex justify-between border-b-2 border-slate-100 py-4 text-slate-700 font-medium"><span>Dasar Pengenaan Pajak (Peredaran Bruto)</span><span className="font-mono font-bold">{formatCurrency(totalRev)}</span></div>
          <div className="flex justify-between border-b-2 border-slate-100 py-4 text-slate-700 font-medium"><span>Tarif Pajak Final (PP 55)</span><span className="font-mono font-black text-emerald-700">0.5 %</span></div>
          <div className="flex justify-between font-black text-xl bg-slate-900 p-6 rounded-2xl text-white mt-8 shadow-xl"><span>PPH FINAL TERUTANG</span><span className="font-mono text-emerald-400">{formatCurrency(pphFinal)}</span></div>
        </div>

        <div className="mt-40 text-right text-sm print:block print:mt-48">
          <div className="w-64 ml-auto text-center">
            <p className="mb-24 text-slate-500">Wajib Pajak / Kuasa Direksi,</p>
            <p className="font-bold border-b border-slate-900 inline-block px-6 pb-1 uppercase text-slate-900">{company.name}</p>
            <p className="text-xs text-slate-400 mt-2">Cap Resmi Perusahaan</p>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   LAPORAN RUPS TAMBAHAN (PERUBAHAN EKUITAS & ARUS KAS)
   ======================================================================== */
function PerubahanModalView({ accounts, balances, company }) {
  const revTotal = accounts.filter(a => a.type === 'revenue').reduce((s, a) => s + (balances[a.id] || 0), 0);
  const expTotal = accounts.filter(a => a.type === 'expense').reduce((s, a) => s + (balances[a.id] || 0), 0);
  const netIncome = revTotal - expTotal;

  const equities = accounts.filter(a => a.type === 'equity');
  const modalAwal = equities.reduce((s, a) => s + (balances[a.id] || 0), 0);
  const modalAkhir = modalAwal + netIncome;

  const handleDownloadExcel = () => {
    const data = [
      { label: 'MODAL AWAL', amount: modalAwal, isHeader: true, isBold: true },
      { label: 'Ditambah: Laba Bersih Tahun Berjalan', amount: netIncome > 0 ? netIncome : 0, isBold: false },
      { label: 'Dikurangi: Rugi Bersih Tahun Berjalan', amount: netIncome < 0 ? Math.abs(netIncome) : 0, isBold: false },
      { label: 'MODAL AKHIR PERIODE', amount: modalAkhir, isBold: true, isTotal: true }
    ];
    exportFinancialToExcel('Perubahan Modal', company.name, data, `Perubahan_Modal_${company.name}.xlsx`);
  };

  return (
    <div className="space-y-8 animate-fade-in">
      <div className="flex justify-end gap-4 print:hidden">
        <button onClick={handleDownloadExcel} className="bg-blue-600 hover:bg-blue-500 border border-blue-500 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Download size={14}/> Unduh Excel Rapi</button>
        <button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 border border-slate-700 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Printer size={14}/> Cetak RUPS</button>
      </div>
      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-4xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[800px] font-sans border border-slate-200">
        <PrintLetterhead company={company} />
        <div className="text-center pb-6 mb-10">
          <h3 className="text-lg font-bold mt-2 text-slate-700 uppercase tracking-widest">Laporan Perubahan Ekuitas</h3>
          <p className="text-sm text-slate-500 font-medium mt-1">Periode Berakhir Per: {new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
        </div>

        <div className="space-y-6 text-sm md:text-base border border-slate-200 rounded-xl p-8 bg-slate-50">
          <div className="flex justify-between items-center py-3 border-b border-slate-300">
            <span className="font-bold text-slate-800 uppercase">Modal Saham Ditempatkan (Awal)</span>
            <span className="font-mono font-bold text-lg">{formatCurrency(modalAwal)}</span>
          </div>
          <div className="flex justify-between items-center py-3 text-slate-600 pl-4">
            <span>{netIncome >= 0 ? 'Ditambah: Laba Bersih Tahun Berjalan' : 'Dikurangi: Rugi Bersih Tahun Berjalan'}</span>
            <span className="font-mono">{formatCurrency(Math.abs(netIncome))}</span>
          </div>
          <div className="flex justify-between items-center py-4 border-t-2 border-slate-900 mt-4">
            <span className="font-black text-slate-900 uppercase text-lg">Modal Akhir Periode</span>
            <span className="font-mono font-black text-xl text-slate-900 border-b-4 border-double border-slate-900 pb-1">{formatCurrency(modalAkhir)}</span>
          </div>
        </div>

        <div className="mt-32 grid grid-cols-2 gap-8 text-center text-sm font-medium print:mt-40">
          <div><p className="mb-24 text-slate-500">Dibuat Oleh,</p><p className="font-bold border-b border-slate-900 inline-block px-10 pb-1 text-slate-900">DIREKSI KEUANGAN</p></div>
          <div><p className="mb-24 text-slate-500">Disahkan Oleh,</p><p className="font-bold border-b border-slate-900 inline-block px-10 pb-1 text-slate-900">DIREKTUR UTAMA</p></div>
        </div>
      </div>
    </div>
  );
}

function ArusKasView({ transactions, accounts, company }) {
  let kasMasukOperasional = 0;
  let kasKeluarOperasional = 0;
  let kasMasukPendanaan = 0;
  
  transactions.forEach(t => {
    if (t.debitAccountId === '101') {
      const crAcc = accounts.find(a => a.id === t.creditAccountId);
      if (crAcc?.type === 'revenue') kasMasukOperasional += t.amount;
      else if (crAcc?.type === 'equity') kasMasukPendanaan += t.amount;
      else kasMasukOperasional += t.amount;
    }
    if (t.creditAccountId === '101') {
      kasKeluarOperasional += t.amount;
    }
  });

  const kasBersihOperasional = kasMasukOperasional - kasKeluarOperasional;
  const kenaikanKas = kasBersihOperasional + kasMasukPendanaan;

  const handleDownloadExcel = () => {
    const data = [
      { label: 'ARUS KAS DARI AKTIVITAS OPERASI', amount: null, isHeader: true, isBold: true },
      { label: '   Penerimaan Kas dari Pelanggan/Pendapatan', amount: kasMasukOperasional, isBold: false },
      { label: '   Pembayaran Kas untuk Beban/Operasional', amount: -kasKeluarOperasional, isBold: false },
      { label: 'Kas Bersih dari Aktivitas Operasi', amount: kasBersihOperasional, isBold: true, isTotal: false },
      { label: '', amount: null },
      { label: 'ARUS KAS DARI AKTIVITAS PENDANAAN', amount: null, isHeader: true, isBold: true },
      { label: '   Penerimaan Modal Saham', amount: kasMasukPendanaan, isBold: false },
      { label: 'Kas Bersih dari Aktivitas Pendanaan', amount: kasMasukPendanaan, isBold: true, isTotal: false },
      { label: '', amount: null },
      { label: 'KENAIKAN (PENURUNAN) KAS BERSIH', amount: kenaikanKas, isBold: true, isTotal: true }
    ];
    exportFinancialToExcel('Arus Kas', company.name, data, `Arus_Kas_${company.name}.xlsx`);
  };

  return (
    <div className="space-y-8 animate-fade-in">
      <div className="flex justify-end gap-4 print:hidden">
        <button onClick={handleDownloadExcel} className="bg-blue-600 hover:bg-blue-500 border border-blue-500 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Download size={14}/> Unduh Excel Rapi</button>
        <button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 border border-slate-700 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Printer size={14}/> Cetak RUPS</button>
      </div>
      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-4xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[800px] font-sans border border-slate-200">
        <PrintLetterhead company={company} />
        <div className="text-center pb-6 mb-10">
          <h3 className="text-lg font-bold mt-2 text-slate-700 uppercase tracking-widest">Laporan Arus Kas (Metode Langsung)</h3>
          <p className="text-sm text-slate-500 font-medium mt-1">Periode Berakhir Per: {new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
        </div>

        <div className="space-y-8 text-sm md:text-base">
          <div>
            <h4 className="font-bold text-slate-900 bg-slate-100 p-2 uppercase tracking-widest border-l-4 border-slate-900">ARUS KAS DARI AKTIVITAS OPERASI</h4>
            <div className="space-y-3 mt-4 pl-4 pr-2">
              <div className="flex justify-between text-slate-700"><span>Penerimaan Kas dari Pelanggan/Operasional</span><span className="font-mono">{formatCurrency(kasMasukOperasional)}</span></div>
              <div className="flex justify-between text-slate-700"><span>Pembayaran Kas untuk Beban/Pembelian</span><span className="font-mono">({formatCurrency(kasKeluarOperasional)})</span></div>
            </div>
            <div className="flex justify-between font-black mt-4 pt-3 border-t-2 border-slate-900 text-slate-900 pl-4 pr-2"><span>Kas Bersih dari Aktivitas Operasi</span><span className="font-mono">{formatCurrency(kasBersihOperasional)}</span></div>
          </div>

          <div>
            <h4 className="font-bold text-slate-900 bg-slate-100 p-2 uppercase tracking-widest border-l-4 border-slate-900">ARUS KAS DARI AKTIVITAS PENDANAAN</h4>
            <div className="space-y-3 mt-4 pl-4 pr-2">
              <div className="flex justify-between text-slate-700"><span>Penerimaan Setoran Modal</span><span className="font-mono">{formatCurrency(kasMasukPendanaan)}</span></div>
            </div>
            <div className="flex justify-between font-black mt-4 pt-3 border-t-2 border-slate-900 text-slate-900 pl-4 pr-2"><span>Kas Bersih dari Aktivitas Pendanaan</span><span className="font-mono">{formatCurrency(kasMasukPendanaan)}</span></div>
          </div>

          <div className="pt-6 border-t-[6px] border-double border-slate-900 flex justify-between font-black text-xl text-slate-900 p-2">
            <span>KENAIKAN (PENURUNAN) KAS BERSIH</span>
            <span className={`font-mono ${kenaikanKas < 0 ? 'text-rose-600' : 'text-slate-900'}`}>{formatCurrency(kenaikanKas)}</span>
          </div>
        </div>

        <div className="mt-32 grid grid-cols-2 gap-8 text-center text-sm font-medium print:mt-40">
          <div><p className="mb-24 text-slate-500">Dibuat Oleh,</p><p className="font-bold border-b border-slate-900 inline-block px-10 pb-1 text-slate-900">DIREKSI KEUANGAN</p></div>
          <div><p className="mb-24 text-slate-500">Disahkan Oleh,</p><p className="font-bold border-b border-slate-900 inline-block px-10 pb-1 text-slate-900">DIREKTUR UTAMA</p></div>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   9. KONFIGURASI PT
   ======================================================================== */
function SettingsView({ company, setCompany, accounts, saveToCloud, user }) {
  const [formData, setFormData] = useState(company);

  const handleLogoUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFormData({ ...formData, logo: reader.result });
      };
      reader.readAsDataURL(file);
    }
  };

  const handleSave = (e) => {
    e.preventDefault();
    if(user.role !== 'admin') return alert("Akses Terkunci!");
    setCompany(formData);
    saveToCloud(null, null, formData);
  };

  return (
    <div className="max-w-3xl bg-slate-900 p-10 rounded-3xl border border-slate-800 shadow-2xl animate-fade-in">
      <div className="mb-8 border-b border-slate-800 pb-4">
        <h3 className="text-xl font-black text-white uppercase tracking-widest">Konfigurasi Identitas Korporasi</h3>
        <p className="text-slate-400 text-xs mt-2">Data ini akan tercetak di seluruh dokumen resmi & laporan pajak.</p>
      </div>
      <form onSubmit={handleSave} className="space-y-6">
        <div className="flex items-center gap-6 p-5 border border-slate-700 bg-slate-950 rounded-xl">
          {formData.logo ? (
            <img src={formData.logo} alt="Logo" className="w-24 h-24 object-contain rounded-lg bg-white p-1" />
          ) : (
            <div className="w-24 h-24 border-2 border-dashed border-slate-600 rounded-lg flex items-center justify-center text-slate-500 text-xs text-center p-2">Belum ada logo</div>
          )}
          <div className="flex-1">
            <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Upload Logo Perusahaan (Kop Surat)</label>
            <input type="file" accept="image/*" onChange={handleLogoUpload} className="w-full text-sm text-slate-400 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-xs file:font-bold file:bg-emerald-500/10 file:text-emerald-400 hover:file:bg-emerald-500/20 cursor-pointer" />
            <p className="text-[9px] text-slate-500 mt-2">Gunakan rasio persegi panjang/kotak format transparan (PNG) untuk hasil cetak dokumen yang terbaik.</p>
          </div>
        </div>

        <div><label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Nama Resmi PT</label><input type="text" className="w-full bg-slate-950 border border-slate-700 p-4 rounded-xl text-white text-sm focus:border-emerald-500 outline-none" value={formData.name} onChange={e=>setFormData({...formData, name: e.target.value})} required /></div>
        <div><label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Alamat Kantor Pusat</label><textarea className="w-full bg-slate-950 border border-slate-700 p-4 rounded-xl text-white text-sm h-24 focus:border-emerald-500 outline-none" value={formData.address} onChange={e=>setFormData({...formData, address: e.target.value})} required /></div>
        <div className="grid grid-cols-2 gap-6">
          <div><label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Nomor Telepon</label><input type="text" className="w-full bg-slate-950 border border-slate-700 p-4 rounded-xl text-white text-sm focus:border-emerald-500 outline-none" value={formData.phone} onChange={e=>setFormData({...formData, phone: e.target.value})} /></div>
          <div><label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Email Resmi</label><input type="email" className="w-full bg-slate-950 border border-slate-700 p-4 rounded-xl text-white text-sm focus:border-emerald-500 outline-none" value={formData.email} onChange={e=>setFormData({...formData, email: e.target.value})} /></div>
        </div>
        <div><label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">NPWP Badan Korporasi</label><input type="text" className="w-full bg-slate-950 border border-slate-700 p-4 rounded-xl text-white text-sm font-mono focus:border-emerald-500 outline-none" value={formData.npwp} onChange={e=>setFormData({...formData, npwp: e.target.value})} /></div>
        
        <div className="pt-6 border-t border-slate-800">
          <button type="submit" className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 px-8 py-4 rounded-xl font-black text-xs uppercase tracking-widest text-white shadow-lg shadow-emerald-900/50 w-full md:w-auto transition-all active:scale-95">Simpan Perubahan Sistem</button>
        </div>
      </form>
    </div>
  );
}
 /*BERITA ACARA RUPS (BISA DI ISI & PRINT)

   ======================================================================== */

function BeritaAcaraView({ company }) {

  const [formData, setFormData] = useState({

    hari: 'Senin',

    tanggal: new Date().toISOString().split('T')[0],

    waktu: '10:00',

    tempat: company.address,

    pimpinan: 'Bpk. Direktur Utama',

    sekretaris: 'Ibu Sekretaris Perusahaan',

    kuorum: '100',

    keputusan: '1. Menyetujui dan mengesahkan Laporan Keuangan Tahunan Perusahaan.\n2. Memberikan pelunasan dan pembebasan tanggung jawab (acquit et decharge) kepada Direksi atas tindakan pengurusan selama tahun buku berjalan.\n3. Menetapkan penggunaan laba bersih perusahaan untuk ditahan sebagai modal kerja.'

  });



  const handleChange = (e) => setFormData({...formData, [e.target.name]: e.target.value});



  const getDayName = (dateStr) => {

    const date = new Date(dateStr);

    const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

    return days[date.getDay()];

  };



  const formatDateIndo = (dateStr) => {

    return new Date(dateStr).toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' });

  };



  return (

    <div className="space-y-8 animate-fade-in relative z-10">

      <div className="flex justify-between items-end print:hidden">

        <div>

          <h2 className="text-2xl font-black text-white uppercase tracking-wider">Draft Berita Acara RUPS</h2>

          <p className="text-slate-400 text-xs font-medium mt-1">Isi formulir di bawah ini, dokumen legal akan ter-generate otomatis saat dicetak.</p>

        </div>

        <button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 border border-slate-700 px-6 py-3 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Printer size={16}/> Cetak Dokumen Legal</button>

      </div>



      {/* Form Input (Hanya tampil di layar) */}

      <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl grid grid-cols-1 md:grid-cols-3 gap-6 print:hidden">

        <div><label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Tanggal Rapat</label><input type="date" name="tanggal" value={formData.tanggal} onChange={handleChange} className="w-full bg-slate-950 border border-slate-700 p-3 rounded-xl text-white text-sm" /></div>

        <div><label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Waktu (WIB)</label><input type="time" name="waktu" value={formData.waktu} onChange={handleChange} className="w-full bg-slate-950 border border-slate-700 p-3 rounded-xl text-white text-sm" /></div>

        <div><label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Tempat Pelaksanaan</label><input type="text" name="tempat" value={formData.tempat} onChange={handleChange} className="w-full bg-slate-950 border border-slate-700 p-3 rounded-xl text-white text-sm" /></div>

        <div><label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Pimpinan Rapat</label><input type="text" name="pimpinan" value={formData.pimpinan} onChange={handleChange} className="w-full bg-slate-950 border border-slate-700 p-3 rounded-xl text-white text-sm" /></div>

        <div><label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Sekretaris Rapat</label><input type="text" name="sekretaris" value={formData.sekretaris} onChange={handleChange} className="w-full bg-slate-950 border border-slate-700 p-3 rounded-xl text-white text-sm" /></div>

        <div><label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Kehadiran Kuorum (%)</label><input type="number" name="kuorum" value={formData.kuorum} onChange={handleChange} className="w-full bg-slate-950 border border-slate-700 p-3 rounded-xl text-white text-sm" /></div>

        <div className="md:col-span-3"><label className="block text-[10px] text-slate-400 font-bold uppercase tracking-widest mb-2">Hasil Keputusan (Pisahkan dengan Enter)</label><textarea name="keputusan" value={formData.keputusan} onChange={handleChange} className="w-full bg-slate-950 border border-slate-700 p-3 rounded-xl text-white text-sm h-32" /></div>

      </div>



      {/* Kertas Cetak A4 */}

      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-4xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[1122px] font-sans border border-slate-200 text-justify">

        <PrintLetterhead company={company} />

       

        <div className="text-center mb-10">

          <h3 className="text-xl font-black text-slate-900 uppercase tracking-widest underline underline-offset-4">BERITA ACARA RAPAT UMUM PEMEGANG SAHAM</h3>

          <p className="text-sm font-bold text-slate-700 mt-2 uppercase">{company.name}</p>

        </div>



        <div className="space-y-6 text-base leading-relaxed">

          <p>

            Pada hari ini, <strong>{getDayName(formData.tanggal)}</strong>, tanggal <strong>{formatDateIndo(formData.tanggal)}</strong>, bertempat di <strong>{formData.tempat}</strong>, telah diselenggarakan Rapat Umum Pemegang Saham (selanjutnya disebut "Rapat") dari <strong>{company.name}</strong>.

          </p>

         

          <div>

            <p className="mb-2">Rapat dibuka pada pukul <strong>{formData.waktu} WIB</strong> dan dipimpin oleh:</p>

            <table className="ml-8 w-full">

              <tbody>

                <tr><td className="w-48 py-1">Pimpinan Rapat</td><td className="w-4">:</td><td className="font-bold">{formData.pimpinan}</td></tr>

                <tr><td className="w-48 py-1">Sekretaris Rapat</td><td className="w-4">:</td><td className="font-bold">{formData.sekretaris}</td></tr>

              </tbody>

            </table>

          </div>



          <p>

            Pimpinan Rapat melaporkan bahwa pemegang saham yang hadir atau diwakili dalam Rapat ini mewakili sejumlah <strong>{formData.kuorum}%</strong> dari seluruh saham dengan hak suara yang sah yang telah dikeluarkan oleh Perusahaan. Dengan demikian, kuorum Rapat telah tercapai dan Rapat sah serta berhak mengambil keputusan yang mengikat.

          </p>



          <div>

            <p className="font-bold uppercase tracking-widest mb-4 border-b-2 border-slate-900 inline-block">Mata Acara & Keputusan Rapat:</p>

            <p className="mb-4">Setelah menelaah Laporan Keuangan (Neraca, Laba/Rugi, Arus Kas, dan Perubahan Ekuitas) dan mendengarkan presentasi kinerja, Rapat secara musyawarah dan mufakat memutuskan hal-hal sebagai berikut:</p>

            <div className="ml-8 space-y-3">

              {formData.keputusan.split('\n').map((item, index) => (

                <p key={index} className="pl-4 -indent-4">{item}</p>

              ))}

            </div>

          </div>



          <p className="pt-4">

            Demikian Berita Acara Rapat ini dibuat dengan sebenar-benarnya untuk dapat dipergunakan sebagaimana mestinya. Rapat ditutup pada pukul {String(parseInt(formData.waktu.split(':')[0]) + 2).padStart(2, '0')}:{formData.waktu.split(':')[1]} WIB.

          </p>

        </div>



        {/* Kolom Tanda Tangan */}

        <div className="mt-20 flex justify-between text-center">

          <div className="w-64">

            <p className="mb-24">Pimpinan Rapat,</p>

            <p className="font-bold uppercase border-b border-slate-900 pb-1">{formData.pimpinan}</p>

          </div>

          <div className="w-64">

            <p className="mb-24">Sekretaris Rapat,</p>

            <p className="font-bold uppercase border-b border-slate-900 pb-1">{formData.sekretaris}</p>

          </div>

        </div>

      </div>

    </div>

  );
function BeritaAcaraView({ company }) {
  const [formData, setFormData] = useState({
    tanggal: new Date().toISOString().split('T')[0],
    pimpinan: 'Direktur Utama',
    sekretaris: 'Sekretaris Perusahaan',
    keputusan: '1. Mengesahkan Laporan Keuangan.\n2. Menetapkan laba ditahan.'
  });

  return (
    <div className="space-y-6 animate-fade-in">
      <div className="flex justify-between items-center print:hidden">
        <h2 className="text-xl font-bold text-white uppercase">Berita Acara RUPS</h2>
        <button onClick={() => window.print()} className="bg-slate-800 px-4 py-2 rounded-lg text-white text-xs flex items-center gap-2"><Printer size={16}/> Cetak Dokumen</button>
      </div>
      
      {/* Form Input */}
      <div className="bg-slate-900 p-6 rounded-xl grid grid-cols-2 gap-4 print:hidden">
        <input type="date" className="bg-slate-950 p-2 rounded text-white text-sm" value={formData.tanggal} onChange={e => setFormData({...formData, tanggal: e.target.value})} />
        <input type="text" className="bg-slate-950 p-2 rounded text-white text-sm" placeholder="Pimpinan Rapat" value={formData.pimpinan} onChange={e => setFormData({...formData, pimpinan: e.target.value})} />
        <textarea className="col-span-2 bg-slate-950 p-2 rounded text-white text-sm h-24" placeholder="Hasil Keputusan" value={formData.keputusan} onChange={e => setFormData({...formData, keputusan: e.target.value})} />
      </div>

      {/* Dokumen Cetak */}
      <div className="bg-white text-black p-12 max-w-2xl mx-auto shadow-2xl min-h-[800px]">
        <h3 className="text-center font-bold underline mb-8">BERITA ACARA RUPS</h3>
        <p>Pada hari ini, {new Date(formData.tanggal).toLocaleDateString('id-ID')}, bertempat di {company.address}, telah dilaksanakan RUPS {company.name}.</p>
        <p className="mt-4">Rapat dipimpin oleh {formData.pimpinan} dan dihadiri oleh jajaran pemegang saham.</p>
        <div className="mt-6">
          <p className="font-bold">Keputusan Rapat:</p>
          <pre className="font-sans whitespace-pre-wrap">{formData.keputusan}</pre>
        </div>
        <div className="mt-12 flex justify-between text-center">
          <div><p>Pimpinan,</p><br/><br/><p className="font-bold border-b">{formData.pimpinan}</p></div>
          <div><p>Sekretaris,</p><br/><br/><p className="font-bold border-b">{formData.sekretaris}</p></div>
        </div>
      </div>
    </div>
  );
}
}
