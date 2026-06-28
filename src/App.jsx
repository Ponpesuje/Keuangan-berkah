import React, { useState, useEffect, useMemo } from 'react';
import {
  LayoutDashboard, Receipt, BookOpen, Settings, LogOut,
  Plus, Printer, Download, UserCircle, Building2, Wallet, TrendingUp,
  Lock, Trash2, FileText, Landmark, Filter, Shield, HelpCircle,
  ChevronRight, AlertCircle, RefreshCw, X, Presentation, FileSignature, Users, ArrowRight, ArrowLeft, CheckCircle2, ImagePlus, Edit2
} from 'lucide-react';
import { AreaChart, Area, CartesianGrid, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend, BarChart, Bar } from 'recharts';

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
   FUNGSI EXPORT EXCEL PROFESIONAL (Tanpa Library Tambahan)
   ======================================================================== */
const exportGridToExcel = (sheetName, columns, data, filename, companyName) => {
  let tableHtml = `<html xmlns:x="urn:schemas-microsoft-com:office:excel">
    <head><meta charset="utf-8"></head><body>
    <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif;">
      <tr><td colspan="${columns.length}" style="font-size: 14pt; font-weight: bold; text-align: center; background-color: #f8fafc;">${companyName.toUpperCase()}</td></tr>
      <tr><td colspan="${columns.length}" style="font-size: 11pt; font-weight: bold; text-align: center; background-color: #f8fafc;">LAPORAN ${sheetName.toUpperCase()}</td></tr>
      <tr><td colspan="${columns.length}"></td></tr>
      <tr style="background-color: #0f172a; color: #ffffff; font-weight: bold;">
        ${columns.map(c => `<th>${c.header}</th>`).join('')}
      </tr>`;

  data.forEach(row => {
    tableHtml += `<tr>`;
    columns.forEach(col => {
      let val = row[col.key] || '';
      if (typeof val === 'number') {
         tableHtml += `<td style="text-align: right;">${val}</td>`;
      } else {
         tableHtml += `<td>${val}</td>`;
      }
    });
    tableHtml += `</tr>`;
  });

  tableHtml += `</table></body></html>`;

  const blob = new Blob([tableHtml], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename.replace('.xlsx', '.xls');
  a.click();
  URL.revokeObjectURL(url);
};

const exportFinancialToExcel = (title, companyName, rowsData, filename) => {
  let tableHtml = `<html xmlns:x="urn:schemas-microsoft-com:office:excel">
    <head><meta charset="utf-8"></head><body>
    <table border="1" style="border-collapse: collapse; font-family: Arial, sans-serif;">
      <tr><td colspan="2" style="font-size: 14pt; font-weight: bold; text-align: center; background-color: #f8fafc;">${companyName.toUpperCase()}</td></tr>
      <tr><td colspan="2" style="font-size: 12pt; font-weight: bold; text-align: center; background-color: #f8fafc;">LAPORAN ${title.toUpperCase()}</td></tr>
      <tr><td colspan="2"></td></tr>`;

  rowsData.forEach(r => {
    let style = "";
    if (r.isHeader) style += "background-color: #f1f5f9; ";
    if (r.isBold) style += "font-weight: bold; ";
    if (r.isTotal) style += "border-bottom: 3px double #000; border-top: 1px solid #000; ";

    tableHtml += `<tr style="${style}">`;
    tableHtml += `<td>${r.label}</td>`;
    if (r.amount !== null) {
      tableHtml += `<td style="text-align: right;">${r.amount}</td>`;
    } else {
      tableHtml += `<td></td>`;
    }
    tableHtml += `</tr>`;
  });

  tableHtml += `</table></body></html>`;

  const blob = new Blob([tableHtml], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename.replace('.xlsx', '.xls');
  a.click();
  URL.revokeObjectURL(url);
};

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
  const [isLoading, setIsLoading] = useState(false);

  const [company, setCompany] = useState(INITIAL_COMPANY);
  const [accounts, setAccounts] = useState(INITIAL_ACCOUNTS);
  const [transactions, setTransactions] = useState([]);

  useEffect(() => {
    const loadData = async () => {
      setIsLoading(true);
      try {
        const res = await fetch(GOOGLE_SHEET_URL);
        const data = await res.json();
        if (data.transactions) setTransactions(data.transactions);
        if (data.accounts) setAccounts(data.accounts);
        if (data.company) setCompany({...INITIAL_COMPANY, ...data.company});
      } catch (e) { console.log("Menggunakan data lokal."); }
      setIsLoading(false);
    };
    loadData();
  }, []);

  const saveToCloud = async (newTransactions, newAccounts, newCompany) => {
    if (currentUser?.role !== 'admin') return;
    setSyncStatus('Menyimpan...');
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
            {company.logo ? (
               <img src={company.logo} alt="Logo Perusahaan" className="h-16 mx-auto mb-4 object-contain" />
            ) : (
               <Building2 size={44} className="mx-auto text-amber-500 mb-3" />
            )}
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
        <div className="p-8 border-b border-slate-800/60 bg-slate-900/80 backdrop-blur-md relative z-10 text-center">
          {company.logo ? (
             <img src={company.logo} alt="Logo" className="h-12 mx-auto mb-3 object-contain drop-shadow-lg" />
          ) : (
             <Building2 size={32} className="mx-auto text-emerald-500 drop-shadow-[0_0_8px_rgba(16,185,129,0.5)] mb-2" />
          )}
          <h1 className="font-black text-sm tracking-widest uppercase truncate text-emerald-400">{company.name}</h1>
          <span className="mt-2 text-[9px] bg-slate-950/50 px-3 py-1 rounded-full inline-block font-mono text-emerald-500 uppercase tracking-widest border border-emerald-900/50 shadow-inner">Akses: {currentUser.role}</span>
        </div>
        
        <nav className="flex-1 p-6 space-y-2 overflow-y-auto relative z-10 scrollbar-thin scrollbar-thumb-slate-800 scrollbar-track-transparent">
          <NavItem icon={<LayoutDashboard size={18}/>} label="Dashboard" active={currentView === 'dashboard'} onClick={() => setCurrentView('dashboard')} />
          
          <div className="pt-4 pb-2">
            <p className="text-[10px] font-black text-slate-600 uppercase tracking-widest px-4">Aktivitas & Siklus</p>
          </div>
          <NavItem icon={<Receipt size={18}/>} label="Transaksi & Invoice" active={currentView === 'transactions'} onClick={() => setCurrentView('transactions')} />
          <NavItem icon={<Wallet size={18}/>} label="Buku Kas Besar" active={currentView === 'buku-kas'} onClick={() => setCurrentView('buku-kas')} />
          <NavItem icon={<BookOpen size={18}/>} label="Jurnal Umum" active={currentView === 'journal'} onClick={() => setCurrentView('journal')} />
          
          <div className="pt-4 pb-2">
            <p className="text-[10px] font-black text-slate-600 uppercase tracking-widest px-4">Laporan Keuangan</p>
          </div>
          <NavItem icon={<TrendingUp size={18}/>} label="Laba Rugi" active={currentView === 'laba-rugi'} onClick={() => setCurrentView('laba-rugi')} />
          <NavItem icon={<FileText size={18}/>} label="Neraca" active={currentView === 'neraca'} onClick={() => setCurrentView('neraca')} />
          <NavItem icon={<HelpCircle size={18}/>} label="Catatan CALK" active={currentView === 'calk'} onClick={() => setCurrentView('calk')} />
          <NavItem icon={<Landmark size={18}/>} label="Pajak Badan" active={currentView === 'laporan-pajak'} onClick={() => setCurrentView('laporan-pajak')} />
          
          <div className="pt-4 pb-2">
             <p className="text-[10px] font-black text-slate-600 uppercase tracking-widest px-4">Dokumen RUPS</p>
          </div>
          <NavItem icon={<Presentation size={18}/>} label="Presentasi Layar" active={currentView === 'presentasi'} onClick={() => setCurrentView('presentasi')} />
          <NavItem icon={<FileSignature size={18}/>} label="Berita Acara" active={currentView === 'berita-acara'} onClick={() => setCurrentView('berita-acara')} />

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
          <div className="flex items-center gap-4">
             {isLoading && <span className="text-[10px] font-mono text-emerald-500 flex items-center gap-2 animate-pulse"><RefreshCw size={12} className="animate-spin"/> Sinkronisasi Cloud...</span>}
             <div className="flex items-center gap-4 bg-slate-900/80 px-4 py-2 rounded-full border border-slate-800">
              <div className="w-8 h-8 rounded-full bg-slate-800 flex items-center justify-center text-slate-400"><UserCircle size={18} /></div>
              <div className="text-right pr-2">
                <p className="text-xs font-bold text-white">{currentUser.username}</p>
                <p className="text-[9px] font-mono text-emerald-400 uppercase tracking-widest">Active Session</p>
              </div>
            </div>
          </div>
        </header>
        
        <div className="flex-1 overflow-y-auto p-6 md:p-10 z-10 print:p-0 print:bg-white print:text-black">
          {currentView === 'dashboard' && <DashboardView transactions={transactions} accounts={accounts} balances={balances} />}
          {currentView === 'transactions' && <TransactionsView transactions={transactions} accounts={accounts} setTransactions={setTransactions} company={company} saveToCloud={saveToCloud} user={currentUser} />}
          {currentView === 'buku-kas' && <BukuKasView transactions={transactions} accounts={accounts} company={company} />}
          {currentView === 'journal' && <JournalView transactions={transactions} accounts={accounts} company={company} />}
          {currentView === 'laba-rugi' && <LabaRugiView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'neraca' && <NeracaView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'calk' && <CalkView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'laporan-pajak' && <LaporanPajakView accounts={accounts} balances={balances} company={company} />}
          {currentView === 'presentasi' && <PresentasiRupsView transactions={transactions} accounts={accounts} balances={balances} company={company} />}
          {currentView === 'berita-acara' && <BeritaAcaraView company={company} />}
          {currentView === 'settings' && <SettingsView company={company} setCompany={setCompany} accounts={accounts} saveToCloud={saveToCloud} user={currentUser} />}
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
   1. DASHBOARD VIEW
   ======================================================================== */
function DashboardView({ transactions, accounts, balances }) {
  const cashAccounts = accounts.filter(a => a.name.toLowerCase().includes('kas') || a.name.toLowerCase().includes('bank'));
  const totalCash = cashAccounts.reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const totalRevenue = accounts.filter(a => a.type === 'revenue').reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const totalExpense = accounts.filter(a => a.type === 'expense').reduce((sum, a) => sum + (balances[a.id] || 0), 0);
  const netProfit = totalRevenue - totalExpense;

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

        <div className="bg-gradient-to-br from-slate-900 to-slate-950 border border-slate-800 p-8 rounded-3xl shadow-xl flex flex-col justify-between">
          <div>
            <h3 className="text-xs font-bold uppercase text-slate-400 tracking-widest mb-6">Ringkasan Eksekutif</h3>
            <div className="space-y-6">
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
          <div className="mt-8 pt-6 border-t border-slate-800/60">
            <p className="text-[10px] text-slate-500 italic text-justify leading-relaxed">Seluruh data yang disajikan dilindungi oleh enkripsi end-to-end dan disinkronisasi secara real-time dengan server cloud korporasi.</p>
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
   2. TRANSAKSI & INVOICE MULTI PILIHAN (DENGAN FITUR EDIT)
   ======================================================================== */
function TransactionsView({ transactions, accounts, setTransactions, company, saveToCloud, user }) {
  const [showForm, setShowForm] = useState(false);
  const [editingTrxId, setEditingTrxId] = useState(null); // State untuk Edit
  const [trxType, setTrxType] = useState('pemasukan');
  const [formData, setFormData] = useState({ date: new Date().toISOString().split('T')[0], description: '', amount: '', accountId: '' });
  
  const [selectedTrxIds, setSelectedTrxIds] = useState([]);
  const [invoiceTransactions, setInvoiceTransactions] = useState(null);

  const toggleSelectTrx = (id) => {
    if (selectedTrxIds.includes(id)) {
      setSelectedTrxIds(selectedTrxIds.filter(item => item !== id));
    } else {
      setSelectedTrxIds([...selectedTrxIds, id]);
    }
  };

  const handlePrintSelectedInvoice = () => {
    const selected = transactions.filter(t => selectedTrxIds.includes(t.id));
    setInvoiceTransactions(selected);
  };

  if (invoiceTransactions) {
    return <InvoiceView transactions={invoiceTransactions} accounts={accounts} company={company} onBack={() => setInvoiceTransactions(null)} />;
  }

  const executeDelete = (id) => {
    if (user.role !== 'admin') return;
    if (window.confirm("Hapus transaksi secara permanen?")) {
      const updated = transactions.filter(t => t.id !== id);
      setTransactions(updated);
      saveToCloud(updated, null, null);
      setSelectedTrxIds(selectedTrxIds.filter(item => item !== id));
    }
  };

  // Fungsi Saat Tombol Edit Ditekan
  const executeEdit = (trx) => {
    if (user.role !== 'admin') return;
    
    // Menentukan jenis transaksi (Pemasukan / Pengeluaran)
    const isPemasukan = trx.debitAccountId === '101';
    setTrxType(isPemasukan ? 'pemasukan' : 'pengeluaran');
    
    // Memasukkan data transaksi lama ke form
    setFormData({
      date: trx.date,
      description: trx.description,
      amount: trx.amount,
      accountId: isPemasukan ? trx.creditAccountId : trx.debitAccountId
    });
    
    setEditingTrxId(trx.id); // Set ID transaksi yang sedang diubah
    setShowForm(true); // Buka form
  };

  const handleCloseForm = () => {
    setShowForm(false);
    setEditingTrxId(null);
    setFormData({ date: new Date().toISOString().split('T')[0], description: '', amount: '', accountId: '' });
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    if (user.role !== 'admin') return;
    const amountNum = parseFloat(formData.amount);
    if (!formData.accountId || isNaN(amountNum) || amountNum <= 0) return;

    const newTrx = {
      id: editingTrxId || `TRX-${Date.now()}`, // Gunakan ID lama jika Edit, ID baru jika Tambah
      date: formData.date,
      description: formData.description,
      amount: amountNum,
      debitAccountId: trxType === 'pemasukan' ? '101' : formData.accountId,
      creditAccountId: trxType === 'pemasukan' ? formData.accountId : '101'
    };

    let updated;
    if (editingTrxId) {
      // Perbarui transaksi yang ada
      updated = transactions.map(t => t.id === editingTrxId ? newTrx : t);
    } else {
      // Tambah transaksi baru
      updated = [...transactions, newTrx];
    }
    
    setTransactions(updated);
    saveToCloud(updated, null, null);
    handleCloseForm();
  };

  return (
    <div className="space-y-8 animate-fade-in">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-2xl font-black text-white uppercase tracking-wider">Transaksi & Invoice</h2>
          <p className="text-slate-400 text-xs font-medium mt-1">Centang kotak untuk menggabungkan beberapa transaksi ke dalam 1 Invoice.</p>
        </div>
        <div className="flex gap-3">
           {selectedTrxIds.length > 0 && (
             <button onClick={handlePrintSelectedInvoice} className="bg-blue-600 hover:bg-blue-500 px-6 py-3.5 rounded-xl text-white font-bold text-xs uppercase tracking-widest flex items-center gap-2 shadow-lg shadow-blue-900/40 transition-all animate-fade-in">
               <Printer size={16}/> Cetak {selectedTrxIds.length} Invoice
             </button>
           )}
          {user.role === 'admin' ? (
            <button onClick={showForm ? handleCloseForm : () => setShowForm(true)} className="bg-emerald-600 hover:bg-emerald-500 px-6 py-3.5 rounded-xl text-white font-bold text-xs uppercase tracking-widest flex items-center gap-2 shadow-lg shadow-emerald-900/40 transition-all active:scale-95">
              {showForm ? <X size={16}/> : <Plus size={16}/>} {showForm ? 'Batal / Tutup' : 'Entri Kas Baru'}
            </button>
          ) : (
            <div className="flex items-center gap-2 bg-amber-500/10 border border-amber-500/20 text-amber-400 px-4 py-2.5 rounded-xl text-xs font-bold uppercase tracking-wider">
              <Lock size={14} /> Akses Input Terkunci
            </div>
          )}
        </div>
      </div>

      {showForm && user.role === 'admin' && (
        <div className="bg-slate-900 border border-slate-800 p-8 rounded-3xl shadow-2xl max-w-3xl">
          <div className="mb-6 flex items-center gap-3">
             {editingTrxId ? <Edit2 size={24} className="text-blue-500"/> : <Plus size={24} className="text-emerald-500"/>}
             <h3 className="text-lg font-black text-white uppercase tracking-widest">{editingTrxId ? 'Ubah / Edit Transaksi' : 'Catat Transaksi Baru'}</h3>
          </div>
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
            <button type="submit" className={`w-full text-white py-4 rounded-xl font-black text-xs uppercase tracking-widest shadow-lg transition-all mt-4 ${editingTrxId ? 'bg-blue-600 hover:bg-blue-500 shadow-blue-900/40' : 'bg-emerald-600 hover:bg-emerald-500 shadow-emerald-900/40'}`}>
              {editingTrxId ? 'Simpan Perubahan' : 'Simpan ke Buku Besar'}
            </button>
          </form>
        </div>
      )}

      <div className="bg-slate-900 border border-slate-800 rounded-3xl overflow-hidden shadow-xl">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm whitespace-nowrap">
            <thead className="bg-slate-950 text-slate-400 border-b border-slate-800">
              <tr>
                <th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-center w-10">Pilih</th>
                <th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Tanggal</th>
                <th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Keterangan</th>
                <th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Debet</th>
                <th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest">Kredit</th>
                <th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-right">Nominal</th>
                {user.role === 'admin' && <th className="px-6 py-4 font-bold text-[10px] uppercase tracking-widest text-center">Aksi</th>}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800/50">
              {[...transactions].reverse().map(t => (
                <tr key={t.id} className={`hover:bg-slate-800/30 transition-colors group cursor-pointer ${selectedTrxIds.includes(t.id) ? 'bg-emerald-900/10' : ''}`} onClick={(e) => {
                   if(e.target.tagName !== 'INPUT' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'svg' && e.target.tagName !== 'path') {
                      toggleSelectTrx(t.id);
                   }
                }}>
                  <td className="px-6 py-4 text-center">
                    <input 
                      type="checkbox" 
                      className="w-4 h-4 rounded border-slate-700 text-emerald-500 focus:ring-emerald-500 cursor-pointer"
                      checked={selectedTrxIds.includes(t.id)}
                      onChange={() => toggleSelectTrx(t.id)}
                    />
                  </td>
                  <td className="px-6 py-4 text-slate-300">{t.date} <br/><span className="text-[9px] font-mono text-slate-600">{t.id}</span></td>
                  <td className="px-6 py-4 text-white font-medium">{t.description}</td>
                  <td className="px-6 py-4 text-emerald-400 text-xs font-semibold">{accounts.find(a=>a.id===t.debitAccountId)?.name}</td>
                  <td className="px-6 py-4 text-rose-400 text-xs font-semibold">{accounts.find(a=>a.id===t.creditAccountId)?.name}</td>
                  <td className="px-6 py-4 text-right font-bold text-white font-mono">{formatCurrency(t.amount)}</td>
                  {user.role === 'admin' && (
                    <td className="px-6 py-4 text-center flex justify-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                      <button onClick={(e) => { e.stopPropagation(); executeEdit(t); }} className="text-slate-500 hover:text-blue-400 hover:bg-blue-500/10 p-2 rounded-lg transition-all" title="Edit Transaksi">
                        <Edit2 size={16}/>
                      </button>
                      <button onClick={(e) => { e.stopPropagation(); executeDelete(t.id); }} className="text-slate-500 hover:text-rose-400 hover:bg-rose-500/10 p-2 rounded-lg transition-all" title="Hapus Transaksi">
                        <Trash2 size={16}/>
                      </button>
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

/* KOMPONEN TAMPILAN INVOICE GABUNGAN (VERSI BERSIH TANPA AKUN D/K) */
function InvoiceView({ transactions, accounts, company, onBack }) {
  const totalAmount = transactions.reduce((sum, t) => sum + t.amount, 0);

  return (
    <div className="space-y-6 animate-fade-in">
      <div className="flex justify-between items-center no-print">
        <button onClick={onBack} className="flex items-center gap-2 text-slate-400 hover:text-white transition-colors bg-slate-800 px-4 py-2 rounded-xl text-xs font-bold uppercase tracking-widest">
          <ArrowLeft size={16} /> Kembali
        </button>
        <button onClick={() => window.print()} className="flex items-center gap-2 px-6 py-3 bg-blue-600 hover:bg-blue-500 text-white rounded-xl transition-all text-xs font-bold uppercase tracking-widest shadow-lg shadow-blue-900/40">
          <Printer size={16} /> Cetak & Download PDF
        </button>
      </div>

      <div className="bg-white text-black p-12 md:p-16 rounded-3xl mx-auto max-w-4xl shadow-2xl print:shadow-none print:m-0 print:p-0 min-h-[800px] border border-slate-200">
        
        {/* Header Invoice dengan Logo */}
        <div className="flex justify-between items-start border-b-4 border-slate-900 pb-8 mb-10">
          <div className="flex gap-4 items-start">
             {company.logo && <img src={company.logo} alt="Logo" className="w-20 h-20 object-contain drop-shadow-md" />}
             <div>
               <h2 className="text-3xl font-black text-slate-900 tracking-tight uppercase">{company.name}</h2>
               <p className="text-slate-600 text-sm mt-2 max-w-xs leading-relaxed">{company.address}</p>
               <p className="text-slate-600 text-xs mt-1 font-medium">{company.phone} • {company.email}</p>
             </div>
          </div>
          <div className="text-right">
            <h1 className="text-4xl font-light text-slate-300 uppercase tracking-widest">INVOICE</h1>
            <p className="text-slate-500 text-sm mt-4 font-medium">Tanggal Cetak: <span className="font-bold text-slate-800">{new Date().toLocaleDateString('id-ID')}</span></p>
            <p className="text-slate-500 text-xs mt-1">Ref: INV-{Date.now().toString().slice(-6)}</p>
          </div>
        </div>

        {/* Tabel Multi Transaksi (Versi Bersih) */}
        <table className="w-full text-left mb-12">
          <thead>
            <tr className="border-b-2 border-slate-900 text-slate-900">
              <th className="py-3 font-bold uppercase tracking-widest text-xs w-16 text-center">No.</th>
              <th className="py-3 font-bold uppercase tracking-widest text-xs">Uraian / Deskripsi Tagihan</th>
              <th className="py-3 font-bold uppercase tracking-widest text-xs text-right">Nominal</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {transactions.map((trx, idx) => {
               return (
                 <tr key={trx.id}>
                    <td className="py-5 text-center text-slate-500 font-mono text-sm align-top">{idx + 1}</td>
                    <td className="py-5 align-top">
                      <p className="text-base text-slate-800 font-bold">{trx.description}</p>
                      <p className="text-xs text-slate-400 font-mono mt-1">{trx.id} • {trx.date}</p>
                    </td>
                    <td className="py-5 text-right font-black text-slate-800 font-mono align-top">{formatCurrency(trx.amount)}</td>
                 </tr>
               )
            })}
          </tbody>
          <tfoot>
            <tr className="border-t-4 border-slate-900 bg-slate-50">
              <td colSpan="2" className="py-4 font-black uppercase tracking-widest text-slate-900 text-right pr-6">TOTAL TAGIHAN KESELURUHAN</td>
              <td className="py-4 font-black text-emerald-600 text-right text-xl font-mono px-2">{formatCurrency(totalAmount)}</td>
            </tr>
          </tfoot>
        </table>

        {/* Kolom Tanda Tangan */}
        <div className="grid grid-cols-2 gap-8 text-center text-sm pt-12 mt-20">
            <div>
                <p className="mb-24 text-slate-500 font-medium">Diterima / Disetujui Oleh Klien,</p>
                <div className="border-b border-slate-400 w-56 mx-auto"></div>
                <p className="text-xs text-slate-400 mt-2">Nama Terang & Cap</p>
            </div>
            <div>
                <p className="mb-24 text-slate-500 font-medium">Hormat Kami, Departemen Finance</p>
                <div className="border-b border-slate-900 w-56 mx-auto font-black pt-1 uppercase tracking-widest">{company.name}</div>
            </div>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   3. BUKU KAS BESAR
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
    exportGridToExcel('Buku Kas', cols, filtered, `Buku_Kas_${company.name}_${startDate||'All'}.xls`, company.name);
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
   4. JURNAL UMUM 
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
    exportGridToExcel('Jurnal Umum', cols, rows, `Jurnal_Umum_${company.name}_${startDate||'All'}.xls`, company.name);
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
                    <td className="px-6 py-4 text-slate-400 font-medium align-top" rowSpan="2">{t.date}<br/><span className="text-[10px] font-mono text-slate-600 mt-1">{t.id}</span></td>
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
   5. LAPORAN LABA RUGI
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
    exportFinancialToExcel('Laba Rugi', company.name, data, `Laba_Rugi_${company.name}.xls`);
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
          <h3 className="text-xl font-bold mt-2 text-slate-700 uppercase tracking-widest underline underline-offset-8">Laporan Laba Rugi Komprehensif</h3>
          <p className="text-sm text-slate-500 font-medium mt-4">Periode Berakhir Per: {new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
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
   6. NERACA POSISI KEUANGAN 
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
    exportFinancialToExcel('Neraca', company.name, data, `Neraca_${company.name}.xls`);
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
          <h3 className="text-xl font-bold mt-2 text-slate-700 uppercase tracking-widest underline underline-offset-8">Laporan Posisi Keuangan (Neraca)</h3>
          <p className="text-sm text-slate-500 font-medium mt-4">Per Tanggal: {new Date().toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
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
   7. CALK 
   ======================================================================== */
function CalkView({ accounts, balances, company }) {
  return (
    <div className="space-y-8 animate-fade-in">
      <div className="flex justify-end print:hidden"><button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 border border-slate-700 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg"><Printer size={14}/> Cetak Dokumen RUPS</button></div>
      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-4xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[800px] font-sans border border-slate-200">
        <PrintLetterhead company={company} />
        <div className="text-center pb-6 mb-10">
          <h3 className="text-xl font-bold mt-2 text-slate-700 uppercase tracking-widest underline underline-offset-8">Catatan Atas Laporan Keuangan (CALK)</h3>
          <p className="text-sm text-slate-500 font-medium mt-4">Tahun Buku Konstitusi Aktif</p>
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

/* ========================================================================
   8. LAPORAN PAJAK
   ======================================================================== */
function LaporanPajakView({ accounts, balances, company }) {
  const totalRev = accounts.filter(a => a.type === 'revenue').reduce((s, a) => s + (balances[a.id] || 0), 0);
  const pphFinal = totalRev * 0.005; // PP 55 / 2022

  return (
    <div className="space-y-8 animate-fade-in">
      <div className="flex justify-end print:hidden"><button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 px-5 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg"><Printer size={14}/> Cetak SPT Masa</button></div>
      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-4xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[800px] font-sans border border-slate-200">
        
        <div className="flex justify-between items-start border-b-4 border-slate-900 pb-6 mb-10">
          <div className="flex items-start gap-4">
             {company.logo && <img src={company.logo} alt="Logo" className="h-16 w-16 object-contain drop-shadow-md" />}
             <div>
               <h2 className="text-xl font-black uppercase text-slate-900 tracking-widest">{company.name}</h2>
               <p className="text-sm text-slate-600 mt-2 max-w-sm">{company.address}</p>
               <p className="text-sm font-mono font-bold text-slate-900 mt-3">NPWP: {company.npwp}</p>
             </div>
          </div>
          <div className="text-right">
            <span className="text-sm font-black bg-slate-100 border border-slate-300 p-3 rounded-md uppercase tracking-widest inline-block">Lampiran SPT PPh Final</span>
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
   9. KONFIGURASI PT (Dengan Fitur Pengecilan Gambar Base64)
   ======================================================================== */
function SettingsView({ company, setCompany, accounts, saveToCloud, user }) {
  const [formData, setFormData] = useState(company);
  const [isCompressing, setIsCompressing] = useState(false);

  // Fungsi Cerdas Mengompresi Gambar agar muat di Google Sheet
  const handleLogoUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setIsCompressing(true);

    const reader = new FileReader();
    reader.onload = (event) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const MAX_WIDTH = 250; // Gambar di-resize ke lebar maksimal 250px agar ringan
        const scaleSize = MAX_WIDTH / img.width;
        canvas.width = MAX_WIDTH;
        canvas.height = img.height * scaleSize;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        
        // Kompres ke JPEG dengan kualitas 80% (mengurangi memori hingga 90%)
        const dataUrl = canvas.toDataURL('image/jpeg', 0.8);
        setFormData({...formData, logo: dataUrl});
        setIsCompressing(false);
      };
      img.src = event.target.result;
    };
    reader.readAsDataURL(file);
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
        
        {/* Fitur Upload Logo */}
        <div className="bg-slate-950 border border-slate-800 p-6 rounded-2xl flex flex-col md:flex-row items-center gap-6">
           <div className="w-24 h-24 shrink-0 rounded-xl bg-slate-900 flex items-center justify-center border-2 border-dashed border-slate-700 overflow-hidden">
              {formData.logo ? (
                <img src={formData.logo} alt="Preview Logo" className="w-full h-full object-contain" />
              ) : (
                <ImagePlus size={32} className="text-slate-600" />
              )}
           </div>
           <div className="flex-1 w-full text-center md:text-left">
              <label className="block text-xs font-bold text-slate-400 uppercase tracking-widest mb-2">Upload Logo Perusahaan</label>
              <input type="file" accept="image/png, image/jpeg" onChange={handleLogoUpload} className="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-bold file:bg-emerald-500/10 file:text-emerald-400 hover:file:bg-emerald-500/20 cursor-pointer" />
              <p className="text-[10px] text-slate-500 mt-2 italic">*Gambar akan otomatis dikompres agar muat tersimpan permanen di cloud Database.</p>
              {isCompressing && <p className="text-[10px] text-amber-400 mt-1 animate-pulse">Sedang mengompres gambar...</p>}
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
          <button type="submit" disabled={isCompressing} className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 px-8 py-4 rounded-xl font-black text-xs uppercase tracking-widest text-white shadow-lg shadow-emerald-900/50 w-full md:w-auto transition-all active:scale-95 disabled:opacity-50 disabled:cursor-not-allowed">Simpan Perubahan Sistem</button>
        </div>
      </form>
    </div>
  );
}

/* ========================================================================
   10. KOMPONEN PEMBANTU KOP SURAT
   ======================================================================== */
function PrintLetterhead({ company }) {
  return (
    <div className="border-b-[3px] border-slate-900 pb-6 mb-8 flex justify-between items-center no-print print:flex">
      <div className="flex gap-5 items-center">
         {company.logo && <img src={company.logo} alt="Logo" className="w-20 h-20 object-contain" />}
         <div>
            <h1 className="font-black text-2xl uppercase tracking-widest text-slate-900">{company.name}</h1>
            <p className="text-xs text-slate-600 font-medium max-w-sm mt-1">{company.address}</p>
         </div>
      </div>
      <div className="text-right text-xs font-medium text-slate-600 space-y-1">
        <p>Telp: {company.phone}</p>
        <p>Email: {company.email}</p>
        <p>NPWP: {company.npwp}</p>
      </div>
    </div>
  );
}

/* ========================================================================
   11. PRESENTASI RUPS
   ======================================================================== */
function PresentasiRupsView() {
  return <div className="text-center p-20 text-slate-500">Fitur Presentasi sedang dinonaktifkan.</div>;
}

/* ========================================================================
   12. BERITA ACARA RUPS
   ======================================================================== */
function BeritaAcaraView({ company }) {
  const [formData, setFormData] = useState({
    tanggal: new Date().toISOString().split('T')[0],
    pimpinan: 'Direktur Utama',
    sekretaris: 'Sekretaris Perusahaan',
    keputusan: '1. Mengesahkan Laporan Keuangan Tahunan.\n2. Menetapkan pelunasan dan pembebasan tanggung jawab (acquit et decharge) kepada Direksi.\n3. Menetapkan laba bersih perusahaan untuk ditahan sebagai modal kerja.'
  });

  return (
    <div className="space-y-6 animate-fade-in relative z-10">
      <div className="flex justify-between items-center print:hidden">
        <div>
          <h2 className="text-2xl font-black text-white uppercase tracking-wider">Berita Acara RUPS</h2>
          <p className="text-slate-400 text-xs font-medium mt-1">Isi formulir ini untuk membuat dokumen Berita Acara resmi secara otomatis.</p>
        </div>
        <button onClick={() => window.print()} className="bg-slate-800 hover:bg-slate-700 border border-slate-700 px-6 py-3 rounded-xl text-xs font-black uppercase tracking-widest text-white flex items-center gap-2 shadow-lg transition-all"><Printer size={16}/> Cetak Dokumen Legal</button>
      </div>
      
      {/* Form Input Rahasia */}
      <div className="bg-slate-900 border border-slate-800 p-6 rounded-2xl grid grid-cols-1 md:grid-cols-2 gap-4 print:hidden">
        <div><label className="block text-[10px] text-slate-400 font-bold uppercase mb-2">Tanggal Rapat</label><input type="date" className="w-full bg-slate-950 p-3 rounded-xl border border-slate-700 text-white text-sm" value={formData.tanggal} onChange={e => setFormData({...formData, tanggal: e.target.value})} /></div>
        <div><label className="block text-[10px] text-slate-400 font-bold uppercase mb-2">Nama Pimpinan</label><input type="text" className="w-full bg-slate-950 p-3 rounded-xl border border-slate-700 text-white text-sm" value={formData.pimpinan} onChange={e => setFormData({...formData, pimpinan: e.target.value})} /></div>
        <div className="md:col-span-2"><label className="block text-[10px] text-slate-400 font-bold uppercase mb-2">Hasil Keputusan Rapat (Enter untuk list baru)</label><textarea className="w-full bg-slate-950 p-3 rounded-xl border border-slate-700 text-white text-sm h-32" value={formData.keputusan} onChange={e => setFormData({...formData, keputusan: e.target.value})} /></div>
      </div>

      {/* Dokumen Kertas A4 yang akan Dicetak */}
      <div className="bg-white text-slate-900 p-12 md:p-16 rounded-3xl max-w-4xl mx-auto shadow-2xl print:shadow-none print:p-0 min-h-[1000px] border border-slate-200 text-justify">
        <PrintLetterhead company={company} />
        <h3 className="text-center font-black text-xl uppercase tracking-widest underline underline-offset-4 mb-10">BERITA ACARA RAPAT UMUM PEMEGANG SAHAM</h3>
        
        <div className="space-y-6 text-base leading-relaxed">
           <p>Pada hari ini, tanggal <strong>{new Date(formData.tanggal).toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</strong>, bertempat di <strong>{company.address}</strong>, telah diselenggarakan Rapat Umum Pemegang Saham (selanjutnya disebut "Rapat") dari Perusahaan <strong>{company.name}</strong>.</p>
           <p>Rapat ini dihadiri oleh para pemegang saham yang memenuhi kuorum dan dipimpin oleh <strong>{formData.pimpinan}</strong>.</p>
           
           <div className="pt-6">
             <p className="font-bold uppercase tracking-widest mb-4 border-b-2 border-slate-900 inline-block">Hasil Keputusan Rapat:</p>
             <p className="mb-4">Setelah menelaah Laporan Keuangan (Neraca dan Laba/Rugi), Rapat secara musyawarah dan mufakat memutuskan hal-hal sebagai berikut:</p>
             <div className="ml-8 space-y-3 font-medium text-slate-800">
               {formData.keputusan.split('\n').map((item, index) => (
                 <p key={index} className="pl-4 -indent-4">{item}</p>
               ))}
             </div>
           </div>

           <p className="pt-8">Demikian Berita Acara Rapat ini dibuat dengan sebenar-benarnya untuk dapat dipergunakan sebagaimana mestinya.</p>
        </div>

        <div className="mt-20 flex justify-between text-center">
          <div className="w-64">
            <p className="mb-24 text-slate-600">Pimpinan Rapat,</p>
            <p className="font-bold uppercase border-b border-slate-900 pb-1">{formData.pimpinan}</p>
          </div>
          <div className="w-64">
            <p className="mb-24 text-slate-600">Sekretaris Rapat,</p>
            <p className="font-bold uppercase border-b border-slate-900 pb-1">{formData.sekretaris}</p>
          </div>
        </div>
      </div>
    </div>
  );
}
