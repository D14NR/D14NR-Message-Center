import React, { useState, useEffect, useCallback, useRef } from 'react';
import axios from 'axios';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { 
  Send, Settings, FileSpreadsheet, CheckCircle2, XCircle, Loader2, Database, RefreshCw, 
  User, Upload, Trash2, X, LayoutDashboard, History as HistoryIcon, 
  MessageSquare, Type, Menu, ShieldCheck, Zap, Globe, Lock, LogOut, 
  ArrowRight, Sun, Moon, Paperclip, ExternalLink, AlertTriangle, Clock, ChevronDown, Download, MoreVertical
} from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Multi-Language Dictionary ---
const translations = {
  id: {
    dashboard: 'Dashboard', broadcast: 'Broadcast', configuration: 'Konfigurasi', history: 'Riwayat',
    status: 'STATUS', total_data: 'TOTAL DATA', pending: 'TERTUNDA', success: 'SUKSES', failed: 'GAGAL',
    ready_broadcast: 'Siap Kirim Pesan?', start_now: 'MULAI SEKARANG', system_config: 'Pengaturan Sistem',
    api_key: 'Fonnte API Token Key', delay: 'Jeda Pesan', auto_check: 'Cek WA Otomatis',
    theme: 'Tampilan Tema', language: 'Pilihan Bahasa', connect: 'HUBUNGKAN', connected: 'TERHUBUNG',
    not_connected: 'BELUM TERHUBUNG', disconnect: 'PUTUSKAN', light: 'TERANG', dark: 'GELAP',
    title: 'Judul Broadcast', excel_upload: 'Unggah Excel', manual_add: 'Tambah Manual',
    queue: 'Antrean', check_wa: 'Cek WA', message: 'Isi Pesan', apply: 'TERAPKAN', cancel: 'BATAL',
    attachment: 'Lampiran', file_upload: 'Upload File', subscribe: 'Berlangganan Paket',
    package_warning: 'Pastikan Anda sudah memiliki paket yang mendukung fitur kirim lampiran.',
    delete_data: 'Hapus Data', export: 'Export', delete_confirm: 'Hapus semua data histori?'
  },
  en: {
    dashboard: 'Dashboard', broadcast: 'Broadcast', configuration: 'Configuration', history: 'History',
    status: 'STATUS', total_data: 'TOTAL DATA', pending: 'PENDING', success: 'SUCCESS', failed: 'FAILED',
    ready_broadcast: 'Ready to Broadcast?', start_now: 'START NOW', system_config: 'System Configuration',
    api_key: 'Fonnte API Token Key', delay: 'Message Delay', auto_check: 'Auto-Check WA',
    theme: 'Display Theme', language: 'Language', connect: 'CONNECT', connected: 'CONNECTED',
    not_connected: 'NOT CONNECTED', disconnect: 'DISCONNECT', light: 'LIGHT', dark: 'DARK',
    title: 'Broadcast Title', excel_upload: 'Excel Upload', manual_add: 'Manual Add',
    queue: 'Queue', check_wa: 'Check WA', message: 'Message', apply: 'APPLY', cancel: 'CANCEL',
    attachment: 'Attachment', file_upload: 'Upload File', subscribe: 'Subscribe Package',
    package_warning: 'Make sure you have a package that supports attachment sending feature.',
    delete_data: 'Delete Data', export: 'Export', delete_confirm: 'Delete all history data?'
  }
};

interface SpreadsheetRow {
  'Nomor Telepon': string; 'Nama': string; 'Pesan': string; 'Status': string; 'Lampiran'?: string; __id?: number;
}

interface HistoryRow {
  'Device': string; 'nomor device': string; 'Judul broadcast': string; 'nomor Tujuan': string; 'tanggal pengiriman': string; 'status': string;
}

type View = 'dashboard' | 'broadcast' | 'configuration' | 'history';

const HISTORY_SHEET_URL = 'https://docs.google.com/spreadsheets/d/17YfjibgKzXX6S7UJEC0OGoDFNwbvqnn5R1rpQOtDhDI/export?format=csv&gid=685286577';
const LOG_WEBHOOK = 'https://script.google.com/macros/s/AKfycbxGvYq05Cr01naPccLtEw9oJgR0-Gwkm7hrWP8ebl9ltmvZWBtLAxwtWGW0vkTRfWo8/exec';

function App() {
  const [isLoggedIn, setIsLoggedIn] = useState(() => localStorage.getItem('isLoggedIn') === 'true');
  const [activeView, setActiveView] = useState<View>('dashboard');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  
  // Config States
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('fonnte_api_key') || '');
  const [isTokenChanged, setIsTokenChanged] = useState(false);
  const [lang, setLang] = useState<'id' | 'en'>(() => (localStorage.getItem('app_lang') as 'id' | 'en') || 'id');
  const [theme, setTheme] = useState<'light' | 'dark'>(() => (localStorage.getItem('app_theme') as 'light' | 'dark') || 'light');
  
  const [data, setData] = useState<SpreadsheetRow[]>([]);
  const [isBroadcasting, setIsBroadcasting] = useState(false);
  const [currentIndex, setCurrentIndex] = useState(-1);
  const [logs, setLogs] = useState<any[]>([]);
  const [delayRange, setDelayRange] = useState('10-25');
  const [broadcastTitle, setBroadcastTitle] = useState('');
  const isBroadcastingRef = useRef(false);

  // States
  const [manualPhone, setManualPhone] = useState('');
  const [manualName, setManualName] = useState('');
  const [manualMessage, setManualMessage] = useState('');
  const [autoCheck, setAutoCheck] = useState(true);
  const [deviceInfo, setDeviceInfo] = useState<any>(null);
  const [isCheckingConnection, setIsCheckingConnection] = useState(false);
  const [isCheckingWA, setIsCheckingWA] = useState(false);

  // History States
  const [historyData, setHistoryData] = useState<HistoryRow[]>([]);
  const [isFetchingHistory, setIsFetchingHistory] = useState(false);
  const [isHistoryMenuOpen, setIsHistoryMenuOpen] = useState(false);

  // Modals
  const [isMessageModalOpen, setIsMessageModalOpen] = useState(false);
  const [isAttachmentModalOpen, setIsAttachmentModalOpen] = useState(false);
  const [bulkMessageInput, setBulkMessageInput] = useState('');
  const [attachmentUrl, setAttachmentUrl] = useState('');
  const [isSelectBroadcastModalOpen, setIsSelectBroadcastModalOpen] = useState(false);
  const [selectBroadcastAction, setSelectBroadcastAction] = useState<'export' | 'delete' | null>(null);
  const [selectedBroadcastForAction, setSelectedBroadcastForAction] = useState<string | null>(null);

  const t = translations[lang];

  useEffect(() => {
    localStorage.setItem('fonnte_api_key', apiKey);
    localStorage.setItem('app_lang', lang);
    localStorage.setItem('app_theme', theme);
    if (theme === 'dark') document.documentElement.classList.add('dark');
    else document.documentElement.classList.remove('dark');
  }, [apiKey, lang, theme]);

  const addLog = useCallback((target: string, status: 'success' | 'error' | 'info', message: string) => {
    setLogs(prev => [{ timestamp: new Date().toLocaleTimeString(), target, status, message }, ...prev].slice(0, 100));
  }, []);

  const formatPhone = (phone: string) => {
    let cleaned = String(phone).replace(/[^0-9]/g, '');
    if (cleaned.startsWith('0')) cleaned = '62' + cleaned.substring(1);
    else if (cleaned.startsWith('8')) cleaned = '62' + cleaned;
    return cleaned;
  };

  const checkConnection = async (silent = false) => {
    if (!apiKey) return;
    if (!silent) setIsCheckingConnection(true);
    try {
      const response = await axios.post('https://api.fonnte.com/device', {}, { headers: { 'Authorization': apiKey } });
      if (response.data.status) {
        setDeviceInfo(response.data);
        setIsTokenChanged(false);
      } else if (!silent) {
        setDeviceInfo(null);
      }
    } catch (error) {
      if (!silent) setDeviceInfo(null);
    } finally {
      if (!silent) setIsCheckingConnection(false);
    }
  };

  const fetchHistory = async () => {
    setIsFetchingHistory(true);
    try {
      const response = await axios.get(HISTORY_SHEET_URL);
      const results = Papa.parse<HistoryRow>(response.data, { header: true, skipEmptyLines: true });
      setHistoryData(results.data);
    } catch (error) { console.error(error); } finally { setIsFetchingHistory(false); }
  };

  const exportHistoryData = () => {
    if (historyRows.length === 0) {
      alert(lang === 'id' ? 'Tidak ada data untuk di export' : 'No data to export');
      return;
    }
    setSelectBroadcastAction('export');
    setIsSelectBroadcastModalOpen(true);
  };

  const deleteHistoryData = () => {
    if (historyRows.length === 0) {
      alert(lang === 'id' ? 'Tidak ada data untuk dihapus' : 'No data to delete');
      return;
    }
    setSelectBroadcastAction('delete');
    setIsSelectBroadcastModalOpen(true);
  };

  const executeExport = (broadcastTitle: string | null) => {
    const dataToExport = broadcastTitle 
      ? historyRows.filter(r => r['Judul broadcast'] === broadcastTitle)
      : historyRows;
    
    if (dataToExport.length === 0) {
      alert(lang === 'id' ? 'Tidak ada data untuk di export' : 'No data to export');
      return;
    }
    
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "History");
    const filename = broadcastTitle 
      ? `history_${broadcastTitle}_${new Date().toISOString().split('T')[0]}.xlsx`
      : `history_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(wb, filename);
    addLog('System', 'success', lang === 'id' ? 'Data berhasil di export' : 'Data exported successfully');
  };

  const executeDelete = (broadcastTitle: string | null) => {
    const confirmMessage = lang === 'id' 
      ? `Hapus data broadcast "${broadcastTitle || 'Semua'}"? Tindakan ini tidak dapat dibatalkan.`
      : `Delete "${broadcastTitle || 'All'}" broadcast data? This action cannot be undone.`;
    
    if (!confirm(confirmMessage)) return;
    
    if (broadcastTitle) {
      setHistoryData(prev => prev.filter(r => r['Judul broadcast'] !== broadcastTitle));
    } else {
      setHistoryData([]);
    }
    
    setIsHistoryMenuOpen(false);
    addLog('System', 'info', lang === 'id' ? 'Data histori dihapus' : 'History data deleted');
  };

  useEffect(() => { if (activeView === 'history' && isLoggedIn) fetchHistory(); }, [activeView, isLoggedIn]);

  const currentDeviceName = deviceInfo?.name?.trim().toLowerCase();
  const historyRows = currentDeviceName
    ? historyData.filter(r => r.Device?.trim().toLowerCase() === currentDeviceName)
    : historyData;

  const logToSheet = async (target: string, status: string) => {
    try {
      const payload = {
        action: 'UPSERT',
        device: deviceInfo?.name || 'Unknown',
        nomor_device: deviceInfo?.device || '-',
        judul: broadcastTitle,
        tujuan: target,
        tanggal: new Date().toLocaleDateString('id-ID'),
        status: status
      };
      await axios.post(LOG_WEBHOOK, JSON.stringify(payload));
    } catch (e) { console.error(e); }
  };

  const bulkCheckWhatsApp = async (rowsToCheck: SpreadsheetRow[]) => {
    if (!apiKey || rowsToCheck.length === 0) return;
    setIsCheckingWA(true);
    try {
      const targets = rowsToCheck.map(r => formatPhone(r['Nomor Telepon'])).join(',');
      const params = new URLSearchParams(); params.append('target', targets);
      const res = await axios.post('https://api.fonnte.com/validate', params, { headers: { 'Authorization': apiKey } });
      if (res.data.status) {
        const reg: string[] = res.data.registered || [];
        const notReg: string[] = res.data.not_registered || [];
        setData(prev => prev.map(row => {
          const f = formatPhone(row['Nomor Telepon']);
          if (reg.includes(f)) return { ...row, Status: 'WA Active' };
          if (notReg.includes(f)) return { ...row, Status: 'No WA' };
          return row;
        }));
      }
    } catch (error) { console.error(error); } finally { setIsCheckingWA(false); }
  };

  const sendMessage = async (row: SpreadsheetRow): Promise<{ success: boolean; message: string }> => {
    const target = formatPhone(row['Nomor Telepon']);
    const message = row['Pesan'].replace(/{nama}/gi, row['Nama']);
    try {
      const params = new URLSearchParams();
      params.append('target', target);
      params.append('message', message);
      if (row.Lampiran) params.append('url', row.Lampiran);
      params.append('delay', '2');
      params.append('countryCode', '62');
      const response = await axios.post('https://api.fonnte.com/send', params, { headers: { 'Authorization': apiKey, 'Content-Type': 'application/x-www-form-urlencoded' } });
      return { success: response.data.status === true || response.data.status === 'true', message: response.data.detail || response.data.reason || 'Processed' };
    } catch (error: any) { return { success: false, message: error.message }; }
  };

  const startBroadcast = async () => {
    if (isBroadcasting) { setIsBroadcasting(false); isBroadcastingRef.current = false; return; }
    if (!broadcastTitle) { alert('WAJIB: Judul Broadcast!'); return; }
    const validStatuses = ['Pending', 'WA Active', ''];
    if (!data.some(row => validStatuses.includes(row.Status))) { alert('No pending.'); return; }
    setIsBroadcasting(true); isBroadcastingRef.current = true;
    for (let i = 0; i < data.length; i++) {
      if (!isBroadcastingRef.current) break;
      const row = data[i]; if (!validStatuses.includes(row.Status)) continue;
      setCurrentIndex(i);
      const res = await sendMessage(row);
      const st = res.success ? 'Success' : 'Failed';
      setData(prev => prev.map((r, idx) => idx === i ? { ...r, Status: st } : r));
      addLog(row['Nomor Telepon'], res.success ? 'success' : 'error', res.message);
      logToSheet(row['Nomor Telepon'], st);
      if (res.success) checkConnection(true);
      if (i < data.length - 1 && isBroadcastingRef.current) {
        const [mn, mx] = delayRange.split('-').map(Number);
        const rd = Math.floor(Math.random() * (mx - mn + 1) + mn);
        addLog('System', 'info', `Waiting ${rd}s for next message...`);
        await new Promise(r => setTimeout(r, rd * 1000));
      }
    }
    setIsBroadcasting(false); isBroadcastingRef.current = false; setCurrentIndex(-1);
  };

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!broadcastTitle) { alert('Isi Judul!'); e.target.value = ''; return; }
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const bstr = ev.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const results = XLSX.utils.sheet_to_json<any>(wb.Sheets[wb.SheetNames[0]]);
      const newRows: SpreadsheetRow[] = results.map((row, index) => ({
        'Nomor Telepon': String(row['Nomor Telepon'] || row['phone'] || ''),
        'Nama': String(row['Nama'] || row['name'] || 'Recipient'),
        'Pesan': '', 'Status': 'Pending', __id: Date.now() + index
      })).filter(r => r['Nomor Telepon'] !== '');
      setData(prev => [...prev, ...newRows]);
      if (autoCheck) bulkCheckWhatsApp(newRows);
      e.target.value = '';
    };
    reader.readAsBinaryString(file);
  };

  const SidebarItem = ({ id, label, icon: Icon }: { id: View, label: string, icon: any }) => {
    const handleClick = () => {
      if (id === 'broadcast' && !deviceInfo) { alert('Silakan hubungkan token Fonnte!'); setActiveView('configuration'); setIsMobileMenuOpen(false); return; }
      setActiveView(id); setIsMobileMenuOpen(false);
    };
    return (
      <button onClick={handleClick} className={cn("w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all active:scale-95", activeView === id ? "bg-emerald-600 text-white shadow-lg" : "text-slate-500 hover:bg-slate-100 dark:text-slate-400 dark:hover:bg-slate-800")}>
        <Icon className="w-5 h-5" />
        <span className="font-semibold text-sm">{label}</span>
      </button>
    );
  };

  const stats = {
    total: data.length,
    pending: data.filter(r => ['Pending', 'WA Active', ''].includes(r.Status)).length,
    success: data.filter(r => r.Status === 'Success').length,
    failed: data.filter(r => r.Status === 'Failed').length,
  };

  const hasPaidPackage = deviceInfo?.package && ['super', 'advanced', 'ultra'].includes(deviceInfo.package.toLowerCase());

  if (!isLoggedIn) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-6 relative overflow-hidden">
        <div className="absolute top-0 -left-20 w-96 h-96 bg-emerald-500/20 rounded-full blur-[120px]" />
        <div className="absolute bottom-0 -right-20 w-96 h-96 bg-blue-500/20 rounded-full blur-[120px]" />
        <div className="w-full max-w-md animate-in fade-in zoom-in duration-500">
          <div className="bg-white/10 backdrop-blur-xl border border-white/10 rounded-[40px] shadow-2xl p-10 space-y-10 text-center">
            <div className="space-y-4">
              <div className="inline-flex bg-gradient-to-br from-emerald-400 to-emerald-600 p-2 rounded-[28px] text-white shadow-xl overflow-hidden"><img src="/logo.png" alt="D14NR Logo" className="w-10 h-10 object-cover" /></div>
              <div><h1 className="text-3xl font-black text-white tracking-tighter uppercase leading-none">D14NR</h1><p className="text-[12px] font-black text-emerald-400 uppercase tracking-[0.3em] mt-2 leading-none">Message Center</p></div>
            </div>
            <form onSubmit={(e) => { e.preventDefault(); const fd = new FormData(e.currentTarget); if (fd.get('user') === 'D14nr' && fd.get('pass') === '290192') { setIsLoggedIn(true); localStorage.setItem('isLoggedIn', 'true'); } else alert('Invalid Credentials'); }} className="space-y-6">
              <div className="space-y-3">
                <div className="relative"><User className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-500 w-5 h-5" /><input name="user" type="text" placeholder="Username" className="w-full pl-14 pr-6 py-5 bg-white/5 border border-white/10 rounded-3xl outline-none text-white font-bold" /></div>
                <div className="relative"><Lock className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-500 w-5 h-5" /><input name="pass" type="password" placeholder="Password" className="w-full pl-14 pr-6 py-5 bg-white/5 border border-white/10 rounded-3xl outline-none text-white font-bold" /></div>
              </div>
              <button type="submit" className="w-full py-5 bg-emerald-500 text-white rounded-3xl font-black text-sm tracking-widest uppercase shadow-2xl hover:bg-emerald-400 active:scale-95 transition-all">SIGN IN <ArrowRight className="ml-2 w-5 h-5 inline" /></button>
            </form>
            <p className="text-center text-[10px] font-black text-slate-600 uppercase tracking-widest leading-none">D14nr System © 2026</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={cn("flex h-screen font-sans overflow-hidden transition-all duration-500", theme === 'dark' ? "bg-slate-950 text-slate-100" : "bg-slate-50 text-slate-900")}>
      {isMobileMenuOpen && <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-40 lg:hidden" onClick={() => setIsMobileMenuOpen(false)} />}
      <aside className={cn("fixed lg:relative inset-y-0 left-0 w-64 border-r flex flex-col z-50 transition-all duration-300 transform lg:translate-x-0 shadow-sm", isMobileMenuOpen ? "translate-x-0" : "-translate-x-full", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
        <div className="p-6 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-3"><div className="rounded-xl overflow-hidden shadow-lg"><img src="/logo.png" alt="D14NR Logo" className="w-10 h-10 object-cover" /></div><div className="flex flex-col"><h1 className="text-sm font-black tracking-tighter leading-none uppercase">D14NR</h1><p className="text-[10px] font-bold text-emerald-600 uppercase tracking-widest leading-none mt-1">Message Center</p></div></div>
          <button className="lg:hidden text-slate-400" onClick={() => setIsMobileMenuOpen(false)}><X className="w-6 h-6" /></button>
        </div>
        <nav className="flex-1 p-4 space-y-2 overflow-y-auto">
          <SidebarItem id="dashboard" label={t.dashboard} icon={LayoutDashboard} />
          <SidebarItem id="broadcast" label={t.broadcast} icon={MessageSquare} />
          <SidebarItem id="history" label={t.history} icon={HistoryIcon} />
          <div className="pt-4 mt-4 border-t border-slate-100 dark:border-slate-800 text-[10px] font-black text-slate-300 dark:text-slate-600 uppercase tracking-widest px-4">System</div>
          <SidebarItem id="configuration" label={t.configuration} icon={Settings} />
        </nav>
        <div className="p-4 border-t border-slate-100 dark:border-slate-800 space-y-4 text-center">
          <div className="bg-slate-50 dark:bg-slate-800/50 p-3 rounded-2xl border border-slate-200 dark:border-slate-800 flex items-center gap-3 text-left">
            <div className={cn("w-2.5 h-2.5 rounded-full", deviceInfo?.device_status === 'connect' ? "bg-emerald-500 animate-pulse shadow-[0_0_8px_rgba(16,185,129,0.5)]" : "bg-slate-300 dark:bg-slate-600")} />
            <div className="flex flex-col leading-tight"><span className="text-[10px] font-bold text-slate-400 uppercase">{t.status}</span><span className="text-[11px] font-black truncate uppercase tracking-tighter">{deviceInfo?.name || 'OFFLINE'}</span></div>
          </div>
          <p className="text-[9px] font-black text-slate-300 dark:text-slate-600 uppercase tracking-widest leading-none">Copyright D14nr@2026</p>
        </div>
      </aside>

      <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
        <header className={cn("h-16 lg:h-14 border-b flex items-center justify-between px-4 lg:px-8 z-30 shrink-0 shadow-sm", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
          <div className="flex items-center gap-3"><button className="lg:hidden p-2 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-lg" onClick={() => setIsMobileMenuOpen(true)}><Menu className="w-6 h-6" /></button><h2 className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{activeView}</h2></div>
          <div className="flex items-center gap-4">
            {isBroadcasting && <div className="hidden sm:flex items-center gap-2 px-3 py-1 bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400 rounded-full animate-pulse text-[10px] font-black uppercase">Broadcasting</div>}
            <button onClick={() => { setIsLoggedIn(false); localStorage.removeItem('isLoggedIn'); }} className="flex items-center gap-2 px-3 py-1.5 bg-slate-100 dark:bg-slate-800 hover:bg-rose-50 dark:hover:bg-rose-900/30 hover:text-rose-600 text-slate-500 rounded-xl transition-all font-black text-[10px] uppercase shadow-sm"><LogOut className="w-3.5 h-3.5" /> LOGOUT</button>
          </div>
        </header>

        <main className="flex-1 overflow-y-auto p-4 lg:p-8 bg-transparent">
          <div className="max-w-6xl mx-auto space-y-6">
            {activeView === 'dashboard' && (
              <div className="space-y-8 animate-in fade-in slide-in-from-bottom-2 duration-500 leading-none">
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 lg:gap-6">
                  {[ { l: t.total_data, v: stats.total, i: Database, c: 'text-slate-600' }, { l: t.pending, v: stats.pending, i: Loader2, c: 'text-amber-600' }, { l: t.success, v: stats.success, i: CheckCircle2, c: 'text-emerald-600' }, { l: t.failed, v: stats.failed, i: XCircle, c: 'text-rose-600' } ].map((s, i) => (
                    <div key={i} className={cn("p-5 lg:p-6 rounded-[32px] border shadow-sm flex items-center gap-4 transition-all", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200 shadow-slate-200/20")}>
                      <div className={cn("p-4 rounded-[20px] bg-slate-50 dark:bg-slate-800", s.c)}><s.i className="w-6 h-6" /></div>
                      <div><p className="text-[10px] font-black text-slate-400 mb-1">{s.l}</p><p className={cn("text-2xl lg:text-3xl font-black", s.c)}>{s.v}</p></div>
                    </div>
                  ))}
                </div>
                <div className="bg-gradient-to-br from-emerald-600 to-emerald-800 p-8 lg:p-10 rounded-[48px] text-white shadow-2xl relative overflow-hidden group">
                  <div className="absolute -right-10 -bottom-10 opacity-10 group-hover:scale-110 transition-transform duration-700 font-black italic text-9xl">D14NR</div>
                  <div className="relative z-10 space-y-8">
                    <div className="space-y-4 leading-none"><h3 className="text-3xl lg:text-4xl font-black uppercase tracking-tight">Empowering Communication</h3><p className="text-emerald-100/80 font-medium max-w-lg text-lg leading-relaxed mt-4">D14NR Message Center bekerjasama dengan Fonnte menghadirkan solusi pengiriman pesan WhatsApp yang aman dan efisien.</p></div>
                    <div className="flex flex-wrap gap-4 leading-none"><div className="bg-white/10 backdrop-blur-md px-5 py-2.5 rounded-[20px] flex items-center gap-2 border border-white/5"><ShieldCheck className="w-4 h-4" /><span className="text-[10px] font-black uppercase tracking-widest">Anti-Spam Tech</span></div><div className="bg-white/10 backdrop-blur-md px-5 py-2.5 rounded-[20px] flex items-center gap-2 border border-white/5"><Zap className="w-4 h-4" /><span className="text-[10px] font-black uppercase tracking-widest">Instant Delivery</span></div><div className="bg-white/10 backdrop-blur-md px-5 py-2.5 rounded-[20px] flex items-center gap-2 border border-white/5"><Globe className="w-4 h-4" /><span className="text-[10px] font-black uppercase tracking-widest">Global Reach</span></div></div>
                    <button onClick={() => setActiveView('broadcast')} className="mt-4 px-12 py-5 bg-white text-emerald-700 rounded-[24px] font-black shadow-2xl hover:scale-105 transition-all uppercase tracking-widest text-sm leading-none">{t.start_now}</button>
                  </div>
                </div>
              </div>
            )}

            {activeView === 'broadcast' && (
              <div className="space-y-6 animate-in fade-in slide-in-from-bottom-2 duration-500">
                <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 lg:gap-4 leading-none">
                  {[ { l: t.total_data, v: stats.total, i: Database, c: 'text-slate-600' }, { l: t.pending, v: stats.pending, i: Loader2, c: 'text-amber-600' }, { l: t.success, v: stats.success, i: CheckCircle2, c: 'text-emerald-600' }, { l: t.failed, v: stats.failed, i: XCircle, c: 'text-rose-600' } ].map((s, i) => (
                    <div key={i} className={cn("p-3 lg:p-4 rounded-[24px] border shadow-sm flex items-center gap-3 transition-all", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200 shadow-slate-200/20")}>
                      <div className={cn("p-2 rounded-xl bg-slate-50 dark:bg-slate-800", s.c)}><s.i className={cn("w-4 lg:w-5 h-4 lg:h-5", s.l === t.pending && isBroadcasting && "animate-spin")} /></div>
                      <div className="min-w-0"><p className="text-[8px] lg:text-[9px] font-black text-slate-400 leading-none mb-1 truncate uppercase">{s.l}</p><p className={cn("text-base lg:text-xl font-black truncate", s.c)}>{s.v}</p></div>
                    </div>
                  ))}
                </div>
                <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 lg:gap-8 pb-20 leading-none">
                  <div className="lg:col-span-4 space-y-6 order-2 lg:order-1">
                    <div className={cn("p-5 lg:p-6 rounded-[32px] border space-y-4 shadow-sm", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
                      <div className="flex items-center gap-2 text-slate-800 dark:text-slate-100 font-black text-[11px] uppercase tracking-widest"><Type className="w-4 h-4 text-emerald-500" /> {t.title}</div>
                      <input type="text" value={broadcastTitle} onChange={(e) => setBroadcastTitle(e.target.value)} className={cn("w-full px-4 py-3 border rounded-2xl text-xs font-bold outline-none", theme === 'dark' ? "bg-slate-800 border-slate-700 text-white" : "bg-slate-50 border-slate-100 text-slate-700")} />
                    </div>
                    <div className={cn("p-5 lg:p-6 rounded-[32px] border space-y-4 shadow-sm", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
                      <div className="flex items-center justify-between font-black text-[11px] uppercase tracking-widest text-slate-800 dark:text-slate-100"><div className="flex items-center gap-2"><Upload className="w-4 h-4 text-emerald-500" /> Excel</div><button onClick={() => { const ws = XLSX.utils.json_to_sheet([{ 'Nomor Telepon': '08123...', 'Nama': 'Budi' }]); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Sheet1"); XLSX.writeFile(wb, "template.xlsx"); }} className="text-[9px] bg-slate-100 dark:bg-slate-800 px-2.5 py-1.5 rounded-lg text-slate-500 font-black border dark:border-slate-700">TEMPLATE</button></div>
                      <div className="relative border-2 border-dashed border-slate-100 dark:border-slate-800 rounded-3xl p-6 bg-slate-50/50 dark:bg-slate-800/50 text-center hover:border-emerald-500 transition-colors"><input type="file" accept=".xlsx, .xls" onChange={handleExcelUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" /><FileSpreadsheet className="w-10 h-10 text-slate-300 mx-auto mb-2" /><p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{t.excel_upload}</p></div>
                    </div>
                    <div className={cn("p-5 lg:p-6 rounded-[32px] border space-y-4 shadow-sm", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
                      <div className="flex items-center gap-2 text-slate-800 dark:text-slate-100 font-black text-[11px] uppercase tracking-widest"><User className="w-4 h-4 text-emerald-500" /> {t.manual_add}</div>
                      <form onSubmit={(e) => { e.preventDefault(); if (!broadcastTitle) { alert('Isi Judul!'); return; } const list = manualPhone.split(/[,\n]/).map(p => p.trim()).filter(p => p !== ''); const newRows = list.map((p, i) => ({ 'Nomor Telepon': p, 'Nama': manualName || 'Recipient', 'Pesan': manualMessage, 'Status': 'Pending', __id: Date.now() + i })); setData(p => [...p, ...newRows]); if (autoCheck) bulkCheckWhatsApp(newRows); setManualPhone(''); setManualName(''); setManualMessage(''); }} className="space-y-3">
                        <textarea value={manualPhone} onChange={(e) => setManualPhone(e.target.value)} placeholder="0812, 0856..." className={cn("w-full p-4 border rounded-2xl text-xs font-bold h-20 resize-none outline-none", theme === 'dark' ? "bg-slate-800 border-slate-700" : "bg-slate-50 border-slate-100")} />
                        <input type="text" value={manualName} onChange={(e) => setManualName(e.target.value)} placeholder="Name" className={cn("w-full p-4 border rounded-2xl text-xs font-bold outline-none", theme === 'dark' ? "bg-slate-800 border-slate-700" : "bg-slate-50 border-slate-100")} />
                        <button type="submit" className="w-full py-4 bg-slate-900 dark:bg-emerald-600 text-white rounded-2xl font-black text-[10px] uppercase tracking-widest shadow-xl active:scale-95 transition-all">ADD TO QUEUE</button>
                      </form>
                    </div>
                  </div>

                  <div className="lg:col-span-8 space-y-6 order-1 lg:order-2">
                    <div className={cn("rounded-[40px] border shadow-sm flex flex-col h-[550px] overflow-hidden transition-all", theme === 'dark' ? "bg-slate-900 border-slate-800 shadow-emerald-950/20" : "bg-white border-slate-200 shadow-slate-200/50")}>
                      <div className="p-4 lg:p-6 bg-slate-50/50 dark:bg-slate-800/50 border-b dark:border-slate-800 flex flex-wrap gap-2 lg:gap-3 justify-between items-center px-6 lg:px-8 shrink-0">
                        <div className="flex gap-2">
                          <button onClick={() => setIsMessageModalOpen(true)} className="px-3 lg:px-4 py-1.5 lg:py-2 bg-emerald-50 dark:bg-emerald-900/30 text-emerald-700 dark:text-emerald-400 text-[10px] font-black uppercase rounded-xl border border-emerald-100 dark:border-emerald-900/50 flex items-center gap-2 hover:bg-emerald-100 transition-all shadow-sm"><MessageSquare className="w-3 h-3 lg:w-3.5 lg:h-3.5" /> {t.message}</button>
                          <button onClick={() => setIsAttachmentModalOpen(true)} className="px-3 lg:px-4 py-1.5 lg:py-2 bg-sky-50 dark:bg-sky-900/30 text-sky-700 dark:text-sky-400 text-[10px] font-black uppercase rounded-xl border border-sky-100 dark:border-sky-900/50 flex items-center gap-2 hover:bg-sky-100 transition-all shadow-sm"><Paperclip className="w-3 h-3 lg:w-3.5 lg:h-3.5" /> {t.attachment}</button>
                          <div className="relative group ml-2">
                            <div className="flex items-center gap-2 bg-slate-100 dark:bg-slate-800 p-1 lg:p-1.5 rounded-xl border dark:border-slate-700 pr-4">
                              <div className="p-1 bg-white dark:bg-slate-900 rounded-lg text-emerald-600 shadow-sm"><Clock className="w-3 h-3" /></div>
                              <select value={delayRange} onChange={(e) => setDelayRange(e.target.value)} className="bg-transparent text-[10px] font-black uppercase outline-none cursor-pointer appearance-none">
                                {['10-25', '10-30', '10-35', '10-40', '10-45', '10-50', '10-55', '10-60'].map(r => <option key={r} value={r}>{r}s</option>)}
                              </select>
                              <ChevronDown className="w-3 h-3 text-slate-400 -ml-2" />
                            </div>
                          </div>
                        </div>
                        <div className="flex items-center gap-3">
                          <button onClick={() => { if(confirm('Clear?')) setData([]) }} className="text-slate-300 hover:text-rose-500 p-2"><Trash2 className="w-4 h-4 lg:w-5 lg:h-5" /></button>
                        </div>
                      </div>
                      <div className="flex-1 overflow-auto px-2 lg:px-4">
                        <table className="w-full text-left border-collapse min-w-[700px]"><thead className="bg-slate-50/50 dark:bg-slate-800 sticky top-0 text-[10px] uppercase font-black text-slate-400 border-b dark:border-slate-800 z-10 tracking-[0.2em]"><tr className="px-6"><th className="pl-6 lg:pl-10 py-5 w-10">#</th><th className="px-4 lg:px-6 py-4">Recipient</th><th className="px-4 lg:px-6 py-4">Preview</th><th className="px-4 lg:px-6 py-4 w-24 text-center">Status</th><th className="px-4 lg:px-6 py-4 w-24 text-center">Action</th></tr></thead>
                        <tbody className="divide-y divide-slate-50 dark:divide-slate-800 text-[11px] lg:text-xs font-bold">{data.map((row, idx) => (<tr key={idx} className={cn("hover:bg-slate-50/50 dark:hover:bg-slate-800/50 transition-all", currentIndex === idx && "bg-amber-50 dark:bg-amber-900/10 animate-pulse")}><td className="pl-6 lg:pl-10 py-5 text-slate-300 dark:text-slate-700 font-black">{idx+1}</td><td className="px-4 lg:px-6 py-4"><div>{row.Nama}</div><div className="text-[10px] text-slate-400 mt-1 font-mono tracking-tighter">{row['Nomor Telepon']}</div></td><td className="px-4 lg:px-6 py-4 text-slate-500 dark:text-slate-400 italic truncate max-w-[200px] leading-relaxed flex flex-col gap-1">{row.Lampiran && <div className="inline-flex items-center gap-1 text-[8px] bg-sky-50 dark:bg-sky-900/20 text-sky-600 px-1.5 py-0.5 rounded uppercase font-black shrink-0 w-fit leading-none"><Paperclip className="w-2.5 h-2.5" /> Media</div>}{row.Pesan.replace(/{nama}/gi, row.Nama) || '-'}</td><td className="px-4 lg:px-6 py-4 text-center"><span className={cn("px-3 py-2 rounded-xl text-[9px] font-black uppercase inline-flex items-center gap-1.5 shadow-inner transition-colors", row.Status === 'Success' ? "bg-emerald-100 text-emerald-700" : row.Status === 'Failed' ? "bg-rose-100 text-rose-700" : "bg-slate-100 dark:bg-slate-800 text-slate-400 shadow-none")}>{currentIndex === idx ? <Loader2 className="w-3 h-3 animate-spin" /> : null}{currentIndex === idx ? 'Sending' : (row.Status || 'Pending')}</span></td><td className="px-4 lg:px-6 py-4 text-center"><button onClick={() => setData(p => p.filter(r => r.__id !== row.__id))} className="text-rose-400 p-2 hover:bg-rose-50 dark:hover:bg-rose-900/30 rounded-xl transition-all"><Trash2 className="w-4 h-4" /></button></td></tr>))}</tbody></table>
                      </div>
                      <div className="p-6 lg:p-10 bg-slate-50 dark:bg-slate-800/50 border-t dark:border-slate-700 shrink-0"><button onClick={startBroadcast} disabled={data.length === 0} className={cn("w-full py-5 rounded-[28px] font-black text-white shadow-2xl transition-all active:scale-95 flex items-center justify-center gap-4 tracking-[0.2em] uppercase text-sm lg:text-base", isBroadcasting ? "bg-rose-500 shadow-rose-200" : "bg-emerald-600 shadow-emerald-500/20 hover:bg-emerald-700 disabled:bg-slate-300 dark:disabled:bg-slate-900 disabled:text-slate-500 shadow-emerald-100")}>{isBroadcasting ? 'STOP ENGINE' : t.start_now}</button></div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {activeView === 'configuration' && (
              <div className="max-w-2xl mx-auto space-y-6 animate-in fade-in slide-in-from-bottom-2 duration-500 pb-20">
                <section className={cn("p-6 lg:p-12 rounded-[48px] border shadow-2xl relative overflow-hidden transition-all", theme === 'dark' ? "bg-slate-900 border-slate-800 shadow-emerald-950/20" : "bg-white border-slate-200 shadow-slate-200/50")}>
                  <div className="absolute top-0 right-0 w-48 h-48 bg-emerald-50 dark:bg-emerald-900/10 rounded-bl-[120px] -z-0 opacity-40" />
                  <div className="relative z-10 space-y-12 leading-none">
                    <div className="flex items-center gap-4"><div className="bg-slate-900 dark:bg-slate-800 p-4 rounded-3xl text-white shadow-lg"><Settings className="w-7 h-7" /></div><h3 className="text-2xl lg:text-3xl font-black uppercase tracking-tight">{t.system_config}</h3></div>
                    <div className="space-y-10">
                      <div className="space-y-4">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">{t.api_key}</label>
                        <div className="flex flex-col sm:flex-row gap-4">
                          <input type="password" value={apiKey} onChange={(e) => { setApiKey(e.target.value); setIsTokenChanged(true); }} placeholder="Paste Fonnte Key..." className={cn("flex-1 px-8 py-5 border rounded-[28px] outline-none font-bold text-sm transition-all focus:ring-4 focus:ring-emerald-500/10", theme === 'dark' ? "bg-slate-800 border-slate-700 text-white" : "bg-slate-50 border-slate-100 text-slate-700")} />
                          <div className="flex gap-2">
                            {deviceInfo ? (
                              <><div className="flex-1 sm:flex-none px-8 py-5 bg-emerald-500 text-white rounded-[28px] font-black text-[10px] tracking-widest flex items-center justify-center gap-2 shadow-lg"><CheckCircle2 className="w-4 h-4" /> {t.connected}</div><button onClick={() => { setDeviceInfo(null); setApiKey(''); setIsTokenChanged(false); }} className="px-8 py-5 bg-rose-500 text-white rounded-[28px] shadow-lg hover:bg-rose-600 transition-all font-black text-[10px] uppercase shadow-rose-100">{t.disconnect}</button></>
                            ) : (
                              <button onClick={() => checkConnection(false)} className={cn("flex-1 sm:flex-none px-10 py-5 rounded-[28px] font-black text-[10px] tracking-[0.2em] transition-all shadow-xl active:scale-95", isTokenChanged ? "bg-amber-400 text-amber-950 shadow-amber-100" : "bg-rose-500 text-white shadow-rose-500/20")}>{isCheckingConnection ? <Loader2 className="w-5 h-5 animate-spin mx-auto" /> : (isTokenChanged ? t.connect : t.not_connected)}</button>
                            )}
                          </div>
                        </div>
                      </div>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                        <div className="space-y-4"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">{t.theme}</label>
                          <div className={cn("grid grid-cols-2 p-2 rounded-[32px] border transition-all", theme === 'dark' ? "bg-slate-800 border-slate-700 shadow-inner" : "bg-slate-100 border-slate-200")}>
                            <button onClick={() => setTheme('light')} className={cn("flex items-center justify-center gap-2 py-4 rounded-[24px] text-[10px] font-black uppercase transition-all", theme === 'light' ? "bg-white text-emerald-600 shadow-xl" : "text-slate-500")}><Sun className="w-4 h-4" /> {t.light}</button>
                            <button onClick={() => setTheme('dark')} className={cn("flex items-center justify-center gap-2 py-4 rounded-[24px] text-[10px] font-black uppercase transition-all", theme === 'dark' ? "bg-slate-900 text-emerald-400 shadow-xl" : "text-slate-500")}><Moon className="w-4 h-4" /> {t.dark}</button>
                          </div>
                        </div>
                        <div className="space-y-4"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">{t.language}</label>
                          <div className={cn("grid grid-cols-2 p-2 rounded-[32px] border transition-all", theme === 'dark' ? "bg-slate-800 border-slate-700 shadow-inner" : "bg-slate-100 border-slate-200")}>
                            <button onClick={() => setLang('id')} className={cn("py-4 rounded-[24px] text-[10px] font-black transition-all uppercase tracking-widest", lang === 'id' ? "bg-emerald-600 text-white shadow-xl" : "text-slate-500")}>ID</button>
                            <button onClick={() => setLang('en')} className={cn("py-4 rounded-[24px] text-[10px] font-black transition-all uppercase tracking-widest", lang === 'en' ? "bg-emerald-600 text-white shadow-xl" : "text-slate-500")}>EN</button>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </section>
                {deviceInfo && (<div className="bg-white p-6 lg:p-10 rounded-[32px] lg:rounded-[40px] border border-slate-200 shadow-sm grid grid-cols-1 sm:grid-cols-2 gap-6 lg:gap-10 relative overflow-hidden transition-all duration-500"><div className="absolute top-0 bottom-0 left-0 w-3 bg-emerald-500" />{[{ l: 'DEVICE NAME', v: deviceInfo.name }, { l: 'STATUS', v: deviceInfo.device_status, active: deviceInfo.device_status === 'connect' }, { l: 'QUOTA', v: `${deviceInfo.quota} MSG` }, { l: 'EXPIRED', v: deviceInfo.expired }].map((d, i) => (<div key={i} className="space-y-1"><p className="text-[10px] font-black text-slate-300 uppercase tracking-widest leading-none">{d.l}</p><p className={cn("text-base lg:text-lg font-black text-slate-800 uppercase tracking-tight", d.active && "text-emerald-600")}>{d.v}</p></div>))}</div>)}
              </div>
            )}

            {activeView === 'history' && (
              <div className="space-y-6 animate-in fade-in duration-500 pb-20">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <div className="flex items-center gap-4"><div className="bg-emerald-100 dark:bg-emerald-900/30 p-4 rounded-3xl text-emerald-600 shadow-sm shrink-0"><HistoryIcon className="w-7 h-7" /></div><div><h3 className="text-2xl lg:text-3xl font-black text-slate-800 dark:text-white uppercase leading-none tracking-tight">{t.history}</h3></div></div>
                  <div className="flex items-center gap-3">
                    <button onClick={fetchHistory} disabled={isFetchingHistory} className="px-8 py-3 bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-2xl text-[10px] font-black text-slate-500 flex items-center justify-center gap-3 shadow-sm hover:shadow-lg transition-all uppercase tracking-widest">{isFetchingHistory ? <Loader2 className="w-4 h-4 animate-spin text-emerald-500" /> : <RefreshCw className="w-4 h-4" />} SYNC DATA</button>
                    <div className="relative">
                      <button onClick={() => setIsHistoryMenuOpen(!isHistoryMenuOpen)} className="px-4 py-3 bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-2xl text-slate-500 flex items-center justify-center gap-2 shadow-sm hover:shadow-lg transition-all"><MoreVertical className="w-5 h-5" /></button>
                      {isHistoryMenuOpen && (
                        <div className="absolute right-0 mt-2 w-48 bg-white dark:bg-slate-800 border dark:border-slate-700 rounded-2xl shadow-xl z-50 overflow-hidden">
                          <button onClick={() => { exportHistoryData(); setIsHistoryMenuOpen(false); }} className="w-full px-6 py-4 text-left flex items-center gap-3 hover:bg-emerald-50 dark:hover:bg-emerald-900/30 transition-all border-b dark:border-slate-700 text-[10px] font-black uppercase tracking-widest text-slate-700 dark:text-slate-300"><Download className="w-4 h-4 text-emerald-600" /> {t.export}</button>
                          <button onClick={() => { deleteHistoryData(); setIsHistoryMenuOpen(false); }} className="w-full px-6 py-4 text-left flex items-center gap-3 hover:bg-rose-50 dark:hover:bg-rose-900/30 transition-all text-[10px] font-black uppercase tracking-widest text-rose-600 dark:text-rose-400"><Trash2 className="w-4 h-4" /> {t.delete_data}</button>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
                <div className={cn("rounded-[48px] border shadow-2xl overflow-hidden min-h-[550px] transition-all duration-500", theme === 'dark' ? "bg-slate-900 border-slate-800 shadow-emerald-950/20" : "bg-white border-slate-200 shadow-slate-200/50")}>
                  <div className="overflow-x-auto"><table className="w-full text-left min-w-[800px] border-collapse"><thead className="bg-slate-50/50 dark:bg-slate-800/80 backdrop-blur sticky top-0 text-[10px] uppercase font-black text-slate-400 border-b dark:border-slate-800 z-10 tracking-[0.2em]"><tr className="px-10"><th className="pl-10 py-6 w-14">#</th><th className="px-6 py-6">Title</th><th className="px-6 py-6">Recipient</th><th className="px-6 py-6 text-center">Status</th></tr></thead><tbody className="divide-y divide-slate-50 dark:divide-slate-800 text-[11px] font-bold">
                    {historyRows.map((row, i) => (<tr key={i} className="hover:bg-slate-50/30 dark:hover:bg-slate-800/30 transition-all duration-300"><td className="pl-10 py-6 text-slate-300 dark:text-slate-700 font-black">{i+1}</td><td className="px-6 py-6"><span className="px-3 py-1.5 bg-emerald-50 dark:bg-emerald-900/30 text-emerald-800 dark:text-emerald-400 rounded-xl text-[9px] font-black uppercase tracking-tight shadow-sm leading-none">{row['Judul broadcast']}</span></td><td className="px-6 py-6 font-mono tracking-tighter"><div>{row['nomor Tujuan']}</div></td><td className="px-6 py-6 text-center"><span className={cn("px-4 py-2 rounded-2xl text-[9px] font-black uppercase inline-flex items-center gap-1.5 shadow-sm", row['status']?.toLowerCase().includes('success') ? "bg-emerald-100 text-emerald-700" : "bg-rose-100 text-rose-700")}>{row['status'] === 'Success' ? <CheckCircle2 className="w-3.5 h-3.5" /> : <XCircle className="w-3.5 h-3.5" />}{row['status']}</span></td></tr>))}
                  </tbody></table></div>
                  {historyRows.length === 0 && <div className="h-[400px] flex flex-col items-center justify-center gap-4 text-slate-300 dark:text-slate-700"><Database className="w-16 h-16 opacity-20" /><p className="font-black uppercase tracking-[0.3em] text-[10px]">No records detected</p></div>}
                </div>
              </div>
            )}
          </div>
        </main>
      </div>

      {/* MODAL: Global Message Editor */}
      {isMessageModalOpen && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-md animate-in fade-in duration-300">
          <div className={cn("rounded-[48px] shadow-[0_32px_80px_rgba(0,0,0,0.3)] w-full max-w-lg overflow-hidden border animate-in zoom-in duration-300", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-100")}>
            <div className="p-8 lg:p-10 border-b dark:border-slate-800 flex justify-between items-center bg-slate-50/50 dark:bg-slate-800/50 px-10 leading-none"><h3 className="font-black text-slate-800 dark:text-white uppercase tracking-tight text-xl flex items-center gap-3"><MessageSquare className="w-6 h-6 text-emerald-500" /> Message Editor</h3><button onClick={() => setIsMessageModalOpen(false)} className="text-slate-300 hover:text-rose-500 transition-colors leading-none p-2"><X className="w-7 h-7" /></button></div>
            <div className="p-10 space-y-8 leading-none">
              <div className="bg-emerald-50 dark:bg-emerald-900/10 p-5 rounded-3xl border border-emerald-100 dark:border-emerald-800 text-[11px] text-emerald-800 dark:text-emerald-400 font-bold leading-relaxed shadow-sm">Pesan ini akan diterapkan ke <strong>seluruh antrean</strong>. Gunakan tag <code className="bg-white dark:bg-slate-800 px-2 py-0.5 rounded text-emerald-600">{"{nama}"}</code>.</div>
              <textarea value={bulkMessageInput} onChange={(e) => setBulkMessageInput(e.target.value)} placeholder="Tulis pesan massal Anda di sini..." className={cn("w-full h-48 p-6 border rounded-[32px] text-sm outline-none transition-all resize-none font-bold shadow-inner leading-relaxed", theme === 'dark' ? "bg-slate-800 border-slate-700 text-white" : "bg-slate-50 border-slate-100 text-slate-700")} />
              <div className="flex gap-4 pt-4"><button onClick={() => setIsMessageModalOpen(false)} className="flex-1 py-5 border rounded-[24px] font-black text-slate-400 uppercase tracking-widest text-[10px] hover:bg-slate-50 transition-all">{t.cancel}</button><button onClick={() => { if (!bulkMessageInput) return; setData(prev => prev.map(r => ({ ...r, Pesan: bulkMessageInput }))); setIsMessageModalOpen(false); addLog('System', 'info', 'Message Applied'); }} className="flex-1 py-5 bg-emerald-600 text-white rounded-[24px] font-black shadow-2xl uppercase tracking-widest text-[10px] hover:bg-emerald-700 transition-all active:scale-95 shadow-emerald-500/20">{t.apply}</button></div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: Attachment / Lampiran */}
      {isAttachmentModalOpen && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-md animate-in fade-in duration-300">
          <div className={cn("rounded-[48px] shadow-[0_32px_80px_rgba(0,0,0,0.3)] w-full max-w-lg overflow-hidden border animate-in zoom-in duration-300", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-100")}>
            <div className="p-8 lg:p-10 border-b dark:border-slate-800 flex justify-between items-center bg-slate-50/50 dark:bg-slate-800/50 px-10 leading-none"><h3 className="font-black text-slate-800 dark:text-white uppercase tracking-tight text-xl flex items-center gap-3 leading-none"><Paperclip className="w-6 h-6 text-emerald-500" /> {t.attachment}</h3><button onClick={() => setIsAttachmentModalOpen(false)} className="text-slate-300 hover:text-rose-500 transition-colors p-2 leading-none"><X className="w-7 h-7" /></button></div>
            <div className="p-10 space-y-10">
              <div className="bg-amber-50 dark:bg-amber-900/10 p-6 rounded-3xl border border-amber-100 dark:border-amber-800 flex gap-4 leading-relaxed">
                <AlertTriangle className="w-6 h-6 text-amber-600 shrink-0" />
                <p className="text-xs lg:text-sm text-amber-800 dark:text-amber-400 font-bold leading-relaxed">{t.package_warning}</p>
              </div>

              {!hasPaidPackage ? (
                <div className="space-y-6 text-center leading-none">
                  <div className="bg-slate-50 dark:bg-slate-800 p-8 rounded-[32px] border dark:border-slate-700">
                    <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-2 italic">DETEKSI PAKET: {deviceInfo?.package || 'Unknown'}</p>
                    <p className="text-slate-600 dark:text-slate-300 font-black text-xs uppercase leading-relaxed">Fitur media memerlukan paket berbayar<br />(Super/Advanced/Ultra).</p>
                  </div>
                  <a href="https://fonnte.com/#harga" target="_blank" rel="noopener noreferrer" className="w-full py-6 bg-emerald-600 text-white rounded-[24px] font-black text-[10px] tracking-widest uppercase shadow-2xl flex items-center justify-center gap-3 hover:bg-emerald-700 transition-all animate-pulse shadow-emerald-500/20">
                    <Zap className="w-5 h-5 fill-current" /> {t.subscribe} <ExternalLink className="w-4 h-4" />
                  </a>
                </div>
              ) : (
                <div className="space-y-6 leading-none">
                  <div className="space-y-3">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Attachment URL (Public Image/File)</label>
                    <input 
                      type="text" 
                      value={attachmentUrl} 
                      onChange={(e) => setAttachmentUrl(e.target.value)}
                      placeholder="https://example.com/image.jpg"
                      className={cn("w-full px-8 py-5 border rounded-3xl outline-none font-bold text-sm focus:ring-4 focus:ring-emerald-500/10 transition-all", theme === 'dark' ? "bg-slate-800 border-slate-700 text-white" : "bg-slate-50 border-slate-100 text-slate-700")} 
                    />
                  </div>
                  <button 
                    onClick={() => {
                      setData(prev => prev.map(r => ({ ...r, Lampiran: attachmentUrl })));
                      setIsAttachmentModalOpen(false);
                      addLog('System', 'success', 'Media Attachment Applied');
                    }}
                    className="w-full py-6 bg-emerald-600 text-white rounded-[24px] font-black text-[10px] tracking-widest uppercase shadow-2xl hover:bg-emerald-700 transition-all shadow-emerald-500/20"
                  >
                    APPLY MEDIA URL
                  </button>
                </div>
              )}
              <button onClick={() => setIsAttachmentModalOpen(false)} className="w-full text-center text-[9px] font-black text-slate-300 uppercase tracking-[0.2em] hover:text-slate-600 transition-colors">{t.cancel}</button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: Select Broadcast for Export/Delete */}
      {isSelectBroadcastModalOpen && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-md animate-in fade-in duration-300">
          <div className={cn("rounded-[48px] shadow-[0_32px_80px_rgba(0,0,0,0.3)] w-full max-w-lg overflow-hidden border animate-in zoom-in duration-300", theme === 'dark' ? "bg-slate-900 border-slate-800" : "bg-white border-slate-100")}>
            <div className="p-8 lg:p-10 border-b dark:border-slate-800 flex justify-between items-center bg-slate-50/50 dark:bg-slate-800/50 px-10 leading-none">
              <h3 className="font-black text-slate-800 dark:text-white uppercase tracking-tight text-lg flex items-center gap-3">
                {selectBroadcastAction === 'export' ? (
                  <>
                    <Download className="w-6 h-6 text-emerald-500" /> Pilih {lang === 'id' ? 'untuk Di Export' : 'to Export'}
                  </>
                ) : (
                  <>
                    <Trash2 className="w-6 h-6 text-rose-500" /> Pilih {lang === 'id' ? 'untuk Dihapus' : 'to Delete'}
                  </>
                )}
              </h3>
              <button onClick={() => setIsSelectBroadcastModalOpen(false)} className="text-slate-300 hover:text-rose-500 transition-colors leading-none p-2">
                <X className="w-7 h-7" />
              </button>
            </div>
            <div className="p-10 space-y-4 leading-none max-h-96 overflow-y-auto">
              {/* Option: All */}
              <button 
                onClick={() => {
                  if (selectBroadcastAction === 'export') {
                    executeExport(null);
                  } else {
                    executeDelete(null);
                  }
                  setIsSelectBroadcastModalOpen(false);
                  setSelectBroadcastAction(null);
                }}
                className={cn("w-full p-4 rounded-2xl border-2 text-left font-black text-sm uppercase tracking-wider transition-all hover:scale-102", 
                  theme === 'dark' 
                    ? "border-slate-700 bg-slate-800 text-slate-100 hover:bg-slate-700" 
                    : "border-slate-200 bg-slate-50 text-slate-800 hover:bg-slate-100"
                )}
              >
                {lang === 'id' ? '📋 Semua Broadcast' : '📋 All Broadcasts'}
              </button>

              {/* Divider */}
              <div className={cn("h-px", theme === 'dark' ? "bg-slate-700" : "bg-slate-200")} />

              {/* Individual Broadcast Options */}
              {Array.from(new Set(historyRows.map(r => r['Judul broadcast']))).map((title, idx) => (
                <button
                  key={idx}
                  onClick={() => {
                    if (selectBroadcastAction === 'export') {
                      executeExport(title);
                    } else {
                      executeDelete(title);
                    }
                    setIsSelectBroadcastModalOpen(false);
                    setSelectBroadcastAction(null);
                  }}
                  className={cn("w-full p-4 rounded-2xl border-2 text-left font-black text-sm uppercase tracking-wider transition-all hover:scale-102", 
                    theme === 'dark' 
                      ? "border-emerald-700/50 bg-slate-800 text-emerald-400 hover:bg-slate-700 hover:border-emerald-600" 
                      : "border-emerald-200 bg-emerald-50 text-emerald-700 hover:bg-emerald-100 hover:border-emerald-400"
                  )}
                >
                  📄 {title}
                </button>
              ))}
            </div>
            <div className="p-4 border-t dark:border-slate-800 bg-slate-50/50 dark:bg-slate-800/50">
              <button onClick={() => setIsSelectBroadcastModalOpen(false)} className="w-full py-3 border rounded-xl font-black text-slate-400 uppercase tracking-widest text-[10px] hover:bg-slate-100 dark:hover:bg-slate-800 transition-all">
                {t.cancel}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
