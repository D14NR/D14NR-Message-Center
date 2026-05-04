import React, { useCallback, useEffect, useRef, useState } from "react";
import axios from "axios";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  AlertTriangle,
  ArrowRight,
  CheckCircle2,
  ChevronDown,
  Clock,
  Database,
  Download,
  ExternalLink,
  FileSpreadsheet,
  Globe,
  History as HistoryIcon,
  Info,
  LayoutDashboard,
  Loader2,
  Lock,
  LogOut,
  Menu,
  MessageSquare,
  Moon,
  MoreVertical,
  Paperclip,
  RefreshCw,
  Search,
  Settings,
  ShieldCheck,
  Sun,
  Trash2,
  Type,
  Upload,
  User,
  X,
  XCircle,
  Zap,
} from "lucide-react";
import { clsx, type ClassValue } from "clsx";
import { motion } from "framer-motion";
import { twMerge } from "tailwind-merge";

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

const translations = {
  id: {
    dashboard: "Dashboard",
    broadcast: "Broadcast",
    configuration: "Konfigurasi",
    history: "Riwayat",
    status: "STATUS",
    total_data: "TOTAL DATA",
    pending: "TERTUNDA",
    success: "SUKSES",
    failed: "GAGAL",
    start_now: "MULAI SEKARANG",
    system_config: "Pengaturan Sistem",
    api_key: "Fonnte API Token Key",
    theme: "Tampilan Tema",
    language: "Pilihan Bahasa",
    connect: "HUBUNGKAN",
    connected: "TERHUBUNG",
    not_connected: "BELUM TERHUBUNG",
    disconnect: "PUTUSKAN",
    light: "TERANG",
    dark: "GELAP",
    title: "Judul Broadcast",
    excel_upload: "Unggah Excel",
    manual_add: "Tambah Manual",
    message: "Isi Pesan",
    apply: "TERAPKAN",
    cancel: "BATAL",
    attachment: "Lampiran",
    subscribe: "Berlangganan Paket",
    package_warning:
      "Pastikan Anda sudah memiliki paket yang mendukung fitur kirim lampiran.",
    delete_data: "Hapus Data",
    export: "Export",
  },
  en: {
    dashboard: "Dashboard",
    broadcast: "Broadcast",
    configuration: "Configuration",
    history: "History",
    status: "STATUS",
    total_data: "TOTAL DATA",
    pending: "PENDING",
    success: "SUCCESS",
    failed: "FAILED",
    start_now: "START NOW",
    system_config: "System Configuration",
    api_key: "Fonnte API Token Key",
    theme: "Display Theme",
    language: "Language",
    connect: "CONNECT",
    connected: "CONNECTED",
    not_connected: "NOT CONNECTED",
    disconnect: "DISCONNECT",
    light: "LIGHT",
    dark: "DARK",
    title: "Broadcast Title",
    excel_upload: "Excel Upload",
    manual_add: "Manual Add",
    message: "Message",
    apply: "APPLY",
    cancel: "CANCEL",
    attachment: "Attachment",
    subscribe: "Subscribe Package",
    package_warning:
      "Make sure you have a package that supports attachment sending feature.",
    delete_data: "Delete Data",
    export: "Export",
  },
};

interface SpreadsheetRow {
  "Nomor Telepon": string;
  Nama: string;
  Pesan: string;
  Status: string;
  Lampiran?: string;
  __id?: number;
}

interface HistoryRow {
  Device: string;
  "nomor device": string;
  "Judul broadcast": string;
  "nomor Tujuan": string;
  "tanggal pengiriman": string;
  status: string;
}

type View = "dashboard" | "broadcast" | "configuration" | "history";
type NotificationVariant = "success" | "error" | "info" | "warning";

interface NotificationItem {
  id: string;
  variant: NotificationVariant;
  title: string;
  message: string;
}

const HISTORY_SHEET_URL =
  "https://docs.google.com/spreadsheets/d/17YfjibgKzXX6S7UJEC0OGoDFNwbvqnn5R1rpQOtDhDI/export?format=csv&gid=685286577";
const LOG_WEBHOOK =
  "https://script.google.com/macros/s/AKfycbxGvYq05Cr01naPccLtEw9oJgR0-Gwkm7hrWP8ebl9ltmvZWBtLAxwtWGW0vkTRfWo8/exec";

export default function App() {
  const [isLoggedIn, setIsLoggedIn] = useState(
    () => localStorage.getItem("isLoggedIn") === "true"
  );
  const [activeView, setActiveView] = useState<View>("dashboard");
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  const [apiKey, setApiKey] = useState(() => localStorage.getItem("fonnte_api_key") || "");
  const [isTokenChanged, setIsTokenChanged] = useState(false);
  const [lang, setLang] = useState<"id" | "en">(
    () => (localStorage.getItem("app_lang") as "id" | "en") || "id"
  );
  const [theme, setTheme] = useState<"light" | "dark">(
    () => (localStorage.getItem("app_theme") as "light" | "dark") || "light"
  );

  const [data, setData] = useState<SpreadsheetRow[]>([]);
  const [isBroadcasting, setIsBroadcasting] = useState(false);
  const [currentIndex, setCurrentIndex] = useState(-1);
  const [logs, setLogs] = useState<any[]>([]);
  const [delayRange, setDelayRange] = useState("25-40");
  const [broadcastTitle, setBroadcastTitle] = useState("");
  const isBroadcastingRef = useRef(false);

  const [manualPhone, setManualPhone] = useState("");
  const [manualName, setManualName] = useState("");
  const [manualMessage, setManualMessage] = useState("");
  const [autoCheck] = useState(true);
  const [deviceInfo, setDeviceInfo] = useState<any>(null);
  const [isCheckingConnection, setIsCheckingConnection] = useState(false);
  const [isCheckingWA, setIsCheckingWA] = useState(false);

  const [historyData, setHistoryData] = useState<HistoryRow[]>([]);
  const [isFetchingHistory, setIsFetchingHistory] = useState(false);
  const [isHistoryMenuOpen, setIsHistoryMenuOpen] = useState(false);

  const [isMessageModalOpen, setIsMessageModalOpen] = useState(false);
  const [isAttachmentModalOpen, setIsAttachmentModalOpen] = useState(false);
  const [bulkMessageInput, setBulkMessageInput] = useState("");
  const [attachmentUrl, setAttachmentUrl] = useState("");
  const [isSelectBroadcastModalOpen, setIsSelectBroadcastModalOpen] = useState(false);
  const [selectBroadcastAction, setSelectBroadcastAction] = useState<
    "export" | "delete" | null
  >(null);
  const [searchBroadcast, setSearchBroadcast] = useState("");
  const [confirmModal, setConfirmModal] = useState<{
    title: string;
    message: string;
    confirmLabel: string;
    cancelLabel: string;
    onConfirm: () => void;
  } | null>(null);
  const [notifications, setNotifications] = useState<NotificationItem[]>([]);

  const t = translations[lang];
  const viewMotion = {
    initial: { opacity: 0, y: 20 },
    animate: { opacity: 1, y: 0 },
    transition: { duration: 0.35, ease: "easeOut" as const },
  };

  const currentDeviceName = deviceInfo?.name?.trim().toLowerCase();
  const historyRows = currentDeviceName
    ? historyData.filter((r) => r.Device?.trim().toLowerCase() === currentDeviceName)
    : historyData;

  const openConfirmModal = useCallback(
    (options: {
      title: string;
      message: string;
      confirmLabel?: string;
      cancelLabel?: string;
      onConfirm: () => void;
    }) => {
      setConfirmModal({
        title: options.title,
        message: options.message,
        confirmLabel: options.confirmLabel ?? (lang === "id" ? "HAPUS" : "DELETE"),
        cancelLabel: options.cancelLabel ?? (lang === "id" ? "BATAL" : "CANCEL"),
        onConfirm: options.onConfirm,
      });
    },
    [lang]
  );

  const showNotification = useCallback(
    (variant: NotificationVariant, title: string, message: string) => {
      const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
      setNotifications((prev) => [...prev, { id, variant, title, message }]);
      window.setTimeout(
        () => setNotifications((prev) => prev.filter((item) => item.id !== id)),
        4200
      );
    },
    []
  );

  useEffect(() => {
    localStorage.setItem("fonnte_api_key", apiKey);
    localStorage.setItem("app_lang", lang);
    localStorage.setItem("app_theme", theme);
    if (theme === "dark") document.documentElement.classList.add("dark");
    else document.documentElement.classList.remove("dark");
  }, [apiKey, lang, theme]);

  const addLog = useCallback((target: string, status: "success" | "error" | "info", message: string) => {
    setLogs((prev) => [
      { timestamp: new Date().toLocaleTimeString(), target, status, message },
      ...prev,
    ].slice(0, 100));
  }, []);

  const formatPhone = (phone: string) => {
    let cleaned = String(phone).replace(/[^0-9]/g, "");
    if (cleaned.startsWith("0")) cleaned = `62${cleaned.substring(1)}`;
    else if (cleaned.startsWith("8")) cleaned = `62${cleaned}`;
    return cleaned;
  };

  const checkConnection = async (silent = false) => {
    if (!apiKey) return;
    if (!silent) setIsCheckingConnection(true);
    try {
      const response = await axios.post(
        "https://api.fonnte.com/device",
        {},
        { headers: { Authorization: apiKey } }
      );
      if (response.data.status) {
        setDeviceInfo(response.data);
        setIsTokenChanged(false);
      } else if (!silent) {
        setDeviceInfo(null);
      }
    } catch {
      if (!silent) setDeviceInfo(null);
    } finally {
      if (!silent) setIsCheckingConnection(false);
    }
  };

  const fetchHistory = async () => {
    setIsFetchingHistory(true);
    try {
      const response = await axios.get(HISTORY_SHEET_URL);
      const results = Papa.parse<HistoryRow>(response.data, {
        header: true,
        skipEmptyLines: true,
      });
      setHistoryData(results.data);
    } catch (error) {
      console.error(error);
    } finally {
      setIsFetchingHistory(false);
    }
  };

  const exportHistoryData = () => {
    if (historyRows.length === 0) {
      showNotification(
        "info",
        lang === "id" ? "Data Kosong" : "No Data",
        lang === "id" ? "Tidak ada data untuk di export" : "No data to export"
      );
      return;
    }
    setSelectBroadcastAction("export");
    setIsSelectBroadcastModalOpen(true);
  };

  const deleteHistoryData = () => {
    if (historyRows.length === 0) {
      showNotification(
        "info",
        lang === "id" ? "Data Kosong" : "No Data",
        lang === "id" ? "Tidak ada data untuk dihapus" : "No data to delete"
      );
      return;
    }
    setSelectBroadcastAction("delete");
    setIsSelectBroadcastModalOpen(true);
  };

  const executeExport = (title: string | null) => {
    const dataToExport = title
      ? historyRows.filter((r) => r["Judul broadcast"] === title)
      : historyRows;

    if (dataToExport.length === 0) {
      showNotification(
        "info",
        lang === "id" ? "Data Kosong" : "No Data",
        lang === "id" ? "Tidak ada data untuk di export" : "No data to export"
      );
      return;
    }

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "History");
    const filename = title
      ? `history_${title}_${new Date().toISOString().split("T")[0]}.xlsx`
      : `history_${new Date().toISOString().split("T")[0]}.xlsx`;
    XLSX.writeFile(wb, filename);
    addLog("System", "success", lang === "id" ? "Data berhasil di export" : "Data exported");
  };

  const executeDelete = (title: string | null) => {
    const confirmMessage =
      lang === "id"
        ? `Hapus data broadcast "${title || "Semua"}"? Tindakan ini tidak dapat dibatalkan.`
        : `Delete "${title || "All"}" broadcast data? This action cannot be undone.`;

    openConfirmModal({
      title: lang === "id" ? "Konfirmasi Hapus" : "Delete Confirmation",
      message: confirmMessage,
      onConfirm: () => {
        if (title) {
          setHistoryData((prev) => prev.filter((r) => r["Judul broadcast"] !== title));
        } else {
          setHistoryData([]);
        }
        setIsHistoryMenuOpen(false);
        addLog("System", "info", lang === "id" ? "Data histori dihapus" : "History data deleted");
      },
    });
  };

  const getBroadcastList = () => {
    const allBroadcasts = Array.from(new Set(historyRows.map((r) => r["Judul broadcast"])));
    if (!searchBroadcast.trim()) return allBroadcasts;
    return allBroadcasts.filter((b) => b.toLowerCase().includes(searchBroadcast.toLowerCase()));
  };

  const getCountForBroadcast = (title: string | null) => {
    if (!title) return historyRows.length;
    return historyRows.filter((r) => r["Judul broadcast"] === title).length;
  };

  useEffect(() => {
    if (activeView === "history" && isLoggedIn) fetchHistory();
  }, [activeView, isLoggedIn]);

  const logToSheet = async (target: string, status: string) => {
    try {
      const payload = {
        action: "UPSERT",
        device: deviceInfo?.name || "Unknown",
        nomor_device: deviceInfo?.device || "-",
        judul: broadcastTitle,
        tujuan: target,
        tanggal: new Date().toLocaleDateString("id-ID"),
        status,
      };
      await axios.post(LOG_WEBHOOK, JSON.stringify(payload));
    } catch (error) {
      console.error(error);
    }
  };

  const bulkCheckWhatsApp = async (rowsToCheck: SpreadsheetRow[]) => {
    if (!apiKey || rowsToCheck.length === 0) return;
    setIsCheckingWA(true);
    try {
      const targets = rowsToCheck.map((r) => formatPhone(r["Nomor Telepon"])).join(",");
      const params = new URLSearchParams();
      params.append("target", targets);
      const res = await axios.post("https://api.fonnte.com/validate", params, {
        headers: { Authorization: apiKey },
      });
      if (res.data.status) {
        const reg: string[] = res.data.registered || [];
        const notReg: string[] = res.data.not_registered || [];
        setData((prev) =>
          prev.map((row) => {
            const f = formatPhone(row["Nomor Telepon"]);
            if (reg.includes(f)) return { ...row, Status: "WA Active" };
            if (notReg.includes(f)) return { ...row, Status: "No WA" };
            return row;
          })
        );
      }
    } catch (error) {
      console.error(error);
    } finally {
      setIsCheckingWA(false);
    }
  };

  const sendMessage = async (row: SpreadsheetRow): Promise<{ success: boolean; message: string }> => {
    const target = formatPhone(row["Nomor Telepon"]);
    const message = row.Pesan.replace(/{nama}/gi, row.Nama);
    try {
      const params = new URLSearchParams();
      params.append("target", target);
      params.append("message", message);
      if (row.Lampiran) params.append("url", row.Lampiran);
      params.append("delay", "2");
      params.append("countryCode", "62");
      const response = await axios.post("https://api.fonnte.com/send", params, {
        headers: {
          Authorization: apiKey,
          "Content-Type": "application/x-www-form-urlencoded",
        },
      });
      return {
        success: response.data.status === true || response.data.status === "true",
        message: response.data.detail || response.data.reason || "Processed",
      };
    } catch (error: any) {
      return { success: false, message: error.message };
    }
  };

  const waitSeconds = (seconds: number) =>
    new Promise<void>((resolve) => {
      setTimeout(resolve, seconds * 1000);
    });

  const getRandomDelaySeconds = () => {
    const [mn, mx] = delayRange.split("-").map(Number);
    return Math.floor(Math.random() * (mx - mn + 1) + mn);
  };

  const getRandomStageDelaySeconds = () =>
    Math.floor(Math.random() * (60 - 30 + 1) + 30);

  const startBroadcast = async () => {
    if (isBroadcasting) {
      setIsBroadcasting(false);
      isBroadcastingRef.current = false;
      return;
    }
    if (!broadcastTitle) {
      showNotification("warning", "Judul Diperlukan", "WAJIB: Judul Broadcast!");
      return;
    }
    const validStatuses = ["Pending", "WA Active", ""];
    const queuedIndexes = data
      .map((row, index) => (validStatuses.includes(row.Status) ? index : -1))
      .filter((index) => index >= 0);

    if (queuedIndexes.length === 0) {
      showNotification("info", "Tidak Ada Pending", "No pending rows available.");
      return;
    }

    setIsBroadcasting(true);
    isBroadcastingRef.current = true;
    const maxRetries = 2;
    const stageSize = 25;
    const totalStages = Math.ceil(queuedIndexes.length / stageSize);

    for (let position = 0; position < queuedIndexes.length; position += 1) {
      if (!isBroadcastingRef.current) break;
      const i = queuedIndexes[position];
      const row = data[i];
      const currentStage = Math.floor(position / stageSize) + 1;

      if (position % stageSize === 0) {
        addLog(
          "System",
          "info",
          `Starting stage ${currentStage}/${totalStages} (${Math.min(
            stageSize,
            queuedIndexes.length - position
          )} data)`
        );
      }

      setCurrentIndex(i);
      let attempt = 1;
      let res = await sendMessage(row);

      while (!res.success && attempt <= maxRetries && isBroadcastingRef.current) {
        attempt += 1;
        addLog(
          row["Nomor Telepon"],
          "info",
          `Retry ${attempt - 1}/${maxRetries} in 5s...`
        );
        await waitSeconds(5);
        if (!isBroadcastingRef.current) break;
        res = await sendMessage(row);
      }

      const status = res.success ? "Success" : "Failed";
      setData((prev) => prev.map((r, idx) => (idx === i ? { ...r, Status: status } : r)));
      addLog(
        row["Nomor Telepon"],
        res.success ? "success" : "error",
        `${res.message} (attempt ${attempt})`
      );
      logToSheet(row["Nomor Telepon"], status);
      if (res.success) checkConnection(true);

      const hasNextMessage = position < queuedIndexes.length - 1;
      if (!hasNextMessage || !isBroadcastingRef.current) continue;

      const endOfStage = (position + 1) % stageSize === 0;
      if (endOfStage) {
        const stageDelay = getRandomStageDelaySeconds();
        addLog(
          "System",
          "info",
          `Stage ${currentStage} completed. Waiting ${stageDelay}s before next stage...`
        );
        await waitSeconds(stageDelay);
      } else {
        const rd = getRandomDelaySeconds();
        addLog("System", "info", `Waiting ${rd}s for next message...`);
        await waitSeconds(rd);
      }
    }
    setIsBroadcasting(false);
    isBroadcastingRef.current = false;
    setCurrentIndex(-1);
  };

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!broadcastTitle) {
      showNotification("warning", "Judul Diperlukan", "Isi Judul!");
      e.target.value = "";
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const bstr = ev.target?.result;
      const wb = XLSX.read(bstr, { type: "binary" });
      const results = XLSX.utils.sheet_to_json<any>(wb.Sheets[wb.SheetNames[0]]);
      const newRows: SpreadsheetRow[] = results
        .map((row, index) => ({
          "Nomor Telepon": String(row["Nomor Telepon"] || row.phone || ""),
          Nama: String(row.Nama || row.name || "Recipient"),
          Pesan: "",
          Status: "Pending",
          __id: Date.now() + index,
        }))
        .filter((row) => row["Nomor Telepon"] !== "");
      setData((prev) => [...prev, ...newRows]);
      if (autoCheck) bulkCheckWhatsApp(newRows);
      e.target.value = "";
    };
    reader.readAsBinaryString(file);
  };

  const SidebarItem = ({
    id,
    label,
    icon: Icon,
  }: {
    id: View;
    label: string;
    icon: any;
  }) => {
    const handleClick = () => {
      if (id === "broadcast" && !deviceInfo) {
        showNotification("warning", "Butuh Token", "Silakan hubungkan token Fonnte!");
        setActiveView("configuration");
        setIsMobileMenuOpen(false);
        return;
      }
      setActiveView(id);
      setIsMobileMenuOpen(false);
    };
    return (
      <button
        onClick={handleClick}
        className={cn(
          "w-full rounded-xl px-3 py-2.5 flex items-center gap-3 transition-colors",
          activeView === id
            ? "bg-emerald-500 text-white"
            : "text-slate-500 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800"
        )}
      >
        <Icon size={18} />
        <span className="font-semibold text-sm">{label}</span>
      </button>
    );
  };

  const stats = {
    total: data.length,
    pending: data.filter((r) => ["Pending", "WA Active", ""].includes(r.Status)).length,
    success: data.filter((r) => r.Status === "Success").length,
    failed: data.filter((r) => r.Status === "Failed").length,
  };
  const invalidWaCount = data.filter((r) => r.Status === "No WA").length;
  const allWaActive = data.length > 0 && data.every((r) => r.Status === "WA Active");

  const hasPaidPackage =
    deviceInfo?.package &&
    ["super", "advanced", "ultra"].includes(deviceInfo.package.toLowerCase());

  if (!isLoggedIn) {
    return (
      <div className="relative min-h-screen overflow-hidden bg-[radial-gradient(circle_at_top_right,#4adbe533,transparent_40%),radial-gradient(circle_at_bottom_left,#0f8ea133,transparent_45%)]">
        <div className="pointer-events-none absolute inset-0">
          <div className="absolute -top-24 -left-12 h-72 w-72 rounded-full bg-cyan-400/20 blur-3xl" />
          <div className="absolute -bottom-24 -right-12 h-72 w-72 rounded-full bg-indigo-500/20 blur-3xl" />
        </div>

        <div className="relative z-10 mx-auto flex min-h-screen max-w-6xl items-center px-4 py-10 lg:px-8">
          <div className="grid w-full overflow-hidden rounded-[2rem] border border-cyan-200/70 bg-white/90 backdrop-blur-xl shadow-2xl lg:grid-cols-2">
            <motion.section
              initial={{ opacity: 0, x: -20 }}
              animate={{ opacity: 1, x: 0 }}
              transition={{ duration: 0.45 }}
              className="relative hidden lg:flex flex-col justify-between bg-gradient-to-br from-indigo-700 via-indigo-600 to-cyan-600 p-10 text-white"
            >
              <div className="space-y-6">
                <div className="inline-flex items-center gap-3">
                  <img src="/logo.png" alt="D14NR Logo" className="h-12 w-12 rounded-2xl object-cover ring-2 ring-white/30" />
                  <div>
                    <p className="text-2xl font-black tracking-tight">D14NR</p>
                    <p className="text-[11px] font-bold tracking-[0.2em] uppercase text-cyan-100">Message Center</p>
                  </div>
                </div>
                <div className="space-y-3">
                  <p className="text-xs font-black tracking-[0.24em] uppercase text-cyan-100">Broadcast Console</p>
                  <h2 className="text-4xl font-black leading-tight">Kirim pesan WA secara terstruktur, cepat, dan aman.</h2>
                  <p className="max-w-sm text-sm leading-relaxed text-cyan-50">Satu dashboard untuk upload data, validasi nomor, broadcast bertahap, dan monitoring riwayat pengiriman.</p>
                </div>
              </div>
              <div className="inline-flex items-center gap-2 rounded-xl border border-white/30 bg-white/10 px-3 py-2 text-xs font-bold uppercase tracking-widest">
                <ShieldCheck className="h-4 w-4" /> Secure Internal Access
              </div>
            </motion.section>

            <motion.section
              initial={{ opacity: 0, y: 16 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.4 }}
              className="p-6 sm:p-10 lg:p-12"
            >
              <div className="mb-8 text-center lg:text-left">
                <div className="mx-auto mb-4 inline-flex rounded-2xl overflow-hidden shadow lg:mx-0">
                  <img src="/logo.png" alt="D14NR Logo" className="h-12 w-12 object-cover" />
                </div>
                <h1 className="text-3xl font-black tracking-tight text-slate-900 dark:text-white">Welcome Back</h1>
                <p className="mt-2 text-sm font-semibold text-slate-600">Masuk ke D14NR Message Center</p>
              </div>

              <form
                onSubmit={(e) => {
                  e.preventDefault();
                  const fd = new FormData(e.currentTarget);
                  if (fd.get("user") === "D14nr" && fd.get("pass") === "290192") {
                    setIsLoggedIn(true);
                    localStorage.setItem("isLoggedIn", "true");
                  } else {
                    showNotification("error", "Login Gagal", "Invalid Credentials");
                  }
                }}
                className="space-y-4"
              >
                <label className="relative block">
                  <User className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-500" />
                  <input
                    name="user"
                    type="text"
                    placeholder="Username"
                    className="w-full rounded-xl border border-slate-300 bg-slate-50 py-3 pl-10 pr-3 text-sm font-semibold text-slate-800 outline-none transition focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20"
                  />
                </label>
                <label className="relative block">
                  <Lock className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-500" />
                  <input
                    name="pass"
                    type="password"
                    placeholder="Password"
                    className="w-full rounded-xl border border-slate-300 bg-slate-50 py-3 pl-10 pr-3 text-sm font-semibold text-slate-800 outline-none transition focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20"
                  />
                </label>
                <motion.button whileHover={{ scale: 1.01 }} whileTap={{ scale: 0.98 }} className="flex w-full items-center justify-center gap-2 rounded-xl bg-indigo-600 py-3 font-black uppercase tracking-wider text-white shadow-lg shadow-indigo-500/30">
                  SIGN IN <ArrowRight className="h-4 w-4" />
                </motion.button>
              </form>
            </motion.section>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div
      className={cn(
        "app-shell relative flex h-screen font-sans overflow-hidden transition-all duration-500",
        theme === "dark" ? "bg-slate-950 text-slate-100" : "bg-slate-50 text-slate-900"
      )}
    >
      <div className="pointer-events-none absolute inset-0 overflow-hidden">
        <div className="absolute -top-28 -right-24 h-72 w-72 rounded-full bg-cyan-400/20 blur-3xl" />
        <div className="absolute -bottom-28 -left-20 h-72 w-72 rounded-full bg-indigo-500/20 blur-3xl" />
      </div>
      {isMobileMenuOpen && (
        <div
          className="fixed inset-0 bg-black/40 backdrop-blur-sm z-40 lg:hidden"
          onClick={() => setIsMobileMenuOpen(false)}
        />
      )}
      <aside
        className={cn(
          "fixed lg:relative inset-y-0 left-0 w-64 border-r flex flex-col z-50 transition-all duration-300 transform lg:translate-x-0 backdrop-blur-xl",
          isMobileMenuOpen ? "translate-x-0" : "-translate-x-full",
          theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200"
        )}
      >
        <div className="p-6 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="rounded-xl overflow-hidden shadow-lg">
              <img src="/logo.png" alt="D14NR Logo" className="w-10 h-10 object-cover" />
            </div>
            <div className="flex flex-col">
              <h1 className="text-sm font-black tracking-tighter leading-none uppercase">D14NR</h1>
              <p className="text-[10px] font-bold text-indigo-600 dark:text-indigo-400 uppercase tracking-widest leading-none mt-1">
                Message Center
              </p>
            </div>
          </div>
            <button className="lg:hidden text-slate-500 dark:text-slate-200" onClick={() => setIsMobileMenuOpen(false)}>
            <X className="w-6 h-6" />
          </button>
        </div>
        <nav className="flex-1 p-4 space-y-1 overflow-y-auto">
          <SidebarItem id="dashboard" label={t.dashboard} icon={LayoutDashboard} />
          <SidebarItem id="broadcast" label={t.broadcast} icon={MessageSquare} />
          <SidebarItem id="history" label={t.history} icon={HistoryIcon} />
          <div className="pt-4 mt-4 border-t border-slate-100 dark:border-slate-800 text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">
            System
          </div>
          <SidebarItem id="configuration" label={t.configuration} icon={Settings} />
        </nav>
        <div className="p-4 border-t border-slate-100 dark:border-slate-800 space-y-4 text-center">
          <div className="bg-slate-50 dark:bg-slate-800/50 p-3 rounded-2xl border border-slate-200 dark:border-slate-800 flex items-center gap-3 text-left">
            <div
              className={cn(
                "w-2.5 h-2.5 rounded-full",
                deviceInfo?.device_status === "connect"
                  ? "bg-indigo-500 animate-pulse"
                  : "bg-slate-300 dark:bg-slate-600"
              )}
            />
            <div className="flex flex-col leading-tight">
              <span className="text-[10px] font-bold text-slate-400 dark:text-slate-300 uppercase">
                {t.status}
              </span>
              <span className="text-[11px] font-black truncate uppercase tracking-tighter">
                {deviceInfo?.name || "OFFLINE"}
              </span>
            </div>
          </div>
          <p className="text-[9px] font-black text-slate-400 dark:text-slate-500 uppercase tracking-widest leading-none">
            Copyright D14nr 2026
          </p>
        </div>
      </aside>

      <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
        <header
          className={cn(
            "h-16 lg:h-14 border-b flex items-center justify-between px-4 lg:px-8 z-30 shrink-0 backdrop-blur-xl",
            theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200"
          )}
        >
          <div className="flex items-center gap-3">
            <button
              className="lg:hidden p-2 text-slate-500 dark:text-white hover:bg-slate-100 dark:hover:bg-slate-800 rounded-lg"
              onClick={() => setIsMobileMenuOpen(true)}
            >
              <Menu className="w-6 h-6" />
            </button>
            <h2 className="text-[10px] font-black text-slate-400 dark:text-slate-300 uppercase tracking-widest">
              {activeView}
            </h2>
            {isCheckingWA && (
              <span className="ui-chip ui-chip-info">
                checking wa
              </span>
            )}
          </div>
          <div className="flex items-center gap-4">
            {isBroadcasting && (
              <div className="hidden sm:flex ui-chip ui-chip-info animate-pulse">
                Broadcasting
              </div>
            )}
            <button
              onClick={() => {
                setIsLoggedIn(false);
                localStorage.removeItem("isLoggedIn");
              }}
              className="flex items-center gap-2 px-4 py-2 bg-indigo-100 dark:bg-indigo-900/20 hover:bg-indigo-200 text-indigo-700 dark:text-indigo-400 rounded-lg font-bold text-[10px] uppercase"
            >
              <LogOut className="w-3.5 h-3.5" /> LOGOUT
            </button>
          </div>
        </header>

        <main className="flex-1 overflow-y-auto p-4 lg:p-8 bg-transparent">
          <div className="max-w-6xl mx-auto space-y-6">
            {activeView === "dashboard" && (
              <motion.div className="space-y-8" {...viewMotion}>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 lg:gap-6">
                  {[
                    { l: t.total_data, v: stats.total, i: Database, c: theme === "dark" ? "text-slate-200" : "text-slate-600" },
                    { l: t.pending, v: stats.pending, i: Loader2, c: "text-amber-600" },
                    { l: t.success, v: stats.success, i: CheckCircle2, c: "text-emerald-600" },
                    { l: t.failed, v: stats.failed, i: XCircle, c: "text-rose-600" },
                  ].map((s, i) => (
                    <motion.div
                      key={i}
                      initial={{ opacity: 0, y: 14 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{ delay: i * 0.08, duration: 0.3 }}
                      whileHover={{ y: -4 }}
                      className={cn(
                        "p-5 rounded-3xl border flex items-center gap-4 shadow-sm hover:shadow-xl transition-all",
                        theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200"
                      )}
                    >
                      <div className={cn("p-3 rounded-2xl bg-slate-50 dark:bg-slate-800", s.c)}>
                        <s.i className="w-5 h-5" />
                      </div>
                      <div>
                        <p className="text-[10px] font-black text-slate-400 dark:text-slate-300 mb-1">{s.l}</p>
                        <p className={cn("text-2xl font-black", s.c)}>{s.v}</p>
                      </div>
                    </motion.div>
                  ))}
                </div>
                <section className="bg-gradient-to-br from-indigo-600 via-indigo-700 to-indigo-800 p-8 rounded-[40px] text-white shadow-xl overflow-hidden relative border border-white/20">
                  <div className="absolute inset-0 bg-[radial-gradient(circle_at_20%_20%,rgba(255,255,255,0.16),transparent_45%)]" />
                  <div className="space-y-6 relative z-10">
                    <h3 className="text-3xl font-black uppercase tracking-tight">D14NR</h3>
                    <p className="text-cyan-50 max-w-xl leading-relaxed text-sm">
                      D14NR Message Center bekerja sama dengan Fonnte menghadirkan solusi
                      pengiriman pesan WhatsApp yang aman dan efisien.
                    </p>
                    <div className="flex flex-wrap gap-3 text-[10px] font-black uppercase tracking-widest">
                      <span className="bg-white/10 px-4 py-2 rounded-xl flex items-center gap-2"><ShieldCheck className="w-4 h-4" /> Anti-Spam Tech</span>
                      <span className="bg-white/10 px-4 py-2 rounded-xl flex items-center gap-2"><Zap className="w-4 h-4" /> Instant Delivery</span>
                      <span className="bg-white/10 px-4 py-2 rounded-xl flex items-center gap-2"><Globe className="w-4 h-4" /> Global Reach</span>
                    </div>
                    <button
                      onClick={() => setActiveView("broadcast")}
                      className="px-8 py-3 bg-white text-indigo-700 rounded-2xl font-black uppercase text-xs shadow-lg hover:scale-[1.03] transition-transform"
                    >
                      {t.start_now}
                    </button>
                  </div>
                </section>
                <section className={cn("rounded-3xl border p-5 shadow-sm", theme === "dark" ? "border-slate-800 bg-slate-900" : "border-slate-200 bg-white")}>
                  <div className="text-xs font-black uppercase tracking-widest text-slate-500 dark:text-slate-300 mb-4">Recent Logs</div>
                  <div className="space-y-2 max-h-52 overflow-auto">
                    {logs.length === 0 && <div className="text-sm text-slate-500 dark:text-slate-300">No activity yet.</div>}
                    {logs.slice(0, 8).map((log, idx) => (
                      <div key={idx} className="text-xs flex gap-2 items-center">
                        <span className="text-slate-400 dark:text-slate-500 w-20 shrink-0">{log.timestamp}</span>
                        <span className={cn("w-2 h-2 rounded-full shrink-0", log.status === "success" ? "bg-emerald-500" : log.status === "error" ? "bg-rose-500" : "bg-indigo-500")} />
                        <span className="font-semibold">{log.target}</span>
                        <span className="text-slate-500 dark:text-slate-300 truncate">{log.message}</span>
                      </div>
                    ))}
                  </div>
                </section>
              </motion.div>
            )}

            {activeView === "broadcast" && (
              <motion.div className="space-y-6" {...viewMotion}>
                <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
                  {[
                    { l: t.total_data, v: stats.total, i: Database, c: theme === "dark" ? "text-slate-200" : "text-slate-600" },
                    { l: t.pending, v: stats.pending, i: Loader2, c: "text-amber-600" },
                    { l: t.success, v: stats.success, i: CheckCircle2, c: "text-emerald-600" },
                    { l: t.failed, v: stats.failed, i: XCircle, c: "text-rose-600" },
                  ].map((s, i) => (
                    <motion.div
                      key={i}
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{ delay: i * 0.06, duration: 0.25 }}
                      className={cn(
                        "p-3 rounded-2xl border flex items-center gap-3 shadow-sm",
                        theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200"
                      )}
                    >
                      <div className={cn("p-2 rounded-xl bg-slate-50 dark:bg-slate-800", s.c)}>
                        <s.i className={cn("w-4 h-4", s.l === t.pending && isBroadcasting && "animate-spin")} />
                      </div>
                      <div className="min-w-0">
                        <p className="text-[9px] font-black text-slate-400 dark:text-slate-300 truncate uppercase">{s.l}</p>
                        <p className={cn("text-base font-black truncate", s.c)}>{s.v}</p>
                      </div>
                    </motion.div>
                  ))}
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 pb-20">
                  <div className="lg:col-span-4 space-y-6 order-2 lg:order-1">
                    <div className={cn("p-5 rounded-3xl border space-y-4", theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
                      <div className="flex items-center gap-2 text-slate-900 dark:text-white font-black text-[11px] uppercase tracking-widest">
                        <Type className="w-4 h-4 text-indigo-500" /> {t.title}
                      </div>
                      <input
                        type="text"
                        value={broadcastTitle}
                        onChange={(e) => setBroadcastTitle(e.target.value)}
                        placeholder="Enter broadcast title..."
                        className={cn(
                          "w-full px-4 py-3 border rounded-2xl text-xs font-bold outline-none",
                          theme === "dark"
                            ? "bg-slate-800 border-slate-700 text-white placeholder-slate-300"
                            : "bg-slate-50 border-slate-100 text-slate-700 placeholder-slate-400"
                        )}
                      />
                    </div>

                    <div className={cn("p-5 rounded-3xl border space-y-4", theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
                      <div className="flex items-center justify-between font-black text-[11px] uppercase tracking-widest text-slate-800 dark:text-white">
                        <div className="flex items-center gap-2">
                          <Upload className="w-4 h-4 text-indigo-500" /> Excel
                        </div>
                        <button
                          onClick={() => {
                            const ws = XLSX.utils.json_to_sheet([
                              { "Nomor Telepon": "08123...", Nama: "Budi" },
                            ]);
                            const wb = XLSX.utils.book_new();
                            XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
                            XLSX.writeFile(wb, "template.xlsx");
                          }}
                          className="text-[9px] bg-indigo-100 dark:bg-indigo-900/20 px-2.5 py-1.5 rounded-lg text-indigo-700 dark:text-indigo-300 font-black border dark:border-indigo-800"
                        >
                          TEMPLATE
                        </button>
                      </div>
                      <div className="relative border-2 border-dashed border-slate-100 dark:border-slate-800 rounded-3xl p-6 text-center hover:border-indigo-500 transition-colors">
                        <input
                          type="file"
                          accept=".xlsx, .xls"
                          onChange={handleExcelUpload}
                          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                        />
                        <FileSpreadsheet className="w-10 h-10 text-indigo-300 dark:text-indigo-700 mx-auto mb-2" />
                        <p className="text-[10px] font-black text-slate-400 dark:text-slate-300 uppercase tracking-widest">
                          {t.excel_upload}
                        </p>
                      </div>
                    </div>

                    <div className={cn("p-5 rounded-3xl border space-y-4", theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
                      <div className="flex items-center gap-2 text-slate-900 dark:text-white font-black text-[11px] uppercase tracking-widest">
                        <User className="w-4 h-4 text-indigo-500" /> {t.manual_add}
                      </div>
                      <form
                        onSubmit={(e) => {
                          e.preventDefault();
                          if (!broadcastTitle) {
                            showNotification("warning", "Judul Diperlukan", "Isi Judul!");
                            return;
                          }
                          const list = manualPhone
                            .split(/[,\n]/)
                            .map((p) => p.trim())
                            .filter((p) => p !== "");
                          const newRows = list.map((p, i) => ({
                            "Nomor Telepon": p,
                            Nama: manualName || "Recipient",
                            Pesan: manualMessage,
                            Status: "Pending",
                            __id: Date.now() + i,
                          }));
                          setData((prev) => [...prev, ...newRows]);
                          if (autoCheck) bulkCheckWhatsApp(newRows);
                          setManualPhone("");
                          setManualName("");
                          setManualMessage("");
                        }}
                        className="space-y-3"
                      >
                        <textarea
                          value={manualPhone}
                          onChange={(e) => setManualPhone(e.target.value)}
                          placeholder="0812, 0856..."
                          className={cn(
                            "w-full p-4 border rounded-2xl text-xs font-bold h-20 resize-none outline-none",
                            theme === "dark"
                              ? "bg-slate-800 border-slate-700 text-white placeholder-slate-300"
                              : "bg-slate-50 border-slate-100 text-slate-700 placeholder-slate-400"
                          )}
                        />
                        <input
                          type="text"
                          value={manualName}
                          onChange={(e) => setManualName(e.target.value)}
                          placeholder="Name"
                          className={cn(
                            "w-full p-4 border rounded-2xl text-xs font-bold outline-none",
                            theme === "dark"
                              ? "bg-slate-800 border-slate-700 text-white placeholder-slate-300"
                              : "bg-slate-50 border-slate-100 text-slate-700 placeholder-slate-400"
                          )}
                        />
                        <button type="submit" className="w-full py-4 bg-indigo-600 text-white rounded-2xl font-black text-[10px] uppercase tracking-widest">
                          ADD TO QUEUE
                        </button>
                      </form>
                    </div>
                  </div>

                  <div className="lg:col-span-8 space-y-6 order-1 lg:order-2">
                    <div className={cn("rounded-[40px] border flex flex-col h-[550px] overflow-hidden shadow-xl", theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
                      <div className="p-4 bg-slate-50/50 dark:bg-slate-800/50 border-b dark:border-slate-800 flex flex-wrap gap-2 justify-between items-center">
                        <div className="flex gap-2">
                          <button
                            onClick={() => setIsMessageModalOpen(true)}
                            className="px-3 py-2 bg-cyan-50 dark:bg-cyan-900/30 text-cyan-800 dark:text-cyan-200 text-[10px] font-black uppercase rounded-xl border border-cyan-200 dark:border-cyan-800 flex items-center gap-2"
                          >
                            <MessageSquare className="w-3.5 h-3.5" /> {t.message}
                          </button>
                          <button
                            onClick={() => setIsAttachmentModalOpen(true)}
                            className="px-3 py-2 bg-cyan-50 dark:bg-cyan-900/30 text-cyan-800 dark:text-cyan-200 text-[10px] font-black uppercase rounded-xl border border-cyan-200 dark:border-cyan-800 flex items-center gap-2"
                          >
                            <Paperclip className="w-3.5 h-3.5" /> {t.attachment}
                          </button>
                          <div className="flex items-center gap-2 bg-slate-100 dark:bg-slate-800 p-1.5 rounded-xl border dark:border-slate-700 pr-4">
                            <div className="p-1 bg-white dark:bg-slate-900 rounded-lg text-emerald-600">
                              <Clock className="w-3 h-3" />
                            </div>
                            <select
                              value={delayRange}
                              onChange={(e) => setDelayRange(e.target.value)}
                              className="bg-transparent text-slate-700 dark:text-white text-[10px] font-black uppercase outline-none cursor-pointer appearance-none"
                            >
                              {["25-40", "25-50", "25-60"].map((r) => (
                                <option key={r} value={r}>
                                  {r}s
                                </option>
                              ))}
                            </select>
                            <ChevronDown className="w-3 h-3 text-slate-400 dark:text-slate-300" />
                          </div>
                        </div>
                        <div className="flex items-center gap-2">
                          {invalidWaCount > 0 && (
                            <button
                              onClick={() => {
                                setData((prev) => prev.filter((row) => row.Status !== "No WA"));
                                showNotification(
                                  "info",
                                  lang === "id" ? "Nomor Tidak Valid Dihapus" : "Invalid Numbers Removed",
                                  lang === "id"
                                    ? `${invalidWaCount} nomor tidak terdaftar telah dihapus`
                                    : `${invalidWaCount} unregistered numbers removed`
                                );
                              }}
                              className="px-3 py-2 rounded-xl text-[10px] font-black uppercase tracking-wider bg-rose-100 text-rose-700 border border-rose-200 dark:bg-rose-900/30 dark:text-rose-200 dark:border-rose-800"
                            >
                              {lang === "id" ? "Hapus Nomor Tidak Valid" : "Remove Invalid Numbers"}
                            </button>
                          )}
                          {invalidWaCount === 0 && allWaActive && (
                            <span className="ui-chip ui-chip-success">
                              {lang === "id" ? "Siap" : "Ready"}
                            </span>
                          )}
                          <button
                            onClick={() =>
                              openConfirmModal({
                                title: lang === "id" ? "Konfirmasi Clear" : "Confirm Clear",
                                message:
                                  lang === "id" ? "Hapus semua data antrean?" : "Clear all queued data?",
                                confirmLabel: lang === "id" ? "HAPUS" : "CLEAR",
                                cancelLabel: lang === "id" ? "BATAL" : "CANCEL",
                                onConfirm: () => setData([]),
                              })
                            }
                            className="text-slate-400 dark:text-slate-300 hover:text-rose-500 p-2"
                          >
                            <Trash2 className="w-5 h-5" />
                          </button>
                        </div>
                      </div>
                      <div className="flex-1 overflow-auto px-2 lg:px-4">
                        <table className="w-full text-left border-collapse min-w-[700px]">
                          <thead className="bg-slate-50/50 dark:bg-slate-800 sticky top-0 text-[10px] uppercase font-black text-slate-400 dark:text-slate-300 border-b dark:border-slate-800 z-10 tracking-[0.2em]">
                            <tr>
                              <th className="pl-6 py-5 w-10">#</th>
                              <th className="px-4 py-4">Recipient</th>
                              <th className="px-4 py-4">Preview</th>
                              <th className="px-4 py-4 w-24 text-center">Status</th>
                              <th className="px-4 py-4 w-24 text-center">Action</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-50 dark:divide-slate-800 text-xs font-bold">
                            {data.map((row, idx) => (
                              <tr
                                key={row.__id ?? idx}
                                className={cn(
                                  "hover:bg-slate-50/50 dark:hover:bg-slate-800/50 transition-all",
                                  currentIndex === idx && "bg-amber-50 dark:bg-amber-900/10 animate-pulse"
                                )}
                              >
                                <td className="pl-6 py-5 text-slate-400 dark:text-slate-500 font-black">{idx + 1}</td>
                                <td className="px-4 py-4">
                                  <div className="text-slate-900 dark:text-slate-100">{row.Nama}</div>
                                  <div className="text-[10px] text-slate-500 dark:text-slate-300 mt-1 font-mono tracking-tighter">
                                    {row["Nomor Telepon"]}
                                  </div>
                                </td>
                                <td className="px-4 py-4 text-slate-600 dark:text-slate-200 italic truncate max-w-[200px] leading-relaxed">
                                  {row.Lampiran && (
                                    <div className="inline-flex items-center gap-1 text-[8px] bg-sky-50 dark:bg-sky-900/20 text-sky-600 px-1.5 py-0.5 rounded uppercase font-black">
                                      <Paperclip className="w-2.5 h-2.5" /> Media
                                    </div>
                                  )}
                                  <div>{row.Pesan.replace(/{nama}/gi, row.Nama) || "-"}</div>
                                </td>
                                <td className="px-4 py-4 text-center">
                                  <span
                                    className={cn(
                                      "ui-chip",
                                      row.Status === "Success"
                                        ? "ui-chip-success"
                                        : row.Status === "Failed"
                                          ? "ui-chip-danger"
                                          : "ui-chip-neutral"
                                    )}
                                  >
                                    {currentIndex === idx ? <Loader2 className="w-3 h-3 animate-spin" /> : null}
                                    {currentIndex === idx ? "Sending" : row.Status || "Pending"}
                                  </span>
                                </td>
                                <td className="px-4 py-4 text-center">
                                  <button
                                    onClick={() =>
                                      setData((prev) => prev.filter((r) => r.__id !== row.__id))
                                    }
                                    className="text-rose-400 p-2 hover:bg-rose-50 rounded-xl"
                                  >
                                    <Trash2 className="w-4 h-4" />
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <div className="p-6 bg-gradient-to-r from-indigo-600/10 to-violet-600/10 border-t dark:border-slate-700">
                        <button
                          onClick={startBroadcast}
                          disabled={data.length === 0}
                          className={cn(
                            "w-full py-4 rounded-2xl font-black text-white transition-all flex items-center justify-center gap-3 tracking-[0.2em] uppercase text-sm",
                            isBroadcasting
                              ? "bg-rose-500 hover:bg-rose-600"
                              : "bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-300 disabled:text-slate-500"
                          )}
                        >
                          {isBroadcasting ? "STOP ENGINE" : t.start_now}
                        </button>
                      </div>
                    </div>

                    <section className={cn("rounded-3xl border p-4", theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
                      <div className="flex items-center justify-between gap-3 mb-3">
                        <h4 className="text-[11px] font-black uppercase tracking-widest text-slate-600 dark:text-slate-300">
                          Broadcast Logs
                        </h4>
                        <div className="flex items-center gap-2">
                          <span className="text-[10px] font-semibold text-slate-500 dark:text-slate-300">
                            {logs.length} entries
                          </span>
                          <button
                            onClick={() => {
                              if (logs.length === 0) return;
                              openConfirmModal({
                                title: lang === "id" ? "Bersihkan Log" : "Clear Logs",
                                message:
                                  lang === "id"
                                    ? "Hapus semua broadcast logs?"
                                    : "Delete all broadcast logs?",
                                confirmLabel: lang === "id" ? "HAPUS" : "DELETE",
                                cancelLabel: lang === "id" ? "BATAL" : "CANCEL",
                                onConfirm: () => setLogs([]),
                              });
                            }}
                            disabled={logs.length === 0}
                            className="px-2.5 py-1 rounded-lg text-[10px] font-black uppercase tracking-wider border border-rose-300 text-rose-700 disabled:opacity-50 disabled:cursor-not-allowed"
                          >
                            {lang === "id" ? "Bersihkan" : "Clear"}
                          </button>
                        </div>
                      </div>
                      <div className="max-h-44 overflow-auto space-y-2 pr-1">
                        {logs.length === 0 && (
                          <p className="text-xs text-slate-500 dark:text-slate-300">Belum ada log broadcast.</p>
                        )}
                        {logs.slice(0, 12).map((log, idx) => (
                          <div key={`${log.timestamp}-${idx}`} className="text-xs flex items-center gap-2">
                            <span className="text-slate-400 dark:text-slate-500 w-20 shrink-0">{log.timestamp}</span>
                            <span
                              className={cn(
                                "w-2 h-2 rounded-full shrink-0",
                                log.status === "success"
                                  ? "bg-emerald-500"
                                  : log.status === "error"
                                    ? "bg-rose-500"
                                    : "bg-indigo-500"
                              )}
                            />
                            <span className="font-semibold truncate max-w-32">{log.target}</span>
                            <span className="text-slate-500 dark:text-slate-300 truncate">{log.message}</span>
                          </div>
                        ))}
                      </div>
                    </section>
                  </div>
                </div>
              </motion.div>
            )}

            {activeView === "configuration" && (
              <motion.div className="max-w-2xl mx-auto space-y-6 pb-20" {...viewMotion}>
                <section
                  className={cn(
                    "p-6 lg:p-10 rounded-[40px] border shadow-sm relative overflow-hidden",
                    theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200"
                  )}
                >
                  <div className="space-y-10">
                    <div className="flex items-center gap-4">
                      <div className="bg-slate-900 dark:bg-slate-800 p-4 rounded-3xl text-white shadow-lg">
                        <Settings className="w-7 h-7" />
                      </div>
                      <h3 className="text-2xl font-black uppercase tracking-tight">{t.system_config}</h3>
                    </div>

                    <div className="space-y-4">
                      <label className="text-[10px] font-black text-slate-400 dark:text-slate-300 uppercase tracking-widest px-2">
                        {t.api_key}
                      </label>
                      <div className="flex flex-col sm:flex-row gap-4">
                        <input
                          type="password"
                          value={apiKey}
                          onChange={(e) => {
                            setApiKey(e.target.value);
                            setIsTokenChanged(true);
                          }}
                          placeholder="Paste Fonnte Key..."
                          className={cn(
                            "flex-1 px-6 py-4 border rounded-2xl outline-none font-bold text-sm",
                            theme === "dark"
                              ? "bg-slate-800 border-slate-700 text-white placeholder-slate-300"
                              : "bg-slate-50 border-slate-100 text-slate-700 placeholder-slate-400"
                          )}
                        />
                        <div className="flex gap-2">
                          {deviceInfo ? (
                            <>
                              <div className="px-6 py-4 bg-indigo-600 text-white rounded-2xl font-black text-[10px] tracking-widest flex items-center justify-center gap-2">
                                <CheckCircle2 className="w-4 h-4" /> {t.connected}
                              </div>
                              <button
                                onClick={() => {
                                  setDeviceInfo(null);
                                  setApiKey("");
                                  setIsTokenChanged(false);
                                }}
                                className="px-6 py-4 bg-rose-500 text-white rounded-2xl hover:bg-rose-600 font-black text-[10px] uppercase"
                              >
                                {t.disconnect}
                              </button>
                            </>
                          ) : (
                            <button
                              onClick={() => checkConnection(false)}
                              className={cn(
                                "px-8 py-4 rounded-2xl font-black text-[10px] tracking-[0.2em]",
                                isTokenChanged ? "bg-amber-400 text-amber-950" : "bg-rose-500 text-white"
                              )}
                            >
                              {isCheckingConnection ? (
                                <Loader2 className="w-5 h-5 animate-spin mx-auto" />
                              ) : isTokenChanged ? (
                                t.connect
                              ) : (
                                t.not_connected
                              )}
                            </button>
                          )}
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                      <div className="space-y-4">
                        <label className="text-[10px] font-black text-slate-400 dark:text-slate-300 uppercase tracking-widest px-2">
                          {t.theme}
                        </label>
                        <div className={cn("grid grid-cols-2 p-2 rounded-2xl border", theme === "dark" ? "bg-slate-800 border-slate-700" : "bg-slate-100 border-slate-200")}>
                          <button
                            onClick={() => setTheme("light")}
                            className={cn(
                              "flex items-center justify-center gap-2 py-3 rounded-xl text-[10px] font-black uppercase",
                              theme === "light" ? "bg-white text-emerald-600" : "text-slate-500"
                            )}
                          >
                            <Sun className="w-4 h-4" /> {t.light}
                          </button>
                          <button
                            onClick={() => setTheme("dark")}
                            className={cn(
                              "flex items-center justify-center gap-2 py-3 rounded-xl text-[10px] font-black uppercase",
                              theme === "dark" ? "bg-slate-900 text-emerald-400" : "text-slate-500"
                            )}
                          >
                            <Moon className="w-4 h-4" /> {t.dark}
                          </button>
                        </div>
                      </div>
                      <div className="space-y-4">
                        <label className="text-[10px] font-black text-slate-400 dark:text-slate-300 uppercase tracking-widest px-2">
                          {t.language}
                        </label>
                        <div className={cn("grid grid-cols-2 p-2 rounded-2xl border", theme === "dark" ? "bg-slate-800 border-slate-700" : "bg-slate-100 border-slate-200")}>
                          <button
                            onClick={() => setLang("id")}
                            className={cn(
                              "py-3 rounded-xl text-[10px] font-black uppercase tracking-widest",
                              lang === "id" ? "bg-indigo-600 text-white" : "text-slate-500"
                            )}
                          >
                            ID
                          </button>
                          <button
                            onClick={() => setLang("en")}
                            className={cn(
                              "py-3 rounded-xl text-[10px] font-black uppercase tracking-widest",
                              lang === "en" ? "bg-indigo-600 text-white" : "text-slate-500"
                            )}
                          >
                            EN
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                </section>
              </motion.div>
            )}

            {activeView === "history" && (
              <motion.div className="space-y-6 pb-20" {...viewMotion}>
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <div className="flex items-center gap-4">
                    <div className="bg-indigo-100 dark:bg-indigo-900/30 p-4 rounded-3xl text-indigo-600 shrink-0">
                      <HistoryIcon className="w-7 h-7" />
                    </div>
                    <h3 className="text-2xl font-black uppercase leading-none tracking-tight">{t.history}</h3>
                  </div>
                  <div className="flex items-center gap-3">
                    <button
                      onClick={fetchHistory}
                      disabled={isFetchingHistory}
                      className="px-6 py-3 bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-2xl text-[10px] font-black text-slate-500 dark:text-slate-300 flex items-center justify-center gap-3 uppercase tracking-widest"
                    >
                      {isFetchingHistory ? (
                        <Loader2 className="w-4 h-4 animate-spin text-indigo-600" />
                      ) : (
                        <RefreshCw className="w-4 h-4" />
                      )}
                      SYNC DATA
                    </button>
                    <div className="relative">
                      <button
                        onClick={() => setIsHistoryMenuOpen(!isHistoryMenuOpen)}
                        className="px-4 py-3 bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-2xl text-slate-500 dark:text-slate-300"
                      >
                        <MoreVertical className="w-5 h-5" />
                      </button>
                      {isHistoryMenuOpen && (
                        <div className="absolute right-0 mt-2 w-48 bg-white dark:bg-slate-800 border dark:border-slate-700 rounded-2xl shadow-xl z-50 overflow-hidden">
                          <button
                            onClick={() => {
                              exportHistoryData();
                              setIsHistoryMenuOpen(false);
                            }}
                            className="w-full px-6 py-4 text-left flex items-center gap-3 hover:bg-indigo-50 dark:hover:bg-indigo-900/30 border-b dark:border-slate-700 text-[10px] font-black uppercase tracking-widest text-slate-700 dark:text-slate-200"
                          >
                            <Download className="w-4 h-4 text-indigo-600" /> {t.export}
                          </button>
                          <button
                            onClick={() => {
                              deleteHistoryData();
                              setIsHistoryMenuOpen(false);
                            }}
                            className="w-full px-6 py-4 text-left flex items-center gap-3 hover:bg-rose-50 dark:hover:bg-rose-900/30 text-[10px] font-black uppercase tracking-widest text-rose-600"
                          >
                            <Trash2 className="w-4 h-4" /> {t.delete_data}
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
                <div className={cn("rounded-[40px] border overflow-hidden min-h-[550px]", theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-200")}>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left min-w-[800px] border-collapse">
                      <thead className="bg-slate-50/50 dark:bg-slate-800/80 sticky top-0 text-[10px] uppercase font-black text-slate-400 dark:text-slate-300 border-b dark:border-slate-800 z-10 tracking-[0.2em]">
                        <tr>
                          <th className="pl-10 py-6 w-14">#</th>
                          <th className="px-6 py-6">Title</th>
                          <th className="px-6 py-6">Recipient</th>
                          <th className="px-6 py-6 text-center">Status</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-50 dark:divide-slate-800 text-[11px] font-bold">
                        {historyRows.map((row, i) => (
                          <tr key={`${row["nomor Tujuan"]}-${i}`} className="hover:bg-slate-50/30 dark:hover:bg-slate-800/30 transition-all duration-300">
                            <td className="pl-10 py-6 text-slate-400 font-black">{i + 1}</td>
                            <td className="px-6 py-6">
                              <span className="ui-chip ui-chip-info">
                                {row["Judul broadcast"]}
                              </span>
                            </td>
                            <td className="px-6 py-6 font-mono tracking-tighter">{row["nomor Tujuan"]}</td>
                            <td className="px-6 py-6 text-center">
                              <span
                                className={cn(
                                  "ui-chip",
                                  row.status?.toLowerCase().includes("success")
                                    ? "ui-chip-success"
                                    : "ui-chip-danger"
                                )}
                              >
                                {row.status === "Success" ? (
                                  <CheckCircle2 className="w-3.5 h-3.5" />
                                ) : (
                                  <XCircle className="w-3.5 h-3.5" />
                                )}
                                {row.status}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {historyRows.length === 0 && (
                    <div className="h-[400px] flex flex-col items-center justify-center gap-4 text-slate-400 dark:text-slate-300">
                      <Database className="w-16 h-16 opacity-20" />
                      <p className="font-black uppercase tracking-[0.3em] text-[10px]">No records detected</p>
                    </div>
                  )}
                </div>
              </motion.div>
            )}
          </div>
        </main>
      </div>

      {isMessageModalOpen && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-md">
          <div className={cn("rounded-[40px] shadow-2xl w-full max-w-lg overflow-hidden border", theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-100")}>
            <div className="p-8 border-b dark:border-slate-800 flex justify-between items-center bg-slate-50/50 dark:bg-slate-800/50">
              <h3 className="font-black text-slate-800 dark:text-white uppercase tracking-tight text-xl flex items-center gap-3">
                <MessageSquare className="w-6 h-6 text-indigo-600" /> Message Editor
              </h3>
              <button onClick={() => setIsMessageModalOpen(false)} className="text-slate-400 dark:text-slate-300 hover:text-rose-500 p-2">
                <X className="w-7 h-7" />
              </button>
            </div>
            <div className="p-8 space-y-6">
              <div className="bg-indigo-50 dark:bg-indigo-900/10 p-5 rounded-3xl border border-indigo-100 dark:border-indigo-800 text-[11px] text-indigo-800 dark:text-indigo-400 font-bold leading-relaxed">
                Pesan ini akan diterapkan ke seluruh antrean. Gunakan tag <code className="bg-white dark:bg-slate-800 px-2 py-0.5 rounded text-indigo-600">{"{nama}"}</code>.
              </div>
              <textarea
                value={bulkMessageInput}
                onChange={(e) => setBulkMessageInput(e.target.value)}
                placeholder="Tulis pesan massal Anda di sini..."
                className={cn(
                  "w-full h-48 p-6 border rounded-[24px] text-sm outline-none resize-none font-bold",
                  theme === "dark"
                    ? "bg-slate-800 border-slate-700 text-white placeholder-slate-300"
                    : "bg-slate-50 border-slate-100 text-slate-700 placeholder-slate-400"
                )}
              />
              <div className="flex gap-4">
                <button onClick={() => setIsMessageModalOpen(false)} className="flex-1 py-4 border rounded-2xl font-black text-slate-400 dark:text-slate-300 uppercase tracking-widest text-[10px]">
                  {t.cancel}
                </button>
                <button
                  onClick={() => {
                    if (!bulkMessageInput) return;
                    setData((prev) => prev.map((r) => ({ ...r, Pesan: bulkMessageInput })));
                    setIsMessageModalOpen(false);
                    addLog("System", "info", "Message Applied");
                  }}
                  className="flex-1 py-4 bg-indigo-600 text-white rounded-2xl font-black uppercase tracking-widest text-[10px]"
                >
                  {t.apply}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {isAttachmentModalOpen && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-md">
          <div className={cn("rounded-[40px] shadow-2xl w-full max-w-lg overflow-hidden border", theme === "dark" ? "bg-slate-900 border-slate-800" : "bg-white border-slate-100")}>
            <div className="p-8 border-b dark:border-slate-800 flex justify-between items-center bg-slate-50/50 dark:bg-slate-800/50">
              <h3 className="font-black text-slate-800 dark:text-white uppercase tracking-tight text-xl flex items-center gap-3">
                <Paperclip className="w-6 h-6 text-indigo-600" /> {t.attachment}
              </h3>
              <button onClick={() => setIsAttachmentModalOpen(false)} className="text-slate-400 dark:text-slate-300 hover:text-rose-500 p-2">
                <X className="w-7 h-7" />
              </button>
            </div>
            <div className="p-8 space-y-8">
              <div className="bg-amber-50 dark:bg-amber-900/10 p-6 rounded-3xl border border-amber-100 dark:border-amber-800 flex gap-4 leading-relaxed">
                <AlertTriangle className="w-6 h-6 text-amber-600 shrink-0" />
                <p className="text-xs text-amber-800 dark:text-amber-400 font-bold leading-relaxed">{t.package_warning}</p>
              </div>

              {!hasPaidPackage ? (
                <div className="space-y-6 text-center">
                  <div className="bg-slate-50 dark:bg-slate-800 p-8 rounded-[24px] border dark:border-slate-700">
                    <p className="text-slate-400 dark:text-slate-300 text-[10px] font-black uppercase tracking-widest mb-2">
                      DETEKSI PAKET: {deviceInfo?.package || "Unknown"}
                    </p>
                    <p className="text-slate-600 dark:text-slate-300 font-black text-xs uppercase leading-relaxed">
                      Fitur media memerlukan paket berbayar
                    </p>
                  </div>
                  <a
                    href="https://fonnte.com/#harga"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="w-full py-5 bg-indigo-600 text-white rounded-[24px] font-black text-[10px] tracking-widest uppercase flex items-center justify-center gap-3"
                  >
                    <Zap className="w-5 h-5" /> {t.subscribe} <ExternalLink className="w-4 h-4" />
                  </a>
                </div>
              ) : (
                <div className="space-y-6">
                  <div className="space-y-3">
                    <label className="text-[10px] font-black text-slate-400 dark:text-slate-300 uppercase tracking-widest px-2">
                      Attachment URL (Public Image/File)
                    </label>
                    <input
                      type="text"
                      value={attachmentUrl}
                      onChange={(e) => setAttachmentUrl(e.target.value)}
                      placeholder="https://example.com/image.jpg"
                      className={cn(
                        "w-full px-6 py-4 border rounded-3xl outline-none font-bold text-sm",
                        theme === "dark"
                          ? "bg-slate-800 border-slate-700 text-white placeholder-slate-300"
                          : "bg-slate-50 border-slate-100 text-slate-700 placeholder-slate-400"
                      )}
                    />
                  </div>
                  <button
                    onClick={() => {
                      setData((prev) => prev.map((r) => ({ ...r, Lampiran: attachmentUrl })));
                      setIsAttachmentModalOpen(false);
                      addLog("System", "success", "Media Attachment Applied");
                    }}
                    className="w-full py-5 bg-indigo-600 text-white rounded-[24px] font-black text-[10px] tracking-widest uppercase"
                  >
                    APPLY MEDIA URL
                  </button>
                </div>
              )}
              <button onClick={() => setIsAttachmentModalOpen(false)} className="w-full text-center text-[9px] font-black text-slate-400 dark:text-slate-300 uppercase tracking-[0.2em]">
                {t.cancel}
              </button>
            </div>
          </div>
        </div>
      )}

      {isSelectBroadcastModalOpen && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-black/50 backdrop-blur-xl">
          <div className={cn("rounded-[32px] shadow-2xl w-full max-w-2xl overflow-hidden border max-h-[85vh] flex flex-col", theme === "dark" ? "bg-slate-900 border-slate-700" : "bg-white border-slate-200")}>
            <div className={cn("p-6 border-b flex justify-between items-start gap-4", theme === "dark" ? "border-slate-700" : "border-slate-200")}>
              <div className="space-y-2 flex-1">
                <div className="flex items-center gap-3">
                  <div className={cn("p-2.5 rounded-2xl", selectBroadcastAction === "export" ? "bg-emerald-100 dark:bg-emerald-900/30" : "bg-rose-100 dark:bg-rose-900/30")}>
                    {selectBroadcastAction === "export" ? (
                      <Download className="w-5 h-5 text-emerald-600" />
                    ) : (
                      <Trash2 className="w-5 h-5 text-rose-600" />
                    )}
                  </div>
                  <h3 className={cn("font-black text-lg uppercase tracking-tight", theme === "dark" ? "text-white" : "text-slate-900")}>
                    {selectBroadcastAction === "export"
                      ? lang === "id"
                        ? "Pilih Data untuk Export"
                        : "Select Data to Export"
                      : lang === "id"
                        ? "Pilih Data untuk Dihapus"
                        : "Select Data to Delete"}
                  </h3>
                </div>
                <p className={cn("text-[11px] font-semibold uppercase tracking-widest", theme === "dark" ? "text-slate-300" : "text-slate-500")}>
                  {lang === "id"
                    ? `${getCountForBroadcast(null)} data tersimpan`
                    : `${getCountForBroadcast(null)} records stored`}
                </p>
              </div>
              <button
                onClick={() => {
                  setIsSelectBroadcastModalOpen(false);
                  setSearchBroadcast("");
                }}
                className="p-2 rounded-xl"
              >
                <X className={cn("w-6 h-6", theme === "dark" ? "text-slate-300" : "text-slate-500")} />
              </button>
            </div>

            <div className={cn("p-6 border-b space-y-4", theme === "dark" ? "border-slate-700 bg-slate-800/20" : "border-slate-200 bg-slate-50/50")}>
              <div className="relative">
                <Search className={cn("absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5", theme === "dark" ? "text-slate-300" : "text-slate-400")} />
                <input
                  type="text"
                  value={searchBroadcast}
                  onChange={(e) => setSearchBroadcast(e.target.value)}
                  placeholder={lang === "id" ? "Cari nama broadcast..." : "Search broadcast name..."}
                  className={cn(
                    "w-full pl-12 pr-4 py-3 rounded-xl border-2 outline-none transition-all text-sm font-semibold",
                    theme === "dark"
                      ? "bg-slate-800 border-slate-700 text-white placeholder-slate-300"
                      : "bg-white border-slate-200 text-slate-900 placeholder-slate-400"
                  )}
                />
              </div>
              {searchBroadcast && (
                <p className={cn("text-[10px] font-black uppercase tracking-widest", theme === "dark" ? "text-slate-300" : "text-slate-600")}>
                  {lang === "id"
                    ? `Ditemukan ${getBroadcastList().length} broadcast`
                    : `Found ${getBroadcastList().length} broadcasts`}
                </p>
              )}
            </div>

            <div className={cn("flex-1 overflow-y-auto p-6 space-y-3", theme === "dark" ? "bg-slate-800/40" : "bg-slate-50/50")}>
              <button
                onClick={() => {
                  if (selectBroadcastAction === "export") executeExport(null);
                  else executeDelete(null);
                  setIsSelectBroadcastModalOpen(false);
                  setSearchBroadcast("");
                  setSelectBroadcastAction(null);
                }}
                className={cn(
                  "w-full p-4 rounded-2xl border-2 transition-all duration-200 active:scale-95 text-left",
                  theme === "dark"
                    ? "border-emerald-700/50 bg-emerald-900/20"
                    : "border-emerald-300 bg-emerald-50"
                )}
              >
                <div className="flex items-center gap-3">
                  <div className={cn("w-5 h-5 rounded-lg border-2 flex items-center justify-center", selectBroadcastAction === "export" ? "border-indigo-600 bg-indigo-600" : "border-rose-500 bg-rose-500")}>
                    <CheckCircle2 className="w-3.5 h-3.5 text-white" />
                  </div>
                  <div className="flex-1 text-left">
                    <p className={cn("font-black text-sm uppercase tracking-wide", theme === "dark" ? "text-emerald-400" : "text-emerald-700")}>
                      All Broadcasts
                    </p>
                    <p className={cn("text-[10px] font-semibold", theme === "dark" ? "text-slate-300" : "text-slate-600")}>
                      {getCountForBroadcast(null)} {lang === "id" ? "data" : "records"}
                    </p>
                  </div>
                </div>
              </button>

              <div className={cn("h-px my-2", theme === "dark" ? "bg-slate-700" : "bg-slate-300")} />

              {getBroadcastList().length > 0 ? (
                <div className="grid grid-cols-1 gap-2.5">
                  {getBroadcastList().map((title, idx) => (
                    <button
                      key={`${title}-${idx}`}
                      onClick={() => {
                        if (selectBroadcastAction === "export") executeExport(title);
                        else executeDelete(title);
                        setIsSelectBroadcastModalOpen(false);
                        setSearchBroadcast("");
                        setSelectBroadcastAction(null);
                      }}
                      className={cn(
                        "w-full p-4 rounded-2xl border-2 transition-all duration-200 active:scale-95 text-left",
                        theme === "dark"
                          ? "border-slate-700 bg-slate-800/50 hover:bg-slate-700/50"
                          : "border-slate-200 bg-slate-100/50 hover:bg-slate-200/50"
                      )}
                    >
                      <div className="flex items-center gap-3">
                        <div className={cn("w-5 h-5 rounded-lg border-2 flex items-center justify-center", theme === "dark" ? "border-slate-600 bg-slate-700" : "border-slate-300 bg-slate-200")}>
                          <span className={cn("text-[10px] font-black", theme === "dark" ? "text-slate-300" : "text-slate-600")}>OK</span>
                        </div>
                        <div className="flex-1 min-w-0">
                          <p className={cn("font-black text-sm uppercase tracking-wide truncate", theme === "dark" ? "text-white" : "text-slate-800")}>
                            {title}
                          </p>
                          <p className={cn("text-[10px] font-semibold", theme === "dark" ? "text-slate-200" : "text-slate-600")}>
                            {getCountForBroadcast(title)} {lang === "id" ? "data" : "records"}
                          </p>
                        </div>
                      </div>
                    </button>
                  ))}
                </div>
              ) : (
                <div className="h-32 flex flex-col items-center justify-center gap-2 text-center">
                  <Database className={cn("w-10 h-10 opacity-20", theme === "dark" ? "text-slate-200" : "text-slate-400")} />
                  <p className={cn("text-xs font-black uppercase tracking-wide", theme === "dark" ? "text-slate-300" : "text-slate-600")}>
                    {lang === "id" ? "Tidak ada broadcast ditemukan" : "No broadcasts found"}
                  </p>
                </div>
              )}
            </div>

            <div className={cn("p-4 border-t", theme === "dark" ? "bg-slate-800/50 border-slate-700" : "bg-slate-50/50 border-slate-200")}>
              <button
                onClick={() => {
                  setIsSelectBroadcastModalOpen(false);
                  setSearchBroadcast("");
                }}
                className={cn(
                  "w-full py-3 px-4 rounded-xl font-black text-[10px] uppercase tracking-widest",
                  theme === "dark" ? "bg-slate-800 text-white" : "bg-slate-200 text-slate-700"
                )}
              >
                {t.cancel}
              </button>
            </div>
          </div>
        </div>
      )}

      {confirmModal && (
        <div className="fixed inset-0 z-[95] flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-sm">
          <div className={cn("w-full max-w-md rounded-[32px] border shadow-2xl overflow-hidden", theme === "dark" ? "bg-slate-900 border-slate-700" : "bg-white border-slate-200")}>
            <div className={cn("p-6 border-b", theme === "dark" ? "border-slate-700 bg-slate-950/90" : "border-slate-200 bg-slate-50")}>
              <div className="flex items-center justify-between gap-4">
                <div className="space-y-2">
                  <h3 className="text-lg font-black text-slate-900 dark:text-white">{confirmModal.title}</h3>
                  <p className="text-sm text-slate-500 dark:text-slate-300">{confirmModal.message}</p>
                </div>
                <button onClick={() => setConfirmModal(null)} className="p-2 rounded-full text-slate-400 hover:text-slate-700 dark:hover:text-slate-100">
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>
            <div className="p-6 bg-slate-50 dark:bg-slate-950/90 flex gap-3">
              <button
                onClick={() => setConfirmModal(null)}
                className={cn(
                  "flex-1 py-3 rounded-2xl border font-black text-sm uppercase tracking-[0.2em]",
                  theme === "dark"
                    ? "border-slate-700 bg-slate-900 text-white"
                    : "border-slate-200 bg-white text-slate-700"
                )}
              >
                {confirmModal.cancelLabel}
              </button>
              <button
                onClick={() => {
                  confirmModal.onConfirm();
                  setConfirmModal(null);
                }}
                className="flex-1 py-3 rounded-2xl bg-rose-500 text-white font-black text-sm uppercase tracking-[0.2em]"
              >
                {confirmModal.confirmLabel}
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="fixed bottom-4 right-4 z-[100] flex flex-col gap-3 items-end">
        {notifications.map((notification) => (
          <div
            key={notification.id}
            className={cn(
              "w-full max-w-sm rounded-3xl border p-4 shadow-2xl backdrop-blur-xl text-left overflow-hidden border-opacity-70",
              notification.variant === "success"
                ? "bg-emerald-600/95 border-emerald-300 text-white"
                : notification.variant === "error"
                  ? "bg-rose-600/95 border-rose-300 text-white"
                  : notification.variant === "warning"
                    ? "bg-amber-500/95 border-amber-300 text-slate-950"
                    : "bg-slate-950/95 border-slate-300 text-slate-100"
            )}
          >
            <div className="flex items-start gap-3">
              <div className="mt-0.5">
                {notification.variant === "success" && <CheckCircle2 className="w-5 h-5 text-white" />}
                {notification.variant === "error" && <XCircle className="w-5 h-5 text-white" />}
                {notification.variant === "warning" && <AlertTriangle className="w-5 h-5 text-slate-950" />}
                {notification.variant === "info" && <Info className="w-5 h-5 text-white" />}
              </div>
              <div className="min-w-0">
                <div className="font-black text-sm uppercase tracking-[0.2em]">{notification.title}</div>
                <p className="text-xs leading-relaxed mt-1 opacity-90">{notification.message}</p>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
