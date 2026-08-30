import React, { useState, useEffect, useMemo, useRef } from "react";
import { saveFileHandle, loadFileHandle, clearFileHandle } from "./fileHandleStore.js";
import * as XLSX from "xlsx";
import {
  Package, Plus, Pencil, Trash2, Search, Download, Receipt,
  ShoppingCart, X, Check, AlertTriangle, ChevronRight, Boxes,
  Printer, ArrowLeft, FileText, Link2, RefreshCw, Image as ImageIcon,
  DatabaseBackup, UploadCloud, Users, RotateCcw, Minus, Share2, FileSpreadsheet, EyeOff, Eye, Sun, Moon
} from "lucide-react";

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');`;

// Colors are CSS variables, not raw hex, so the whole app can switch
// between light/dark by swapping one attribute on the root — every
// existing `color: INK` / `background: PAPER` style below picks up
// the new value automatically, with no other changes needed.
const INK = "var(--sr-ink)";
const INK_SOFT = "var(--sr-ink-soft)";
const PAPER = "var(--sr-paper)";
const PAPER_DIM = "var(--sr-paper-dim)";
const AMBER = "var(--sr-amber)";
const AMBER_DARK = "var(--sr-amber-dark)";
const AMBER_TINT = "var(--sr-amber-tint)";
const SLATE = "var(--sr-slate)";
const GREEN = "var(--sr-green)";
const RED = "var(--sr-red)";
const LINE = "var(--sr-line)";
const SHADOW_SM = "var(--sr-shadow-sm)";
const SHADOW_MD = "var(--sr-shadow-md)";
const SHADOW_LG = "var(--sr-shadow-lg)";

const THEME_KEY = "stockroom:theme";

// The sidebar is a fixed dark-navy panel by design in both themes — it
// should NOT flip when the rest of the app switches to dark mode, so
// it uses its own constant colors instead of the themed variables above.
const SIDEBAR_BG = "#1C2431";
const SIDEBAR_BG_ACTIVE = "#2A3444";
const SIDEBAR_TEXT = "#FAF6EE";

const THEME_VARS = `
  :root, [data-sr-theme="light"] {
    --sr-ink: #1C2431;
    --sr-ink-soft: #2A3444;
    --sr-paper: #FAF6EE;
    --sr-paper-dim: #F1EBDD;
    --sr-amber: #C68A2E;
    --sr-amber-dark: #A8741F;
    --sr-amber-tint: #FBF0DC;
    --sr-slate: #5B6472;
    --sr-green: #3F7A57;
    --sr-red: #B5483D;
    --sr-line: #E4DCC8;
    --sr-shadow-sm: 0 1px 2px rgba(28,36,49,0.06), 0 1px 1px rgba(28,36,49,0.04);
    --sr-shadow-md: 0 6px 16px rgba(28,36,49,0.08), 0 2px 6px rgba(28,36,49,0.05);
    --sr-shadow-lg: 0 20px 48px rgba(28,36,49,0.16), 0 6px 18px rgba(28,36,49,0.08);
    --sr-card-bg: #ffffff;
    --sr-input-bg: #ffffff;
    color-scheme: light;
  }
  [data-sr-theme="dark"] {
    --sr-ink: #EDEFF4;
    --sr-ink-soft: #D6DAE3;
    --sr-paper: #171B24;
    --sr-paper-dim: #1E232E;
    --sr-amber: #D9A03F;
    --sr-amber-dark: #E8B863;
    --sr-amber-tint: #2E2515;
    --sr-slate: #97A0B0;
    --sr-green: #5FAE7C;
    --sr-red: #E0685A;
    --sr-line: #313847;
    --sr-shadow-sm: 0 1px 2px rgba(0,0,0,0.35), 0 1px 1px rgba(0,0,0,0.25);
    --sr-shadow-md: 0 6px 16px rgba(0,0,0,0.4), 0 2px 6px rgba(0,0,0,0.3);
    --sr-shadow-lg: 0 20px 48px rgba(0,0,0,0.55), 0 6px 18px rgba(0,0,0,0.35);
    --sr-card-bg: #212734;
    --sr-input-bg: #1B202A;
    color-scheme: dark;
  }
`;

function uid(prefix = "") {
  return prefix + Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-4);
}

function currency(n) {
  return "₹" + Number(n || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function todayStr() {
  return new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
}

const STOCK_KEY = "stockroom:items";
const SALES_KEY = "stockroom:invoices";
const SEQ_KEY = "stockroom:invoice-seq";
const RETURNS_KEY = "stockroom:returns";
const RETURN_SEQ_KEY = "stockroom:return-seq";
const AUTH_KEY = "stockroom:auth";
const SESSION_KEY = "stockroom:session";
const DEFAULT_AUTH = { username: "Admin", password: "Admin@123" };

// ---- lazy-load jsPDF from CDN, once ----
let jsPDFPromise = null;
function loadJsPDF() {
  if (window.jspdf) return Promise.resolve(window.jspdf);
  if (jsPDFPromise) return jsPDFPromise;
  jsPDFPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    script.onload = () => resolve(window.jspdf);
    script.onerror = () => reject(new Error("Could not load PDF library"));
    document.head.appendChild(script);
  });
  return jsPDFPromise;
}

function resizeImage(file, maxDim = 240) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          if (width > height) { height = Math.round((height * maxDim) / width); width = maxDim; }
          else { width = Math.round((width * maxDim) / height); height = maxDim; }
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", 0.82));
      };
      img.onerror = () => reject(new Error("Invalid image"));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error("Could not read file"));
    reader.readAsDataURL(file);
  });
}

function buildWorkbook(invoices, returns = []) {
  const rows = [];
  invoices.forEach((inv) => {
    inv.lines.forEach((l) => {
      rows.push({
        Invoice: inv.number,
        Date: inv.date,
        Customer: inv.customer,
        SKU: l.sku,
        Item: l.name,
        Quantity: l.qty,
        "Unit Price": l.price,
        "Line Total": +(l.qty * l.price).toFixed(2),
        "Invoice Total": inv.total,
      });
    });
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{ wch: 10 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 22 }, { wch: 9 }, { wch: 11 }, { wch: 11 }, { wch: 13 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sales Log");

  const returnRows = [];
  returns.forEach((ret) => {
    ret.lines.forEach((l) => {
      returnRows.push({
        "Return #": ret.number,
        "Original Invoice": ret.invoiceNumber,
        Date: ret.date,
        Customer: ret.customer,
        SKU: l.sku,
        Item: l.name,
        "Qty Returned": l.qty,
        "Unit Price": l.price,
        "Refund Amount": +(l.qty * l.price).toFixed(2),
        Reason: ret.reason || "",
      });
    });
  });
  const wsReturns = XLSX.utils.json_to_sheet(returnRows);
  wsReturns["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 22 }, { wch: 12 }, { wch: 11 }, { wch: 13 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, wsReturns, "Returns");

  return wb;
}

export default function StockroomApp() {
  const [items, setItems] = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [returns, setReturns] = useState([]);
  const [seq, setSeq] = useState(1);
  const [returnSeq, setReturnSeq] = useState(1);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("inventory");
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [toast, setToast] = useState(null);
  const [viewInvoice, setViewInvoice] = useState(null);
  const [viewReturn, setViewReturn] = useState(null);
  const [returningInvoice, setReturningInvoice] = useState(null);
  const [fileHandle, setFileHandle] = useState(null);
  const [linkedName, setLinkedName] = useState("");
  const [needsReconnect, setNeedsReconnect] = useState(false);
  const [savingFile, setSavingFile] = useState(false);
  const [pendingRestore, setPendingRestore] = useState(null);
  const [shareOpen, setShareOpen] = useState(false);
  const [preparingCatalog, setPreparingCatalog] = useState(false);
  const [auth, setAuth] = useState(DEFAULT_AUTH);
  const [loggedIn, setLoggedIn] = useState(false);
  const [theme, setTheme] = useState("light");
  const [accountOpen, setAccountOpen] = useState(false);
  const toastTimer = useRef(null);
  const restoreInputRef = useRef(null);
  const fsaSupported = typeof window !== "undefined" && !!window.showSaveFilePicker;

  useEffect(() => {
    try {
      const saved = window.localStorage.getItem(THEME_KEY);
      if (saved === "dark" || saved === "light") setTheme(saved);
      else if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) setTheme("dark");
    } catch (e) {}
  }, []);

  function toggleTheme() {
    setTheme((t) => {
      const next = t === "dark" ? "light" : "dark";
      try { window.localStorage.setItem(THEME_KEY, next); } catch (e) {}
      return next;
    });
  }

  useEffect(() => {
    (async () => {
      try { const s = await window.storage.get(STOCK_KEY); if (s?.value) setItems(JSON.parse(s.value)); } catch (e) {}
      try { const s = await window.storage.get(SALES_KEY); if (s?.value) setInvoices(JSON.parse(s.value)); } catch (e) {}
      try { const s = await window.storage.get(SEQ_KEY); if (s?.value) setSeq(JSON.parse(s.value)); } catch (e) {}
      try { const s = await window.storage.get(RETURNS_KEY); if (s?.value) setReturns(JSON.parse(s.value)); } catch (e) {}
      try { const s = await window.storage.get(RETURN_SEQ_KEY); if (s?.value) setReturnSeq(JSON.parse(s.value)); } catch (e) {}
      setLoaded(true);

      // ---- auth: load a saved login in the background; default already active ----
      try {
        const s = await window.storage.get(AUTH_KEY);
        if (s?.value) {
          const parsed = JSON.parse(s.value);
          if (parsed?.username && parsed?.password) setAuth(parsed);
        } else {
          window.storage.set(AUTH_KEY, JSON.stringify(DEFAULT_AUTH)).catch(() => {});
        }
      } catch (e) {}

      try {
        const s = await window.storage.get(SESSION_KEY);
        if (s?.value === "1") setLoggedIn(true);
      } catch (e) {}

      // ---- restore the previously linked Excel file, if any ----
      try {
        const handle = await loadFileHandle();
        if (handle) {
          const opts = { mode: "readwrite" };
          let permission = await handle.queryPermission(opts);
          if (permission === "prompt") {
            // Some browsers allow a silent re-grant if it was already
            // approved before; if not, this quietly fails and we ask
            // the person to reconnect with one click instead of
            // re-picking the file from scratch.
            try { permission = await handle.requestPermission(opts); } catch (e) {}
          }
          if (permission === "granted") {
            setFileHandle(handle);
            setLinkedName(handle.name);
          } else {
            setLinkedName(handle.name);
            setNeedsReconnect(true);
          }
        }
      } catch (e) {}
    })();
  }, []);

  useEffect(() => { if (loaded) window.storage.set(STOCK_KEY, JSON.stringify(items)).catch(() => {}); }, [items, loaded]);
  useEffect(() => { if (loaded) window.storage.set(SALES_KEY, JSON.stringify(invoices)).catch(() => {}); }, [invoices, loaded]);
  useEffect(() => { if (loaded) window.storage.set(SEQ_KEY, JSON.stringify(seq)).catch(() => {}); }, [seq, loaded]);
  useEffect(() => { if (loaded) window.storage.set(RETURNS_KEY, JSON.stringify(returns)).catch(() => {}); }, [returns, loaded]);
  useEffect(() => { if (loaded) window.storage.set(RETURN_SEQ_KEY, JSON.stringify(returnSeq)).catch(() => {}); }, [returnSeq, loaded]);

  function showToast(msg, kind = "ok") {
    setToast({ msg, kind });
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(null), 3200);
  }

  // ---- login / account ----
  function attemptLogin(username, password) {
    const current = auth || DEFAULT_AUTH;
    if (username.trim() === current.username && password === current.password) {
      setLoggedIn(true);
      window.storage.set(SESSION_KEY, "1").catch(() => {});
      return true;
    }
    return false;
  }

  function logout() {
    setLoggedIn(false);
    window.storage.set(SESSION_KEY, "0").catch(() => {});
  }

  async function resetAuthToDefault() {
    setAuth(DEFAULT_AUTH);
    try {
      await window.storage.set(AUTH_KEY, JSON.stringify(DEFAULT_AUTH));
      await window.storage.set(SESSION_KEY, "0");
    } catch (e) {}
    showToast(`Login reset — use ${DEFAULT_AUTH.username} / ${DEFAULT_AUTH.password}`);
  }

  // returns "" on success, or an error message to show in the form
  async function changeCredentials({ currentPassword, newUsername, newPassword }) {
    const current = auth || DEFAULT_AUTH;
    if (currentPassword !== current.password) {
      return "Current password is incorrect.";
    }
    const nextAuth = {
      username: newUsername?.trim() ? newUsername.trim() : current.username,
      password: newPassword ? newPassword : current.password,
    };
    setAuth(nextAuth);
    try {
      await window.storage.set(AUTH_KEY, JSON.stringify(nextAuth));
    } catch (e) {
      return "Couldn't save the new login — try again.";
    }
    setAccountOpen(false);
    showToast("Login details updated");
    return "";
  }

  // ---- stock CRUD ----
  function saveItem(item) {
    setItems((prev) => {
      const exists = prev.some((p) => p.id === item.id);
      if (exists) return prev.map((p) => (p.id === item.id ? item : p));
      return [...prev, item];
    });
    setEditing(null);
    showToast(item._isNew ? "Item added to inventory" : "Item updated");
  }

  function deleteItem(id) {
    setItems((prev) => prev.filter((p) => p.id !== id));
    setConfirmDelete(null);
    showToast("Item removed", "warn");
  }

  const filteredItems = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return items;
    return items.filter(
      (i) => i.name.toLowerCase().includes(q) || i.sku.toLowerCase().includes(q) || (i.category || "").toLowerCase().includes(q)
    );
  }, [items, search]);

  // ---- excel: inventory manual export ----
  // ---- full data backup / restore (JSON safety net, independent of Excel) ----
  function downloadBackup() {
    const payload = {
      type: "stockroom-backup",
      version: 2,
      exportedAt: new Date().toISOString(),
      items,
      invoices,
      seq,
      returns,
      returnSeq,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `stockroom_backup_${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showToast("Backup downloaded");
  }

  function pickRestoreFile() {
    restoreInputRef.current?.click();
  }

  function handleRestoreFile(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target.result);
        if (!parsed || !Array.isArray(parsed.items) || !Array.isArray(parsed.invoices)) {
          throw new Error("bad shape");
        }
        setPendingRestore(parsed);
      } catch (err) {
        showToast("That file doesn't look like a valid Stockroom backup", "warn");
      }
    };
    reader.onerror = () => showToast("Couldn't read that file", "warn");
    reader.readAsText(file);
  }

  function confirmRestore() {
    if (!pendingRestore) return;
    setItems(pendingRestore.items || []);
    setInvoices(pendingRestore.invoices || []);
    setSeq(pendingRestore.seq || 1);
    setReturns(pendingRestore.returns || []);
    setReturnSeq(pendingRestore.returnSeq || 1);
    setPendingRestore(null);
    showToast("Backup restored");
  }

  function exportInventory() {
    const rows = items.map((i) => ({
      SKU: i.sku, Name: i.name, Category: i.category || "",
      Quantity: i.quantity, "Unit Price": i.price, "Stock Value": +(i.price * i.quantity).toFixed(2),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 26 }, { wch: 16 }, { wch: 10 }, { wch: 12 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Inventory");
    XLSX.writeFile(wb, `inventory_${Date.now()}.xlsx`);
    showToast("Inventory exported");
  }

  // ---- shareable stock lists for customers ----
  // Option 1: editable Excel with prices — customer/owner can open and tweak in Excel.
  function downloadPriceListExcel() {
    const rows = items.map((i) => ({
      Item: i.name,
      Category: i.category || "",
      Price: i.price,
      Availability: i.quantity > 0 ? "In Stock" : "Out of Stock",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 28 }, { wch: 18 }, { wch: 12 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Price List");
    XLSX.writeFile(wb, `price_list_${Date.now()}.xlsx`);
    showToast("Price list downloaded (editable in Excel)");
    setShareOpen(false);
  }

  // Option 2: PDF catalog with no prices — safe to hand straight to customers.
  async function downloadCatalogPdf() {
    setPreparingCatalog(true);
    try {
      const { jsPDF } = await loadJsPDF();
      const doc = new jsPDF({ unit: "pt", format: "a4" });
      const marginX = 48;
      const rightX = 547;
      const rowH = 46;
      let y = 70;

      function drawHeader() {
        doc.setFont("helvetica", "bold");
        doc.setFontSize(19);
        doc.setTextColor(28, 36, 49);
        doc.text("Our Stock", marginX, y);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(10);
        doc.setTextColor(91, 100, 114);
        doc.text("Catalog — prices on request", marginX, y + 16);
        y += 40;
        doc.setDrawColor(28, 36, 49);
        doc.setLineWidth(1.2);
        doc.line(marginX, y, rightX, y);
        y += 8;
      }

      drawHeader();

      for (const it of items) {
        if (y + rowH > 790) {
          doc.addPage();
          y = 60;
        }
        const inStock = it.quantity > 0;
        if (it.image) {
          try { doc.addImage(it.image, "JPEG", marginX, y, 34, 34); } catch (e) {}
        } else {
          doc.setDrawColor(228, 220, 200);
          doc.rect(marginX, y, 34, 34);
        }
        doc.setFont("helvetica", "bold");
        doc.setFontSize(11.5);
        doc.setTextColor(28, 36, 49);
        doc.text(it.name, marginX + 46, y + 15);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(9.5);
        doc.setTextColor(91, 100, 114);
        doc.text(it.category || "", marginX + 46, y + 29);

        const badgeText = inStock ? "In Stock" : "Out of Stock";
        doc.setFontSize(9);
        const badgeW = doc.getTextWidth(badgeText) + 16;
        const badgeX = rightX - badgeW;
        doc.setFillColor(...(inStock ? [237, 247, 240] : [251, 234, 231]));
        doc.roundedRect(badgeX, y + 8, badgeW, 18, 4, 4, "F");
        doc.setTextColor(...(inStock ? [63, 122, 87] : [181, 72, 61]));
        doc.text(badgeText, badgeX + 8, y + 20);

        y += rowH;
        doc.setDrawColor(228, 220, 200);
        doc.setLineWidth(0.5);
        doc.line(marginX, y - 12, rightX, y - 12);
      }

      doc.save(`stock_catalog_${Date.now()}.pdf`);
      showToast("Customer catalog downloaded");
      setShareOpen(false);
    } catch (e) {
      showToast("Couldn't generate the catalog — check your connection and try again.", "warn");
    } finally {
      setPreparingCatalog(false);
    }
  }

  // ---- excel: link a single local file that gets overwritten in place ----
  async function linkExcelFile() {
    if (!fsaSupported) {
      showToast("This browser can't link a local file — sales will download a fresh Excel file instead.", "warn");
      return;
    }
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: "sales_log.xlsx",
        types: [{ description: "Excel Workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }],
      });
      setFileHandle(handle);
      setLinkedName(handle.name);
      setNeedsReconnect(false);
      await saveFileHandle(handle);
      await writeToLinkedFile(handle, invoices, returns);
      showToast(`Linked to "${handle.name}" — every sale and return now updates this file, and it'll stay linked next time you log in.`);
    } catch (e) {
      // user cancelled the picker; do nothing
    }
  }

  // re-grant permission to the already-remembered file without making
  // the person pick it again (only needed if the browser didn't allow
  // a silent re-grant on load)
  async function reconnectExcelFile() {
    try {
      const handle = await loadFileHandle();
      if (!handle) return;
      const permission = await handle.requestPermission({ mode: "readwrite" });
      if (permission === "granted") {
        setFileHandle(handle);
        setLinkedName(handle.name);
        setNeedsReconnect(false);
        showToast(`Reconnected to "${handle.name}".`);
      } else {
        showToast("Permission wasn't granted — try linking the file again.", "warn");
      }
    } catch (e) {
      showToast("Couldn't reconnect — try linking the file again.", "warn");
    }
  }

  async function writeToLinkedFile(handle, invoiceList, returnList) {
    const wb = buildWorkbook(invoiceList, returnList);
    const buffer = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    const writable = await handle.createWritable();
    await writable.write(buffer);
    await writable.close();
  }

  // manual "export a copy" button in Invoices tab — always available regardless of linking
  function exportSalesLogCopy() {
    const wb = buildWorkbook(invoices, returns);
    XLSX.writeFile(wb, `sales_log_${Date.now()}.xlsx`);
    showToast("Sales log exported as a new file");
  }

  async function persistSalesLog(invoiceList, returnList) {
    if (fileHandle) {
      try {
        setSavingFile(true);
        await writeToLinkedFile(fileHandle, invoiceList, returnList);
        setSavingFile(false);
        return "linked";
      } catch (e) {
        setSavingFile(false);
        showToast("Couldn't write to the linked file — downloading a copy instead.", "warn");
      }
    }
    // fallback: no linked file (or write failed) -> download a fresh snapshot
    const wb = buildWorkbook(invoiceList, returnList);
    XLSX.writeFile(wb, `sales_log_${Date.now()}.xlsx`);
    return "downloaded";
  }

  // ---- sale / invoice ----
  async function completeSale(customer, lines) {
    const number = "INV-" + String(seq).padStart(4, "0");
    const total = +lines.reduce((s, l) => s + l.qty * l.price, 0).toFixed(2);
    const invoice = { id: uid("inv_"), number, customer, date: todayStr(), lines, total };
    const nextInvoices = [invoice, ...invoices];

    setItems((prev) =>
      prev.map((it) => {
        const line = lines.find((l) => l.stockId === it.id);
        return line ? { ...it, quantity: it.quantity - line.qty } : it;
      })
    );
    setInvoices(nextInvoices);
    setSeq((s) => s + 1);
    setTab("invoices");
    setViewInvoice(invoice);

    const result = await persistSalesLog(nextInvoices, returns);
    showToast(
      result === "linked"
        ? `${number} created — saved to "${linkedName}"`
        : `${number} created — sales log downloaded`
    );
  }

  // ---- goods return ----
  function returnedQtyFor(invoiceId, stockId) {
    return returns
      .filter((r) => r.invoiceId === invoiceId)
      .reduce((sum, r) => sum + r.lines.filter((l) => l.stockId === stockId).reduce((s, l) => s + l.qty, 0), 0);
  }

  async function processReturn(invoice, returnLines, reason) {
    const number = "RET-" + String(returnSeq).padStart(4, "0");
    const total = +returnLines.reduce((s, l) => s + l.qty * l.price, 0).toFixed(2);
    const record = {
      id: uid("ret_"),
      number,
      invoiceId: invoice.id,
      invoiceNumber: invoice.number,
      customer: invoice.customer,
      date: todayStr(),
      lines: returnLines,
      total,
      reason: reason || "",
    };
    const nextReturns = [record, ...returns];
    const missingItems = returnLines.filter((l) => !items.some((it) => it.id === l.stockId));

    setItems((prev) =>
      prev.map((it) => {
        const line = returnLines.find((l) => l.stockId === it.id);
        return line ? { ...it, quantity: it.quantity + line.qty } : it;
      })
    );
    setReturns(nextReturns);
    setReturnSeq((s) => s + 1);
    setReturningInvoice(null);
    setViewInvoice(null);
    setTab("returns");
    setViewReturn(record);

    const result = await persistSalesLog(invoices, nextReturns);
    if (missingItems.length > 0) {
      showToast(`${number} recorded — some items no longer exist in inventory so stock wasn't restored for them`, "warn");
    } else {
      showToast(
        result === "linked"
          ? `${number} recorded — stock restored and saved to "${linkedName}"`
          : `${number} recorded — stock restored, sales log downloaded`
      );
    }
  }

  // ---- PDF invoice ----
  async function downloadInvoicePdf(invoice) {
    try {
      const { jsPDF } = await loadJsPDF();
      const doc = new jsPDF({ unit: "pt", format: "a4" });
      const marginX = 48;
      let y = 64;

      doc.setFont("helvetica", "bold");
      doc.setFontSize(20);
      doc.text("Invoice", marginX, y);
      doc.setTextColor(198, 138, 46);
      doc.setFontSize(12);
      doc.text(invoice.number, marginX, y + 20);

      doc.setTextColor(60, 66, 78);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10.5);
      const rightX = 547;
      doc.text(invoice.date, rightX, 58, { align: "right" });
      doc.text(`Billed to: ${invoice.customer}`, rightX, 74, { align: "right" });

      y = 110;
      doc.setDrawColor(28, 36, 49);
      doc.setLineWidth(1.2);
      doc.line(marginX, y, rightX, y);
      y += 22;

      doc.setFont("helvetica", "bold");
      doc.setFontSize(9.5);
      doc.setTextColor(91, 100, 114);
      doc.text("ITEM", marginX, y);
      doc.text("QTY", 330, y, { align: "right" });
      doc.text("UNIT PRICE", 430, y, { align: "right" });
      doc.text("AMOUNT", rightX, y, { align: "right" });
      y += 8;
      doc.setDrawColor(228, 220, 200);
      doc.setLineWidth(0.75);
      doc.line(marginX, y, rightX, y);
      y += 18;

      doc.setFont("helvetica", "normal");
      doc.setFontSize(10.5);
      doc.setTextColor(28, 36, 49);
      invoice.lines.forEach((l) => {
        doc.setFont("helvetica", "bold");
        doc.text(l.name, marginX, y);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(8.5);
        doc.setTextColor(91, 100, 114);
        doc.text(l.sku, marginX, y + 12);
        doc.setFontSize(10.5);
        doc.setTextColor(28, 36, 49);
        doc.text(String(l.qty), 330, y, { align: "right" });
        doc.text(currency(l.price), 430, y, { align: "right" });
        doc.text(currency(l.qty * l.price), rightX, y, { align: "right" });
        y += 30;
        doc.setDrawColor(228, 220, 200);
        doc.setLineWidth(0.5);
        doc.line(marginX, y - 12, rightX, y - 12);
      });

      y += 6;
      doc.setDrawColor(28, 36, 49);
      doc.setLineWidth(1.2);
      doc.line(400, y, rightX, y);
      y += 20;
      doc.setFont("helvetica", "bold");
      doc.setFontSize(13);
      doc.text("Total", 400, y);
      doc.text(currency(invoice.total), rightX, y, { align: "right" });

      y += 40;
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9.5);
      doc.setTextColor(91, 100, 114);
      doc.text("Thank you for your purchase.", 300, y, { align: "center" });

      doc.save(`${invoice.number}.pdf`);
    } catch (e) {
      showToast("Couldn't generate the PDF — check your connection and try again.", "warn");
    }
  }

  if (!loggedIn) {
    return (
      <div data-sr-theme={theme} style={{ fontFamily: "'IBM Plex Sans', sans-serif", background: PAPER, minHeight: "100dvh", color: INK }}>
        <style>{`
          ${THEME_VARS}
          ${FONT_IMPORT}
          * { box-sizing: border-box; }
          button { font-family: inherit; cursor: pointer; }
          input { font-family: inherit; }
          button, input { transition: background 140ms ease, border-color 140ms ease, transform 120ms ease, box-shadow 140ms ease; }
          button:focus-visible, input:focus-visible { outline: 2px solid ${AMBER}; outline-offset: 2px; }
          @keyframes sr-modal-in { from { opacity: 0; transform: translateY(10px) scale(0.98); } to { opacity: 1; transform: translateY(0) scale(1); } }
          @media (prefers-reduced-motion: reduce) { *, *::before, *::after { animation-duration: 0.01ms !important; transition-duration: 0.01ms !important; } }
          @media (max-width: 640px) {
            .sr-login-shell { flex-direction: column !important; max-width: 400px !important; }
            .sr-login-brand { flex: none !important; padding: 26px 24px !important; }
            .sr-login-form { padding: 28px 24px !important; }
          }
        `}</style>
        <LoginScreen onLogin={attemptLogin} onResetDefault={resetAuthToDefault} theme={theme} onToggleTheme={toggleTheme} />
      </div>
    );
  }

  return (
    <div data-sr-theme={theme} style={{ fontFamily: "'IBM Plex Sans', sans-serif", background: PAPER, minHeight: "100dvh", color: INK, display: "flex" }}>
      <style>{`
        ${THEME_VARS}
        ${FONT_IMPORT}
        * { box-sizing: border-box; }
        .stockroom-scroll::-webkit-scrollbar { width: 8px; height: 8px; }
        .stockroom-scroll::-webkit-scrollbar-thumb { background: ${LINE}; border-radius: 8px; }
        button { font-family: inherit; cursor: pointer; }
        input, select { font-family: inherit; }
        button, input, select, .sr-row, .navbtn { transition: background 140ms ease, border-color 140ms ease, color 140ms ease, transform 120ms ease, box-shadow 140ms ease; }
        .navbtn:focus-visible, button:focus-visible, input:focus-visible, select:focus-visible {
          outline: 2px solid ${AMBER}; outline-offset: 2px;
        }
        @media (prefers-reduced-motion: reduce) {
          *, *::before, *::after { animation-duration: 0.01ms !important; transition-duration: 0.01ms !important; }
        }
        @keyframes sr-modal-in {
          from { opacity: 0; transform: translateY(10px) scale(0.98); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        @keyframes sr-overlay-in {
          from { opacity: 0; }
          to { opacity: 1; }
        }
        @keyframes sr-toast-in {
          from { opacity: 0; transform: translate(-50%, 10px); }
          to { opacity: 1; transform: translate(-50%, 0); }
        }
        @keyframes sr-fade-up {
          from { opacity: 0; transform: translateY(6px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @media (max-width: 760px) {
          .sr-sidebar { display: none !important; }
          .sr-mobile-nav {
            display: flex !important;
            flex-wrap: nowrap !important;
            overflow: hidden;
            position: sticky;
            top: 0;
            z-index: 10;
          }
          .sr-mobile-nav button {
            min-width: 0;
            padding: 10px 2px !important;
            font-size: 10.5px !important;
          }
          .sr-mobile-nav button svg { width: 16px; height: 16px; }
          .sr-main { padding: 16px 14px 40px !important; }
          .sr-sell-grid { grid-template-columns: 1fr !important; }
          .sr-summary-card { position: static !important; }
          .sr-search { max-width: none !important; }
          .sr-table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
          .sr-table-wrap table { min-width: 620px; }
          .sr-modal-pad { padding: 18px 16px !important; }
          .sr-header { flex-direction: column !important; align-items: stretch !important; }
          .sr-header-actions { width: 100%; }
          .sr-header-actions button { flex: 1; justify-content: center; }
          .sr-mobile-backup { display: flex !important; }
          .sr-mobile-backup button { flex: 1; justify-content: center; }
          .sr-stat-grid { grid-template-columns: repeat(2, 1fr) !important; }
        }
        @media (max-width: 480px) {
          .sr-form-grid { grid-template-columns: 1fr !important; }
          .sr-invoice-actions { width: 100%; justify-content: stretch !important; }
          .sr-invoice-actions button { flex: 1; justify-content: center; }
          .sr-stat-grid { grid-template-columns: 1fr !important; }
        }
      `}</style>

      <Sidebar
        tab={tab}
        setTab={setTab}
        counts={{ items: items.length, invoices: invoices.length, customers: new Set(invoices.map((i) => i.customer.trim().toLowerCase())).size, returns: returns.length }}
        linkedName={linkedName}
        onBackup={downloadBackup}
        onRestore={pickRestoreFile}
        username={(auth || DEFAULT_AUTH).username}
        onAccount={() => setAccountOpen(true)}
        onLogout={logout}
        theme={theme}
        onToggleTheme={toggleTheme}
      />
      <input ref={restoreInputRef} type="file" accept="application/json" onChange={handleRestoreFile} style={{ display: "none" }} />

      <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column" }}>
        <MobileNav tab={tab} setTab={setTab} />
        <main style={{ flex: 1, padding: "28px 32px 60px", overflowY: "auto" }} className="stockroom-scroll sr-main">
          {tab === "inventory" && (
            <InventoryView
              items={filteredItems}
              allCount={items.length}
              search={search}
              setSearch={setSearch}
              onAdd={() => setEditing("new")}
              onEdit={(it) => setEditing(it)}
              onDelete={(it) => setConfirmDelete(it)}
              onExport={exportInventory}
              onShare={() => setShareOpen(true)}
            />
          )}
          {tab === "sell" && <SellView items={items} onComplete={completeSale} />}
          {tab === "invoices" && (
            <InvoicesView
              invoices={invoices}
              returns={returns}
              onExportCopy={exportSalesLogCopy}
              onView={(inv) => setViewInvoice(inv)}
              onLink={linkExcelFile}
              linkedName={linkedName}
              needsReconnect={needsReconnect}
              onReconnect={reconnectExcelFile}
              fsaSupported={fsaSupported}
              savingFile={savingFile}
              onBackup={downloadBackup}
              onRestore={pickRestoreFile}
            />
          )}
          {tab === "customers" && (
            <CustomersView invoices={invoices} returns={returns} onView={(inv) => setViewInvoice(inv)} />
          )}
          {tab === "returns" && (
            <ReturnsView returns={returns} onView={(r) => setViewReturn(r)} />
          )}
        </main>
      </div>

      {editing && <ItemEditor item={editing === "new" ? null : editing} onCancel={() => setEditing(null)} onSave={saveItem} />}
      {confirmDelete && (
        <ConfirmDialog
          title="Remove item"
          message={`Remove "${confirmDelete.name}" from inventory? This can't be undone.`}
          confirmLabel="Remove"
          danger
          onCancel={() => setConfirmDelete(null)}
          onConfirm={() => deleteItem(confirmDelete.id)}
        />
      )}
      {viewInvoice && (
        <InvoiceModal
          invoice={viewInvoice}
          onClose={() => setViewInvoice(null)}
          onDownloadPdf={() => downloadInvoicePdf(viewInvoice)}
          onStartReturn={() => setReturningInvoice(viewInvoice)}
          returnedQtyFor={returnedQtyFor}
        />
      )}
      {returningInvoice && (
        <ReturnModal
          invoice={returningInvoice}
          returnedQtyFor={returnedQtyFor}
          onCancel={() => setReturningInvoice(null)}
          onSubmit={(lines, reason) => processReturn(returningInvoice, lines, reason)}
        />
      )}
      {viewReturn && <ReturnDetailModal record={viewReturn} onClose={() => setViewReturn(null)} />}
      {shareOpen && (
        <ShareStockListModal
          itemCount={items.length}
          preparingCatalog={preparingCatalog}
          onClose={() => setShareOpen(false)}
          onDownloadCatalog={downloadCatalogPdf}
          onDownloadPriceList={downloadPriceListExcel}
        />
      )}
      {pendingRestore && (
        <ConfirmDialog
          title="Restore backup"
          message={`This will replace your current data with the backup (${pendingRestore.items?.length || 0} items, ${pendingRestore.invoices?.length || 0} invoices, ${pendingRestore.returns?.length || 0} returns) from ${pendingRestore.exportedAt ? new Date(pendingRestore.exportedAt).toLocaleString() : "an unknown date"}. This can't be undone.`}
          confirmLabel="Restore"
          danger
          onCancel={() => setPendingRestore(null)}
          onConfirm={confirmRestore}
        />
      )}
      {accountOpen && (
        <AccountModal
          username={(auth || DEFAULT_AUTH).username}
          onClose={() => setAccountOpen(false)}
          onSubmit={changeCredentials}
        />
      )}
      {toast && <Toast msg={toast.msg} kind={toast.kind} />}
    </div>
  );
}

// ---------------- Login ----------------

function LoginScreen({ onLogin, onResetDefault, theme, onToggleTheme }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [resetArmed, setResetArmed] = useState(false);

  function submit(e) {
    e?.preventDefault();
    if (!username.trim() || !password) {
      setError("Enter both the ID and password.");
      return;
    }
    setSubmitting(true);
    const ok = onLogin(username, password);
    if (!ok) {
      setError("Incorrect ID or password.");
      setSubmitting(false);
    }
  }

  function handleReset() {
    if (!resetArmed) {
      setResetArmed(true);
      return;
    }
    onResetDefault();
    setResetArmed(false);
    setError("");
  }

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 20, position: "relative" }}>
      <button
        onClick={onToggleTheme}
        title={theme === "dark" ? "Switch to light mode" : "Switch to dark mode"}
        style={{ position: "absolute", top: 20, right: 20, width: 34, height: 34, borderRadius: 9, border: `1px solid ${LINE}`, background: "var(--sr-card-bg)", color: INK, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: SHADOW_SM }}
      >
        {theme === "dark" ? <Sun size={16} /> : <Moon size={16} />}
      </button>
      <div
        className="sr-login-shell"
        style={{
          width: "100%", maxWidth: 760, display: "flex", borderRadius: 20, overflow: "hidden",
          boxShadow: SHADOW_LG, animation: "sr-modal-in 320ms cubic-bezier(0.16,1,0.3,1)",
        }}
      >
        {/* left: brand / ledger panel */}
        <div
          className="sr-login-brand"
          style={{
            flex: "0 0 42%", background: SIDEBAR_BG, color: SIDEBAR_TEXT, padding: "38px 32px",
            display: "flex", flexDirection: "column", justifyContent: "space-between", position: "relative", overflow: "hidden",
          }}
        >
          <div style={{ position: "absolute", inset: 0, backgroundImage: `repeating-linear-gradient(${SIDEBAR_BG_ACTIVE} 0 1px, transparent 1px 34px)`, opacity: 0.5 }} />
          <div style={{ position: "relative" }}>
            <div style={{ width: 40, height: 40, background: `linear-gradient(155deg, ${AMBER} 0%, ${AMBER_DARK} 100%)`, borderRadius: 11, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 6px 14px rgba(198,138,46,0.35)", marginBottom: 18 }}>
              <Package size={21} color={SIDEBAR_BG} strokeWidth={2.4} />
            </div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 22, letterSpacing: -0.2 }}>Stockroom</div>
            <div style={{ fontSize: 12.5, color: "#9CA5B4", marginTop: 6, lineHeight: 1.6, maxWidth: 220 }}>
              Inventory, sales and returns — kept like a ledger.
            </div>
          </div>
          <div style={{ position: "relative", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: "#7C8494", letterSpacing: 0.4 }}>
            LEDGER v1 · ACCESS CONTROLLED
          </div>
        </div>

        {/* right: form */}
        <div style={{ flex: 1, background: "var(--sr-card-bg)", padding: "38px 34px" }} className="sr-login-form">
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 19, marginBottom: 3 }}>Sign in</div>
          <div style={{ fontSize: 12.5, color: SLATE, marginBottom: 22 }}>Enter your ID and password to continue.</div>

          <Field label="ID">
            <input
              autoFocus
              style={inputStyle}
              value={username}
              onChange={(e) => { setUsername(e.target.value); setError(""); }}
              onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
              placeholder="Admin"
            />
          </Field>

          <div style={{ marginTop: 14 }}>
            <Field label="Password">
              <div style={{ position: "relative" }}>
                <input
                  type={showPw ? "text" : "password"}
                  style={{ ...inputStyle, paddingRight: 38 }}
                  value={password}
                  onChange={(e) => { setPassword(e.target.value); setError(""); }}
                  onKeyDown={(e) => { if (e.key === "Enter") submit(); }}
                  placeholder="••••••••"
                />
                <button
                  type="button"
                  onClick={() => setShowPw((v) => !v)}
                  aria-label={showPw ? "Hide password" : "Show password"}
                  style={{ position: "absolute", right: 8, top: 0, bottom: 0, background: "none", border: "none", color: SLATE, display: "flex", alignItems: "center" }}
                >
                  {showPw ? <EyeOff size={15} /> : <Eye size={15} />}
                </button>
              </div>
            </Field>
          </div>

          {error && (
            <div style={{ marginTop: 12, fontSize: 12.5, color: RED, display: "flex", alignItems: "center", gap: 6 }}>
              <AlertTriangle size={13} /> {error}
            </div>
          )}

          <button
            type="button"
            onClick={() => submit()}
            disabled={submitting}
            style={{ width: "100%", marginTop: 20, padding: "11px 0", borderRadius: 9, border: "none", background: INK, color: PAPER, fontWeight: 700, fontSize: 13.5, boxShadow: SHADOW_SM, opacity: submitting ? 0.75 : 1 }}
          >
            Sign in
          </button>

          <div style={{ marginTop: 16, textAlign: "center" }}>
            <button
              type="button"
              onClick={handleReset}
              style={{ background: "none", border: "none", color: resetArmed ? RED : SLATE, fontSize: 12, fontWeight: 600, padding: "4px 6px" }}
            >
              {resetArmed ? `Click again to reset to ${DEFAULT_AUTH.username} / ${DEFAULT_AUTH.password}` : "Forgot login details? Reset to default"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function AccountModal({ username, onClose, onSubmit }) {
  const [currentPassword, setCurrentPassword] = useState("");
  const [newUsername, setNewUsername] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function submit() {
    setError("");
    if (!currentPassword) return setError("Enter your current password to confirm.");
    if (newPassword && newPassword.length < 4) return setError("New password should be at least 4 characters.");
    if (newPassword && newPassword !== confirmPassword) return setError("New passwords don't match.");
    setBusy(true);
    const err = await onSubmit({ currentPassword, newUsername, newPassword });
    setBusy(false);
    if (err) setError(err);
  }

  return (
    <Overlay onClose={onClose}>
      <div className="sr-modal-pad" style={{ padding: "22px 24px 24px" }}>
        <ModalHeader title="Login details" onClose={onClose} />
        <div style={{ color: SLATE, fontSize: 12.5, marginTop: 4 }}>
          Signed in as <strong style={{ color: INK }}>{username}</strong>. Change the ID and/or password below.
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 14, marginTop: 18 }}>
          <Field label="New ID (leave blank to keep current)">
            <input style={inputStyle} value={newUsername} onChange={(e) => setNewUsername(e.target.value)} placeholder={username} />
          </Field>
          <Field label="New password (leave blank to keep current)">
            <input type="password" style={inputStyle} value={newPassword} onChange={(e) => setNewPassword(e.target.value)} placeholder="••••••••" />
          </Field>
          <Field label="Confirm new password">
            <input type="password" style={inputStyle} value={confirmPassword} onChange={(e) => setConfirmPassword(e.target.value)} placeholder="••••••••" />
          </Field>
          <Field label="Current password (required to confirm)">
            <input type="password" style={inputStyle} value={currentPassword} onChange={(e) => setCurrentPassword(e.target.value)} placeholder="••••••••" />
          </Field>
        </div>

        {error && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: RED, display: "flex", alignItems: "center", gap: 6 }}>
            <AlertTriangle size={13} /> {error}
          </div>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
          <TextButton onClick={onClose}>Cancel</TextButton>
          <button
            onClick={submit}
            disabled={busy}
            style={{ padding: "9px 16px", borderRadius: 8, border: "none", background: INK, color: PAPER, fontWeight: 700, fontSize: 13, opacity: busy ? 0.7 : 1 }}
          >
            {busy ? "Saving…" : "Save changes"}
          </button>
        </div>
      </div>
    </Overlay>
  );
}

// ---------------- Sidebar ----------------

function Sidebar({ tab, setTab, counts, linkedName, onBackup, onRestore, username, onAccount, onLogout, theme, onToggleTheme }) {
  const items = [
    { key: "inventory", label: "Inventory", icon: Boxes, count: counts.items },
    { key: "sell", label: "New sale", icon: ShoppingCart },
    { key: "invoices", label: "Invoices", icon: Receipt, count: counts.invoices },
    { key: "returns", label: "Returns", icon: RotateCcw, count: counts.returns },
    { key: "customers", label: "Customers", icon: Users, count: counts.customers },
  ];
  return (
    <aside className="sr-sidebar" style={{ width: 236, background: SIDEBAR_BG, color: SIDEBAR_TEXT, display: "flex", flexDirection: "column", padding: "26px 16px", flexShrink: 0 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 11, padding: "0 8px 26px" }}>
        <div style={{ width: 36, height: 36, background: `linear-gradient(155deg, ${AMBER} 0%, ${AMBER_DARK} 100%)`, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 4px 10px rgba(198,138,46,0.35)" }}>
          <Package size={19} color={SIDEBAR_BG} strokeWidth={2.4} />
        </div>
        <div style={{ flex: 1 }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 17, letterSpacing: 0.2 }}>Stockroom</div>
          <div style={{ fontSize: 11, color: "#8B93A3", fontFamily: "'IBM Plex Mono', monospace", letterSpacing: 0.6 }}>LEDGER v1</div>
        </div>
        <button
          onClick={onToggleTheme}
          title={theme === "dark" ? "Switch to light mode" : "Switch to dark mode"}
          style={{ width: 30, height: 30, borderRadius: 8, border: "1px solid #333E52", background: "transparent", color: "#C7CCD6", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}
        >
          {theme === "dark" ? <Sun size={15} /> : <Moon size={15} />}
        </button>
      </div>

      <nav style={{ display: "flex", flexDirection: "column", gap: 3 }}>
        {items.map(({ key, label, icon: Icon, count }) => {
          const active = tab === key;
          return (
            <button
              key={key}
              className="navbtn"
              onClick={() => setTab(key)}
              style={{
                position: "relative", display: "flex", alignItems: "center", gap: 11, padding: "10px 12px 10px 15px",
                borderRadius: 9, border: "none", textAlign: "left",
                background: active ? SIDEBAR_BG_ACTIVE : "transparent",
                color: active ? SIDEBAR_TEXT : "#A7AEBB",
                fontSize: 14.5, fontWeight: active ? 600 : 500,
              }}
            >
              {active && (
                <span style={{ position: "absolute", left: 0, top: "22%", bottom: "22%", width: 3, borderRadius: 3, background: AMBER }} />
              )}
              <Icon size={17} strokeWidth={2.2} color={active ? AMBER : undefined} />
              <span style={{ flex: 1 }}>{label}</span>
              {typeof count === "number" && (
                <span style={{ fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace", color: active ? AMBER : "#6B7280" }}>{count}</span>
              )}
            </button>
          );
        })}
      </nav>

      <div style={{ marginTop: "auto", padding: "13px 12px", borderTop: "1px solid #333E52", fontSize: 11.5, color: "#7C8494", lineHeight: 1.5 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6, color: linkedName ? "#8FD1A8" : "#7C8494", marginBottom: 4 }}>
          <Link2 size={12} />
          {linkedName ? `Auto-saving to ${linkedName}` : "No file linked yet"}
        </div>
        Link a local Excel file from the Invoices tab so every sale updates it in place.
      </div>

      <div style={{ padding: "13px 12px 2px", borderTop: "1px solid #333E52", display: "flex", flexDirection: "column", gap: 6 }}>
        <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, color: "#6B7280", marginBottom: 2 }}>Data safety net</div>
        <SidebarSmallButton icon={DatabaseBackup} label="Download backup" onClick={onBackup} />
        <SidebarSmallButton icon={UploadCloud} label="Restore from backup" onClick={onRestore} />
      </div>

      <div style={{ padding: "13px 12px 2px", borderTop: "1px solid #333E52", display: "flex", flexDirection: "column", gap: 6 }}>
        <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, color: "#6B7280", marginBottom: 2 }}>{username}</div>
        <SidebarSmallButton icon={Users} label="Change login details" onClick={onAccount} />
        <SidebarSmallButton icon={X} label="Log out" onClick={onLogout} />
      </div>
    </aside>
  );
}

function SidebarSmallButton({ icon: Icon, label, onClick }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "flex", alignItems: "center", gap: 8, padding: "8px 10px", borderRadius: 7,
        border: "1px solid #333E52", background: "transparent", color: "#C7CCD6", fontSize: 12.5, fontWeight: 500,
      }}
      onMouseEnter={(e) => (e.currentTarget.style.background = "#2A3444")}
      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
    >
      <Icon size={14} /> {label}
    </button>
  );
}

function MobileNav({ tab, setTab }) {
  const items = [
    { key: "inventory", label: "Stock", icon: Boxes },
    { key: "sell", label: "Sell", icon: ShoppingCart },
    { key: "invoices", label: "Invoices", icon: Receipt },
    { key: "returns", label: "Returns", icon: RotateCcw },
    { key: "customers", label: "Customers", icon: Users },
  ];
  return (
    <div className="sr-mobile-nav" style={{ display: "none", borderBottom: `1px solid ${LINE}`, background: PAPER }}>
      {items.map(({ key, label, icon: Icon }) => {
        const active = tab === key;
        return (
          <button
            key={key}
            onClick={() => setTab(key)}
            style={{
              flex: 1, padding: "13px 0", border: "none", background: "none",
              display: "flex", flexDirection: "column", alignItems: "center", gap: 4,
              color: active ? INK : SLATE, borderBottom: active ? `2px solid ${AMBER}` : "2px solid transparent",
              fontSize: 12, fontWeight: active ? 600 : 500,
            }}
          >
            <Icon size={17} />
            {label}
          </button>
        );
      })}
    </div>
  );
}

// ---------------- Inventory ----------------

function InventoryView({ items, allCount, search, setSearch, onAdd, onEdit, onDelete, onExport, onShare }) {
  return (
    <div>
      <Header
        title="Inventory"
        subtitle={`${allCount} item${allCount === 1 ? "" : "s"} tracked`}
        actions={
          <>
            <IconButton icon={Share2} label="Share list" onClick={onShare} variant="ghost" />
            <IconButton icon={Download} label="Export Excel" onClick={onExport} variant="ghost" />
            <IconButton icon={Plus} label="Add item" onClick={onAdd} variant="solid" />
          </>
        }
      />

      <div className="sr-search" style={{ position: "relative", maxWidth: 340, marginBottom: 18 }}>
        <Search size={15} color={SLATE} style={{ position: "absolute", left: 12, top: 12 }} />
        <input
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Search by name, SKU, category…"
          style={{ width: "100%", padding: "10px 12px 10px 34px", borderRadius: 9, border: `1px solid ${LINE}`, background: "var(--sr-card-bg)", fontSize: 13.5, color: INK }}
        />
      </div>

      {items.length === 0 ? (
        <EmptyState
          icon={Boxes}
          title={allCount === 0 ? "No stock yet" : "No matches"}
          body={allCount === 0 ? "Add your first item to start tracking inventory." : "Try a different search term."}
          action={allCount === 0 ? { label: "Add item", onClick: onAdd } : null}
        />
      ) : (
        <div style={{ background: "var(--sr-card-bg)", border: `1px solid ${LINE}`, borderRadius: 13, overflow: "hidden", boxShadow: SHADOW_SM }}>
          <div className="sr-table-wrap">
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13.5 }}>
            <thead>
              <tr style={{ background: PAPER_DIM, textAlign: "left" }}>
                {["SKU", "Name", "Category", "Qty", "Price", "Value", ""].map((h, i) => (
                  <th key={i} style={{ padding: "11px 16px", fontWeight: 600, fontSize: 11.5, letterSpacing: 0.4, textTransform: "uppercase", color: SLATE, borderBottom: `1px solid ${LINE}` }}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {items.map((it) => {
                const low = it.quantity <= (it.lowStockAt ?? 5);
                return (
                  <tr key={it.id} className="sr-row" style={{ borderBottom: `1px solid ${LINE}` }} onMouseEnter={(e) => (e.currentTarget.style.background = PAPER_DIM)} onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}>
                    <td style={{ padding: "12px 16px", fontFamily: "'IBM Plex Mono', monospace", color: SLATE, fontSize: 12.5 }}>{it.sku}</td>
                    <td style={{ padding: "12px 16px", fontWeight: 600 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <div style={{ width: 30, height: 30, borderRadius: 7, background: PAPER_DIM, flexShrink: 0, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }}>
                          {it.image ? <img src={it.image} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <ImageIcon size={13} color={SLATE} />}
                        </div>
                        {it.name}
                      </div>
                    </td>
                    <td style={{ padding: "12px 16px", color: SLATE }}>{it.category || "—"}</td>
                    <td style={{ padding: "12px 16px" }}>
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: low ? RED : INK, display: "inline-flex", alignItems: "center", gap: 5 }}>
                        {low && <AlertTriangle size={12} />}
                        {it.quantity}
                      </span>
                    </td>
                    <td style={{ padding: "12px 16px", fontFamily: "'IBM Plex Mono', monospace" }}>{currency(it.price)}</td>
                    <td style={{ padding: "12px 16px", fontFamily: "'IBM Plex Mono', monospace", color: SLATE }}>{currency(it.price * it.quantity)}</td>
                    <td style={{ padding: "12px 16px", textAlign: "right", whiteSpace: "nowrap" }}>
                      <RowIconBtn icon={Pencil} onClick={() => onEdit(it)} />
                      <RowIconBtn icon={Trash2} onClick={() => onDelete(it)} danger />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          </div>
        </div>
      )}
    </div>
  );
}

function RowIconBtn({ icon: Icon, onClick, danger }) {
  return (
    <button
      onClick={onClick}
      style={{ background: "none", border: "none", padding: 6, borderRadius: 6, marginLeft: 2, color: danger ? RED : SLATE, display: "inline-flex" }}
      onMouseEnter={(e) => (e.currentTarget.style.background = danger ? "#FBEAE7" : PAPER_DIM)}
      onMouseLeave={(e) => (e.currentTarget.style.background = "none")}
    >
      <Icon size={15} />
    </button>
  );
}

// ---------------- Item editor ----------------

function ShareStockListModal({ itemCount, preparingCatalog, onClose, onDownloadCatalog, onDownloadPriceList }) {
  return (
    <Overlay onClose={onClose}>
      <div className="sr-modal-pad" style={{ padding: "22px 24px 24px" }}>
        <ModalHeader title="Share stock list" onClose={onClose} />
        <div style={{ color: SLATE, fontSize: 12.5, marginTop: 4 }}>
          Choose a format to send to customers — based on your current {itemCount} item{itemCount === 1 ? "" : "s"}.
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 18 }}>
          <div style={{ border: `1px solid ${LINE}`, borderRadius: 11, padding: 16 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
              <div style={{ width: 32, height: 32, borderRadius: 8, background: PAPER_DIM, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <EyeOff size={15} color={AMBER} />
              </div>
              <div style={{ fontWeight: 700, fontSize: 14 }}>Customer catalog — no prices</div>
            </div>
            <div style={{ fontSize: 12.5, color: SLATE, marginBottom: 12, lineHeight: 1.5 }}>
              A clean PDF with item names, categories, photos and stock status. Safe to send straight to customers as-is.
            </div>
            <button
              onClick={onDownloadCatalog}
              disabled={preparingCatalog}
              style={{ width: "100%", padding: "10px 0", borderRadius: 8, border: "none", background: INK, color: PAPER, fontWeight: 700, fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center", gap: 7, opacity: preparingCatalog ? 0.7 : 1 }}
            >
              <FileText size={14} /> {preparingCatalog ? "Preparing…" : "Download PDF catalog"}
            </button>
          </div>

          <div style={{ border: `1px solid ${LINE}`, borderRadius: 11, padding: 16 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
              <div style={{ width: 32, height: 32, borderRadius: 8, background: PAPER_DIM, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <Eye size={15} color={GREEN} />
              </div>
              <div style={{ fontWeight: 700, fontSize: 14 }}>Price list — editable Excel</div>
            </div>
            <div style={{ fontSize: 12.5, color: SLATE, marginBottom: 12, lineHeight: 1.5 }}>
              An .xlsx with item, category, price and availability. Fully editable — tweak prices or trim rows before sending it on.
            </div>
            <button
              onClick={onDownloadPriceList}
              style={{ width: "100%", padding: "10px 0", borderRadius: 8, border: `1px solid ${LINE}`, background: "var(--sr-card-bg)", color: INK, fontWeight: 700, fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center", gap: 7 }}
            >
              <FileSpreadsheet size={14} /> Download Excel price list
            </button>
          </div>
        </div>
      </div>
    </Overlay>
  );
}

function ItemEditor({ item, onSave, onCancel }) {
  const isNew = !item;
  const [form, setForm] = useState(item || { id: uid("sku_"), name: "", sku: "", category: "", quantity: 0, price: 0, lowStockAt: 5, image: null });
  const [error, setError] = useState("");
  const [imageBusy, setImageBusy] = useState(false);
  const fileInputRef = useRef(null);

  function update(field, value) { setForm((f) => ({ ...f, [field]: value })); }

  async function handleImagePick(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    if (!file.type.startsWith("image/")) return setError("Please choose an image file.");
    setError("");
    setImageBusy(true);
    try {
      const dataUrl = await resizeImage(file);
      update("image", dataUrl);
    } catch (err) {
      setError("Couldn't read that image — try a different file.");
    }
    setImageBusy(false);
  }

  function submit() {
    if (!form.name.trim()) return setError("Item name is required.");
    if (!form.sku.trim()) return setError("SKU is required.");
    if (form.quantity < 0 || form.price < 0) return setError("Quantity and price can't be negative.");
    onSave({ ...form, quantity: Number(form.quantity), price: Number(form.price), _isNew: isNew });
  }

  return (
    <Overlay onClose={onCancel}>
      <div className="sr-modal-pad" style={{ padding: "22px 24px 24px" }}>
        <ModalHeader title={isNew ? "Add item" : "Edit item"} onClose={onCancel} />

        <div style={{ marginTop: 16 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: SLATE, marginBottom: 6 }}>Photo (optional)</div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{
              width: 72, height: 72, borderRadius: 9, border: `1px dashed ${LINE}`, background: PAPER_DIM,
              display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", flexShrink: 0,
            }}>
              {form.image ? (
                <img src={form.image} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
              ) : (
                <ImageIcon size={20} color={SLATE} />
              )}
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <input ref={fileInputRef} type="file" accept="image/*" onChange={handleImagePick} style={{ display: "none" }} />
              <TextButtonOutline onClick={() => fileInputRef.current?.click()} disabled={imageBusy}>
                {imageBusy ? "Uploading…" : form.image ? "Replace photo" : "Upload photo"}
              </TextButtonOutline>
              {form.image && (
                <TextButtonOutline onClick={() => update("image", null)} muted>Remove</TextButtonOutline>
              )}
            </div>
          </div>
        </div>

        <div className="sr-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginTop: 16 }}>
          <Field label="Item name" span={2}>
            <input autoFocus value={form.name} onChange={(e) => update("name", e.target.value)} style={inputStyle} placeholder="e.g. Cotton T-Shirt — Blue, M" />
          </Field>
          <Field label="SKU / code">
            <input value={form.sku} onChange={(e) => update("sku", e.target.value)} style={{ ...inputStyle, fontFamily: "'IBM Plex Mono', monospace" }} placeholder="TSH-BLU-M" />
          </Field>
          <Field label="Category">
            <input value={form.category} onChange={(e) => update("category", e.target.value)} style={inputStyle} placeholder="Apparel" />
          </Field>
          <Field label="Quantity in stock">
            <input type="number" min="0" value={form.quantity} onChange={(e) => update("quantity", e.target.value)} style={{ ...inputStyle, fontFamily: "'IBM Plex Mono', monospace" }} />
          </Field>
          <Field label="Unit price (₹)">
            <input type="number" min="0" step="0.01" value={form.price} onChange={(e) => update("price", e.target.value)} style={{ ...inputStyle, fontFamily: "'IBM Plex Mono', monospace" }} />
          </Field>
          <Field label="Low-stock alert below">
            <input type="number" min="0" value={form.lowStockAt} onChange={(e) => update("lowStockAt", Number(e.target.value))} style={{ ...inputStyle, fontFamily: "'IBM Plex Mono', monospace" }} />
          </Field>
        </div>
        {error && <div style={{ color: RED, fontSize: 12.5, marginTop: 12 }}>{error}</div>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 22 }}>
          <TextButton onClick={onCancel}>Cancel</TextButton>
          <IconButton icon={Check} label={isNew ? "Add item" : "Save changes"} onClick={submit} variant="solid" />
        </div>
      </div>
    </Overlay>
  );
}

// ---------------- Sell / invoice creation ----------------

function SellView({ items, onComplete }) {
  const [customer, setCustomer] = useState("");
  const [cart, setCart] = useState([]);
  const [pickId, setPickId] = useState("");
  const [pickQty, setPickQty] = useState(1);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const available = items.filter((i) => i.quantity > 0);
  const cartLines = cart.map((c) => {
    const it = items.find((i) => i.id === c.stockId);
    return it ? { stockId: it.id, sku: it.sku, name: it.name, price: it.price, qty: c.qty, image: it.image || null } : null;
  }).filter(Boolean);
  const total = cartLines.reduce((s, l) => s + l.qty * l.price, 0);

  function addToCart() {
    setError("");
    if (!pickId) return setError("Choose an item to add.");
    const it = items.find((i) => i.id === pickId);
    const qty = Number(pickQty);
    if (!qty || qty <= 0) return setError("Quantity must be at least 1.");
    const already = cart.find((c) => c.stockId === pickId)?.qty || 0;
    if (already + qty > it.quantity) return setError(`Only ${it.quantity} in stock for "${it.name}".`);
    setCart((prev) => {
      const exists = prev.find((c) => c.stockId === pickId);
      if (exists) return prev.map((c) => (c.stockId === pickId ? { ...c, qty: c.qty + qty } : c));
      return [...prev, { stockId: pickId, qty }];
    });
    setPickId("");
    setPickQty(1);
  }

  function removeLine(stockId) { setCart((prev) => prev.filter((c) => c.stockId !== stockId)); }

  async function finalize() {
    setError("");
    if (!customer.trim()) return setError("Enter a customer name.");
    if (cartLines.length === 0) return setError("Add at least one item to the sale.");
    setBusy(true);
    await onComplete(customer.trim(), cartLines);
    setBusy(false);
    setCustomer("");
    setCart([]);
  }

  return (
    <div>
      <Header title="New sale" subtitle="Sell stock and generate an invoice" />
      <div className="sr-sell-grid" style={{ display: "grid", gridTemplateColumns: "1.1fr 0.9fr", gap: 22, alignItems: "start" }}>
        <div style={{ background: "var(--sr-card-bg)", border: `1px solid ${LINE}`, borderRadius: 12, padding: 20 }}>
          <Field label="Customer name">
            <input value={customer} onChange={(e) => setCustomer(e.target.value)} style={inputStyle} placeholder="e.g. Ravi Sharma" />
          </Field>

          <div style={{ height: 16 }} />
          <div style={{ fontSize: 12.5, fontWeight: 600, color: SLATE, marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.4 }}>Add item to sale</div>
          <div style={{ display: "flex", gap: 8 }}>
            <select value={pickId} onChange={(e) => setPickId(e.target.value)} style={{ ...inputStyle, flex: 1 }}>
              <option value="">Choose item…</option>
              {available.map((it) => (
                <option key={it.id} value={it.id}>{it.name} — {it.quantity} in stock — {currency(it.price)}</option>
              ))}
            </select>
            <input type="number" min="1" value={pickQty} onChange={(e) => setPickQty(e.target.value)} style={{ ...inputStyle, width: 76 }} />
            <IconButton icon={Plus} label="Add" onClick={addToCart} variant="ghost" />
          </div>

          {error && <div style={{ color: RED, fontSize: 12.5, marginTop: 12 }}>{error}</div>}

          <div style={{ height: 20 }} />
          {cartLines.length === 0 ? (
            <div style={{ color: SLATE, fontSize: 13, padding: "18px 0", textAlign: "center", border: `1px dashed ${LINE}`, borderRadius: 9 }}>No items added yet</div>
          ) : (
            <div style={{ border: `1px solid ${LINE}`, borderRadius: 9, overflow: "hidden" }}>
              {cartLines.map((l) => (
                <div key={l.stockId} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", borderBottom: `1px solid ${LINE}`, fontSize: 13.5 }}>
                  <div style={{ width: 28, height: 28, borderRadius: 6, background: PAPER_DIM, flexShrink: 0, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    {l.image ? <img src={l.image} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <ImageIcon size={12} color={SLATE} />}
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600 }}>{l.name}</div>
                    <div style={{ fontSize: 11.5, color: SLATE, fontFamily: "'IBM Plex Mono', monospace" }}>{l.sku}</div>
                  </div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", color: SLATE }}>{l.qty} × {currency(l.price)}</div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, width: 84, textAlign: "right" }}>{currency(l.qty * l.price)}</div>
                  <RowIconBtn icon={X} onClick={() => removeLine(l.stockId)} danger />
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="sr-summary-card" style={{ background: INK, color: PAPER, borderRadius: 12, padding: 22, position: "sticky", top: 0 }}>
          <div style={{ fontSize: 12, textTransform: "uppercase", letterSpacing: 0.6, color: "#9AA3B2", marginBottom: 4 }}>Sale total</div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 32, marginBottom: 18 }}>{currency(total)}</div>
          <div style={{ fontSize: 13, color: "#C7CCD6", marginBottom: 18, lineHeight: 1.6 }}>
            {cartLines.length} line{cartLines.length === 1 ? "" : "s"} · {cartLines.reduce((s, l) => s + l.qty, 0)} unit{cartLines.reduce((s, l) => s + l.qty, 0) === 1 ? "" : "s"} · customer {customer.trim() ? `"${customer.trim()}"` : "not set"}
          </div>
          <button
            onClick={finalize}
            disabled={busy}
            style={{ width: "100%", padding: "12px 0", borderRadius: 9, border: "none", background: AMBER, color: INK, fontWeight: 700, fontSize: 14, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, opacity: busy ? 0.7 : 1 }}
          >
            <Receipt size={16} /> {busy ? "Saving…" : "Complete sale & generate invoice"}
          </button>
          <div style={{ fontSize: 11.5, color: "#8B93A3", marginTop: 10, lineHeight: 1.5 }}>
            Completing the sale reduces stock and saves the sales log to Excel.
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------------- Invoices list ----------------

function InvoicesView({ invoices, returns, onExportCopy, onView, onLink, linkedName, needsReconnect, onReconnect, fsaSupported, savingFile, onBackup, onRestore }) {
  const returnedTotalFor = (invoiceId) =>
    (returns || []).filter((r) => r.invoiceId === invoiceId).reduce((s, r) => s + r.total, 0);
  return (
    <div>
      <Header
        title="Invoices"
        subtitle={`${invoices.length} invoice${invoices.length === 1 ? "" : "s"} recorded`}
        actions={
          <>
            <IconButton icon={Download} label="Export a copy" onClick={onExportCopy} variant="ghost" />
            <IconButton icon={savingFile ? RefreshCw : Link2} label={linkedName ? "Change linked file" : "Link Excel file"} onClick={onLink} variant="solid" />
          </>
        }
      />

      <div className="sr-mobile-backup" style={{ display: "none", gap: 8, marginBottom: 14 }}>
        <IconButton icon={DatabaseBackup} label="Download backup" onClick={onBackup} variant="ghost" />
        <IconButton icon={UploadCloud} label="Restore backup" onClick={onRestore} variant="ghost" />
      </div>

      {needsReconnect && (
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, fontSize: 12.5, color: AMBER_DARK, background: AMBER_TINT, border: `1px solid ${AMBER}`, borderRadius: 9, padding: "9px 13px" }}>
          <AlertTriangle size={13} color={AMBER_DARK} />
          <span style={{ flex: 1 }}>
            <strong style={{ color: INK }}>{linkedName}</strong> is remembered, but this browser needs one click to reconnect before saving to it again.
          </span>
          <button onClick={onReconnect} style={{ border: "none", background: AMBER_DARK, color: "#fff", borderRadius: 7, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer", flexShrink: 0 }}>
            Reconnect
          </button>
        </div>
      )}

      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 18, fontSize: 12.5, color: SLATE, background: linkedName ? "#EEF6F0" : PAPER_DIM, border: `1px solid ${linkedName ? "#CFE6D5" : LINE}`, borderRadius: 9, padding: "9px 13px" }}>
        <Link2 size={13} color={linkedName ? GREEN : SLATE} />
        {linkedName ? (
          <span>Auto-saving every sale and return to <strong style={{ color: INK }}>{linkedName}</strong> — no repeat downloads.</span>
        ) : fsaSupported ? (
          <span>No file linked yet. Link one and every sale or return updates it in place instead of downloading a new file.</span>
        ) : (
          <span>Your browser doesn't support linking a local file, so each sale or return will download a fresh sales-log file.</span>
        )}
      </div>

      {invoices.length === 0 ? (
        <EmptyState icon={Receipt} title="No invoices yet" body="Completed sales will show up here." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {invoices.map((inv) => {
            const returnedAmt = returnedTotalFor(inv.id);
            return (
              <button
                key={inv.id}
                onClick={() => onView(inv)}
                style={{ display: "flex", alignItems: "center", gap: 14, textAlign: "left", background: "var(--sr-card-bg)", border: `1px solid ${LINE}`, borderRadius: 11, padding: "14px 16px" }}
              >
                <div style={{ width: 38, height: 38, borderRadius: 8, background: PAPER_DIM, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <Receipt size={17} color={AMBER} />
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", fontSize: 13 }}>{inv.number}</div>
                  <div style={{ fontSize: 12.5, color: SLATE }}>{inv.customer} · {inv.date} · {inv.lines.length} item{inv.lines.length === 1 ? "" : "s"}</div>
                  {returnedAmt > 0 && (
                    <div style={{ fontSize: 11, color: RED, marginTop: 2, display: "flex", alignItems: "center", gap: 3 }}>
                      <RotateCcw size={10} /> {currency(returnedAmt)} returned
                    </div>
                  )}
                </div>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 15 }}>{currency(inv.total)}</div>
                <ChevronRight size={16} color={SLATE} />
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------------- Customers ----------------

function CustomersView({ invoices, returns, onView }) {
  const [search, setSearch] = useState("");
  const [selectedKey, setSelectedKey] = useState(null);

  const customers = useMemo(() => {
    const map = new Map();
    invoices.forEach((inv) => {
      const name = (inv.customer || "Unknown").trim();
      const key = name.toLowerCase();
      if (!map.has(key)) map.set(key, { key, name, orders: [], total: 0, unitsBought: 0 });
      const c = map.get(key);
      c.orders.push(inv);
      c.total += inv.total;
      c.unitsBought += inv.lines.reduce((s, l) => s + l.qty, 0);
    });
    (returns || []).forEach((r) => {
      const key = (r.customer || "Unknown").trim().toLowerCase();
      const c = map.get(key);
      if (c) c.returnedTotal = (c.returnedTotal || 0) + r.total;
    });
    return Array.from(map.values()).sort((a, b) => b.orders.length - a.orders.length || b.total - a.total);
  }, [invoices, returns]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return customers;
    return customers.filter((c) => c.name.toLowerCase().includes(q));
  }, [customers, search]);

  const selected = customers.find((c) => c.key === selectedKey);
  const returnedTotalFor = (invoiceId) => (returns || []).filter((r) => r.invoiceId === invoiceId).reduce((s, r) => s + r.total, 0);

  if (selected) {
    const netSpent = selected.total - (selected.returnedTotal || 0);
    return (
      <div>
        <button onClick={() => setSelectedKey(null)} style={{ background: "none", border: "none", display: "flex", alignItems: "center", gap: 6, color: SLATE, fontSize: 13, padding: 0, marginBottom: 18 }}>
          <ArrowLeft size={15} /> All customers
        </button>

        <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 22 }}>
          <div style={{ width: 46, height: 46, borderRadius: 10, background: INK, color: AMBER, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 18, flexShrink: 0 }}>
            {selected.name.slice(0, 1).toUpperCase()}
          </div>
          <div>
            <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 22, margin: 0 }}>{selected.name}</h1>
            <div style={{ color: SLATE, fontSize: 13, marginTop: 2 }}>
              {selected.orders.length} order{selected.orders.length === 1 ? "" : "s"} · {selected.unitsBought} unit{selected.unitsBought === 1 ? "" : "s"} bought
            </div>
          </div>
        </div>

        <div className="sr-stat-grid" style={{ display: "grid", gridTemplateColumns: selected.returnedTotal ? "repeat(4, 1fr)" : "repeat(3, 1fr)", gap: 12, marginBottom: 24 }}>
          <StatCard label="Total spent" value={currency(selected.total)} />
          <StatCard label="Orders" value={selected.orders.length} />
          <StatCard label="Avg. order value" value={currency(selected.total / selected.orders.length)} />
          {selected.returnedTotal > 0 && <StatCard label="Net spent" value={currency(netSpent)} accent={RED} />}
        </div>

        <div style={{ fontSize: 12.5, fontWeight: 600, color: SLATE, marginBottom: 10, textTransform: "uppercase", letterSpacing: 0.4 }}>Purchase history</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {selected.orders.map((inv) => {
            const returnedAmt = returnedTotalFor(inv.id);
            return (
              <button
                key={inv.id}
                onClick={() => onView(inv)}
                style={{ display: "flex", alignItems: "center", gap: 14, textAlign: "left", background: "var(--sr-card-bg)", border: `1px solid ${LINE}`, borderRadius: 11, padding: "14px 16px" }}
              >
                <div style={{ width: 38, height: 38, borderRadius: 8, background: PAPER_DIM, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <Receipt size={17} color={AMBER} />
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", fontSize: 13 }}>{inv.number}</div>
                  <div style={{ fontSize: 12.5, color: SLATE }}>
                    {inv.date} · {inv.lines.map((l) => `${l.qty}× ${l.name}`).join(", ")}
                  </div>
                  {returnedAmt > 0 && (
                    <div style={{ fontSize: 11, color: RED, marginTop: 2, display: "flex", alignItems: "center", gap: 3 }}>
                      <RotateCcw size={10} /> {currency(returnedAmt)} returned
                    </div>
                  )}
                </div>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 15 }}>{currency(inv.total)}</div>
                <ChevronRight size={16} color={SLATE} />
              </button>
            );
          })}
        </div>
      </div>
    );
  }

  return (
    <div>
      <Header title="Customers" subtitle={`${customers.length} customer${customers.length === 1 ? "" : "s"} on record`} />

      <div className="sr-search" style={{ position: "relative", maxWidth: 340, marginBottom: 18 }}>
        <Search size={15} color={SLATE} style={{ position: "absolute", left: 12, top: 12 }} />
        <input
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Search customers by name…"
          style={{ width: "100%", padding: "10px 12px 10px 34px", borderRadius: 9, border: `1px solid ${LINE}`, background: "var(--sr-card-bg)", fontSize: 13.5, color: INK }}
        />
      </div>

      {filtered.length === 0 ? (
        <EmptyState icon={Users} title={customers.length === 0 ? "No customers yet" : "No matches"} body={customers.length === 0 ? "Customers appear here after you complete a sale." : "Try a different name."} />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {filtered.map((c) => (
            <button
              key={c.key}
              onClick={() => setSelectedKey(c.key)}
              style={{ display: "flex", alignItems: "center", gap: 14, textAlign: "left", background: "var(--sr-card-bg)", border: `1px solid ${LINE}`, borderRadius: 11, padding: "14px 16px" }}
            >
              <div style={{ width: 38, height: 38, borderRadius: 8, background: PAPER_DIM, color: AMBER, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 15 }}>
                {c.name.slice(0, 1).toUpperCase()}
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{c.name}</div>
                <div style={{ fontSize: 12.5, color: SLATE }}>
                  {c.orders.length} order{c.orders.length === 1 ? "" : "s"} · last on {c.orders[0].date}
                </div>
              </div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 15 }}>{currency(c.total)}</div>
              <ChevronRight size={16} color={SLATE} />
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function StatCard({ label, value, accent }) {
  return (
    <div style={{ background: "var(--sr-card-bg)", border: `1px solid ${LINE}`, borderRadius: 12, padding: "15px 16px", boxShadow: SHADOW_SM }}>
      <div style={{ fontSize: 11.5, color: SLATE, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 5, fontWeight: 600 }}>{label}</div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 19, color: accent || INK }}>{value}</div>
    </div>
  );
}

function InvoiceModal({ invoice, onClose, onDownloadPdf, onStartReturn, returnedQtyFor }) {
  const [downloading, setDownloading] = useState(false);
  async function handleDownload() {
    setDownloading(true);
    await onDownloadPdf();
    setDownloading(false);
  }
  const totalReturnable = invoice.lines.reduce((s, l) => s + (l.qty - (returnedQtyFor ? returnedQtyFor(invoice.id, l.stockId) : 0)), 0);
  return (
    <Overlay onClose={onClose}>
      <div className="sr-modal-pad" style={{ padding: "24px 26px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <button onClick={onClose} style={{ background: "none", border: "none", display: "flex", alignItems: "center", gap: 6, color: SLATE, fontSize: 13, padding: 0 }}>
            <ArrowLeft size={15} /> Back
          </button>
          <div className="sr-invoice-actions" style={{ display: "flex", gap: 8 }}>
            <button onClick={() => window.print()} style={{ background: "none", border: `1px solid ${LINE}`, borderRadius: 8, padding: "6px 12px", display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: INK }}>
              <Printer size={14} /> Print
            </button>
            <button onClick={handleDownload} disabled={downloading} style={{ background: INK, color: PAPER, border: "none", borderRadius: 8, padding: "6px 12px", display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 600, opacity: downloading ? 0.7 : 1 }}>
              <FileText size={14} /> {downloading ? "Preparing…" : "Download PDF"}
            </button>
          </div>
        </div>

        <div style={{ marginTop: 18, paddingBottom: 16, borderBottom: `2px solid ${INK}`, display: "flex", justifyContent: "space-between", alignItems: "flex-end" }}>
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 22 }}>Invoice</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", color: AMBER, fontWeight: 700, fontSize: 14, marginTop: 2 }}>{invoice.number}</div>
          </div>
          <div style={{ textAlign: "right", fontSize: 12.5, color: SLATE }}>
            <div>{invoice.date}</div>
            <div style={{ marginTop: 2 }}>Billed to: <strong style={{ color: INK }}>{invoice.customer}</strong></div>
          </div>
        </div>

        <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 16, fontSize: 13.5 }}>
          <thead>
            <tr style={{ borderBottom: `1px solid ${LINE}`, textAlign: "left", color: SLATE, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.4 }}>
              <th style={{ padding: "6px 0" }}>Item</th>
              <th style={{ padding: "6px 0", textAlign: "center" }}>Qty</th>
              <th style={{ padding: "6px 0", textAlign: "right" }}>Unit price</th>
              <th style={{ padding: "6px 0", textAlign: "right" }}>Amount</th>
            </tr>
          </thead>
          <tbody>
            {invoice.lines.map((l, i) => {
              const returnedQty = returnedQtyFor ? returnedQtyFor(invoice.id, l.stockId) : 0;
              return (
                <tr key={i} style={{ borderBottom: `1px solid ${LINE}` }}>
                  <td style={{ padding: "10px 0" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <div style={{ width: 26, height: 26, borderRadius: 6, background: PAPER_DIM, flexShrink: 0, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }}>
                        {l.image ? <img src={l.image} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <ImageIcon size={11} color={SLATE} />}
                      </div>
                      <div>
                        <div style={{ fontWeight: 600 }}>{l.name}</div>
                        <div style={{ fontSize: 11, color: SLATE, fontFamily: "'IBM Plex Mono', monospace" }}>{l.sku}</div>
                        {returnedQty > 0 && (
                          <div style={{ fontSize: 10.5, color: RED, marginTop: 2, display: "flex", alignItems: "center", gap: 3 }}>
                            <RotateCcw size={9} /> {returnedQty} returned
                          </div>
                        )}
                      </div>
                    </div>
                  </td>
                  <td style={{ padding: "10px 0", textAlign: "center", fontFamily: "'IBM Plex Mono', monospace" }}>{l.qty}</td>
                  <td style={{ padding: "10px 0", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>{currency(l.price)}</td>
                  <td style={{ padding: "10px 0", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{currency(l.qty * l.price)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>

        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 14 }}>
          <div style={{ width: 200 }}>
            <div style={{ display: "flex", justifyContent: "space-between", padding: "10px 0", borderTop: `2px solid ${INK}`, fontWeight: 700, fontSize: 16 }}>
              <span>Total</span>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{currency(invoice.total)}</span>
            </div>
          </div>
        </div>

        {onStartReturn && (
          <div style={{ marginTop: 16, paddingTop: 16, borderTop: `1px dashed ${LINE}` }}>
            <button
              onClick={onStartReturn}
              disabled={totalReturnable <= 0}
              style={{
                width: "100%", padding: "11px 0", borderRadius: 9, border: `1px solid ${totalReturnable <= 0 ? LINE : RED}`,
                background: "none", color: totalReturnable <= 0 ? SLATE : RED, fontWeight: 700, fontSize: 13.5,
                display: "flex", alignItems: "center", justifyContent: "center", gap: 8, opacity: totalReturnable <= 0 ? 0.6 : 1,
              }}
            >
              <RotateCcw size={15} /> {totalReturnable <= 0 ? "All items already returned" : "Process a return"}
            </button>
          </div>
        )}

        <div style={{ marginTop: 18, fontSize: 11.5, color: SLATE, textAlign: "center" }}>Thank you for your purchase.</div>
      </div>
    </Overlay>
  );
}

// ---------------- Returns ----------------

function ReturnModal({ invoice, returnedQtyFor, onCancel, onSubmit }) {
  const lineState = invoice.lines.map((l) => {
    const already = returnedQtyFor ? returnedQtyFor(invoice.id, l.stockId) : 0;
    const remaining = l.qty - already;
    return { ...l, already, remaining, returnQty: 0 };
  });
  const [rows, setRows] = useState(lineState);
  const [reason, setReason] = useState("");
  const [error, setError] = useState("");

  function updateQty(i, value) {
    setRows((prev) => prev.map((r, idx) => (idx === i ? { ...r, returnQty: value } : r)));
  }

  const returnLines = rows.filter((r) => Number(r.returnQty) > 0).map((r) => ({
    stockId: r.stockId, sku: r.sku, name: r.name, price: r.price, qty: Number(r.returnQty), image: r.image,
  }));
  const total = returnLines.reduce((s, l) => s + l.qty * l.price, 0);

  function submit() {
    setError("");
    if (returnLines.length === 0) return setError("Enter a quantity to return for at least one item.");
    for (const r of rows) {
      const q = Number(r.returnQty) || 0;
      if (q < 0) return setError("Return quantity can't be negative.");
      if (q > r.remaining) return setError(`Can't return more than ${r.remaining} of "${r.name}".`);
    }
    onSubmit(returnLines, reason.trim());
  }

  return (
    <Overlay onClose={onCancel}>
      <div className="sr-modal-pad" style={{ padding: "22px 24px 24px" }}>
        <ModalHeader title={`Process return — ${invoice.number}`} onClose={onCancel} />
        <div style={{ color: SLATE, fontSize: 12.5, marginTop: 4 }}>Choose how many units {invoice.customer} is returning for each item. Stock is added back automatically.</div>

        <div style={{ marginTop: 16, border: `1px solid ${LINE}`, borderRadius: 9, overflow: "hidden" }}>
          {rows.map((r, i) => (
            <div key={r.stockId + i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px", borderBottom: i < rows.length - 1 ? `1px solid ${LINE}` : "none" }}>
              <div style={{ width: 30, height: 30, borderRadius: 7, background: PAPER_DIM, flexShrink: 0, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }}>
                {r.image ? <img src={r.image} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <ImageIcon size={12} color={SLATE} />}
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 600, fontSize: 13.5 }}>{r.name}</div>
                <div style={{ fontSize: 11, color: SLATE }}>
                  Sold {r.qty} · {r.already > 0 ? `${r.already} already returned · ` : ""}{r.remaining} eligible
                </div>
              </div>
              <input
                type="number"
                min="0"
                max={r.remaining}
                value={r.returnQty}
                disabled={r.remaining <= 0}
                onChange={(e) => updateQty(i, e.target.value)}
                style={{ ...inputStyle, width: 64, textAlign: "center", opacity: r.remaining <= 0 ? 0.5 : 1 }}
              />
            </div>
          ))}
        </div>

        <div style={{ marginTop: 14 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: SLATE, marginBottom: 6 }}>Reason (optional)</div>
          <input value={reason} onChange={(e) => setReason(e.target.value)} style={inputStyle} placeholder="e.g. Wrong size, damaged, changed mind…" />
        </div>

        {error && <div style={{ color: RED, fontSize: 12.5, marginTop: 12 }}>{error}</div>}

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 20 }}>
          <div style={{ fontSize: 13, color: SLATE }}>
            Refund total: <strong style={{ color: INK, fontFamily: "'IBM Plex Mono', monospace" }}>{currency(total)}</strong>
          </div>
          <div style={{ display: "flex", gap: 10 }}>
            <TextButton onClick={onCancel}>Cancel</TextButton>
            <button
              onClick={submit}
              style={{ padding: "9px 16px", borderRadius: 8, border: "none", background: RED, color: "#fff", fontWeight: 700, fontSize: 13, display: "flex", alignItems: "center", gap: 7 }}
            >
              <RotateCcw size={14} /> Confirm return
            </button>
          </div>
        </div>
      </div>
    </Overlay>
  );
}

function ReturnsView({ returns, onView }) {
  return (
    <div>
      <Header title="Returns" subtitle={`${returns.length} return${returns.length === 1 ? "" : "s"} recorded`} />
      {returns.length === 0 ? (
        <EmptyState icon={RotateCcw} title="No returns yet" body="When a customer returns an item, process it from the invoice and it'll show up here." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {returns.map((r) => (
            <button
              key={r.id}
              onClick={() => onView(r)}
              style={{ display: "flex", alignItems: "center", gap: 14, textAlign: "left", background: "var(--sr-card-bg)", border: `1px solid ${LINE}`, borderRadius: 11, padding: "14px 16px" }}
            >
              <div style={{ width: 38, height: 38, borderRadius: 8, background: "#FBEAE7", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <RotateCcw size={17} color={RED} />
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", fontSize: 13 }}>{r.number}</div>
                <div style={{ fontSize: 12.5, color: SLATE }}>
                  {r.customer} · {r.date} · against {r.invoiceNumber} · {r.lines.reduce((s, l) => s + l.qty, 0)} unit{r.lines.reduce((s, l) => s + l.qty, 0) === 1 ? "" : "s"}
                </div>
              </div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 15, color: RED }}>-{currency(r.total)}</div>
              <ChevronRight size={16} color={SLATE} />
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function ReturnDetailModal({ record, onClose }) {
  return (
    <Overlay onClose={onClose}>
      <div className="sr-modal-pad" style={{ padding: "24px 26px" }}>
        <button onClick={onClose} style={{ background: "none", border: "none", display: "flex", alignItems: "center", gap: 6, color: SLATE, fontSize: 13, padding: 0 }}>
          <ArrowLeft size={15} /> Back
        </button>

        <div style={{ marginTop: 18, paddingBottom: 16, borderBottom: `2px solid ${INK}`, display: "flex", justifyContent: "space-between", alignItems: "flex-end" }}>
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 22 }}>Return</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", color: RED, fontWeight: 700, fontSize: 14, marginTop: 2 }}>{record.number}</div>
          </div>
          <div style={{ textAlign: "right", fontSize: 12.5, color: SLATE }}>
            <div>{record.date}</div>
            <div style={{ marginTop: 2 }}>Against: <strong style={{ color: INK }}>{record.invoiceNumber}</strong></div>
            <div style={{ marginTop: 2 }}>Customer: <strong style={{ color: INK }}>{record.customer}</strong></div>
          </div>
        </div>

        <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 16, fontSize: 13.5 }}>
          <thead>
            <tr style={{ borderBottom: `1px solid ${LINE}`, textAlign: "left", color: SLATE, fontSize: 11.5, textTransform: "uppercase", letterSpacing: 0.4 }}>
              <th style={{ padding: "6px 0" }}>Item</th>
              <th style={{ padding: "6px 0", textAlign: "center" }}>Qty returned</th>
              <th style={{ padding: "6px 0", textAlign: "right" }}>Unit price</th>
              <th style={{ padding: "6px 0", textAlign: "right" }}>Refund</th>
            </tr>
          </thead>
          <tbody>
            {record.lines.map((l, i) => (
              <tr key={i} style={{ borderBottom: `1px solid ${LINE}` }}>
                <td style={{ padding: "10px 0" }}>
                  <div style={{ fontWeight: 600 }}>{l.name}</div>
                  <div style={{ fontSize: 11, color: SLATE, fontFamily: "'IBM Plex Mono', monospace" }}>{l.sku}</div>
                </td>
                <td style={{ padding: "10px 0", textAlign: "center", fontFamily: "'IBM Plex Mono', monospace" }}>{l.qty}</td>
                <td style={{ padding: "10px 0", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>{currency(l.price)}</td>
                <td style={{ padding: "10px 0", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{currency(l.qty * l.price)}</td>
              </tr>
            ))}
          </tbody>
        </table>

        {record.reason && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: SLATE, background: PAPER_DIM, borderRadius: 8, padding: "9px 12px" }}>
            <strong style={{ color: INK }}>Reason: </strong>{record.reason}
          </div>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 14 }}>
          <div style={{ width: 200 }}>
            <div style={{ display: "flex", justifyContent: "space-between", padding: "10px 0", borderTop: `2px solid ${INK}`, fontWeight: 700, fontSize: 16 }}>
              <span>Total refund</span>
              <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: RED }}>{currency(record.total)}</span>
            </div>
          </div>
        </div>

        <div style={{ marginTop: 12, fontSize: 11.5, color: GREEN, textAlign: "center", display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
          <Check size={13} /> Stock was restored to inventory for this return.
        </div>
      </div>
    </Overlay>
  );
}

// ---------------- shared bits ----------------

const inputStyle = { width: "100%", padding: "9px 11px", borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13.5, color: INK, background: "var(--sr-card-bg)", boxShadow: "inset 0 1px 2px rgba(28,36,49,0.03)" };

function Header({ title, subtitle, actions }) {
  return (
    <div className="sr-header" style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 20, flexWrap: "wrap", gap: 12 }}>
      <div>
        <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 25, margin: 0, letterSpacing: -0.3 }}>{title}</h1>
        {subtitle && <div style={{ color: SLATE, fontSize: 13, marginTop: 4 }}>{subtitle}</div>}
      </div>
      {actions && <div className="sr-header-actions" style={{ display: "flex", gap: 8 }}>{actions}</div>}
    </div>
  );
}

function Field({ label, children, span }) {
  return (
    <div style={{ gridColumn: span === 2 ? "span 2" : undefined }}>
      <div style={{ fontSize: 12, fontWeight: 600, color: SLATE, marginBottom: 6 }}>{label}</div>
      {children}
    </div>
  );
}

function IconButton({ icon: Icon, label, onClick, variant }) {
  const solid = variant === "solid";
  const [hover, setHover] = useState(false);
  return (
    <button
      onClick={onClick}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        display: "flex", alignItems: "center", gap: 7, padding: "9px 14px", borderRadius: 9,
        border: solid ? "none" : `1px solid ${hover ? SLATE : LINE}`,
        background: solid ? (hover ? INK_SOFT : INK) : (hover ? PAPER_DIM : "var(--sr-card-bg)"),
        color: solid ? PAPER : INK, fontWeight: 600, fontSize: 13, whiteSpace: "nowrap",
        boxShadow: solid ? SHADOW_SM : "none",
        transform: hover ? "translateY(-1px)" : "none",
      }}
    >
      <Icon size={15} /> {label}
    </button>
  );
}

function TextButton({ children, onClick }) {
  return <button onClick={onClick} style={{ background: "none", border: "none", padding: "9px 14px", color: SLATE, fontSize: 13, fontWeight: 600, borderRadius: 8 }}>{children}</button>;
}

function TextButtonOutline({ children, onClick, disabled, muted }) {
  const [hover, setHover] = useState(false);
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        background: hover && !disabled ? PAPER_DIM : "var(--sr-card-bg)", border: `1px solid ${LINE}`, borderRadius: 8, padding: "8px 13px",
        color: muted ? RED : INK, fontSize: 12.5, fontWeight: 600, opacity: disabled ? 0.6 : 1,
      }}
    >
      {children}
    </button>
  );
}

function EmptyState({ icon: Icon, title, body, action }) {
  return (
    <div style={{ textAlign: "center", padding: "56px 20px", border: `1.5px dashed ${LINE}`, borderRadius: 14, background: "var(--sr-card-bg)" }}>
      <div style={{ width: 46, height: 46, borderRadius: 12, background: AMBER_TINT, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
        <Icon size={22} color={AMBER_DARK} />
      </div>
      <div style={{ fontWeight: 700, fontSize: 15.5, fontFamily: "'Space Grotesk', sans-serif" }}>{title}</div>
      <div style={{ color: SLATE, fontSize: 13, marginTop: 5 }}>{body}</div>
      {action && (
        <div style={{ marginTop: 18, display: "flex", justifyContent: "center" }}>
          <IconButton icon={Plus} label={action.label} onClick={action.onClick} variant="solid" />
        </div>
      )}
    </div>
  );
}

function Overlay({ children, onClose }) {
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(28,36,49,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 16, animation: "sr-overlay-in 160ms ease" }}>
      <div
        onClick={(e) => e.stopPropagation()}
        style={{ background: PAPER, borderRadius: 16, width: "100%", maxWidth: 520, maxHeight: "88vh", overflowY: "auto", boxShadow: SHADOW_LG, animation: "sr-modal-in 190ms cubic-bezier(0.16,1,0.3,1)" }}
        className="stockroom-scroll"
      >
        {children}
      </div>
    </div>
  );
}

function ModalHeader({ title, onClose }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
      <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 18, margin: 0 }}>{title}</h2>
      <button onClick={onClose} style={{ background: "none", border: "none", color: SLATE, padding: 6, borderRadius: 8, display: "flex" }}>
        <X size={18} />
      </button>
    </div>
  );
}

function ConfirmDialog({ title, message, confirmLabel, danger, onCancel, onConfirm }) {
  return (
    <Overlay onClose={onCancel}>
      <div className="sr-modal-pad" style={{ padding: "22px 24px" }}>
        <ModalHeader title={title} onClose={onCancel} />
        <p style={{ color: SLATE, fontSize: 13.5, marginTop: 12, lineHeight: 1.5 }}>{message}</p>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
          <TextButton onClick={onCancel}>Cancel</TextButton>
          <button onClick={onConfirm} style={{ padding: "9px 16px", borderRadius: 8, border: "none", background: danger ? RED : GREEN, color: "#fff", fontWeight: 700, fontSize: 13, boxShadow: SHADOW_SM }}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </Overlay>
  );
}

function Toast({ msg, kind }) {
  const bg = kind === "warn" ? RED : GREEN;
  return (
    <div style={{ position: "fixed", bottom: 22, left: "50%", transform: "translateX(-50%)", background: bg, color: "#fff", padding: "11px 20px", borderRadius: 10, fontSize: 13, fontWeight: 600, zIndex: 60, boxShadow: "0 12px 28px rgba(0,0,0,0.22)", animation: "sr-toast-in 220ms cubic-bezier(0.16,1,0.3,1)", display: "flex", alignItems: "center", gap: 8 }}>
      {kind === "warn" ? <AlertTriangle size={14} /> : <Check size={14} />}
      {msg}
    </div>
  );
}
