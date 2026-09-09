// src/entry.jsx
import React from "react";
import { createRoot } from "react-dom/client";

// src/App.jsx
import { useState, useEffect, useMemo, useCallback, useRef, Fragment } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard,
  Package,
  ArrowLeftRight,
  Building2,
  Tablet as TabletIcon,
  Users,
  Settings,
  Plus,
  Pencil,
  Trash2,
  X,
  Search,
  AlertTriangle,
  ChevronDown,
  ArrowUpCircle,
  ArrowDownCircle,
  RefreshCcw,
  CornerDownLeft,
  ShieldAlert,
  Ban,
  Loader2,
  Boxes,
  ClipboardCheck,
  Send,
  History,
  Inbox,
  FileText,
  Printer,
  ArrowLeft,
  Download,
  Upload,
  Menu
} from "lucide-react";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Legend
} from "recharts";
import { Fragment as Fragment2, jsx, jsxs } from "react/jsx-runtime";
var SUPABASE_URL = "https://frkbeuzymwuxnwohvmix.supabase.co";
var SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZya2JldXp5bXd1eG53b2h2bWl4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM1ODI0NzQsImV4cCI6MjA5OTE1ODQ3NH0.W6qIF_XRQtpLcPpN7ECR6N-EA9jU5BSjuT6dmQ2iGgY";
var APP_STATE_ROW_ID = "main";
async function supabaseRequest(path, options = {}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
      "Content-Type": "application/json",
      Prefer: "return=representation",
      ...options.headers || {}
    }
  });
  if (!res.ok) {
    const text2 = await res.text().catch(() => "");
    throw new Error(`Supabase ${res.status}: ${text2 || res.statusText}`);
  }
  if (res.status === 204) return null;
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}
async function loadAppState() {
  const rows = await supabaseRequest(`app_state?id=eq.${APP_STATE_ROW_ID}&select=data`);
  return rows && rows[0] ? rows[0].data : null;
}
async function saveAppState(data) {
  await supabaseRequest(`app_state?id=eq.${APP_STATE_ROW_ID}`, {
    method: "PATCH",
    body: JSON.stringify({ data })
  });
}
var TX_TYPES = [
  { id: "RECEIVE", label: "Receive Stock", dir: 1, icon: ArrowDownCircle, color: "#0F6E56" },
  { id: "ISSUE", label: "Issue Stock", dir: -1, icon: ArrowUpCircle, color: "#993C1D" },
  { id: "TRANSFER", label: "Transfer Stock", dir: 0, icon: ArrowLeftRight, color: "#185FA5" },
  { id: "ADJUST", label: "Adjust Stock", dir: null, icon: RefreshCcw, color: "#854F0B" },
  { id: "RETURN", label: "Return Stock", dir: 1, icon: CornerDownLeft, color: "#3B6D11" },
  { id: "DAMAGED", label: "Damaged Stock", dir: -1, icon: AlertTriangle, color: "#A32D2D" },
  { id: "LOST", label: "Lost Stock", dir: -1, icon: ShieldAlert, color: "#712B13" },
  { id: "DISPOSED", label: "Disposed Stock", dir: -1, icon: Ban, color: "#444441" }
];
var WAREHOUSE_TYPES = ["Central Warehouse", "Regional Warehouse", "Hospital Warehouse"];
var TABLET_STATUSES = ["In Stock", "Assigned", "Under Repair", "Damaged", "Lost", "Disposed"];
var DEFAULT_DATA = {
  categories: [
    { id: "cat-1", name: "RDT Kits", group: "Medical Supplies" },
    { id: "cat-2", name: "Tablets", group: "Equipment" },
    { id: "cat-3", name: "Blood Testing Devices", group: "Equipment" },
    { id: "cat-4", name: "Printers", group: "Equipment" },
    { id: "cat-5", name: "Barcode Scanners", group: "Equipment" }
  ],
  units: ["Piece", "Box", "Pack", "Kit", "Unit", "Carton"],
  hospitals: [
    { id: "wh-1", name: "Central Warehouse - Islamabad", type: "Central Warehouse", location: "Islamabad" },
    { id: "wh-2", name: "Regional Warehouse - Punjab", type: "Regional Warehouse", location: "Lahore" },
    { id: "wh-3", name: "DHQ Hospital Rawalpindi", type: "Hospital Warehouse", location: "Rawalpindi" }
  ],
  designations: ["Medical Officer", "Lab Technician", "Data Encoder", "Program Coordinator", "IT Support"],
  departments: ["Hepatitis Clinic", "Laboratory", "IT", "Administration", "Pharmacy"],
  sources: ["DHO Office G9", "Program HQ Islamabad", "PITB"],
  handoverOfficers: [
    { id: "off-1", name: "Asim Rauf", designation: "IT Manager" },
    { id: "off-2", name: "Wajahat Ahmed", designation: "IT Data Officer" },
    { id: "off-3", name: "Naveed Sheikh", designation: "IT Data Officer" }
  ],
  items: [],
  stockLevels: [],
  transactions: [],
  tablets: [],
  staff: []
};
var uid = (p = "id") => `${p}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
var todayISO = () => (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
var fmtDate = (d) => {
  if (!d) return "-";
  try {
    return new Date(d).toLocaleDateString(void 0, { year: "numeric", month: "short", day: "numeric" });
  } catch {
    return d;
  }
};
var cx = (...a) => a.filter(Boolean).join(" ");
function Btn({ children, variant = "secondary", className = "", ...props }) {
  const base = "inline-flex items-center gap-1.5 rounded-md px-3 py-2 text-sm font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed";
  const styles = {
    primary: "bg-slate-900 text-white hover:bg-slate-800",
    secondary: "bg-white text-slate-700 border border-slate-300 hover:bg-slate-50",
    danger: "bg-red-600 text-white hover:bg-red-700",
    ghost: "text-slate-600 hover:bg-slate-100"
  };
  return /* @__PURE__ */ jsx("button", { className: cx(base, styles[variant], className), ...props, children });
}
function Field({ label, children, required, className = "" }) {
  return /* @__PURE__ */ jsxs("label", { className: cx("flex flex-col gap-1 text-sm", className), children: [
    /* @__PURE__ */ jsxs("span", { className: "font-medium text-slate-700", children: [
      label,
      " ",
      required && /* @__PURE__ */ jsx("span", { className: "text-red-500", children: "*" })
    ] }),
    children
  ] });
}
var inputCls = "w-full rounded-md border border-slate-300 px-3 py-2 text-sm text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-900/20 focus:border-slate-400 bg-white";
function TextInput(props) {
  return /* @__PURE__ */ jsx("input", { ...props, className: cx(inputCls, props.className) });
}
function Select({ children, ...props }) {
  return /* @__PURE__ */ jsx("select", { ...props, className: cx(inputCls, "cursor-pointer", props.className), children });
}
function TextArea(props) {
  return /* @__PURE__ */ jsx("textarea", { ...props, className: cx(inputCls, "resize-none", props.className) });
}
function Badge({ children, tone = "slate" }) {
  const tones = {
    slate: "bg-slate-100 text-slate-700",
    green: "bg-emerald-100 text-emerald-800",
    red: "bg-red-100 text-red-800",
    amber: "bg-amber-100 text-amber-800",
    blue: "bg-blue-100 text-blue-800"
  };
  return /* @__PURE__ */ jsx("span", { className: cx("rounded-full px-2 py-0.5 text-xs font-medium whitespace-nowrap", tones[tone]), children });
}
function Modal({ title, onClose, children, width = "max-w-lg" }) {
  return /* @__PURE__ */ jsx("div", { className: "fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4", onMouseDown: (e) => e.target === e.currentTarget && onClose(), children: /* @__PURE__ */ jsxs("div", { className: cx("w-full rounded-lg bg-white shadow-xl max-h-[90vh] flex flex-col", width), children: [
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between border-b border-slate-200 px-5 py-3.5", children: [
      /* @__PURE__ */ jsx("h3", { className: "text-base font-semibold text-slate-900", children: title }),
      /* @__PURE__ */ jsx("button", { onClick: onClose, className: "rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600", "aria-label": "Close", children: /* @__PURE__ */ jsx(X, { size: 18 }) })
    ] }),
    /* @__PURE__ */ jsx("div", { className: "overflow-y-auto px-5 py-4", children })
  ] }) });
}
function ConfirmDialog({ title = "Confirm delete", message, confirmLabel = "Delete", onConfirm, onCancel }) {
  return /* @__PURE__ */ jsxs(Modal, { title, onClose: onCancel, width: "max-w-sm", children: [
    /* @__PURE__ */ jsx("p", { className: "text-sm text-slate-600", children: message }),
    /* @__PURE__ */ jsxs("div", { className: "mt-5 flex justify-end gap-2", children: [
      /* @__PURE__ */ jsx(Btn, { onClick: onCancel, children: "Cancel" }),
      /* @__PURE__ */ jsxs(Btn, { variant: "danger", onClick: onConfirm, children: [
        /* @__PURE__ */ jsx(Trash2, { size: 14 }),
        " ",
        confirmLabel
      ] })
    ] })
  ] });
}
function EmptyState({ icon: Icon, title, subtitle }) {
  return /* @__PURE__ */ jsxs("div", { className: "flex flex-col items-center justify-center gap-2 py-14 text-center", children: [
    /* @__PURE__ */ jsx(Icon, { size: 32, className: "text-slate-300" }),
    /* @__PURE__ */ jsx("p", { className: "text-sm font-medium text-slate-500", children: title }),
    subtitle && /* @__PURE__ */ jsx("p", { className: "text-xs text-slate-400", children: subtitle })
  ] });
}
function ConfigList({ label, items, onAdd, onEdit, onDelete, extraFields }) {
  const [draft, setDraft] = useState(extraFields ? extraFields.initial : "");
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const [justAddedId, setJustAddedId] = useState(null);
  useEffect(() => {
    if (!justAddedId) return;
    const t = setTimeout(() => setJustAddedId(null), 1800);
    return () => clearTimeout(t);
  }, [justAddedId]);
  return /* @__PURE__ */ jsxs("div", { className: "rounded-lg border border-slate-200 bg-white p-4", children: [
    /* @__PURE__ */ jsxs("h4", { className: "mb-3 text-sm font-semibold text-slate-800", children: [
      label,
      " ",
      /* @__PURE__ */ jsxs("span", { className: "font-normal text-slate-400", children: [
        "(",
        items.length,
        ")"
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mb-3 flex gap-2", children: [
      extraFields ? extraFields.render(draft, setDraft) : /* @__PURE__ */ jsx(TextInput, { placeholder: `Add new ${label.toLowerCase()}...`, value: draft, onChange: (e) => setDraft(e.target.value), onKeyDown: (e) => e.key === "Enter" && submit() }),
      /* @__PURE__ */ jsxs(Btn, { type: "button", variant: "primary", onClick: submit, children: [
        /* @__PURE__ */ jsx(Plus, { size: 15 }),
        " Add"
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex flex-col gap-1.5 max-h-64 overflow-y-auto pr-1", children: [
      items.length === 0 && /* @__PURE__ */ jsx("p", { className: "text-xs text-slate-400 py-2", children: "No entries yet - add your first one above." }),
      items.map((it) => {
        const id = it.id ?? it;
        const isEditing = editingId === id;
        const isNew = justAddedId === id;
        return /* @__PURE__ */ jsxs("div", { className: cx("flex items-center justify-between gap-2 rounded-md px-3 py-2 text-sm transition-colors", isNew ? "bg-emerald-50 ring-1 ring-emerald-300" : "bg-slate-50"), children: [
          isEditing ? /* @__PURE__ */ jsx("div", { className: "flex flex-1 gap-2", children: extraFields ? extraFields.render(editDraft, setEditDraft) : /* @__PURE__ */ jsx(TextInput, { value: editDraft, onChange: (e) => setEditDraft(e.target.value), className: "py-1" }) }) : /* @__PURE__ */ jsx("span", { className: "text-slate-700", children: extraFields ? extraFields.display(it) : it }),
          /* @__PURE__ */ jsx("div", { className: "flex shrink-0 gap-1", children: isEditing ? /* @__PURE__ */ jsxs(Fragment2, { children: [
            /* @__PURE__ */ jsx(Btn, { type: "button", variant: "ghost", className: "px-2 py-1", onClick: () => {
              onEdit(id, editDraft);
              setEditingId(null);
            }, children: "Save" }),
            /* @__PURE__ */ jsx(Btn, { type: "button", variant: "ghost", className: "px-2 py-1", onClick: () => setEditingId(null), children: "Cancel" })
          ] }) : /* @__PURE__ */ jsxs(Fragment2, { children: [
            /* @__PURE__ */ jsx("button", { type: "button", className: "rounded p-1 text-slate-400 hover:bg-slate-200 hover:text-slate-700", onClick: () => {
              setEditingId(id);
              setEditDraft(extraFields ? it : it);
            }, children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
            /* @__PURE__ */ jsx("button", { type: "button", className: "rounded p-1 text-slate-400 hover:bg-red-100 hover:text-red-600", onClick: () => onDelete(id), children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
          ] }) })
        ] }, id);
      })
    ] })
  ] });
  function submit() {
    if (extraFields) {
      if (!extraFields.valid(draft)) return;
      const newId = onAdd(draft);
      setJustAddedId(newId ?? null);
      setDraft(extraFields.initial);
    } else {
      const value = draft.trim();
      if (!value) return;
      onAdd(value);
      setJustAddedId(value);
      setDraft("");
    }
  }
}
var TABS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "inventory", label: "Inventory Items", icon: Package },
  { id: "transactions", label: "Stock Transactions", icon: ArrowLeftRight },
  { id: "warehouses", label: "Hospitals & Warehouses", icon: Building2 },
  { id: "tablets", label: "Tablet Assets", icon: TabletIcon },
  { id: "staff", label: "Staff", icon: Users },
  { id: "settings", label: "Admin Settings", icon: Settings }
];
function App() {
  const [data, setData] = useState(DEFAULT_DATA);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [connectionError, setConnectionError] = useState(null);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  useEffect(() => {
    (async () => {
      try {
        const remote = await loadAppState();
        if (remote && Object.keys(remote).length > 0) {
          setData({ ...DEFAULT_DATA, ...remote });
        } else {
          await saveAppState(DEFAULT_DATA);
        }
        setConnectionError(null);
      } catch (e) {
        console.error("Failed to load from Supabase", e);
        setConnectionError(`Could not connect to the database (${e.message}). Working locally for this session - changes won't be saved until the connection is restored.`);
      } finally {
        setLoading(false);
      }
    })();
  }, []);
  const persist = useCallback(async (next) => {
    setSaving(true);
    try {
      await saveAppState(next);
      setConnectionError(null);
    } catch (e) {
      console.error("Failed to save to Supabase", e);
      setConnectionError(`Could not save to the database (${e.message}). Your last change is only kept in this browser tab for now.`);
    } finally {
      setSaving(false);
    }
  }, []);
  const update = useCallback((updater) => {
    setData((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      persist(next);
      return next;
    });
  }, [persist]);
  const selectTab = (id) => {
    setActiveTab(id);
    setMobileMenuOpen(false);
  };
  if (loading) {
    return /* @__PURE__ */ jsxs("div", { className: "flex h-[70vh] items-center justify-center gap-2 text-slate-400", children: [
      /* @__PURE__ */ jsx(Loader2, { className: "animate-spin", size: 20 }),
      " Loading inventory system..."
    ] });
  }
  return /* @__PURE__ */ jsxs("div", { className: "app-shell relative flex w-full min-w-0 flex-col overflow-hidden bg-slate-50 text-slate-900 md:flex-row", children: [
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between border-b border-slate-200 bg-white px-4 py-3 md:hidden", children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2", children: [
        /* @__PURE__ */ jsx("div", { className: "flex h-8 w-8 items-center justify-center rounded-lg bg-slate-900 text-white", children: /* @__PURE__ */ jsx(Boxes, { size: 16 }) }),
        /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold text-slate-900", children: "HCV Program" })
      ] }),
      /* @__PURE__ */ jsx("button", { onClick: () => setMobileMenuOpen(true), className: "rounded-md p-2 text-slate-600 hover:bg-slate-100", "aria-label": "Open menu", children: /* @__PURE__ */ jsx(Menu, { size: 20 }) })
    ] }),
    mobileMenuOpen && /* @__PURE__ */ jsxs("div", { className: "absolute inset-0 z-50 flex md:hidden", children: [
      /* @__PURE__ */ jsx("div", { className: "absolute inset-0 bg-slate-900/40", onClick: () => setMobileMenuOpen(false) }),
      /* @__PURE__ */ jsxs("div", { className: "relative flex h-full w-72 max-w-[85vw] flex-col bg-white shadow-xl", children: [
        /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between border-b border-slate-200 px-4 py-4", children: [
          /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold text-slate-900", children: "Menu" }),
          /* @__PURE__ */ jsx("button", { onClick: () => setMobileMenuOpen(false), className: "rounded p-1 text-slate-400 hover:bg-slate-100", "aria-label": "Close menu", children: /* @__PURE__ */ jsx(X, { size: 18 }) })
        ] }),
        /* @__PURE__ */ jsx(SidebarNav, { activeTab, setActiveTab: selectTab, saving, connectionError })
      ] })
    ] }),
    /* @__PURE__ */ jsxs("aside", { className: "hidden w-64 shrink-0 flex-col border-r border-slate-200 bg-white md:flex", children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-center gap-2 border-b border-slate-200 px-5 py-4", children: [
        /* @__PURE__ */ jsx("div", { className: "flex h-9 w-9 items-center justify-center rounded-lg bg-slate-900 text-white", children: /* @__PURE__ */ jsx(Boxes, { size: 18 }) }),
        /* @__PURE__ */ jsxs("div", { children: [
          /* @__PURE__ */ jsx("p", { className: "text-sm font-semibold leading-tight text-slate-900", children: "HCV Program" }),
          /* @__PURE__ */ jsx("p", { className: "text-xs leading-tight text-slate-500", children: "Inventory & Assets" })
        ] })
      ] }),
      /* @__PURE__ */ jsx(SidebarNav, { activeTab, setActiveTab, saving, connectionError })
    ] }),
    /* @__PURE__ */ jsx("main", { className: "min-w-0 flex-1 overflow-y-auto overflow-x-hidden", children: /* @__PURE__ */ jsxs("div", { className: "mx-auto max-w-6xl p-3 sm:p-6", children: [
      connectionError && /* @__PURE__ */ jsxs("div", { className: "mb-4 flex items-start gap-2 rounded-md border border-amber-200 bg-amber-50 px-3 py-2.5 text-sm text-amber-800", children: [
        /* @__PURE__ */ jsx(AlertTriangle, { size: 16, className: "mt-0.5 shrink-0" }),
        /* @__PURE__ */ jsx("span", { children: connectionError })
      ] }),
      activeTab === "dashboard" && /* @__PURE__ */ jsx(Dashboard, { data, setActiveTab }),
      activeTab === "inventory" && /* @__PURE__ */ jsx(InventoryTab, { data, update }),
      activeTab === "transactions" && /* @__PURE__ */ jsx(TransactionsTab, { data }),
      activeTab === "warehouses" && /* @__PURE__ */ jsx(WarehousesTab, { data, update }),
      activeTab === "tablets" && /* @__PURE__ */ jsx(TabletsTab, { data, update }),
      activeTab === "staff" && /* @__PURE__ */ jsx(StaffTab, { data, update }),
      activeTab === "settings" && /* @__PURE__ */ jsx(SettingsTab, { data, update })
    ] }) })
  ] });
}
function SidebarNav({ activeTab, setActiveTab, saving, connectionError }) {
  return /* @__PURE__ */ jsxs(Fragment2, { children: [
    /* @__PURE__ */ jsx("nav", { className: "flex-1 space-y-0.5 overflow-y-auto p-3", children: TABS.map((t) => {
      const Icon = t.icon;
      const active = activeTab === t.id;
      return /* @__PURE__ */ jsxs(
        "button",
        {
          onClick: () => setActiveTab(t.id),
          className: cx(
            "flex w-full items-center gap-2.5 rounded-md px-3 py-2.5 text-sm font-medium transition-colors",
            active ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-100"
          ),
          children: [
            /* @__PURE__ */ jsx(Icon, { size: 16 }),
            t.label
          ]
        },
        t.id
      );
    }) }),
    /* @__PURE__ */ jsx("div", { className: cx("border-t border-slate-200 px-4 py-3 text-xs", connectionError ? "text-amber-600" : "text-slate-400"), children: connectionError ? "Offline - not connected" : saving ? "Saving to database..." : "Connected - all changes saved" })
  ] });
}
function Dashboard({ data, setActiveTab }) {
  const stockByItem = useMemo(() => {
    const map = {};
    data.stockLevels.forEach((s) => {
      map[s.itemId] = (map[s.itemId] || 0) + s.quantity;
    });
    return map;
  }, [data.stockLevels]);
  const totalItems = data.items.length;
  const totalStockQty = Object.values(stockByItem).reduce((a, b) => a + b, 0);
  const lowStockItems = data.items.filter((it) => (stockByItem[it.id] || 0) < (it.minThreshold ?? 0));
  const tabletsAssigned = data.tablets.filter((t) => t.status === "Assigned").length;
  const byCategory = useMemo(() => {
    return data.categories.map((c) => ({
      name: c.name,
      qty: data.items.filter((i) => i.categoryId === c.id).reduce((sum, i) => sum + (stockByItem[i.id] || 0), 0)
    })).filter((c) => c.qty > 0);
  }, [data.categories, data.items, stockByItem]);
  const byWarehouse = useMemo(() => {
    return data.hospitals.map((w) => ({
      name: w.name,
      qty: data.stockLevels.filter((s) => s.hospitalId === w.id).reduce((sum, s) => sum + s.quantity, 0)
    })).filter((w) => w.qty > 0);
  }, [data.hospitals, data.stockLevels]);
  const COLORS = ["#0F6E56", "#185FA5", "#854F0B", "#993C1D", "#534AB7", "#3B6D11", "#72243E"];
  return /* @__PURE__ */ jsxs("div", { className: "flex flex-col gap-6", children: [
    /* @__PURE__ */ jsxs("div", { children: [
      /* @__PURE__ */ jsx("h1", { className: "text-xl font-semibold text-slate-900", children: "Dashboard" }),
      /* @__PURE__ */ jsx("p", { className: "text-sm text-slate-500", children: "Program-wide overview of inventory, warehouses and tablet assets." })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-2 gap-4 md:grid-cols-4", children: [
      /* @__PURE__ */ jsx(StatCard, { label: "Inventory items", value: totalItems, icon: Package }),
      /* @__PURE__ */ jsx(StatCard, { label: "Total stock quantity", value: totalStockQty.toLocaleString(), icon: Boxes }),
      /* @__PURE__ */ jsx(StatCard, { label: "Low stock alerts", value: lowStockItems.length, icon: AlertTriangle, tone: lowStockItems.length ? "red" : "slate", onClick: () => setActiveTab("inventory") }),
      /* @__PURE__ */ jsx(StatCard, { label: "Tablets assigned", value: `${tabletsAssigned} / ${data.tablets.length}`, icon: TabletIcon, onClick: () => setActiveTab("tablets") })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 gap-4 lg:grid-cols-2", children: [
      /* @__PURE__ */ jsxs("div", { className: "rounded-lg border border-slate-200 bg-white p-4", children: [
        /* @__PURE__ */ jsx("h3", { className: "mb-3 text-sm font-semibold text-slate-800", children: "Stock quantity by category" }),
        byCategory.length === 0 ? /* @__PURE__ */ jsx(EmptyState, { icon: Package, title: "No stock recorded yet", subtitle: "Receive stock against an item to see category breakdown." }) : /* @__PURE__ */ jsx(ResponsiveContainer, { width: "100%", height: 260, children: /* @__PURE__ */ jsxs(BarChart, { data: byCategory, layout: "vertical", margin: { left: 10, right: 20 }, children: [
          /* @__PURE__ */ jsx(CartesianGrid, { strokeDasharray: "3 3", horizontal: false, stroke: "#e2e8f0" }),
          /* @__PURE__ */ jsx(XAxis, { type: "number", tick: { fontSize: 12 } }),
          /* @__PURE__ */ jsx(YAxis, { type: "category", dataKey: "name", tick: { fontSize: 12 }, width: 110 }),
          /* @__PURE__ */ jsx(Tooltip, {}),
          /* @__PURE__ */ jsx(Bar, { dataKey: "qty", fill: "#185FA5", radius: [0, 4, 4, 0] })
        ] }) })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "rounded-lg border border-slate-200 bg-white p-4", children: [
        /* @__PURE__ */ jsx("h3", { className: "mb-3 text-sm font-semibold text-slate-800", children: "Stock distribution by warehouse" }),
        byWarehouse.length === 0 ? /* @__PURE__ */ jsx(EmptyState, { icon: Building2, title: "No stock recorded yet" }) : /* @__PURE__ */ jsx(ResponsiveContainer, { width: "100%", height: 260, children: /* @__PURE__ */ jsxs(PieChart, { children: [
          /* @__PURE__ */ jsx(Pie, { data: byWarehouse, dataKey: "qty", nameKey: "name", innerRadius: 55, outerRadius: 90, paddingAngle: 2, children: byWarehouse.map((_, i) => /* @__PURE__ */ jsx(Cell, { fill: COLORS[i % COLORS.length] }, i)) }),
          /* @__PURE__ */ jsx(Tooltip, {}),
          /* @__PURE__ */ jsx(Legend, { wrapperStyle: { fontSize: 12 } })
        ] }) })
      ] })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "rounded-lg border border-slate-200 bg-white p-4", children: [
      /* @__PURE__ */ jsxs("div", { className: "mb-3 flex items-center justify-between", children: [
        /* @__PURE__ */ jsx("h3", { className: "text-sm font-semibold text-slate-800", children: "Low stock items" }),
        lowStockItems.length > 0 && /* @__PURE__ */ jsxs(Badge, { tone: "red", children: [
          lowStockItems.length,
          " below threshold"
        ] })
      ] }),
      lowStockItems.length === 0 ? /* @__PURE__ */ jsx("p", { className: "py-6 text-center text-sm text-slate-400", children: "All items are above their minimum threshold." }) : /* @__PURE__ */ jsxs("table", { className: "w-full min-w-[720px] text-sm", children: [
        /* @__PURE__ */ jsx("thead", { children: /* @__PURE__ */ jsxs("tr", { className: "border-b border-slate-200 text-left text-xs uppercase text-slate-400", children: [
          /* @__PURE__ */ jsx("th", { className: "py-2", children: "Item" }),
          /* @__PURE__ */ jsx("th", { children: "SKU" }),
          /* @__PURE__ */ jsx("th", { children: "Current qty" }),
          /* @__PURE__ */ jsx("th", { children: "Threshold" })
        ] }) }),
        /* @__PURE__ */ jsx("tbody", { children: lowStockItems.map((it) => /* @__PURE__ */ jsxs("tr", { className: "border-b border-slate-100 last:border-0", children: [
          /* @__PURE__ */ jsx("td", { className: "py-2 font-medium text-slate-700", children: it.name }),
          /* @__PURE__ */ jsx("td", { className: "text-slate-500", children: it.sku }),
          /* @__PURE__ */ jsx("td", { className: "text-red-600 font-medium", children: stockByItem[it.id] || 0 }),
          /* @__PURE__ */ jsx("td", { className: "text-slate-500", children: it.minThreshold })
        ] }, it.id)) })
      ] })
    ] })
  ] });
}
function StatCard({ label, value, icon: Icon, tone = "slate", onClick }) {
  const tones = { slate: "text-slate-900", red: "text-red-600" };
  return /* @__PURE__ */ jsxs("button", { onClick, className: cx("flex flex-col gap-2 rounded-lg border border-slate-200 bg-white p-4 text-left", onClick && "hover:border-slate-300 cursor-pointer"), children: [
    /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between", children: [
      /* @__PURE__ */ jsx("span", { className: "text-xs font-medium text-slate-500", children: label }),
      /* @__PURE__ */ jsx(Icon, { size: 16, className: "text-slate-400" })
    ] }),
    /* @__PURE__ */ jsx("span", { className: cx("text-2xl font-semibold", tones[tone]), children: value })
  ] });
}
function Toolbar({ title, subtitle, search, setSearch, children }) {
  return /* @__PURE__ */ jsxs("div", { className: "mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between", children: [
    /* @__PURE__ */ jsxs("div", { children: [
      /* @__PURE__ */ jsx("h1", { className: "text-xl font-semibold text-slate-900", children: title }),
      subtitle && /* @__PURE__ */ jsx("p", { className: "text-sm text-slate-500", children: subtitle })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "flex flex-wrap items-center gap-2", children: [
      setSearch && /* @__PURE__ */ jsxs("div", { className: "relative w-full sm:w-auto", children: [
        /* @__PURE__ */ jsx(Search, { size: 14, className: "absolute left-2.5 top-2.5 text-slate-400" }),
        /* @__PURE__ */ jsx(TextInput, { placeholder: "Search...", value: search, onChange: (e) => setSearch(e.target.value), className: "w-full pl-8 sm:w-56" })
      ] }),
      children
    ] })
  ] });
}
function InventoryTab({ data, update }) {
  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [txItem, setTxItem] = useState(null);
  const [expanded, setExpanded] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const stockByItem = useMemo(() => {
    const map = {};
    data.stockLevels.forEach((s) => {
      map[s.itemId] = (map[s.itemId] || 0) + s.quantity;
    });
    return map;
  }, [data.stockLevels]);
  const filtered = data.items.filter((it) => {
    const matchesSearch = !search || [it.name, it.sku, it.itemCode, it.serialNumber].join(" ").toLowerCase().includes(search.toLowerCase());
    const matchesCat = !categoryFilter || it.categoryId === categoryFilter;
    return matchesSearch && matchesCat;
  });
  const categoryName = (id) => data.categories.find((c) => c.id === id)?.name || "-";
  function saveItem(item, openingStock, stockEdits) {
    update((prev) => {
      const exists = prev.items.some((i) => i.id === item.id);
      const items = exists ? prev.items.map((i) => i.id === item.id ? item : i) : [...prev.items, item];
      let stockLevels = [...prev.stockLevels];
      let transactions = [...prev.transactions];
      if (openingStock) {
        const idx = stockLevels.findIndex((s) => s.itemId === item.id && s.hospitalId === openingStock.warehouseId);
        if (idx >= 0) stockLevels[idx] = { ...stockLevels[idx], quantity: stockLevels[idx].quantity + openingStock.qty };
        else stockLevels.push({ id: uid("stk"), itemId: item.id, hospitalId: openingStock.warehouseId, quantity: openingStock.qty });
        transactions = [{
          id: uid("txn"),
          itemId: item.id,
          itemName: item.name,
          type: "RECEIVE",
          qty: openingStock.qty,
          fromHospitalId: null,
          toHospitalId: openingStock.warehouseId,
          date: todayISO(),
          remarks: "Opening stock recorded with item creation"
        }, ...transactions];
      }
      if (stockEdits) {
        Object.entries(stockEdits).forEach(([hospitalId, newQty]) => {
          const idx = stockLevels.findIndex((s) => s.itemId === item.id && s.hospitalId === hospitalId);
          const currentQty = idx >= 0 ? stockLevels[idx].quantity : 0;
          const delta = Number(newQty) - currentQty;
          if (delta === 0) return;
          if (idx >= 0) stockLevels[idx] = { ...stockLevels[idx], quantity: Number(newQty) };
          else stockLevels.push({ id: uid("stk"), itemId: item.id, hospitalId, quantity: Number(newQty) });
          transactions = [{
            id: uid("txn"),
            itemId: item.id,
            itemName: item.name,
            type: "ADJUST",
            qty: delta,
            fromHospitalId: null,
            toHospitalId: hospitalId,
            date: todayISO(),
            remarks: "Quantity edited directly from Inventory Items"
          }, ...transactions];
        });
      }
      return { ...prev, items, stockLevels, transactions };
    });
    setShowForm(false);
    setEditingItem(null);
  }
  function deleteItem(id) {
    update((prev) => ({
      ...prev,
      items: prev.items.filter((i) => i.id !== id),
      stockLevels: prev.stockLevels.filter((s) => s.itemId !== id)
    }));
  }
  function recordTransaction(tx) {
    update((prev) => {
      let stockLevels = [...prev.stockLevels];
      const applyDelta = (itemId, hospitalId, delta) => {
        const idx = stockLevels.findIndex((s) => s.itemId === itemId && s.hospitalId === hospitalId);
        if (idx >= 0) stockLevels[idx] = { ...stockLevels[idx], quantity: stockLevels[idx].quantity + delta };
        else stockLevels.push({ id: uid("stk"), itemId, hospitalId, quantity: delta });
      };
      if (tx.type === "TRANSFER") {
        applyDelta(tx.itemId, tx.fromHospitalId, -tx.qty);
        applyDelta(tx.itemId, tx.toHospitalId, tx.qty);
      } else if (tx.type === "ADJUST") {
        applyDelta(tx.itemId, tx.toHospitalId, tx.qty);
      } else if (["RECEIVE", "RETURN"].includes(tx.type)) {
        applyDelta(tx.itemId, tx.toHospitalId, tx.qty);
      } else {
        applyDelta(tx.itemId, tx.fromHospitalId, -tx.qty);
      }
      const record = { ...tx, id: uid("txn"), date: tx.date || todayISO() };
      return { ...prev, stockLevels, transactions: [record, ...prev.transactions] };
    });
    setTxItem(null);
  }
  return /* @__PURE__ */ jsxs("div", { children: [
    /* @__PURE__ */ jsxs(Toolbar, { title: "Inventory items", subtitle: "Stock & supplies catalog (RDT kits, printers, scanners, etc). For individually tracked tablets/devices with IMEI, use the 'Tablet Assets' tab instead.", search, setSearch, children: [
      /* @__PURE__ */ jsxs(Select, { value: categoryFilter, onChange: (e) => setCategoryFilter(e.target.value), className: "w-full sm:w-48", children: [
        /* @__PURE__ */ jsx("option", { value: "", children: "All categories" }),
        data.categories.map((c) => /* @__PURE__ */ jsx("option", { value: c.id, children: c.name }, c.id))
      ] }),
      /* @__PURE__ */ jsxs(Btn, { variant: "primary", onClick: () => {
        setEditingItem(null);
        setShowForm(true);
      }, children: [
        /* @__PURE__ */ jsx(Plus, { size: 15 }),
        " Add item"
      ] })
    ] }),
    data.items.length === 0 ? /* @__PURE__ */ jsx(EmptyState, { icon: Package, title: "No items yet", subtitle: "Add your first inventory item to get started." }) : /* @__PURE__ */ jsx("div", { className: "overflow-x-auto rounded-lg border border-slate-200 bg-white", children: /* @__PURE__ */ jsxs("table", { className: "w-full min-w-[720px] text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-slate-50 text-left text-xs uppercase text-slate-400", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "w-6 py-2.5 px-4" }),
        /* @__PURE__ */ jsx("th", { className: "py-2.5", children: "Item" }),
        /* @__PURE__ */ jsx("th", { children: "SKU / Code" }),
        /* @__PURE__ */ jsx("th", { children: "Category" }),
        /* @__PURE__ */ jsx("th", { children: "Unit" }),
        /* @__PURE__ */ jsx("th", { children: "Total qty" }),
        /* @__PURE__ */ jsx("th", { className: "text-right pr-4", children: "Actions" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: filtered.map((it) => {
        const qty = stockByItem[it.id] || 0;
        const low = qty < (it.minThreshold ?? 0);
        const isOpen = expanded === it.id;
        return /* @__PURE__ */ jsxs(Fragment, { children: [
          /* @__PURE__ */ jsxs("tr", { className: "border-t border-slate-100 hover:bg-slate-50", children: [
            /* @__PURE__ */ jsx("td", { className: "pl-4", children: /* @__PURE__ */ jsx("button", { onClick: () => setExpanded(isOpen ? null : it.id), className: "text-slate-400 hover:text-slate-700", children: /* @__PURE__ */ jsx(ChevronDown, { size: 15, className: cx("transition-transform", isOpen && "rotate-180") }) }) }),
            /* @__PURE__ */ jsxs("td", { className: "py-2.5", children: [
              /* @__PURE__ */ jsx("p", { className: "font-medium text-slate-800", children: it.name }),
              it.serialNumber && /* @__PURE__ */ jsxs("p", { className: "text-xs text-slate-400", children: [
                "S/N: ",
                it.serialNumber
              ] })
            ] }),
            /* @__PURE__ */ jsxs("td", { className: "text-slate-500", children: [
              it.sku,
              " ",
              it.itemCode && /* @__PURE__ */ jsxs("span", { className: "text-slate-300", children: [
                "/ ",
                it.itemCode
              ] })
            ] }),
            /* @__PURE__ */ jsx("td", { children: /* @__PURE__ */ jsx(Badge, { children: categoryName(it.categoryId) }) }),
            /* @__PURE__ */ jsx("td", { className: "text-slate-500", children: it.unit }),
            /* @__PURE__ */ jsxs("td", { className: cx("font-medium", low ? "text-red-600" : "text-slate-700"), children: [
              qty,
              " ",
              low && /* @__PURE__ */ jsx(AlertTriangle, { size: 13, className: "inline ml-1" })
            ] }),
            /* @__PURE__ */ jsx("td", { className: "pr-4", children: /* @__PURE__ */ jsxs("div", { className: "flex justify-end gap-1", children: [
              /* @__PURE__ */ jsx(Btn, { className: "py-1 px-2 text-xs", onClick: () => setTxItem(it), children: "Record txn" }),
              /* @__PURE__ */ jsx("button", { title: "Edit item", className: "rounded p-1.5 text-slate-400 hover:bg-slate-200 hover:text-slate-700", onClick: () => {
                setEditingItem(it);
                setShowForm(true);
              }, children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
              /* @__PURE__ */ jsx("button", { title: "Delete item", className: "rounded p-1.5 text-slate-400 hover:bg-red-100 hover:text-red-600", onClick: () => setConfirmDeleteId(it.id), children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
            ] }) })
          ] }),
          isOpen && /* @__PURE__ */ jsx("tr", { className: "bg-slate-50/70 border-t border-slate-100", children: /* @__PURE__ */ jsxs("td", { colSpan: 7, className: "px-4 py-3", children: [
            /* @__PURE__ */ jsx("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: "Stock by warehouse / hospital" }),
            /* @__PURE__ */ jsxs("div", { className: "flex flex-wrap gap-2", children: [
              data.hospitals.map((w) => {
                const lvl = data.stockLevels.find((s) => s.itemId === it.id && s.hospitalId === w.id);
                if (!lvl || lvl.quantity === 0) return null;
                return /* @__PURE__ */ jsxs(Badge, { tone: "blue", children: [
                  w.name,
                  ": ",
                  lvl.quantity
                ] }, w.id);
              }),
              qty === 0 && /* @__PURE__ */ jsx("span", { className: "text-xs text-slate-400", children: "No stock recorded anywhere yet." })
            ] })
          ] }) })
        ] }, it.id);
      }) })
    ] }) }),
    showForm && /* @__PURE__ */ jsx(ItemFormModal, { item: editingItem, categories: data.categories, units: data.units, hospitals: data.hospitals, stockLevels: data.stockLevels, onSave: saveItem, onClose: () => {
      setShowForm(false);
      setEditingItem(null);
    } }),
    txItem && /* @__PURE__ */ jsx(TransactionModal, { item: txItem, hospitals: data.hospitals, onSave: recordTransaction, onClose: () => setTxItem(null) }),
    confirmDeleteId && /* @__PURE__ */ jsx(
      ConfirmDialog,
      {
        message: "Delete this item? Related stock levels will also be removed. This cannot be undone.",
        onCancel: () => setConfirmDeleteId(null),
        onConfirm: () => {
          deleteItem(confirmDeleteId);
          setConfirmDeleteId(null);
        }
      }
    )
  ] });
}
function ItemFormModal({ item, categories, units, hospitals, stockLevels, onSave, onClose }) {
  const [form, setForm] = useState(item || {
    id: uid("item"),
    sku: "",
    itemCode: "",
    name: "",
    categoryId: categories[0]?.id || "",
    unit: units[0] || "",
    serialNumber: "",
    minThreshold: 0,
    image: ""
  });
  const [openingQty, setOpeningQty] = useState(0);
  const [openingWarehouseId, setOpeningWarehouseId] = useState(hospitals[0]?.id || "");
  const [showAdvanced, setShowAdvanced] = useState(false);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const isNew = !item;
  const [stockEdits, setStockEdits] = useState(() => {
    if (isNew) return {};
    const map = {};
    hospitals.forEach((h) => {
      const lvl = (stockLevels || []).find((s) => s.itemId === item.id && s.hospitalId === h.id);
      map[h.id] = lvl ? lvl.quantity : 0;
    });
    return map;
  });
  const setStockFor = (hospitalId, val) => setStockEdits((prev) => ({ ...prev, [hospitalId]: val }));
  return /* @__PURE__ */ jsxs(Modal, { title: item ? "Edit item" : "Add inventory item", onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Item name", required: true, className: "col-span-2", children: /* @__PURE__ */ jsx(TextInput, { value: form.name, onChange: set("name"), placeholder: "e.g. HCV RDT Kit", autoFocus: true }) }),
      /* @__PURE__ */ jsx(Field, { label: "Item code", required: true, children: /* @__PURE__ */ jsx(TextInput, { value: form.itemCode, onChange: set("itemCode"), placeholder: "e.g. RDT-001" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Category", required: true, children: /* @__PURE__ */ jsx(Select, { value: form.categoryId, onChange: set("categoryId"), children: categories.map((c) => /* @__PURE__ */ jsx("option", { value: c.id, children: c.name }, c.id)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Unit", required: true, className: "col-span-2", children: /* @__PURE__ */ jsx(Select, { value: form.unit, onChange: set("unit"), children: units.map((u) => /* @__PURE__ */ jsx("option", { value: u, children: u }, u)) }) })
    ] }),
    isNew && /* @__PURE__ */ jsxs("div", { className: "col-span-2 mt-3 rounded-md border border-slate-200 bg-slate-50 p-3", children: [
      /* @__PURE__ */ jsx("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: "Opening stock (optional)" }),
      hospitals.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-amber-600", children: 'No warehouse or hospital configured yet. Add one from "Hospitals & Warehouses" first, or leave this blank and use "Record txn" later.' }) : /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
        /* @__PURE__ */ jsx(Field, { label: "Quantity", children: /* @__PURE__ */ jsx(TextInput, { type: "number", min: 0, value: openingQty, onChange: (e) => setOpeningQty(Number(e.target.value)) }) }),
        /* @__PURE__ */ jsx(Field, { label: "Into warehouse / hospital", children: /* @__PURE__ */ jsx(Select, { value: openingWarehouseId, onChange: (e) => setOpeningWarehouseId(e.target.value), children: hospitals.map((h) => /* @__PURE__ */ jsx("option", { value: h.id, children: h.name }, h.id)) }) })
      ] })
    ] }),
    !isNew && /* @__PURE__ */ jsxs("div", { className: "mt-3 rounded-md border border-slate-200 bg-slate-50 p-3", children: [
      /* @__PURE__ */ jsx("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: "Stock by warehouse" }),
      hospitals.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-amber-600", children: "No warehouse or hospital configured yet." }) : /* @__PURE__ */ jsx("div", { className: "flex flex-col gap-2 max-h-48 overflow-y-auto pr-1", children: hospitals.map((h) => /* @__PURE__ */ jsxs("div", { className: "flex items-center justify-between gap-3", children: [
        /* @__PURE__ */ jsx("span", { className: "text-sm text-slate-700", children: h.name }),
        /* @__PURE__ */ jsx(
          TextInput,
          {
            type: "number",
            min: 0,
            value: stockEdits[h.id] ?? 0,
            onChange: (e) => setStockFor(h.id, Number(e.target.value)),
            className: "w-28 text-right"
          }
        )
      ] }, h.id)) }),
      /* @__PURE__ */ jsx("p", { className: "mt-2 text-xs text-slate-400", children: 'Changing a number here directly adjusts stock and records an "Adjust Stock" transaction automatically.' })
    ] }),
    /* @__PURE__ */ jsxs("button", { type: "button", onClick: () => setShowAdvanced((v) => !v), className: "mt-3 flex items-center gap-1 text-xs font-medium text-slate-500 hover:text-slate-700", children: [
      /* @__PURE__ */ jsx(ChevronDown, { size: 13, className: cx("transition-transform", showAdvanced && "rotate-180") }),
      showAdvanced ? "Hide" : "Show",
      " more fields (SKU, serial number, threshold, image)"
    ] }),
    showAdvanced && /* @__PURE__ */ jsxs("div", { className: "mt-2 grid grid-cols-2 gap-3 rounded-md border border-slate-200 p-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "SKU", children: /* @__PURE__ */ jsx(TextInput, { value: form.sku, onChange: set("sku"), placeholder: "Defaults to item code" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Serial number", children: /* @__PURE__ */ jsx(TextInput, { value: form.serialNumber, onChange: set("serialNumber"), placeholder: "If applicable" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Minimum stock threshold", children: /* @__PURE__ */ jsx(TextInput, { type: "number", min: 0, value: form.minThreshold, onChange: (e) => setForm((f) => ({ ...f, minThreshold: Number(e.target.value) })) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Image URL", children: /* @__PURE__ */ jsx(TextInput, { value: form.image, onChange: set("image"), placeholder: "https://..." }) })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mt-5 flex justify-end gap-2", children: [
      /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
      /* @__PURE__ */ jsx(
        Btn,
        {
          variant: "primary",
          disabled: !form.name || !form.itemCode || !form.categoryId,
          onClick: () => onSave(
            { ...form, sku: form.sku || form.itemCode },
            isNew && openingQty > 0 ? { qty: openingQty, warehouseId: openingWarehouseId } : null,
            !isNew ? stockEdits : null
          ),
          children: "Save item"
        }
      )
    ] })
  ] });
}
function TransactionModal({ item, hospitals, onSave, onClose }) {
  const [type, setType] = useState("RECEIVE");
  const [qty, setQty] = useState(1);
  const [fromHospitalId, setFromHospitalId] = useState(hospitals[0]?.id || "");
  const [toHospitalId, setToHospitalId] = useState(hospitals[0]?.id || "");
  const [date, setDate] = useState(todayISO());
  const [remarks, setRemarks] = useState("");
  const needsFrom = ["TRANSFER", "ISSUE", "DAMAGED", "LOST", "DISPOSED"].includes(type);
  const needsTo = ["TRANSFER", "RECEIVE", "RETURN", "ADJUST"].includes(type);
  return /* @__PURE__ */ jsxs(Modal, { title: `Record transaction - ${item.name}`, onClose, children: [
    hospitals.length === 0 && /* @__PURE__ */ jsx("div", { className: "mb-3 rounded-md bg-amber-50 px-3 py-2 text-xs text-amber-700", children: 'No warehouse or hospital configured yet. Add one from "Hospitals & Warehouses" before recording a transaction.' }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Transaction type", required: true, className: "col-span-2", children: /* @__PURE__ */ jsx(Select, { value: type, onChange: (e) => setType(e.target.value), children: TX_TYPES.map((t) => /* @__PURE__ */ jsx("option", { value: t.id, children: t.label }, t.id)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Quantity", required: true, children: /* @__PURE__ */ jsx(TextInput, { type: "number", min: 1, value: qty, onChange: (e) => setQty(Number(e.target.value)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Date", children: /* @__PURE__ */ jsx(TextInput, { type: "date", value: date, onChange: (e) => setDate(e.target.value) }) }),
      needsFrom && /* @__PURE__ */ jsx(Field, { label: "From warehouse / hospital", required: true, className: "col-span-2", children: /* @__PURE__ */ jsx(Select, { value: fromHospitalId, onChange: (e) => setFromHospitalId(e.target.value), children: hospitals.map((h) => /* @__PURE__ */ jsx("option", { value: h.id, children: h.name }, h.id)) }) }),
      needsTo && /* @__PURE__ */ jsx(Field, { label: "To warehouse / hospital", required: true, className: "col-span-2", children: /* @__PURE__ */ jsx(Select, { value: toHospitalId, onChange: (e) => setToHospitalId(e.target.value), children: hospitals.map((h) => /* @__PURE__ */ jsx("option", { value: h.id, children: h.name }, h.id)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Remarks", className: "col-span-2", children: /* @__PURE__ */ jsx(TextArea, { rows: 2, value: remarks, onChange: (e) => setRemarks(e.target.value), placeholder: "Optional notes" }) })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mt-5 flex justify-end gap-2", children: [
      /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
      /* @__PURE__ */ jsx(Btn, { variant: "primary", disabled: !qty || qty <= 0 || hospitals.length === 0, onClick: () => onSave({ itemId: item.id, itemName: item.name, type, qty: Number(qty), fromHospitalId: needsFrom ? fromHospitalId : null, toHospitalId: needsTo ? toHospitalId : null, date, remarks }), children: "Save transaction" })
    ] })
  ] });
}
function TransactionsTab({ data }) {
  const [search, setSearch] = useState("");
  const [typeFilter, setTypeFilter] = useState("");
  const hospitalName = (id) => data.hospitals.find((h) => h.id === id)?.name || "-";
  const filtered = data.transactions.filter((tx) => {
    const matchesSearch = !search || tx.itemName.toLowerCase().includes(search.toLowerCase());
    const matchesType = !typeFilter || tx.type === typeFilter;
    return matchesSearch && matchesType;
  });
  return /* @__PURE__ */ jsxs("div", { children: [
    /* @__PURE__ */ jsx(Toolbar, { title: "Stock transactions", subtitle: "Full audit trail of every stock movement.", search, setSearch, children: /* @__PURE__ */ jsxs(Select, { value: typeFilter, onChange: (e) => setTypeFilter(e.target.value), className: "w-full sm:w-48", children: [
      /* @__PURE__ */ jsx("option", { value: "", children: "All types" }),
      TX_TYPES.map((t) => /* @__PURE__ */ jsx("option", { value: t.id, children: t.label }, t.id))
    ] }) }),
    filtered.length === 0 ? /* @__PURE__ */ jsx(EmptyState, { icon: ArrowLeftRight, title: "No transactions recorded", subtitle: "Record a transaction from the Inventory Items tab." }) : /* @__PURE__ */ jsx("div", { className: "overflow-x-auto rounded-lg border border-slate-200 bg-white", children: /* @__PURE__ */ jsxs("table", { className: "w-full min-w-[720px] text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-slate-50 text-left text-xs uppercase text-slate-400", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "py-2.5 pl-4", children: "Date" }),
        /* @__PURE__ */ jsx("th", { children: "Item" }),
        /* @__PURE__ */ jsx("th", { children: "Type" }),
        /* @__PURE__ */ jsx("th", { children: "Qty" }),
        /* @__PURE__ */ jsx("th", { children: "From" }),
        /* @__PURE__ */ jsx("th", { children: "To" }),
        /* @__PURE__ */ jsx("th", { className: "pr-4", children: "Remarks" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: filtered.map((tx) => {
        const meta = TX_TYPES.find((t) => t.id === tx.type);
        const Icon = meta?.icon || RefreshCcw;
        return /* @__PURE__ */ jsxs("tr", { className: "border-t border-slate-100", children: [
          /* @__PURE__ */ jsx("td", { className: "py-2.5 pl-4 text-slate-500", children: fmtDate(tx.date) }),
          /* @__PURE__ */ jsx("td", { className: "font-medium text-slate-700", children: tx.itemName }),
          /* @__PURE__ */ jsx("td", { children: /* @__PURE__ */ jsxs("span", { className: "inline-flex items-center gap-1.5 text-xs font-medium", style: { color: meta?.color }, children: [
            /* @__PURE__ */ jsx(Icon, { size: 14 }),
            " ",
            meta?.label
          ] }) }),
          /* @__PURE__ */ jsx("td", { className: "text-slate-700", children: tx.qty }),
          /* @__PURE__ */ jsx("td", { className: "text-slate-500", children: tx.fromHospitalId ? hospitalName(tx.fromHospitalId) : "-" }),
          /* @__PURE__ */ jsx("td", { className: "text-slate-500", children: tx.toHospitalId ? hospitalName(tx.toHospitalId) : "-" }),
          /* @__PURE__ */ jsx("td", { className: "pr-4 text-slate-400", children: tx.remarks || "-" })
        ] }, tx.id);
      }) })
    ] }) })
  ] });
}
function WarehousesTab({ data, update }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  function save(w) {
    update((prev) => {
      const exists = prev.hospitals.some((h) => h.id === w.id);
      return { ...prev, hospitals: exists ? prev.hospitals.map((h) => h.id === w.id ? w : h) : [...prev.hospitals, w] };
    });
    setShowForm(false);
    setEditing(null);
  }
  function remove(id) {
    update((prev) => ({ ...prev, hospitals: prev.hospitals.filter((h) => h.id !== id) }));
  }
  return /* @__PURE__ */ jsxs("div", { children: [
    /* @__PURE__ */ jsx(Toolbar, { title: "Hospitals & warehouses", subtitle: "Configure the warehouse hierarchy - Central, Regional, and Hospital level.", children: /* @__PURE__ */ jsxs(Btn, { variant: "primary", onClick: () => {
      setEditing(null);
      setShowForm(true);
    }, children: [
      /* @__PURE__ */ jsx(Plus, { size: 15 }),
      " Add location"
    ] }) }),
    data.hospitals.length === 0 ? /* @__PURE__ */ jsx(EmptyState, { icon: Building2, title: "No warehouses or hospitals configured" }) : /* @__PURE__ */ jsx("div", { className: "grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3", children: data.hospitals.map((w) => /* @__PURE__ */ jsxs("div", { className: "rounded-lg border border-slate-200 bg-white p-4", children: [
      /* @__PURE__ */ jsxs("div", { className: "mb-2 flex items-start justify-between", children: [
        /* @__PURE__ */ jsx(Badge, { tone: w.type === "Central Warehouse" ? "blue" : w.type === "Regional Warehouse" ? "amber" : "green", children: w.type }),
        /* @__PURE__ */ jsxs("div", { className: "flex gap-1", children: [
          /* @__PURE__ */ jsx("button", { title: "Edit", className: "rounded p-1 text-slate-400 hover:bg-slate-200 hover:text-slate-700", onClick: () => {
            setEditing(w);
            setShowForm(true);
          }, children: /* @__PURE__ */ jsx(Pencil, { size: 13 }) }),
          /* @__PURE__ */ jsx("button", { title: "Delete", className: "rounded p-1 text-slate-400 hover:bg-red-100 hover:text-red-600", onClick: () => setConfirmDeleteId(w.id), children: /* @__PURE__ */ jsx(Trash2, { size: 13 }) })
        ] })
      ] }),
      /* @__PURE__ */ jsx("p", { className: "font-medium text-slate-800", children: w.name }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-slate-400", children: w.location })
    ] }, w.id)) }),
    showForm && /* @__PURE__ */ jsx(WarehouseFormModal, { warehouse: editing, onSave: save, onClose: () => {
      setShowForm(false);
      setEditing(null);
    } }),
    confirmDeleteId && /* @__PURE__ */ jsx(
      ConfirmDialog,
      {
        message: "Delete this warehouse / hospital? This cannot be undone.",
        onCancel: () => setConfirmDeleteId(null),
        onConfirm: () => {
          remove(confirmDeleteId);
          setConfirmDeleteId(null);
        }
      }
    )
  ] });
}
function WarehouseFormModal({ warehouse, onSave, onClose }) {
  const [form, setForm] = useState(warehouse || { id: uid("wh"), name: "", type: WAREHOUSE_TYPES[0], location: "" });
  return /* @__PURE__ */ jsxs(Modal, { title: warehouse ? "Edit location" : "Add warehouse / hospital", onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "flex flex-col gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Name", required: true, children: /* @__PURE__ */ jsx(TextInput, { value: form.name, onChange: (e) => setForm((f) => ({ ...f, name: e.target.value })), placeholder: "e.g. DHQ Hospital Multan" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Type", required: true, children: /* @__PURE__ */ jsx(Select, { value: form.type, onChange: (e) => setForm((f) => ({ ...f, type: e.target.value })), children: WAREHOUSE_TYPES.map((t) => /* @__PURE__ */ jsx("option", { value: t, children: t }, t)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Location / city", children: /* @__PURE__ */ jsx(TextInput, { value: form.location, onChange: (e) => setForm((f) => ({ ...f, location: e.target.value })) }) })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mt-5 flex justify-end gap-2", children: [
      /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
      /* @__PURE__ */ jsx(Btn, { variant: "primary", disabled: !form.name, onClick: () => onSave(form), children: "Save" })
    ] })
  ] });
}
var CUSTODY_TYPES = {
  TAKEOVER_SOURCE: { label: "Takeover (from source)", icon: Inbox, color: "#0F6E56", tone: "green" },
  HANDOVER: { label: "Handover (to hospital)", icon: Send, color: "#185FA5", tone: "blue" },
  TAKEOVER_HOSPITAL: { label: "Takeover (from hospital)", icon: ClipboardCheck, color: "#854F0B", tone: "amber" },
  REPLACED_OUT: { label: "Replaced - removed from hospital", icon: RefreshCcw, color: "#A32D2D", tone: "red" },
  REPLACED_IN: { label: "Replaced - issued to hospital", icon: RefreshCcw, color: "#0F6E56", tone: "green" },
  ADDED: { label: "Record created", icon: Plus, color: "#5F5E5A", tone: "slate" },
  EDITED: { label: "Details edited", icon: Pencil, color: "#5F5E5A", tone: "slate" }
};
function custodyMeta(type) {
  return CUSTODY_TYPES[type] || { label: type || "Event", icon: History, color: "#5F5E5A", tone: "slate" };
}
function TabletsTab({ data, update }) {
  const [view, setView] = useState("devices");
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [takeoverSourceTablet, setTakeoverSourceTablet] = useState(null);
  const [handoverTablet, setHandoverTablet] = useState(null);
  const [takeoverHospitalTablet, setTakeoverHospitalTablet] = useState(null);
  const [showLetterModal, setShowLetterModal] = useState(false);
  const [showReplacementModal, setShowReplacementModal] = useState(false);
  const [expanded, setExpanded] = useState(null);
  const [importSummary, setImportSummary] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [selectedForHandover, setSelectedForHandover] = useState(/* @__PURE__ */ new Set());
  const fileInputRef = useRef(null);
  const toggleSelectForHandover = (id) => setSelectedForHandover((prev) => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    return next;
  });
  const hospitalName = (id) => data.hospitals.find((h) => h.id === id)?.name || "-";
  const staffName = (id) => data.staff.find((s) => s.id === id)?.name || "-";
  const filtered = data.tablets.filter((t) => !search || [t.serialNumber, t.model].join(" ").toLowerCase().includes(search.toLowerCase()));
  function processImportRows(rows) {
    const warnings = [];
    let created = 0, updated = 0;
    const KNOWN_FIELD_ALIASES = {
      imei: ["imei", "imei no", "imei number", "imei/serial no", "serial number", "serial no", "serial", "device serial"],
      appInstalled: ["app installed", "nadra app", "nadra app installed", "app"],
      status: ["status", "device status", "condition"],
      hospitalName: ["hospital name", "hospital", "facility", "facility name"],
      handedTo: ["handed over to person", "handed over to", "handed to", "assigned staff", "assigned to", "taken over by", "person"]
    };
    update((prev) => {
      let tablets = [...prev.tablets];
      rows.forEach((row, idx) => {
        const rowKeys = Object.keys(row);
        const consumedKeys = /* @__PURE__ */ new Set();
        const findKey = (aliases) => {
          let found = rowKeys.find((rk) => aliases.includes(rk.trim().toLowerCase()));
          if (!found) found = rowKeys.find((rk) => aliases.some((a) => rk.trim().toLowerCase().includes(a)));
          return found;
        };
        const get = (fieldKey) => {
          const found = findKey(KNOWN_FIELD_ALIASES[fieldKey]);
          if (!found) return "";
          consumedKeys.add(found);
          const val = row[found];
          return val === void 0 || val === null ? "" : String(val).trim();
        };
        const imei = get("imei");
        if (!imei) {
          warnings.push(`Row ${idx + 2}: missing IMEI/Serial column value - skipped.`);
          return;
        }
        const appInstalledRaw = get("appInstalled").toLowerCase();
        const nadraInstalled = ["yes", "true", "installed", "1", "y"].includes(appInstalledRaw);
        const statusRaw = get("status");
        const hospitalNameRaw = get("hospitalName");
        const handedToRaw = get("handedTo");
        const hospital = hospitalNameRaw ? prev.hospitals.find((h) => h.name.toLowerCase() === hospitalNameRaw.toLowerCase()) : null;
        if (hospitalNameRaw && !hospital) warnings.push(`Row ${idx + 2}: hospital "${hospitalNameRaw}" not found - left unassigned.`);
        let assignedStaffId = null, assignedStaffName = null;
        if (handedToRaw) {
          const matchedStaff = prev.staff.find((s) => s.name.toLowerCase() === handedToRaw.toLowerCase() && (!hospital || s.hospitalId === hospital.id));
          if (matchedStaff) assignedStaffId = matchedStaff.id;
          else assignedStaffName = handedToRaw;
        }
        const validStatus = TABLET_STATUSES.find((s) => s.toLowerCase() === statusRaw.toLowerCase());
        const status = validStatus || (handedToRaw ? "Assigned" : "In Stock");
        const customFields = {};
        rowKeys.forEach((k) => {
          if (consumedKeys.has(k)) return;
          const val = row[k];
          if (val !== void 0 && val !== null && String(val).trim() !== "") customFields[k.trim()] = String(val).trim();
        });
        const existingIdx = tablets.findIndex((t) => t.serialNumber.trim().toLowerCase() === imei.toLowerCase());
        const extraNote = Object.keys(customFields).length ? ` Extra columns: ${Object.entries(customFields).map(([k, v]) => `${k}=${v}`).join(", ")}.` : "";
        const historyEntry = {
          date: todayISO(),
          type: existingIdx >= 0 ? "EDITED" : "ADDED",
          detail: `Imported from Excel.${hospital ? ` Hospital: ${hospital.name}.` : ""}${handedToRaw ? ` Handed to: ${handedToRaw}.` : ""} Status: ${status}.${extraNote}`
        };
        if (existingIdx >= 0) {
          tablets[existingIdx] = {
            ...tablets[existingIdx],
            nadraInstalled,
            status,
            currentHospitalId: hospital?.id || tablets[existingIdx].currentHospitalId,
            assignedStaffId,
            assignedStaffName,
            customFields: { ...tablets[existingIdx].customFields || {}, ...customFields },
            history: [historyEntry, ...tablets[existingIdx].history || []]
          };
          updated++;
        } else {
          tablets.push({
            id: uid("tab"),
            serialNumber: imei,
            model: "IVAS Tablet",
            currentHospitalId: hospital?.id || "",
            assignedStaffId,
            assignedStaffName,
            nadraInstalled,
            batteryHealth: 100,
            status,
            remarks: "",
            customFields,
            history: [historyEntry]
          });
          created++;
        }
      });
      return { ...prev, tablets };
    });
    setImportSummary({ created, updated, warnings });
  }
  async function handleImportFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (rows.length === 0) {
        setImportSummary({ created: 0, updated: 0, warnings: ["The file has no data rows."] });
        return;
      }
      processImportRows(rows);
    } catch (err) {
      setImportSummary({ created: 0, updated: 0, warnings: ["Could not read this file. Please upload a valid .xlsx, .xls, or .csv file."] });
    } finally {
      e.target.value = "";
    }
  }
  function save(t) {
    update((prev) => {
      const exists = prev.tablets.some((x) => x.id === t.id);
      const history = exists ? [{ date: todayISO(), type: "EDITED", detail: "Device details updated" }, ...t.history || []] : [{ date: todayISO(), type: "ADDED", detail: "Device record created" }];
      return {
        ...prev,
        tablets: exists ? prev.tablets.map((x) => x.id === t.id ? { ...t, history } : x) : [...prev.tablets, { ...t, history }]
      };
    });
    setShowForm(false);
    setEditing(null);
  }
  function remove(id) {
    update((prev) => ({ ...prev, tablets: prev.tablets.filter((x) => x.id !== id) }));
  }
  function takeoverFromSource(payload) {
    update((prev) => {
      const conditionOk = payload.overallCondition !== "Damaged";
      const historyEntry = {
        date: payload.date,
        type: "TAKEOVER_SOURCE",
        detail: `Received from ${payload.source}${payload.receivedBy ? ` by ${payload.receivedBy}` : ""} into ${hospitalName(payload.warehouseId)}. Condition: ${payload.overallCondition}. Checklist passed: ${payload.checklist.filter(Boolean).length}/${payload.checklistLabels.length}.${payload.remarks ? " Remarks: " + payload.remarks : ""}`
      };
      const existing = prev.tablets.find((t) => t.id === payload.tabletId);
      if (existing) {
        return {
          ...prev,
          tablets: prev.tablets.map((t) => t.id !== payload.tabletId ? t : {
            ...t,
            currentHospitalId: payload.warehouseId,
            assignedStaffId: null,
            status: conditionOk ? "In Stock" : "Damaged",
            history: [historyEntry, ...t.history || []]
          })
        };
      }
      const newTablet = {
        id: uid("tab"),
        serialNumber: payload.serialNumber,
        model: payload.model,
        currentHospitalId: payload.warehouseId,
        assignedStaffId: null,
        nadraInstalled: false,
        batteryHealth: payload.batteryHealth || 100,
        status: conditionOk ? "In Stock" : "Damaged",
        remarks: payload.remarks,
        history: [historyEntry]
      };
      return { ...prev, tablets: [...prev.tablets, newTablet] };
    });
    setTakeoverSourceTablet(null);
  }
  function handover(tablet, payload) {
    update((prev) => ({
      ...prev,
      tablets: prev.tablets.map((t) => t.id !== tablet.id ? t : {
        ...t,
        currentHospitalId: payload.hospitalId,
        assignedStaffId: payload.staffId,
        status: "Assigned",
        history: [{
          date: payload.date,
          type: "HANDOVER",
          detail: `Handed over to ${hospitalName(payload.hospitalId)} / ${staffName(payload.staffId)}${payload.handedOverBy ? ` by ${payload.handedOverBy}` : ""}. Condition: ${payload.condition}. Accessories given: ${payload.accessories.join(", ") || "None"}.${payload.remarks ? " Remarks: " + payload.remarks : ""}`
        }, ...t.history || []]
      })
    }));
    setHandoverTablet(null);
  }
  function takeoverFromHospital(tablet, payload) {
    update((prev) => ({
      ...prev,
      tablets: prev.tablets.map((t) => t.id !== tablet.id ? t : {
        ...t,
        assignedStaffId: null,
        status: payload.condition === "Damaged" ? "Damaged" : "In Stock",
        remarks: payload.remarks,
        history: [{
          date: payload.date,
          type: "TAKEOVER_HOSPITAL",
          detail: `Taken over from ${hospitalName(t.currentHospitalId)}${payload.receivedBy ? ` by ${payload.receivedBy}` : ""}. Condition: ${payload.condition}. Accessories returned: ${payload.accessories.join(", ") || "None"}.${payload.inspectionNotes ? " Inspection: " + payload.inspectionNotes : ""}${payload.remarks ? " Remarks: " + payload.remarks : ""}`
        }, ...t.history || []]
      })
    }));
    setTakeoverHospitalTablet(null);
  }
  function handoverViaLetter(letter) {
    update((prev) => ({
      ...prev,
      tablets: prev.tablets.map((t) => {
        const row = letter.rows.find((r) => r.tabletId === t.id);
        if (!row) return t;
        return {
          ...t,
          currentHospitalId: letter.hospitalId,
          assignedStaffId: letter.takenOverStaffId || null,
          assignedStaffName: letter.takenOverStaffId ? null : letter.takenOverName,
          status: "Assigned",
          history: [{
            date: row.date,
            type: "HANDOVER",
            detail: `Handed over to ${hospitalName(letter.hospitalId)} / ${letter.takenOverName}${letter.takenOverDesignation ? ` (${letter.takenOverDesignation})` : ""} by ${letter.handingOverName}${letter.handingOverDesignation ? ` (${letter.handingOverDesignation})` : ""}. Counter: ${row.counter || "-"}. Status: ${row.status || "-"}.${letter.refNo ? ` Ref: ${letter.refNo}.` : ""}`
          }, ...t.history || []]
        };
      })
    }));
  }
  function replaceViaLetter(letter) {
    update((prev) => ({
      ...prev,
      tablets: prev.tablets.map((t) => {
        const asOld = letter.rows.find((r) => r.oldTabletId === t.id);
        const asNew = letter.rows.find((r) => r.newTabletId === t.id);
        const newSerial = (id) => prev.tablets.find((x) => x.id === id)?.serialNumber || "-";
        if (asOld) {
          return {
            ...t,
            status: asOld.oldStatus || "Damaged",
            assignedStaffId: null,
            assignedStaffName: null,
            history: [{
              date: asOld.date,
              type: "REPLACED_OUT",
              detail: `Replaced with new device IMEI ${newSerial(asOld.newTabletId)} at ${hospitalName(letter.hospitalId)}. Reason: ${asOld.reason || "-"}. Processed by ${letter.handingOverName}.${letter.refNo ? ` Ref: ${letter.refNo}.` : ""}`
            }, ...t.history || []]
          };
        }
        if (asNew) {
          return {
            ...t,
            currentHospitalId: letter.hospitalId,
            assignedStaffId: letter.takenOverStaffId || null,
            assignedStaffName: letter.takenOverStaffId ? null : letter.takenOverName,
            status: "Assigned",
            history: [{
              date: asNew.date,
              type: "REPLACED_IN",
              detail: `Issued to ${hospitalName(letter.hospitalId)} / ${letter.takenOverName} replacing old device IMEI ${newSerial(asNew.oldTabletId)}. Reason: ${asNew.reason || "-"}. Counter: ${asNew.counter || "-"}.${letter.refNo ? ` Ref: ${letter.refNo}.` : ""}`
            }, ...t.history || []]
          };
        }
        return t;
      })
    }));
  }
  const staffLabel = (t) => t.assignedStaffId ? staffName(t.assignedStaffId) : t.assignedStaffName || "-";
  const allLogEntries = useMemo(() => {
    const rows = [];
    data.tablets.forEach((t) => (t.history || []).forEach((h) => rows.push({ ...h, tablet: t })));
    return rows.sort((a, b) => a.date < b.date ? 1 : -1);
  }, [data.tablets]);
  const filteredLog = allLogEntries.filter((e) => !search || [e.tablet.serialNumber, e.detail].join(" ").toLowerCase().includes(search.toLowerCase()));
  return /* @__PURE__ */ jsxs("div", { children: [
    /* @__PURE__ */ jsxs(
      Toolbar,
      {
        title: "Tablet assets - handover / takeover",
        subtitle: "Full chain of custody: receive from source, inspect, hand over to hospitals, and take back.",
        search,
        setSearch,
        children: [
          /* @__PURE__ */ jsxs(Btn, { variant: "primary", onClick: () => setTakeoverSourceTablet({}), children: [
            /* @__PURE__ */ jsx(Inbox, { size: 15 }),
            " Takeover device"
          ] }),
          /* @__PURE__ */ jsxs(Btn, { onClick: () => setShowLetterModal(true), children: [
            /* @__PURE__ */ jsx(FileText, { size: 15 }),
            " Handover letter"
          ] }),
          /* @__PURE__ */ jsxs(Btn, { onClick: () => setShowReplacementModal(true), children: [
            /* @__PURE__ */ jsx(RefreshCcw, { size: 15 }),
            " Replacement letter"
          ] }),
          /* @__PURE__ */ jsxs(Btn, { onClick: () => fileInputRef.current?.click(), children: [
            /* @__PURE__ */ jsx(Upload, { size: 15 }),
            " Import from Excel"
          ] }),
          /* @__PURE__ */ jsx("input", { ref: fileInputRef, type: "file", accept: ".xlsx,.xls,.csv", className: "hidden", onChange: handleImportFile }),
          /* @__PURE__ */ jsxs(Btn, { onClick: () => {
            setEditing(null);
            setShowForm(true);
          }, children: [
            /* @__PURE__ */ jsx(Plus, { size: 15 }),
            " Add manually"
          ] })
        ]
      }
    ),
    /* @__PURE__ */ jsxs("p", { className: "mb-3 text-xs text-slate-400", children: [
      "Recognized columns: ",
      /* @__PURE__ */ jsx("span", { className: "font-medium text-slate-500", children: "IMEI, App Installed, Status, Hospital Name, Handed Over To Person" }),
      " (column names are matched flexibly, so slight variations are fine). Existing devices are matched and updated by IMEI; new IMEIs create new device records. Any extra columns in your file are kept too and shown under each device's custody history."
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mb-4 flex gap-1 rounded-md bg-slate-100 p-1 w-fit", children: [
      /* @__PURE__ */ jsx("button", { onClick: () => setView("devices"), className: cx("rounded px-3 py-1.5 text-sm font-medium", view === "devices" ? "bg-white shadow-sm text-slate-900" : "text-slate-500"), children: "Devices" }),
      /* @__PURE__ */ jsx("button", { onClick: () => setView("log"), className: cx("rounded px-3 py-1.5 text-sm font-medium", view === "log" ? "bg-white shadow-sm text-slate-900" : "text-slate-500"), children: "Handover / takeover log" })
    ] }),
    importSummary && /* @__PURE__ */ jsxs("div", { className: "mb-4 rounded-md border border-blue-200 bg-blue-50 p-3 text-sm text-blue-800", children: [
      /* @__PURE__ */ jsxs("div", { className: "flex items-start justify-between gap-2", children: [
        /* @__PURE__ */ jsxs("p", { className: "font-medium", children: [
          "Import complete: ",
          importSummary.created,
          " device",
          importSummary.created === 1 ? "" : "s",
          " created, ",
          importSummary.updated,
          " updated."
        ] }),
        /* @__PURE__ */ jsx("button", { onClick: () => setImportSummary(null), className: "shrink-0 text-blue-400 hover:text-blue-700", children: /* @__PURE__ */ jsx(X, { size: 14 }) })
      ] }),
      importSummary.warnings.length > 0 && /* @__PURE__ */ jsx("ul", { className: "mt-1.5 list-disc pl-5 text-xs text-amber-700", children: importSummary.warnings.map((w, i) => /* @__PURE__ */ jsx("li", { children: w }, i)) })
    ] }),
    view === "devices" && selectedForHandover.size > 0 && /* @__PURE__ */ jsxs("div", { className: "mb-4 flex flex-wrap items-center justify-between gap-2 rounded-md border border-slate-900 bg-slate-50 px-3 py-2.5", children: [
      /* @__PURE__ */ jsxs("p", { className: "text-sm font-medium text-slate-800", children: [
        selectedForHandover.size,
        " device",
        selectedForHandover.size === 1 ? "" : "s",
        " selected for handover"
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
        /* @__PURE__ */ jsx(Btn, { className: "py-1.5 px-2.5 text-xs", onClick: () => {
          const inStockIds = filtered.filter((t) => t.status === "In Stock").map((t) => t.id);
          setSelectedForHandover(new Set(inStockIds));
        }, children: "Select all in stock" }),
        /* @__PURE__ */ jsx(Btn, { className: "py-1.5 px-2.5 text-xs", onClick: () => setSelectedForHandover(/* @__PURE__ */ new Set()), children: "Clear selection" }),
        /* @__PURE__ */ jsxs(Btn, { variant: "primary", className: "py-1.5 px-2.5 text-xs", onClick: () => setShowLetterModal(true), children: [
          /* @__PURE__ */ jsx(FileText, { size: 13 }),
          " Generate handover letter"
        ] })
      ] })
    ] }),
    view === "devices" ? data.tablets.length === 0 ? /* @__PURE__ */ jsx(EmptyState, { icon: TabletIcon, title: "No devices registered", subtitle: "Use 'Takeover device' to record a device received from DHO office or another source." }) : /* @__PURE__ */ jsx("div", { className: "overflow-x-auto rounded-lg border border-slate-200 bg-white", children: /* @__PURE__ */ jsxs("table", { className: "w-full min-w-[720px] text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-slate-50 text-left text-xs uppercase text-slate-400", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "w-6 py-2.5 pl-4" }),
        /* @__PURE__ */ jsx("th", { className: "w-6" }),
        /* @__PURE__ */ jsx("th", { children: "Serial number" }),
        /* @__PURE__ */ jsx("th", { children: "Status" }),
        /* @__PURE__ */ jsx("th", { children: "Current location" }),
        /* @__PURE__ */ jsx("th", { children: "Assigned staff" }),
        /* @__PURE__ */ jsx("th", { children: "NADRA app" }),
        /* @__PURE__ */ jsx("th", { children: "Battery" }),
        /* @__PURE__ */ jsx("th", { className: "text-right pr-4", children: "Actions" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: filtered.map((t) => {
        const isOpen = expanded === t.id;
        const eligibleForHandover = t.status === "In Stock";
        return /* @__PURE__ */ jsxs(Fragment, { children: [
          /* @__PURE__ */ jsxs("tr", { className: "border-t border-slate-100 hover:bg-slate-50", children: [
            /* @__PURE__ */ jsx("td", { className: "pl-4", children: eligibleForHandover && /* @__PURE__ */ jsx(
              "input",
              {
                type: "checkbox",
                checked: selectedForHandover.has(t.id),
                onChange: () => toggleSelectForHandover(t.id),
                title: "Select for handover"
              }
            ) }),
            /* @__PURE__ */ jsx("td", { children: /* @__PURE__ */ jsx("button", { onClick: () => setExpanded(isOpen ? null : t.id), className: "text-slate-400 hover:text-slate-700", children: /* @__PURE__ */ jsx(ChevronDown, { size: 15, className: cx("transition-transform", isOpen && "rotate-180") }) }) }),
            /* @__PURE__ */ jsxs("td", { className: "py-2.5 font-medium text-slate-700", children: [
              t.serialNumber,
              /* @__PURE__ */ jsx("p", { className: "text-xs font-normal text-slate-400", children: t.model })
            ] }),
            /* @__PURE__ */ jsx("td", { children: /* @__PURE__ */ jsx(Badge, { tone: t.status === "Assigned" ? "blue" : t.status === "In Stock" ? "green" : t.status === "Damaged" || t.status === "Lost" ? "red" : "amber", children: t.status }) }),
            /* @__PURE__ */ jsx("td", { className: "text-slate-600", children: t.currentHospitalId ? hospitalName(t.currentHospitalId) : "-" }),
            /* @__PURE__ */ jsx("td", { className: "text-slate-600", children: staffLabel(t) }),
            /* @__PURE__ */ jsx("td", { children: t.nadraInstalled ? /* @__PURE__ */ jsx(Badge, { tone: "green", children: "Installed" }) : /* @__PURE__ */ jsx(Badge, { children: "Not installed" }) }),
            /* @__PURE__ */ jsx("td", { className: "text-slate-600", children: t.batteryHealth ? `${t.batteryHealth}%` : "-" }),
            /* @__PURE__ */ jsx("td", { className: "pr-4", children: /* @__PURE__ */ jsxs("div", { className: "flex justify-end gap-1", children: [
              t.status === "Assigned" ? /* @__PURE__ */ jsxs(Btn, { className: "py-1 px-2 text-xs", onClick: () => setTakeoverHospitalTablet(t), children: [
                /* @__PURE__ */ jsx(ClipboardCheck, { size: 13 }),
                " Takeover"
              ] }) : /* @__PURE__ */ jsxs(Btn, { className: "py-1 px-2 text-xs", onClick: () => setHandoverTablet(t), children: [
                /* @__PURE__ */ jsx(Send, { size: 13 }),
                " Handover"
              ] }),
              /* @__PURE__ */ jsx("button", { className: "rounded p-1.5 text-slate-400 hover:bg-slate-200 hover:text-slate-700", onClick: () => {
                setEditing(t);
                setShowForm(true);
              }, children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
              /* @__PURE__ */ jsx("button", { title: "Delete device", className: "rounded p-1.5 text-slate-400 hover:bg-red-100 hover:text-red-600", onClick: () => setConfirmDeleteId(t.id), children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
            ] }) })
          ] }),
          isOpen && /* @__PURE__ */ jsx("tr", { className: "bg-slate-50/70 border-t border-slate-100", children: /* @__PURE__ */ jsxs("td", { colSpan: 9, className: "px-4 py-3", children: [
            t.customFields && Object.keys(t.customFields).length > 0 && /* @__PURE__ */ jsxs("div", { className: "mb-3", children: [
              /* @__PURE__ */ jsx("p", { className: "mb-1.5 text-xs font-semibold uppercase text-slate-400", children: "Additional imported fields" }),
              /* @__PURE__ */ jsx("div", { className: "flex flex-wrap gap-1.5", children: Object.entries(t.customFields).map(([k, v]) => /* @__PURE__ */ jsxs(Badge, { tone: "slate", children: [
                k,
                ": ",
                v
              ] }, k)) })
            ] }),
            /* @__PURE__ */ jsx("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: "Custody history" }),
            !t.history || t.history.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-slate-400", children: "No custody events recorded yet." }) : /* @__PURE__ */ jsx("ol", { className: "flex flex-col gap-2.5", children: t.history.map((h, i) => {
              const meta = custodyMeta(h.type);
              const Icon = meta.icon;
              return /* @__PURE__ */ jsxs("li", { className: "flex gap-2.5 text-xs", children: [
                /* @__PURE__ */ jsx("span", { className: "mt-0.5 shrink-0", style: { color: meta.color }, children: /* @__PURE__ */ jsx(Icon, { size: 14 }) }),
                /* @__PURE__ */ jsxs("div", { children: [
                  /* @__PURE__ */ jsxs("p", { children: [
                    /* @__PURE__ */ jsx("span", { className: "font-medium text-slate-700", children: meta.label }),
                    " ",
                    /* @__PURE__ */ jsxs("span", { className: "text-slate-400", children: [
                      "- ",
                      fmtDate(h.date)
                    ] })
                  ] }),
                  /* @__PURE__ */ jsx("p", { className: "text-slate-500", children: h.detail })
                ] })
              ] }, i);
            }) })
          ] }) })
        ] }, t.id);
      }) })
    ] }) }) : filteredLog.length === 0 ? /* @__PURE__ */ jsx(EmptyState, { icon: History, title: "No handover / takeover events yet", subtitle: "Events appear here as devices are taken over, handed over, or returned." }) : /* @__PURE__ */ jsx("div", { className: "overflow-x-auto rounded-lg border border-slate-200 bg-white", children: /* @__PURE__ */ jsxs("table", { className: "w-full min-w-[720px] text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-slate-50 text-left text-xs uppercase text-slate-400", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "py-2.5 pl-4", children: "Date" }),
        /* @__PURE__ */ jsx("th", { children: "Device" }),
        /* @__PURE__ */ jsx("th", { children: "Event" }),
        /* @__PURE__ */ jsx("th", { className: "pr-4", children: "Details" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: filteredLog.map((e, i) => {
        const meta = custodyMeta(e.type);
        const Icon = meta.icon;
        return /* @__PURE__ */ jsxs("tr", { className: "border-t border-slate-100", children: [
          /* @__PURE__ */ jsx("td", { className: "py-2.5 pl-4 text-slate-500 whitespace-nowrap", children: fmtDate(e.date) }),
          /* @__PURE__ */ jsx("td", { className: "font-medium text-slate-700 whitespace-nowrap", children: e.tablet.serialNumber }),
          /* @__PURE__ */ jsx("td", { children: /* @__PURE__ */ jsxs("span", { className: "inline-flex items-center gap-1.5 text-xs font-medium whitespace-nowrap", style: { color: meta.color }, children: [
            /* @__PURE__ */ jsx(Icon, { size: 14 }),
            " ",
            meta.label
          ] }) }),
          /* @__PURE__ */ jsx("td", { className: "pr-4 text-slate-500", children: e.detail })
        ] }, i);
      }) })
    ] }) }),
    showForm && /* @__PURE__ */ jsx(TabletFormModal, { tablet: editing, hospitals: data.hospitals, staff: data.staff, onSave: save, onClose: () => {
      setShowForm(false);
      setEditing(null);
    } }),
    takeoverSourceTablet && /* @__PURE__ */ jsx(
      TakeoverSourceModal,
      {
        hospitals: data.hospitals,
        sources: data.sources,
        existingTablets: data.tablets,
        onSave: takeoverFromSource,
        onClose: () => setTakeoverSourceTablet(null)
      }
    ),
    handoverTablet && /* @__PURE__ */ jsx(HandoverModal, { tablet: handoverTablet, hospitals: data.hospitals, staff: data.staff, onSave: handover, onClose: () => setHandoverTablet(null) }),
    takeoverHospitalTablet && /* @__PURE__ */ jsx(TakeoverHospitalModal, { tablet: takeoverHospitalTablet, onSave: takeoverFromHospital, onClose: () => setTakeoverHospitalTablet(null) }),
    showLetterModal && /* @__PURE__ */ jsx(
      HandoverLetterModal,
      {
        tablets: data.tablets,
        hospitals: data.hospitals,
        staff: data.staff,
        handoverOfficers: data.handoverOfficers,
        initialDeviceIds: Array.from(selectedForHandover),
        onConfirm: (letter) => {
          handoverViaLetter(letter);
          setSelectedForHandover(/* @__PURE__ */ new Set());
        },
        onClose: () => setShowLetterModal(false)
      }
    ),
    showReplacementModal && /* @__PURE__ */ jsx(
      ReplacementLetterModal,
      {
        tablets: data.tablets,
        hospitals: data.hospitals,
        staff: data.staff,
        handoverOfficers: data.handoverOfficers,
        onConfirm: replaceViaLetter,
        onClose: () => setShowReplacementModal(false)
      }
    ),
    confirmDeleteId && /* @__PURE__ */ jsx(
      ConfirmDialog,
      {
        message: "Delete this device record? Its custody history will also be removed. This cannot be undone.",
        onCancel: () => setConfirmDeleteId(null),
        onConfirm: () => {
          remove(confirmDeleteId);
          setConfirmDeleteId(null);
        }
      }
    )
  ] });
}
function TabletFormModal({ tablet, hospitals, staff, onSave, onClose }) {
  const [form, setForm] = useState(tablet || {
    id: uid("tab"),
    serialNumber: "",
    model: "",
    currentHospitalId: hospitals[0]?.id || "",
    assignedStaffId: null,
    nadraInstalled: false,
    batteryHealth: 100,
    status: "In Stock",
    remarks: "",
    history: []
  });
  const eligibleStaff = staff.filter((s) => s.hospitalId === form.currentHospitalId);
  const isAssigned = form.status === "Assigned";
  return /* @__PURE__ */ jsxs(Modal, { title: tablet ? "Edit device" : "Add device manually", onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Serial number", required: true, className: "col-span-2", children: /* @__PURE__ */ jsx(TextInput, { value: form.serialNumber, onChange: (e) => setForm((f) => ({ ...f, serialNumber: e.target.value })) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Model", children: /* @__PURE__ */ jsx(TextInput, { value: form.model, onChange: (e) => setForm((f) => ({ ...f, model: e.target.value })), placeholder: "e.g. Samsung Galaxy Tab A9" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Status", children: /* @__PURE__ */ jsx(Select, { value: form.status, onChange: (e) => setForm((f) => ({ ...f, status: e.target.value, assignedStaffId: e.target.value === "Assigned" ? f.assignedStaffId : null })), children: TABLET_STATUSES.map((s) => /* @__PURE__ */ jsx("option", { value: s, children: s }, s)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Current location", children: /* @__PURE__ */ jsx(Select, { value: form.currentHospitalId, onChange: (e) => setForm((f) => ({ ...f, currentHospitalId: e.target.value, assignedStaffId: null })), children: hospitals.map((h) => /* @__PURE__ */ jsx("option", { value: h.id, children: h.name }, h.id)) }) }),
      isAssigned && /* @__PURE__ */ jsx(Field, { label: "Assigned staff", required: true, className: "col-span-2", children: eligibleStaff.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-amber-600", children: "No staff registered at this location. Add staff first, or change the location above." }) : /* @__PURE__ */ jsxs(Select, { value: form.assignedStaffId || "", onChange: (e) => setForm((f) => ({ ...f, assignedStaffId: e.target.value })), children: [
        /* @__PURE__ */ jsx("option", { value: "", disabled: true, children: "Select staff member" }),
        eligibleStaff.map((s) => /* @__PURE__ */ jsxs("option", { value: s.id, children: [
          s.name,
          " - ",
          s.designation
        ] }, s.id))
      ] }) }),
      /* @__PURE__ */ jsx(Field, { label: "Battery health (%)", children: /* @__PURE__ */ jsx(TextInput, { type: "number", min: 0, max: 100, value: form.batteryHealth, onChange: (e) => setForm((f) => ({ ...f, batteryHealth: Number(e.target.value) })) }) }),
      /* @__PURE__ */ jsxs("label", { className: "col-span-2 flex items-center gap-2 text-sm text-slate-700", children: [
        /* @__PURE__ */ jsx("input", { type: "checkbox", checked: form.nadraInstalled, onChange: (e) => setForm((f) => ({ ...f, nadraInstalled: e.target.checked })) }),
        "NADRA app installed"
      ] }),
      /* @__PURE__ */ jsx(Field, { label: "Remarks", className: "col-span-2", children: /* @__PURE__ */ jsx(TextArea, { rows: 2, value: form.remarks, onChange: (e) => setForm((f) => ({ ...f, remarks: e.target.value })) }) })
    ] }),
    /* @__PURE__ */ jsx("p", { className: "mt-3 text-xs text-slate-400", children: `This form edits the device's current details directly without logging a custody event. For a proper chain-of-custody record, use "Handover" / "Takeover" instead.` }),
    /* @__PURE__ */ jsxs("div", { className: "mt-3 flex justify-end gap-2", children: [
      /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
      /* @__PURE__ */ jsx(Btn, { variant: "primary", disabled: !form.serialNumber || isAssigned && !form.assignedStaffId, onClick: () => onSave(form), children: "Save device" })
    ] })
  ] });
}
var TAKEOVER_CHECKLIST = ["Physical condition OK", "Screen intact", "Charger included", "Device powers on", "Functional test passed"];
function TakeoverSourceModal({ hospitals, sources, existingTablets, onSave, onClose }) {
  const [serialNumber, setSerialNumber] = useState("");
  const [model, setModel] = useState("");
  const [source, setSource] = useState(sources[0] || "");
  const [customSource, setCustomSource] = useState("");
  const [receivedBy, setReceivedBy] = useState("");
  const [warehouseId, setWarehouseId] = useState(hospitals[0]?.id || "");
  const [date, setDate] = useState(todayISO());
  const [checklist, setChecklist] = useState(TAKEOVER_CHECKLIST.map(() => true));
  const [overallCondition, setOverallCondition] = useState("Good");
  const [batteryHealth, setBatteryHealth] = useState(100);
  const [remarks, setRemarks] = useState("");
  const matchedTablet = existingTablets.find((t) => t.serialNumber.trim().toLowerCase() === serialNumber.trim().toLowerCase());
  const toggleCheck = (i) => setChecklist((prev) => prev.map((v, idx) => idx === i ? !v : v));
  const finalSource = source === "__other__" ? customSource : source;
  return /* @__PURE__ */ jsxs(Modal, { title: "Takeover device from source", onClose, width: "max-w-xl", children: [
    /* @__PURE__ */ jsx("p", { className: "mb-3 text-xs text-slate-500", children: "Record a device received from an office (e.g. DHO office), inspect it, and bring it into warehouse stock." }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsxs(Field, { label: "Device serial number", required: true, className: "col-span-2", children: [
        /* @__PURE__ */ jsx(TextInput, { value: serialNumber, onChange: (e) => setSerialNumber(e.target.value), placeholder: "Scan or type serial number" }),
        matchedTablet && /* @__PURE__ */ jsxs("p", { className: "mt-1 text-xs text-blue-600", children: [
          'Matches existing device "',
          matchedTablet.model || matchedTablet.serialNumber,
          '" - this will update it instead of creating a duplicate.'
        ] })
      ] }),
      !matchedTablet && /* @__PURE__ */ jsx(Field, { label: "Model", className: "col-span-2", children: /* @__PURE__ */ jsx(TextInput, { value: model, onChange: (e) => setModel(e.target.value), placeholder: "e.g. Samsung Galaxy Tab A9" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Received from (source)", required: true, children: /* @__PURE__ */ jsxs(Select, { value: source, onChange: (e) => setSource(e.target.value), children: [
        sources.map((s) => /* @__PURE__ */ jsx("option", { value: s, children: s }, s)),
        /* @__PURE__ */ jsx("option", { value: "__other__", children: "Other..." })
      ] }) }),
      source === "__other__" && /* @__PURE__ */ jsx(Field, { label: "Specify source", children: /* @__PURE__ */ jsx(TextInput, { value: customSource, onChange: (e) => setCustomSource(e.target.value), placeholder: "e.g. DHO Office G9" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Received by", children: /* @__PURE__ */ jsx(TextInput, { value: receivedBy, onChange: (e) => setReceivedBy(e.target.value), placeholder: "Name of person receiving" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Date", children: /* @__PURE__ */ jsx(TextInput, { type: "date", value: date, onChange: (e) => setDate(e.target.value) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Bring into warehouse", required: true, className: "col-span-2", children: /* @__PURE__ */ jsx(Select, { value: warehouseId, onChange: (e) => setWarehouseId(e.target.value), children: hospitals.map((h) => /* @__PURE__ */ jsx("option", { value: h.id, children: h.name }, h.id)) }) }),
      /* @__PURE__ */ jsxs("div", { className: "col-span-2 rounded-md border border-slate-200 p-3", children: [
        /* @__PURE__ */ jsx("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: "Inspection checklist" }),
        /* @__PURE__ */ jsx("div", { className: "grid grid-cols-2 gap-1.5", children: TAKEOVER_CHECKLIST.map((label, i) => /* @__PURE__ */ jsxs("label", { className: "flex items-center gap-1.5 text-sm text-slate-700", children: [
          /* @__PURE__ */ jsx("input", { type: "checkbox", checked: checklist[i], onChange: () => toggleCheck(i) }),
          " ",
          label
        ] }, label)) })
      ] }),
      /* @__PURE__ */ jsx(Field, { label: "Overall condition", required: true, children: /* @__PURE__ */ jsxs(Select, { value: overallCondition, onChange: (e) => setOverallCondition(e.target.value), children: [
        /* @__PURE__ */ jsx("option", { children: "Good" }),
        /* @__PURE__ */ jsx("option", { children: "Minor issues" }),
        /* @__PURE__ */ jsx("option", { children: "Damaged" })
      ] }) }),
      /* @__PURE__ */ jsx(Field, { label: "Battery health (%)", children: /* @__PURE__ */ jsx(TextInput, { type: "number", min: 0, max: 100, value: batteryHealth, onChange: (e) => setBatteryHealth(Number(e.target.value)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Remarks", className: "col-span-2", children: /* @__PURE__ */ jsx(TextArea, { rows: 2, value: remarks, onChange: (e) => setRemarks(e.target.value), placeholder: "Any notes about the handover from source" }) })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mt-5 flex justify-end gap-2", children: [
      /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
      /* @__PURE__ */ jsx(
        Btn,
        {
          variant: "primary",
          disabled: !serialNumber.trim() || !finalSource || !warehouseId,
          onClick: () => onSave({
            tabletId: matchedTablet?.id,
            serialNumber: serialNumber.trim(),
            model,
            source: finalSource,
            receivedBy,
            warehouseId,
            date,
            checklist,
            checklistLabels: TAKEOVER_CHECKLIST,
            overallCondition,
            batteryHealth,
            remarks
          }),
          children: "Confirm takeover"
        }
      )
    ] })
  ] });
}
function HandoverModal({ tablet, hospitals, staff, onSave, onClose }) {
  const [hospitalId, setHospitalId] = useState(tablet.currentHospitalId || hospitals[0]?.id || "");
  const eligibleStaff = staff.filter((s) => s.hospitalId === hospitalId);
  const [staffId, setStaffId] = useState(eligibleStaff[0]?.id || "");
  const [handedOverBy, setHandedOverBy] = useState("");
  const [date, setDate] = useState(todayISO());
  const [condition, setCondition] = useState("Good");
  const [accessories, setAccessories] = useState(["Charger"]);
  const [remarks, setRemarks] = useState("");
  const toggle = (a) => setAccessories((prev) => prev.includes(a) ? prev.filter((x) => x !== a) : [...prev, a]);
  useEffect(() => {
    setStaffId(eligibleStaff[0]?.id || "");
  }, [hospitalId]);
  return /* @__PURE__ */ jsxs(Modal, { title: `Handover device - ${tablet.serialNumber}`, onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "flex flex-col gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Hospital", required: true, children: /* @__PURE__ */ jsx(Select, { value: hospitalId, onChange: (e) => setHospitalId(e.target.value), children: hospitals.map((h) => /* @__PURE__ */ jsx("option", { value: h.id, children: h.name }, h.id)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Staff member", required: true, children: eligibleStaff.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-amber-600", children: "No staff registered at this hospital yet. Add staff first." }) : /* @__PURE__ */ jsx(Select, { value: staffId, onChange: (e) => setStaffId(e.target.value), children: eligibleStaff.map((s) => /* @__PURE__ */ jsxs("option", { value: s.id, children: [
        s.name,
        " - ",
        s.designation
      ] }, s.id)) }) }),
      /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
        /* @__PURE__ */ jsx(Field, { label: "Handed over by", children: /* @__PURE__ */ jsx(TextInput, { value: handedOverBy, onChange: (e) => setHandedOverBy(e.target.value), placeholder: "Warehouse keeper name" }) }),
        /* @__PURE__ */ jsx(Field, { label: "Date", children: /* @__PURE__ */ jsx(TextInput, { type: "date", value: date, onChange: (e) => setDate(e.target.value) }) })
      ] }),
      /* @__PURE__ */ jsx(Field, { label: "Condition at handover", children: /* @__PURE__ */ jsxs(Select, { value: condition, onChange: (e) => setCondition(e.target.value), children: [
        /* @__PURE__ */ jsx("option", { children: "Good" }),
        /* @__PURE__ */ jsx("option", { children: "Minor wear" }),
        /* @__PURE__ */ jsx("option", { children: "Damaged" })
      ] }) }),
      /* @__PURE__ */ jsx(Field, { label: "Accessories given", children: /* @__PURE__ */ jsx("div", { className: "flex flex-wrap gap-4 text-sm text-slate-700", children: ["Charger", "Cable", "Case", "SIM card"].map((a) => /* @__PURE__ */ jsxs("label", { className: "flex items-center gap-1.5", children: [
        /* @__PURE__ */ jsx("input", { type: "checkbox", checked: accessories.includes(a), onChange: () => toggle(a) }),
        " ",
        a
      ] }, a)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Remarks", children: /* @__PURE__ */ jsx(TextArea, { rows: 2, value: remarks, onChange: (e) => setRemarks(e.target.value) }) })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mt-5 flex justify-end gap-2", children: [
      /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
      /* @__PURE__ */ jsx(Btn, { variant: "primary", disabled: !hospitalId || !staffId, onClick: () => onSave(tablet, { hospitalId, staffId, handedOverBy, date, condition, accessories, remarks }), children: "Confirm handover" })
    ] })
  ] });
}
function TakeoverHospitalModal({ tablet, onSave, onClose }) {
  const [receivedBy, setReceivedBy] = useState("");
  const [date, setDate] = useState(todayISO());
  const [condition, setCondition] = useState("Good");
  const [accessories, setAccessories] = useState([]);
  const [inspectionNotes, setInspectionNotes] = useState("");
  const [remarks, setRemarks] = useState("");
  const toggle = (a) => setAccessories((prev) => prev.includes(a) ? prev.filter((x) => x !== a) : [...prev, a]);
  return /* @__PURE__ */ jsxs(Modal, { title: `Takeover device from hospital - ${tablet.serialNumber}`, onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "flex flex-col gap-3", children: [
      /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
        /* @__PURE__ */ jsx(Field, { label: "Received by", children: /* @__PURE__ */ jsx(TextInput, { value: receivedBy, onChange: (e) => setReceivedBy(e.target.value), placeholder: "Warehouse keeper name" }) }),
        /* @__PURE__ */ jsx(Field, { label: "Date", children: /* @__PURE__ */ jsx(TextInput, { type: "date", value: date, onChange: (e) => setDate(e.target.value) }) })
      ] }),
      /* @__PURE__ */ jsx(Field, { label: "Condition on return", required: true, children: /* @__PURE__ */ jsxs(Select, { value: condition, onChange: (e) => setCondition(e.target.value), children: [
        /* @__PURE__ */ jsx("option", { children: "Good" }),
        /* @__PURE__ */ jsx("option", { children: "Minor wear" }),
        /* @__PURE__ */ jsx("option", { children: "Damaged" })
      ] }) }),
      /* @__PURE__ */ jsx(Field, { label: "Accessories returned", children: /* @__PURE__ */ jsx("div", { className: "flex flex-wrap gap-4 text-sm text-slate-700", children: ["Charger", "Cable", "Case", "SIM card"].map((a) => /* @__PURE__ */ jsxs("label", { className: "flex items-center gap-1.5", children: [
        /* @__PURE__ */ jsx("input", { type: "checkbox", checked: accessories.includes(a), onChange: () => toggle(a) }),
        " ",
        a
      ] }, a)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Inspection notes", children: /* @__PURE__ */ jsx(TextArea, { rows: 2, value: inspectionNotes, onChange: (e) => setInspectionNotes(e.target.value), placeholder: "Physical condition, screen, functionality..." }) }),
      /* @__PURE__ */ jsx(Field, { label: "Remarks", children: /* @__PURE__ */ jsx(TextArea, { rows: 2, value: remarks, onChange: (e) => setRemarks(e.target.value) }) }),
      /* @__PURE__ */ jsx("p", { className: "text-xs text-slate-400", children: "Photos of returned condition can be attached once file upload is enabled for this deployment." })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mt-5 flex justify-end gap-2", children: [
      /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
      /* @__PURE__ */ jsx(Btn, { variant: "primary", onClick: () => onSave(tablet, { receivedBy, date, condition, accessories, inspectionNotes, remarks }), children: "Confirm takeover" })
    ] })
  ] });
}
var NUM_WORDS = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen", "Twenty"];
var numWord = (n) => NUM_WORDS[n] || String(n);
var pad2 = (n) => String(n).padStart(2, "0");
var ordinal = (n) => {
  const s = ["th", "st", "nd", "rd"], v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
};
var formatLetterDate = (iso) => {
  if (!iso) return "-";
  const d = new Date(iso);
  return `${ordinal(d.getDate())} ${d.toLocaleString(void 0, { month: "long" })}, ${d.getFullYear()}`;
};
function buildLetterHTML({ refNo, letterDate, hospitalName, hospitalLocation, handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, focalPerson, rows, tablets }) {
  const subjectText = `Handing-Over of ${numWord(rows.length)} (${pad2(rows.length)}) Android Tablet Device${rows.length > 1 ? "s" : ""} for Data Entry in Electronic Medical Record (EMR) System under &quot;Prime Minister Programme for the Elimination of Hepatitis C Infection&quot;.`;
  const tableRows = rows.map((r, i) => {
    const t = tablets.find((x) => x.id === r.tabletId);
    return `<tr>
      <td style="border:1px solid #1e293b;padding:6px;text-align:center;">${i + 1}.</td>
      <td style="border:1px solid #1e293b;padding:6px;">IVAS Tablet</td>
      <td style="border:1px solid #1e293b;padding:6px;">${t?.serialNumber || ""}</td>
      <td style="border:1px solid #1e293b;padding:6px;">${r.counter || ""}</td>
      <td style="border:1px solid #1e293b;padding:6px;text-align:center;">${fmtDate(r.date)}</td>
      <td style="border:1px solid #1e293b;padding:6px;">${r.status || ""}</td>
    </tr>`;
  }).join("");
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<title>Handover Letter ${refNo || ""}</title>
<style>
  @page { size: A4; margin: 20mm; }
  body { font-family: Georgia, 'Times New Roman', serif; font-size: 13px; line-height: 1.6; color: #0f172a; max-width: 800px; margin: 0 auto; padding: 24px; }
  table { border-collapse: collapse; width: 100%; font-size: 12px; margin-top: 12px; }
  .center { text-align: center; }
  .bold { font-weight: bold; }
  .underline { text-decoration: underline; }
  .row { display: flex; justify-content: space-between; margin-top: 16px; }
  .sig-grid { display: flex; justify-content: space-between; gap: 24px; margin-top: 48px; }
  .sig-line { margin-top: 44px; }
  @media print { body { padding: 0; } }
</style>
</head>
<body>
  <div class="center">
    <p class="bold">GOVERNMENT OF PAKISTAN</p>
    <p class="bold">Ministry of National Health Services, Regulations &amp; Coordination,</p>
    <p>3rd Floor Kohsar Block, New Pak Secretariat, Islamabad</p>
    <p style="letter-spacing:3px;">**********</p>
  </div>
  <div class="row">
    <p>F. No: ${refNo || "_______________"}</p>
    <p>Islamabad, the ${formatLetterDate(letterDate)}</p>
  </div>
  <p><span class="bold underline">Subject:</span> <span class="bold underline">${subjectText}</span></p>
  <p>I, <b>${handingOverName}</b>${handingOverDesignation ? `, Designation <b>${handingOverDesignation}</b>` : ""} (Hep-C Elimination Program) is hereby handing over ${numWord(rows.length)} (${pad2(rows.length)}) Android Tablet Device${rows.length > 1 ? "s" : ""} to <b>${takenOverName}</b> with following details: -</p>
  <table>
    <thead><tr style="background:#f8fafc;">
      <th style="border:1px solid #1e293b;padding:6px;">S.#</th>
      <th style="border:1px solid #1e293b;padding:6px;">Item Description</th>
      <th style="border:1px solid #1e293b;padding:6px;">Serial No./IMEI No.</th>
      <th style="border:1px solid #1e293b;padding:6px;">Counter</th>
      <th style="border:1px solid #1e293b;padding:6px;">Handing over Date</th>
      <th style="border:1px solid #1e293b;padding:6px;">Status</th>
    </tr></thead>
    <tbody>${tableRows}</tbody>
  </table>
  <p style="margin-top:12px;">The tablet device shall be utilized exclusively for data entry of citizens/individuals related to Hepatitis C screening, testing, and treatment${focalPerson ? ` under the supervision of focal person <b>${focalPerson}</b>` : ""} from ${hospitalName}.</p>
  <div class="sig-grid">
    <div>
      <p class="bold">HANDED OVER BY</p>
      <div class="sig-line"></div>
      <p>Name: ${handingOverName}</p>
      <p>Designation: ${handingOverDesignation || "-"}</p>
      <p style="margin-top:16px;">On behalf of.</p>
      <p class="bold">PM's Hepatitis C Elimination Program</p>
      <p class="bold">Ministry of NHSR&amp;C</p>
    </div>
    <div>
      <p class="bold">TAKEN OVER BY</p>
      <div class="sig-line"></div>
      <p>Name: ${takenOverName}</p>
      <p>Designation: ${takenOverDesignation || "-"}</p>
      <p>Contact No.: ${takenOverContact || "-"}</p>
      <p style="margin-top:16px;">On behalf of.</p>
      <p class="bold">${hospitalName}</p>
      ${hospitalLocation ? `<p>${hospitalLocation}</p>` : ""}
    </div>
  </div>
</body>
</html>`;
}
function downloadFile(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2e3);
}
function downloadLetterHTML(html, filename) {
  downloadFile(html, filename, "text/html");
}
function downloadLetterWord(html, filename) {
  const wordHtml = html.replace("<html>", '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">').replace("<head>", '<head>\n<meta name="ProgId" content="Word.Document">\n<meta name="Generator" content="Microsoft Word">');
  downloadFile(wordHtml, filename, "application/msword");
}
function HandoverLetterModal({ tablets, hospitals, staff, handoverOfficers = [], initialDeviceIds = [], onConfirm, onClose }) {
  const [step, setStep] = useState("form");
  const [refNo, setRefNo] = useState("");
  const [letterDate, setLetterDate] = useState(todayISO());
  const [hospitalId, setHospitalId] = useState(hospitals[0]?.id || "");
  const [handingOverOfficerId, setHandingOverOfficerId] = useState("");
  const [manualOfficerName, setManualOfficerName] = useState("");
  const [manualOfficerDesignation, setManualOfficerDesignation] = useState("");
  const [takenOverStaffId, setTakenOverStaffId] = useState("");
  const [manualName, setManualName] = useState("");
  const [manualDesignation, setManualDesignation] = useState("");
  const [manualContact, setManualContact] = useState("");
  const [focalPerson, setFocalPerson] = useState("");
  const [rows, setRows] = useState(
    () => tablets.filter((t) => initialDeviceIds.includes(t.id) && t.status === "In Stock").map((t) => ({ tabletId: t.id, counter: "", date: todayISO(), status: "Working" }))
  );
  const [deviceSearch, setDeviceSearch] = useState("");
  const availableTablets = tablets.filter((t) => t.status === "In Stock");
  const filteredAvailableTablets = availableTablets.filter(
    (t) => !deviceSearch || [t.serialNumber, t.model].join(" ").toLowerCase().includes(deviceSearch.toLowerCase())
  );
  const eligibleStaff = staff.filter((s) => s.hospitalId === hospitalId);
  const hospital = hospitals.find((h) => h.id === hospitalId);
  const isOfficerManual = handingOverOfficerId === "__manual__";
  const selectedOfficer = handoverOfficers.find((o) => o.id === handingOverOfficerId);
  const handingOverName = isOfficerManual ? manualOfficerName : selectedOfficer?.name || "";
  const handingOverDesignation = isOfficerManual ? manualOfficerDesignation : selectedOfficer?.designation || "";
  const isManual = takenOverStaffId === "__manual__";
  const selectedStaff = eligibleStaff.find((s) => s.id === takenOverStaffId);
  const takenOverName = isManual ? manualName : selectedStaff?.name || "";
  const takenOverDesignation = isManual ? manualDesignation : selectedStaff?.designation || "";
  const takenOverContact = isManual ? manualContact : "";
  const toggleDevice = (t) => setRows(
    (prev) => prev.some((r) => r.tabletId === t.id) ? prev.filter((r) => r.tabletId !== t.id) : [...prev, { tabletId: t.id, counter: "", date: letterDate, status: "Working" }]
  );
  const updateRow = (id, patch) => setRows((prev) => prev.map((r) => r.tabletId === id ? { ...r, ...patch } : r));
  const selectAllFiltered = () => setRows((prev) => {
    const existingIds = new Set(prev.map((r) => r.tabletId));
    const toAdd = filteredAvailableTablets.filter((t) => !existingIds.has(t.id)).map((t) => ({ tabletId: t.id, counter: "", date: letterDate, status: "Working" }));
    return [...prev, ...toAdd];
  });
  const clearAllFiltered = () => setRows((prev) => {
    const filteredIds = new Set(filteredAvailableTablets.map((t) => t.id));
    return prev.filter((r) => !filteredIds.has(r.tabletId));
  });
  const canPreview = hospitalId && handingOverName.trim() && takenOverName.trim() && rows.length > 0;
  const subjectText = `Handing-Over of ${numWord(rows.length)} (${pad2(rows.length)}) Android Tablet Device${rows.length > 1 ? "s" : ""} for Data Entry in Electronic Medical Record (EMR) System under "Prime Minister Programme for the Elimination of Hepatitis C Infection".`;
  function handleConfirm() {
    onConfirm({
      refNo,
      letterDate,
      hospitalId,
      handingOverName,
      handingOverDesignation,
      takenOverStaffId: isManual ? null : takenOverStaffId,
      takenOverName,
      takenOverDesignation,
      takenOverContact,
      focalPerson,
      rows
    });
    onClose();
  }
  return /* @__PURE__ */ jsxs(Modal, { title: step === "form" ? "Generate handover letter" : "Handover letter preview", onClose, width: "max-w-3xl", children: [
    /* @__PURE__ */ jsx("style", { children: `
        @page { size: A4; margin: 15mm; }
        @media print {
          body * { visibility: hidden; }
          .print-area, .print-area * { visibility: visible; }
          .print-area { position: absolute; top: 0; left: 0; width: 100%; padding: 0; margin: 0; border: none !important; box-shadow: none !important; }
          .no-print { display: none !important; }
        }
      ` }),
    step === "form" && /* @__PURE__ */ jsxs("div", { className: "flex flex-col gap-4", children: [
      /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
        /* @__PURE__ */ jsx(Field, { label: "Reference / F. No", children: /* @__PURE__ */ jsx(TextInput, { value: refNo, onChange: (e) => setRefNo(e.target.value), placeholder: "1-2025/26/Hep-C/Proc/" }) }),
        /* @__PURE__ */ jsx(Field, { label: "Letter date", children: /* @__PURE__ */ jsx(TextInput, { type: "date", value: letterDate, onChange: (e) => setLetterDate(e.target.value) }) }),
        /* @__PURE__ */ jsx(Field, { label: "Hospital", required: true, className: "col-span-2", children: /* @__PURE__ */ jsx(Select, { value: hospitalId, onChange: (e) => {
          setHospitalId(e.target.value);
          setTakenOverStaffId("");
        }, children: hospitals.map((h) => /* @__PURE__ */ jsx("option", { value: h.id, children: h.name }, h.id)) }) })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "rounded-md border border-slate-200 p-3", children: [
        /* @__PURE__ */ jsx("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: "Handed over by" }),
        /* @__PURE__ */ jsx(Field, { label: "Officer", required: true, children: /* @__PURE__ */ jsxs(Select, { value: handingOverOfficerId, onChange: (e) => setHandingOverOfficerId(e.target.value), children: [
          /* @__PURE__ */ jsx("option", { value: "", children: "Select handing-over officer..." }),
          handoverOfficers.map((o) => /* @__PURE__ */ jsxs("option", { value: o.id, children: [
            o.name,
            " - ",
            o.designation
          ] }, o.id)),
          /* @__PURE__ */ jsx("option", { value: "__manual__", children: "Enter manually..." })
        ] }) }),
        isOfficerManual && /* @__PURE__ */ jsxs("div", { className: "mt-3 grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
          /* @__PURE__ */ jsx(Field, { label: "Name", required: true, children: /* @__PURE__ */ jsx(TextInput, { value: manualOfficerName, onChange: (e) => setManualOfficerName(e.target.value), placeholder: "e.g. Wajahat Ahmed" }) }),
          /* @__PURE__ */ jsx(Field, { label: "Designation", children: /* @__PURE__ */ jsx(TextInput, { value: manualOfficerDesignation, onChange: (e) => setManualOfficerDesignation(e.target.value), placeholder: "e.g. IT / Data Officer" }) })
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "rounded-md border border-slate-200 p-3", children: [
        /* @__PURE__ */ jsx("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: "Taken over by" }),
        /* @__PURE__ */ jsx(Field, { label: "Person", required: true, children: /* @__PURE__ */ jsxs(Select, { value: takenOverStaffId, onChange: (e) => setTakenOverStaffId(e.target.value), children: [
          /* @__PURE__ */ jsx("option", { value: "", children: "Select hospital staff..." }),
          eligibleStaff.map((s) => /* @__PURE__ */ jsxs("option", { value: s.id, children: [
            s.name,
            " - ",
            s.designation
          ] }, s.id)),
          /* @__PURE__ */ jsx("option", { value: "__manual__", children: "Enter manually..." })
        ] }) }),
        isManual && /* @__PURE__ */ jsxs("div", { className: "mt-3 grid grid-cols-2 gap-3", children: [
          /* @__PURE__ */ jsx(Field, { label: "Name", required: true, children: /* @__PURE__ */ jsx(TextInput, { value: manualName, onChange: (e) => setManualName(e.target.value) }) }),
          /* @__PURE__ */ jsx(Field, { label: "Designation", children: /* @__PURE__ */ jsx(TextInput, { value: manualDesignation, onChange: (e) => setManualDesignation(e.target.value) }) }),
          /* @__PURE__ */ jsx(Field, { label: "Contact No.", className: "col-span-2", children: /* @__PURE__ */ jsx(TextInput, { value: manualContact, onChange: (e) => setManualContact(e.target.value) }) })
        ] })
      ] }),
      /* @__PURE__ */ jsx(Field, { label: "Focal person supervising usage (optional)", children: /* @__PURE__ */ jsx(TextInput, { value: focalPerson, onChange: (e) => setFocalPerson(e.target.value), placeholder: "e.g. Mr. Ihtisham Afzal Khan" }) }),
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("div", { className: "mb-2 flex flex-wrap items-center justify-between gap-2", children: [
          /* @__PURE__ */ jsxs("p", { className: "text-xs font-semibold uppercase text-slate-400", children: [
            "Devices (",
            rows.length,
            ' selected) - only devices currently "In Stock" are shown'
          ] }),
          availableTablets.length > 0 && /* @__PURE__ */ jsxs("div", { className: "flex gap-2", children: [
            /* @__PURE__ */ jsxs("button", { type: "button", onClick: selectAllFiltered, className: "text-xs font-medium text-slate-600 hover:text-slate-900 hover:underline", children: [
              "Select all",
              deviceSearch ? " filtered" : ""
            ] }),
            /* @__PURE__ */ jsx("span", { className: "text-xs text-slate-300", children: "|" }),
            /* @__PURE__ */ jsxs("button", { type: "button", onClick: clearAllFiltered, className: "text-xs font-medium text-slate-600 hover:text-slate-900 hover:underline", children: [
              "Clear",
              deviceSearch ? " filtered" : ""
            ] })
          ] })
        ] }),
        availableTablets.length === 0 ? /* @__PURE__ */ jsx("p", { className: "rounded-md bg-amber-50 px-3 py-2 text-xs text-amber-700", children: 'No devices with status "In Stock". Use "Takeover device" first to bring devices into a warehouse.' }) : /* @__PURE__ */ jsxs(Fragment2, { children: [
          /* @__PURE__ */ jsxs("div", { className: "relative mb-2", children: [
            /* @__PURE__ */ jsx(Search, { size: 14, className: "absolute left-2.5 top-2.5 text-slate-400" }),
            /* @__PURE__ */ jsx(TextInput, { placeholder: "Search by serial number...", value: deviceSearch, onChange: (e) => setDeviceSearch(e.target.value), className: "pl-8" })
          ] }),
          filteredAvailableTablets.length === 0 ? /* @__PURE__ */ jsxs("p", { className: "rounded-md bg-slate-50 px-3 py-2 text-xs text-slate-400", children: [
            'No devices match "',
            deviceSearch,
            '".'
          ] }) : /* @__PURE__ */ jsx("div", { className: "flex flex-col gap-2 max-h-72 overflow-y-auto pr-1", children: filteredAvailableTablets.map((t) => {
            const row = rows.find((r) => r.tabletId === t.id);
            const checked = !!row;
            return /* @__PURE__ */ jsxs("div", { className: cx("rounded-md border p-2.5", checked ? "border-slate-900 bg-slate-50" : "border-slate-200"), children: [
              /* @__PURE__ */ jsxs("label", { className: "flex items-center gap-2 text-sm", children: [
                /* @__PURE__ */ jsx("input", { type: "checkbox", checked, onChange: () => toggleDevice(t) }),
                /* @__PURE__ */ jsx("span", { className: "font-medium text-slate-800", children: t.serialNumber }),
                /* @__PURE__ */ jsx("span", { className: "text-xs text-slate-400", children: "IVAS Tablet" })
              ] }),
              checked && /* @__PURE__ */ jsxs("div", { className: "mt-2 grid grid-cols-3 gap-2 pl-6", children: [
                /* @__PURE__ */ jsx(TextInput, { placeholder: "Counter (e.g. OPD Counter - Ground Floor)", value: row.counter, onChange: (e) => updateRow(t.id, { counter: e.target.value }), className: "col-span-3 sm:col-span-1" }),
                /* @__PURE__ */ jsx(TextInput, { type: "date", value: row.date, onChange: (e) => updateRow(t.id, { date: e.target.value }) }),
                /* @__PURE__ */ jsx(TextInput, { placeholder: "Status (e.g. Working)", value: row.status, onChange: (e) => updateRow(t.id, { status: e.target.value }) })
              ] })
            ] }, t.id);
          }) })
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex justify-end gap-2", children: [
        /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
        /* @__PURE__ */ jsxs(Btn, { variant: "primary", disabled: !canPreview, onClick: () => setStep("preview"), children: [
          /* @__PURE__ */ jsx(FileText, { size: 15 }),
          " Preview letter"
        ] })
      ] })
    ] }),
    step === "preview" && /* @__PURE__ */ jsxs("div", { children: [
      /* @__PURE__ */ jsxs("div", { className: "print-area rounded-md border border-slate-300 bg-white p-6 text-[13px] leading-relaxed text-slate-900", style: { fontFamily: "Georgia, 'Times New Roman', serif" }, children: [
        /* @__PURE__ */ jsxs("div", { className: "text-center", children: [
          /* @__PURE__ */ jsx("p", { className: "font-bold", children: "GOVERNMENT OF PAKISTAN" }),
          /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "Ministry of National Health Services, Regulations & Coordination," }),
          /* @__PURE__ */ jsx("p", { children: "3rd Floor Kohsar Block, New Pak Secretariat, Islamabad" }),
          /* @__PURE__ */ jsx("p", { className: "tracking-widest", children: "**********" })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "mt-4 flex justify-between text-sm", children: [
          /* @__PURE__ */ jsxs("p", { children: [
            "F. No: ",
            refNo || "_______________"
          ] }),
          /* @__PURE__ */ jsxs("p", { children: [
            "Islamabad, the ",
            formatLetterDate(letterDate)
          ] })
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "mt-4", children: [
          /* @__PURE__ */ jsx("span", { className: "font-semibold underline", children: "Subject:" }),
          " ",
          /* @__PURE__ */ jsx("span", { className: "font-semibold underline", children: subjectText })
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "mt-3", children: [
          "I, ",
          /* @__PURE__ */ jsx("b", { children: handingOverName }),
          handingOverDesignation && /* @__PURE__ */ jsxs(Fragment2, { children: [
            ", Designation ",
            /* @__PURE__ */ jsx("b", { children: handingOverDesignation })
          ] }),
          " (Hep-C Elimination Program) is hereby handing over ",
          numWord(rows.length),
          " (",
          pad2(rows.length),
          ") Android Tablet Device",
          rows.length > 1 ? "s" : "",
          " to ",
          /* @__PURE__ */ jsx("b", { children: takenOverName }),
          " with following details: -"
        ] }),
        /* @__PURE__ */ jsxs("table", { className: "mt-3 w-full border-collapse text-xs", children: [
          /* @__PURE__ */ jsx("thead", { children: /* @__PURE__ */ jsxs("tr", { className: "bg-slate-50", children: [
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "S.#" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Item Description" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Serial No./IMEI No." }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Counter" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Handing over Date" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Status" })
          ] }) }),
          /* @__PURE__ */ jsx("tbody", { children: rows.map((r, i) => {
            const t = tablets.find((x) => x.id === r.tabletId);
            return /* @__PURE__ */ jsxs("tr", { children: [
              /* @__PURE__ */ jsxs("td", { className: "border border-slate-800 p-1.5 text-center", children: [
                i + 1,
                "."
              ] }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5", children: "IVAS Tablet" }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5", children: t?.serialNumber }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5", children: r.counter }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5 text-center", children: fmtDate(r.date) }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5", children: r.status })
            ] }, r.tabletId);
          }) })
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "mt-3", children: [
          "The tablet device shall be utilized exclusively for data entry of citizens/individuals related to Hepatitis C screening, testing, and treatment",
          focalPerson ? /* @__PURE__ */ jsxs(Fragment2, { children: [
            " under the supervision of focal person ",
            /* @__PURE__ */ jsx("b", { children: focalPerson })
          ] }) : "",
          " from ",
          hospital?.name,
          "."
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "mt-10 grid grid-cols-2 gap-6 text-sm", children: [
          /* @__PURE__ */ jsxs("div", { children: [
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "HANDED OVER BY" }),
            /* @__PURE__ */ jsx("div", { className: "mt-11" }),
            /* @__PURE__ */ jsxs("p", { className: "mt-1", children: [
              "Name: ",
              handingOverName
            ] }),
            /* @__PURE__ */ jsxs("p", { children: [
              "Designation: ",
              handingOverDesignation || "-"
            ] }),
            /* @__PURE__ */ jsx("p", { className: "mt-4", children: "On behalf of." }),
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "PM's Hepatitis C Elimination Program" }),
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "Ministry of NHSR&C" })
          ] }),
          /* @__PURE__ */ jsxs("div", { children: [
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "TAKEN OVER BY" }),
            /* @__PURE__ */ jsx("div", { className: "mt-11" }),
            /* @__PURE__ */ jsxs("p", { className: "mt-1", children: [
              "Name: ",
              takenOverName
            ] }),
            /* @__PURE__ */ jsxs("p", { children: [
              "Designation: ",
              takenOverDesignation || "-"
            ] }),
            /* @__PURE__ */ jsxs("p", { children: [
              "Contact No.: ",
              takenOverContact || "-"
            ] }),
            /* @__PURE__ */ jsx("p", { className: "mt-4", children: "On behalf of." }),
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: hospital?.name }),
            hospital?.location && /* @__PURE__ */ jsx("p", { children: hospital.location })
          ] })
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "no-print mt-4 flex justify-between gap-2", children: [
        /* @__PURE__ */ jsxs(Btn, { onClick: () => setStep("form"), children: [
          /* @__PURE__ */ jsx(ArrowLeft, { size: 15 }),
          " Back to form"
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "flex flex-wrap justify-end gap-2", children: [
          /* @__PURE__ */ jsxs(Btn, { onClick: () => window.print(), children: [
            /* @__PURE__ */ jsx(Printer, { size: 15 }),
            " Print"
          ] }),
          /* @__PURE__ */ jsxs(
            Btn,
            {
              onClick: () => downloadLetterWord(
                buildLetterHTML({ refNo, letterDate, hospitalName: hospital?.name || "", hospitalLocation: hospital?.location || "", handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, focalPerson, rows, tablets }),
                `Handover-Letter-${(refNo || todayISO()).replace(/[^\w-]+/g, "_")}.doc`
              ),
              children: [
                /* @__PURE__ */ jsx(FileText, { size: 15 }),
                " Export Word"
              ]
            }
          ),
          /* @__PURE__ */ jsxs(
            Btn,
            {
              onClick: () => downloadLetterHTML(
                buildLetterHTML({ refNo, letterDate, hospitalName: hospital?.name || "", hospitalLocation: hospital?.location || "", handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, focalPerson, rows, tablets }),
                `Handover-Letter-${(refNo || todayISO()).replace(/[^\w-]+/g, "_")}.html`
              ),
              children: [
                /* @__PURE__ */ jsx(Download, { size: 15 }),
                " Export PDF"
              ]
            }
          ),
          /* @__PURE__ */ jsxs(Btn, { variant: "primary", onClick: handleConfirm, children: [
            /* @__PURE__ */ jsx(Send, { size: 15 }),
            " Confirm handover"
          ] })
        ] })
      ] }),
      /* @__PURE__ */ jsx("p", { className: "no-print mt-2 text-xs text-slate-400", children: `"Print" opens your browser's print dialog directly. If it doesn't respond (some browsers block printing from embedded previews), use "Export PDF" - it downloads a print-ready file; open it and choose "Save as PDF" from its print dialog. "Export Word" downloads an editable .doc file that opens directly in Microsoft Word.` })
    ] })
  ] });
}
function buildReplacementLetterHTML({ refNo, letterDate, hospitalName, hospitalLocation, handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, rows, tablets }) {
  const subjectText = `Replacement of ${numWord(rows.length)} (${pad2(rows.length)}) Android Tablet Device${rows.length > 1 ? "s" : ""} under &quot;Prime Minister Programme for the Elimination of Hepatitis C Infection&quot;.`;
  const tableRows = rows.map((r, i) => {
    const oldT = tablets.find((x) => x.id === r.oldTabletId);
    const newT = tablets.find((x) => x.id === r.newTabletId);
    return `<tr>
      <td style="border:1px solid #1e293b;padding:6px;text-align:center;">${i + 1}.</td>
      <td style="border:1px solid #1e293b;padding:6px;">${oldT?.serialNumber || ""}</td>
      <td style="border:1px solid #1e293b;padding:6px;">${newT?.serialNumber || ""}</td>
      <td style="border:1px solid #1e293b;padding:6px;">${r.reason || ""}</td>
      <td style="border:1px solid #1e293b;padding:6px;">${r.counter || ""}</td>
      <td style="border:1px solid #1e293b;padding:6px;text-align:center;">${fmtDate(r.date)}</td>
      <td style="border:1px solid #1e293b;padding:6px;">${r.status || ""}</td>
    </tr>`;
  }).join("");
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<title>Replacement Letter ${refNo || ""}</title>
<style>
  @page { size: A4; margin: 20mm; }
  body { font-family: Georgia, 'Times New Roman', serif; font-size: 13px; line-height: 1.6; color: #0f172a; max-width: 800px; margin: 0 auto; padding: 24px; }
  table { border-collapse: collapse; width: 100%; font-size: 11px; margin-top: 12px; }
  .center { text-align: center; }
  .bold { font-weight: bold; }
  .underline { text-decoration: underline; }
  .row { display: flex; justify-content: space-between; margin-top: 16px; }
  .sig-grid { display: flex; justify-content: space-between; gap: 24px; margin-top: 48px; }
  .sig-line { margin-top: 44px; }
  @media print { body { padding: 0; } }
</style>
</head>
<body>
  <div class="center">
    <p class="bold">GOVERNMENT OF PAKISTAN</p>
    <p class="bold">Ministry of National Health Services, Regulations &amp; Coordination,</p>
    <p>3rd Floor Kohsar Block, New Pak Secretariat, Islamabad</p>
    <p style="letter-spacing:3px;">**********</p>
  </div>
  <div class="row">
    <p>F. No: ${refNo || "_______________"}</p>
    <p>Islamabad, the ${formatLetterDate(letterDate)}</p>
  </div>
  <p><span class="bold underline">Subject:</span> <span class="bold underline">${subjectText}</span></p>
  <p>I, <b>${handingOverName}</b>${handingOverDesignation ? `, Designation <b>${handingOverDesignation}</b>` : ""} (Hep-C Elimination Program) am hereby replacing the following faulty/damaged device(s) at <b>${hospitalName}</b> and handing over the replacement device(s) to <b>${takenOverName}</b> with following details: -</p>
  <table>
    <thead><tr style="background:#f8fafc;">
      <th style="border:1px solid #1e293b;padding:6px;">S.#</th>
      <th style="border:1px solid #1e293b;padding:6px;">Old Device<br/>IMEI (Returned)</th>
      <th style="border:1px solid #1e293b;padding:6px;">New Device<br/>IMEI (Issued)</th>
      <th style="border:1px solid #1e293b;padding:6px;">Reason for Replacement</th>
      <th style="border:1px solid #1e293b;padding:6px;">Counter</th>
      <th style="border:1px solid #1e293b;padding:6px;">Replacement Date</th>
      <th style="border:1px solid #1e293b;padding:6px;">Status</th>
    </tr></thead>
    <tbody>${tableRows}</tbody>
  </table>
  <p style="margin-top:12px;">The replacement device(s) shall be utilized exclusively for data entry of citizens/individuals related to Hepatitis C screening, testing, and treatment at ${hospitalName}. The returned/faulty device(s) shall be taken back into program custody for inspection and repair.</p>
  <div class="sig-grid">
    <div>
      <p class="bold">HANDED OVER BY</p>
      <div class="sig-line"></div>
      <p>Name: ${handingOverName}</p>
      <p>Designation: ${handingOverDesignation || "-"}</p>
      <p style="margin-top:16px;">On behalf of.</p>
      <p class="bold">PM's Hepatitis C Elimination Program</p>
      <p class="bold">Ministry of NHSR&amp;C</p>
    </div>
    <div>
      <p class="bold">TAKEN OVER BY</p>
      <div class="sig-line"></div>
      <p>Name: ${takenOverName}</p>
      <p>Designation: ${takenOverDesignation || "-"}</p>
      <p>Contact No.: ${takenOverContact || "-"}</p>
      <p style="margin-top:16px;">On behalf of.</p>
      <p class="bold">${hospitalName}</p>
      ${hospitalLocation ? `<p>${hospitalLocation}</p>` : ""}
    </div>
  </div>
</body>
</html>`;
}
function ReplacementLetterModal({ tablets, hospitals, staff, handoverOfficers = [], onConfirm, onClose }) {
  const [step, setStep] = useState("form");
  const [refNo, setRefNo] = useState("");
  const [letterDate, setLetterDate] = useState(todayISO());
  const [hospitalId, setHospitalId] = useState(hospitals[0]?.id || "");
  const [handingOverOfficerId, setHandingOverOfficerId] = useState("");
  const [manualOfficerName, setManualOfficerName] = useState("");
  const [manualOfficerDesignation, setManualOfficerDesignation] = useState("");
  const [takenOverStaffId, setTakenOverStaffId] = useState("");
  const [manualName, setManualName] = useState("");
  const [manualDesignation, setManualDesignation] = useState("");
  const [manualContact, setManualContact] = useState("");
  const [rows, setRows] = useState([]);
  const [deviceSearch, setDeviceSearch] = useState("");
  const assignedAtHospital = tablets.filter((t) => t.status === "Assigned" && t.currentHospitalId === hospitalId);
  const filteredAssignedAtHospital = assignedAtHospital.filter(
    (t) => !deviceSearch || [t.serialNumber, t.model].join(" ").toLowerCase().includes(deviceSearch.toLowerCase())
  );
  const availableTablets = tablets.filter((t) => t.status === "In Stock");
  const eligibleStaff = staff.filter((s) => s.hospitalId === hospitalId);
  const hospital = hospitals.find((h) => h.id === hospitalId);
  const isOfficerManual = handingOverOfficerId === "__manual__";
  const selectedOfficer = handoverOfficers.find((o) => o.id === handingOverOfficerId);
  const handingOverName = isOfficerManual ? manualOfficerName : selectedOfficer?.name || "";
  const handingOverDesignation = isOfficerManual ? manualOfficerDesignation : selectedOfficer?.designation || "";
  const isManual = takenOverStaffId === "__manual__";
  const selectedStaff = eligibleStaff.find((s) => s.id === takenOverStaffId);
  const takenOverName = isManual ? manualName : selectedStaff?.name || "";
  const takenOverDesignation = isManual ? manualDesignation : selectedStaff?.designation || "";
  const takenOverContact = isManual ? manualContact : "";
  const usedNewIds = rows.map((r) => r.newTabletId).filter(Boolean);
  const toggleOld = (t) => setRows(
    (prev) => prev.some((r) => r.oldTabletId === t.id) ? prev.filter((r) => r.oldTabletId !== t.id) : [...prev, { oldTabletId: t.id, newTabletId: "", reason: "", counter: "", date: letterDate, status: "Working", oldStatus: "Damaged" }]
  );
  const updateRow = (oldId, patch) => setRows((prev) => prev.map((r) => r.oldTabletId === oldId ? { ...r, ...patch } : r));
  const canPreview = hospitalId && handingOverName.trim() && takenOverName.trim() && rows.length > 0 && rows.every((r) => r.newTabletId);
  const subjectText = `Replacement of ${numWord(rows.length)} (${pad2(rows.length)}) Android Tablet Device${rows.length > 1 ? "s" : ""} under "Prime Minister Programme for the Elimination of Hepatitis C Infection".`;
  function handleConfirm() {
    onConfirm({
      refNo,
      letterDate,
      hospitalId,
      handingOverName,
      handingOverDesignation,
      takenOverStaffId: isManual ? null : takenOverStaffId,
      takenOverName,
      takenOverDesignation,
      takenOverContact,
      rows
    });
    onClose();
  }
  return /* @__PURE__ */ jsxs(Modal, { title: step === "form" ? "Generate replacement letter" : "Replacement letter preview", onClose, width: "max-w-3xl", children: [
    /* @__PURE__ */ jsx("style", { children: `
        @page { size: A4; margin: 15mm; }
        @media print {
          body * { visibility: hidden; }
          .print-area, .print-area * { visibility: visible; }
          .print-area { position: absolute; top: 0; left: 0; width: 100%; padding: 0; margin: 0; border: none !important; box-shadow: none !important; }
          .no-print { display: none !important; }
        }
      ` }),
    step === "form" && /* @__PURE__ */ jsxs("div", { className: "flex flex-col gap-4", children: [
      /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
        /* @__PURE__ */ jsx(Field, { label: "Reference / F. No", children: /* @__PURE__ */ jsx(TextInput, { value: refNo, onChange: (e) => setRefNo(e.target.value) }) }),
        /* @__PURE__ */ jsx(Field, { label: "Letter date", children: /* @__PURE__ */ jsx(TextInput, { type: "date", value: letterDate, onChange: (e) => setLetterDate(e.target.value) }) }),
        /* @__PURE__ */ jsx(Field, { label: "Hospital", required: true, className: "col-span-2", children: /* @__PURE__ */ jsx(Select, { value: hospitalId, onChange: (e) => {
          setHospitalId(e.target.value);
          setTakenOverStaffId("");
          setRows([]);
          setDeviceSearch("");
        }, children: hospitals.map((h) => /* @__PURE__ */ jsx("option", { value: h.id, children: h.name }, h.id)) }) })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "rounded-md border border-slate-200 p-3", children: [
        /* @__PURE__ */ jsx("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: "Handed over by" }),
        /* @__PURE__ */ jsx(Field, { label: "Officer", required: true, children: /* @__PURE__ */ jsxs(Select, { value: handingOverOfficerId, onChange: (e) => setHandingOverOfficerId(e.target.value), children: [
          /* @__PURE__ */ jsx("option", { value: "", children: "Select handing-over officer..." }),
          handoverOfficers.map((o) => /* @__PURE__ */ jsxs("option", { value: o.id, children: [
            o.name,
            " - ",
            o.designation
          ] }, o.id)),
          /* @__PURE__ */ jsx("option", { value: "__manual__", children: "Enter manually..." })
        ] }) }),
        isOfficerManual && /* @__PURE__ */ jsxs("div", { className: "mt-3 grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
          /* @__PURE__ */ jsx(Field, { label: "Name", required: true, children: /* @__PURE__ */ jsx(TextInput, { value: manualOfficerName, onChange: (e) => setManualOfficerName(e.target.value), placeholder: "e.g. Wajahat Ahmed" }) }),
          /* @__PURE__ */ jsx(Field, { label: "Designation", children: /* @__PURE__ */ jsx(TextInput, { value: manualOfficerDesignation, onChange: (e) => setManualOfficerDesignation(e.target.value), placeholder: "e.g. IT / Data Officer" }) })
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "rounded-md border border-slate-200 p-3", children: [
        /* @__PURE__ */ jsx("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: "Taken over by (receiving replacement)" }),
        /* @__PURE__ */ jsx(Field, { label: "Person", required: true, children: /* @__PURE__ */ jsxs(Select, { value: takenOverStaffId, onChange: (e) => setTakenOverStaffId(e.target.value), children: [
          /* @__PURE__ */ jsx("option", { value: "", children: "Select hospital staff..." }),
          eligibleStaff.map((s) => /* @__PURE__ */ jsxs("option", { value: s.id, children: [
            s.name,
            " - ",
            s.designation
          ] }, s.id)),
          /* @__PURE__ */ jsx("option", { value: "__manual__", children: "Enter manually..." })
        ] }) }),
        isManual && /* @__PURE__ */ jsxs("div", { className: "mt-3 grid grid-cols-2 gap-3", children: [
          /* @__PURE__ */ jsx(Field, { label: "Name", required: true, children: /* @__PURE__ */ jsx(TextInput, { value: manualName, onChange: (e) => setManualName(e.target.value) }) }),
          /* @__PURE__ */ jsx(Field, { label: "Designation", children: /* @__PURE__ */ jsx(TextInput, { value: manualDesignation, onChange: (e) => setManualDesignation(e.target.value) }) }),
          /* @__PURE__ */ jsx(Field, { label: "Contact No.", className: "col-span-2", children: /* @__PURE__ */ jsx(TextInput, { value: manualContact, onChange: (e) => setManualContact(e.target.value) }) })
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { children: [
        /* @__PURE__ */ jsxs("p", { className: "mb-2 text-xs font-semibold uppercase text-slate-400", children: [
          "Devices to replace (",
          rows.length,
          " selected) - devices currently Assigned at this hospital"
        ] }),
        !hospitalId ? /* @__PURE__ */ jsx("p", { className: "text-xs text-slate-400", children: "Select a hospital first." }) : assignedAtHospital.length === 0 ? /* @__PURE__ */ jsx("p", { className: "rounded-md bg-amber-50 px-3 py-2 text-xs text-amber-700", children: "No devices are currently Assigned at this hospital." }) : /* @__PURE__ */ jsxs(Fragment2, { children: [
          /* @__PURE__ */ jsxs("div", { className: "relative mb-2", children: [
            /* @__PURE__ */ jsx(Search, { size: 14, className: "absolute left-2.5 top-2.5 text-slate-400" }),
            /* @__PURE__ */ jsx(TextInput, { placeholder: "Search by serial number...", value: deviceSearch, onChange: (e) => setDeviceSearch(e.target.value), className: "pl-8" })
          ] }),
          filteredAssignedAtHospital.length === 0 ? /* @__PURE__ */ jsxs("p", { className: "rounded-md bg-slate-50 px-3 py-2 text-xs text-slate-400", children: [
            'No devices match "',
            deviceSearch,
            '".'
          ] }) : /* @__PURE__ */ jsx("div", { className: "flex flex-col gap-2 max-h-80 overflow-y-auto pr-1", children: filteredAssignedAtHospital.map((t) => {
            const row = rows.find((r) => r.oldTabletId === t.id);
            const checked = !!row;
            const newOptions = availableTablets.filter((nt) => nt.id === row?.newTabletId || !usedNewIds.includes(nt.id));
            return /* @__PURE__ */ jsxs("div", { className: cx("rounded-md border p-2.5", checked ? "border-slate-900 bg-slate-50" : "border-slate-200"), children: [
              /* @__PURE__ */ jsxs("label", { className: "flex items-center gap-2 text-sm", children: [
                /* @__PURE__ */ jsx("input", { type: "checkbox", checked, onChange: () => toggleOld(t) }),
                /* @__PURE__ */ jsx("span", { className: "font-medium text-slate-800", children: t.serialNumber }),
                /* @__PURE__ */ jsx("span", { className: "text-xs text-slate-400", children: "IVAS Tablet (old device)" })
              ] }),
              checked && /* @__PURE__ */ jsxs("div", { className: "mt-2 flex flex-col gap-2 pl-6", children: [
                /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-2", children: [
                  /* @__PURE__ */ jsx(Field, { label: "Replacement device", required: true, children: newOptions.length === 0 ? /* @__PURE__ */ jsx("p", { className: "text-xs text-amber-600", children: 'No available "In Stock" devices left. Takeover a device first.' }) : /* @__PURE__ */ jsxs(Select, { value: row.newTabletId, onChange: (e) => updateRow(t.id, { newTabletId: e.target.value }), children: [
                    /* @__PURE__ */ jsx("option", { value: "", children: "Select new device..." }),
                    newOptions.map((nt) => /* @__PURE__ */ jsx("option", { value: nt.id, children: nt.serialNumber }, nt.id))
                  ] }) }),
                  /* @__PURE__ */ jsx(Field, { label: "Old device becomes", children: /* @__PURE__ */ jsx(Select, { value: row.oldStatus, onChange: (e) => updateRow(t.id, { oldStatus: e.target.value }), children: TABLET_STATUSES.filter((s) => s !== "Assigned").map((s) => /* @__PURE__ */ jsx("option", { value: s, children: s }, s)) }) })
                ] }),
                /* @__PURE__ */ jsx(TextInput, { placeholder: "Reason for replacement (e.g. screen damage, malfunction)", value: row.reason, onChange: (e) => updateRow(t.id, { reason: e.target.value }) }),
                /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-3 gap-2", children: [
                  /* @__PURE__ */ jsx(TextInput, { placeholder: "Counter", value: row.counter, onChange: (e) => updateRow(t.id, { counter: e.target.value }), className: "col-span-1" }),
                  /* @__PURE__ */ jsx(TextInput, { type: "date", value: row.date, onChange: (e) => updateRow(t.id, { date: e.target.value }) }),
                  /* @__PURE__ */ jsx(TextInput, { placeholder: "Status", value: row.status, onChange: (e) => updateRow(t.id, { status: e.target.value }) })
                ] })
              ] })
            ] }, t.id);
          }) })
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "flex justify-end gap-2", children: [
        /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
        /* @__PURE__ */ jsxs(Btn, { variant: "primary", disabled: !canPreview, onClick: () => setStep("preview"), children: [
          /* @__PURE__ */ jsx(FileText, { size: 15 }),
          " Preview letter"
        ] })
      ] })
    ] }),
    step === "preview" && /* @__PURE__ */ jsxs("div", { children: [
      /* @__PURE__ */ jsxs("div", { className: "print-area rounded-md border border-slate-300 bg-white p-6 text-[13px] leading-relaxed text-slate-900", style: { fontFamily: "Georgia, 'Times New Roman', serif" }, children: [
        /* @__PURE__ */ jsxs("div", { className: "text-center", children: [
          /* @__PURE__ */ jsx("p", { className: "font-bold", children: "GOVERNMENT OF PAKISTAN" }),
          /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "Ministry of National Health Services, Regulations & Coordination," }),
          /* @__PURE__ */ jsx("p", { children: "3rd Floor Kohsar Block, New Pak Secretariat, Islamabad" }),
          /* @__PURE__ */ jsx("p", { className: "tracking-widest", children: "**********" })
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "mt-4 flex justify-between text-sm", children: [
          /* @__PURE__ */ jsxs("p", { children: [
            "F. No: ",
            refNo || "_______________"
          ] }),
          /* @__PURE__ */ jsxs("p", { children: [
            "Islamabad, the ",
            formatLetterDate(letterDate)
          ] })
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "mt-4", children: [
          /* @__PURE__ */ jsx("span", { className: "font-semibold underline", children: "Subject:" }),
          " ",
          /* @__PURE__ */ jsx("span", { className: "font-semibold underline", children: subjectText })
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "mt-3", children: [
          "I, ",
          /* @__PURE__ */ jsx("b", { children: handingOverName }),
          handingOverDesignation && /* @__PURE__ */ jsxs(Fragment2, { children: [
            ", Designation ",
            /* @__PURE__ */ jsx("b", { children: handingOverDesignation })
          ] }),
          " (Hep-C Elimination Program) am hereby replacing the following faulty/damaged device(s) at ",
          /* @__PURE__ */ jsx("b", { children: hospital?.name }),
          " and handing over the replacement device(s) to ",
          /* @__PURE__ */ jsx("b", { children: takenOverName }),
          " with following details: -"
        ] }),
        /* @__PURE__ */ jsxs("table", { className: "mt-3 w-full border-collapse text-xs", children: [
          /* @__PURE__ */ jsx("thead", { children: /* @__PURE__ */ jsxs("tr", { className: "bg-slate-50", children: [
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "S.#" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Old Device IMEI (Returned)" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "New Device IMEI (Issued)" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Reason for Replacement" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Counter" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Replacement Date" }),
            /* @__PURE__ */ jsx("th", { className: "border border-slate-800 p-1.5", children: "Status" })
          ] }) }),
          /* @__PURE__ */ jsx("tbody", { children: rows.map((r, i) => {
            const oldT = tablets.find((x) => x.id === r.oldTabletId);
            const newT = tablets.find((x) => x.id === r.newTabletId);
            return /* @__PURE__ */ jsxs("tr", { children: [
              /* @__PURE__ */ jsxs("td", { className: "border border-slate-800 p-1.5 text-center", children: [
                i + 1,
                "."
              ] }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5", children: oldT?.serialNumber }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5", children: newT?.serialNumber }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5", children: r.reason }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5", children: r.counter }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5 text-center", children: fmtDate(r.date) }),
              /* @__PURE__ */ jsx("td", { className: "border border-slate-800 p-1.5", children: r.status })
            ] }, r.oldTabletId);
          }) })
        ] }),
        /* @__PURE__ */ jsxs("p", { className: "mt-3", children: [
          "The replacement device(s) shall be utilized exclusively for data entry of citizens/individuals related to Hepatitis C screening, testing, and treatment at ",
          hospital?.name,
          ". The returned/faulty device(s) shall be taken back into program custody for inspection and repair."
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "mt-10 grid grid-cols-2 gap-6 text-sm", children: [
          /* @__PURE__ */ jsxs("div", { children: [
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "HANDED OVER BY" }),
            /* @__PURE__ */ jsx("div", { className: "mt-11" }),
            /* @__PURE__ */ jsxs("p", { className: "mt-1", children: [
              "Name: ",
              handingOverName
            ] }),
            /* @__PURE__ */ jsxs("p", { children: [
              "Designation: ",
              handingOverDesignation || "-"
            ] }),
            /* @__PURE__ */ jsx("p", { className: "mt-4", children: "On behalf of." }),
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "PM's Hepatitis C Elimination Program" }),
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "Ministry of NHSR&C" })
          ] }),
          /* @__PURE__ */ jsxs("div", { children: [
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: "TAKEN OVER BY" }),
            /* @__PURE__ */ jsx("div", { className: "mt-11" }),
            /* @__PURE__ */ jsxs("p", { className: "mt-1", children: [
              "Name: ",
              takenOverName
            ] }),
            /* @__PURE__ */ jsxs("p", { children: [
              "Designation: ",
              takenOverDesignation || "-"
            ] }),
            /* @__PURE__ */ jsxs("p", { children: [
              "Contact No.: ",
              takenOverContact || "-"
            ] }),
            /* @__PURE__ */ jsx("p", { className: "mt-4", children: "On behalf of." }),
            /* @__PURE__ */ jsx("p", { className: "font-semibold", children: hospital?.name }),
            hospital?.location && /* @__PURE__ */ jsx("p", { children: hospital.location })
          ] })
        ] })
      ] }),
      /* @__PURE__ */ jsxs("div", { className: "no-print mt-4 flex justify-between gap-2", children: [
        /* @__PURE__ */ jsxs(Btn, { onClick: () => setStep("form"), children: [
          /* @__PURE__ */ jsx(ArrowLeft, { size: 15 }),
          " Back to form"
        ] }),
        /* @__PURE__ */ jsxs("div", { className: "flex flex-wrap justify-end gap-2", children: [
          /* @__PURE__ */ jsxs(Btn, { onClick: () => window.print(), children: [
            /* @__PURE__ */ jsx(Printer, { size: 15 }),
            " Print"
          ] }),
          /* @__PURE__ */ jsxs(
            Btn,
            {
              onClick: () => downloadLetterWord(
                buildReplacementLetterHTML({ refNo, letterDate, hospitalName: hospital?.name || "", hospitalLocation: hospital?.location || "", handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, rows, tablets }),
                `Replacement-Letter-${(refNo || todayISO()).replace(/[^\w-]+/g, "_")}.doc`
              ),
              children: [
                /* @__PURE__ */ jsx(FileText, { size: 15 }),
                " Export Word"
              ]
            }
          ),
          /* @__PURE__ */ jsxs(
            Btn,
            {
              onClick: () => downloadLetterHTML(
                buildReplacementLetterHTML({ refNo, letterDate, hospitalName: hospital?.name || "", hospitalLocation: hospital?.location || "", handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, rows, tablets }),
                `Replacement-Letter-${(refNo || todayISO()).replace(/[^\w-]+/g, "_")}.html`
              ),
              children: [
                /* @__PURE__ */ jsx(Download, { size: 15 }),
                " Export PDF"
              ]
            }
          ),
          /* @__PURE__ */ jsxs(Btn, { variant: "primary", onClick: handleConfirm, children: [
            /* @__PURE__ */ jsx(Send, { size: 15 }),
            " Confirm replacement"
          ] })
        ] })
      ] }),
      /* @__PURE__ */ jsx("p", { className: "no-print mt-2 text-xs text-slate-400", children: `"Print" opens your browser's print dialog directly. If it doesn't respond, use "Export PDF" - it downloads a print-ready file; open it and choose "Save as PDF". "Export Word" downloads an editable .doc file that opens directly in Microsoft Word.` })
    ] })
  ] });
}
function StaffTab({ data, update }) {
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const hospitalName = (id) => data.hospitals.find((h) => h.id === id)?.name || "-";
  const filtered = data.staff.filter((s) => !search || [s.name, s.employeeId, s.designation].join(" ").toLowerCase().includes(search.toLowerCase()));
  function save(s) {
    update((prev) => {
      const exists = prev.staff.some((x) => x.id === s.id);
      return { ...prev, staff: exists ? prev.staff.map((x) => x.id === s.id ? s : x) : [...prev.staff, s] };
    });
    setShowForm(false);
    setEditing(null);
  }
  function remove(id) {
    update((prev) => ({ ...prev, staff: prev.staff.filter((x) => x.id !== id) }));
  }
  return /* @__PURE__ */ jsxs("div", { children: [
    /* @__PURE__ */ jsx(Toolbar, { title: "Hospital staff", subtitle: "Directory of staff eligible for tablet and asset assignment.", search, setSearch, children: /* @__PURE__ */ jsxs(Btn, { variant: "primary", onClick: () => {
      setEditing(null);
      setShowForm(true);
    }, children: [
      /* @__PURE__ */ jsx(Plus, { size: 15 }),
      " Add staff"
    ] }) }),
    data.staff.length === 0 ? /* @__PURE__ */ jsx(EmptyState, { icon: Users, title: "No staff registered yet" }) : /* @__PURE__ */ jsx("div", { className: "overflow-x-auto rounded-lg border border-slate-200 bg-white", children: /* @__PURE__ */ jsxs("table", { className: "w-full min-w-[720px] text-sm", children: [
      /* @__PURE__ */ jsx("thead", { className: "bg-slate-50 text-left text-xs uppercase text-slate-400", children: /* @__PURE__ */ jsxs("tr", { children: [
        /* @__PURE__ */ jsx("th", { className: "py-2.5 pl-4", children: "Employee ID" }),
        /* @__PURE__ */ jsx("th", { children: "Name" }),
        /* @__PURE__ */ jsx("th", { children: "Designation" }),
        /* @__PURE__ */ jsx("th", { children: "Department" }),
        /* @__PURE__ */ jsx("th", { children: "Hospital" }),
        /* @__PURE__ */ jsx("th", { className: "text-right pr-4", children: "Actions" })
      ] }) }),
      /* @__PURE__ */ jsx("tbody", { children: filtered.map((s) => /* @__PURE__ */ jsxs("tr", { className: "border-t border-slate-100 hover:bg-slate-50", children: [
        /* @__PURE__ */ jsx("td", { className: "py-2.5 pl-4 text-slate-500", children: s.employeeId }),
        /* @__PURE__ */ jsxs("td", { className: "font-medium text-slate-700", children: [
          s.name,
          /* @__PURE__ */ jsx("p", { className: "text-xs font-normal text-slate-400", children: s.cnic })
        ] }),
        /* @__PURE__ */ jsx("td", { className: "text-slate-600", children: s.designation }),
        /* @__PURE__ */ jsx("td", { className: "text-slate-600", children: s.department }),
        /* @__PURE__ */ jsx("td", { className: "text-slate-600", children: hospitalName(s.hospitalId) }),
        /* @__PURE__ */ jsx("td", { className: "pr-4", children: /* @__PURE__ */ jsxs("div", { className: "flex justify-end gap-1", children: [
          /* @__PURE__ */ jsx("button", { title: "Edit", className: "rounded p-1.5 text-slate-400 hover:bg-slate-200 hover:text-slate-700", onClick: () => {
            setEditing(s);
            setShowForm(true);
          }, children: /* @__PURE__ */ jsx(Pencil, { size: 14 }) }),
          /* @__PURE__ */ jsx("button", { title: "Delete", className: "rounded p-1.5 text-slate-400 hover:bg-red-100 hover:text-red-600", onClick: () => setConfirmDeleteId(s.id), children: /* @__PURE__ */ jsx(Trash2, { size: 14 }) })
        ] }) })
      ] }, s.id)) })
    ] }) }),
    showForm && /* @__PURE__ */ jsx(StaffFormModal, { staff: editing, hospitals: data.hospitals, designations: data.designations, departments: data.departments, onSave: save, onClose: () => {
      setShowForm(false);
      setEditing(null);
    } }),
    confirmDeleteId && /* @__PURE__ */ jsx(
      ConfirmDialog,
      {
        message: "Delete this staff member? This cannot be undone.",
        onCancel: () => setConfirmDeleteId(null),
        onConfirm: () => {
          remove(confirmDeleteId);
          setConfirmDeleteId(null);
        }
      }
    )
  ] });
}
function StaffFormModal({ staff, hospitals, designations, departments, onSave, onClose }) {
  const [form, setForm] = useState(staff || {
    id: uid("staff"),
    employeeId: "",
    name: "",
    cnic: "",
    designation: designations[0] || "",
    department: departments[0] || "",
    hospitalId: hospitals[0]?.id || ""
  });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  return /* @__PURE__ */ jsxs(Modal, { title: staff ? "Edit staff member" : "Add staff member", onClose, children: [
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3", children: [
      /* @__PURE__ */ jsx(Field, { label: "Employee ID", required: true, children: /* @__PURE__ */ jsx(TextInput, { value: form.employeeId, onChange: set("employeeId") }) }),
      /* @__PURE__ */ jsx(Field, { label: "Full name", required: true, children: /* @__PURE__ */ jsx(TextInput, { value: form.name, onChange: set("name") }) }),
      /* @__PURE__ */ jsx(Field, { label: "CNIC", children: /* @__PURE__ */ jsx(TextInput, { value: form.cnic, onChange: set("cnic"), placeholder: "XXXXX-XXXXXXX-X" }) }),
      /* @__PURE__ */ jsx(Field, { label: "Hospital", required: true, children: /* @__PURE__ */ jsx(Select, { value: form.hospitalId, onChange: set("hospitalId"), children: hospitals.map((h) => /* @__PURE__ */ jsx("option", { value: h.id, children: h.name }, h.id)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Designation", required: true, children: /* @__PURE__ */ jsx(Select, { value: form.designation, onChange: set("designation"), children: designations.map((d) => /* @__PURE__ */ jsx("option", { value: d, children: d }, d)) }) }),
      /* @__PURE__ */ jsx(Field, { label: "Department", required: true, children: /* @__PURE__ */ jsx(Select, { value: form.department, onChange: set("department"), children: departments.map((d) => /* @__PURE__ */ jsx("option", { value: d, children: d }, d)) }) })
    ] }),
    /* @__PURE__ */ jsxs("div", { className: "mt-5 flex justify-end gap-2", children: [
      /* @__PURE__ */ jsx(Btn, { onClick: onClose, children: "Cancel" }),
      /* @__PURE__ */ jsx(Btn, { variant: "primary", disabled: !form.name || !form.employeeId, onClick: () => onSave(form), children: "Save staff member" })
    ] })
  ] });
}
function SettingsTab({ data, update }) {
  const setCategories = (categories) => update((prev) => ({ ...prev, categories }));
  const setUnits = (units) => update((prev) => ({ ...prev, units }));
  const setDesignations = (designations) => update((prev) => ({ ...prev, designations }));
  const setDepartments = (departments) => update((prev) => ({ ...prev, departments }));
  const setSources = (sources) => update((prev) => ({ ...prev, sources }));
  const setHandoverOfficers = (handoverOfficers) => update((prev) => ({ ...prev, handoverOfficers }));
  return /* @__PURE__ */ jsxs("div", { children: [
    /* @__PURE__ */ jsx(Toolbar, { title: "Admin settings", subtitle: "Configure master data used across the system - no code changes required." }),
    /* @__PURE__ */ jsxs("div", { className: "grid grid-cols-1 gap-4 md:grid-cols-2", children: [
      /* @__PURE__ */ jsx(
        ConfigList,
        {
          label: "Inventory categories",
          items: data.categories,
          extraFields: {
            initial: { name: "", group: "" },
            valid: (d) => d.name.trim(),
            display: (c) => `${c.name}${c.group ? ` (${c.group})` : ""}`,
            render: (d, setD) => /* @__PURE__ */ jsxs("div", { className: "flex flex-1 gap-2", children: [
              /* @__PURE__ */ jsx(TextInput, { placeholder: "Category name", value: d.name, onChange: (e) => setD({ ...d, name: e.target.value }) }),
              /* @__PURE__ */ jsx(TextInput, { placeholder: "Group (optional)", value: d.group, onChange: (e) => setD({ ...d, group: e.target.value }) })
            ] })
          },
          onAdd: (d) => {
            const id = uid("cat");
            setCategories([{ id, ...d }, ...data.categories]);
            return id;
          },
          onEdit: (id, d) => setCategories(data.categories.map((c) => c.id === id ? { ...c, ...d } : c)),
          onDelete: (id) => setCategories(data.categories.filter((c) => c.id !== id))
        }
      ),
      /* @__PURE__ */ jsx(
        ConfigList,
        {
          label: "Units of measure",
          items: data.units,
          onAdd: (v) => setUnits([v, ...data.units]),
          onEdit: (id, v) => setUnits(data.units.map((u) => u === id ? v : u)),
          onDelete: (id) => setUnits(data.units.filter((u) => u !== id))
        }
      ),
      /* @__PURE__ */ jsx(
        ConfigList,
        {
          label: "Staff designations",
          items: data.designations,
          onAdd: (v) => setDesignations([v, ...data.designations]),
          onEdit: (id, v) => setDesignations(data.designations.map((d) => d === id ? v : d)),
          onDelete: (id) => setDesignations(data.designations.filter((d) => d !== id))
        }
      ),
      /* @__PURE__ */ jsx(
        ConfigList,
        {
          label: "Departments",
          items: data.departments,
          onAdd: (v) => setDepartments([v, ...data.departments]),
          onEdit: (id, v) => setDepartments(data.departments.map((d) => d === id ? v : d)),
          onDelete: (id) => setDepartments(data.departments.filter((d) => d !== id))
        }
      ),
      /* @__PURE__ */ jsx(
        ConfigList,
        {
          label: "Device handover / takeover sources",
          items: data.sources,
          onAdd: (v) => setSources([v, ...data.sources]),
          onEdit: (id, v) => setSources(data.sources.map((s) => s === id ? v : s)),
          onDelete: (id) => setSources(data.sources.filter((s) => s !== id))
        }
      ),
      /* @__PURE__ */ jsx(
        ConfigList,
        {
          label: "Handing-over officers",
          items: data.handoverOfficers,
          extraFields: {
            initial: { name: "", designation: "" },
            valid: (d) => d.name.trim(),
            display: (o) => `${o.name}${o.designation ? ` (${o.designation})` : ""}`,
            render: (d, setD) => /* @__PURE__ */ jsxs("div", { className: "flex flex-1 gap-2", children: [
              /* @__PURE__ */ jsx(TextInput, { placeholder: "Name", value: d.name, onChange: (e) => setD({ ...d, name: e.target.value }) }),
              /* @__PURE__ */ jsx(TextInput, { placeholder: "Designation", value: d.designation, onChange: (e) => setD({ ...d, designation: e.target.value }) })
            ] })
          },
          onAdd: (d) => {
            const id = uid("off");
            setHandoverOfficers([{ id, ...d }, ...data.handoverOfficers]);
            return id;
          },
          onEdit: (id, d) => setHandoverOfficers(data.handoverOfficers.map((o) => o.id === id ? { ...o, ...d } : o)),
          onDelete: (id) => setHandoverOfficers(data.handoverOfficers.filter((o) => o.id !== id))
        }
      )
    ] }),
    /* @__PURE__ */ jsx("p", { className: "mt-4 text-xs text-slate-400", children: 'Warehouses and hospitals are configured from the "Hospitals & Warehouses" tab since they carry extra fields (type, location).' })
  ] });
}

// src/entry.jsx
import { jsx as jsx2 } from "react/jsx-runtime";
createRoot(document.getElementById("root")).render(/* @__PURE__ */ jsx2(App, {}));
