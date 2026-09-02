import { useState, useEffect, useMemo, useCallback, useRef, Fragment } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Package, ArrowLeftRight, Building2, Tablet as TabletIcon,
  Users, Settings, Plus, Pencil, Trash2, X, Search, AlertTriangle,
  ChevronDown, ArrowUpCircle, ArrowDownCircle, RefreshCcw, CornerDownLeft,
  ShieldAlert, Ban, Loader2, Boxes, ClipboardCheck, Send, History, Inbox,
  FileText, Printer, ArrowLeft, Download, Upload, Menu
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend
} from "recharts";

/* ---------------------------------------------------------------------- */
/* Constants & helpers                                                     */
/* ---------------------------------------------------------------------- */

const SUPABASE_URL = "https://frkbeuzymwuxnwohvmix.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZya2JldXp5bXd1eG53b2h2bWl4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM1ODI0NzQsImV4cCI6MjA5OTE1ODQ3NH0.W6qIF_XRQtpLcPpN7ECR6N-EA9jU5BSjuT6dmQ2iGgY";
const APP_STATE_ROW_ID = "main";

async function supabaseRequest(path, options = {}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
      "Content-Type": "application/json",
      Prefer: "return=representation",
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Supabase ${res.status}: ${text || res.statusText}`);
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
    body: JSON.stringify({ data }),
  });
}

const TX_TYPES = [
  { id: "RECEIVE", label: "Receive Stock", dir: 1, icon: ArrowDownCircle, color: "#0F6E56" },
  { id: "ISSUE", label: "Issue Stock", dir: -1, icon: ArrowUpCircle, color: "#993C1D" },
  { id: "TRANSFER", label: "Transfer Stock", dir: 0, icon: ArrowLeftRight, color: "#185FA5" },
  { id: "ADJUST", label: "Adjust Stock", dir: null, icon: RefreshCcw, color: "#854F0B" },
  { id: "RETURN", label: "Return Stock", dir: 1, icon: CornerDownLeft, color: "#3B6D11" },
  { id: "DAMAGED", label: "Damaged Stock", dir: -1, icon: AlertTriangle, color: "#A32D2D" },
  { id: "LOST", label: "Lost Stock", dir: -1, icon: ShieldAlert, color: "#712B13" },
  { id: "DISPOSED", label: "Disposed Stock", dir: -1, icon: Ban, color: "#444441" },
];

const WAREHOUSE_TYPES = ["Central Warehouse", "Regional Warehouse", "Hospital Warehouse"];
const TABLET_STATUSES = ["In Stock", "Assigned", "Under Repair", "Damaged", "Lost", "Disposed"];

const DEFAULT_DATA = {
  categories: [
    { id: "cat-1", name: "RDT Kits", group: "Medical Supplies" },
    { id: "cat-2", name: "Tablets", group: "Equipment" },
    { id: "cat-3", name: "Blood Testing Devices", group: "Equipment" },
    { id: "cat-4", name: "Printers", group: "Equipment" },
    { id: "cat-5", name: "Barcode Scanners", group: "Equipment" },
  ],
  units: ["Piece", "Box", "Pack", "Kit", "Unit", "Carton"],
  hospitals: [
    { id: "wh-1", name: "Central Warehouse - Islamabad", type: "Central Warehouse", location: "Islamabad" },
    { id: "wh-2", name: "Regional Warehouse - Punjab", type: "Regional Warehouse", location: "Lahore" },
    { id: "wh-3", name: "DHQ Hospital Rawalpindi", type: "Hospital Warehouse", location: "Rawalpindi" },
  ],
  designations: ["Medical Officer", "Lab Technician", "Data Encoder", "Program Coordinator", "IT Support"],
  departments: ["Hepatitis Clinic", "Laboratory", "IT", "Administration", "Pharmacy"],
  sources: ["DHO Office G9", "Program HQ Islamabad", "PITB"],
  items: [],
  stockLevels: [],
  transactions: [],
  tablets: [],
  staff: [],
};

const uid = (p = "id") =>
  `${p}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

const todayISO = () => new Date().toISOString().slice(0, 10);

const fmtDate = (d) => {
  if (!d) return "-";
  try {
    return new Date(d).toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
  } catch {
    return d;
  }
};

const cx = (...a) => a.filter(Boolean).join(" ");

/* ---------------------------------------------------------------------- */
/* Generic UI primitives                                                   */
/* ---------------------------------------------------------------------- */

function Btn({ children, variant = "secondary", className = "", ...props }) {
  const base = "inline-flex items-center gap-1.5 rounded-md px-3 py-2 text-sm font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed";
  const styles = {
    primary: "bg-slate-900 text-white hover:bg-slate-800",
    secondary: "bg-white text-slate-700 border border-slate-300 hover:bg-slate-50",
    danger: "bg-red-600 text-white hover:bg-red-700",
    ghost: "text-slate-600 hover:bg-slate-100",
  };
  return (
    <button className={cx(base, styles[variant], className)} {...props}>
      {children}
    </button>
  );
}

function Field({ label, children, required, className = "" }) {
  return (
    <label className={cx("flex flex-col gap-1 text-sm", className)}>
      <span className="font-medium text-slate-700">
        {label} {required && <span className="text-red-500">*</span>}
      </span>
      {children}
    </label>
  );
}

const inputCls = "w-full rounded-md border border-slate-300 px-3 py-2 text-sm text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-900/20 focus:border-slate-400 bg-white";

function TextInput(props) {
  return <input {...props} className={cx(inputCls, props.className)} />;
}
function Select({ children, ...props }) {
  return (
    <select {...props} className={cx(inputCls, "cursor-pointer", props.className)}>
      {children}
    </select>
  );
}
function TextArea(props) {
  return <textarea {...props} className={cx(inputCls, "resize-none", props.className)} />;
}

function Badge({ children, tone = "slate" }) {
  const tones = {
    slate: "bg-slate-100 text-slate-700",
    green: "bg-emerald-100 text-emerald-800",
    red: "bg-red-100 text-red-800",
    amber: "bg-amber-100 text-amber-800",
    blue: "bg-blue-100 text-blue-800",
  };
  return <span className={cx("rounded-full px-2 py-0.5 text-xs font-medium whitespace-nowrap", tones[tone])}>{children}</span>;
}

function Modal({ title, onClose, children, width = "max-w-lg" }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className={cx("w-full rounded-lg bg-white shadow-xl max-h-[90vh] flex flex-col", width)}>
        <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5">
          <h3 className="text-base font-semibold text-slate-900">{title}</h3>
          <button onClick={onClose} className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600" aria-label="Close">
            <X size={18} />
          </button>
        </div>
        <div className="overflow-y-auto px-5 py-4">{children}</div>
      </div>
    </div>
  );
}

function ConfirmDialog({ title = "Confirm delete", message, confirmLabel = "Delete", onConfirm, onCancel }) {
  return (
    <Modal title={title} onClose={onCancel} width="max-w-sm">
      <p className="text-sm text-slate-600">{message}</p>
      <div className="mt-5 flex justify-end gap-2">
        <Btn onClick={onCancel}>Cancel</Btn>
        <Btn variant="danger" onClick={onConfirm}><Trash2 size={14} /> {confirmLabel}</Btn>
      </div>
    </Modal>
  );
}

function EmptyState({ icon: Icon, title, subtitle }) {
  return (
    <div className="flex flex-col items-center justify-center gap-2 py-14 text-center">
      <Icon size={32} className="text-slate-300" />
      <p className="text-sm font-medium text-slate-500">{title}</p>
      {subtitle && <p className="text-xs text-slate-400">{subtitle}</p>}
    </div>
  );
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

  return (
    <div className="rounded-lg border border-slate-200 bg-white p-4">
      <h4 className="mb-3 text-sm font-semibold text-slate-800">{label} <span className="font-normal text-slate-400">({items.length})</span></h4>
      <div className="mb-3 flex gap-2">
        {extraFields ? (
          extraFields.render(draft, setDraft)
        ) : (
          <TextInput placeholder={`Add new ${label.toLowerCase()}...`} value={draft} onChange={(e) => setDraft(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
        )}
        <Btn type="button" variant="primary" onClick={submit}><Plus size={15} /> Add</Btn>
      </div>
      <div className="flex flex-col gap-1.5 max-h-64 overflow-y-auto pr-1">
        {items.length === 0 && <p className="text-xs text-slate-400 py-2">No entries yet - add your first one above.</p>}
        {items.map((it) => {
          const id = it.id ?? it;
          const isEditing = editingId === id;
          const isNew = justAddedId === id;
          return (
            <div key={id} className={cx("flex items-center justify-between gap-2 rounded-md px-3 py-2 text-sm transition-colors", isNew ? "bg-emerald-50 ring-1 ring-emerald-300" : "bg-slate-50")}>
              {isEditing ? (
                <div className="flex flex-1 gap-2">
                  {extraFields ? extraFields.render(editDraft, setEditDraft) : <TextInput value={editDraft} onChange={(e) => setEditDraft(e.target.value)} className="py-1" />}
                </div>
              ) : (
                <span className="text-slate-700">{extraFields ? extraFields.display(it) : it}</span>
              )}
              <div className="flex shrink-0 gap-1">
                {isEditing ? (
                  <>
                    <Btn type="button" variant="ghost" className="px-2 py-1" onClick={() => { onEdit(id, editDraft); setEditingId(null); }}>Save</Btn>
                    <Btn type="button" variant="ghost" className="px-2 py-1" onClick={() => setEditingId(null)}>Cancel</Btn>
                  </>
                ) : (
                  <>
                    <button type="button" className="rounded p-1 text-slate-400 hover:bg-slate-200 hover:text-slate-700" onClick={() => { setEditingId(id); setEditDraft(extraFields ? it : it); }}><Pencil size={14} /></button>
                    <button type="button" className="rounded p-1 text-slate-400 hover:bg-red-100 hover:text-red-600" onClick={() => onDelete(id)}><Trash2 size={14} /></button>
                  </>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );

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

/* ---------------------------------------------------------------------- */
/* Main App                                                                */
/* ---------------------------------------------------------------------- */

const TABS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "inventory", label: "Inventory Items", icon: Package },
  { id: "transactions", label: "Stock Transactions", icon: ArrowLeftRight },
  { id: "warehouses", label: "Hospitals & Warehouses", icon: Building2 },
  { id: "tablets", label: "Tablet Assets", icon: TabletIcon },
  { id: "staff", label: "Staff", icon: Users },
  { id: "settings", label: "Admin Settings", icon: Settings },
];

export default function App() {
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
          // First run: nothing saved yet in Supabase - seed it with the defaults.
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

  const selectTab = (id) => { setActiveTab(id); setMobileMenuOpen(false); };

  if (loading) {
    return (
      <div className="flex h-[70vh] items-center justify-center gap-2 text-slate-400">
        <Loader2 className="animate-spin" size={20} /> Loading inventory system...
      </div>
    );
  }

  return (
    <div className="app-shell relative flex w-full min-w-0 flex-col overflow-hidden bg-slate-50 text-slate-900 md:flex-row">
      {/* Mobile top bar */}
      <div className="flex items-center justify-between border-b border-slate-200 bg-white px-4 py-3 md:hidden">
        <div className="flex items-center gap-2">
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-slate-900 text-white">
            <Boxes size={16} />
          </div>
          <p className="text-sm font-semibold text-slate-900">HCV Program</p>
        </div>
        <button onClick={() => setMobileMenuOpen(true)} className="rounded-md p-2 text-slate-600 hover:bg-slate-100" aria-label="Open menu">
          <Menu size={20} />
        </button>
      </div>

      {/* Mobile slide-in drawer */}
      {mobileMenuOpen && (
        <div className="absolute inset-0 z-50 flex md:hidden">
          <div className="absolute inset-0 bg-slate-900/40" onClick={() => setMobileMenuOpen(false)} />
          <div className="relative flex h-full w-72 max-w-[85vw] flex-col bg-white shadow-xl">
            <div className="flex items-center justify-between border-b border-slate-200 px-4 py-4">
              <p className="text-sm font-semibold text-slate-900">Menu</p>
              <button onClick={() => setMobileMenuOpen(false)} className="rounded p-1 text-slate-400 hover:bg-slate-100" aria-label="Close menu">
                <X size={18} />
              </button>
            </div>
            <SidebarNav activeTab={activeTab} setActiveTab={selectTab} saving={saving} connectionError={connectionError} />
          </div>
        </div>
      )}

      {/* Desktop sidebar */}
      <aside className="hidden w-64 shrink-0 flex-col border-r border-slate-200 bg-white md:flex">
        <div className="flex items-center gap-2 border-b border-slate-200 px-5 py-4">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-slate-900 text-white">
            <Boxes size={18} />
          </div>
          <div>
            <p className="text-sm font-semibold leading-tight text-slate-900">HCV Program</p>
            <p className="text-xs leading-tight text-slate-500">Inventory & Assets</p>
          </div>
        </div>
        <SidebarNav activeTab={activeTab} setActiveTab={setActiveTab} saving={saving} connectionError={connectionError} />
      </aside>

      <main className="min-w-0 flex-1 overflow-y-auto overflow-x-hidden">
        <div className="mx-auto max-w-6xl p-3 sm:p-6">
          {connectionError && (
            <div className="mb-4 flex items-start gap-2 rounded-md border border-amber-200 bg-amber-50 px-3 py-2.5 text-sm text-amber-800">
              <AlertTriangle size={16} className="mt-0.5 shrink-0" />
              <span>{connectionError}</span>
            </div>
          )}
          {activeTab === "dashboard" && <Dashboard data={data} setActiveTab={setActiveTab} />}
          {activeTab === "inventory" && <InventoryTab data={data} update={update} />}
          {activeTab === "transactions" && <TransactionsTab data={data} />}
          {activeTab === "warehouses" && <WarehousesTab data={data} update={update} />}
          {activeTab === "tablets" && <TabletsTab data={data} update={update} />}
          {activeTab === "staff" && <StaffTab data={data} update={update} />}
          {activeTab === "settings" && <SettingsTab data={data} update={update} />}
        </div>
      </main>
    </div>
  );
}

function SidebarNav({ activeTab, setActiveTab, saving, connectionError }) {
  return (
    <>
      <nav className="flex-1 space-y-0.5 overflow-y-auto p-3">
        {TABS.map((t) => {
          const Icon = t.icon;
          const active = activeTab === t.id;
          return (
            <button
              key={t.id}
              onClick={() => setActiveTab(t.id)}
              className={cx(
                "flex w-full items-center gap-2.5 rounded-md px-3 py-2.5 text-sm font-medium transition-colors",
                active ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-100"
              )}
            >
              <Icon size={16} />
              {t.label}
            </button>
          );
        })}
      </nav>
      <div className={cx("border-t border-slate-200 px-4 py-3 text-xs", connectionError ? "text-amber-600" : "text-slate-400")}>
        {connectionError ? "Offline - not connected" : saving ? "Saving to database..." : "Connected - all changes saved"}
      </div>
    </>
  );
}

/* ---------------------------------------------------------------------- */
/* Dashboard                                                                */
/* ---------------------------------------------------------------------- */

function Dashboard({ data, setActiveTab }) {
  const stockByItem = useMemo(() => {
    const map = {};
    data.stockLevels.forEach((s) => { map[s.itemId] = (map[s.itemId] || 0) + s.quantity; });
    return map;
  }, [data.stockLevels]);

  const totalItems = data.items.length;
  const totalStockQty = Object.values(stockByItem).reduce((a, b) => a + b, 0);
  const lowStockItems = data.items.filter((it) => (stockByItem[it.id] || 0) < (it.minThreshold ?? 0));
  const tabletsAssigned = data.tablets.filter((t) => t.status === "Assigned").length;

  const byCategory = useMemo(() => {
    return data.categories.map((c) => ({
      name: c.name,
      qty: data.items.filter((i) => i.categoryId === c.id).reduce((sum, i) => sum + (stockByItem[i.id] || 0), 0),
    })).filter((c) => c.qty > 0);
  }, [data.categories, data.items, stockByItem]);

  const byWarehouse = useMemo(() => {
    return data.hospitals.map((w) => ({
      name: w.name,
      qty: data.stockLevels.filter((s) => s.hospitalId === w.id).reduce((sum, s) => sum + s.quantity, 0),
    })).filter((w) => w.qty > 0);
  }, [data.hospitals, data.stockLevels]);

  const COLORS = ["#0F6E56", "#185FA5", "#854F0B", "#993C1D", "#534AB7", "#3B6D11", "#72243E"];

  return (
    <div className="flex flex-col gap-6">
      <div>
        <h1 className="text-xl font-semibold text-slate-900">Dashboard</h1>
        <p className="text-sm text-slate-500">Program-wide overview of inventory, warehouses and tablet assets.</p>
      </div>

      <div className="grid grid-cols-2 gap-4 md:grid-cols-4">
        <StatCard label="Inventory items" value={totalItems} icon={Package} />
        <StatCard label="Total stock quantity" value={totalStockQty.toLocaleString()} icon={Boxes} />
        <StatCard label="Low stock alerts" value={lowStockItems.length} icon={AlertTriangle} tone={lowStockItems.length ? "red" : "slate"} onClick={() => setActiveTab("inventory")} />
        <StatCard label="Tablets assigned" value={`${tabletsAssigned} / ${data.tablets.length}`} icon={TabletIcon} onClick={() => setActiveTab("tablets")} />
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <h3 className="mb-3 text-sm font-semibold text-slate-800">Stock quantity by category</h3>
          {byCategory.length === 0 ? (
            <EmptyState icon={Package} title="No stock recorded yet" subtitle="Receive stock against an item to see category breakdown." />
          ) : (
            <ResponsiveContainer width="100%" height={260}>
              <BarChart data={byCategory} layout="vertical" margin={{ left: 10, right: 20 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e2e8f0" />
                <XAxis type="number" tick={{ fontSize: 12 }} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={110} />
                <Tooltip />
                <Bar dataKey="qty" fill="#185FA5" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>

        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <h3 className="mb-3 text-sm font-semibold text-slate-800">Stock distribution by warehouse</h3>
          {byWarehouse.length === 0 ? (
            <EmptyState icon={Building2} title="No stock recorded yet" />
          ) : (
            <ResponsiveContainer width="100%" height={260}>
              <PieChart>
                <Pie data={byWarehouse} dataKey="qty" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>
                  {byWarehouse.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                </Pie>
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: 12 }} />
              </PieChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>

      <div className="rounded-lg border border-slate-200 bg-white p-4">
        <div className="mb-3 flex items-center justify-between">
          <h3 className="text-sm font-semibold text-slate-800">Low stock items</h3>
          {lowStockItems.length > 0 && <Badge tone="red">{lowStockItems.length} below threshold</Badge>}
        </div>
        {lowStockItems.length === 0 ? (
          <p className="py-6 text-center text-sm text-slate-400">All items are above their minimum threshold.</p>
        ) : (
          <table className="w-full min-w-[720px] text-sm">
            <thead>
              <tr className="border-b border-slate-200 text-left text-xs uppercase text-slate-400">
                <th className="py-2">Item</th><th>SKU</th><th>Current qty</th><th>Threshold</th>
              </tr>
            </thead>
            <tbody>
              {lowStockItems.map((it) => (
                <tr key={it.id} className="border-b border-slate-100 last:border-0">
                  <td className="py-2 font-medium text-slate-700">{it.name}</td>
                  <td className="text-slate-500">{it.sku}</td>
                  <td className="text-red-600 font-medium">{stockByItem[it.id] || 0}</td>
                  <td className="text-slate-500">{it.minThreshold}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function StatCard({ label, value, icon: Icon, tone = "slate", onClick }) {
  const tones = { slate: "text-slate-900", red: "text-red-600" };
  return (
    <button onClick={onClick} className={cx("flex flex-col gap-2 rounded-lg border border-slate-200 bg-white p-4 text-left", onClick && "hover:border-slate-300 cursor-pointer")}>
      <div className="flex items-center justify-between">
        <span className="text-xs font-medium text-slate-500">{label}</span>
        <Icon size={16} className="text-slate-400" />
      </div>
      <span className={cx("text-2xl font-semibold", tones[tone])}>{value}</span>
    </button>
  );
}

/* ---------------------------------------------------------------------- */
/* Toolbar (search)                                                        */
/* ---------------------------------------------------------------------- */

function Toolbar({ title, subtitle, search, setSearch, children }) {
  return (
    <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
      <div>
        <h1 className="text-xl font-semibold text-slate-900">{title}</h1>
        {subtitle && <p className="text-sm text-slate-500">{subtitle}</p>}
      </div>
      <div className="flex flex-wrap items-center gap-2">
        {setSearch && (
          <div className="relative w-full sm:w-auto">
            <Search size={14} className="absolute left-2.5 top-2.5 text-slate-400" />
            <TextInput placeholder="Search..." value={search} onChange={(e) => setSearch(e.target.value)} className="w-full pl-8 sm:w-56" />
          </div>
        )}
        {children}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Inventory Items Tab                                                     */
/* ---------------------------------------------------------------------- */

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
    data.stockLevels.forEach((s) => { map[s.itemId] = (map[s.itemId] || 0) + s.quantity; });
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
      const items = exists ? prev.items.map((i) => (i.id === item.id ? item : i)) : [...prev.items, item];
      let stockLevels = [...prev.stockLevels];
      let transactions = [...prev.transactions];

      if (openingStock) {
        const idx = stockLevels.findIndex((s) => s.itemId === item.id && s.hospitalId === openingStock.warehouseId);
        if (idx >= 0) stockLevels[idx] = { ...stockLevels[idx], quantity: stockLevels[idx].quantity + openingStock.qty };
        else stockLevels.push({ id: uid("stk"), itemId: item.id, hospitalId: openingStock.warehouseId, quantity: openingStock.qty });
        transactions = [{
          id: uid("txn"), itemId: item.id, itemName: item.name, type: "RECEIVE",
          qty: openingStock.qty, fromHospitalId: null, toHospitalId: openingStock.warehouseId,
          date: todayISO(), remarks: "Opening stock recorded with item creation",
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
            id: uid("txn"), itemId: item.id, itemName: item.name, type: "ADJUST",
            qty: delta, fromHospitalId: null, toHospitalId: hospitalId,
            date: todayISO(), remarks: "Quantity edited directly from Inventory Items",
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
      stockLevels: prev.stockLevels.filter((s) => s.itemId !== id),
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

  return (
    <div>
      <Toolbar title="Inventory items" subtitle="Stock & supplies catalog (RDT kits, printers, scanners, etc). For individually tracked tablets/devices with IMEI, use the 'Tablet Assets' tab instead." search={search} setSearch={setSearch}>
        <Select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} className="w-full sm:w-48">
          <option value="">All categories</option>
          {data.categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
        </Select>
        <Btn variant="primary" onClick={() => { setEditingItem(null); setShowForm(true); }}><Plus size={15} /> Add item</Btn>
      </Toolbar>

      {data.items.length === 0 ? (
        <EmptyState icon={Package} title="No items yet" subtitle="Add your first inventory item to get started." />
      ) : (
        <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
          <table className="w-full min-w-[720px] text-sm">
            <thead className="bg-slate-50 text-left text-xs uppercase text-slate-400">
              <tr>
                <th className="w-6 py-2.5 px-4"></th>
                <th className="py-2.5">Item</th><th>SKU / Code</th><th>Category</th><th>Unit</th>
                <th>Total qty</th><th className="text-right pr-4">Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((it) => {
                const qty = stockByItem[it.id] || 0;
                const low = qty < (it.minThreshold ?? 0);
                const isOpen = expanded === it.id;
                return (
                  <Fragment key={it.id}>
                    <tr className="border-t border-slate-100 hover:bg-slate-50">
                      <td className="pl-4">
                        <button onClick={() => setExpanded(isOpen ? null : it.id)} className="text-slate-400 hover:text-slate-700">
                          <ChevronDown size={15} className={cx("transition-transform", isOpen && "rotate-180")} />
                        </button>
                      </td>
                      <td className="py-2.5">
                        <p className="font-medium text-slate-800">{it.name}</p>
                        {it.serialNumber && <p className="text-xs text-slate-400">S/N: {it.serialNumber}</p>}
                      </td>
                      <td className="text-slate-500">{it.sku} {it.itemCode && <span className="text-slate-300">/ {it.itemCode}</span>}</td>
                      <td><Badge>{categoryName(it.categoryId)}</Badge></td>
                      <td className="text-slate-500">{it.unit}</td>
                      <td className={cx("font-medium", low ? "text-red-600" : "text-slate-700")}>
                        {qty} {low && <AlertTriangle size={13} className="inline ml-1" />}
                      </td>
                      <td className="pr-4">
                        <div className="flex justify-end gap-1">
                          <Btn className="py-1 px-2 text-xs" onClick={() => setTxItem(it)}>Record txn</Btn>
                          <button title="Edit item" className="rounded p-1.5 text-slate-400 hover:bg-slate-200 hover:text-slate-700" onClick={() => { setEditingItem(it); setShowForm(true); }}><Pencil size={14} /></button>
                          <button title="Delete item" className="rounded p-1.5 text-slate-400 hover:bg-red-100 hover:text-red-600" onClick={() => setConfirmDeleteId(it.id)}><Trash2 size={14} /></button>
                        </div>
                      </td>
                    </tr>
                    {isOpen && (
                      <tr className="bg-slate-50/70 border-t border-slate-100">
                        <td colSpan={7} className="px-4 py-3">
                          <p className="mb-2 text-xs font-semibold uppercase text-slate-400">Stock by warehouse / hospital</p>
                          <div className="flex flex-wrap gap-2">
                            {data.hospitals.map((w) => {
                              const lvl = data.stockLevels.find((s) => s.itemId === it.id && s.hospitalId === w.id);
                              if (!lvl || lvl.quantity === 0) return null;
                              return <Badge key={w.id} tone="blue">{w.name}: {lvl.quantity}</Badge>;
                            })}
                            {qty === 0 && <span className="text-xs text-slate-400">No stock recorded anywhere yet.</span>}
                          </div>
                        </td>
                      </tr>
                    )}
                  </Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {showForm && (
        <ItemFormModal item={editingItem} categories={data.categories} units={data.units} hospitals={data.hospitals} stockLevels={data.stockLevels} onSave={saveItem} onClose={() => { setShowForm(false); setEditingItem(null); }} />
      )}
      {txItem && (
        <TransactionModal item={txItem} hospitals={data.hospitals} onSave={recordTransaction} onClose={() => setTxItem(null)} />
      )}
      {confirmDeleteId && (
        <ConfirmDialog
          message="Delete this item? Related stock levels will also be removed. This cannot be undone."
          onCancel={() => setConfirmDeleteId(null)}
          onConfirm={() => { deleteItem(confirmDeleteId); setConfirmDeleteId(null); }}
        />
      )}
    </div>
  );
}

function ItemFormModal({ item, categories, units, hospitals, stockLevels, onSave, onClose }) {
  const [form, setForm] = useState(item || {
    id: uid("item"), sku: "", itemCode: "", name: "", categoryId: categories[0]?.id || "",
    unit: units[0] || "", serialNumber: "", minThreshold: 0, image: "",
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

  return (
    <Modal title={item ? "Edit item" : "Add inventory item"} onClose={onClose}>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <Field label="Item name" required className="col-span-2">
          <TextInput value={form.name} onChange={set("name")} placeholder="e.g. HCV RDT Kit" autoFocus />
        </Field>
        <Field label="Item code" required>
          <TextInput value={form.itemCode} onChange={set("itemCode")} placeholder="e.g. RDT-001" />
        </Field>
        <Field label="Category" required>
          <Select value={form.categoryId} onChange={set("categoryId")}>
            {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </Select>
        </Field>
        <Field label="Unit" required className="col-span-2">
          <Select value={form.unit} onChange={set("unit")}>
            {units.map((u) => <option key={u} value={u}>{u}</option>)}
          </Select>
        </Field>
      </div>

      {isNew && (
        <div className="col-span-2 mt-3 rounded-md border border-slate-200 bg-slate-50 p-3">
          <p className="mb-2 text-xs font-semibold uppercase text-slate-400">Opening stock (optional)</p>
          {hospitals.length === 0 ? (
            <p className="text-xs text-amber-600">No warehouse or hospital configured yet. Add one from "Hospitals & Warehouses" first, or leave this blank and use "Record txn" later.</p>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <Field label="Quantity">
                <TextInput type="number" min={0} value={openingQty} onChange={(e) => setOpeningQty(Number(e.target.value))} />
              </Field>
              <Field label="Into warehouse / hospital">
                <Select value={openingWarehouseId} onChange={(e) => setOpeningWarehouseId(e.target.value)}>
                  {hospitals.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
                </Select>
              </Field>
            </div>
          )}
        </div>
      )}

      {!isNew && (
        <div className="mt-3 rounded-md border border-slate-200 bg-slate-50 p-3">
          <p className="mb-2 text-xs font-semibold uppercase text-slate-400">Stock by warehouse</p>
          {hospitals.length === 0 ? (
            <p className="text-xs text-amber-600">No warehouse or hospital configured yet.</p>
          ) : (
            <div className="flex flex-col gap-2 max-h-48 overflow-y-auto pr-1">
              {hospitals.map((h) => (
                <div key={h.id} className="flex items-center justify-between gap-3">
                  <span className="text-sm text-slate-700">{h.name}</span>
                  <TextInput
                    type="number" min={0} value={stockEdits[h.id] ?? 0}
                    onChange={(e) => setStockFor(h.id, Number(e.target.value))}
                    className="w-28 text-right"
                  />
                </div>
              ))}
            </div>
          )}
          <p className="mt-2 text-xs text-slate-400">Changing a number here directly adjusts stock and records an "Adjust Stock" transaction automatically.</p>
        </div>
      )}

      <button type="button" onClick={() => setShowAdvanced((v) => !v)} className="mt-3 flex items-center gap-1 text-xs font-medium text-slate-500 hover:text-slate-700">
        <ChevronDown size={13} className={cx("transition-transform", showAdvanced && "rotate-180")} />
        {showAdvanced ? "Hide" : "Show"} more fields (SKU, serial number, threshold, image)
      </button>

      {showAdvanced && (
        <div className="mt-2 grid grid-cols-2 gap-3 rounded-md border border-slate-200 p-3">
          <Field label="SKU">
            <TextInput value={form.sku} onChange={set("sku")} placeholder="Defaults to item code" />
          </Field>
          <Field label="Serial number">
            <TextInput value={form.serialNumber} onChange={set("serialNumber")} placeholder="If applicable" />
          </Field>
          <Field label="Minimum stock threshold">
            <TextInput type="number" min={0} value={form.minThreshold} onChange={(e) => setForm((f) => ({ ...f, minThreshold: Number(e.target.value) }))} />
          </Field>
          <Field label="Image URL">
            <TextInput value={form.image} onChange={set("image")} placeholder="https://..." />
          </Field>
        </div>
      )}

      <div className="mt-5 flex justify-end gap-2">
        <Btn onClick={onClose}>Cancel</Btn>
        <Btn
          variant="primary"
          disabled={!form.name || !form.itemCode || !form.categoryId}
          onClick={() => onSave(
            { ...form, sku: form.sku || form.itemCode },
            isNew && openingQty > 0 ? { qty: openingQty, warehouseId: openingWarehouseId } : null,
            !isNew ? stockEdits : null
          )}
        >
          Save item
        </Btn>
      </div>
    </Modal>
  );
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

  return (
    <Modal title={`Record transaction - ${item.name}`} onClose={onClose}>
      {hospitals.length === 0 && (
        <div className="mb-3 rounded-md bg-amber-50 px-3 py-2 text-xs text-amber-700">
          No warehouse or hospital configured yet. Add one from "Hospitals & Warehouses" before recording a transaction.
        </div>
      )}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <Field label="Transaction type" required className="col-span-2">
          <Select value={type} onChange={(e) => setType(e.target.value)}>
            {TX_TYPES.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
          </Select>
        </Field>
        <Field label="Quantity" required>
          <TextInput type="number" min={1} value={qty} onChange={(e) => setQty(Number(e.target.value))} />
        </Field>
        <Field label="Date">
          <TextInput type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </Field>
        {needsFrom && (
          <Field label="From warehouse / hospital" required className="col-span-2">
            <Select value={fromHospitalId} onChange={(e) => setFromHospitalId(e.target.value)}>
              {hospitals.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
            </Select>
          </Field>
        )}
        {needsTo && (
          <Field label="To warehouse / hospital" required className="col-span-2">
            <Select value={toHospitalId} onChange={(e) => setToHospitalId(e.target.value)}>
              {hospitals.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
            </Select>
          </Field>
        )}
        <Field label="Remarks" className="col-span-2">
          <TextArea rows={2} value={remarks} onChange={(e) => setRemarks(e.target.value)} placeholder="Optional notes" />
        </Field>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <Btn onClick={onClose}>Cancel</Btn>
        <Btn variant="primary" disabled={!qty || qty <= 0 || hospitals.length === 0} onClick={() => onSave({ itemId: item.id, itemName: item.name, type, qty: Number(qty), fromHospitalId: needsFrom ? fromHospitalId : null, toHospitalId: needsTo ? toHospitalId : null, date, remarks })}>
          Save transaction
        </Btn>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/* Transactions Tab                                                        */
/* ---------------------------------------------------------------------- */

function TransactionsTab({ data }) {
  const [search, setSearch] = useState("");
  const [typeFilter, setTypeFilter] = useState("");

  const hospitalName = (id) => data.hospitals.find((h) => h.id === id)?.name || "-";

  const filtered = data.transactions.filter((tx) => {
    const matchesSearch = !search || tx.itemName.toLowerCase().includes(search.toLowerCase());
    const matchesType = !typeFilter || tx.type === typeFilter;
    return matchesSearch && matchesType;
  });

  return (
    <div>
      <Toolbar title="Stock transactions" subtitle="Full audit trail of every stock movement." search={search} setSearch={setSearch}>
        <Select value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)} className="w-full sm:w-48">
          <option value="">All types</option>
          {TX_TYPES.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
        </Select>
      </Toolbar>

      {filtered.length === 0 ? (
        <EmptyState icon={ArrowLeftRight} title="No transactions recorded" subtitle="Record a transaction from the Inventory Items tab." />
      ) : (
        <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
          <table className="w-full min-w-[720px] text-sm">
            <thead className="bg-slate-50 text-left text-xs uppercase text-slate-400">
              <tr><th className="py-2.5 pl-4">Date</th><th>Item</th><th>Type</th><th>Qty</th><th>From</th><th>To</th><th className="pr-4">Remarks</th></tr>
            </thead>
            <tbody>
              {filtered.map((tx) => {
                const meta = TX_TYPES.find((t) => t.id === tx.type);
                const Icon = meta?.icon || RefreshCcw;
                return (
                  <tr key={tx.id} className="border-t border-slate-100">
                    <td className="py-2.5 pl-4 text-slate-500">{fmtDate(tx.date)}</td>
                    <td className="font-medium text-slate-700">{tx.itemName}</td>
                    <td>
                      <span className="inline-flex items-center gap-1.5 text-xs font-medium" style={{ color: meta?.color }}>
                        <Icon size={14} /> {meta?.label}
                      </span>
                    </td>
                    <td className="text-slate-700">{tx.qty}</td>
                    <td className="text-slate-500">{tx.fromHospitalId ? hospitalName(tx.fromHospitalId) : "-"}</td>
                    <td className="text-slate-500">{tx.toHospitalId ? hospitalName(tx.toHospitalId) : "-"}</td>
                    <td className="pr-4 text-slate-400">{tx.remarks || "-"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Warehouses / Hospitals Tab                                              */
/* ---------------------------------------------------------------------- */

function WarehousesTab({ data, update }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);

  function save(w) {
    update((prev) => {
      const exists = prev.hospitals.some((h) => h.id === w.id);
      return { ...prev, hospitals: exists ? prev.hospitals.map((h) => (h.id === w.id ? w : h)) : [...prev.hospitals, w] };
    });
    setShowForm(false); setEditing(null);
  }

  function remove(id) {
    update((prev) => ({ ...prev, hospitals: prev.hospitals.filter((h) => h.id !== id) }));
  }

  return (
    <div>
      <Toolbar title="Hospitals & warehouses" subtitle="Configure the warehouse hierarchy - Central, Regional, and Hospital level.">
        <Btn variant="primary" onClick={() => { setEditing(null); setShowForm(true); }}><Plus size={15} /> Add location</Btn>
      </Toolbar>

      {data.hospitals.length === 0 ? (
        <EmptyState icon={Building2} title="No warehouses or hospitals configured" />
      ) : (
        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
          {data.hospitals.map((w) => (
            <div key={w.id} className="rounded-lg border border-slate-200 bg-white p-4">
              <div className="mb-2 flex items-start justify-between">
                <Badge tone={w.type === "Central Warehouse" ? "blue" : w.type === "Regional Warehouse" ? "amber" : "green"}>{w.type}</Badge>
                <div className="flex gap-1">
                  <button title="Edit" className="rounded p-1 text-slate-400 hover:bg-slate-200 hover:text-slate-700" onClick={() => { setEditing(w); setShowForm(true); }}><Pencil size={13} /></button>
                  <button title="Delete" className="rounded p-1 text-slate-400 hover:bg-red-100 hover:text-red-600" onClick={() => setConfirmDeleteId(w.id)}><Trash2 size={13} /></button>
                </div>
              </div>
              <p className="font-medium text-slate-800">{w.name}</p>
              <p className="text-xs text-slate-400">{w.location}</p>
            </div>
          ))}
        </div>
      )}

      {showForm && <WarehouseFormModal warehouse={editing} onSave={save} onClose={() => { setShowForm(false); setEditing(null); }} />}
      {confirmDeleteId && (
        <ConfirmDialog
          message="Delete this warehouse / hospital? This cannot be undone."
          onCancel={() => setConfirmDeleteId(null)}
          onConfirm={() => { remove(confirmDeleteId); setConfirmDeleteId(null); }}
        />
      )}
    </div>
  );
}

function WarehouseFormModal({ warehouse, onSave, onClose }) {
  const [form, setForm] = useState(warehouse || { id: uid("wh"), name: "", type: WAREHOUSE_TYPES[0], location: "" });
  return (
    <Modal title={warehouse ? "Edit location" : "Add warehouse / hospital"} onClose={onClose}>
      <div className="flex flex-col gap-3">
        <Field label="Name" required><TextInput value={form.name} onChange={(e) => setForm((f) => ({ ...f, name: e.target.value }))} placeholder="e.g. DHQ Hospital Multan" /></Field>
        <Field label="Type" required>
          <Select value={form.type} onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}>
            {WAREHOUSE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </Select>
        </Field>
        <Field label="Location / city"><TextInput value={form.location} onChange={(e) => setForm((f) => ({ ...f, location: e.target.value }))} /></Field>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <Btn onClick={onClose}>Cancel</Btn>
        <Btn variant="primary" disabled={!form.name} onClick={() => onSave(form)}>Save</Btn>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/* Tablets Tab - includes Device Handover / Takeover custody workflow      */
/* ---------------------------------------------------------------------- */

const CUSTODY_TYPES = {
  TAKEOVER_SOURCE: { label: "Takeover (from source)", icon: Inbox, color: "#0F6E56", tone: "green" },
  HANDOVER: { label: "Handover (to hospital)", icon: Send, color: "#185FA5", tone: "blue" },
  TAKEOVER_HOSPITAL: { label: "Takeover (from hospital)", icon: ClipboardCheck, color: "#854F0B", tone: "amber" },
  REPLACED_OUT: { label: "Replaced - removed from hospital", icon: RefreshCcw, color: "#A32D2D", tone: "red" },
  REPLACED_IN: { label: "Replaced - issued to hospital", icon: RefreshCcw, color: "#0F6E56", tone: "green" },
  ADDED: { label: "Record created", icon: Plus, color: "#5F5E5A", tone: "slate" },
  EDITED: { label: "Details edited", icon: Pencil, color: "#5F5E5A", tone: "slate" },
};

function custodyMeta(type) {
  return CUSTODY_TYPES[type] || { label: type || "Event", icon: History, color: "#5F5E5A", tone: "slate" };
}

function TabletsTab({ data, update }) {
  const [view, setView] = useState("devices"); // 'devices' | 'log'
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [takeoverSourceTablet, setTakeoverSourceTablet] = useState(null); // null=closed, {} = new device, obj = existing
  const [handoverTablet, setHandoverTablet] = useState(null);
  const [takeoverHospitalTablet, setTakeoverHospitalTablet] = useState(null);
  const [showLetterModal, setShowLetterModal] = useState(false);
  const [showReplacementModal, setShowReplacementModal] = useState(false);
  const [expanded, setExpanded] = useState(null);
  const [importSummary, setImportSummary] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const fileInputRef = useRef(null);

  const hospitalName = (id) => data.hospitals.find((h) => h.id === id)?.name || "-";
  const staffName = (id) => data.staff.find((s) => s.id === id)?.name || "-";

  const filtered = data.tablets.filter((t) => !search || [t.serialNumber, t.model].join(" ").toLowerCase().includes(search.toLowerCase()));

  // Import devices in bulk from an Excel/CSV file: IMEI, App Installed, Status, Hospital Name, Handed Over To Person
  function processImportRows(rows) {
    const warnings = [];
    let created = 0, updated = 0;

    const KNOWN_FIELD_ALIASES = {
      imei: ["imei", "imei no", "imei number", "imei/serial no", "serial number", "serial no", "serial", "device serial"],
      appInstalled: ["app installed", "nadra app", "nadra app installed", "app"],
      status: ["status", "device status", "condition"],
      hospitalName: ["hospital name", "hospital", "facility", "facility name"],
      handedTo: ["handed over to person", "handed over to", "handed to", "assigned staff", "assigned to", "taken over by", "person"],
    };

    update((prev) => {
      let tablets = [...prev.tablets];
      rows.forEach((row, idx) => {
        const rowKeys = Object.keys(row);
        const consumedKeys = new Set();

        // Exact (normalized) match first, then fuzzy "contains" match as a fallback.
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
          return val === undefined || val === null ? "" : String(val).trim();
        };

        const imei = get("imei");
        if (!imei) { warnings.push(`Row ${idx + 2}: missing IMEI/Serial column value - skipped.`); return; }

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
          if (matchedStaff) assignedStaffId = matchedStaff.id; else assignedStaffName = handedToRaw;
        }

        const validStatus = TABLET_STATUSES.find((s) => s.toLowerCase() === statusRaw.toLowerCase());
        const status = validStatus || (handedToRaw ? "Assigned" : "In Stock");

        // Anything in the sheet that wasn't one of the 5 known columns is preserved, not dropped.
        const customFields = {};
        rowKeys.forEach((k) => {
          if (consumedKeys.has(k)) return;
          const val = row[k];
          if (val !== undefined && val !== null && String(val).trim() !== "") customFields[k.trim()] = String(val).trim();
        });

        const existingIdx = tablets.findIndex((t) => t.serialNumber.trim().toLowerCase() === imei.toLowerCase());
        const extraNote = Object.keys(customFields).length ? ` Extra columns: ${Object.entries(customFields).map(([k, v]) => `${k}=${v}`).join(", ")}.` : "";
        const historyEntry = {
          date: todayISO(), type: existingIdx >= 0 ? "EDITED" : "ADDED",
          detail: `Imported from Excel.${hospital ? ` Hospital: ${hospital.name}.` : ""}${handedToRaw ? ` Handed to: ${handedToRaw}.` : ""} Status: ${status}.${extraNote}`,
        };

        if (existingIdx >= 0) {
          tablets[existingIdx] = {
            ...tablets[existingIdx], nadraInstalled, status,
            currentHospitalId: hospital?.id || tablets[existingIdx].currentHospitalId,
            assignedStaffId, assignedStaffName,
            customFields: { ...(tablets[existingIdx].customFields || {}), ...customFields },
            history: [historyEntry, ...(tablets[existingIdx].history || [])],
          };
          updated++;
        } else {
          tablets.push({
            id: uid("tab"), serialNumber: imei, model: "IVAS Tablet",
            currentHospitalId: hospital?.id || "", assignedStaffId, assignedStaffName,
            nadraInstalled, batteryHealth: 100, status, remarks: "", customFields, history: [historyEntry],
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
      if (rows.length === 0) { setImportSummary({ created: 0, updated: 0, warnings: ["The file has no data rows."] }); return; }
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
      const history = exists
        ? [{ date: todayISO(), type: "EDITED", detail: "Device details updated" }, ...(t.history || [])]
        : [{ date: todayISO(), type: "ADDED", detail: "Device record created" }];
      return {
        ...prev,
        tablets: exists ? prev.tablets.map((x) => (x.id === t.id ? { ...t, history } : x)) : [...prev.tablets, { ...t, history }],
      };
    });
    setShowForm(false); setEditing(null);
  }

  function remove(id) {
    update((prev) => ({ ...prev, tablets: prev.tablets.filter((x) => x.id !== id) }));
  }

  // Takeover from an external source (e.g. DHO office) into a warehouse, with inspection
  function takeoverFromSource(payload) {
    update((prev) => {
      const conditionOk = payload.overallCondition !== "Damaged";
      const historyEntry = {
        date: payload.date, type: "TAKEOVER_SOURCE",
        detail: `Received from ${payload.source}${payload.receivedBy ? ` by ${payload.receivedBy}` : ""} into ${hospitalName(payload.warehouseId)}. Condition: ${payload.overallCondition}. Checklist passed: ${payload.checklist.filter(Boolean).length}/${payload.checklistLabels.length}.${payload.remarks ? " Remarks: " + payload.remarks : ""}`,
      };
      const existing = prev.tablets.find((t) => t.id === payload.tabletId);
      if (existing) {
        return {
          ...prev,
          tablets: prev.tablets.map((t) => t.id !== payload.tabletId ? t : {
            ...t, currentHospitalId: payload.warehouseId, assignedStaffId: null,
            status: conditionOk ? "In Stock" : "Damaged",
            history: [historyEntry, ...(t.history || [])],
          }),
        };
      }
      const newTablet = {
        id: uid("tab"), serialNumber: payload.serialNumber, model: payload.model,
        currentHospitalId: payload.warehouseId, assignedStaffId: null,
        nadraInstalled: false, batteryHealth: payload.batteryHealth || 100,
        status: conditionOk ? "In Stock" : "Damaged", remarks: payload.remarks,
        history: [historyEntry],
      };
      return { ...prev, tablets: [...prev.tablets, newTablet] };
    });
    setTakeoverSourceTablet(null);
  }

  // Handover: dispatch a device from warehouse to a hospital + staff member
  function handover(tablet, payload) {
    update((prev) => ({
      ...prev,
      tablets: prev.tablets.map((t) => t.id !== tablet.id ? t : {
        ...t, currentHospitalId: payload.hospitalId, assignedStaffId: payload.staffId, status: "Assigned",
        history: [{
          date: payload.date, type: "HANDOVER",
          detail: `Handed over to ${hospitalName(payload.hospitalId)} / ${staffName(payload.staffId)}${payload.handedOverBy ? ` by ${payload.handedOverBy}` : ""}. Condition: ${payload.condition}. Accessories given: ${payload.accessories.join(", ") || "None"}.${payload.remarks ? " Remarks: " + payload.remarks : ""}`,
        }, ...(t.history || [])],
      }),
    }));
    setHandoverTablet(null);
  }

  // Takeover from hospital: device returns from hospital back into custody
  function takeoverFromHospital(tablet, payload) {
    update((prev) => ({
      ...prev,
      tablets: prev.tablets.map((t) => t.id !== tablet.id ? t : {
        ...t, assignedStaffId: null, status: payload.condition === "Damaged" ? "Damaged" : "In Stock",
        remarks: payload.remarks,
        history: [{
          date: payload.date, type: "TAKEOVER_HOSPITAL",
          detail: `Taken over from ${hospitalName(t.currentHospitalId)}${payload.receivedBy ? ` by ${payload.receivedBy}` : ""}. Condition: ${payload.condition}. Accessories returned: ${payload.accessories.join(", ") || "None"}.${payload.inspectionNotes ? " Inspection: " + payload.inspectionNotes : ""}${payload.remarks ? " Remarks: " + payload.remarks : ""}`,
        }, ...(t.history || [])],
      }),
    }));
    setTakeoverHospitalTablet(null);
  }

  // Batch handover via a formal handover letter (multiple devices, one signed document)
  function handoverViaLetter(letter) {
    update((prev) => ({
      ...prev,
      tablets: prev.tablets.map((t) => {
        const row = letter.rows.find((r) => r.tabletId === t.id);
        if (!row) return t;
        return {
          ...t, currentHospitalId: letter.hospitalId, assignedStaffId: letter.takenOverStaffId || null,
          assignedStaffName: letter.takenOverStaffId ? null : letter.takenOverName, status: "Assigned",
          history: [{
            date: row.date, type: "HANDOVER",
            detail: `Handed over to ${hospitalName(letter.hospitalId)} / ${letter.takenOverName}${letter.takenOverDesignation ? ` (${letter.takenOverDesignation})` : ""} by ${letter.handingOverName}${letter.handingOverDesignation ? ` (${letter.handingOverDesignation})` : ""}. Counter: ${row.counter || "-"}. Status: ${row.status || "-"}.${letter.refNo ? ` Ref: ${letter.refNo}.` : ""}`,
          }, ...(t.history || [])],
        };
      }),
    }));
  }

  // Batch device replacement via a formal replacement letter (old device swapped for new device)
  function replaceViaLetter(letter) {
    update((prev) => ({
      ...prev,
      tablets: prev.tablets.map((t) => {
        const asOld = letter.rows.find((r) => r.oldTabletId === t.id);
        const asNew = letter.rows.find((r) => r.newTabletId === t.id);
        const newSerial = (id) => prev.tablets.find((x) => x.id === id)?.serialNumber || "-";

        if (asOld) {
          return {
            ...t, status: asOld.oldStatus || "Damaged", assignedStaffId: null, assignedStaffName: null,
            history: [{
              date: asOld.date, type: "REPLACED_OUT",
              detail: `Replaced with new device IMEI ${newSerial(asOld.newTabletId)} at ${hospitalName(letter.hospitalId)}. Reason: ${asOld.reason || "-"}. Processed by ${letter.handingOverName}.${letter.refNo ? ` Ref: ${letter.refNo}.` : ""}`,
            }, ...(t.history || [])],
          };
        }
        if (asNew) {
          return {
            ...t, currentHospitalId: letter.hospitalId, assignedStaffId: letter.takenOverStaffId || null,
            assignedStaffName: letter.takenOverStaffId ? null : letter.takenOverName, status: "Assigned",
            history: [{
              date: asNew.date, type: "REPLACED_IN",
              detail: `Issued to ${hospitalName(letter.hospitalId)} / ${letter.takenOverName} replacing old device IMEI ${newSerial(asNew.oldTabletId)}. Reason: ${asNew.reason || "-"}. Counter: ${asNew.counter || "-"}.${letter.refNo ? ` Ref: ${letter.refNo}.` : ""}`,
            }, ...(t.history || [])],
          };
        }
        return t;
      }),
    }));
  }

  const staffLabel = (t) => t.assignedStaffId ? staffName(t.assignedStaffId) : (t.assignedStaffName || "-");

  const allLogEntries = useMemo(() => {
    const rows = [];
    data.tablets.forEach((t) => (t.history || []).forEach((h) => rows.push({ ...h, tablet: t })));
    return rows.sort((a, b) => (a.date < b.date ? 1 : -1));
  }, [data.tablets]);

  const filteredLog = allLogEntries.filter((e) => !search || [e.tablet.serialNumber, e.detail].join(" ").toLowerCase().includes(search.toLowerCase()));

  return (
    <div>
      <Toolbar
        title="Tablet assets - handover / takeover"
        subtitle="Full chain of custody: receive from source, inspect, hand over to hospitals, and take back."
        search={search} setSearch={setSearch}
      >
        <Btn variant="primary" onClick={() => setTakeoverSourceTablet({})}><Inbox size={15} /> Takeover device</Btn>
        <Btn onClick={() => setShowLetterModal(true)}><FileText size={15} /> Handover letter</Btn>
        <Btn onClick={() => setShowReplacementModal(true)}><RefreshCcw size={15} /> Replacement letter</Btn>
        <Btn onClick={() => fileInputRef.current?.click()}><Upload size={15} /> Import from Excel</Btn>
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleImportFile} />
        <Btn onClick={() => { setEditing(null); setShowForm(true); }}><Plus size={15} /> Add manually</Btn>
      </Toolbar>

      <p className="mb-3 text-xs text-slate-400">
        Recognized columns: <span className="font-medium text-slate-500">IMEI, App Installed, Status, Hospital Name, Handed Over To Person</span> (column names are matched flexibly, so slight variations are fine). Existing devices are matched and updated by IMEI; new IMEIs create new device records. Any extra columns in your file are kept too and shown under each device's custody history.
      </p>

      <div className="mb-4 flex gap-1 rounded-md bg-slate-100 p-1 w-fit">
        <button onClick={() => setView("devices")} className={cx("rounded px-3 py-1.5 text-sm font-medium", view === "devices" ? "bg-white shadow-sm text-slate-900" : "text-slate-500")}>Devices</button>
        <button onClick={() => setView("log")} className={cx("rounded px-3 py-1.5 text-sm font-medium", view === "log" ? "bg-white shadow-sm text-slate-900" : "text-slate-500")}>Handover / takeover log</button>
      </div>

      {importSummary && (
        <div className="mb-4 rounded-md border border-blue-200 bg-blue-50 p-3 text-sm text-blue-800">
          <div className="flex items-start justify-between gap-2">
            <p className="font-medium">
              Import complete: {importSummary.created} device{importSummary.created === 1 ? "" : "s"} created, {importSummary.updated} updated.
            </p>
            <button onClick={() => setImportSummary(null)} className="shrink-0 text-blue-400 hover:text-blue-700"><X size={14} /></button>
          </div>
          {importSummary.warnings.length > 0 && (
            <ul className="mt-1.5 list-disc pl-5 text-xs text-amber-700">
              {importSummary.warnings.map((w, i) => <li key={i}>{w}</li>)}
            </ul>
          )}
        </div>
      )}

      {view === "devices" ? (
        data.tablets.length === 0 ? (
          <EmptyState icon={TabletIcon} title="No devices registered" subtitle="Use 'Takeover device' to record a device received from DHO office or another source." />
        ) : (
          <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
            <table className="w-full min-w-[720px] text-sm">
              <thead className="bg-slate-50 text-left text-xs uppercase text-slate-400">
                <tr>
                  <th className="w-6 py-2.5 pl-4"></th>
                  <th>Serial number</th><th>Status</th><th>Current location</th><th>Assigned staff</th><th>NADRA app</th><th>Battery</th><th className="text-right pr-4">Actions</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((t) => {
                  const isOpen = expanded === t.id;
                  return (
                    <Fragment key={t.id}>
                      <tr className="border-t border-slate-100 hover:bg-slate-50">
                        <td className="pl-4">
                          <button onClick={() => setExpanded(isOpen ? null : t.id)} className="text-slate-400 hover:text-slate-700">
                            <ChevronDown size={15} className={cx("transition-transform", isOpen && "rotate-180")} />
                          </button>
                        </td>
                        <td className="py-2.5 font-medium text-slate-700">{t.serialNumber}<p className="text-xs font-normal text-slate-400">{t.model}</p></td>
                        <td><Badge tone={t.status === "Assigned" ? "blue" : t.status === "In Stock" ? "green" : t.status === "Damaged" || t.status === "Lost" ? "red" : "amber"}>{t.status}</Badge></td>
                        <td className="text-slate-600">{t.currentHospitalId ? hospitalName(t.currentHospitalId) : "-"}</td>
                        <td className="text-slate-600">{staffLabel(t)}</td>
                        <td>{t.nadraInstalled ? <Badge tone="green">Installed</Badge> : <Badge>Not installed</Badge>}</td>
                        <td className="text-slate-600">{t.batteryHealth ? `${t.batteryHealth}%` : "-"}</td>
                        <td className="pr-4">
                          <div className="flex justify-end gap-1">
                            {t.status === "Assigned" ? (
                              <Btn className="py-1 px-2 text-xs" onClick={() => setTakeoverHospitalTablet(t)}><ClipboardCheck size={13} /> Takeover</Btn>
                            ) : (
                              <Btn className="py-1 px-2 text-xs" onClick={() => setHandoverTablet(t)}><Send size={13} /> Handover</Btn>
                            )}
                            <button className="rounded p-1.5 text-slate-400 hover:bg-slate-200 hover:text-slate-700" onClick={() => { setEditing(t); setShowForm(true); }}><Pencil size={14} /></button>
                            <button title="Delete device" className="rounded p-1.5 text-slate-400 hover:bg-red-100 hover:text-red-600" onClick={() => setConfirmDeleteId(t.id)}><Trash2 size={14} /></button>
                          </div>
                        </td>
                      </tr>
                      {isOpen && (
                        <tr className="bg-slate-50/70 border-t border-slate-100">
                          <td colSpan={8} className="px-4 py-3">
                            {t.customFields && Object.keys(t.customFields).length > 0 && (
                              <div className="mb-3">
                                <p className="mb-1.5 text-xs font-semibold uppercase text-slate-400">Additional imported fields</p>
                                <div className="flex flex-wrap gap-1.5">
                                  {Object.entries(t.customFields).map(([k, v]) => (
                                    <Badge key={k} tone="slate">{k}: {v}</Badge>
                                  ))}
                                </div>
                              </div>
                            )}
                            <p className="mb-2 text-xs font-semibold uppercase text-slate-400">Custody history</p>
                            {(!t.history || t.history.length === 0) ? (
                              <p className="text-xs text-slate-400">No custody events recorded yet.</p>
                            ) : (
                              <ol className="flex flex-col gap-2.5">
                                {t.history.map((h, i) => {
                                  const meta = custodyMeta(h.type);
                                  const Icon = meta.icon;
                                  return (
                                    <li key={i} className="flex gap-2.5 text-xs">
                                      <span className="mt-0.5 shrink-0" style={{ color: meta.color }}><Icon size={14} /></span>
                                      <div>
                                        <p><span className="font-medium text-slate-700">{meta.label}</span> <span className="text-slate-400">- {fmtDate(h.date)}</span></p>
                                        <p className="text-slate-500">{h.detail}</p>
                                      </div>
                                    </li>
                                  );
                                })}
                              </ol>
                            )}
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        )
      ) : (
        filteredLog.length === 0 ? (
          <EmptyState icon={History} title="No handover / takeover events yet" subtitle="Events appear here as devices are taken over, handed over, or returned." />
        ) : (
          <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
            <table className="w-full min-w-[720px] text-sm">
              <thead className="bg-slate-50 text-left text-xs uppercase text-slate-400">
                <tr><th className="py-2.5 pl-4">Date</th><th>Device</th><th>Event</th><th className="pr-4">Details</th></tr>
              </thead>
              <tbody>
                {filteredLog.map((e, i) => {
                  const meta = custodyMeta(e.type);
                  const Icon = meta.icon;
                  return (
                    <tr key={i} className="border-t border-slate-100">
                      <td className="py-2.5 pl-4 text-slate-500 whitespace-nowrap">{fmtDate(e.date)}</td>
                      <td className="font-medium text-slate-700 whitespace-nowrap">{e.tablet.serialNumber}</td>
                      <td>
                        <span className="inline-flex items-center gap-1.5 text-xs font-medium whitespace-nowrap" style={{ color: meta.color }}>
                          <Icon size={14} /> {meta.label}
                        </span>
                      </td>
                      <td className="pr-4 text-slate-500">{e.detail}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )
      )}

      {showForm && <TabletFormModal tablet={editing} hospitals={data.hospitals} staff={data.staff} onSave={save} onClose={() => { setShowForm(false); setEditing(null); }} />}
      {takeoverSourceTablet && (
        <TakeoverSourceModal
          hospitals={data.hospitals} sources={data.sources} existingTablets={data.tablets}
          onSave={takeoverFromSource} onClose={() => setTakeoverSourceTablet(null)}
        />
      )}
      {handoverTablet && <HandoverModal tablet={handoverTablet} hospitals={data.hospitals} staff={data.staff} onSave={handover} onClose={() => setHandoverTablet(null)} />}
      {takeoverHospitalTablet && <TakeoverHospitalModal tablet={takeoverHospitalTablet} onSave={takeoverFromHospital} onClose={() => setTakeoverHospitalTablet(null)} />}
      {showLetterModal && (
        <HandoverLetterModal
          tablets={data.tablets} hospitals={data.hospitals} staff={data.staff}
          onConfirm={handoverViaLetter} onClose={() => setShowLetterModal(false)}
        />
      )}
      {showReplacementModal && (
        <ReplacementLetterModal
          tablets={data.tablets} hospitals={data.hospitals} staff={data.staff}
          onConfirm={replaceViaLetter} onClose={() => setShowReplacementModal(false)}
        />
      )}
      {confirmDeleteId && (
        <ConfirmDialog
          message="Delete this device record? Its custody history will also be removed. This cannot be undone."
          onCancel={() => setConfirmDeleteId(null)}
          onConfirm={() => { remove(confirmDeleteId); setConfirmDeleteId(null); }}
        />
      )}
    </div>
  );
}

function TabletFormModal({ tablet, hospitals, staff, onSave, onClose }) {
  const [form, setForm] = useState(tablet || {
    id: uid("tab"), serialNumber: "", model: "", currentHospitalId: hospitals[0]?.id || "",
    assignedStaffId: null, nadraInstalled: false, batteryHealth: 100, status: "In Stock", remarks: "", history: [],
  });
  const eligibleStaff = staff.filter((s) => s.hospitalId === form.currentHospitalId);
  const isAssigned = form.status === "Assigned";

  return (
    <Modal title={tablet ? "Edit device" : "Add device manually"} onClose={onClose}>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <Field label="Serial number" required className="col-span-2"><TextInput value={form.serialNumber} onChange={(e) => setForm((f) => ({ ...f, serialNumber: e.target.value }))} /></Field>
        <Field label="Model"><TextInput value={form.model} onChange={(e) => setForm((f) => ({ ...f, model: e.target.value }))} placeholder="e.g. Samsung Galaxy Tab A9" /></Field>
        <Field label="Status">
          <Select value={form.status} onChange={(e) => setForm((f) => ({ ...f, status: e.target.value, assignedStaffId: e.target.value === "Assigned" ? f.assignedStaffId : null }))}>
            {TABLET_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </Select>
        </Field>
        <Field label="Current location">
          <Select value={form.currentHospitalId} onChange={(e) => setForm((f) => ({ ...f, currentHospitalId: e.target.value, assignedStaffId: null }))}>
            {hospitals.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
          </Select>
        </Field>
        {isAssigned && (
          <Field label="Assigned staff" required className="col-span-2">
            {eligibleStaff.length === 0 ? (
              <p className="text-xs text-amber-600">No staff registered at this location. Add staff first, or change the location above.</p>
            ) : (
              <Select value={form.assignedStaffId || ""} onChange={(e) => setForm((f) => ({ ...f, assignedStaffId: e.target.value }))}>
                <option value="" disabled>Select staff member</option>
                {eligibleStaff.map((s) => <option key={s.id} value={s.id}>{s.name} - {s.designation}</option>)}
              </Select>
            )}
          </Field>
        )}
        <Field label="Battery health (%)"><TextInput type="number" min={0} max={100} value={form.batteryHealth} onChange={(e) => setForm((f) => ({ ...f, batteryHealth: Number(e.target.value) }))} /></Field>
        <label className="col-span-2 flex items-center gap-2 text-sm text-slate-700">
          <input type="checkbox" checked={form.nadraInstalled} onChange={(e) => setForm((f) => ({ ...f, nadraInstalled: e.target.checked }))} />
          NADRA app installed
        </label>
        <Field label="Remarks" className="col-span-2"><TextArea rows={2} value={form.remarks} onChange={(e) => setForm((f) => ({ ...f, remarks: e.target.value }))} /></Field>
      </div>
      <p className="mt-3 text-xs text-slate-400">This form edits the device's current details directly without logging a custody event. For a proper chain-of-custody record, use "Handover" / "Takeover" instead.</p>
      <div className="mt-3 flex justify-end gap-2">
        <Btn onClick={onClose}>Cancel</Btn>
        <Btn variant="primary" disabled={!form.serialNumber || (isAssigned && !form.assignedStaffId)} onClick={() => onSave(form)}>Save device</Btn>
      </div>
    </Modal>
  );
}

const TAKEOVER_CHECKLIST = ["Physical condition OK", "Screen intact", "Charger included", "Device powers on", "Functional test passed"];

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
  const toggleCheck = (i) => setChecklist((prev) => prev.map((v, idx) => (idx === i ? !v : v)));
  const finalSource = source === "__other__" ? customSource : source;

  return (
    <Modal title="Takeover device from source" onClose={onClose} width="max-w-xl">
      <p className="mb-3 text-xs text-slate-500">Record a device received from an office (e.g. DHO office), inspect it, and bring it into warehouse stock.</p>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <Field label="Device serial number" required className="col-span-2">
          <TextInput value={serialNumber} onChange={(e) => setSerialNumber(e.target.value)} placeholder="Scan or type serial number" />
          {matchedTablet && <p className="mt-1 text-xs text-blue-600">Matches existing device "{matchedTablet.model || matchedTablet.serialNumber}" - this will update it instead of creating a duplicate.</p>}
        </Field>
        {!matchedTablet && (
          <Field label="Model" className="col-span-2"><TextInput value={model} onChange={(e) => setModel(e.target.value)} placeholder="e.g. Samsung Galaxy Tab A9" /></Field>
        )}
        <Field label="Received from (source)" required>
          <Select value={source} onChange={(e) => setSource(e.target.value)}>
            {sources.map((s) => <option key={s} value={s}>{s}</option>)}
            <option value="__other__">Other...</option>
          </Select>
        </Field>
        {source === "__other__" && (
          <Field label="Specify source"><TextInput value={customSource} onChange={(e) => setCustomSource(e.target.value)} placeholder="e.g. DHO Office G9" /></Field>
        )}
        <Field label="Received by">
          <TextInput value={receivedBy} onChange={(e) => setReceivedBy(e.target.value)} placeholder="Name of person receiving" />
        </Field>
        <Field label="Date"><TextInput type="date" value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        <Field label="Bring into warehouse" required className="col-span-2">
          <Select value={warehouseId} onChange={(e) => setWarehouseId(e.target.value)}>
            {hospitals.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
          </Select>
        </Field>

        <div className="col-span-2 rounded-md border border-slate-200 p-3">
          <p className="mb-2 text-xs font-semibold uppercase text-slate-400">Inspection checklist</p>
          <div className="grid grid-cols-2 gap-1.5">
            {TAKEOVER_CHECKLIST.map((label, i) => (
              <label key={label} className="flex items-center gap-1.5 text-sm text-slate-700">
                <input type="checkbox" checked={checklist[i]} onChange={() => toggleCheck(i)} /> {label}
              </label>
            ))}
          </div>
        </div>

        <Field label="Overall condition" required>
          <Select value={overallCondition} onChange={(e) => setOverallCondition(e.target.value)}>
            <option>Good</option><option>Minor issues</option><option>Damaged</option>
          </Select>
        </Field>
        <Field label="Battery health (%)"><TextInput type="number" min={0} max={100} value={batteryHealth} onChange={(e) => setBatteryHealth(Number(e.target.value))} /></Field>
        <Field label="Remarks" className="col-span-2"><TextArea rows={2} value={remarks} onChange={(e) => setRemarks(e.target.value)} placeholder="Any notes about the handover from source" /></Field>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <Btn onClick={onClose}>Cancel</Btn>
        <Btn
          variant="primary"
          disabled={!serialNumber.trim() || !finalSource || !warehouseId}
          onClick={() => onSave({
            tabletId: matchedTablet?.id, serialNumber: serialNumber.trim(), model,
            source: finalSource, receivedBy, warehouseId, date, checklist, checklistLabels: TAKEOVER_CHECKLIST,
            overallCondition, batteryHealth, remarks,
          })}
        >
          Confirm takeover
        </Btn>
      </div>
    </Modal>
  );
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

  useEffect(() => { setStaffId(eligibleStaff[0]?.id || ""); }, [hospitalId]);

  return (
    <Modal title={`Handover device - ${tablet.serialNumber}`} onClose={onClose}>
      <div className="flex flex-col gap-3">
        <Field label="Hospital" required>
          <Select value={hospitalId} onChange={(e) => setHospitalId(e.target.value)}>
            {hospitals.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
          </Select>
        </Field>
        <Field label="Staff member" required>
          {eligibleStaff.length === 0 ? (
            <p className="text-xs text-amber-600">No staff registered at this hospital yet. Add staff first.</p>
          ) : (
            <Select value={staffId} onChange={(e) => setStaffId(e.target.value)}>
              {eligibleStaff.map((s) => <option key={s.id} value={s.id}>{s.name} - {s.designation}</option>)}
            </Select>
          )}
        </Field>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <Field label="Handed over by"><TextInput value={handedOverBy} onChange={(e) => setHandedOverBy(e.target.value)} placeholder="Warehouse keeper name" /></Field>
          <Field label="Date"><TextInput type="date" value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        </div>
        <Field label="Condition at handover">
          <Select value={condition} onChange={(e) => setCondition(e.target.value)}>
            <option>Good</option><option>Minor wear</option><option>Damaged</option>
          </Select>
        </Field>
        <Field label="Accessories given">
          <div className="flex flex-wrap gap-4 text-sm text-slate-700">
            {["Charger", "Cable", "Case", "SIM card"].map((a) => (
              <label key={a} className="flex items-center gap-1.5">
                <input type="checkbox" checked={accessories.includes(a)} onChange={() => toggle(a)} /> {a}
              </label>
            ))}
          </div>
        </Field>
        <Field label="Remarks"><TextArea rows={2} value={remarks} onChange={(e) => setRemarks(e.target.value)} /></Field>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <Btn onClick={onClose}>Cancel</Btn>
        <Btn variant="primary" disabled={!hospitalId || !staffId} onClick={() => onSave(tablet, { hospitalId, staffId, handedOverBy, date, condition, accessories, remarks })}>Confirm handover</Btn>
      </div>
    </Modal>
  );
}

function TakeoverHospitalModal({ tablet, onSave, onClose }) {
  const [receivedBy, setReceivedBy] = useState("");
  const [date, setDate] = useState(todayISO());
  const [condition, setCondition] = useState("Good");
  const [accessories, setAccessories] = useState([]);
  const [inspectionNotes, setInspectionNotes] = useState("");
  const [remarks, setRemarks] = useState("");
  const toggle = (a) => setAccessories((prev) => prev.includes(a) ? prev.filter((x) => x !== a) : [...prev, a]);

  return (
    <Modal title={`Takeover device from hospital - ${tablet.serialNumber}`} onClose={onClose}>
      <div className="flex flex-col gap-3">
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <Field label="Received by"><TextInput value={receivedBy} onChange={(e) => setReceivedBy(e.target.value)} placeholder="Warehouse keeper name" /></Field>
          <Field label="Date"><TextInput type="date" value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        </div>
        <Field label="Condition on return" required>
          <Select value={condition} onChange={(e) => setCondition(e.target.value)}>
            <option>Good</option><option>Minor wear</option><option>Damaged</option>
          </Select>
        </Field>
        <Field label="Accessories returned">
          <div className="flex flex-wrap gap-4 text-sm text-slate-700">
            {["Charger", "Cable", "Case", "SIM card"].map((a) => (
              <label key={a} className="flex items-center gap-1.5">
                <input type="checkbox" checked={accessories.includes(a)} onChange={() => toggle(a)} /> {a}
              </label>
            ))}
          </div>
        </Field>
        <Field label="Inspection notes"><TextArea rows={2} value={inspectionNotes} onChange={(e) => setInspectionNotes(e.target.value)} placeholder="Physical condition, screen, functionality..." /></Field>
        <Field label="Remarks"><TextArea rows={2} value={remarks} onChange={(e) => setRemarks(e.target.value)} /></Field>
        <p className="text-xs text-slate-400">Photos of returned condition can be attached once file upload is enabled for this deployment.</p>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <Btn onClick={onClose}>Cancel</Btn>
        <Btn variant="primary" onClick={() => onSave(tablet, { receivedBy, date, condition, accessories, inspectionNotes, remarks })}>Confirm takeover</Btn>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/* Handover Letter - dynamic form + printable government-format document  */
/* ---------------------------------------------------------------------- */

const NUM_WORDS = ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen", "Twenty"];
const numWord = (n) => NUM_WORDS[n] || String(n);
const pad2 = (n) => String(n).padStart(2, "0");
const ordinal = (n) => { const s = ["th", "st", "nd", "rd"], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); };
const formatLetterDate = (iso) => {
  if (!iso) return "-";
  const d = new Date(iso);
  return `${ordinal(d.getDate())} ${d.toLocaleString(undefined, { month: "long" })}, ${d.getFullYear()}`;
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
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function downloadLetterHTML(html, filename) {
  downloadFile(html, filename, "text/html");
}

// Wraps the same letter HTML with Microsoft Word namespaces so it opens directly and editable in MS Word.
function downloadLetterWord(html, filename) {
  const wordHtml = html
    .replace("<html>", '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">')
    .replace("<head>", '<head>\n<meta name="ProgId" content="Word.Document">\n<meta name="Generator" content="Microsoft Word">');
  downloadFile(wordHtml, filename, "application/msword");
}

function HandoverLetterModal({ tablets, hospitals, staff, onConfirm, onClose }) {
  const [step, setStep] = useState("form"); // 'form' | 'preview'
  const [refNo, setRefNo] = useState("");
  const [letterDate, setLetterDate] = useState(todayISO());
  const [hospitalId, setHospitalId] = useState(hospitals[0]?.id || "");
  const [handingOverName, setHandingOverName] = useState("");
  const [handingOverDesignation, setHandingOverDesignation] = useState("");
  const [takenOverStaffId, setTakenOverStaffId] = useState("");
  const [manualName, setManualName] = useState("");
  const [manualDesignation, setManualDesignation] = useState("");
  const [manualContact, setManualContact] = useState("");
  const [focalPerson, setFocalPerson] = useState("");
  const [rows, setRows] = useState([]); // [{tabletId, counter, date, status}]
  const [deviceSearch, setDeviceSearch] = useState("");

  const availableTablets = tablets.filter((t) => t.status === "In Stock");
  const filteredAvailableTablets = availableTablets.filter((t) =>
    !deviceSearch || [t.serialNumber, t.model].join(" ").toLowerCase().includes(deviceSearch.toLowerCase())
  );
  const eligibleStaff = staff.filter((s) => s.hospitalId === hospitalId);
  const hospital = hospitals.find((h) => h.id === hospitalId);
  const isManual = takenOverStaffId === "__manual__";
  const selectedStaff = eligibleStaff.find((s) => s.id === takenOverStaffId);
  const takenOverName = isManual ? manualName : (selectedStaff?.name || "");
  const takenOverDesignation = isManual ? manualDesignation : (selectedStaff?.designation || "");
  const takenOverContact = isManual ? manualContact : "";

  const toggleDevice = (t) => setRows((prev) =>
    prev.some((r) => r.tabletId === t.id)
      ? prev.filter((r) => r.tabletId !== t.id)
      : [...prev, { tabletId: t.id, counter: "", date: letterDate, status: "Working" }]
  );
  const updateRow = (id, patch) => setRows((prev) => prev.map((r) => (r.tabletId === id ? { ...r, ...patch } : r)));
  const selectAllFiltered = () => setRows((prev) => {
    const existingIds = new Set(prev.map((r) => r.tabletId));
    const toAdd = filteredAvailableTablets.filter((t) => !existingIds.has(t.id))
      .map((t) => ({ tabletId: t.id, counter: "", date: letterDate, status: "Working" }));
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
      refNo, letterDate, hospitalId, handingOverName, handingOverDesignation,
      takenOverStaffId: isManual ? null : takenOverStaffId, takenOverName, takenOverDesignation, takenOverContact,
      focalPerson, rows,
    });
    onClose();
  }

  return (
    <Modal title={step === "form" ? "Generate handover letter" : "Handover letter preview"} onClose={onClose} width="max-w-3xl">
      <style>{`
        @page { size: A4; margin: 15mm; }
        @media print {
          body * { visibility: hidden; }
          .print-area, .print-area * { visibility: visible; }
          .print-area { position: absolute; top: 0; left: 0; width: 100%; padding: 0; margin: 0; border: none !important; box-shadow: none !important; }
          .no-print { display: none !important; }
        }
      `}</style>

      {step === "form" && (
        <div className="flex flex-col gap-4">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <Field label="Reference / F. No"><TextInput value={refNo} onChange={(e) => setRefNo(e.target.value)} placeholder="1-2025/26/Hep-C/Proc/" /></Field>
            <Field label="Letter date"><TextInput type="date" value={letterDate} onChange={(e) => setLetterDate(e.target.value)} /></Field>
            <Field label="Hospital" required className="col-span-2">
              <Select value={hospitalId} onChange={(e) => { setHospitalId(e.target.value); setTakenOverStaffId(""); }}>
                {hospitals.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
              </Select>
            </Field>
          </div>

          <div className="rounded-md border border-slate-200 p-3">
            <p className="mb-2 text-xs font-semibold uppercase text-slate-400">Handed over by</p>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <Field label="Name" required><TextInput value={handingOverName} onChange={(e) => setHandingOverName(e.target.value)} placeholder="e.g. Wajahat Ahmed" /></Field>
              <Field label="Designation"><TextInput value={handingOverDesignation} onChange={(e) => setHandingOverDesignation(e.target.value)} placeholder="e.g. IT / Data Officer" /></Field>
            </div>
          </div>

          <div className="rounded-md border border-slate-200 p-3">
            <p className="mb-2 text-xs font-semibold uppercase text-slate-400">Taken over by</p>
            <Field label="Person" required>
              <Select value={takenOverStaffId} onChange={(e) => setTakenOverStaffId(e.target.value)}>
                <option value="">Select hospital staff...</option>
                {eligibleStaff.map((s) => <option key={s.id} value={s.id}>{s.name} - {s.designation}</option>)}
                <option value="__manual__">Enter manually...</option>
              </Select>
            </Field>
            {isManual && (
              <div className="mt-3 grid grid-cols-2 gap-3">
                <Field label="Name" required><TextInput value={manualName} onChange={(e) => setManualName(e.target.value)} /></Field>
                <Field label="Designation"><TextInput value={manualDesignation} onChange={(e) => setManualDesignation(e.target.value)} /></Field>
                <Field label="Contact No." className="col-span-2"><TextInput value={manualContact} onChange={(e) => setManualContact(e.target.value)} /></Field>
              </div>
            )}
          </div>

          <Field label="Focal person supervising usage (optional)">
            <TextInput value={focalPerson} onChange={(e) => setFocalPerson(e.target.value)} placeholder="e.g. Mr. Ihtisham Afzal Khan" />
          </Field>

          <div>
            <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
              <p className="text-xs font-semibold uppercase text-slate-400">
                Devices ({rows.length} selected) - only devices currently "In Stock" are shown
              </p>
              {availableTablets.length > 0 && (
                <div className="flex gap-2">
                  <button type="button" onClick={selectAllFiltered} className="text-xs font-medium text-slate-600 hover:text-slate-900 hover:underline">Select all{deviceSearch ? " filtered" : ""}</button>
                  <span className="text-xs text-slate-300">|</span>
                  <button type="button" onClick={clearAllFiltered} className="text-xs font-medium text-slate-600 hover:text-slate-900 hover:underline">Clear{deviceSearch ? " filtered" : ""}</button>
                </div>
              )}
            </div>
            {availableTablets.length === 0 ? (
              <p className="rounded-md bg-amber-50 px-3 py-2 text-xs text-amber-700">No devices with status "In Stock". Use "Takeover device" first to bring devices into a warehouse.</p>
            ) : (
              <>
                <div className="relative mb-2">
                  <Search size={14} className="absolute left-2.5 top-2.5 text-slate-400" />
                  <TextInput placeholder="Search by serial number..." value={deviceSearch} onChange={(e) => setDeviceSearch(e.target.value)} className="pl-8" />
                </div>
                {filteredAvailableTablets.length === 0 ? (
                  <p className="rounded-md bg-slate-50 px-3 py-2 text-xs text-slate-400">No devices match "{deviceSearch}".</p>
                ) : (
                  <div className="flex flex-col gap-2 max-h-72 overflow-y-auto pr-1">
                    {filteredAvailableTablets.map((t) => {
                      const row = rows.find((r) => r.tabletId === t.id);
                      const checked = !!row;
                      return (
                        <div key={t.id} className={cx("rounded-md border p-2.5", checked ? "border-slate-900 bg-slate-50" : "border-slate-200")}>
                          <label className="flex items-center gap-2 text-sm">
                            <input type="checkbox" checked={checked} onChange={() => toggleDevice(t)} />
                            <span className="font-medium text-slate-800">{t.serialNumber}</span>
                            <span className="text-xs text-slate-400">IVAS Tablet</span>
                          </label>
                          {checked && (
                            <div className="mt-2 grid grid-cols-3 gap-2 pl-6">
                              <TextInput placeholder="Counter (e.g. OPD Counter - Ground Floor)" value={row.counter} onChange={(e) => updateRow(t.id, { counter: e.target.value })} className="col-span-3 sm:col-span-1" />
                              <TextInput type="date" value={row.date} onChange={(e) => updateRow(t.id, { date: e.target.value })} />
                              <TextInput placeholder="Status (e.g. Working)" value={row.status} onChange={(e) => updateRow(t.id, { status: e.target.value })} />
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </>
            )}
          </div>

          <div className="flex justify-end gap-2">
            <Btn onClick={onClose}>Cancel</Btn>
            <Btn variant="primary" disabled={!canPreview} onClick={() => setStep("preview")}>
              <FileText size={15} /> Preview letter
            </Btn>
          </div>
        </div>
      )}

      {step === "preview" && (
        <div>
          <div className="print-area rounded-md border border-slate-300 bg-white p-6 text-[13px] leading-relaxed text-slate-900" style={{ fontFamily: "Georgia, 'Times New Roman', serif" }}>
            <div className="text-center">
              <p className="font-bold">GOVERNMENT OF PAKISTAN</p>
              <p className="font-semibold">Ministry of National Health Services, Regulations &amp; Coordination,</p>
              <p>3rd Floor Kohsar Block, New Pak Secretariat, Islamabad</p>
              <p className="tracking-widest">**********</p>
            </div>
            <div className="mt-4 flex justify-between text-sm">
              <p>F. No: {refNo || "_______________"}</p>
              <p>Islamabad, the {formatLetterDate(letterDate)}</p>
            </div>
            <p className="mt-4"><span className="font-semibold underline">Subject:</span> <span className="font-semibold underline">{subjectText}</span></p>
            <p className="mt-3">
              I, <b>{handingOverName}</b>{handingOverDesignation && <>, Designation <b>{handingOverDesignation}</b></>} (Hep-C Elimination Program) is hereby handing over {numWord(rows.length)} ({pad2(rows.length)}) Android Tablet Device{rows.length > 1 ? "s" : ""} to <b>{takenOverName}</b> with following details: -
            </p>
            <table className="mt-3 w-full border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50">
                  <th className="border border-slate-800 p-1.5">S.#</th>
                  <th className="border border-slate-800 p-1.5">Item Description</th>
                  <th className="border border-slate-800 p-1.5">Serial No./IMEI No.</th>
                  <th className="border border-slate-800 p-1.5">Counter</th>
                  <th className="border border-slate-800 p-1.5">Handing over Date</th>
                  <th className="border border-slate-800 p-1.5">Status</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const t = tablets.find((x) => x.id === r.tabletId);
                  return (
                    <tr key={r.tabletId}>
                      <td className="border border-slate-800 p-1.5 text-center">{i + 1}.</td>
                      <td className="border border-slate-800 p-1.5">IVAS Tablet</td>
                      <td className="border border-slate-800 p-1.5">{t?.serialNumber}</td>
                      <td className="border border-slate-800 p-1.5">{r.counter}</td>
                      <td className="border border-slate-800 p-1.5 text-center">{fmtDate(r.date)}</td>
                      <td className="border border-slate-800 p-1.5">{r.status}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <p className="mt-3">
              The tablet device shall be utilized exclusively for data entry of citizens/individuals related to Hepatitis C screening, testing, and treatment{focalPerson ? <> under the supervision of focal person <b>{focalPerson}</b></> : ""} from {hospital?.name}.
            </p>
            <div className="mt-10 grid grid-cols-2 gap-6 text-sm">
              <div>
                <p className="font-semibold">HANDED OVER BY</p>
                <div className="mt-11" />
                <p className="mt-1">Name: {handingOverName}</p>
                <p>Designation: {handingOverDesignation || "-"}</p>
                <p className="mt-4">On behalf of.</p>
                <p className="font-semibold">PM's Hepatitis C Elimination Program</p>
                <p className="font-semibold">Ministry of NHSR&amp;C</p>
              </div>
              <div>
                <p className="font-semibold">TAKEN OVER BY</p>
                <div className="mt-11" />
                <p className="mt-1">Name: {takenOverName}</p>
                <p>Designation: {takenOverDesignation || "-"}</p>
                <p>Contact No.: {takenOverContact || "-"}</p>
                <p className="mt-4">On behalf of.</p>
                <p className="font-semibold">{hospital?.name}</p>
                {hospital?.location && <p>{hospital.location}</p>}
              </div>
            </div>
          </div>

          <div className="no-print mt-4 flex justify-between gap-2">
            <Btn onClick={() => setStep("form")}><ArrowLeft size={15} /> Back to form</Btn>
            <div className="flex flex-wrap justify-end gap-2">
              <Btn onClick={() => window.print()}><Printer size={15} /> Print</Btn>
              <Btn
                onClick={() => downloadLetterWord(
                  buildLetterHTML({ refNo, letterDate, hospitalName: hospital?.name || "", hospitalLocation: hospital?.location || "", handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, focalPerson, rows, tablets }),
                  `Handover-Letter-${(refNo || todayISO()).replace(/[^\w-]+/g, "_")}.doc`
                )}
              >
                <FileText size={15} /> Export Word
              </Btn>
              <Btn
                onClick={() => downloadLetterHTML(
                  buildLetterHTML({ refNo, letterDate, hospitalName: hospital?.name || "", hospitalLocation: hospital?.location || "", handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, focalPerson, rows, tablets }),
                  `Handover-Letter-${(refNo || todayISO()).replace(/[^\w-]+/g, "_")}.html`
                )}
              >
                <Download size={15} /> Export PDF
              </Btn>
              <Btn variant="primary" onClick={handleConfirm}><Send size={15} /> Confirm handover</Btn>
            </div>
          </div>
          <p className="no-print mt-2 text-xs text-slate-400">
            "Print" opens your browser's print dialog directly. If it doesn't respond (some browsers block printing from embedded previews), use "Export PDF" - it downloads a print-ready file; open it and choose "Save as PDF" from its print dialog. "Export Word" downloads an editable .doc file that opens directly in Microsoft Word.
          </p>
        </div>
      )}
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/* Replacement Letter - old device swapped for a new device                */
/* ---------------------------------------------------------------------- */

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

function ReplacementLetterModal({ tablets, hospitals, staff, onConfirm, onClose }) {
  const [step, setStep] = useState("form"); // 'form' | 'preview'
  const [refNo, setRefNo] = useState("");
  const [letterDate, setLetterDate] = useState(todayISO());
  const [hospitalId, setHospitalId] = useState(hospitals[0]?.id || "");
  const [handingOverName, setHandingOverName] = useState("");
  const [handingOverDesignation, setHandingOverDesignation] = useState("");
  const [takenOverStaffId, setTakenOverStaffId] = useState("");
  const [manualName, setManualName] = useState("");
  const [manualDesignation, setManualDesignation] = useState("");
  const [manualContact, setManualContact] = useState("");
  const [rows, setRows] = useState([]); // [{oldTabletId, newTabletId, reason, counter, date, status, oldStatus}]
  const [deviceSearch, setDeviceSearch] = useState("");

  const assignedAtHospital = tablets.filter((t) => t.status === "Assigned" && t.currentHospitalId === hospitalId);
  const filteredAssignedAtHospital = assignedAtHospital.filter((t) =>
    !deviceSearch || [t.serialNumber, t.model].join(" ").toLowerCase().includes(deviceSearch.toLowerCase())
  );
  const availableTablets = tablets.filter((t) => t.status === "In Stock");
  const eligibleStaff = staff.filter((s) => s.hospitalId === hospitalId);
  const hospital = hospitals.find((h) => h.id === hospitalId);
  const isManual = takenOverStaffId === "__manual__";
  const selectedStaff = eligibleStaff.find((s) => s.id === takenOverStaffId);
  const takenOverName = isManual ? manualName : (selectedStaff?.name || "");
  const takenOverDesignation = isManual ? manualDesignation : (selectedStaff?.designation || "");
  const takenOverContact = isManual ? manualContact : "";
  const usedNewIds = rows.map((r) => r.newTabletId).filter(Boolean);

  const toggleOld = (t) => setRows((prev) =>
    prev.some((r) => r.oldTabletId === t.id)
      ? prev.filter((r) => r.oldTabletId !== t.id)
      : [...prev, { oldTabletId: t.id, newTabletId: "", reason: "", counter: "", date: letterDate, status: "Working", oldStatus: "Damaged" }]
  );
  const updateRow = (oldId, patch) => setRows((prev) => prev.map((r) => (r.oldTabletId === oldId ? { ...r, ...patch } : r)));

  const canPreview = hospitalId && handingOverName.trim() && takenOverName.trim() && rows.length > 0 && rows.every((r) => r.newTabletId);
  const subjectText = `Replacement of ${numWord(rows.length)} (${pad2(rows.length)}) Android Tablet Device${rows.length > 1 ? "s" : ""} under "Prime Minister Programme for the Elimination of Hepatitis C Infection".`;

  function handleConfirm() {
    onConfirm({
      refNo, letterDate, hospitalId, handingOverName, handingOverDesignation,
      takenOverStaffId: isManual ? null : takenOverStaffId, takenOverName, takenOverDesignation, takenOverContact,
      rows,
    });
    onClose();
  }

  return (
    <Modal title={step === "form" ? "Generate replacement letter" : "Replacement letter preview"} onClose={onClose} width="max-w-3xl">
      <style>{`
        @page { size: A4; margin: 15mm; }
        @media print {
          body * { visibility: hidden; }
          .print-area, .print-area * { visibility: visible; }
          .print-area { position: absolute; top: 0; left: 0; width: 100%; padding: 0; margin: 0; border: none !important; box-shadow: none !important; }
          .no-print { display: none !important; }
        }
      `}</style>

      {step === "form" && (
        <div className="flex flex-col gap-4">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <Field label="Reference / F. No"><TextInput value={refNo} onChange={(e) => setRefNo(e.target.value)} /></Field>
            <Field label="Letter date"><TextInput type="date" value={letterDate} onChange={(e) => setLetterDate(e.target.value)} /></Field>
            <Field label="Hospital" required className="col-span-2">
              <Select value={hospitalId} onChange={(e) => { setHospitalId(e.target.value); setTakenOverStaffId(""); setRows([]); setDeviceSearch(""); }}>
                {hospitals.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
              </Select>
            </Field>
          </div>

          <div className="rounded-md border border-slate-200 p-3">
            <p className="mb-2 text-xs font-semibold uppercase text-slate-400">Handed over by</p>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <Field label="Name" required><TextInput value={handingOverName} onChange={(e) => setHandingOverName(e.target.value)} placeholder="e.g. Wajahat Ahmed" /></Field>
              <Field label="Designation"><TextInput value={handingOverDesignation} onChange={(e) => setHandingOverDesignation(e.target.value)} placeholder="e.g. IT / Data Officer" /></Field>
            </div>
          </div>

          <div className="rounded-md border border-slate-200 p-3">
            <p className="mb-2 text-xs font-semibold uppercase text-slate-400">Taken over by (receiving replacement)</p>
            <Field label="Person" required>
              <Select value={takenOverStaffId} onChange={(e) => setTakenOverStaffId(e.target.value)}>
                <option value="">Select hospital staff...</option>
                {eligibleStaff.map((s) => <option key={s.id} value={s.id}>{s.name} - {s.designation}</option>)}
                <option value="__manual__">Enter manually...</option>
              </Select>
            </Field>
            {isManual && (
              <div className="mt-3 grid grid-cols-2 gap-3">
                <Field label="Name" required><TextInput value={manualName} onChange={(e) => setManualName(e.target.value)} /></Field>
                <Field label="Designation"><TextInput value={manualDesignation} onChange={(e) => setManualDesignation(e.target.value)} /></Field>
                <Field label="Contact No." className="col-span-2"><TextInput value={manualContact} onChange={(e) => setManualContact(e.target.value)} /></Field>
              </div>
            )}
          </div>

          <div>
            <p className="mb-2 text-xs font-semibold uppercase text-slate-400">
              Devices to replace ({rows.length} selected) - devices currently Assigned at this hospital
            </p>
            {!hospitalId ? (
              <p className="text-xs text-slate-400">Select a hospital first.</p>
            ) : assignedAtHospital.length === 0 ? (
              <p className="rounded-md bg-amber-50 px-3 py-2 text-xs text-amber-700">No devices are currently Assigned at this hospital.</p>
            ) : (
              <>
                <div className="relative mb-2">
                  <Search size={14} className="absolute left-2.5 top-2.5 text-slate-400" />
                  <TextInput placeholder="Search by serial number..." value={deviceSearch} onChange={(e) => setDeviceSearch(e.target.value)} className="pl-8" />
                </div>
                {filteredAssignedAtHospital.length === 0 ? (
                  <p className="rounded-md bg-slate-50 px-3 py-2 text-xs text-slate-400">No devices match "{deviceSearch}".</p>
                ) : (
              <div className="flex flex-col gap-2 max-h-80 overflow-y-auto pr-1">
                {filteredAssignedAtHospital.map((t) => {
                  const row = rows.find((r) => r.oldTabletId === t.id);
                  const checked = !!row;
                  const newOptions = availableTablets.filter((nt) => nt.id === row?.newTabletId || !usedNewIds.includes(nt.id));
                  return (
                    <div key={t.id} className={cx("rounded-md border p-2.5", checked ? "border-slate-900 bg-slate-50" : "border-slate-200")}>
                      <label className="flex items-center gap-2 text-sm">
                        <input type="checkbox" checked={checked} onChange={() => toggleOld(t)} />
                        <span className="font-medium text-slate-800">{t.serialNumber}</span>
                        <span className="text-xs text-slate-400">IVAS Tablet (old device)</span>
                      </label>
                      {checked && (
                        <div className="mt-2 flex flex-col gap-2 pl-6">
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            <Field label="Replacement device" required>
                              {newOptions.length === 0 ? (
                                <p className="text-xs text-amber-600">No available "In Stock" devices left. Takeover a device first.</p>
                              ) : (
                                <Select value={row.newTabletId} onChange={(e) => updateRow(t.id, { newTabletId: e.target.value })}>
                                  <option value="">Select new device...</option>
                                  {newOptions.map((nt) => <option key={nt.id} value={nt.id}>{nt.serialNumber}</option>)}
                                </Select>
                              )}
                            </Field>
                            <Field label="Old device becomes">
                              <Select value={row.oldStatus} onChange={(e) => updateRow(t.id, { oldStatus: e.target.value })}>
                                {TABLET_STATUSES.filter((s) => s !== "Assigned").map((s) => <option key={s} value={s}>{s}</option>)}
                              </Select>
                            </Field>
                          </div>
                          <TextInput placeholder="Reason for replacement (e.g. screen damage, malfunction)" value={row.reason} onChange={(e) => updateRow(t.id, { reason: e.target.value })} />
                          <div className="grid grid-cols-3 gap-2">
                            <TextInput placeholder="Counter" value={row.counter} onChange={(e) => updateRow(t.id, { counter: e.target.value })} className="col-span-1" />
                            <TextInput type="date" value={row.date} onChange={(e) => updateRow(t.id, { date: e.target.value })} />
                            <TextInput placeholder="Status" value={row.status} onChange={(e) => updateRow(t.id, { status: e.target.value })} />
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
                )}
              </>
            )}
          </div>

          <div className="flex justify-end gap-2">
            <Btn onClick={onClose}>Cancel</Btn>
            <Btn variant="primary" disabled={!canPreview} onClick={() => setStep("preview")}>
              <FileText size={15} /> Preview letter
            </Btn>
          </div>
        </div>
      )}

      {step === "preview" && (
        <div>
          <div className="print-area rounded-md border border-slate-300 bg-white p-6 text-[13px] leading-relaxed text-slate-900" style={{ fontFamily: "Georgia, 'Times New Roman', serif" }}>
            <div className="text-center">
              <p className="font-bold">GOVERNMENT OF PAKISTAN</p>
              <p className="font-semibold">Ministry of National Health Services, Regulations &amp; Coordination,</p>
              <p>3rd Floor Kohsar Block, New Pak Secretariat, Islamabad</p>
              <p className="tracking-widest">**********</p>
            </div>
            <div className="mt-4 flex justify-between text-sm">
              <p>F. No: {refNo || "_______________"}</p>
              <p>Islamabad, the {formatLetterDate(letterDate)}</p>
            </div>
            <p className="mt-4"><span className="font-semibold underline">Subject:</span> <span className="font-semibold underline">{subjectText}</span></p>
            <p className="mt-3">
              I, <b>{handingOverName}</b>{handingOverDesignation && <>, Designation <b>{handingOverDesignation}</b></>} (Hep-C Elimination Program) am hereby replacing the following faulty/damaged device(s) at <b>{hospital?.name}</b> and handing over the replacement device(s) to <b>{takenOverName}</b> with following details: -
            </p>
            <table className="mt-3 w-full border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50">
                  <th className="border border-slate-800 p-1.5">S.#</th>
                  <th className="border border-slate-800 p-1.5">Old Device IMEI (Returned)</th>
                  <th className="border border-slate-800 p-1.5">New Device IMEI (Issued)</th>
                  <th className="border border-slate-800 p-1.5">Reason for Replacement</th>
                  <th className="border border-slate-800 p-1.5">Counter</th>
                  <th className="border border-slate-800 p-1.5">Replacement Date</th>
                  <th className="border border-slate-800 p-1.5">Status</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const oldT = tablets.find((x) => x.id === r.oldTabletId);
                  const newT = tablets.find((x) => x.id === r.newTabletId);
                  return (
                    <tr key={r.oldTabletId}>
                      <td className="border border-slate-800 p-1.5 text-center">{i + 1}.</td>
                      <td className="border border-slate-800 p-1.5">{oldT?.serialNumber}</td>
                      <td className="border border-slate-800 p-1.5">{newT?.serialNumber}</td>
                      <td className="border border-slate-800 p-1.5">{r.reason}</td>
                      <td className="border border-slate-800 p-1.5">{r.counter}</td>
                      <td className="border border-slate-800 p-1.5 text-center">{fmtDate(r.date)}</td>
                      <td className="border border-slate-800 p-1.5">{r.status}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <p className="mt-3">
              The replacement device(s) shall be utilized exclusively for data entry of citizens/individuals related to Hepatitis C screening, testing, and treatment at {hospital?.name}. The returned/faulty device(s) shall be taken back into program custody for inspection and repair.
            </p>
            <div className="mt-10 grid grid-cols-2 gap-6 text-sm">
              <div>
                <p className="font-semibold">HANDED OVER BY</p>
                <div className="mt-11" />
                <p className="mt-1">Name: {handingOverName}</p>
                <p>Designation: {handingOverDesignation || "-"}</p>
                <p className="mt-4">On behalf of.</p>
                <p className="font-semibold">PM's Hepatitis C Elimination Program</p>
                <p className="font-semibold">Ministry of NHSR&amp;C</p>
              </div>
              <div>
                <p className="font-semibold">TAKEN OVER BY</p>
                <div className="mt-11" />
                <p className="mt-1">Name: {takenOverName}</p>
                <p>Designation: {takenOverDesignation || "-"}</p>
                <p>Contact No.: {takenOverContact || "-"}</p>
                <p className="mt-4">On behalf of.</p>
                <p className="font-semibold">{hospital?.name}</p>
                {hospital?.location && <p>{hospital.location}</p>}
              </div>
            </div>
          </div>

          <div className="no-print mt-4 flex justify-between gap-2">
            <Btn onClick={() => setStep("form")}><ArrowLeft size={15} /> Back to form</Btn>
            <div className="flex flex-wrap justify-end gap-2">
              <Btn onClick={() => window.print()}><Printer size={15} /> Print</Btn>
              <Btn
                onClick={() => downloadLetterWord(
                  buildReplacementLetterHTML({ refNo, letterDate, hospitalName: hospital?.name || "", hospitalLocation: hospital?.location || "", handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, rows, tablets }),
                  `Replacement-Letter-${(refNo || todayISO()).replace(/[^\w-]+/g, "_")}.doc`
                )}
              >
                <FileText size={15} /> Export Word
              </Btn>
              <Btn
                onClick={() => downloadLetterHTML(
                  buildReplacementLetterHTML({ refNo, letterDate, hospitalName: hospital?.name || "", hospitalLocation: hospital?.location || "", handingOverName, handingOverDesignation, takenOverName, takenOverDesignation, takenOverContact, rows, tablets }),
                  `Replacement-Letter-${(refNo || todayISO()).replace(/[^\w-]+/g, "_")}.html`
                )}
              >
                <Download size={15} /> Export PDF
              </Btn>
              <Btn variant="primary" onClick={handleConfirm}><Send size={15} /> Confirm replacement</Btn>
            </div>
          </div>
          <p className="no-print mt-2 text-xs text-slate-400">
            "Print" opens your browser's print dialog directly. If it doesn't respond, use "Export PDF" - it downloads a print-ready file; open it and choose "Save as PDF". "Export Word" downloads an editable .doc file that opens directly in Microsoft Word.
          </p>
        </div>
      )}
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/* Staff Tab                                                               */
/* ---------------------------------------------------------------------- */

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
      return { ...prev, staff: exists ? prev.staff.map((x) => (x.id === s.id ? s : x)) : [...prev.staff, s] };
    });
    setShowForm(false); setEditing(null);
  }

  function remove(id) {
    update((prev) => ({ ...prev, staff: prev.staff.filter((x) => x.id !== id) }));
  }

  return (
    <div>
      <Toolbar title="Hospital staff" subtitle="Directory of staff eligible for tablet and asset assignment." search={search} setSearch={setSearch}>
        <Btn variant="primary" onClick={() => { setEditing(null); setShowForm(true); }}><Plus size={15} /> Add staff</Btn>
      </Toolbar>

      {data.staff.length === 0 ? (
        <EmptyState icon={Users} title="No staff registered yet" />
      ) : (
        <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
          <table className="w-full min-w-[720px] text-sm">
            <thead className="bg-slate-50 text-left text-xs uppercase text-slate-400">
              <tr><th className="py-2.5 pl-4">Employee ID</th><th>Name</th><th>Designation</th><th>Department</th><th>Hospital</th><th className="text-right pr-4">Actions</th></tr>
            </thead>
            <tbody>
              {filtered.map((s) => (
                <tr key={s.id} className="border-t border-slate-100 hover:bg-slate-50">
                  <td className="py-2.5 pl-4 text-slate-500">{s.employeeId}</td>
                  <td className="font-medium text-slate-700">{s.name}<p className="text-xs font-normal text-slate-400">{s.cnic}</p></td>
                  <td className="text-slate-600">{s.designation}</td>
                  <td className="text-slate-600">{s.department}</td>
                  <td className="text-slate-600">{hospitalName(s.hospitalId)}</td>
                  <td className="pr-4">
                    <div className="flex justify-end gap-1">
                      <button title="Edit" className="rounded p-1.5 text-slate-400 hover:bg-slate-200 hover:text-slate-700" onClick={() => { setEditing(s); setShowForm(true); }}><Pencil size={14} /></button>
                      <button title="Delete" className="rounded p-1.5 text-slate-400 hover:bg-red-100 hover:text-red-600" onClick={() => setConfirmDeleteId(s.id)}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {showForm && <StaffFormModal staff={editing} hospitals={data.hospitals} designations={data.designations} departments={data.departments} onSave={save} onClose={() => { setShowForm(false); setEditing(null); }} />}
      {confirmDeleteId && (
        <ConfirmDialog
          message="Delete this staff member? This cannot be undone."
          onCancel={() => setConfirmDeleteId(null)}
          onConfirm={() => { remove(confirmDeleteId); setConfirmDeleteId(null); }}
        />
      )}
    </div>
  );
}

function StaffFormModal({ staff, hospitals, designations, departments, onSave, onClose }) {
  const [form, setForm] = useState(staff || {
    id: uid("staff"), employeeId: "", name: "", cnic: "", designation: designations[0] || "",
    department: departments[0] || "", hospitalId: hospitals[0]?.id || "",
  });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  return (
    <Modal title={staff ? "Edit staff member" : "Add staff member"} onClose={onClose}>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <Field label="Employee ID" required><TextInput value={form.employeeId} onChange={set("employeeId")} /></Field>
        <Field label="Full name" required><TextInput value={form.name} onChange={set("name")} /></Field>
        <Field label="CNIC"><TextInput value={form.cnic} onChange={set("cnic")} placeholder="XXXXX-XXXXXXX-X" /></Field>
        <Field label="Hospital" required>
          <Select value={form.hospitalId} onChange={set("hospitalId")}>
            {hospitals.map((h) => <option key={h.id} value={h.id}>{h.name}</option>)}
          </Select>
        </Field>
        <Field label="Designation" required>
          <Select value={form.designation} onChange={set("designation")}>
            {designations.map((d) => <option key={d} value={d}>{d}</option>)}
          </Select>
        </Field>
        <Field label="Department" required>
          <Select value={form.department} onChange={set("department")}>
            {departments.map((d) => <option key={d} value={d}>{d}</option>)}
          </Select>
        </Field>
      </div>
      <div className="mt-5 flex justify-end gap-2">
        <Btn onClick={onClose}>Cancel</Btn>
        <Btn variant="primary" disabled={!form.name || !form.employeeId} onClick={() => onSave(form)}>Save staff member</Btn>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/* Settings / Admin Tab                                                    */
/* ---------------------------------------------------------------------- */

function SettingsTab({ data, update }) {
  const setCategories = (categories) => update((prev) => ({ ...prev, categories }));
  const setUnits = (units) => update((prev) => ({ ...prev, units }));
  const setDesignations = (designations) => update((prev) => ({ ...prev, designations }));
  const setDepartments = (departments) => update((prev) => ({ ...prev, departments }));
  const setSources = (sources) => update((prev) => ({ ...prev, sources }));

  return (
    <div>
      <Toolbar title="Admin settings" subtitle="Configure master data used across the system - no code changes required." />
      <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
        <ConfigList
          label="Inventory categories"
          items={data.categories}
          extraFields={{
            initial: { name: "", group: "" },
            valid: (d) => d.name.trim(),
            display: (c) => `${c.name}${c.group ? ` (${c.group})` : ""}`,
            render: (d, setD) => (
              <div className="flex flex-1 gap-2">
                <TextInput placeholder="Category name" value={d.name} onChange={(e) => setD({ ...d, name: e.target.value })} />
                <TextInput placeholder="Group (optional)" value={d.group} onChange={(e) => setD({ ...d, group: e.target.value })} />
              </div>
            ),
          }}
          onAdd={(d) => {
            const id = uid("cat");
            setCategories([{ id, ...d }, ...data.categories]);
            return id;
          }}
          onEdit={(id, d) => setCategories(data.categories.map((c) => (c.id === id ? { ...c, ...d } : c)))}
          onDelete={(id) => setCategories(data.categories.filter((c) => c.id !== id))}
        />

        <ConfigList
          label="Units of measure"
          items={data.units}
          onAdd={(v) => setUnits([v, ...data.units])}
          onEdit={(id, v) => setUnits(data.units.map((u) => (u === id ? v : u)))}
          onDelete={(id) => setUnits(data.units.filter((u) => u !== id))}
        />

        <ConfigList
          label="Staff designations"
          items={data.designations}
          onAdd={(v) => setDesignations([v, ...data.designations])}
          onEdit={(id, v) => setDesignations(data.designations.map((d) => (d === id ? v : d)))}
          onDelete={(id) => setDesignations(data.designations.filter((d) => d !== id))}
        />

        <ConfigList
          label="Departments"
          items={data.departments}
          onAdd={(v) => setDepartments([v, ...data.departments])}
          onEdit={(id, v) => setDepartments(data.departments.map((d) => (d === id ? v : d)))}
          onDelete={(id) => setDepartments(data.departments.filter((d) => d !== id))}
        />

        <ConfigList
          label="Device handover / takeover sources"
          items={data.sources}
          onAdd={(v) => setSources([v, ...data.sources])}
          onEdit={(id, v) => setSources(data.sources.map((s) => (s === id ? v : s)))}
          onDelete={(id) => setSources(data.sources.filter((s) => s !== id))}
        />
      </div>
      <p className="mt-4 text-xs text-slate-400">
        Warehouses and hospitals are configured from the "Hospitals & Warehouses" tab since they carry extra fields (type, location).
      </p>
    </div>
  );
}
