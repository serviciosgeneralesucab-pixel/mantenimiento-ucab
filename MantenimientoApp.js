"use client";

import React, { useState, useEffect, useRef } from "react";
import {
  Zap, Wrench, Snowflake, Truck, Building2, Plus, Check, X, RefreshCw,
  Users, ClipboardList, CalendarClock, AlertTriangle, LayoutDashboard,
  Trash2, ChevronLeft, ChevronRight, ClipboardCheck, PackageSearch,
  Upload, Download, Mail,
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Cell,
} from "recharts";
import * as XLSX from "xlsx";
import { supabase } from "../lib/supabaseClient";

/* ---------------------------------------------------------------------- */
/* Config & helpers                                                       */
/* ---------------------------------------------------------------------- */

const GROUP_META = {
  electrico: { nombre: "Grupo Eléctrico", color: "#C4880A", Icon: Zap },
  mecanico: { nombre: "Grupo Mecánico", color: "#3E5C9A", Icon: Wrench },
  refrigeracion: { nombre: "Grupo Refrigeración", color: "#1592A3", Icon: Snowflake },
  logistica: { nombre: "Grupo Logística", color: "#C9622B", Icon: Truck },
  infraestructura: { nombre: "Grupo Infraestructura", color: "#3C8F5C", Icon: Building2 },
};
const GROUP_ORDER = ["electrico", "mecanico", "refrigeracion", "logistica", "infraestructura"];

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

function monthKey(d = new Date()) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}
function monthLabel(key) {
  const [y, m] = key.split("-").map(Number);
  return `${MESES[m - 1]} ${y}`;
}
function shiftMonth(key, delta) {
  const [y, m] = key.split("-").map(Number);
  const d = new Date(y, m - 1 + delta, 1);
  return monthKey(d);
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function normalizeHeader(s) {
  return String(s || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .trim();
}
function matchGroupId(text) {
  const n = normalizeHeader(text);
  if (n.includes("ELECTR")) return "electrico";
  if (n.includes("MECAN")) return "mecanico";
  if (n.includes("REFRIG")) return "refrigeracion";
  if (n.includes("LOGIST")) return "logistica";
  if (n.includes("INFRAESTRUCT")) return "infraestructura";
  return null;
}

const FREQ_WORDS = {
  DIARIA: 1, DIARIO: 1, SEMANAL: 7, QUINCENAL: 15, MENSUAL: 30,
  BIMESTRAL: 60, TRIMESTRAL: 90, CUATRIMESTRAL: 120, SEMESTRAL: 180, ANUAL: 365,
};
function parseFrecuenciaDias(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  if (!isNaN(n) && n > 0) return Math.round(n);
  const norm = normalizeHeader(v);
  for (const [word, days] of Object.entries(FREQ_WORDS)) {
    if (norm.includes(word)) return days;
  }
  return null;
}
const PLAN_FREQ_OPTIONS = ["MENSUAL", "BIMESTRAL", "TRIMESTRAL", "CUATRIMESTRAL", "SEMESTRAL", "ANUAL"];
function freqLabel(word) {
  if (!word) return "Según condición";
  const w = String(word);
  return w.charAt(0) + w.slice(1).toLowerCase();
}
function freqWordToMonths(word) {
  const norm = normalizeHeader(word);
  const days = FREQ_WORDS[norm];
  if (!days) return null;
  return Math.max(1, Math.round(days / 30));
}
function monthsDiff(fromISODate, toMonthKey) {
  if (!fromISODate) return null;
  const parts = String(fromISODate).split("-").map(Number);
  const [fy, fm] = parts;
  if (!fy || !fm) return null;
  const [ty, tm] = toMonthKey.split("-").map(Number);
  return (ty * 12 + tm) - (fy * 12 + fm);
}
function isDueInMonth(fechaUltimaEjecucion, freqMonths, mesKey) {
  if (!freqMonths) return true;
  const diff = monthsDiff(fechaUltimaEjecucion, mesKey);
  if (diff === null) return true;
  return diff >= freqMonths;
}
function excelValueToISO(v) {
  if (!v) return "";
  if (v instanceof Date && !isNaN(v)) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;
  }
  if (typeof v === "string") {
    let m = v.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = v.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) return `${m[3]}-${String(m[1]).padStart(2, "0")}-${String(m[2]).padStart(2, "0")}`;
  }
  return "";
}
function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        resolve(XLSX.read(e.target.result, { type: "array", cellDates: true }));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
function sheetToRows(wb, sheetName) {
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "", raw: true });
}
function findHeaderRow(rows, anchorKeywords) {
  for (let r = 0; r < Math.min(rows.length, 30); r++) {
    const norm = (rows[r] || []).map(normalizeHeader);
    if (anchorKeywords.some((k) => norm.some((c) => c.includes(k)))) {
      return { rowIndex: r, headers: norm };
    }
  }
  return null;
}
function colIndex(headers, keywords) {
  for (let i = 0; i < headers.length; i++) {
    if (keywords.some((k) => headers[i].includes(k))) return i;
  }
  return -1;
}
function exportToExcel(filename, sheetName, rows) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  XLSX.writeFile(wb, filename);
}

/* ---------------------------------------------------------------------- */
/* UI atoms                                                                */
/* ---------------------------------------------------------------------- */

function GroupDot({ id, size = 9 }) {
  const meta = GROUP_META[id];
  if (!meta) return null;
  return <span style={{ display: "inline-block", width: size, height: size, borderRadius: "50%", background: meta.color, flexShrink: 0 }} />;
}
function GroupBadge({ id }) {
  const meta = GROUP_META[id];
  if (!meta) return <span style={{ color: "var(--ink-soft)" }}>—</span>;
  const Icon = meta.Icon;
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12.5 }}>
      <Icon size={13} color={meta.color} strokeWidth={2.4} />
      <span style={{ color: "var(--ink)" }}>{meta.nombre.replace("Grupo ", "")}</span>
    </span>
  );
}

const STATUS_STYLES = {
  pendiente: { bg: "#EEF0F3", fg: "#5B6472", label: "PENDIENTE" },
  realizado: { bg: "#E4F5EA", fg: "#1E7A45", label: "REALIZADO" },
  no_realizado: { bg: "#FBE9E6", fg: "#B4432A", label: "NO REALIZADO" },
  realizado_sin_obs: { bg: "#E4F5EA", fg: "#1E7A45", label: "REALIZADO S/OBS" },
  realizado_con_obs: { bg: "#FFF3D9", fg: "#8A6100", label: "REALIZADO C/OBS" },
  reprogramado: { bg: "#E7ECFB", fg: "#33459E", label: "REPROGRAMADO" },
  en_proceso: { bg: "#E7ECFB", fg: "#33459E", label: "EN PROCESO" },
  completada: { bg: "#E4F5EA", fg: "#1E7A45", label: "COMPLETADA" },
  anulada: { bg: "#F1E7F5", fg: "#6B3B8A", label: "ANULADA" },
};
function StatusChip({ status }) {
  const s = STATUS_STYLES[status] || STATUS_STYLES.pendiente;
  return (
    <span style={{ fontFamily: "var(--mono)", fontSize: 10.5, fontWeight: 700, letterSpacing: "0.04em", padding: "3px 7px", borderRadius: 5, background: s.bg, color: s.fg, whiteSpace: "nowrap" }}>
      {s.label}
    </span>
  );
}

function Btn({ children, onClick, kind = "default", size = "md", icon: Icon, disabled, title }) {
  const base = {
    display: "inline-flex", alignItems: "center", gap: 6, borderRadius: 7, fontWeight: 600,
    fontSize: size === "sm" ? 12 : 13, padding: size === "sm" ? "5px 10px" : "7px 13px",
    cursor: disabled ? "not-allowed" : "pointer", border: "1px solid transparent",
    transition: "background .12s, border-color .12s, opacity .12s", opacity: disabled ? 0.5 : 1,
  };
  const kinds = {
    default: { background: "#fff", border: "1px solid var(--line)", color: "var(--ink)" },
    primary: { background: "var(--primary)", color: "#fff" },
    ghost: { background: "transparent", color: "var(--ink-soft)" },
    danger: { background: "transparent", color: "#B4432A" },
    success: { background: "#1E7A45", color: "#fff" },
  };
  return (
    <button disabled={disabled} title={title} onClick={onClick} style={{ ...base, ...kinds[kind] }}>
      {Icon && <Icon size={14} strokeWidth={2.3} />}
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12.5 }}>
      <span style={{ color: "var(--ink-soft)", fontWeight: 600 }}>{label}</span>
      {children}
    </label>
  );
}
const inputStyle = {
  border: "1px solid var(--line)", borderRadius: 6, padding: "6px 9px", fontSize: 13,
  fontFamily: "var(--sans)", color: "var(--ink)", outline: "none", background: "#fff",
};
function Empty({ icon: Icon, text }) {
  return (
    <div style={{ padding: "38px 20px", textAlign: "center", color: "var(--ink-soft)", fontSize: 13 }}>
      <Icon size={22} style={{ opacity: 0.4, marginBottom: 8 }} />
      <div>{text}</div>
    </div>
  );
}
function MonthSwitcher({ mes, setMes }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
      <button onClick={() => setMes(shiftMonth(mes, -1))} style={{ border: "1px solid var(--line)", background: "#fff", borderRadius: 6, padding: 5, cursor: "pointer" }}>
        <ChevronLeft size={14} />
      </button>
      <div style={{ fontFamily: "var(--mono)", fontSize: 12.5, fontWeight: 700, color: "var(--ink)", minWidth: 128, textAlign: "center", letterSpacing: "0.02em" }}>
        {monthLabel(mes).toUpperCase()}
      </div>
      <button onClick={() => setMes(shiftMonth(mes, 1))} style={{ border: "1px solid var(--line)", background: "#fff", borderRadius: 6, padding: 5, cursor: "pointer" }}>
        <ChevronRight size={14} />
      </button>
    </div>
  );
}
function Card({ title, action, children, style }) {
  return (
    <section style={{ background: "var(--surface)", border: "1px solid var(--line)", borderRadius: 10, overflow: "hidden", ...style }}>
      {(title || action) && (
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 14px", borderBottom: "1px solid var(--line)" }}>
          <h2 style={{ margin: 0, fontSize: 13.5, fontWeight: 700, color: "var(--ink)" }}>{title}</h2>
          {action}
        </div>
      )}
      {children}
    </section>
  );
}
function StatBox({ label, value, sub, alert }) {
  return (
    <div style={{ background: "var(--surface)", border: "1px solid var(--line)", borderRadius: 10, padding: 14 }}>
      <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.05em", color: "var(--ink-soft)", marginBottom: 6 }}>{label.toUpperCase()}</div>
      <div style={{ fontSize: 26, fontWeight: 800, fontFamily: "var(--mono)", color: alert ? "#B4432A" : "var(--ink)" }}>{value}</div>
      {sub && <div style={{ fontSize: 11.5, color: alert ? "#B4432A" : "var(--ink-soft)", marginTop: 2 }}>{sub}</div>}
    </div>
  );
}
function ImportExcelButton({ onFile, label = "Importar desde Excel" }) {
  const inputRef = useRef(null);
  const [status, setStatus] = useState(null);
  async function handleChange(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setStatus("Leyendo archivo…");
    try {
      const result = await onFile(file);
      setStatus(
        result.imported === 0 && result.skipped === 0
          ? "No se encontraron filas reconocibles en el archivo."
          : `Importados: ${result.imported} · omitidos (duplicados o incompletos): ${result.skipped}`
      );
    } catch (err) {
      setStatus("No se pudo leer el archivo. Verifica que sea un .xlsx válido.");
    }
    e.target.value = "";
  }
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
      <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleChange} />
      <Btn size="sm" icon={Upload} onClick={() => inputRef.current?.click()}>{label}</Btn>
      {status && <span style={{ fontSize: 11, color: "var(--ink-soft)" }}>{status}</span>}
    </span>
  );
}

/* ---------------------------------------------------------------------- */
/* App principal                                                           */
/* ---------------------------------------------------------------------- */

export default function MantenimientoApp() {
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("dashboard");
  const [mes, setMes] = useState(monthKey());
  const [errorCarga, setErrorCarga] = useState(null);

  const [groups, setGroups] = useState([]);
  const [rutinaDefs, setRutinaDefs] = useState([]);
  const [rutinaInst, setRutinaInst] = useState([]);
  const [planDefs, setPlanDefs] = useState([]);
  const [planInst, setPlanInst] = useState([]);
  const [correctivas, setCorrectivas] = useState([]);

  useEffect(() => {
    (async () => {
      const [g, rd, ri, pd, pi, c] = await Promise.all([
        supabase.from("groups").select("*"),
        supabase.from("rutina_defs").select("*").order("created_at"),
        supabase.from("rutina_inst").select("*").order("created_at"),
        supabase.from("plan_defs").select("*").order("created_at"),
        supabase.from("plan_inst").select("*").order("created_at"),
        supabase.from("correctivas").select("*").order("created_at"),
      ]);
      const anyError = [g, rd, ri, pd, pi, c].find((r) => r.error);
      if (anyError) setErrorCarga(anyError.error.message);
      const gData = (g.data || []).slice().sort((a, b) => GROUP_ORDER.indexOf(a.id) - GROUP_ORDER.indexOf(b.id));
      setGroups(gData);
      setRutinaDefs(rd.data || []);
      setRutinaInst(ri.data || []);
      setPlanDefs(pd.data || []);
      setPlanInst(pi.data || []);
      setCorrectivas(c.data || []);
      setLoading(false);
    })();
  }, []);

  const TABS = [
    { id: "dashboard", label: "Panel general", icon: LayoutDashboard },
    { id: "rutinas", label: "Rutinas de inspección", icon: ClipboardCheck },
    { id: "planes", label: "Planes de mantenimiento", icon: ClipboardList },
    { id: "correctivas", label: "Órdenes correctivas", icon: PackageSearch },
    { id: "grupos", label: "Grupos", icon: Users },
  ];

  if (loading) {
    return (
      <div style={{ ...shellStyle, alignItems: "center", justifyContent: "center", minHeight: 420 }}>
        <style>{GLOBAL_CSS}</style>
        <div style={{ color: "var(--ink-soft)", fontSize: 13 }}>Cargando…</div>
      </div>
    );
  }

  return (
    <div style={shellStyle}>
      <style>{GLOBAL_CSS}</style>

      <header style={{ padding: "18px 22px 14px", borderBottom: "1px solid var(--line)" }}>
        <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.12em", color: "var(--accent-teal)", marginBottom: 3 }}>
              COORDINACIÓN DE MANTENIMIENTO
            </div>
            <h1 style={{ margin: 0, fontSize: 21, fontWeight: 800, color: "var(--ink)", letterSpacing: "-0.01em" }}>
              Seguimiento de cumplimiento
            </h1>
          </div>
          {tab !== "grupos" && <MonthSwitcher mes={mes} setMes={setMes} />}
        </div>

        <nav style={{ display: "flex", gap: 4, marginTop: 16, flexWrap: "wrap" }}>
          {TABS.map((t) => {
            const Icon = t.icon;
            const active = tab === t.id;
            return (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                style={{
                  display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", borderRadius: 7,
                  border: "1px solid", borderColor: active ? "var(--primary)" : "var(--line)",
                  background: active ? "var(--primary)" : "#fff", color: active ? "#fff" : "var(--ink-soft)",
                  fontSize: 12.5, fontWeight: 600, cursor: "pointer",
                }}
              >
                <Icon size={14} strokeWidth={2.3} />
                {t.label}
              </button>
            );
          })}
        </nav>
      </header>

      {errorCarga && (
        <div style={{ padding: "10px 22px", background: "#FBE9E6", color: "#B4432A", fontSize: 12.5 }}>
          No se pudo conectar con la base de datos ({errorCarga}). Verifica la configuración de Supabase.
        </div>
      )}

      <main style={{ padding: 22, flex: 1 }}>
        {tab === "dashboard" && (
          <Dashboard mes={mes} groups={groups} rutinaInst={rutinaInst} planInst={planInst} correctivas={correctivas} />
        )}
        {tab === "rutinas" && (
          <RutinasTab mes={mes} groups={groups} defs={rutinaDefs} setDefs={setRutinaDefs} inst={rutinaInst} setInst={setRutinaInst} />
        )}
        {tab === "planes" && (
          <PlanesTab mes={mes} groups={groups} defs={planDefs} setDefs={setPlanDefs} inst={planInst} setInst={setPlanInst} />
        )}
        {tab === "correctivas" && (
          <CorrectivasTab mes={mes} groups={groups} items={correctivas} setItems={setCorrectivas} />
        )}
        {tab === "grupos" && <GruposTab groups={groups} setGroups={setGroups} />}
      </main>

      <footer style={{ padding: "10px 22px", borderTop: "1px solid var(--line)", fontSize: 11, color: "var(--ink-soft)" }}>
        Datos guardados en la base de datos del proyecto — compartidos entre todos los usuarios de esta app.
      </footer>
    </div>
  );
}

const shellStyle = {
  fontFamily: "var(--sans)", background: "var(--bg)", color: "var(--ink)", minHeight: "100vh",
  display: "flex", flexDirection: "column",
};
const GLOBAL_CSS = `
:root {
  --bg: #F3F5F8; --surface: #FFFFFF; --ink: #1A2333; --ink-soft: #626C7A; --line: #E3E7EC;
  --primary: #1B2A4A; --accent-teal: #0F766E;
  --sans: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
  --mono: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;
}
table { border-collapse: collapse; width: 100%; }
th { text-align: left; font-size: 10.5px; letter-spacing: 0.06em; color: var(--ink-soft); font-weight: 700; padding: 8px 10px; border-bottom: 1px solid var(--line); white-space: nowrap; }
td { padding: 9px 10px; font-size: 12.8px; border-bottom: 1px solid var(--line); vertical-align: middle; }
tr:last-child td { border-bottom: none; }
input, select, textarea { font-family: var(--sans); }
::placeholder { color: #A6ADB8; }
`;

/* ---------------------------------------------------------------------- */
/* Dashboard                                                                */
/* ---------------------------------------------------------------------- */

function Dashboard({ mes, groups, rutinaInst, planInst, correctivas }) {
  const rInMonth = rutinaInst.filter((i) => i.mes === mes);
  const pInMonth = planInst.filter((i) => i.mes === mes);

  const compliance = groups.map((g) => {
    const rg = rInMonth.filter((i) => i.grupo === g.id);
    const pg = pInMonth.filter((i) => i.grupo === g.id);
    const rDone = rg.filter((i) => i.estado === "realizado").length;
    const pDone = pg.filter((i) => i.estado.startsWith("realizado")).length;
    return {
      nombre: g.nombre.replace("Grupo ", ""),
      color: GROUP_META[g.id]?.color || "#999",
      rutinas: rg.length ? Math.round((rDone / rg.length) * 100) : null,
      planes: pg.length ? Math.round((pDone / pg.length) * 100) : null,
    };
  });

  const rTotal = rInMonth.length;
  const rDoneTotal = rInMonth.filter((i) => i.estado === "realizado").length;
  const pTotal = pInMonth.length;
  const pDoneTotal = pInMonth.filter((i) => i.estado.startsWith("realizado")).length;

  const correctivasAbiertas = correctivas.filter((c) => c.estado !== "completada" && c.estado !== "anulada");
  const correctivasVencidas = correctivasAbiertas.filter((c) => c.fecha_compromiso && c.fecha_compromiso < todayISO());

  const chartData = compliance.map((c) => ({ ...c, rutinas: c.rutinas ?? 0, planes: c.planes ?? 0 }));

  function enviarProgramaPorCorreo() {
    const groupName = (id) => groups.find((g) => g.id === id)?.nombre || id;
    let cuerpo = `Programa de mantenimiento — ${monthLabel(mes)}\n\n`;
    cuerpo += `RUTINAS DE INSPECCIÓN (${rInMonth.length})\n`;
    if (rInMonth.length === 0) cuerpo += "(sin carga generada para este mes)\n";
    rInMonth.forEach((i) => { cuerpo += `- [${groupName(i.grupo)}] ${i.nombre} — programado ${i.fecha}\n`; });
    cuerpo += `\nPLANES DE MANTENIMIENTO (${pInMonth.length})\n`;
    if (pInMonth.length === 0) cuerpo += "(sin carga generada para este mes)\n";
    pInMonth.forEach((i) => { cuerpo += `- [${groupName(i.grupo)}] ${i.nombre} — programado ${i.fecha || mes + "-01"}\n`; });
    const asunto = `Programa de mantenimiento — ${monthLabel(mes)}`;
    const url = `mailto:servicios.generales.ucab@gmail.com?subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}`;
    window.open(url, "_blank");
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <Btn size="sm" icon={Mail} kind="primary" onClick={enviarProgramaPorCorreo}>Enviar programa por correo</Btn>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12 }}>
        <StatBox label="Rutinas del mes" value={rTotal} sub={rTotal ? `${rDoneTotal} realizadas · ${Math.round((rDoneTotal / rTotal) * 100)}%` : "sin carga generada"} />
        <StatBox label="Planes del mes" value={pTotal} sub={pTotal ? `${pDoneTotal} realizados · ${Math.round((pDoneTotal / pTotal) * 100)}%` : "sin carga generada"} />
        <StatBox label="Órdenes correctivas abiertas" value={correctivasAbiertas.length} sub={`${correctivas.length} en total`} />
        <StatBox label="Correctivas vencidas" value={correctivasVencidas.length} sub="fecha de compromiso superada" alert={correctivasVencidas.length > 0} />
      </div>

      <Card title={`Cumplimiento por grupo — ${monthLabel(mes)}`}>
        <div style={{ padding: 14, height: 260 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={chartData} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--line)" vertical={false} />
              <XAxis dataKey="nombre" tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={{ stroke: "var(--line)" }} tickLine={false} />
              <YAxis domain={[0, 100]} tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={false} tickLine={false} unit="%" />
              <Tooltip formatter={(v, name) => [`${v}%`, name === "rutinas" ? "Rutinas" : "Planes"]} contentStyle={{ fontSize: 12, borderRadius: 8, border: "1px solid var(--line)" }} />
              <Bar dataKey="rutinas" name="Rutinas" radius={[4, 4, 0, 0]} maxBarSize={22}>
                {chartData.map((d, i) => <Cell key={i} fill={d.color} />)}
              </Bar>
              <Bar dataKey="planes" name="Planes" radius={[4, 4, 0, 0]} maxBarSize={22} fillOpacity={0.45}>
                {chartData.map((d, i) => <Cell key={i} fill={d.color} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>

      <Card title="Estatus de órdenes correctivas">
        <div style={{ padding: 14, display: "flex", gap: 22, flexWrap: "wrap" }}>
          {["pendiente", "en_proceso", "completada", "anulada"].map((s) => {
            const n = correctivas.filter((c) => c.estado === s).length;
            return (
              <div key={s} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <StatusChip status={s} />
                <div style={{ fontSize: 20, fontWeight: 800, fontFamily: "var(--mono)" }}>{n}</div>
              </div>
            );
          })}
        </div>
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Rutinas de inspección                                                   */
/* ---------------------------------------------------------------------- */

function RutinasTab({ mes, groups, defs, setDefs, inst, setInst }) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ nombre: "", grupo: groups[0]?.id, frecuencia_dias: 7, fecha_ultima_ejecucion: "" });
  const [filterGrupo, setFilterGrupo] = useState("todos");

  const monthInst = inst.filter((i) => i.mes === mes && (filterGrupo === "todos" || i.grupo === filterGrupo));
  const generatedDefIds = new Set(inst.filter((i) => i.mes === mes).map((i) => i.def_id));
  const dueDefs = defs.filter(
    (d) => !generatedDefIds.has(d.id) && isDueInMonth(d.fecha_ultima_ejecucion, Math.max(1, Math.round((d.frecuencia_dias || 30) / 30)), mes)
  );

  async function addDef() {
    if (!form.nombre.trim()) return;
    const payload = { nombre: form.nombre.trim(), grupo: form.grupo, frecuencia_dias: Number(form.frecuencia_dias) || 1, fecha_ultima_ejecucion: form.fecha_ultima_ejecucion || null };
    const { data, error } = await supabase.from("rutina_defs").insert(payload).select().single();
    if (!error && data) setDefs([...defs, data]);
    setForm({ nombre: "", grupo: groups[0]?.id, frecuencia_dias: 7, fecha_ultima_ejecucion: "" });
    setShowForm(false);
  }
  async function removeDef(id) {
    await supabase.from("rutina_defs").delete().eq("id", id);
    setDefs(defs.filter((d) => d.id !== id));
    setInst(inst.filter((i) => i.def_id !== id));
  }
  async function generarCarga() {
    const nuevas = dueDefs.map((d) => ({
      def_id: d.id, nombre: d.nombre, grupo: d.grupo, mes, fecha: `${mes}-01`,
      estado: "pendiente", observaciones: "", fecha_ejecucion: null,
    }));
    if (nuevas.length === 0) return;
    const { data, error } = await supabase.from("rutina_inst").insert(nuevas).select();
    if (!error && data) setInst([...inst, ...data]);
  }
  async function updateInst(id, patch) {
    const { data, error } = await supabase.from("rutina_inst").update(patch).eq("id", id).select().single();
    if (error || !data) return;
    setInst((prev) => prev.map((i) => (i.id === id ? data : i)));
    if (patch.estado === "realizado" && patch.fecha_ejecucion) {
      const { data: defData } = await supabase.from("rutina_defs").update({ fecha_ultima_ejecucion: patch.fecha_ejecucion }).eq("id", data.def_id).select().single();
      if (defData) setDefs((prev) => prev.map((d) => (d.id === data.def_id ? defData : d)));
    }
  }
  async function importarDesdeExcel(file) {
    const wb = await readWorkbook(file);
    let imported = 0, skipped = 0;
    for (const sheetName of wb.SheetNames) {
      const rows = sheetToRows(wb, sheetName);
      const header = findHeaderRow(rows, ["ACTIVIDAD"]);
      if (!header) continue;
      const idxNombre = colIndex(header.headers, ["ACTIVIDAD"]);
      const idxGrupo = colIndex(header.headers, ["AREA", "GRUPO"]);
      const idxFrecuencia = colIndex(header.headers, ["FRECUENCIA"]);
      const idxFecha = colIndex(header.headers, ["ULTIMA INTERVENCION", "ULTIMA EJECUCION"]);
      if (idxNombre === -1) continue;
      const nuevos = [];
      for (let r = header.rowIndex + 1; r < rows.length; r++) {
        const row = rows[r];
        const nombre = row[idxNombre];
        if (!nombre || String(nombre).trim() === "") continue;
        const grupoTxt = idxGrupo !== -1 ? row[idxGrupo] : "";
        const grupo = matchGroupId(grupoTxt) || matchGroupId(sheetName) || groups[0]?.id;
        const frecuencia_dias = idxFrecuencia !== -1 ? parseFrecuenciaDias(row[idxFrecuencia]) : null;
        const fecha_ultima_ejecucion = idxFecha !== -1 ? excelValueToISO(row[idxFecha]) : "";
        const nombreLimpio = String(nombre).trim();
        const yaExiste = [...defs, ...nuevos].some((d) => d.nombre.trim().toLowerCase() === nombreLimpio.toLowerCase() && d.grupo === grupo);
        if (yaExiste) { skipped++; continue; }
        nuevos.push({ nombre: nombreLimpio, grupo, frecuencia_dias: frecuencia_dias || 30, fecha_ultima_ejecucion: fecha_ultima_ejecucion || null });
        imported++;
      }
      if (nuevos.length) {
        const { data } = await supabase.from("rutina_defs").insert(nuevos).select();
        if (data) setDefs((prev) => [...prev, ...data]);
      }
      break;
    }
    return { imported, skipped };
  }
  function exportarCarga() {
    exportToExcel(`rutinas_${mes}.xlsx`, `Rutinas ${mes}`, monthInst.map((i) => ({
      Rutina: i.nombre,
      Grupo: groups.find((g) => g.id === i.grupo)?.nombre || i.grupo,
      "Fecha programada": i.fecha,
      "Fecha última ejecución": i.fecha_ejecucion || "",
      Estado: (STATUS_STYLES[i.estado] || {}).label || i.estado,
      Observaciones: i.observaciones || "",
    })));
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <Card title="Rutinas definidas" action={
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <ImportExcelButton onFile={importarDesdeExcel} />
          <Btn icon={Plus} size="sm" onClick={() => setShowForm((s) => !s)}>Nueva rutina</Btn>
        </div>
      }>
        {showForm && (
          <div style={{ padding: 14, borderBottom: "1px solid var(--line)", display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
            <Field label="Nombre de la rutina">
              <input style={{ ...inputStyle, width: 220 }} value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })} placeholder="Ej. Inspección tablero eléctrico" />
            </Field>
            <Field label="Grupo">
              <select style={{ ...inputStyle, width: 170 }} value={form.grupo} onChange={(e) => setForm({ ...form, grupo: e.target.value })}>
                {groups.map((g) => <option key={g.id} value={g.id}>{g.nombre}</option>)}
              </select>
            </Field>
            <Field label="Frecuencia (días)">
              <input type="number" min={1} style={{ ...inputStyle, width: 100 }} value={form.frecuencia_dias} onChange={(e) => setForm({ ...form, frecuencia_dias: e.target.value })} />
            </Field>
            <Field label="Fecha última ejecución">
              <input type="date" style={inputStyle} value={form.fecha_ultima_ejecucion} onChange={(e) => setForm({ ...form, fecha_ultima_ejecucion: e.target.value })} />
            </Field>
            <Btn kind="primary" onClick={addDef}>Guardar</Btn>
          </div>
        )}
        {defs.length === 0 ? (
          <Empty icon={ClipboardCheck} text="Aún no hay rutinas definidas. Crea la primera con “Nueva rutina”." />
        ) : (
          <table>
            <thead><tr><th>Rutina</th><th>Grupo</th><th>Frecuencia</th><th>Fecha últ. ejecución</th><th></th></tr></thead>
            <tbody>
              {defs.map((d) => (
                <tr key={d.id}>
                  <td>{d.nombre}</td>
                  <td><GroupBadge id={d.grupo} /></td>
                  <td style={{ fontFamily: "var(--mono)" }}>cada {d.frecuencia_dias} días</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{d.fecha_ultima_ejecucion || "—"}</td>
                  <td style={{ textAlign: "right" }}><Btn kind="danger" size="sm" icon={Trash2} onClick={() => removeDef(d.id)}>Eliminar</Btn></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>

      <Card title={`Carga de ${monthLabel(mes)}`} action={
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <select style={{ ...inputStyle, padding: "5px 8px", fontSize: 12 }} value={filterGrupo} onChange={(e) => setFilterGrupo(e.target.value)}>
            <option value="todos">Todos los grupos</option>
            {groups.map((g) => <option key={g.id} value={g.id}>{g.nombre}</option>)}
          </select>
          <Btn size="sm" icon={RefreshCw} kind="primary" onClick={generarCarga} disabled={dueDefs.length === 0}>
            Generar carga del mes {dueDefs.length ? `(${dueDefs.length})` : ""}
          </Btn>
          <Btn size="sm" icon={Download} onClick={exportarCarga} disabled={monthInst.length === 0}>Exportar</Btn>
        </div>
      }>
        {monthInst.length === 0 ? (
          <Empty icon={CalendarClock} text="Sin rutinas cargadas para este mes todavía. Usa “Generar carga del mes”." />
        ) : (
          <table>
            <thead><tr><th>Rutina</th><th>Grupo</th><th>Fecha prog.</th><th>Fecha últ. ejecución</th><th>Estado</th><th>Observaciones</th><th></th></tr></thead>
            <tbody>
              {monthInst.map((i) => <RutinaRow key={i.id} item={i} onUpdate={updateInst} />)}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

function RutinaRow({ item, onUpdate }) {
  const [obs, setObs] = useState(item.observaciones || "");
  const [fechaEj, setFechaEj] = useState(item.fecha_ejecucion || "");

  function guardarCampos() {
    onUpdate(item.id, { observaciones: obs, fecha_ejecucion: fechaEj || null });
  }
  function marcar(estado) {
    const fe = fechaEj || (estado === "realizado" ? todayISO() : fechaEj);
    setFechaEj(fe);
    onUpdate(item.id, { estado, observaciones: obs, fecha_ejecucion: fe || null });
  }

  return (
    <tr>
      <td>{item.nombre}</td>
      <td><GroupBadge id={item.grupo} /></td>
      <td style={{ fontFamily: "var(--mono)" }}>{item.fecha}</td>
      <td>
        <input type="date" style={{ ...inputStyle, fontSize: 12, padding: "4px 6px" }} value={fechaEj} onChange={(e) => setFechaEj(e.target.value)} onBlur={guardarCampos} />
      </td>
      <td><StatusChip status={item.estado} /></td>
      <td style={{ minWidth: 200 }}>
        <input style={{ ...inputStyle, width: "100%", fontSize: 12 }} placeholder="Observaciones…" value={obs} onChange={(e) => setObs(e.target.value)} onBlur={guardarCampos} />
      </td>
      <td style={{ textAlign: "right" }}>
        {item.estado === "pendiente" ? (
          <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
            <Btn size="sm" kind="success" icon={Check} onClick={() => marcar("realizado")}>Sí</Btn>
            <Btn size="sm" kind="danger" icon={X} onClick={() => marcar("no_realizado")}>No</Btn>
          </div>
        ) : (
          <Btn size="sm" kind="ghost" onClick={() => onUpdate(item.id, { estado: "pendiente" })}>Reabrir</Btn>
        )}
      </td>
    </tr>
  );
}

/* ---------------------------------------------------------------------- */
/* Planes de mantenimiento                                                 */
/* ---------------------------------------------------------------------- */

function PlanesTab({ mes, groups, defs, setDefs, inst, setInst }) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ nombre: "", grupo: groups[0]?.id, frecuencia: "MENSUAL", fecha_ultima_ejecucion: "" });
  const [filterGrupo, setFilterGrupo] = useState("todos");

  const monthInst = inst.filter((i) => i.mes === mes && (filterGrupo === "todos" || i.grupo === filterGrupo));
  const generatedDefIds = new Set(inst.filter((i) => i.mes === mes).map((i) => i.def_id));
  const dueDefs = defs.filter(
    (d) => !generatedDefIds.has(d.id) && isDueInMonth(d.fecha_ultima_ejecucion, freqWordToMonths(d.frecuencia), mes)
  );

  async function addDef() {
    if (!form.nombre.trim()) return;
    const payload = { nombre: form.nombre.trim(), grupo: form.grupo, frecuencia: form.frecuencia, fecha_ultima_ejecucion: form.fecha_ultima_ejecucion || null };
    const { data, error } = await supabase.from("plan_defs").insert(payload).select().single();
    if (!error && data) setDefs([...defs, data]);
    setForm({ nombre: "", grupo: groups[0]?.id, frecuencia: "MENSUAL", fecha_ultima_ejecucion: "" });
    setShowForm(false);
  }
  async function removeDef(id) {
    await supabase.from("plan_defs").delete().eq("id", id);
    setDefs(defs.filter((d) => d.id !== id));
    setInst(inst.filter((i) => i.def_id !== id));
  }
  async function generarCarga() {
    const nuevas = dueDefs.map((d) => ({
      def_id: d.id, nombre: d.nombre, grupo: d.grupo, mes, fecha: `${mes}-01`,
      estado: "pendiente", observaciones: "", fecha_reprogramada: null, fecha_ultima_ejecucion: null,
    }));
    if (nuevas.length === 0) return;
    const { data, error } = await supabase.from("plan_inst").insert(nuevas).select();
    if (!error && data) setInst([...inst, ...data]);
  }
  async function updateInst(id, patch) {
    const { data, error } = await supabase.from("plan_inst").update(patch).eq("id", id).select().single();
    if (error || !data) return;
    setInst((prev) => prev.map((i) => (i.id === id ? data : i)));
    if (patch.estado && patch.estado.startsWith("realizado") && patch.fecha_ultima_ejecucion) {
      const { data: defData } = await supabase.from("plan_defs").update({ fecha_ultima_ejecucion: patch.fecha_ultima_ejecucion }).eq("id", data.def_id).select().single();
      if (defData) setDefs((prev) => prev.map((d) => (d.id === data.def_id ? defData : d)));
    }
  }
  async function importarDesdeExcel(file) {
    const wb = await readWorkbook(file);
    let imported = 0, skipped = 0;
    for (const sheetName of wb.SheetNames) {
      const rows = sheetToRows(wb, sheetName);
      const header = findHeaderRow(rows, ["ACTIVIDAD"]);
      if (!header) continue;
      const idxNombre = colIndex(header.headers, ["ACTIVIDAD"]);
      const idxGrupo = colIndex(header.headers, ["AREA", "GRUPO"]);
      const idxFecha = colIndex(header.headers, ["ULTIMA INTERVENCION", "ULTIMA EJECUCION"]);
      const idxFrecuencia = colIndex(header.headers, ["FRECUENCIA"]);
      if (idxNombre === -1) continue;
      const nuevos = [];
      for (let r = header.rowIndex + 1; r < rows.length; r++) {
        const row = rows[r];
        const nombre = row[idxNombre];
        if (!nombre || String(nombre).trim() === "") continue;
        const grupoTxt = idxGrupo !== -1 ? row[idxGrupo] : "";
        const grupo = matchGroupId(grupoTxt) || matchGroupId(sheetName) || groups[0]?.id;
        const fecha_ultima_ejecucion = idxFecha !== -1 ? excelValueToISO(row[idxFecha]) : "";
        const freqNorm = idxFrecuencia !== -1 ? normalizeHeader(row[idxFrecuencia]) : "";
        const frecuencia = PLAN_FREQ_OPTIONS.includes(freqNorm) ? freqNorm : null;
        const nombreLimpio = String(nombre).trim();
        const yaExiste = [...defs, ...nuevos].some((d) => d.nombre.trim().toLowerCase() === nombreLimpio.toLowerCase() && d.grupo === grupo);
        if (yaExiste) { skipped++; continue; }
        nuevos.push({ nombre: nombreLimpio, grupo, frecuencia, fecha_ultima_ejecucion: fecha_ultima_ejecucion || null });
        imported++;
      }
      if (nuevos.length) {
        const { data } = await supabase.from("plan_defs").insert(nuevos).select();
        if (data) setDefs((prev) => [...prev, ...data]);
      }
      break;
    }
    return { imported, skipped };
  }
  function exportarCarga() {
    exportToExcel(`planes_${mes}.xlsx`, `Planes ${mes}`, monthInst.map((i) => ({
      Plan: i.nombre,
      Grupo: groups.find((g) => g.id === i.grupo)?.nombre || i.grupo,
      "Fecha programada": i.fecha || `${mes}-01`,
      Estado: (STATUS_STYLES[i.estado] || {}).label || i.estado,
      "Fecha última ejecución": i.fecha_ultima_ejecucion || "",
      Observaciones: i.observaciones || "",
    })));
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <Card title="Planes definidos" action={
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <ImportExcelButton onFile={importarDesdeExcel} />
          <Btn icon={Plus} size="sm" onClick={() => setShowForm((s) => !s)}>Nuevo plan</Btn>
        </div>
      }>
        {showForm && (
          <div style={{ padding: 14, borderBottom: "1px solid var(--line)", display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
            <Field label="Nombre del plan">
              <input style={{ ...inputStyle, width: 240 }} value={form.nombre} onChange={(e) => setForm({ ...form, nombre: e.target.value })} placeholder="Ej. Mantenimiento preventivo compresor 3" />
            </Field>
            <Field label="Grupo">
              <select style={{ ...inputStyle, width: 170 }} value={form.grupo} onChange={(e) => setForm({ ...form, grupo: e.target.value })}>
                {groups.map((g) => <option key={g.id} value={g.id}>{g.nombre}</option>)}
              </select>
            </Field>
            <Field label="Frecuencia">
              <select style={{ ...inputStyle, width: 150 }} value={form.frecuencia} onChange={(e) => setForm({ ...form, frecuencia: e.target.value })}>
                {PLAN_FREQ_OPTIONS.map((f) => <option key={f} value={f}>{freqLabel(f)}</option>)}
              </select>
            </Field>
            <Field label="Fecha última ejecución">
              <input type="date" style={inputStyle} value={form.fecha_ultima_ejecucion} onChange={(e) => setForm({ ...form, fecha_ultima_ejecucion: e.target.value })} />
            </Field>
            <Btn kind="primary" onClick={addDef}>Guardar</Btn>
          </div>
        )}
        {defs.length === 0 ? (
          <Empty icon={ClipboardList} text="Aún no hay planes definidos. Crea el primero con “Nuevo plan”." />
        ) : (
          <table>
            <thead><tr><th>Plan</th><th>Grupo</th><th>Frecuencia</th><th>Fecha últ. ejecución</th><th></th></tr></thead>
            <tbody>
              {defs.map((d) => (
                <tr key={d.id}>
                  <td>{d.nombre}</td>
                  <td><GroupBadge id={d.grupo} /></td>
                  <td style={{ fontFamily: "var(--mono)" }}>{freqLabel(d.frecuencia)}</td>
                  <td style={{ fontFamily: "var(--mono)" }}>{d.fecha_ultima_ejecucion || "—"}</td>
                  <td style={{ textAlign: "right" }}><Btn kind="danger" size="sm" icon={Trash2} onClick={() => removeDef(d.id)}>Eliminar</Btn></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>

      <Card title={`Carga de ${monthLabel(mes)}`} action={
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <select style={{ ...inputStyle, padding: "5px 8px", fontSize: 12 }} value={filterGrupo} onChange={(e) => setFilterGrupo(e.target.value)}>
            <option value="todos">Todos los grupos</option>
            {groups.map((g) => <option key={g.id} value={g.id}>{g.nombre}</option>)}
          </select>
          <Btn size="sm" icon={RefreshCw} kind="primary" onClick={generarCarga} disabled={dueDefs.length === 0}>
            Generar carga del mes {dueDefs.length ? `(${dueDefs.length})` : ""}
          </Btn>
          <Btn size="sm" icon={Download} onClick={exportarCarga} disabled={monthInst.length === 0}>Exportar</Btn>
        </div>
      }>
        {monthInst.length === 0 ? (
          <Empty icon={CalendarClock} text="Sin planes cargados para este mes todavía. Usa “Generar carga del mes”." />
        ) : (
          <table>
            <thead><tr><th>Plan</th><th>Grupo</th><th>Fecha prog.</th><th>Estado</th><th>Fecha últ. ejecución</th><th>Observaciones</th><th></th></tr></thead>
            <tbody>
              {monthInst.map((i) => <PlanRow key={i.id} item={i} onUpdate={updateInst} />)}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

function PlanRow({ item, onUpdate }) {
  const [note, setNote] = useState(item.observaciones || "");
  const [fechaEj, setFechaEj] = useState(item.fecha_ultima_ejecucion || "");
  const [reprog, setReprog] = useState(item.fecha_reprogramada || "");
  const [showReprog, setShowReprog] = useState(false);

  function guardarCampos() {
    onUpdate(item.id, { observaciones: note, fecha_ultima_ejecucion: fechaEj || null });
  }
  function marcar(estado) {
    const fe = fechaEj || todayISO();
    setFechaEj(fe);
    onUpdate(item.id, { estado, observaciones: note, fecha_ultima_ejecucion: fe });
  }

  return (
    <tr>
      <td>{item.nombre}</td>
      <td><GroupBadge id={item.grupo} /></td>
      <td style={{ fontFamily: "var(--mono)" }}>{item.fecha || "—"}</td>
      <td>
        <StatusChip status={item.estado} />
        {item.estado === "reprogramado" && item.fecha_reprogramada && (
          <div style={{ fontFamily: "var(--mono)", fontSize: 10.5, color: "var(--ink-soft)", marginTop: 3 }}>nueva fecha: {item.fecha_reprogramada}</div>
        )}
      </td>
      <td>
        <input type="date" style={{ ...inputStyle, fontSize: 12, padding: "4px 6px" }} value={fechaEj} onChange={(e) => setFechaEj(e.target.value)} onBlur={guardarCampos} />
      </td>
      <td style={{ minWidth: 200 }}>
        <input style={{ ...inputStyle, width: "100%", fontSize: 12 }} placeholder="Observaciones…" value={note} onChange={(e) => setNote(e.target.value)} onBlur={guardarCampos} />
      </td>
      <td style={{ textAlign: "right" }}>
        {item.estado === "pendiente" ? (
          showReprog ? (
            <div style={{ display: "flex", gap: 6, justifyContent: "flex-end", alignItems: "center" }}>
              <input type="date" style={{ ...inputStyle, fontSize: 12 }} value={reprog} onChange={(e) => setReprog(e.target.value)} />
              <Btn size="sm" kind="primary" onClick={() => { if (reprog) onUpdate(item.id, { estado: "reprogramado", fecha_reprogramada: reprog }); }}>Confirmar</Btn>
            </div>
          ) : (
            <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
              <Btn size="sm" kind="success" onClick={() => marcar("realizado_sin_obs")}>Sin obs.</Btn>
              <Btn size="sm" kind="default" onClick={() => marcar("realizado_con_obs")}>Con obs.</Btn>
              <Btn size="sm" kind="danger" onClick={() => setShowReprog(true)}>Reprogramar</Btn>
            </div>
          )
        ) : (
          <Btn size="sm" kind="ghost" onClick={() => onUpdate(item.id, { estado: "pendiente", fecha_reprogramada: null })}>Reabrir</Btn>
        )}
      </td>
    </tr>
  );
}

/* ---------------------------------------------------------------------- */
/* Órdenes correctivas                                                     */
/* ---------------------------------------------------------------------- */

function CorrectivasTab({ mes, groups, items, setItems }) {
  const [showForm, setShowForm] = useState(false);
  const [filterGrupo, setFilterGrupo] = useState("todos");
  const [filterEstado, setFilterEstado] = useState("todos");
  const [form, setForm] = useState({ numero: "", descripcion: "", unidad: "", grupo: groups[0]?.id, fecha_compromiso: "", fecha_ejecucion: "" });

  const filtered = items.filter((i) => (filterGrupo === "todos" || i.grupo === filterGrupo) && (filterEstado === "todos" || i.estado === filterEstado));

  async function addItem() {
    if (!form.numero.trim() || !form.descripcion.trim()) return;
    const payload = { ...form, fecha_compromiso: form.fecha_compromiso || null, fecha_ejecucion: form.fecha_ejecucion || null, estado: "pendiente", tratamiento: "" };
    const { data, error } = await supabase.from("correctivas").insert(payload).select().single();
    if (!error && data) setItems([...items, data]);
    setForm({ numero: "", descripcion: "", unidad: "", grupo: groups[0]?.id, fecha_compromiso: "", fecha_ejecucion: "" });
    setShowForm(false);
  }
  async function updateItem(id, patch) {
    const { data, error } = await supabase.from("correctivas").update(patch).eq("id", id).select().single();
    if (!error && data) setItems((prev) => prev.map((i) => (i.id === id ? data : i)));
  }
  async function removeItem(id) {
    await supabase.from("correctivas").delete().eq("id", id);
    setItems(items.filter((i) => i.id !== id));
  }
  async function importarDesdeExcel(file) {
    const wb = await readWorkbook(file);
    let imported = 0, skipped = 0;
    for (const sheetName of wb.SheetNames) {
      const rows = sheetToRows(wb, sheetName);
      const header = findHeaderRow(rows, ["DESCRIPCION", "N ORDEN", "NRO ORDEN", "NUMERO DE ORDEN"]);
      if (!header) continue;
      const idxNumero = colIndex(header.headers, ["N ORDEN", "NRO ORDEN", "NUMERO DE ORDEN", "NUMERO", "ORDEN"]);
      const idxDescripcion = colIndex(header.headers, ["DESCRIPCION", "ACTIVIDAD", "FALLA"]);
      const idxUnidad = colIndex(header.headers, ["UNIDAD"]);
      const idxGrupo = colIndex(header.headers, ["AREA", "GRUPO"]);
      const idxFechaCompromiso = colIndex(header.headers, ["FECHA COMPROMISO", "FECHA DE COMPROMISO"]);
      const idxFechaEjecucion = colIndex(header.headers, ["FECHA EJECUCION", "FECHA DE EJECUCION", "ULTIMA INTERVENCION"]);
      const idxEstado = colIndex(header.headers, ["ESTADO", "ESTATUS"]);
      if (idxDescripcion === -1) continue;
      const nuevos = [];
      for (let r = header.rowIndex + 1; r < rows.length; r++) {
        const row = rows[r];
        const descripcion = row[idxDescripcion];
        if (!descripcion || String(descripcion).trim() === "") continue;
        const numero = idxNumero !== -1 && row[idxNumero] ? String(row[idxNumero]).trim() : `IMP-${imported + 1}`;
        const grupoTxt = idxGrupo !== -1 ? row[idxGrupo] : "";
        const grupo = matchGroupId(grupoTxt) || matchGroupId(sheetName) || groups[0]?.id;
        const unidad = idxUnidad !== -1 ? String(row[idxUnidad] || "").trim() : "";
        const fecha_compromiso = idxFechaCompromiso !== -1 ? excelValueToISO(row[idxFechaCompromiso]) : "";
        const fecha_ejecucion = idxFechaEjecucion !== -1 ? excelValueToISO(row[idxFechaEjecucion]) : "";
        const estadoTxt = idxEstado !== -1 ? normalizeHeader(row[idxEstado]) : "";
        let estado = "pendiente";
        if (estadoTxt.includes("COMPLET")) estado = "completada";
        else if (estadoTxt.includes("PROCES")) estado = "en_proceso";
        else if (estadoTxt.includes("ANUL")) estado = "anulada";
        const yaExiste = [...items, ...nuevos].some((it) => it.numero === numero);
        if (yaExiste) { skipped++; continue; }
        nuevos.push({ numero, descripcion: String(descripcion).trim(), unidad, grupo, fecha_compromiso: fecha_compromiso || null, fecha_ejecucion: fecha_ejecucion || null, estado, tratamiento: "" });
        imported++;
      }
      if (nuevos.length) {
        const { data } = await supabase.from("correctivas").insert(nuevos).select();
        if (data) setItems((prev) => [...prev, ...data]);
      }
      break;
    }
    return { imported, skipped };
  }
  function exportarListado() {
    exportToExcel("ordenes_correctivas.xlsx", "Correctivas", filtered.map((i) => ({
      Orden: i.numero, Descripción: i.descripcion, "Unidad usuaria": i.unidad || "",
      Grupo: groups.find((g) => g.id === i.grupo)?.nombre || i.grupo,
      "Fecha compromiso": i.fecha_compromiso || "", "Fecha ejecución": i.fecha_ejecucion || "",
      Estado: (STATUS_STYLES[i.estado] || {}).label || i.estado,
    })));
  }

  return (
    <Card title="Órdenes de servicio / correctivas" action={
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        <ImportExcelButton onFile={importarDesdeExcel} />
        <Btn size="sm" icon={Download} onClick={exportarListado} disabled={filtered.length === 0}>Exportar</Btn>
        <Btn icon={Plus} size="sm" onClick={() => setShowForm((s) => !s)}>Nueva orden</Btn>
      </div>
    }>
      {showForm && (
        <div style={{ padding: 14, borderBottom: "1px solid var(--line)", display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Field label="N° de orden">
            <input style={{ ...inputStyle, width: 110 }} value={form.numero} onChange={(e) => setForm({ ...form, numero: e.target.value })} placeholder="OS-0231" />
          </Field>
          <Field label="Descripción">
            <input style={{ ...inputStyle, width: 240 }} value={form.descripcion} onChange={(e) => setForm({ ...form, descripcion: e.target.value })} placeholder="Falla en…" />
          </Field>
          <Field label="Unidad usuaria">
            <input style={{ ...inputStyle, width: 160 }} value={form.unidad} onChange={(e) => setForm({ ...form, unidad: e.target.value })} placeholder="Ej. Quirófano 2" />
          </Field>
          <Field label="Grupo">
            <select style={{ ...inputStyle, width: 170 }} value={form.grupo} onChange={(e) => setForm({ ...form, grupo: e.target.value })}>
              {groups.map((g) => <option key={g.id} value={g.id}>{g.nombre}</option>)}
            </select>
          </Field>
          <Field label="Fecha de compromiso">
            <input type="date" style={inputStyle} value={form.fecha_compromiso} onChange={(e) => setForm({ ...form, fecha_compromiso: e.target.value })} />
          </Field>
          <Field label="Fecha última ejecución">
            <input type="date" style={inputStyle} value={form.fecha_ejecucion} onChange={(e) => setForm({ ...form, fecha_ejecucion: e.target.value })} />
          </Field>
          <Btn kind="primary" onClick={addItem}>Guardar</Btn>
        </div>
      )}

      <div style={{ padding: "10px 14px", display: "flex", gap: 8, borderBottom: "1px solid var(--line)" }}>
        <select style={{ ...inputStyle, fontSize: 12 }} value={filterGrupo} onChange={(e) => setFilterGrupo(e.target.value)}>
          <option value="todos">Todos los grupos</option>
          {groups.map((g) => <option key={g.id} value={g.id}>{g.nombre}</option>)}
        </select>
        <select style={{ ...inputStyle, fontSize: 12 }} value={filterEstado} onChange={(e) => setFilterEstado(e.target.value)}>
          <option value="todos">Todos los estados</option>
          <option value="pendiente">Pendiente</option>
          <option value="en_proceso">En proceso</option>
          <option value="completada">Completada</option>
          <option value="anulada">Anulada</option>
        </select>
      </div>

      {filtered.length === 0 ? (
        <Empty icon={PackageSearch} text="No hay órdenes correctivas registradas con estos filtros." />
      ) : (
        <table>
          <thead><tr><th>Orden</th><th>Descripción</th><th>Unidad</th><th>Grupo</th><th>Fecha compromiso</th><th>Fecha ejecución</th><th>Estado</th><th></th></tr></thead>
          <tbody>
            {filtered.map((i) => {
              const vencida = i.fecha_compromiso && i.fecha_compromiso < todayISO() && i.estado !== "completada" && i.estado !== "anulada";
              return (
                <tr key={i.id}>
                  <td style={{ fontFamily: "var(--mono)", fontWeight: 700 }}>{i.numero}</td>
                  <td>{i.descripcion}</td>
                  <td>{i.unidad || "—"}</td>
                  <td><GroupBadge id={i.grupo} /></td>
                  <td style={{ fontFamily: "var(--mono)", color: vencida ? "#B4432A" : "var(--ink)" }}>
                    {i.fecha_compromiso || "—"} {vencida && <AlertTriangle size={11} style={{ verticalAlign: -1, marginLeft: 3 }} />}
                  </td>
                  <td>
                    <input type="date" style={{ ...inputStyle, fontSize: 12, padding: "4px 6px" }} value={i.fecha_ejecucion || ""} onChange={(e) => updateItem(i.id, { fecha_ejecucion: e.target.value || null })} />
                  </td>
                  <td>
                    <select value={i.estado} onChange={(e) => updateItem(i.id, { estado: e.target.value })} style={{ ...inputStyle, fontSize: 11.5, padding: "3px 6px" }}>
                      <option value="pendiente">Pendiente</option>
                      <option value="en_proceso">En proceso</option>
                      <option value="completada">Completada</option>
                      <option value="anulada">Anulada</option>
                    </select>
                  </td>
                  <td style={{ textAlign: "right" }}><Btn kind="danger" size="sm" icon={Trash2} onClick={() => removeItem(i.id)} /></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
    </Card>
  );
}

/* ---------------------------------------------------------------------- */
/* Grupos                                                                  */
/* ---------------------------------------------------------------------- */

function GruposTab({ groups, setGroups }) {
  async function updateGroup(id, patch) {
    const { data, error } = await supabase.from("groups").update(patch).eq("id", id).select().single();
    if (!error && data) setGroups((prev) => prev.map((g) => (g.id === id ? data : g)));
  }
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 14 }}>
      {groups.map((g) => (
        <GroupCard key={g.id} group={g} onChange={(patch) => updateGroup(g.id, patch)} />
      ))}
    </div>
  );
}

function GroupCard({ group, onChange }) {
  const [supervisor, setSupervisor] = useState(group.supervisor || "");
  const [nuevoOperario, setNuevoOperario] = useState("");
  const meta = GROUP_META[group.id];
  const Icon = meta?.Icon || Users;

  function guardarSupervisor() {
    if (supervisor !== (group.supervisor || "")) onChange({ supervisor });
  }
  function addOperario() {
    if (!nuevoOperario.trim()) return;
    onChange({ operarios: [...(group.operarios || []), nuevoOperario.trim()] });
    setNuevoOperario("");
  }
  function removeOperario(idx) {
    onChange({ operarios: (group.operarios || []).filter((_, i) => i !== idx) });
  }

  return (
    <Card>
      <div style={{ padding: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
          <span style={{ width: 30, height: 30, borderRadius: 8, background: `${meta?.color || "#999"}1A`, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Icon size={16} color={meta?.color} strokeWidth={2.3} />
          </span>
          <h3 style={{ margin: 0, fontSize: 14.5, fontWeight: 700 }}>{group.nombre}</h3>
        </div>

        <Field label="Supervisor">
          <input style={{ ...inputStyle, width: "100%" }} value={supervisor} onChange={(e) => setSupervisor(e.target.value)} onBlur={guardarSupervisor} placeholder="Nombre del supervisor" />
        </Field>

        <div style={{ marginTop: 12 }}>
          <div style={{ fontSize: 12.5, fontWeight: 600, color: "var(--ink-soft)", marginBottom: 6 }}>Operarios ({(group.operarios || []).length})</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 5, marginBottom: 8 }}>
            {(group.operarios || []).map((op, idx) => (
              <div key={idx} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "var(--bg)", borderRadius: 6, padding: "5px 9px", fontSize: 12.5 }}>
                <span>{op}</span>
                <button onClick={() => removeOperario(idx)} style={{ border: "none", background: "none", color: "var(--ink-soft)", cursor: "pointer" }}><X size={13} /></button>
              </div>
            ))}
            {(!group.operarios || group.operarios.length === 0) && <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>Sin operarios registrados.</div>}
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <input style={{ ...inputStyle, flex: 1 }} placeholder="Nombre del operario" value={nuevoOperario} onChange={(e) => setNuevoOperario(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addOperario()} />
            <Btn size="sm" icon={Plus} onClick={addOperario}>Añadir</Btn>
          </div>
        </div>
      </div>
    </Card>
  );
}
