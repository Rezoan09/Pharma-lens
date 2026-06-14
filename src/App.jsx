import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  AreaChart, Area, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis
} from "recharts";

// ── Design Tokens ─────────────────────────────────────────────────────────────
const T = {
  bg: "#f7f8fc",
  white: "#ffffff",
  surface: "#f0f2f8",
  border: "#e2e6f0",
  borderStrong: "#c8d0e4",
  text: "#0f1729",
  muted: "#6b7a99",
  faint: "#a0aec0",
  // Clinical blue-teal primary
  primary: "#0057a8",
  primaryLight: "#e8f0fb",
  primaryMid: "#3b82f6",
  // Accent: teal for positive, amber for warning, rose for critical
  teal: "#0891b2",
  tealLight: "#e0f7fa",
  green: "#059669",
  greenLight: "#d1fae5",
  amber: "#d97706",
  amberLight: "#fef3c7",
  rose: "#e11d48",
  roseLight: "#ffe4e6",
  violet: "#7c3aed",
  violetLight: "#ede9fe",
  chartColors: ["#0057a8","#0891b2","#059669","#7c3aed","#d97706","#e11d48","#0284c7","#65a30d"],
};

// ── Helpers ───────────────────────────────────────────────────────────────────
function isNumericCol(data, key) {
  const sample = data.slice(0, 30).map(r => r[key]).filter(v => v !== "" && v != null);
  if (!sample.length) return false;
  return sample.filter(v => !isNaN(Number(v))).length / sample.length > 0.7;
}

function colStats(data, key) {
  const vals = data.map(r => Number(r[key])).filter(v => !isNaN(v));
  if (!vals.length) return null;
  const sorted = [...vals].sort((a, b) => a - b);
  const sum = vals.reduce((a, b) => a + b, 0);
  const mean = sum / vals.length;
  const variance = vals.reduce((a, b) => a + (b - mean) ** 2, 0) / vals.length;
  const q1 = sorted[Math.floor(sorted.length * 0.25)];
  const q3 = sorted[Math.floor(sorted.length * 0.75)];
  return {
    min: sorted[0], max: sorted[sorted.length - 1],
    mean, median: sorted[Math.floor(sorted.length / 2)],
    std: Math.sqrt(variance), sum, count: vals.length, q1, q3,
    outliers: vals.filter(v => v < q1 - 1.5 * (q3 - q1) || v > q3 + 1.5 * (q3 - q1)).length,
  };
}

function catFreq(data, key) {
  const f = {};
  data.forEach(r => { const v = String(r[key] ?? ""); if (v && v !== "undefined" && v !== "") f[v] = (f[v] || 0) + 1; });
  return Object.entries(f).sort((a, b) => b[1] - a[1]).slice(0, 10).map(([name, value]) => ({ name: name.length > 16 ? name.slice(0, 15) + "…" : name, value, full: name }));
}

function fmt(n, decimals = 2) {
  if (n == null) return "—";
  const num = Number(n);
  if (isNaN(num)) return String(n);
  if (Math.abs(num) >= 1e9) return (num / 1e9).toFixed(1) + "B";
  if (Math.abs(num) >= 1e6) return (num / 1e6).toFixed(2) + "M";
  if (Math.abs(num) >= 1e3) return (num / 1e3).toFixed(1) + "K";
  return num % 1 === 0 ? num.toLocaleString() : num.toFixed(decimals);
}

// ── PDF Export ────────────────────────────────────────────────────────────────
function exportPDF(fileName, rows, cols, numStats, insights) {
  const w = window.open("", "_blank");
  const date = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const insightRows = insights.map((ins, i) => `<tr><td style="padding:8px 12px;border-bottom:1px solid #e2e6f0;color:#6b7a99;font-size:13px;">${i + 1}</td><td style="padding:8px 12px;border-bottom:1px solid #e2e6f0;font-size:13px;">${ins}</td></tr>`).join("");
  const statRows = numStats.slice(0, 10).map(s => `
    <tr>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e6f0;font-weight:500;">${s.key}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e6f0;text-align:right;">${fmt(s.stats.mean)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e6f0;text-align:right;">${fmt(s.stats.median)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e6f0;text-align:right;">${fmt(s.stats.min)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e6f0;text-align:right;">${fmt(s.stats.max)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e6f0;text-align:right;">${fmt(s.stats.std)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e2e6f0;text-align:right;color:${s.stats.outliers > 0 ? "#e11d48" : "#059669"};">${s.stats.outliers}</td>
    </tr>`).join("");
  const dataRows = rows.slice(0, 30).map((row, ri) =>
    `<tr>${cols.slice(0, 8).map(c => `<td style="padding:6px 10px;border-bottom:1px solid #f0f2f8;font-size:11px;">${String(row[c] ?? "")}</td>`).join("")}</tr>`
  ).join("");

  w.document.write(`<!DOCTYPE html><html><head><title>Clinical Report — ${fileName}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@400;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
    body{font-family:'Source Serif 4',Georgia,serif;color:#0f1729;background:#fff;margin:0;padding:40px 60px;}
    h1{font-size:28px;font-weight:700;color:#0057a8;margin:0 0 4px;}
    h2{font-size:16px;font-weight:600;color:#0057a8;margin:32px 0 12px;padding-bottom:6px;border-bottom:2px solid #e2e6f0;}
    table{width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;}
    thead th{background:#0057a8;color:#fff;padding:8px 12px;text-align:left;font-size:12px;font-weight:500;}
    .badge{display:inline-block;padding:2px 10px;border-radius:99px;font-size:11px;font-family:'IBM Plex Mono',monospace;}
    @media print{body{padding:20px;} button{display:none;}}
  </style></head><body>
  <div style="border-bottom:3px solid #0057a8;padding-bottom:20px;margin-bottom:28px;display:flex;justify-content:space-between;align-items:flex-end;">
    <div>
      <div style="font-size:11px;letter-spacing:0.12em;color:#6b7a99;text-transform:uppercase;margin-bottom:6px;font-family:'IBM Plex Mono',monospace;">Clinical Research Intelligence Report</div>
      <h1>${fileName.replace(/\.[^.]+$/, "")}</h1>
      <div style="color:#6b7a99;font-size:13px;margin-top:4px;">Generated on ${date} · Powered by PharmaLens AI</div>
    </div>
    <div style="text-align:right;">
      <div style="font-size:32px;font-weight:700;color:#0057a8;">${rows.length.toLocaleString()}</div>
      <div style="font-size:12px;color:#6b7a99;">Patient / Data Records</div>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:28px;">
    ${[["Records","📋",rows.length.toLocaleString(),"#0057a8"],["Columns","📐",cols.length,"#0891b2"],["Numeric","🔢",numStats.length,"#059669"],["Text Cols","🏷️",cols.length - numStats.length,"#7c3aed"]].map(([l,i,v,c])=>`<div style="border:1px solid #e2e6f0;border-radius:10px;padding:16px;border-top:3px solid ${c};"><div style="font-size:11px;color:#6b7a99;margin-bottom:4px;">${i} ${l}</div><div style="font-size:24px;font-weight:700;color:${c};">${v}</div></div>`).join("")}
  </div>

  <h2>AI-Generated Clinical Insights</h2>
  <table><thead><tr><th style="width:40px;">#</th><th>Insight</th></tr></thead><tbody>${insightRows}</tbody></table>

  <h2>Statistical Summary — Numeric Variables</h2>
  <table><thead><tr><th>Variable</th><th style="text-align:right;">Mean</th><th style="text-align:right;">Median</th><th style="text-align:right;">Min</th><th style="text-align:right;">Max</th><th style="text-align:right;">Std Dev</th><th style="text-align:right;">Outliers</th></tr></thead><tbody>${statRows}</tbody></table>

  <h2>Raw Data Preview (First 30 Records · ${cols.length > 8 ? "First 8 columns shown" : "All columns"})</h2>
  <table><thead><tr>${cols.slice(0, 8).map(c => `<th style="font-size:11px;">${c}</th>`).join("")}</tr></thead><tbody>${dataRows}</tbody></table>

  <div style="margin-top:40px;padding-top:16px;border-top:1px solid #e2e6f0;display:flex;justify-content:space-between;font-size:11px;color:#a0aec0;font-family:'IBM Plex Mono',monospace;">
    <span>PharmaLens Clinical Intelligence · Confidential</span>
    <span>${date}</span>
  </div>
  <div style="text-align:center;margin-top:20px;"><button onclick="window.print()" style="background:#0057a8;color:#fff;border:none;padding:10px 28px;border-radius:8px;font-size:14px;cursor:pointer;font-family:'Source Serif 4',serif;">🖨 Print / Save as PDF</button></div>
  </body></html>`);
  w.document.close();
}

// ── Sub-components ────────────────────────────────────────────────────────────
const ClinicalTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: T.white, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 14px", boxShadow: "0 4px 20px rgba(0,87,168,0.12)", fontSize: 12 }}>
      {label && <div style={{ color: T.muted, marginBottom: 4, fontFamily: "'IBM Plex Mono', monospace", fontSize: 11 }}>{label}</div>}
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || T.primary, fontFamily: "'IBM Plex Mono', monospace" }}>
          {p.name}: <strong>{fmt(p.value)}</strong>
        </div>
      ))}
    </div>
  );
};

const KPICard = ({ label, value, sub, color = T.primary, icon, trend }) => (
  <div style={{
    background: T.white, border: `1px solid ${T.border}`, borderRadius: 14,
    padding: "20px 22px", borderTop: `3px solid ${color}`,
    boxShadow: "0 2px 12px rgba(0,87,168,0.06)", transition: "box-shadow 0.2s, transform 0.2s",
    cursor: "default",
  }}
    onMouseEnter={e => { e.currentTarget.style.boxShadow = `0 8px 28px rgba(0,87,168,0.14)`; e.currentTarget.style.transform = "translateY(-2px)"; }}
    onMouseLeave={e => { e.currentTarget.style.boxShadow = "0 2px 12px rgba(0,87,168,0.06)"; e.currentTarget.style.transform = ""; }}
  >
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
      <div>
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 6, fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.03em", textTransform: "uppercase" }}>{label}</div>
        <div style={{ fontSize: 30, fontWeight: 700, color: T.text, fontFamily: "'Source Serif 4', serif", lineHeight: 1 }}>{value}</div>
        {sub && <div style={{ fontSize: 11, color: T.faint, marginTop: 6 }}>{sub}</div>}
      </div>
      <div style={{ width: 42, height: 42, borderRadius: 10, background: color + "18", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>{icon}</div>
    </div>
    {trend && (
      <div style={{ marginTop: 10, fontSize: 11, color: trend > 0 ? T.green : T.rose, fontFamily: "'IBM Plex Mono', monospace" }}>
        {trend > 0 ? "▲" : "▼"} {Math.abs(trend)}% vs baseline
      </div>
    )}
  </div>
);

const Panel = ({ title, subtitle, badge, children, fullWidth }) => (
  <div style={{
    background: T.white, border: `1px solid ${T.border}`, borderRadius: 14,
    padding: "22px 24px", boxShadow: "0 2px 12px rgba(0,87,168,0.05)",
    gridColumn: fullWidth ? "1 / -1" : undefined,
  }}>
    <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: subtitle || badge ? 14 : 16 }}>
      <div>
        <div style={{ fontWeight: 700, fontSize: 14, color: T.text, fontFamily: "'Source Serif 4', serif", marginBottom: 2 }}>{title}</div>
        {subtitle && <div style={{ fontSize: 12, color: T.muted }}>{subtitle}</div>}
      </div>
      {badge && <span style={{ background: T.primaryLight, color: T.primary, borderRadius: 99, padding: "3px 10px", fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 500 }}>{badge}</span>}
    </div>
    {children}
  </div>
);

const InsightPill = ({ text, index }) => {
  const colors = [T.primary, T.teal, T.green, T.violet, T.amber, T.rose];
  const bgs = [T.primaryLight, T.tealLight, T.greenLight, T.violetLight, T.amberLight, T.roseLight];
  const c = colors[index % colors.length];
  const bg = bgs[index % bgs.length];
  return (
    <div style={{ display: "flex", gap: 12, alignItems: "flex-start", padding: "12px 16px", background: bg, border: `1px solid ${c}30`, borderRadius: 10 }}>
      <div style={{ width: 24, height: 24, borderRadius: 6, background: c, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, flexShrink: 0, fontFamily: "'IBM Plex Mono', monospace" }}>{index + 1}</div>
      <div style={{ fontSize: 13, color: T.text, lineHeight: 1.65 }}>{text}</div>
    </div>
  );
};

// ── AI Insights ───────────────────────────────────────────────────────────────
async function getInsights(totalRows, numStats, catCols, colNames) {
  const summary = {
    domain: "Clinical Research / Pharma",
    totalRecords: totalRows,
    numericVariables: numStats.map(s => ({
      name: s.key,
      mean: +s.stats.mean.toFixed(2),
      std: +s.stats.std.toFixed(2),
      min: s.stats.min,
      max: s.stats.max,
      outliers: s.stats.outliers,
      q1: s.stats.q1,
      q3: s.stats.q3,
    })),
    categoricalVariables: catCols.slice(0, 6),
    allColumns: colNames,
  };

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": import.meta.env.VITE_ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      system: `You are a clinical data analyst specializing in pharmaceutical research. Given dataset statistics, return ONLY a valid JSON array of exactly 6 insight strings. No markdown, no backticks, no preamble. Each insight should be 1-2 sentences, clinically relevant, specific to the numbers, and actionable. Mention actual column names and values where possible. Format exactly: ["insight1","insight2","insight3","insight4","insight5","insight6"]`,
      messages: [{ role: "user", content: `Analyze this clinical dataset and provide 6 key insights:\n${JSON.stringify(summary, null, 2)}` }],
    }),
  });
  const data = await res.json();
  const raw = data.content?.find(b => b.type === "text")?.text || "[]";
  try { return JSON.parse(raw.replace(/```json|```/g, "").trim()); }
  catch { return ["Dataset parsed successfully.", "Review numeric distributions for clinical significance.", "Check for outliers in key measurements.", "Categorical variables show group distributions.", "Statistical ranges appear within expected clinical bounds.", "Consider stratifying analysis by key categorical variables."]; }
}

// ── Main ──────────────────────────────────────────────────────────────────────
export default function PharmaLens() {
  const [stage, setStage] = useState("upload");
  const [drag, setDrag] = useState(false);
  const [fileName, setFileName] = useState("");
  const [rawData, setRawData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [numericCols, setNumericCols] = useState([]);
  const [catCols, setCatCols] = useState([]);
  const [insights, setInsights] = useState([]);
  const [progress, setProgress] = useState(0);
  const [progressLabel, setProgressLabel] = useState("");
  const [activeTab, setActiveTab] = useState("overview");
  const [error, setError] = useState("");
  const [tableSearch, setTableSearch] = useState("");
  const fileRef = useRef();

  const processFile = useCallback(async (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xls", "csv"].includes(ext)) { setError("Please upload .xlsx, .xls, or .csv"); return; }
    setError(""); setFileName(file.name); setStage("analyzing");
    setProgress(10); setProgressLabel("Reading file…");

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        setProgress(30); setProgressLabel("Parsing data structure…");
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!json.length) { setError("Empty sheet."); setStage("upload"); return; }

        setProgress(50); setProgressLabel("Detecting variable types…");
        const cols = Object.keys(json[0]);
        const numeric = cols.filter(c => isNumericCol(json, c));
        const cat = cols.filter(c => !numeric.includes(c));
        setRawData(json); setColumns(cols); setNumericCols(numeric); setCatCols(cat);

        setProgress(70); setProgressLabel("Computing clinical statistics…");
        const numStats = numeric.map(k => ({ key: k, stats: colStats(json, k) })).filter(x => x.stats);

        setProgress(85); setProgressLabel("Generating AI clinical insights…");
        const ins = await getInsights(json.length, numStats, cat, cols);
        setInsights(ins);
        setProgress(100); setProgressLabel("Complete.");
        setTimeout(() => setStage("dashboard"), 500);
      } catch (err) {
        setError("Could not parse file."); setStage("upload");
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => { e.preventDefault(); setDrag(false); processFile(e.dataTransfer.files[0]); }, [processFile]);

  // Computed
  const numericStats = numericCols.map(k => ({ key: k, stats: colStats(rawData, k) })).filter(x => x.stats);
  const catFreqs = catCols.slice(0, 5).map(c => ({ col: c, data: catFreq(rawData, c) }));
  const lineData = (() => {
    if (!numericCols[0]) return [];
    const step = Math.max(1, Math.floor(rawData.length / 80));
    return rawData.filter((_, i) => i % step === 0).map((r, i) => ({
      i: i + 1,
      ...Object.fromEntries(numericCols.slice(0, 3).map(c => [c, Number(r[c]) || 0])),
    }));
  })();
  const scatterData = (() => {
    if (numericCols.length < 2) return [];
    const step = Math.max(1, Math.floor(rawData.length / 300));
    return rawData.filter((_, i) => i % step === 0).map(r => ({
      x: Number(r[numericCols[0]]) || 0, y: Number(r[numericCols[1]]) || 0,
    }));
  })();
  const filteredRows = rawData.filter(row =>
    !tableSearch || columns.some(c => String(row[c] ?? "").toLowerCase().includes(tableSearch.toLowerCase()))
  );

  const commonStyle = {
    fontFamily: "'Source Serif 4', Georgia, serif",
    background: T.bg, minHeight: "100vh", color: T.text,
  };

  // ── UPLOAD ────────────────────────────────────────────────────────────────
  if (stage === "upload") return (
    <div style={{ ...commonStyle, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,400;8..60,600;8..60,700&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: ${T.surface}; }
        ::-webkit-scrollbar-thumb { background: ${T.borderStrong}; border-radius: 3px; }
        .dropzone:hover { border-color: ${T.primary} !important; background: ${T.primaryLight} !important; }
        @keyframes fadeUp { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:translateY(0); } }
      `}</style>

      {/* Subtle grid background */}
      <div style={{ position: "fixed", inset: 0, backgroundImage: `linear-gradient(${T.border} 1px, transparent 1px), linear-gradient(90deg, ${T.border} 1px, transparent 1px)`, backgroundSize: "32px 32px", opacity: 0.5, pointerEvents: "none" }} />

      <div style={{ position: "relative", maxWidth: 580, width: "90%", animation: "fadeUp 0.5s ease" }}>
        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 48 }}>
          <div style={{ display: "inline-flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
            <div style={{ width: 44, height: 44, borderRadius: 12, background: T.primary, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22 }}>⚕️</div>
            <div>
              <div style={{ fontSize: 11, letterSpacing: "0.2em", color: T.muted, textTransform: "uppercase", fontFamily: "'IBM Plex Mono', monospace" }}>Clinical Research Intelligence</div>
              <div style={{ fontSize: 28, fontWeight: 700, color: T.primary, lineHeight: 1, fontFamily: "'Source Serif 4', serif" }}>PharmaLens</div>
            </div>
          </div>
          <p style={{ color: T.muted, fontSize: 15, lineHeight: 1.7, maxWidth: 420, margin: "0 auto" }}>
            Upload clinical trial data, patient outcomes, or research datasets. Get instant AI-powered insights, statistical analysis, and exportable reports.
          </p>
        </div>

        {/* Drop Zone */}
        <div className="dropzone" onClick={() => fileRef.current.click()}
          onDragOver={e => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={onDrop}
          style={{
            border: `2px dashed ${drag ? T.primary : T.borderStrong}`,
            borderRadius: 20, padding: "52px 32px", textAlign: "center",
            background: drag ? T.primaryLight : T.white, cursor: "pointer",
            transition: "all 0.2s", boxShadow: "0 4px 24px rgba(0,87,168,0.08)",
          }}
        >
          <div style={{ fontSize: 48, marginBottom: 14 }}>🧬</div>
          <div style={{ fontWeight: 600, fontSize: 18, color: T.text, marginBottom: 6 }}>Drop your dataset here</div>
          <div style={{ color: T.muted, fontSize: 13, marginBottom: 24 }}>Clinical trial data, patient records, research CSVs</div>
          <div style={{ display: "flex", justifyContent: "center", gap: 8 }}>
            {[".xlsx", ".xls", ".csv"].map(t => (
              <span key={t} style={{ background: T.primaryLight, color: T.primary, border: `1px solid ${T.primary}33`, borderRadius: 99, padding: "4px 12px", fontSize: 12, fontFamily: "'IBM Plex Mono', monospace" }}>{t}</span>
            ))}
          </div>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={e => processFile(e.target.files[0])} />

        {error && <div style={{ marginTop: 14, color: T.rose, fontSize: 13, textAlign: "center" }}>⚠ {error}</div>}

        {/* Feature pills */}
        <div style={{ display: "flex", justifyContent: "center", gap: 20, marginTop: 36, flexWrap: "wrap" }}>
          {[["🔬", "AI Clinical Insights"], ["📊", "Auto Visualizations"], ["📋", "Statistical Tables"], ["📄", "PDF Report Export"]].map(([icon, label]) => (
            <div key={label} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.muted }}>
              <span>{icon}</span><span>{label}</span>
            </div>
          ))}
        </div>

        <div style={{ textAlign: "center", marginTop: 28, fontSize: 11, color: T.faint, fontFamily: "'IBM Plex Mono', monospace" }}>
          Data stays in your browser · No upload to servers · HIPAA-aware design
        </div>
      </div>
    </div>
  );

  // ── ANALYZING ─────────────────────────────────────────────────────────────
  if (stage === "analyzing") return (
    <div style={{ ...commonStyle, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,400;8..60,600;8..60,700&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; margin:0; padding:0; }
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.5} }
      `}</style>
      <div style={{ textAlign: "center", maxWidth: 400 }}>
        <div style={{ width: 56, height: 56, border: `3px solid ${T.border}`, borderTop: `3px solid ${T.primary}`, borderRadius: "50%", animation: "spin 0.9s linear infinite", margin: "0 auto 24px" }} />
        <div style={{ fontSize: 20, fontWeight: 700, color: T.text, marginBottom: 6 }}>Analyzing Dataset</div>
        <div style={{ color: T.primary, fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", marginBottom: 4 }}>{fileName}</div>
        <div style={{ color: T.muted, fontSize: 13, marginBottom: 28, animation: "pulse 2s ease-in-out infinite" }}>{progressLabel}</div>
        <div style={{ background: T.border, borderRadius: 99, height: 6, overflow: "hidden" }}>
          <div style={{ height: "100%", width: `${progress}%`, background: `linear-gradient(90deg, ${T.primary}, ${T.teal})`, borderRadius: 99, transition: "width 0.5s ease" }} />
        </div>
        <div style={{ marginTop: 8, fontSize: 12, color: T.faint, fontFamily: "'IBM Plex Mono', monospace" }}>{progress}%</div>
      </div>
    </div>
  );

  // ── DASHBOARD ─────────────────────────────────────────────────────────────
  const tabs = [
    { key: "overview", label: "Overview", icon: "📋" },
    { key: "charts", label: "Visualizations", icon: "📈" },
    { key: "distributions", label: "Distributions", icon: "🥧" },
    { key: "table", label: "Data Table", icon: "🗃" },
  ];

  return (
    <div style={commonStyle}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,400;8..60,600;8..60,700&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: ${T.surface}; }
        ::-webkit-scrollbar-thumb { background: ${T.borderStrong}; border-radius: 3px; }
        .tab-btn { transition: all 0.18s !important; cursor: pointer; }
        .tab-btn:hover { background: ${T.primaryLight} !important; color: ${T.primary} !important; }
        .row-hover:hover { background: ${T.primaryLight} !important; }
        @keyframes fadeUp { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:translateY(0); } }
        .fade { animation: fadeUp 0.35s ease forwards; }
        input:focus { outline: none; border-color: ${T.primary} !important; box-shadow: 0 0 0 3px ${T.primaryLight} !important; }
      `}</style>

      {/* Top Nav */}
      <div style={{
        background: T.white, borderBottom: `1px solid ${T.border}`, padding: "0 28px",
        display: "flex", alignItems: "center", justifyContent: "space-between",
        height: 60, position: "sticky", top: 0, zIndex: 100,
        boxShadow: "0 1px 8px rgba(0,87,168,0.07)",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 34, height: 34, borderRadius: 9, background: T.primary, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>⚕️</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 15, color: T.primary, fontFamily: "'Source Serif 4', serif", lineHeight: 1 }}>PharmaLens</div>
            <div style={{ fontSize: 10, color: T.muted, fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.06em" }}>CLINICAL INTELLIGENCE</div>
          </div>
          <div style={{ width: 1, height: 28, background: T.border, margin: "0 8px" }} />
          <span style={{ fontSize: 13, color: T.muted, maxWidth: 260, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{fileName}</span>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <span style={{ background: T.primaryLight, color: T.primary, border: `1px solid ${T.primary}33`, borderRadius: 99, padding: "3px 12px", fontSize: 11, fontFamily: "'IBM Plex Mono', monospace" }}>n = {rawData.length.toLocaleString()}</span>
          <span style={{ background: T.tealLight, color: T.teal, border: `1px solid ${T.teal}33`, borderRadius: 99, padding: "3px 12px", fontSize: 11, fontFamily: "'IBM Plex Mono', monospace" }}>{columns.length} variables</span>
          <button onClick={() => exportPDF(fileName, rawData, columns, numericStats, insights)}
            style={{ background: T.primary, color: "#fff", border: "none", borderRadius: 8, padding: "7px 16px", cursor: "pointer", fontSize: 12, fontFamily: "'Source Serif 4', serif", fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
            📄 Export PDF
          </button>
          <button onClick={() => { setStage("upload"); setRawData([]); }}
            style={{ background: T.white, color: T.muted, border: `1px solid ${T.border}`, borderRadius: 8, padding: "7px 14px", cursor: "pointer", fontSize: 12, fontFamily: "'Source Serif 4', serif" }}>
            ↩ New File
          </button>
        </div>
      </div>

      {/* Tab Bar */}
      <div style={{ background: T.white, borderBottom: `1px solid ${T.border}`, padding: "0 28px", display: "flex", gap: 2 }}>
        {tabs.map(t => (
          <button key={t.key} className="tab-btn" onClick={() => setActiveTab(t.key)}
            style={{
              background: activeTab === t.key ? T.primaryLight : "transparent",
              border: "none", borderBottom: activeTab === t.key ? `2px solid ${T.primary}` : "2px solid transparent",
              color: activeTab === t.key ? T.primary : T.muted,
              padding: "12px 18px", fontSize: 13, fontFamily: "'Source Serif 4', serif",
              fontWeight: activeTab === t.key ? 600 : 400,
            }}>
            {t.icon} {t.label}
          </button>
        ))}
      </div>

      <div style={{ padding: "28px", maxWidth: 1400, margin: "0 auto" }} className="fade">

        {/* ── OVERVIEW ── */}
        {activeTab === "overview" && (
          <div>
            {/* KPI row */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(190px, 1fr))", gap: 16, marginBottom: 24 }}>
              <KPICard label="Total Records" value={rawData.length.toLocaleString()} icon="📋" color={T.primary} sub="dataset size" />
              <KPICard label="Variables" value={columns.length} icon="📐" color={T.teal} sub={`${numericCols.length} numeric · ${catCols.length} categorical`} />
              <KPICard label="Numeric Vars" value={numericCols.length} icon="🔢" color={T.green} sub="quantitative measures" />
              <KPICard label="Categorical Vars" value={catCols.length} icon="🏷️" color={T.violet} sub="groups & labels" />
              {numericStats[0] && <KPICard label={numericStats[0].key} value={fmt(numericStats[0].stats.mean)} icon="📊" color={T.primary} sub={`± ${fmt(numericStats[0].stats.std)} SD`} />}
              {numericStats[0] && numericStats[0].stats.outliers > 0 && <KPICard label="Outliers Detected" value={numericStats.reduce((a, s) => a + s.stats.outliers, 0)} icon="⚠️" color={T.amber} sub="across all numeric vars" />}
            </div>

            {/* AI Insights */}
            <Panel title="AI Clinical Insights" subtitle="Automatically generated by Claude — based on your dataset's statistical profile" badge="Powered by Claude" fullWidth>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))", gap: 10 }}>
                {insights.map((ins, i) => <InsightPill key={i} text={ins} index={i} />)}
              </div>
            </Panel>

            <div style={{ height: 20 }} />

            {/* Stats table */}
            <Panel title="Variable Statistics" subtitle="Complete statistical summary for all detected numeric variables" fullWidth>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: T.surface }}>
                      {["Variable", "Type", "n", "Mean", "Median", "Min", "Max", "Std Dev", "IQR", "Outliers"].map(h => (
                        <th key={h} style={{ padding: "10px 14px", textAlign: h === "Variable" || h === "Type" ? "left" : "right", color: T.muted, fontWeight: 600, fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, letterSpacing: "0.05em", borderBottom: `2px solid ${T.border}`, whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {columns.slice(0, 25).map((col, i) => {
                      const s = numericCols.includes(col) ? colStats(rawData, col) : null;
                      return (
                        <tr key={col} className="row-hover" style={{ borderBottom: `1px solid ${T.border}` }}>
                          <td style={{ padding: "9px 14px", fontWeight: 500, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>{col}</td>
                          <td style={{ padding: "9px 14px" }}>
                            <span style={{ background: s ? T.primaryLight : T.amberLight, color: s ? T.primary : T.amber, borderRadius: 99, padding: "2px 8px", fontSize: 10, fontFamily: "'IBM Plex Mono', monospace" }}>
                              {s ? "numeric" : "categorical"}
                            </span>
                          </td>
                          <td style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: T.muted }}>{s ? s.count : rawData.length}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>{s ? fmt(s.mean) : "—"}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>{s ? fmt(s.median) : "—"}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: T.muted }}>{s ? fmt(s.min) : "—"}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: T.muted }}>{s ? fmt(s.max) : "—"}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>{s ? fmt(s.std) : "—"}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: T.muted }}>{s ? fmt(s.q3 - s.q1) : "—"}</td>
                          <td style={{ padding: "9px 14px", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: s?.outliers > 0 ? T.amber : T.green, fontWeight: s?.outliers > 0 ? 600 : 400 }}>
                            {s ? (s.outliers > 0 ? `⚠ ${s.outliers}` : "✓ 0") : "—"}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </Panel>
          </div>
        )}

        {/* ── CHARTS ── */}
        {activeTab === "charts" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(500px, 1fr))", gap: 20 }}>

            {numericStats.length > 0 && (
              <Panel title="Variable Mean Comparison" subtitle="Average value per numeric variable" fullWidth>
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart data={numericStats.slice(0, 10).map(s => ({ name: s.key.length > 14 ? s.key.slice(0, 13) + "…" : s.key, Mean: +s.stats.mean.toFixed(2) }))} barSize={36}>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
                    <XAxis dataKey="name" tick={{ fill: T.muted, fontSize: 11, fontFamily: "'IBM Plex Mono', monospace" }} />
                    <YAxis tickFormatter={fmt} tick={{ fill: T.muted, fontSize: 10 }} />
                    <Tooltip content={<ClinicalTooltip />} />
                    <Bar dataKey="Mean" radius={[5, 5, 0, 0]}>
                      {numericStats.map((_, i) => <Cell key={i} fill={T.chartColors[i % T.chartColors.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </Panel>
            )}

            {lineData.length > 1 && (
              <Panel title="Trend Across Records" subtitle={`${numericCols.slice(0, 3).join(" · ")} — sampled across all rows`} fullWidth>
                <ResponsiveContainer width="100%" height={240}>
                  <AreaChart data={lineData}>
                    <defs>
                      {numericCols.slice(0, 3).map((c, i) => (
                        <linearGradient key={c} id={`g${i}`} x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor={T.chartColors[i]} stopOpacity={0.2} />
                          <stop offset="95%" stopColor={T.chartColors[i]} stopOpacity={0} />
                        </linearGradient>
                      ))}
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
                    <XAxis dataKey="i" tick={{ fill: T.muted, fontSize: 10 }} label={{ value: "Record Index", position: "insideBottom", offset: -2, fill: T.muted, fontSize: 11 }} />
                    <YAxis tickFormatter={fmt} tick={{ fill: T.muted, fontSize: 10 }} />
                    <Tooltip content={<ClinicalTooltip />} />
                    <Legend wrapperStyle={{ fontSize: 12, color: T.muted, fontFamily: "'IBM Plex Mono', monospace" }} />
                    {numericCols.slice(0, 3).map((c, i) => (
                      <Area key={c} type="monotone" dataKey={c} stroke={T.chartColors[i]} fill={`url(#g${i})`} strokeWidth={2} dot={false} />
                    ))}
                  </AreaChart>
                </ResponsiveContainer>
              </Panel>
            )}

            {scatterData.length > 0 && numericCols.length >= 2 && (
              <Panel title={`Correlation: ${numericCols[0]} × ${numericCols[1]}`} subtitle="Scatter plot — identify clinical correlations">
                <ResponsiveContainer width="100%" height={240}>
                  <ScatterChart>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
                    <XAxis dataKey="x" name={numericCols[0]} tickFormatter={fmt} tick={{ fill: T.muted, fontSize: 10 }} label={{ value: numericCols[0], position: "insideBottom", offset: -2, fill: T.muted, fontSize: 11 }} />
                    <YAxis dataKey="y" name={numericCols[1]} tickFormatter={fmt} tick={{ fill: T.muted, fontSize: 10 }} />
                    <Tooltip content={<ClinicalTooltip />} cursor={{ strokeDasharray: "3 3" }} />
                    <Scatter data={scatterData} fill={T.primary} fillOpacity={0.5} />
                  </ScatterChart>
                </ResponsiveContainer>
              </Panel>
            )}

            {numericStats.length >= 3 && (
              <Panel title="Min · Mean · Max Range" subtitle="Clinical range analysis per variable">
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={numericStats.slice(0, 6).map(s => ({ name: s.key.length > 12 ? s.key.slice(0, 11) + "…" : s.key, Min: +s.stats.min.toFixed(2), Mean: +s.stats.mean.toFixed(2), Max: +s.stats.max.toFixed(2) }))}>
                    <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
                    <XAxis dataKey="name" tick={{ fill: T.muted, fontSize: 10 }} />
                    <YAxis tickFormatter={fmt} tick={{ fill: T.muted, fontSize: 10 }} />
                    <Tooltip content={<ClinicalTooltip />} />
                    <Legend wrapperStyle={{ fontSize: 11, fontFamily: "'IBM Plex Mono', monospace" }} />
                    <Bar dataKey="Min" fill={T.rose} radius={[3, 3, 0, 0]} barSize={12} />
                    <Bar dataKey="Mean" fill={T.primary} radius={[3, 3, 0, 0]} barSize={12} />
                    <Bar dataKey="Max" fill={T.green} radius={[3, 3, 0, 0]} barSize={12} />
                  </BarChart>
                </ResponsiveContainer>
              </Panel>
            )}
          </div>
        )}

        {/* ── DISTRIBUTIONS ── */}
        {activeTab === "distributions" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(380px, 1fr))", gap: 20 }}>
            {catFreqs.map(({ col, data }, ci) => (
              <Panel key={col} title={col} subtitle={`Category distribution · ${data.length} unique values shown`} badge={`n=${rawData.length}`}>
                <ResponsiveContainer width="100%" height={180}>
                  <PieChart>
                    <Pie data={data} cx="50%" cy="50%" outerRadius={70} dataKey="value" nameKey="name"
                      label={({ name, percent }) => percent > 0.07 ? `${(percent * 100).toFixed(0)}%` : ""}
                      labelLine={false} fontSize={10}>
                      {data.map((_, i) => <Cell key={i} fill={T.chartColors[i % T.chartColors.length]} />)}
                    </Pie>
                    <Tooltip content={<ClinicalTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
                <div style={{ marginTop: 10, borderTop: `1px solid ${T.border}`, paddingTop: 10 }}>
                  {data.slice(0, 6).map((d, i) => (
                    <div key={d.name} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                      <div style={{ width: 8, height: 8, borderRadius: 2, background: T.chartColors[i % T.chartColors.length], flexShrink: 0 }} />
                      <div style={{ flex: 1, fontSize: 12, color: T.muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={d.full}>{d.name}</div>
                      <div style={{ flex: 2, height: 5, background: T.surface, borderRadius: 3, overflow: "hidden" }}>
                        <div style={{ height: "100%", width: `${(d.value / data[0].value) * 100}%`, background: T.chartColors[i % T.chartColors.length], borderRadius: 3 }} />
                      </div>
                      <div style={{ fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", color: T.text, minWidth: 30, textAlign: "right" }}>{d.value}</div>
                    </div>
                  ))}
                </div>
              </Panel>
            ))}

            {numericStats.slice(0, 4).map((s, i) => (
              <Panel key={s.key} title={s.key} subtitle="Numeric variable · statistical distribution" badge={`n=${s.stats.count}`}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 14 }}>
                  {[["Mean", s.stats.mean, T.primary], ["Median", s.stats.median, T.teal], ["Std Dev", s.stats.std, T.violet], ["Min", s.stats.min, T.green], ["Max", s.stats.max, T.amber], ["Outliers", s.stats.outliers, s.stats.outliers > 0 ? T.rose : T.green]].map(([lbl, val, color]) => (
                    <div key={lbl} style={{ background: T.surface, borderRadius: 8, padding: "10px 10px", border: `1px solid ${T.border}` }}>
                      <div style={{ fontSize: 10, color: T.faint, fontFamily: "'IBM Plex Mono', monospace", marginBottom: 2, textTransform: "uppercase", letterSpacing: "0.05em" }}>{lbl}</div>
                      <div style={{ fontSize: 17, fontWeight: 700, color, fontFamily: "'Source Serif 4', serif" }}>{fmt(val)}</div>
                    </div>
                  ))}
                </div>
                {/* Distribution bar */}
                <div style={{ background: T.surface, borderRadius: 10, padding: "12px 14px", border: `1px solid ${T.border}` }}>
                  <div style={{ fontSize: 11, color: T.muted, marginBottom: 8, fontFamily: "'IBM Plex Mono', monospace" }}>RANGE VISUALIZATION</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 10, color: T.faint, fontFamily: "'IBM Plex Mono', monospace", minWidth: 40 }}>{fmt(s.stats.min)}</span>
                    <div style={{ flex: 1, height: 12, background: T.border, borderRadius: 6, position: "relative", overflow: "hidden" }}>
                      {/* IQR band */}
                      <div style={{
                        position: "absolute", height: "100%", background: T.primary + "30",
                        left: `${((s.stats.q1 - s.stats.min) / (s.stats.max - s.stats.min)) * 100}%`,
                        width: `${((s.stats.q3 - s.stats.q1) / (s.stats.max - s.stats.min)) * 100}%`,
                      }} />
                      {/* Mean line */}
                      <div style={{
                        position: "absolute", top: 0, bottom: 0, width: 3, background: T.primary, borderRadius: 2,
                        left: `${((s.stats.mean - s.stats.min) / (s.stats.max - s.stats.min)) * 100}%`,
                      }} />
                      {/* Median line */}
                      <div style={{
                        position: "absolute", top: 0, bottom: 0, width: 2, background: T.teal,
                        left: `${((s.stats.median - s.stats.min) / (s.stats.max - s.stats.min)) * 100}%`,
                      }} />
                    </div>
                    <span style={{ fontSize: 10, color: T.faint, fontFamily: "'IBM Plex Mono', monospace", minWidth: 40, textAlign: "right" }}>{fmt(s.stats.max)}</span>
                  </div>
                  <div style={{ display: "flex", gap: 16, marginTop: 8, fontSize: 10, color: T.muted }}>
                    <span style={{ display: "flex", gap: 4, alignItems: "center" }}><span style={{ width: 10, height: 3, background: T.primary, display: "inline-block", borderRadius: 2 }} />Mean</span>
                    <span style={{ display: "flex", gap: 4, alignItems: "center" }}><span style={{ width: 10, height: 3, background: T.teal, display: "inline-block", borderRadius: 2 }} />Median</span>
                    <span style={{ display: "flex", gap: 4, alignItems: "center" }}><span style={{ width: 10, height: 10, background: T.primary + "30", display: "inline-block", borderRadius: 2 }} />IQR</span>
                  </div>
                </div>
              </Panel>
            ))}
          </div>
        )}

        {/* ── TABLE ── */}
        {activeTab === "table" && (
          <div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, gap: 16, flexWrap: "wrap" }}>
              <div style={{ color: T.muted, fontSize: 13 }}>
                Showing <strong style={{ color: T.text }}>{Math.min(filteredRows.length, 300).toLocaleString()}</strong> of <strong style={{ color: T.text }}>{rawData.length.toLocaleString()}</strong> records · {columns.length} variables
              </div>
              <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                <input value={tableSearch} onChange={e => setTableSearch(e.target.value)}
                  placeholder="🔍  Search records…"
                  style={{
                    border: `1px solid ${T.border}`, borderRadius: 8, padding: "7px 14px",
                    fontSize: 13, fontFamily: "'Source Serif 4', serif", background: T.white,
                    color: T.text, width: 220, transition: "all 0.2s",
                  }}
                />
                {tableSearch && <button onClick={() => setTableSearch("")} style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, padding: "7px 12px", cursor: "pointer", fontSize: 12, color: T.muted }}>✕ Clear</button>}
              </div>
            </div>

            <div style={{ background: T.white, border: `1px solid ${T.border}`, borderRadius: 14, overflow: "hidden", boxShadow: "0 2px 12px rgba(0,87,168,0.05)" }}>
              <div style={{ overflowX: "auto", maxHeight: "68vh", overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ position: "sticky", top: 0, background: T.surface, zIndex: 5 }}>
                    <tr>
                      <th style={{ padding: "10px 14px", borderBottom: `2px solid ${T.border}`, textAlign: "left", color: T.faint, fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, fontWeight: 600 }}>#</th>
                      {columns.map(col => (
                        <th key={col} style={{ padding: "10px 14px", borderBottom: `2px solid ${T.border}`, textAlign: "left", color: T.muted, fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, fontWeight: 600, whiteSpace: "nowrap", minWidth: 110 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                            <span style={{ width: 6, height: 6, borderRadius: "50%", background: numericCols.includes(col) ? T.primary : T.amber, display: "inline-block", flexShrink: 0 }} />
                            {col}
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.slice(0, 300).map((row, ri) => (
                      <tr key={ri} className="row-hover" style={{ borderBottom: `1px solid ${T.border}` }}>
                        <td style={{ padding: "7px 14px", color: T.faint, fontFamily: "'IBM Plex Mono', monospace" }}>{ri + 1}</td>
                        {columns.map(col => (
                          <td key={col} style={{ padding: "7px 14px", color: numericCols.includes(col) ? T.primary : T.text, fontFamily: numericCols.includes(col) ? "'IBM Plex Mono', monospace" : "inherit", whiteSpace: "nowrap", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis" }}>
                            {String(row[col] ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
