"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { FormEvent } from "react";
import Image from "next/image";
import * as XLSX from "xlsx";

type Cell = string | number | boolean | Date | null | undefined;
type Matrix = Cell[][];
type Sheets = Record<string, Matrix>;
type MovementSortKey = "date" | "provider" | "department" | "expense" | "concept" | "subtotal" | "total";

const BRANCHES = [
  { id: "machinery", label: "Kit Machinery", short: "KM", sheet: "P&L KIT MACHINERY" },
  { id: "express", label: "Kit Express", short: "KE", sheet: "P&L KIT EXPRESS" },
  { id: "conversion", label: "Conversión", short: "CV", sheet: "P&L CONVERSION" },
  { id: "distribution", label: "Distribución", short: "DIST", sheet: "P&L DISTRIBUCION QRO" },
  { id: "qro", label: "QRO", short: "QRO", sheet: "P&L DISTRIBUCION QRO" },
  { id: "puebla", label: "PUE", short: "PUE", sheet: "P&L DISTRIBUCION PUEBLA" },
  { id: "guadalajara", label: "GDL", short: "GDL", sheet: "P&L DISTRIBUCION GUADALAJARA" },
  { id: "slp", label: "SLP", short: "SLP", sheet: "P&L DISTRIBUCION SLP" },
  { id: "leon", label: "LEÓN", short: "GTO", sheet: "P&L DISTRIBUCION LEON" },
] as const;

const STATE_BRANCHES = new Set(["qro", "puebla", "guadalajara", "slp", "leon"]);
const DISTRIBUTION_BRANCH_CODES = new Set(["Q", "P", "G", "S", "L"]);
const DISTRIBUTION_SHEETS = [
  "P&L DISTRIBUCION QRO",
  "P&L DISTRIBUCION PUEBLA",
  "P&L DISTRIBUCION GUADALAJARA",
  "P&L DISTRIBUCION SLP",
  "P&L DISTRIBUCION LEON",
];

const MONTHS = [
  ["Enero", 2], ["Febrero", 4], ["Marzo", 6], ["Abril", 8],
  ["Mayo", 10], ["Junio", 12], ["Julio", 14], ["Agosto", 16],
  ["Septiembre", 18], ["Octubre", 20], ["Noviembre", 22], ["Diciembre", 24],
] as const;

const money = new Intl.NumberFormat("es-MX", { style: "currency", currency: "MXN", minimumFractionDigits: 2, maximumFractionDigits: 2 });
const moneyFull = new Intl.NumberFormat("es-MX", { style: "currency", currency: "MXN", minimumFractionDigits: 2, maximumFractionDigits: 2 });
const moneyFull2 = new Intl.NumberFormat("es-MX", { style: "currency", currency: "MXN", minimumFractionDigits: 2, maximumFractionDigits: 2 });
const moneyCompact = new Intl.NumberFormat("es-MX", { notation: "compact", maximumFractionDigits: 1 });
const PIE_COLORS = ["#2367a8", "#4d8fd0", "#77afd9", "#694da8", "#9674c5", "#d07a48", "#e2aa5a", "#2d8b79", "#70b69e"];
const EXPENSE_GROUP_LABELS: Record<string, string> = {
  GA: "GASTO ADMINISTRATIVO",
  GO: "GASTO OPERATIVO",
  GV: "GASTO DE VENTA",
  GF: "GASTO FINANCIERO",
};
const STORED_WORKBOOK_KEY = "kit-dashboard-workbook-v1";
const STORED_WORKBOOK_NAME_KEY = "kit-dashboard-workbook-name-v1";
const WORKBOOK_DB_NAME = "kit-dashboard-storage";
const AUTH_STORAGE_KEY = "kit-dashboard-auth-v1";
const APP_USERNAME = "moramay";
const APP_PASSWORD = "Afatelamordemivida2026";
const DEFAULT_WORKBOOK_PATH = process.env.GITHUB_ACTIONS === "true" ? "/kit/datos-kit.xlsx" : "/datos-kit.xlsx";

function bytesToBase64(bytes: Uint8Array) {
  let binary = "";
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
  return btoa(binary);
}

function base64ToBytes(value: string) {
  const binary = atob(value);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function saveWorkbook(data: ArrayBuffer, name: string) {
  return new Promise<void>((resolve, reject) => {
    const request = indexedDB.open(WORKBOOK_DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore("workbook");
    request.onerror = () => reject(request.error);
    request.onsuccess = () => { const transaction = request.result.transaction("workbook", "readwrite"); transaction.objectStore("workbook").put({ data, name }, "latest"); transaction.oncomplete = () => resolve(); transaction.onerror = () => reject(transaction.error); };
  });
}

function loadWorkbook() {
  return new Promise<{ data: ArrayBuffer; name: string } | null>((resolve, reject) => {
    const request = indexedDB.open(WORKBOOK_DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore("workbook");
    request.onerror = () => reject(request.error);
    request.onsuccess = () => { const transaction = request.result.transaction("workbook", "readonly"); const read = transaction.objectStore("workbook").get("latest"); read.onsuccess = () => resolve(read.result ?? null); read.onerror = () => reject(read.error); };
  });
}

function numberOf(value: Cell): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const text = String(value ?? "").trim();
  const compact = text.replace(/\s/g, "");
  if (!compact || compact === "-" || compact === "$-" || compact === "—") return 0;
  const parsed = Number(compact.replace(/[$,% ,]/g, "").replace(/^\((.*)\)$/, "-$1"));
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalize(value: Cell) {
  return String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
}

function isoDate(value: Cell) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  const parsed = new Date(String(value ?? ""));
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString().slice(0, 10);
}

function monthRange(period: number | "acc") {
  if (period === "acc") return null;
  const month = Math.max(0, Math.min(11, Math.round((period - 2) / 2)));
  const start = new Date(Date.UTC(2026, month, 1));
  const end = new Date(Date.UTC(2026, month + 1, 0));
  return { from: start.toISOString().slice(0, 10), to: end.toISOString().slice(0, 10) };
}

function belongsToBranch(row: Cell[], branchId: string) {
  if (branchId === "general") return true;
  const branch = normalize(row[10]);
  const costCenter = normalize(row[11]);
  if (branchId === "machinery") return costCenter === "M";
  if (branchId === "express") return costCenter === "E";
  if (branchId === "conversion") return costCenter === "C";
  if (branchId === "distribution") return costCenter === "L" && DISTRIBUTION_BRANCH_CODES.has(branch);
  const codes: Record<string, string> = { qro: "Q", puebla: "P", guadalajara: "G", slp: "S", leon: "L" };
  return branch === codes[branchId] && costCenter === "L";
}

function workbookToSheets(workbook: XLSX.WorkBook): Sheets {
  const sheets = Object.fromEntries(workbook.SheetNames.map((name) => [name, XLSX.utils.sheet_to_json<Cell[]>(workbook.Sheets[name], { header: 1, defval: null, raw: true })]));
  return sheets;
}

function findRow(matrix: Matrix, pattern: RegExp, fallback = -1) {
  const index = matrix.findIndex((row) => pattern.test(normalize(row?.[1])));
  return index >= 0 ? index : fallback;
}

function findRevenueRow(matrix: Matrix) {
  const headerRevenue = matrix.findIndex((row, index) => index >= 2 && index <= 3 && !/COSTO|MARGEN|GASTOS|UTILIDAD/.test(normalize(row?.[1])) && MONTHS.some(([, column]) => numberOf(row?.[column]) !== 0));
  if (headerRevenue >= 0) return headerRevenue;
  const labeled = findRow(matrix, /VENTAS|LAMINA/, -1);
  if (labeled >= 0) return labeled;
  const firstNumeric = matrix.findIndex((row, index) => index >= 2 && MONTHS.some(([, column]) => numberOf(row?.[column]) !== 0));
  return firstNumeric >= 0 ? firstNumeric : 2;
}

function availablePeriods(matrix: Matrix) {
  const revenueRow = findRevenueRow(matrix);
  const periods = MONTHS.filter(([, column]) => numberOf(matrix[revenueRow]?.[column]) !== 0);
  return periods.length ? periods : MONTHS.slice(0, 1);
}

function valueAt(matrix: Matrix, row: number, period: number | "acc") {
  return row < 0 ? 0 : numberOf(matrix[row]?.[period === "acc" ? 27 : period]);
}

function rebatesAt(matrix: Matrix, period: number | "acc") {
  return matrix.reduce((sum, row, rowIndex) => {
    // Algunos P&L de sucursal colocan el concepto en la primera columna y
    // otros en la segunda; revisamos ambas para no perder rebates.
    const label = normalize(`${row?.[0] ?? ""} ${row?.[1] ?? ""}`);
    // En los P&L de sucursal los rebates están normalizados en las filas 6 y 7
    // (índices 5 y 6), aun cuando el texto de la etiqueta venga vacío.
    if (!/^REBATE(?:\s|$)/.test(label) && !(rowIndex === 5 || rowIndex === 6)) return sum;
    if (period === "acc") return sum + MONTHS.reduce((total, [, column]) => total + numberOf(row?.[column]), 0);
    return sum + numberOf(row?.[period]);
  }, 0);
}

function percentAt(matrix: Matrix, row: number, period: number | "acc", base: number) {
  if (row < 0) return 0;
  const raw = matrix[row]?.[period === "acc" ? 28 : period + 1];
  if (raw !== null && raw !== undefined && String(raw).trim() !== "") {
    const parsed = numberOf(raw);
    return Math.abs(parsed) <= 1 ? parsed * 100 : parsed;
  }
  return base ? (valueAt(matrix, row, period) / base) * 100 : 0;
}

function KpiSparkline({ values, color }: { values: number[]; color: string }) {
  const min = Math.min(...values, 0);
  const max = Math.max(...values, 1);
  const span = max - min || 1;
  const points = values.map((value, index) => `${values.length === 1 ? 50 : index / (values.length - 1) * 100},${28 - (value - min) / span * 24}`).join(" ");
  return <svg className="kpiSparkline" viewBox="0 0 100 32" preserveAspectRatio="none" aria-hidden="true"><polygon points={`0,32 ${points} 100,32`} fill={color} opacity=".12" /><polyline points={points} fill="none" stroke={color} strokeWidth="2" vectorEffect="non-scaling-stroke" /></svg>;
}

function ExpandChartButton() {
  return <button type="button" className="chartExpandButton" aria-label="Mostrar gráfica en grande" title="Mostrar en grande" onClick={(event) => { event.stopPropagation(); if (document.fullscreenElement) { void document.exitFullscreen(); return; } const panel = event.currentTarget.closest(".panel") as HTMLElement | null; if (panel?.requestFullscreen) void panel.requestFullscreen(); }}>⛶</button>;
}

function DownloadPdfButton() {
  return <button type="button" className="pdfButton" onClick={() => window.print()} title="Descargar dashboard en PDF">Descargar PDF</button>;
}

function LoginScreen({ onLogin }: { onLogin: () => void }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (username.trim() === APP_USERNAME && password === APP_PASSWORD) {
      window.localStorage.setItem(AUTH_STORAGE_KEY, "1");
      onLogin();
    } else setLoginError("Usuario o contraseña incorrectos.");
  }
  return <main className="loginShell"><form className="loginCard" onSubmit={submit}><div className="loginMark">K</div><span className="eyebrow">KIT</span><h1>Acceso al dashboard</h1><p>Ingresa tus credenciales para continuar.</p><label>Usuario<input value={username} onChange={(event) => setUsername(event.target.value)} autoComplete="username" required /></label><label>Contraseña<input type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete="current-password" required /></label>{loginError && <div className="loginError" role="alert">{loginError}</div>}<button type="submit">Entrar</button></form></main>;
}

function MetricCard({ label, value, displayValue, secondaryValue, comparisonValue, previousValue, values, tone, icon, inverse = false, margin = false, active, onClick }: { label: string; value: number; displayValue?: string; secondaryValue?: string; comparisonValue: number; previousValue: number; values: number[]; tone: string; icon: string; inverse?: boolean; margin?: boolean; active?: boolean; onClick: () => void }) {
  const rawChange = margin ? comparisonValue - previousValue : previousValue ? (comparisonValue - previousValue) / Math.abs(previousValue) * 100 : 0;
  const favorable = inverse ? rawChange <= 0 : rawChange >= 0;
  const difference = comparisonValue - previousValue;
  return <article className={`metricCard premiumKpi ${tone} ${active ? "active" : ""}`} onClick={onClick} role="button" tabIndex={0} onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") onClick(); }}><span className="metricLabel">{label}</span><span className="kpiIcon" aria-hidden="true">{icon}</span><strong>{displayValue ?? moneyFull2.format(value)}</strong>{secondaryValue && <span className="kpiSecondaryValue">{secondaryValue}</span>}<div className={`kpiChange ${favorable ? "positive" : "negative"}`}><span>{rawChange >= 0 ? "↑" : "↓"} {Math.abs(rawChange).toFixed(1)}{margin ? " pp" : "%"}</span><small>vs. mes anterior</small></div><div className="kpiDifference">{difference >= 0 ? "+" : "−"} {margin ? `${Math.abs(difference).toFixed(2)} pp` : moneyFull2.format(Math.abs(difference))}</div><KpiSparkline values={values} color={tone === "sky" ? "#2584ff" : tone === "slate" ? "#88a1c2" : tone === "orange" ? "#ffb02e" : tone === "purple" ? "#9b4dff" : tone === "green" ? "#13d69f" : "#17d2ff"} /></article>;
}

function ExecutiveKpis({ matrix, period, selectedMetric, onSelect }: { matrix: Matrix; period: number | "acc"; selectedMetric: string; onSelect: (metric: string) => void }) {
  const periods = availablePeriods(matrix);
  const revenueRow = findRevenueRow(matrix);
  const costRow = findRow(matrix, /^COSTO$/);
  const marginRow = findRow(matrix, /MARGEN SOBRE MATERIA PRIMA/);
  const fixedRow = findRow(matrix, /GASTOS FIJOS/);
  const variableRow = findRow(matrix, /GASTOS VARIABLES/);
  const profitRow = findRow(matrix, /^UTILIDAD$/);
  const detailColumn = period === "acc" ? periods.at(-1)?.[1] ?? 2 : period;
  const previousColumn = periods[Math.max(0, periods.findIndex(([, column]) => column === detailColumn) - 1)]?.[1] ?? detailColumn;
  const monthly = (row: number) => periods.map(([, column]) => valueAt(matrix, row, column));
  const revenue = valueAt(matrix, revenueRow, period);
  const cost = valueAt(matrix, costRow, period);
  const marginPercent = percentAt(matrix, marginRow, period, revenue);
  const rebatesMarginPercent = percentAt(matrix, marginRow + 1, period, revenue);
  const rebates = rebatesAt(matrix, period);
  const expenses = valueAt(matrix, fixedRow, period) + valueAt(matrix, variableRow, period);
  const profit = valueAt(matrix, profitRow, period);
  const netMargin = revenue ? profit / revenue * 100 : 0;
  const previousRevenue = valueAt(matrix, revenueRow, previousColumn);
  const previousCost = valueAt(matrix, costRow, previousColumn);
  const previousRebates = rebatesAt(matrix, previousColumn);
  const previousExpenses = valueAt(matrix, fixedRow, previousColumn) + valueAt(matrix, variableRow, previousColumn);
  const previousProfit = valueAt(matrix, profitRow, previousColumn);
  const previousNetMargin = previousRevenue ? previousProfit / previousRevenue * 100 : 0;
  const comparisonRevenue = valueAt(matrix, revenueRow, detailColumn);
  const comparisonCost = valueAt(matrix, costRow, detailColumn);
  const comparisonRebates = rebatesAt(matrix, detailColumn);
  const comparisonExpenses = valueAt(matrix, fixedRow, detailColumn) + valueAt(matrix, variableRow, detailColumn);
  const comparisonProfit = valueAt(matrix, profitRow, detailColumn);
  const comparisonNetMargin = comparisonRevenue ? comparisonProfit / comparisonRevenue * 100 : 0;
  const monthlyRevenue = monthly(revenueRow);
  const monthlyCost = monthly(costRow);
  const monthlyMargin = monthly(marginRow);
  const monthlyRebates = periods.map(([, column]) => rebatesAt(matrix, column));
  const monthlyExpenses = periods.map(([, column]) => valueAt(matrix, fixedRow, column) + valueAt(matrix, variableRow, column));
  const monthlyProfit = monthly(profitRow);
  const monthlyNetMargin = periods.map(([, column]) => { const sales = valueAt(matrix, revenueRow, column); return sales ? valueAt(matrix, profitRow, column) / sales * 100 : 0; });
  return <section className="metricsGrid executiveKpis" aria-label="Indicadores ejecutivos">
    <MetricCard label="VENTAS" value={revenue} comparisonValue={comparisonRevenue} previousValue={previousRevenue} values={monthlyRevenue} tone="sky" icon="↗" active={selectedMetric === "revenue"} onClick={() => onSelect("revenue")} />
    <MetricCard label="COSTO" value={cost} comparisonValue={comparisonCost} previousValue={previousCost} values={monthlyCost} tone="slate" icon="◈" inverse active={selectedMetric === "cost"} onClick={() => onSelect("cost")} />
    <MetricCard label="REBATES" value={rebates} secondaryValue={`${rebatesMarginPercent.toFixed(1)}%`} comparisonValue={comparisonRebates} previousValue={previousRebates} values={monthlyRebates} tone="cyan" icon="−" active={selectedMetric === "rebates"} onClick={() => onSelect("rebates")} />
    <MetricCard label="MARGEN SOBRE MATERIA PRIMA" value={valueAt(matrix, marginRow, period)} secondaryValue={`${marginPercent.toFixed(1)}%`} comparisonValue={valueAt(matrix, marginRow, detailColumn)} previousValue={valueAt(matrix, marginRow, previousColumn)} values={monthlyMargin} tone="orange" icon="◒" active={selectedMetric === "grossMargin"} onClick={() => onSelect("grossMargin")} />
    <MetricCard label="GASTOS" value={expenses} comparisonValue={comparisonExpenses} previousValue={previousExpenses} values={monthlyExpenses} tone="purple" icon="↘" inverse active={selectedMetric === "expenses"} onClick={() => onSelect("expenses")} />
    <MetricCard label="UTILIDAD" value={profit} comparisonValue={comparisonProfit} previousValue={previousProfit} values={monthlyProfit} tone="green" icon="◆" active={selectedMetric === "profit"} onClick={() => onSelect("profit")} />
  </section>;
}

function PnlWaterfall({ revenue, cost, grossMargin, fixed, variable, rebates, profit, onSelect }: { revenue: number; cost: number; grossMargin: number; fixed: number; variable: number; rebates: number; profit: number; onSelect: () => void }) {
  const expenses = fixed + variable;
  const operatingProfit = grossMargin - expenses + rebates;
  const items = [
    { label: "Ventas", value: revenue, start: 0, end: revenue, color: "#2584ff" },
    { label: "Costo", value: -Math.abs(cost), start: revenue, end: revenue - Math.abs(cost), color: "#ff5068" },
    { label: "Margen bruto", value: grossMargin, start: 0, end: grossMargin, color: "#13d69f" },
    { label: "Gastos", value: -Math.abs(expenses), start: grossMargin, end: grossMargin - Math.abs(expenses), color: "#ff5068" },
    { label: "Rebates", value: rebates, start: grossMargin - Math.abs(expenses), end: grossMargin - Math.abs(expenses) + rebates, color: "#17d2ff" },
    { label: "Utilidad", value: operatingProfit, start: 0, end: operatingProfit, color: operatingProfit < 0 ? "#ff5068" : "#13d69f" },
    { label: "Utilidad neta", value: profit, start: 0, end: profit, color: profit < 0 ? "#ff5068" : "#2584ff" },
  ];
  const chartValues = items.flatMap((item) => [item.start, item.end]);
  const range = Math.max(1, ...chartValues.map((value) => Math.abs(value)), Math.abs(revenue));
  const chartMin = Math.min(0, ...chartValues);
  const chartMax = Math.max(range, ...chartValues);
  const chartSpan = chartMax - chartMin || 1;
  const y = (value: number) => 8 + (chartMax - value) / chartSpan * 68;
  const zero = y(0);
  const width = 100 / items.length;
  const format = (value: number) => `${value < 0 ? "-" : ""}${moneyFull2.format(Math.abs(value))}`;
  const axisLabel = (value: number) => `${value < 0 ? "-$" : "$"}${(Math.abs(value) / 1_000_000).toFixed(1)}M`;
  const [hoveredIndex, setHoveredIndex] = useState<number | null>(null);
  return <section className="panel waterfallPanel" onClick={onSelect} role="button" tabIndex={0} onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") onSelect(); }}><div className="panelHeading"><div><span className="eyebrow">PUENTE DE RENTABILIDAD</span><h3>Del ingreso a la utilidad neta</h3><p>Impacto acumulado de costos, gastos y rebates.</p></div><div className="panelHeadingActions"><span className="waterfallPeriod">Valores reales</span><ExpandChartButton /></div></div><div className="waterfallChart"><div className="waterfallYAxis">{[chartMax, chartMax * .5, 0].map((value) => <span key={value} style={{ top: `${y(value)}%` }}>{axisLabel(value)}</span>)}</div><svg viewBox="0 0 100 100" preserveAspectRatio="none" role="img" aria-label="Puente de rentabilidad"><line className="waterfallGrid" x1="7" x2="98" y1={y(chartMax)} y2={y(chartMax)} /><line className="waterfallGrid" x1="7" x2="98" y1={y(chartMax * .5)} y2={y(chartMax * .5)} /><line className="waterfallAxis" x1="7" x2="98" y1={zero} y2={zero} />{items.map((item, index) => { const x = index * width + width * .2; const barWidth = width * .6; const top = y(Math.max(item.start, item.end)); const height = Math.max(item.value === 0 ? .35 : 1.2, Math.abs(y(item.start) - y(item.end))); return <g key={item.label} onMouseEnter={() => setHoveredIndex(index)} onMouseLeave={() => setHoveredIndex(null)}><rect className="waterfallBar" x={x} y={top} width={barWidth} height={height} rx=".8" fill={item.color} />{index < items.length - 1 && <line className="waterfallConnector" x1={x + barWidth} x2={(index + 1) * width + width * .2} y1={y(item.end)} y2={y(item.end)} />}</g>; })}</svg>{hoveredIndex !== null && <div className="waterfallTooltip" style={{ left: `${(hoveredIndex + .5) / items.length * 100}%` }}><strong>{items[hoveredIndex].label}</strong><span>Monto <b>{format(items[hoveredIndex].value)}</b></span><span>% de ventas <b>{(revenue ? items[hoveredIndex].value / revenue * 100 : 0).toFixed(2)}%</b></span></div>}<div className="waterfallLabels">{items.map((item) => <div key={item.label}><strong>{item.label}</strong><b>{format(item.value)}</b></div>)}</div></div></section>;
}

function TrendChart({ matrix }: { matrix: Matrix }) {
  const revenueRow = findRevenueRow(matrix);
  const profitRow = findRow(matrix, /^UTILIDAD$/);
  const periods = availablePeriods(matrix);
  const values = periods.map(([label, column]) => ({ label: label.slice(0, 3), revenue: valueAt(matrix, revenueRow, column), profit: valueAt(matrix, profitRow, column) }));
  const chartValues = values.flatMap((item) => [item.revenue, item.profit]);
  const min = Math.min(0, ...chartValues);
  const max = Math.max(1, ...chartValues);
  const span = max - min || 1;
  const points = values.map((item, index) => {
    const x = values.length === 1 ? 50 : 8 + (index / (values.length - 1)) * 84;
    const y = 88 - (item.profit - min) / span * 76;
    return `${x},${y}`;
  }).join(" ");

  return <section className="panel trendPanel">
    <div className="panelHeading"><div><span className="eyebrow">EVOLUCIÓN 2026</span><h3>Ventas y utilidad mensual</h3></div><div className="legend"><span><i className="salesDot" /> Ventas</span><span><i className="profitDot" /> Utilidad</span></div></div>
    <div className="trendChart"><div className="gridLines"><i /><i /><i /></div><div className="bars">{values.map((item) => <div className="barColumn" key={item.label}><div className="salesBar" style={{ height: `${Math.max(4, item.revenue / max * 78)}%` }} /><span>{item.label}</span></div>)}</div><svg viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><polyline points={points} /><g>{points.split(" ").map((point, i) => { const [cx, cy] = point.split(","); return <circle key={i} cx={cx} cy={cy} r="1.25" />; })}</g></svg></div>
  </section>;
}

function expenseBreakdown(matrix: Matrix, kind: "fixed" | "variable", period: number | "acc") {
  const fixedRow = findRow(matrix, /GASTOS FIJOS/);
  const variableRow = findRow(matrix, /GASTOS VARIABLES/);
  const profitRow = findRow(matrix, /^UTILIDAD$/);
  const start = (kind === "fixed" ? fixedRow : variableRow) + 1;
  const end = kind === "fixed" ? variableRow : profitRow;
  let group = "";
  const totals = new Map<string, number>();

  matrix.slice(start, end).forEach((row, offset) => {
    const label = String(row?.[1] ?? "").trim();
    const normalized = normalize(label);
    if (/^(GA|GO|GV|GF|SUELDO DIRECCION)$/.test(normalized)) {
      group = normalized;
      return;
    }
    const value = valueAt(matrix, start + offset, period);
    if (!label || value <= 0) return;
    const groupLabel = EXPENSE_GROUP_LABELS[group] ?? group;
    const category = group === "SUELDO DIRECCION" ? "SUELDO DIRECCIÓN" : group ? `${groupLabel} · ${label}` : label;
    totals.set(category, (totals.get(category) ?? 0) + value);
  });

  return [...totals.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
}

function expenseCategoryBreakdown(matrix: Matrix, kind: "fixed" | "variable", period: number | "acc") {
  const totals = new Map<string, number>();
  expenseBreakdown(matrix, kind, period).forEach(({ label, value }) => {
    const separator = label.indexOf(" · ");
    const category = separator >= 0 ? label.slice(0, separator) : label;
    totals.set(category, (totals.get(category) ?? 0) + value);
  });
  return [...totals.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
}

function compactPie(items: { label: string; value: number }[], limit = 9) {
  if (items.length <= limit) return items;
  const visible = items.slice(0, limit - 1);
  const remainder = items.slice(limit - 1);
  return [...visible, { label: `Otros conceptos (${remainder.length})`, value: remainder.reduce((sum, item) => sum + item.value, 0) }];
}

function ExpenseDonutCard({ eyebrow, title, subtitle, items, selectedCategory, onSelect }: { eyebrow: string; title: string; subtitle: string; items: { label: string; value: number }[]; selectedCategory: string; onSelect: (category: string) => void }) {
  const slices = compactPie(items);
  const total = slices.reduce((sum, item) => sum + item.value, 0);
  const dominant = slices.reduce<{ label: string; value: number } | null>((largest, item) => !largest || item.value > largest.value ? item : largest, null);
  const dominantPercent = dominant && total ? dominant.value / total * 100 : 0;
  const dominantLabel = dominant?.label.replace(/^[A-Z]{2}\s*·\s*/, "") ?? "Sin datos";
  const tableName = title.toLowerCase().includes("categoría") ? "CATEGORÍA" : title === "Fijo vs. variable" ? "TIPO" : "CONCEPTO";
  const tablePercent = title === "Fijo vs. variable" ? "% DEL GASTO" : title.toLowerCase().includes("variables") ? "% DE VARIABLES" : title.toLowerCase().includes("fijos") ? "% DE FIJOS" : "% DEL GASTO";
  const selectItem = (label: string) => onSelect(selectedCategory === label ? "" : label);
  const selectedInChart = slices.some((item) => item.label === selectedCategory);
  let cursor = 0;
  const segments = slices.map((item, index) => { const start = cursor; const percent = total ? item.value / total * 100 : 0; cursor += percent; return { ...item, start, percent, color: PIE_COLORS[index % PIE_COLORS.length] }; });
  const centerSegment = segments.find((item) => item.label === selectedCategory) ?? segments.find((item) => item.label === dominant?.label) ?? segments[0];
  const centerPercent = centerSegment?.percent ?? dominantPercent;
  const centerLabel = centerSegment?.label.replace(/^[A-Z]{2}\s*·\s*/, "") ?? dominantLabel;

  return <section className="panel unifiedExpenseChart"><div className="panelHeading"><div><span className="eyebrow">{eyebrow}</span><h3>{title}</h3><p>{subtitle}</p></div><div className="panelHeadingActions"><strong className="unifiedExpenseTotal">{moneyFull2.format(total)}</strong><ExpandChartButton /></div></div>{total ? <div className="unifiedExpenseBody"><div className="unifiedDonut"><svg viewBox="0 0 120 120" role="group" aria-label={`${title}: ${moneyFull2.format(total)}. Selecciona un segmento para filtrar movimientos.`}>{segments.map((item) => <circle key={item.label} cx="60" cy="60" r="42" pathLength="100" fill="none" stroke={item.color} strokeWidth="22" strokeDasharray={`${item.percent} ${100 - item.percent}`} strokeDashoffset={-item.start} className={selectedInChart && selectedCategory !== item.label ? "muted" : ""} role="button" aria-label={`${item.label}: ${moneyFull2.format(item.value)}, ${item.percent.toFixed(1)}% del gasto`} tabIndex={0} onClick={() => selectItem(item.label)} onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") selectItem(item.label); }}><title>{item.label}: {moneyFull2.format(item.value)} · {item.percent.toFixed(1)}% del gasto. Haz clic para filtrar; vuelve a pulsar para reiniciar.</title></circle>)}</svg><button type="button" className="unifiedDonutCenter" aria-label="Limpiar filtro de gastos" onClick={() => onSelect("")}><strong>{centerPercent.toFixed(1)}%</strong><span title={centerLabel}>{centerLabel}</span></button></div><div className="unifiedExpenseLegend"><div className="unifiedExpenseTableHeader"><span>{tableName}</span><span>MONTO</span><span>{tablePercent}</span></div>{segments.map((item) => <button type="button" className={selectedCategory === item.label ? "active" : ""} key={item.label} onClick={() => selectItem(item.label)}><i style={{ background: item.color }} /><span title={item.label}>{item.label}</span><b>{moneyFull2.format(item.value)}</b><em>{item.percent.toFixed(1)}%</em></button>)}</div></div> : <div className="noMovements">No hay gastos en este periodo.</div>}</section>;
}

function ExpenseDistribution({ matrix, period, selectedCategory, onSelect }: { matrix: Matrix; period: number | "acc"; selectedCategory: string; onSelect: (category: string) => void }) {
  const items = useMemo(() => [...expenseBreakdown(matrix, "fixed", period), ...expenseBreakdown(matrix, "variable", period)].sort((a, b) => b.value - a.value), [matrix, period]);
  return <ExpenseTreemapCard items={items} selectedCategory={selectedCategory} onSelect={onSelect} />;
}

function ExpenseTreemapCard({ items, selectedCategory, onSelect }: { items: { label: string; value: number }[]; selectedCategory: string; onSelect: (category: string) => void }) {
  const total = items.reduce((sum, item) => sum + Math.abs(item.value), 0);
  const tiles = layoutTreemap(items, total);
  return <section className="panel unifiedExpenseChart expenseTreemapCard"><div className="panelHeading"><div><span className="eyebrow">COMPOSICIÓN DEL GASTO</span><h3>Distribución del egreso</h3><p>Gastos fijos + gastos variables, todos los conceptos.</p></div><div className="panelHeadingActions"><strong className="unifiedExpenseTotal">{moneyFull2.format(total)}</strong><ExpandChartButton /></div></div><div className="expenseTreemap" role="group" aria-label="Distribución completa del egreso: gastos fijos y gastos variables">{tiles.map(({ item, index, percent, x, y, width, height }) => <button type="button" key={item.label} className={`treemapTile ${selectedCategory === item.label ? "active" : ""}`} style={{ left: `${x}%`, top: `${y}%`, width: `${width}%`, height: `${height}%`, background: PIE_COLORS[index % PIE_COLORS.length] }} onClick={() => onSelect(selectedCategory === item.label ? "" : item.label)} title={`${item.label}: ${moneyFull2.format(item.value)} · ${percent.toFixed(1)}% del gasto`} aria-label={`${item.label}: ${moneyFull2.format(item.value)}, ${percent.toFixed(1)}% del gasto`}><strong>{item.label}</strong><span>{percent.toFixed(1)}%</span><em>{moneyFull2.format(item.value)}</em></button>)}</div></section>;
}

function layoutTreemap(items: { label: string; value: number }[], total: number) {
  type Tile = { item: { label: string; value: number }; index: number; percent: number; x: number; y: number; width: number; height: number };
  const result: Tile[] = [];
  const place = (entries: { item: { label: string; value: number }; index: number }[], x: number, y: number, width: number, height: number, horizontal: boolean) => {
    if (!entries.length) return;
    if (entries.length === 1) { const entry = entries[0]; result.push({ ...entry, percent: total ? Math.abs(entry.item.value) / total * 100 : 0, x, y, width, height }); return; }
    const sum = entries.reduce((acc, entry) => acc + Math.abs(entry.item.value), 0);
    let running = Math.abs(entries[0].item.value); let split = 1;
    while (split < entries.length - 1 && running < sum / 2) { running += Math.abs(entries[split].item.value); split += 1; }
    const first = entries.slice(0, split); const second = entries.slice(split); const firstRatio = sum ? running / sum : .5;
    if (horizontal) { place(first, x, y, width * firstRatio, height, !horizontal); place(second, x + width * firstRatio, y, width * (1 - firstRatio), height, !horizontal); }
    else { place(first, x, y, width, height * firstRatio, !horizontal); place(second, x, y + height * firstRatio, width, height * (1 - firstRatio), !horizontal); }
  };
  place(items.map((item, index) => ({ item, index })), 0, 0, 100, 100, true);
  return result;
}

function ExpensePie({ title, subtitle, items, selectedCategory, onSelect }: { title: string; subtitle: string; items: { label: string; value: number }[]; selectedCategory: string; onSelect: (label: string) => void }) {
  return <ExpenseDonutCard eyebrow="COMPOSICIÓN DEL GASTO" title={title} subtitle={subtitle} items={items} selectedCategory={selectedCategory} onSelect={onSelect} />;
}

function ExpenseProportion({ fixed, variable, selectedCategory, onSelect }: { fixed: number; variable: number; selectedCategory: string; onSelect: (category: string) => void }) {
  return <ExpenseDonutCard eyebrow="COMPOSICIÓN DEL GASTO" title="Fijo vs. variable" subtitle="Participación sobre el gasto total." items={[{ label: "Gastos fijos", value: fixed }, { label: "Gastos variables", value: variable }]} selectedCategory={selectedCategory} onSelect={onSelect} />;
}

function ExpensePieCharts({ matrix, period, selectedCategory, onSelect }: { matrix: Matrix; period: number | "acc"; selectedCategory: string; onSelect: (category: string) => void }) {
  const fixed = useMemo(() => expenseBreakdown(matrix, "fixed", period), [matrix, period]);
  const variable = useMemo(() => expenseBreakdown(matrix, "variable", period), [matrix, period]);
  const fixedCategories = useMemo(() => expenseCategoryBreakdown(matrix, "fixed", period), [matrix, period]);
  const variableCategories = useMemo(() => expenseCategoryBreakdown(matrix, "variable", period), [matrix, period]);
  return <><ExpensePie title="Gastos fijos" subtitle="Incluye nómina administrativa, semanal y quincenal." items={fixed} selectedCategory={selectedCategory} onSelect={onSelect} /><ExpensePie title="Gastos variables" subtitle="Operación, ventas y conceptos ligados a actividad." items={variable} selectedCategory={selectedCategory} onSelect={onSelect} /><ExpensePie title="Gastos fijos por categoría" subtitle="Agrupación por gasto administrativo, operativo, de venta y financiero." items={fixedCategories} selectedCategory={selectedCategory} onSelect={onSelect} /><ExpensePie title="Gastos variables por categoría" subtitle="Agrupación por gasto administrativo, operativo, de venta y financiero." items={variableCategories} selectedCategory={selectedCategory} onSelect={onSelect} /></>;
}

function PnlSummaryLineChart({ matrix, metric, onSelect }: { matrix: Matrix; metric: string; onSelect: (period: number, metric: string) => void }) {
  const [granularity, setGranularity] = useState<"monthly" | "quarterly" | "annual">("monthly");
  const revenueRow = findRevenueRow(matrix);
  const costRow = findRow(matrix, /^COSTO$/);
  const marginRow = findRow(matrix, /MARGEN SOBRE MATERIA PRIMA/);
  const fixedRow = findRow(matrix, /GASTOS FIJOS/);
  const variableRow = findRow(matrix, /GASTOS VARIABLES/);
  const profitRow = findRow(matrix, /^UTILIDAD$/);
  const periods = availablePeriods(matrix);
  const periodGroups = granularity === "monthly" ? periods.map(([label, column]) => ({ label, columns: [column] })) : granularity === "quarterly" ? Array.from({ length: Math.ceil(periods.length / 3) }, (_, index) => ({ label: `T${index + 1}`, columns: periods.slice(index * 3, index * 3 + 3).map(([, column]) => column) })) : [{ label: "2026", columns: periods.map(([, column]) => column) }];
  const series = [
    { name: "Ventas", row: revenueRow, color: "#2584ff", axis: "primary" },
    { name: "Costo", row: costRow, color: "#88a1c2", axis: "primary" },
    { name: "Margen bruto", row: marginRow, color: "#ffb02e", axis: "secondary" },
    { name: "Gastos fijos", row: fixedRow, color: "#f2994a", axis: "secondary" },
    { name: "Gastos variables", row: variableRow, color: "#d84d61", axis: "secondary" },
    { name: "Gastos", row: -2, color: "#f2994a", axis: "secondary" },
    { name: "Rebates", row: findRow(matrix, /REBATES?/), color: "#7357e8", axis: "secondary" },
    { name: "Utilidad", row: profitRow, color: "#13d69f", axis: "secondary" },
    { name: "Impuestos", row: findRow(matrix, /IMPUESTOS?/), color: "#d84d61", axis: "secondary" },
    { name: "Margen neto", row: -3, color: "#ffb02e", axis: "secondary" },
  ];
  const activeSeries = metric === "all" ? series.filter((item) => ["Ventas", "Costo", "Utilidad", "Margen neto"].includes(item.name)) : [series.find((item) => item.name === metric) ?? series[0]];
  const valuesFor = (selected: typeof series[number]) => periodGroups.map(({ label, columns }) => { const sales = columns.reduce((sum, column) => sum + valueAt(matrix, revenueRow, column), 0); const value = selected.name === "Gastos" ? columns.reduce((sum, column) => sum + valueAt(matrix, fixedRow, column) + valueAt(matrix, variableRow, column), 0) : selected.name === "Rebates" ? columns.reduce((sum, column) => sum + rebatesAt(matrix, column), 0) : selected.name === "Margen neto" ? (sales ? columns.reduce((sum, column) => sum + valueAt(matrix, profitRow, column), 0) / sales * 100 : 0) : columns.reduce((sum, column) => sum + valueAt(matrix, selected.row, column), 0); return { label, column: columns.at(-1) ?? 2, value }; });
  const allValues = activeSeries.flatMap((selected) => valuesFor(selected).map((item) => item.value));
  const min = Math.min(0, ...allValues);
  const max = Math.max(1, ...allValues);
  const span = max - min || 1;
  const pointFor = (value: number, index: number) => {
    const x = periodGroups.length === 1 ? 50 : 6 + index / (periodGroups.length - 1) * 88;
    const y = 8 + (max - value) / span * 78;
    return { x, y: Math.max(4, Math.min(94, y)) };
  };
  const chartSeries = activeSeries.map((selected) => ({ selected, values: valuesFor(selected), points: valuesFor(selected).map((item, index) => pointFor(item.value, index)) }));
  const primaryValues = valuesFor(activeSeries[0]);

  return <section className="panel pnlLinePanel">
    <div className="panelHeading"><div><span className="eyebrow">EVOLUCIÓN DEL NEGOCIO</span><h3>{metric === "all" ? "Ventas, costo, utilidad y margen" : `Evolución de ${activeSeries[0].name.toLowerCase()}`}</h3><p>Selecciona un punto para actualizar el periodo financiero.</p></div><div className="panelHeadingActions"><div className="chartPeriodTabs" aria-label="Granularidad de la gráfica"><button type="button" className={granularity === "monthly" ? "active" : ""} onClick={() => setGranularity("monthly")}>Mensual</button><button type="button" className={granularity === "quarterly" ? "active" : ""} onClick={() => setGranularity("quarterly")}>Trimestral</button><button type="button" className={granularity === "annual" ? "active" : ""} onClick={() => setGranularity("annual")}>Anual</button></div><ExpandChartButton /></div></div><div className="lineLegend">{activeSeries.map((item) => <span key={item.name}><i style={{ background: item.color }} />{item.name}</span>)}</div>
    <div className="lineChart areaLineChart"><div className="lineScaleLabels"><span>{activeSeries[0].name === "Margen neto" ? `${max.toFixed(2)}%` : moneyFull2.format(max)}</span></div><svg viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><line className="lineGrid" x1="6" x2="94" y1="8" y2="8" /><line className="lineGrid" x1="6" x2="94" y1="47" y2="47" /><line className="lineGrid" x1="6" x2="94" y1="86" y2="86" />{min < 0 && <line className="lineZero" x1="6" x2="94" y1={pointFor(0, 0).y} y2={pointFor(0, 0).y} />}{chartSeries.map(({ selected, points }) => <g key={selected.name}><polygon points={`6,86 ${points.map((point) => `${point.x},${point.y}`).join(" ")} 94,86`} style={{ fill: selected.color }} /><polyline points={points.map((point) => `${point.x},${point.y}`).join(" ")} style={{ stroke: selected.color }} /></g>)}</svg><div className="areaPointLayer">{chartSeries.map(({ selected, values, points }) => points.map((point, index) => <button type="button" className="areaPoint" key={`${selected.name}-${values[index].label}`} style={{ left: `${point.x}%`, top: `${point.y}%`, borderColor: selected.color, background: selected.color }} onClick={() => onSelect(values[index].column, selected.name)} title={`${values[index].label}: ${selected.name === "Margen neto" ? `${values[index].value.toFixed(2)}%` : moneyFull2.format(values[index].value)}`}>{metric !== "all" && <span className="areaDataLabel">{selected.name === "Margen neto" ? `${values[index].value.toFixed(2)}%` : moneyFull2.format(values[index].value)}</span>}</button>))}</div><div className="lineChartLabels">{primaryValues.map(({ label }) => <span key={label}>{label.slice(0, 3)}</span>)}</div></div>
  </section>;
}

function PnlMonthDetail({ matrix, period }: { matrix: Matrix; period: number | "acc" }) {
  const periods = availablePeriods(matrix);
  const detailColumn = period === "acc" ? periods.at(-1)?.[1] ?? 2 : period;
  const detailIndex = Math.max(0, periods.findIndex(([, column]) => column === detailColumn));
  const previousColumn = periods[Math.max(0, detailIndex - 1)]?.[1] ?? detailColumn;
  const monthLabel = MONTHS.find(([, column]) => column === detailColumn)?.[0] ?? "Periodo";
  const revenueRow = findRevenueRow(matrix);
  const costRow = findRow(matrix, /^COSTO$/);
  const marginRow = findRow(matrix, /MARGEN SOBRE MATERIA PRIMA/);
  const fixedRow = findRow(matrix, /GASTOS FIJOS/);
  const variableRow = findRow(matrix, /GASTOS VARIABLES/);
  const profitRow = findRow(matrix, /^UTILIDAD$/);
  const amount = (row: number, column: number) => valueAt(matrix, row, column);
  const revenue = amount(revenueRow, detailColumn);
  const previousRevenue = amount(revenueRow, previousColumn);
  const profit = amount(profitRow, detailColumn);
  const previousProfit = amount(profitRow, previousColumn);
  const rows = [
    { label: "Ventas", value: revenue, previous: previousRevenue },
    { label: "Costo", value: amount(costRow, detailColumn), previous: amount(costRow, previousColumn), inverse: true },
    { label: "Margen bruto", value: amount(marginRow, detailColumn), previous: amount(marginRow, previousColumn) },
    { label: "Gastos", value: amount(fixedRow, detailColumn) + amount(variableRow, detailColumn), previous: amount(fixedRow, previousColumn) + amount(variableRow, previousColumn), inverse: true },
    { label: "Utilidad", value: profit, previous: previousProfit },
  ];
  const margin = revenue ? profit / revenue * 100 : 0;
  const previousMargin = previousRevenue ? previousProfit / previousRevenue * 100 : 0;
  const currentExpenses = [...expenseBreakdown(matrix, "fixed", detailColumn), ...expenseBreakdown(matrix, "variable", detailColumn)];
  const previousExpenses = new Map([...expenseBreakdown(matrix, "fixed", previousColumn), ...expenseBreakdown(matrix, "variable", previousColumn)].map((item) => [item.label, item.value]));
  const variations = currentExpenses.map((item) => ({ label: item.label.split(" · ").at(-1) ?? item.label, change: item.value - (previousExpenses.get(item.label) ?? 0) })).sort((a, b) => Math.abs(b.change) - Math.abs(a.change)).slice(0, 4);
  return <aside className="panel monthDetail"><div className="panelHeading"><div><span className="eyebrow">DETALLE DEL MES</span><h3>{monthLabel} 2026</h3></div></div><div className="monthDetailRows">{rows.map((item) => { const change = item.previous ? (item.value - item.previous) / Math.abs(item.previous) * 100 : 0; const favorable = item.inverse ? change <= 0 : change >= 0; return <div key={item.label}><span>{item.label}</span><b>{moneyFull2.format(item.value)}</b><em className={favorable ? "positive" : "negative"}>{change >= 0 ? "↑" : "↓"} {Math.abs(change).toFixed(1)}%</em></div>; })}<div><span>Margen neto</span><b>{margin.toFixed(2)}%</b><em className={margin >= previousMargin ? "positive" : "negative"}>{margin >= previousMargin ? "↑" : "↓"} {Math.abs(margin - previousMargin).toFixed(2)} pp</em></div></div><div className="monthVariations"><strong>Principales variaciones</strong>{variations.map((item) => <div key={item.label}><span title={item.label}>{item.label}</span><b className={item.change >= 0 ? "negative" : "positive"}>{item.change >= 0 ? "+" : "−"}{moneyFull2.format(Math.abs(item.change))}</b></div>)}</div></aside>;
}

function PnlSummaryBars({ revenue, cost, fixed, variable, rebates, tax, profit, profitPct, onSelect }: { revenue: number; cost: number; fixed: number; variable: number; rebates: number; tax: number; profit: number; profitPct: number; onSelect: (label: string) => void }) {
  const items: { label: string; value: number; color: string; barValue?: number; display?: string }[] = [
    { label: "VENTAS", value: revenue, color: "#2374e8" },
    { label: "COSTO", value: cost, color: "#8a98ad" },
    { label: "GASTOS", value: fixed + variable, color: "#f2994a" },
    { label: "REBATES", value: rebates, color: "#7357e8" },
    { label: "UTILIDAD", value: profit, color: profit < 0 ? "#d84d61" : "#12a67b" },
  ];
  const max = Math.max(...items.map((item) => item.barValue ?? Math.abs(item.value)), 1);
  const hasNegative = items.some((item) => item.value < 0);
  const [selectedLabel, setSelectedLabel] = useState("VENTAS");

  return <section className="panel summaryBarsPanel">
    <div className="panelHeading"><div><span className="eyebrow">RESUMEN FINANCIERO</span><h3>Comparativo del P&amp;L</h3><p>Importes completos del periodo seleccionado.</p></div><span className="waterfallPeriod">Periodo seleccionado</span></div>
    <div className="summaryBars" aria-label="Comparativo de ventas, costo, gastos, rebates, utilidad, impuestos y margen neto">{items.map((item) => <button type="button" className={`summaryBarRow ${selectedLabel === item.label ? "selected" : ""}`} key={item.label} onClick={() => { setSelectedLabel(item.label); onSelect(item.label); }}><strong>{item.label}</strong><span className={`summaryBarTrack ${hasNegative ? "diverging" : ""}`}><span className={item.value < 0 ? "negativeBar" : "positiveBar"} style={{ width: `${(item.barValue ?? Math.abs(item.value)) / max * (hasNegative ? 50 : 100)}%`, background: item.color }} /></span><b>{item.display ?? moneyFull2.format(item.value)}</b></button>)}</div>
  </section>;
}

function PnlView({ matrix, expenses, title, branchId, period, setPeriod }: { matrix?: Matrix; expenses?: Matrix; title: string; branchId: string; period: number | "acc"; setPeriod: (period: number | "acc") => void }) {
  const periods = useMemo(() => matrix ? availablePeriods(matrix) : [], [matrix]);
  const periodOptions = useMemo(() => {
    if (period === "acc" || periods.some(([, column]) => column === period)) return periods;
    const selected = MONTHS.find(([, column]) => column === period);
    return selected ? [...periods, selected].sort((a, b) => a[1] - b[1]) : periods;
  }, [periods, period]);
  const [transactionCategory, setTransactionCategory] = useState("");
  const [selectedMetric, setSelectedMetric] = useState("all");
  const transactionsRef = useRef<HTMLElement>(null);
  if (!matrix?.length) return <MissingData title={title} />;

  const revenueRow = findRevenueRow(matrix);
  const costRow = findRow(matrix, /^COSTO$/);
  const marginRow = findRow(matrix, /MARGEN SOBRE MATERIA PRIMA/);
  const fixedRow = findRow(matrix, /GASTOS FIJOS/);
  const variableRow = findRow(matrix, /GASTOS VARIABLES/);
  const taxRow = findRow(matrix, /IMPUESTOS?/);
  const profitRow = findRow(matrix, /^UTILIDAD$/);
  const revenue = valueAt(matrix, revenueRow, period);
  const cost = valueAt(matrix, costRow, period);
  const grossMargin = valueAt(matrix, marginRow, period);
  const fixed = valueAt(matrix, fixedRow, period);
  const variable = valueAt(matrix, variableRow, period);
  const tax = valueAt(matrix, taxRow, period);
  const rebates = rebatesAt(matrix, period);
  const profit = valueAt(matrix, profitRow, period);

  return <>
    <div className="viewHeader"><div><span className="eyebrow">ESTADO DE RESULTADOS</span><h2>{title}</h2><p>Resultados financieros y movimientos por fecha</p></div><label className="periodSelect"><span>PERIODO FINANCIERO</span><select value={period} onChange={(e) => setPeriod(e.target.value === "acc" ? "acc" : Number(e.target.value))}>{periodOptions.map(([label, column]) => <option value={column} key={column}>{label} 2026</option>)}<option value="acc">Acumulado 2026</option></select></label></div>
    <ExecutiveKpis matrix={matrix} period={period} selectedMetric={selectedMetric} onSelect={(metric) => { const nextMetric = selectedMetric === metric ? "all" : metric; setSelectedMetric(nextMetric); setTransactionCategory(nextMetric === "expenses" ? "" : transactionCategory); }} />
    <div className="pnlPrimaryGrid"><PnlSummaryLineChart matrix={matrix} metric={selectedMetric === "all" ? "all" : selectedMetric === "expenses" ? "Gastos" : selectedMetric === "rebates" ? "Rebates" : selectedMetric === "grossMargin" ? "Margen bruto" : selectedMetric === "netMargin" ? "Margen neto" : selectedMetric === "profit" ? "Utilidad" : selectedMetric === "cost" ? "Costo" : "Ventas"} onSelect={(nextPeriod, metricName) => { setPeriod(nextPeriod); setTransactionCategory(metricName === "Gastos fijos" ? "Gastos fijos" : metricName === "Gastos variables" ? "Gastos variables" : metricName === "Rebates" ? "REBATE" : ""); }} /></div>
    <div className="analysisHeader"><div><span className="eyebrow">ANÁLISIS DETALLADO</span><h2>Composición y movimientos del gasto</h2><p>Explora la estructura del egreso y consulta las transacciones que componen cada resultado.</p></div><button type="button" onClick={() => transactionsRef.current?.scrollIntoView({ behavior: "smooth", block: "start" })}>Ver movimientos ↓</button></div>
    <div className="expenseOverview"><ExpenseProportion fixed={fixed} variable={variable} selectedCategory={transactionCategory} onSelect={setTransactionCategory} /><ExpenseDistribution matrix={matrix} period={period} selectedCategory={transactionCategory} onSelect={setTransactionCategory} /><ExpensePieCharts matrix={matrix} period={period} selectedCategory={transactionCategory} onSelect={setTransactionCategory} /></div>
    <MovementPanel matrix={expenses} branchId={branchId} period={period} category={transactionCategory} panelRef={transactionsRef} onClearFilter={() => setTransactionCategory("")} />
  </>;
}

function matchesMovementCategory(row: Cell[], category?: string) {
  if (!category || category.startsWith("Otros conceptos")) return true;
  const normalizedCategory = normalize(category);
  if (/^REBATES?$/.test(normalizedCategory)) return /(CARTRO|GRUPAK)/.test(normalize(`${row[8] ?? ""} ${row[13] ?? ""} ${row[14] ?? ""}`));
  if (normalizedCategory === "SUELDO DIRECCION") return normalize(row[12]) === "SUELDO DIRECCION";
  if (normalizedCategory === "GASTOS FIJOS") return normalize(row[9]) === "F";
  if (normalizedCategory === "GASTOS VARIABLES") return normalize(row[9]) === "V";
  const selectedGroupCode = Object.entries(EXPENSE_GROUP_LABELS).find(([, label]) => normalize(label) === normalizedCategory)?.[0];
  if (selectedGroupCode) return normalize(row[12]) === selectedGroupCode;
  if (normalizedCategory.startsWith("AREA · ")) return normalize(row[12]) === normalizedCategory.slice(7).trim();
  const separator = category.indexOf(" · ");
  if (separator < 0) return true;
  const group = normalize(category.slice(0, separator));
  const groupCode = Object.entries(EXPENSE_GROUP_LABELS).find(([, label]) => normalize(label) === group)?.[0] ?? group;
  const name = normalize(category.slice(separator + 3));
  const rowGroup = normalize(row[12]);
  const rowName = normalize(row[13]);
  return rowGroup === groupCode && (rowName === name || rowName.includes(name) || name.includes(rowName));
}

function MovementPanel({ matrix, branchId, period, category, panelRef, onClearFilter }: { matrix?: Matrix; branchId: string; period: number | "acc"; category?: string; panelRef: React.RefObject<HTMLElement | null>; onClearFilter: () => void }) {
  const baseRows = useMemo(() => (matrix ?? []).slice(1).map((row, index) => ({
    index,
    row,
    date: isoDate(row[4]),
  })).filter((item) => item.date && belongsToBranch(item.row, branchId) && matchesMovementCategory(item.row, category)).sort((a, b) => a.date.localeCompare(b.date)), [matrix, branchId, category]);
  const firstDate = baseRows[0]?.date ?? "";
  const lastDate = baseRows.at(-1)?.date ?? "";
  const [sortKey, setSortKey] = useState<MovementSortKey>("date");
  const [sortDirection, setSortDirection] = useState<"asc" | "desc">("asc");
  const [isTableVisible, setIsTableVisible] = useState(true);
  const periodBounds = monthRange(period);
  const effectiveFrom = periodBounds?.from || firstDate;
  const effectiveTo = periodBounds?.to || lastDate;
  const movements = useMemo(() => baseRows.filter((item) => item.date >= effectiveFrom && item.date <= effectiveTo), [baseRows, effectiveFrom, effectiveTo]);
  const sortedMovements = useMemo(() => [...movements].sort((a, b) => {
    const textValue = (item: typeof movements[number]) => ({ date: item.date, provider: String(item.row[8] ?? ""), department: String(item.row[12] ?? ""), expense: String(item.row[13] ?? ""), concept: String(item.row[14] ?? "") })[sortKey as "date" | "provider" | "department" | "expense" | "concept"];
    const numericValue = (item: typeof movements[number]) => sortKey === "subtotal" ? numberOf(item.row[15]) : numberOf(item.row[17]);
    const left = ["subtotal", "total"].includes(sortKey) ? numericValue(a) : textValue(a).toLocaleUpperCase();
    const right = ["subtotal", "total"].includes(sortKey) ? numericValue(b) : textValue(b).toLocaleUpperCase();
    const comparison = typeof left === "number" && typeof right === "number" ? left - right : String(left).localeCompare(String(right), "es-MX");
    return sortDirection === "asc" ? comparison : -comparison;
  }), [movements, sortKey, sortDirection]);
  const changeSort = (nextKey: MovementSortKey) => {
    if (sortKey === nextKey) setSortDirection((current) => current === "asc" ? "desc" : "asc");
    else { setSortKey(nextKey); setSortDirection("asc"); }
  };
  const columns: { label: string; key: MovementSortKey }[] = [
    { label: "Fecha", key: "date" }, { label: "Proveedor", key: "provider" }, { label: "Departamento", key: "department" },
    { label: "Gasto", key: "expense" }, { label: "Concepto", key: "concept" }, { label: "Subtotal", key: "subtotal" }, { label: "Total", key: "total" },
  ];
  const renderHeader = ({ label, key }: { label: string; key: MovementSortKey }) => {
    const activeSort = sortKey === key;
    const arrow = activeSort && sortDirection === "desc" ? "▼" : "▲";
    return <th key={key} aria-sort={activeSort ? (sortDirection === "asc" ? "ascending" : "descending") : "none"}><button type="button" className="sortableHeader" onClick={() => changeSort(key)}>{label}<span aria-hidden="true">{arrow}</span></button></th>;
  };
  const subtotal = movements.reduce((sum, item) => sum + numberOf(item.row[15]), 0);
  const total = movements.reduce((sum, item) => sum + numberOf(item.row[17]), 0);

  return <section className="panel movementsPanel" ref={panelRef}>
    <div className="panelHeading movementHeading"><div><span className="eyebrow">FUENTE · GASTOS</span><h3>Movimientos por fecha</h3><p>Fechas vinculadas al periodo financiero seleccionado arriba.</p>{category && <div className="activeFilterBadge"><span>● Filtro activo</span><b>{category.replace(" · ", " / ")}</b><button type="button" onClick={onClearFilter}>× Limpiar</button></div>}</div><button type="button" className="tableToggle" onClick={() => setIsTableVisible((visible) => !visible)} aria-expanded={isTableVisible} aria-controls="movement-detail-table">{isTableVisible ? "Ocultar tabla" : "Mostrar tabla"}<span aria-hidden="true">{isTableVisible ? "⌃" : "⌄"}</span></button></div>
    {!baseRows.length ? <div className="noMovements">No hay movimientos de GASTOS asociados a esta sucursal.</div> : <><div className="movementSummary"><article><span>MOVIMIENTOS</span><strong>{movements.length.toLocaleString("es-MX")}</strong></article><article><span>SUBTOTAL</span><strong>{money.format(subtotal)}</strong></article><article><span>TOTAL CON IVA</span><strong>{money.format(total)}</strong></article></div>{isTableVisible && <div className="movementTableWrap" id="movement-detail-table"><table className="movementTable"><thead><tr>{columns.map(renderHeader)}</tr></thead><tbody>{sortedMovements.map((item) => { const amount = numberOf(item.row[17]); const rowType = amount < 0 ? "income" : amount > 0 ? "egress" : "neutral"; return <tr className={rowType} key={`${item.date}-${item.index}`}><td>{new Date(`${item.date}T12:00:00`).toLocaleDateString("es-MX")}</td><td>{String(item.row[8] ?? "—")}</td><td>{String(item.row[12] ?? "—")}</td><td>{String(item.row[13] ?? "—")}</td><td>{String(item.row[14] ?? "—")}</td><td>{moneyFull.format(numberOf(item.row[15]))}</td><td>{moneyFull.format(amount)}</td></tr>; })}</tbody></table></div>}</>}
  </section>;
}

function BalanceView({ matrix }: { matrix?: Matrix }) {
  const balanceTitlePattern = /^ESTADO DE (?:POSICION|SITUACION) FINANCIERA AL /i;
  const titleCells = useMemo(() => matrix ? matrix.flatMap((row, rowIndex) => row.map((cell, columnIndex) => ({ rowIndex, columnIndex, value: String(cell ?? "").trim() })).filter((cell) => balanceTitlePattern.test(cell.value))) : [], [matrix]);
  const blocks = titleCells.map((cell) => cell.columnIndex);
  const [offset, setOffset] = useState<number | null>(null);
  if (!matrix?.length) return <MissingData title="Balance general" />;
  const activeOffset = offset !== null && blocks.includes(offset) ? offset : blocks.at(-1) ?? 0;
  const titleCell = titleCells.find((cell) => cell.columnIndex === activeOffset);
  const title = (titleCell?.value ?? "Balance general").replace(/^ESTADO DE (?:POSICION|SITUACION) FINANCIERA AL /i, "Al ");
  const left = matrix.map((row, index) => ({ index, name: String(row?.[activeOffset + 1] ?? "").trim(), value: numberOf(row?.[activeOffset + 2]) })).filter((item) => item.name);
  const right = matrix.map((row, index) => ({ index, name: String(row?.[activeOffset + 4] ?? "").trim(), value: numberOf(row?.[activeOffset + 5]) })).filter((item) => item.name);
  const named = (items: typeof left, pattern: RegExp) => items.find((item) => pattern.test(normalize(item.name)))?.value ?? 0;
  const assets = named(left, /^TOTAL ACTIVO$/);
  const totalRight = named(right, /TOTAL PASIVO \+ CAPITAL/);
  const accountLeft = left.filter((item) => item.index >= 4 && !/^(A C|A F|TOTAL ACTIVO)$/.test(normalize(item.name)));
  const accountRight = right.filter((item) => item.index >= 4 && !/^(P|C|TOTAL PASIVO \+ CAPITAL)$/.test(normalize(item.name)));
  const equityRows = accountRight.filter((item) => /CAPITAL|RESULTADO|UTILIDAD|PERDIDA|PATRIMONIO/.test(normalize(item.name)));
  const liabilityRows = accountRight.filter((item) => !equityRows.some((equity) => equity.index === item.index));
  const liabilities = liabilityRows.reduce((sum, item) => sum + item.value, 0);
  const equity = equityRows.reduce((sum, item) => sum + item.value, 0);
  const difference = assets - (liabilities + equity);
  const assetRows = accountLeft.filter((item) => item.index >= 4 && item.index < 38 && !/^(A C|A F|TOTAL ACTIVO)$/.test(normalize(item.name)));
  const currentMarker = left.find((item) => /^A C$/.test(normalize(item.name)))?.index ?? 0;
  const fixedMarker = left.find((item) => /^A F$/.test(normalize(item.name)))?.index ?? 38;
  const currentAssetRows = assetRows.filter((item) => item.index > currentMarker && item.index < fixedMarker);
  const fixedAssetRows = assetRows.filter((item) => item.index > fixedMarker);
  const ratio = (value: number, base: number) => base ? value / base * 100 : 0;
  const balanced = Math.abs(difference) < 0.01;

  return <>
    <div className="viewHeader"><div><span className="eyebrow">POSICIÓN FINANCIERA</span><h2>Balance general</h2><p>{title}</p></div><label className="periodSelect"><span>PERIODO</span><select value={activeOffset} onChange={(e) => setOffset(Number(e.target.value))}>{blocks.map((block) => <option value={block} key={block}>{String(matrix[1]?.[block]).replace(/^.* AL /, "")}</option>)}</select></label></div>
    <div className="balanceKpis"><BalanceKpi label="ACTIVOS" value={assets} detail="100% del balance" tone="blue" /><BalanceKpi label="PASIVOS" value={liabilities} detail={`${ratio(liabilities, assets).toFixed(1)}% de los activos`} tone="coral" /><BalanceKpi label="CAPITAL" value={equity} detail={`${ratio(equity, assets).toFixed(1)}% de los activos`} tone="purple" /><BalanceKpi label="BALANCE CHECK" value={difference} detail={balanced ? "Activos = Pasivo + Capital" : "Diferencia por conciliar"} tone={balanced ? "green" : "coral"} check={balanced} /></div>
    <div className="balanceDetails"><BalanceAccounts title="Activos circulantes" total={currentAssetRows.reduce((sum, item) => sum + item.value, 0)} rows={currentAssetRows} tone="asset" selectedName="" /><BalanceAccounts title="Activos fijos" total={fixedAssetRows.reduce((sum, item) => sum + item.value, 0)} rows={fixedAssetRows} tone="asset" selectedName="" /><BalanceAccounts title="Pasivos" total={liabilities} rows={liabilityRows} tone="liability" selectedName="" /><BalanceAccounts title="Capital" total={equity} rows={equityRows} tone="equity" selectedName="" /></div>
  </>;
}

function BalanceKpi({ label, value, detail, tone, check = false }: { label: string; value: number; detail: string; tone: string; check?: boolean }) {
  return <article className={`panel balanceKpi ${tone}`}><span className="metricLabel">{label}</span><strong>{check ? "CUADRADO" : moneyFull2.format(value)}</strong><small>{detail}</small></article>;
}

function BalanceAssetComposition({ items, total, selectedAsset, onSelect }: { items: { name: string; value: number }[]; total: number; selectedAsset: string; onSelect: (name: string) => void }) {
  let cursor = 0;
  const segments = items.map((item, index) => { const percent = total ? Math.abs(item.value) / Math.abs(total) * 100 : 0; const start = cursor; cursor += percent; return { ...item, percent, start, color: PIE_COLORS[index % PIE_COLORS.length] }; });
  const selected = segments.find((item) => item.name === selectedAsset) ?? segments[0];
  return <section className="panel balanceComposition"><div className="panelHeading"><div><span className="eyebrow">COMPOSICIÓN</span><h3>Composición de activos</h3><p>Distribución por rubro principal</p></div><strong>{moneyCompact.format(total)}</strong></div><div className="balanceCompositionBody"><div className="balanceDonut"><svg viewBox="0 0 120 120">{segments.map((item) => <circle key={item.name} className={selectedAsset && selectedAsset !== item.name ? "muted" : ""} role="button" tabIndex={0} aria-label={`${item.name}: ${item.percent.toFixed(1)}%`} onClick={() => onSelect(item.name)} onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") onSelect(item.name); }} cx="60" cy="60" r="42" pathLength="100" fill="none" stroke={item.color} strokeWidth="22" strokeDasharray={`${item.percent} ${100 - item.percent}`} strokeDashoffset={-item.start} />)}</svg><button type="button" className="balanceDonutCenter" onClick={() => onSelect("")}><b>{selected?.percent.toFixed(1) ?? "0.0"}%</b><span title={selected?.name}>{selected?.name ?? "ACTIVOS"}</span></button></div><div className="balanceLegend"><div><span>CONCEPTO</span><span>%</span><span>MONTO</span></div>{segments.map((item) => <button type="button" className={selectedAsset === item.name ? "active" : ""} key={item.name} onClick={() => onSelect(item.name)}><span title={item.name}><i style={{ background: item.color }} />{item.name}</span><em>{item.percent.toFixed(1)}%</em><b>{moneyFull2.format(item.value)}</b></button>)}</div></div></section>;
}

function BalanceStructure({ liabilities, equity, assets, difference }: { liabilities: number; equity: number; assets: number; difference: number }) {
  const total = Math.abs(liabilities) + Math.abs(equity);
  const liabilityPercent = ratioValue(liabilities, total);
  return <section className="panel balanceStructure"><div className="panelHeading"><div><span className="eyebrow">ESTRUCTURA</span><h3>Estructura financiera</h3><p>Composición de Pasivo y Capital</p></div></div><div className="structureBar"><span style={{ width: `${liabilityPercent}%` }} /><span style={{ width: `${100 - liabilityPercent}%` }} /></div><div className="structureLabels"><span><i className="coralDot" />PASIVO <b>{liabilityPercent.toFixed(1)}%</b></span><span><i className="purpleDot" />CAPITAL <b>{(100 - liabilityPercent).toFixed(1)}%</b></span></div><div className="structureTotals"><div><span>Pasivo</span><b>{moneyFull2.format(liabilities)}</b></div><div><span>Capital</span><b>{moneyFull2.format(equity)}</b></div><hr /><div><span>Pasivo + Capital</span><b>{moneyFull2.format(liabilities + equity)}</b></div><div><span>vs. Activos</span><b>{assets ? `${((liabilities + equity) / assets * 100).toFixed(1)}%` : "0.0%"}</b></div><div className={Math.abs(difference) < .01 ? "balancedText" : "unbalancedText"}><span>Diferencia</span><b>{moneyFull2.format(difference)}</b></div></div></section>;
}

function ratioValue(value: number, base: number) { return base ? Math.abs(value) / Math.abs(base) * 100 : 0; }

function BalanceRatios({ assets, liabilities, assetRows }: { assets: number; liabilities: number; assetRows: { name: string; value: number }[] }) {
  const customer = assetRows.filter((item) => /CLIENTE|CUENTA POR COBRAR|DEUDOR/.test(normalize(item.name))).reduce((sum, item) => sum + item.value, 0);
  const inventory = assetRows.filter((item) => /INVENT|TARIMA/.test(normalize(item.name))).reduce((sum, item) => sum + item.value, 0);
  return <section className="panel balanceRatios"><div className="panelHeading"><div><span className="eyebrow">INDICADORES</span><h3>Indicadores financieros</h3><p>Relaciones clave del balance</p></div></div><div className="ratioGrid"><div><span>Pasivo / Activo</span><b>{ratioValue(liabilities, assets).toFixed(1)}%</b></div><div><span>Clientes / Activo</span><b>{ratioValue(customer, assets).toFixed(1)}%</b></div><div><span>Inventario / Activo</span><b>{ratioValue(inventory, assets).toFixed(1)}%</b></div><div><span>Activos identificados</span><b>{assetRows.length}</b></div></div></section>;
}

function BalanceAccounts({ title, total, rows, tone, selectedName }: { title: string; total: number; rows: { index: number; name: string; value: number }[]; tone: string; selectedName: string }) {
  const [sort, setSort] = useState<"default" | "desc" | "asc">("default");
  const sortedRows = sort === "default" ? rows : [...rows].sort((a, b) => sort === "desc" ? Math.abs(b.value) - Math.abs(a.value) : Math.abs(a.value) - Math.abs(b.value));
  const nextSort = sort === "default" ? "desc" : sort === "desc" ? "asc" : "default";
  const sortLabel = sort === "default" ? "Orden original" : sort === "desc" ? "Mayor a menor" : "Menor a mayor";
  return <section className={`panel balanceList ${tone}`}><div className="panelHeading"><div><span className="eyebrow">DETALLE DEL BALANCE</span><h3>{title}</h3></div><div className="balanceListActions"><button type="button" onClick={() => setSort(nextSort)} aria-label={`Ordenar ${title}`}>↕ {sortLabel}</button><strong>{money.format(total)}</strong></div></div><div className="balanceTableHead"><span>CONCEPTO</span><span>MONTO</span><span>%</span></div>{sortedRows.map((row) => <div className={`accountRow ${selectedName === row.name ? "selected" : ""}`} key={`${row.index}-${row.name}`}><span title={row.name}>{row.name}</span><b>{moneyFull.format(row.value)}</b><em>{total ? (Math.abs(row.value) / Math.abs(total) * 100).toFixed(1) : "0.0"}%</em><i><span style={{ width: `${Math.min(100, total ? Math.abs(row.value) / Math.abs(total) * 100 : 0)}%` }} /></i></div>)}</section>;
}

function MissingData({ title }: { title: string }) {
  return <section className="empty"><div className="emptyIcon">↗</div><h2>{title}</h2><p>No encontramos la hoja correspondiente en este archivo. Carga el formato financiero de KIT para visualizarla.</p></section>;
}

function ConversionIcon() {
  return <span className="conversionIcon" aria-hidden="true"><svg viewBox="0 0 32 32" role="presentation"><circle cx="16" cy="16" r="16" /><g><path d="M7.6 12.1 15.8 7.4l3.2 1.9-8.3 4.9z" /><path d="m12.2 15 8.2-4.9 4.1 2.5-8.4 5z" /><path d="M7.2 14.1 15 18.7v8.1l-7.8-4.7z" /><path d="m16.8 18.7 8-4.7v8.1l-8 4.8z" /></g></svg></span>;
}

export default function KitDashboard() {
  const [authenticated, setAuthenticated] = useState(false);
  const [authChecked, setAuthChecked] = useState(false);
  const [sheets, setSheets] = useState<Sheets>({});
  const [active, setActive] = useState("general");
  const [generalView, setGeneralView] = useState<"pnl" | "balance">("pnl");
  const [selectedPeriod, setSelectedPeriod] = useState<number | "acc">("acc");
  const [fileName, setFileName] = useState("datos-kit.xlsx");
  const [error, setError] = useState("");
  const [lastUpdated, setLastUpdated] = useState("");

  useEffect(() => { setAuthenticated(window.localStorage.getItem(AUTH_STORAGE_KEY) === "1"); setAuthChecked(true); }, []);

  useEffect(() => { (async () => { try { const storedWorkbook = await loadWorkbook().catch(() => null); const stored = window.localStorage.getItem(STORED_WORKBOOK_KEY); const storedName = window.localStorage.getItem(STORED_WORKBOOK_NAME_KEY); const data = storedWorkbook?.data ?? (stored ? base64ToBytes(stored).buffer : await fetch(DEFAULT_WORKBOOK_PATH).then((response) => { if (!response.ok) throw new Error(); return response.arrayBuffer(); })); setSheets(workbookToSheets(XLSX.read(data, { type: "array", cellDates: true }))); if (storedWorkbook?.name ?? storedName) setFileName(storedWorkbook?.name ?? storedName ?? "datos-kit.xlsx"); setLastUpdated(new Date().toLocaleString("es-MX")); } catch { setError("No se pudo cargar el archivo financiero inicial."); } })(); }, []);

  async function onFile(file?: File) {
    if (!file) return;
    try { setError(""); const data = await file.arrayBuffer(); const workbook = XLSX.read(data, { type: "array", cellDates: true }); await saveWorkbook(data, file.name).catch(() => undefined); try { window.localStorage.setItem(STORED_WORKBOOK_KEY, bytesToBase64(new Uint8Array(data))); window.localStorage.setItem(STORED_WORKBOOK_NAME_KEY, file.name); } catch { /* IndexedDB conserva el archivo completo cuando está disponible. */ } setSheets(workbookToSheets(workbook)); setFileName(file.name); setLastUpdated(new Date().toLocaleString("es-MX")); setActive("general"); }
    catch { setError("No pudimos leer el archivo. Verifica que sea un Excel válido."); }
  }

  if (!authChecked) return null;
  if (!authenticated) return <LoginScreen onLogin={() => setAuthenticated(true)} />;
  const branch = BRANCHES.find((item) => item.id === active);
  return <main className="appShell">
    <header className="topbar"><div className="brand"><div className="brandMark">K</div><div><strong>KIT</strong></div></div><nav className="branchTabs" aria-label="Sucursales"><button className={active === "general" ? "active" : ""} onClick={() => setActive("general")}><span>GRL</span>General</button>{BRANCHES.map((item) => <button key={item.id} className={`${active === item.id ? "active" : ""} ${item.id === "express" ? "expressTab" : ""}`} onClick={() => setActive(item.id)}>{item.id === "express" ? <span className="expressLogo" aria-hidden="true"><b>express</b></span> : item.id === "machinery" ? <span className="machineryIcon" aria-hidden="true"><Image src="/machinery-icon.png" alt="" width={32} height={32} /></span> : item.id === "conversion" ? <ConversionIcon /> : STATE_BRANCHES.has(item.id) ? <span className={`stateIcon state-${item.id}`} aria-hidden="true" /> : <span>{item.short}</span>}{item.label}</button>)}</nav><span className="brandTagline">Impulsando soluciones en cartón</span><div className="topActions"><div className="fileStatus"><i /> <span>{fileName}</span></div><label className="uploadButton"><input type="file" accept=".xlsx,.xls" onChange={(e) => onFile(e.target.files?.[0])} />Actualizar Excel <span>↑</span></label></div></header>
    <div className="workspace"><div className="workspaceToolbar"><div>{active === "general" && <div className="subTabs"><button className={generalView === "pnl" ? "active" : ""} onClick={() => setGeneralView("pnl")}>P&amp;L consolidado</button><button className={generalView === "balance" ? "active" : ""} onClick={() => setGeneralView("balance")}>Balance general</button></div>}</div><DownloadPdfButton /></div>{error && <div className="error">{error}</div>}{active === "general" ? <>{generalView === "pnl" ? <PnlView key="general" matrix={sheets["P&L"]} expenses={sheets.GASTOS} title="Estado de resultados" branchId="general" period={selectedPeriod} setPeriod={setSelectedPeriod} /> : <BalanceView matrix={sheets.BALANCE} />}</> : <PnlView key={active} matrix={sheets[branch?.sheet ?? ""]} expenses={sheets.GASTOS} title={branch?.label ?? "Sucursal"} branchId={active} period={selectedPeriod} setPeriod={setSelectedPeriod} />}</div>
    <footer className="appFooter"><span>Última actualización: {lastUpdated || "Cargando…"}</span><span>KIT | Dashboard P&amp;L · v2.0</span></footer>
  </main>;
}
