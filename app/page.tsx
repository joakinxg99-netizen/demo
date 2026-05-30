"use client";

import { ChangeEvent, DragEvent, useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

type CellValue = string | number | null;
type DataRow = Record<string, CellValue>;
type FilterState = Record<string, { query: string; min: string; max: string }>;
type SortState = { column: string; direction: "asc" | "desc" } | null;

const SAMPLE_ROWS: DataRow[] = [
  { country: "Japan", region: "East Asia", year: 2024, population_m: 123.2, median_age: 49.1, fertility_rate: 1.2, urban_pct: 92, life_expectancy: 84.8 },
  { country: "Nigeria", region: "West Africa", year: 2024, population_m: 229.2, median_age: 17.2, fertility_rate: 5.0, urban_pct: 54, life_expectancy: 54.6 },
  { country: "Brazil", region: "South America", year: 2024, population_m: 217.6, median_age: 34.6, fertility_rate: 1.6, urban_pct: 87, life_expectancy: 76.4 },
  { country: "Germany", region: "Western Europe", year: 2024, population_m: 83.3, median_age: 45.7, fertility_rate: 1.5, urban_pct: 78, life_expectancy: 81.4 },
  { country: "India", region: "South Asia", year: 2024, population_m: 1450.9, median_age: 28.8, fertility_rate: 2.0, urban_pct: 37, life_expectancy: 72.2 },
  { country: "Mexico", region: "North America", year: 2024, population_m: 129.7, median_age: 30.6, fertility_rate: 1.8, urban_pct: 82, life_expectancy: 75.1 },
  { country: "Egypt", region: "North Africa", year: 2024, population_m: 114.5, median_age: 24.2, fertility_rate: 2.8, urban_pct: 43, life_expectancy: 71.6 },
  { country: "Canada", region: "North America", year: 2024, population_m: 39.7, median_age: 41.8, fertility_rate: 1.4, urban_pct: 82, life_expectancy: 82.9 },
];

const compactNumber = new Intl.NumberFormat("en", {
  maximumFractionDigits: 2,
  notation: "compact",
});

function cleanCell(value: unknown): CellValue {
  if (value === undefined || value === null || value === "") {
    return null;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const text = String(value).trim();
  const normalized = text.replace(/,/g, "");
  const numeric = Number(normalized);

  if (text !== "" && Number.isFinite(numeric) && /^-?\d*\.?\d+(e[-+]?\d+)?$/i.test(normalized)) {
    return numeric;
  }

  return text;
}

function normalizeRows(rawRows: Record<string, unknown>[]): DataRow[] {
  return rawRows
    .map((row) =>
      Object.fromEntries(
        Object.entries(row)
          .filter(([key]) => key.trim())
          .map(([key, value]) => [key.trim(), cleanCell(value)]),
      ),
    )
    .filter((row) => Object.values(row).some((value) => value !== null));
}

function getColumns(rows: DataRow[]) {
  return Array.from(new Set(rows.flatMap((row) => Object.keys(row))));
}

function isNumericColumn(rows: DataRow[], column: string) {
  const values = rows.map((row) => row[column]).filter((value): value is number => typeof value === "number");
  return values.length > 0 && values.length >= Math.max(2, rows.length * 0.35);
}

function getNumericValues(rows: DataRow[], column: string) {
  return rows.map((row) => row[column]).filter((value): value is number => typeof value === "number" && Number.isFinite(value));
}

function mean(values: number[]) {
  return values.reduce((total, value) => total + value, 0) / values.length;
}

function pearson(rows: DataRow[], xColumn: string, yColumn: string) {
  const pairs = rows
    .map((row) => [row[xColumn], row[yColumn]])
    .filter((pair): pair is [number, number] => typeof pair[0] === "number" && typeof pair[1] === "number");

  if (pairs.length < 3) {
    return null;
  }

  const xs = pairs.map(([x]) => x);
  const ys = pairs.map(([, y]) => y);
  const xMean = mean(xs);
  const yMean = mean(ys);
  let numerator = 0;
  let xVariance = 0;
  let yVariance = 0;

  pairs.forEach(([x, y]) => {
    const dx = x - xMean;
    const dy = y - yMean;
    numerator += dx * dy;
    xVariance += dx ** 2;
    yVariance += dy ** 2;
  });

  const denominator = Math.sqrt(xVariance * yVariance);
  return denominator === 0 ? null : numerator / denominator;
}

function formatValue(value: CellValue) {
  if (typeof value === "number") {
    return Number.isInteger(value) ? value.toLocaleString() : value.toLocaleString(undefined, { maximumFractionDigits: 3 });
  }

  return value ?? "—";
}

function toCsv(rows: DataRow[], columns: string[]) {
  const escape = (value: CellValue) => {
    const text = value === null ? "" : String(value);
    return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  };

  return [columns.map(escape).join(","), ...rows.map((row) => columns.map((column) => escape(row[column] ?? null)).join(","))].join("\n");
}

function GlassPanel({ children, className = "" }: { children: React.ReactNode; className?: string }) {
  return <section className={`glass-panel ${className}`}>{children}</section>;
}

function MiniBars({ values }: { values: number[] }) {
  const sample = values.slice(0, 32);
  const min = Math.min(...sample);
  const max = Math.max(...sample);
  const span = max - min || 1;

  return (
    <div className="flex h-16 items-end gap-1">
      {sample.map((value, index) => (
        <div
          className="min-w-1 flex-1 rounded-t bg-cyan-300/75 shadow-[0_0_18px_rgba(34,211,238,0.45)]"
          key={`${value}-${index}`}
          style={{ height: `${18 + ((value - min) / span) * 82}%` }}
          title={formatValue(value)}
        />
      ))}
    </div>
  );
}

function ScatterPlot({ rows, xColumn, yColumn }: { rows: DataRow[]; xColumn: string; yColumn: string }) {
  const pairs = rows
    .map((row) => [row[xColumn], row[yColumn]])
    .filter((pair): pair is [number, number] => typeof pair[0] === "number" && typeof pair[1] === "number");

  if (pairs.length < 2) {
    return <div className="empty-state">Choose two numeric fields to render a scatter field.</div>;
  }

  const xs = pairs.map(([x]) => x);
  const ys = pairs.map(([, y]) => y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  const xSpan = maxX - minX || 1;
  const ySpan = maxY - minY || 1;

  return (
    <svg className="h-[320px] w-full overflow-visible" role="img" viewBox="0 0 640 320">
      <defs>
        <linearGradient id="plotGlow" x1="0" x2="1" y1="0" y2="1">
          <stop stopColor="#22d3ee" />
          <stop offset="1" stopColor="#a78bfa" />
        </linearGradient>
      </defs>
      <rect className="fill-white/[0.03]" height="272" rx="14" width="584" x="40" y="20" />
      {[0, 1, 2, 3, 4].map((tick) => (
        <g key={tick}>
          <line className="stroke-white/10" x1="40" x2="624" y1={20 + tick * 68} y2={20 + tick * 68} />
          <line className="stroke-white/10" x1={40 + tick * 146} x2={40 + tick * 146} y1="20" y2="292" />
        </g>
      ))}
      {pairs.slice(0, 450).map(([x, y], index) => (
        <circle
          className="fill-cyan-200/75"
          cx={40 + ((x - minX) / xSpan) * 584}
          cy={292 - ((y - minY) / ySpan) * 272}
          key={`${x}-${y}-${index}`}
          r="4"
        />
      ))}
      <text className="fill-slate-300 text-[13px]" x="40" y="314">
        {xColumn}
      </text>
      <text className="fill-slate-300 text-[13px]" x="12" y="28">
        {yColumn}
      </text>
      <path className="stroke-[url(#plotGlow)] stroke-2" d="M40 292 L624 20" opacity="0.15" />
    </svg>
  );
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [rows, setRows] = useState<DataRow[]>([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [filters, setFilters] = useState<FilterState>({});
  const [sort, setSort] = useState<SortState>(null);
  const [xColumn, setXColumn] = useState("");
  const [yColumn, setYColumn] = useState("");
  const [dragging, setDragging] = useState(false);

  const columns = useMemo(() => getColumns(rows), [rows]);
  const numericColumns = useMemo(() => columns.filter((column) => isNumericColumn(rows, column)), [columns, rows]);

  const filteredRows = useMemo(() => {
    const nextRows = rows.filter((row) =>
      columns.every((column) => {
        const filter = filters[column];
        if (!filter) {
          return true;
        }

        const value = row[column];
        const textMatch =
          !filter.query || String(value ?? "").toLowerCase().includes(filter.query.trim().toLowerCase());
        const numericValue = typeof value === "number" ? value : null;
        const minMatch = !filter.min || (numericValue !== null && numericValue >= Number(filter.min));
        const maxMatch = !filter.max || (numericValue !== null && numericValue <= Number(filter.max));

        return textMatch && minMatch && maxMatch;
      }),
    );

    if (!sort) {
      return nextRows;
    }

    return [...nextRows].sort((a, b) => {
      const first = a[sort.column];
      const second = b[sort.column];
      const firstValue = typeof first === "number" ? first : String(first ?? "");
      const secondValue = typeof second === "number" ? second : String(second ?? "");
      const comparison = firstValue > secondValue ? 1 : firstValue < secondValue ? -1 : 0;
      return sort.direction === "asc" ? comparison : -comparison;
    });
  }, [columns, filters, rows, sort]);

  const correlations = useMemo(() => {
    return numericColumns
      .flatMap((left, leftIndex) =>
        numericColumns.slice(leftIndex + 1).map((right) => ({
          left,
          right,
          value: pearson(filteredRows, left, right),
        })),
      )
      .filter((item): item is { left: string; right: string; value: number } => item.value !== null)
      .sort((a, b) => Math.abs(b.value) - Math.abs(a.value));
  }, [filteredRows, numericColumns]);

  const primaryStats = useMemo(() => {
    return numericColumns.slice(0, 4).map((column) => {
      const values = getNumericValues(filteredRows, column);
      return {
        column,
        average: values.length ? mean(values) : 0,
        max: values.length ? Math.max(...values) : 0,
        min: values.length ? Math.min(...values) : 0,
        values,
      };
    });
  }, [filteredRows, numericColumns]);

  async function loadFile(file: File) {
    setError("");
    setFileName(file.name);

    try {
      if (file.name.toLowerCase().endsWith(".csv")) {
        const text = await file.text();
        const parsed = Papa.parse<Record<string, unknown>>(text, {
          header: true,
          skipEmptyLines: true,
          dynamicTyping: false,
        });

        if (parsed.errors.length) {
          throw new Error(parsed.errors[0].message);
        }

        setRows(normalizeRows(parsed.data));
      } else {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { defval: null });
        setRows(normalizeRows(json));
      }

      setFilters({});
      setSort(null);
    } catch (caught) {
      setRows([]);
      setError(caught instanceof Error ? caught.message : "Could not parse that file.");
    }
  }

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) {
      void loadFile(file);
    }
  }

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files[0];
    if (file) {
      void loadFile(file);
    }
  }

  function loadSample() {
    setRows(SAMPLE_ROWS);
    setFileName("sample-demography.csv");
    setFilters({});
    setSort(null);
    setError("");
  }

  function updateFilter(column: string, key: "query" | "min" | "max", value: string) {
    setFilters((current) => ({
      ...current,
      [column]: {
        ...(current[column] ?? { query: "", min: "", max: "" }),
        [key]: value,
      },
    }));
  }

  function exportCsv() {
    const blob = new Blob([toCsv(filteredRows, columns)], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `demography-export-${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  const selectedX = xColumn || numericColumns[0] || "";
  const selectedY = yColumn || numericColumns[1] || numericColumns[0] || "";

  return (
    <main className="relative min-h-screen overflow-hidden bg-[#050816] px-4 py-8 text-slate-100 sm:px-6 lg:px-10">
      <div className="cosmic-grid" />
      <div className="star-field" />
      <div className="mx-auto flex w-full max-w-7xl flex-col gap-6">
        <header className="flex flex-col items-center gap-5 py-8 text-center">
          <div className="rounded-full border border-cyan-300/25 bg-cyan-300/10 px-4 py-2 text-xs font-semibold uppercase tracking-[0.28em] text-cyan-200 shadow-[0_0_28px_rgba(34,211,238,0.18)]">
            Astro Demography Lab
          </div>
          <div className="max-w-4xl">
            <h1 className="text-4xl font-semibold tracking-normal text-white sm:text-6xl">Demography Explorer</h1>
            <p className="mt-4 text-base leading-7 text-slate-300 sm:text-lg">
              Upload population, migration, age, fertility, income, or region data and explore patterns in a luminous client-side workspace.
            </p>
          </div>
        </header>

        <GlassPanel className="mx-auto w-full max-w-4xl p-4 sm:p-6">
          <div
            className={`flex min-h-64 flex-col items-center justify-center rounded-lg border border-dashed p-8 text-center transition ${
              dragging ? "border-cyan-200 bg-cyan-300/15" : "border-violet-200/30 bg-white/[0.04]"
            }`}
            onDragLeave={() => setDragging(false)}
            onDragOver={(event) => {
              event.preventDefault();
              setDragging(true);
            }}
            onDrop={handleDrop}
          >
            <input accept=".csv,.xlsx,.xls" className="hidden" onChange={handleFileChange} ref={inputRef} type="file" />
            <div className="mb-5 flex h-16 w-16 items-center justify-center rounded-full border border-cyan-300/40 bg-cyan-300/10 text-2xl text-cyan-200 shadow-[0_0_45px_rgba(34,211,238,0.28)]">
              ↑
            </div>
            <h2 className="text-2xl font-semibold text-white">Drop a CSV or Excel file</h2>
            <p className="mt-2 max-w-xl text-sm leading-6 text-slate-300">
              Data stays in your browser. The first sheet is loaded for Excel workbooks, and numeric-looking values are detected automatically.
            </p>
            <div className="mt-6 flex flex-wrap items-center justify-center gap-3">
              <button className="neon-button" onClick={() => inputRef.current?.click()} type="button">
                Select file
              </button>
              <button className="quiet-button" onClick={loadSample} type="button">
                Load sample
              </button>
            </div>
            {fileName && <p className="mt-4 text-sm text-cyan-100">Loaded: {fileName}</p>}
            {error && <p className="mt-4 text-sm text-rose-200">{error}</p>}
          </div>
        </GlassPanel>

        {rows.length > 0 && (
          <>
            <div className="grid gap-4 md:grid-cols-4">
              <GlassPanel className="p-5">
                <p className="text-sm text-slate-400">Rows</p>
                <p className="mt-2 text-3xl font-semibold text-white">{filteredRows.length.toLocaleString()}</p>
                <p className="mt-1 text-xs text-slate-500">of {rows.length.toLocaleString()} loaded</p>
              </GlassPanel>
              <GlassPanel className="p-5">
                <p className="text-sm text-slate-400">Columns</p>
                <p className="mt-2 text-3xl font-semibold text-white">{columns.length}</p>
                <p className="mt-1 text-xs text-slate-500">{numericColumns.length} numeric fields</p>
              </GlassPanel>
              <GlassPanel className="p-5">
                <p className="text-sm text-slate-400">Strongest signal</p>
                <p className="mt-2 text-3xl font-semibold text-white">{correlations[0] ? correlations[0].value.toFixed(2) : "—"}</p>
                <p className="mt-1 truncate text-xs text-slate-500">{correlations[0] ? `${correlations[0].left} × ${correlations[0].right}` : "Need two numeric fields"}</p>
              </GlassPanel>
              <GlassPanel className="p-5">
                <p className="text-sm text-slate-400">Export</p>
                <button className="mt-3 w-full rounded-md border border-violet-300/35 bg-violet-400/10 px-4 py-3 text-sm font-semibold text-violet-100 transition hover:bg-violet-400/20" onClick={exportCsv} type="button">
                  Download filtered CSV
                </button>
              </GlassPanel>
            </div>

            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
              {primaryStats.map((stat) => (
                <GlassPanel className="p-5" key={stat.column}>
                  <div className="flex items-start justify-between gap-3">
                    <div>
                      <p className="text-sm text-slate-400">{stat.column}</p>
                      <p className="mt-1 text-2xl font-semibold text-white">{compactNumber.format(stat.average)}</p>
                    </div>
                    <span className="rounded-full bg-cyan-300/10 px-2 py-1 text-xs text-cyan-100">avg</span>
                  </div>
                  <div className="mt-4">
                    <MiniBars values={stat.values} />
                  </div>
                  <p className="mt-3 text-xs text-slate-500">
                    {compactNumber.format(stat.min)} min · {compactNumber.format(stat.max)} max
                  </p>
                </GlassPanel>
              ))}
            </div>

            <div className="grid gap-6 xl:grid-cols-[1.1fr_0.9fr]">
              <GlassPanel className="p-5">
                <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <h2 className="text-xl font-semibold text-white">Chart field</h2>
                    <p className="mt-1 text-sm text-slate-400">Compare numeric dimensions from the filtered dataset.</p>
                  </div>
                  <div className="flex gap-3">
                    <select className="field-select" onChange={(event) => setXColumn(event.target.value)} value={selectedX}>
                      {numericColumns.map((column) => (
                        <option key={column}>{column}</option>
                      ))}
                    </select>
                    <select className="field-select" onChange={(event) => setYColumn(event.target.value)} value={selectedY}>
                      {numericColumns.map((column) => (
                        <option key={column}>{column}</option>
                      ))}
                    </select>
                  </div>
                </div>
                <div className="mt-6">
                  <ScatterPlot rows={filteredRows} xColumn={selectedX} yColumn={selectedY} />
                </div>
              </GlassPanel>

              <GlassPanel className="p-5">
                <h2 className="text-xl font-semibold text-white">Correlation analysis</h2>
                <p className="mt-1 text-sm text-slate-400">Pearson coefficients ranked by absolute strength.</p>
                <div className="mt-5 space-y-3">
                  {correlations.slice(0, 8).map((item) => (
                    <div className="rounded-md border border-white/10 bg-white/[0.04] p-3" key={`${item.left}-${item.right}`}>
                      <div className="flex items-center justify-between gap-3 text-sm">
                        <span className="truncate text-slate-200">{item.left} × {item.right}</span>
                        <span className={item.value >= 0 ? "text-cyan-200" : "text-violet-200"}>{item.value.toFixed(3)}</span>
                      </div>
                      <div className="mt-3 h-2 overflow-hidden rounded-full bg-white/10">
                        <div
                          className={item.value >= 0 ? "h-full bg-cyan-300" : "h-full bg-violet-300"}
                          style={{ width: `${Math.abs(item.value) * 100}%` }}
                        />
                      </div>
                    </div>
                  ))}
                  {correlations.length === 0 && <div className="empty-state">Upload at least two numeric columns with three paired values.</div>}
                </div>
              </GlassPanel>
            </div>

            <GlassPanel className="p-5">
              <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
                <div>
                  <h2 className="text-xl font-semibold text-white">Filters and data preview</h2>
                  <p className="mt-1 text-sm text-slate-400">Filter per column, sort headers, and export the filtered view.</p>
                </div>
                <button className="quiet-button self-start lg:self-auto" onClick={() => setFilters({})} type="button">
                  Clear filters
                </button>
              </div>

              <div className="mt-5 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                {columns.map((column) => {
                  const numeric = numericColumns.includes(column);
                  const filter = filters[column] ?? { query: "", min: "", max: "" };

                  return (
                    <div className="rounded-md border border-white/10 bg-white/[0.04] p-3" key={column}>
                      <label className="text-xs font-semibold uppercase tracking-[0.16em] text-slate-400">{column}</label>
                      <input
                        className="filter-input mt-2"
                        onChange={(event) => updateFilter(column, "query", event.target.value)}
                        placeholder="Contains"
                        value={filter.query}
                      />
                      {numeric && (
                        <div className="mt-2 grid grid-cols-2 gap-2">
                          <input className="filter-input" onChange={(event) => updateFilter(column, "min", event.target.value)} placeholder="Min" type="number" value={filter.min} />
                          <input className="filter-input" onChange={(event) => updateFilter(column, "max", event.target.value)} placeholder="Max" type="number" value={filter.max} />
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              <div className="mt-6 overflow-hidden rounded-lg border border-white/10">
                <div className="max-h-[520px] overflow-auto">
                  <table className="w-full min-w-[760px] border-collapse text-left text-sm">
                    <thead className="sticky top-0 z-10 bg-[#10142c]/95 backdrop-blur">
                      <tr>
                        {columns.map((column) => (
                          <th className="border-b border-white/10 px-4 py-3 font-semibold text-slate-200" key={column}>
                            <button
                              className="flex w-full items-center justify-between gap-2 text-left"
                              onClick={() =>
                                setSort((current) =>
                                  current?.column === column
                                    ? { column, direction: current.direction === "asc" ? "desc" : "asc" }
                                    : { column, direction: "asc" },
                                )
                              }
                              type="button"
                            >
                              <span className="truncate">{column}</span>
                              <span className="text-cyan-200">{sort?.column === column ? (sort.direction === "asc" ? "↑" : "↓") : "↕"}</span>
                            </button>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.slice(0, 100).map((row, rowIndex) => (
                        <tr className="odd:bg-white/[0.025] hover:bg-cyan-300/[0.06]" key={rowIndex}>
                          {columns.map((column) => (
                            <td className="border-b border-white/5 px-4 py-3 text-slate-300" key={column}>
                              {formatValue(row[column] ?? null)}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
              <p className="mt-3 text-xs text-slate-500">Showing up to 100 rows in preview.</p>
            </GlassPanel>
          </>
        )}
      </div>
    </main>
  );
}
