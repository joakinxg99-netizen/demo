"use client";

import dynamic from "next/dynamic";
import { ChangeEvent, DragEvent, useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

type CellValue = string | number | null;
type DataRow = Record<string, CellValue>;
type FilterState = Record<string, { query: string; min: string; max: string }>;
type SortState = { column: string; direction: "asc" | "desc" } | null;
type TabKey = "overview" | "visualization" | "correlation" | "preview" | "export";
type ChartType = "scatter" | "bar" | "line" | "histogram";
type PlotTrace = Record<string, unknown>;

const Plot = dynamic(() => import("react-plotly.js"), {
  ssr: false,
  loading: () => <div className="empty-state">Charging the chart renderer...</div>,
});

const TABS: { key: TabKey; label: string }[] = [
  { key: "overview", label: "Overview" },
  { key: "visualization", label: "Visualization" },
  { key: "correlation", label: "Correlation" },
  { key: "preview", label: "Data Preview" },
  { key: "export", label: "Export" },
];

const SAMPLE_ROWS: DataRow[] = [
  { country: "Japan", region: "East Asia", year: 2024, population_m: 123.2, median_age: 49.1, fertility_rate: 1.2, urban_pct: 92, life_expectancy: 84.8 },
  { country: "Nigeria", region: "West Africa", year: 2024, population_m: 229.2, median_age: 17.2, fertility_rate: 5.0, urban_pct: 54, life_expectancy: 54.6 },
  { country: "Brazil", region: "South America", year: 2024, population_m: 217.6, median_age: 34.6, fertility_rate: 1.6, urban_pct: 87, life_expectancy: 76.4 },
  { country: "Germany", region: "Western Europe", year: 2024, population_m: 83.3, median_age: 45.7, fertility_rate: 1.5, urban_pct: 78, life_expectancy: 81.4 },
  { country: "India", region: "South Asia", year: 2024, population_m: 1450.9, median_age: 28.8, fertility_rate: 2.0, urban_pct: 37, life_expectancy: 72.2 },
  { country: "Mexico", region: "North America", year: 2024, population_m: 129.7, median_age: 30.6, fertility_rate: 1.8, urban_pct: 82, life_expectancy: 75.1 },
  { country: "Egypt", region: "North Africa", year: 2024, population_m: 114.5, median_age: 24.2, fertility_rate: 2.8, urban_pct: 43, life_expectancy: 71.6 },
  { country: "Canada", region: "North America", year: 2024, population_m: 39.7, median_age: 41.8, fertility_rate: 1.4, urban_pct: 82, life_expectancy: 82.9 },
  { country: "Indonesia", region: "Southeast Asia", year: 2024, population_m: 283.5, median_age: 30.4, fertility_rate: 2.1, urban_pct: 59, life_expectancy: 72.8 },
  { country: "France", region: "Western Europe", year: 2024, population_m: 66.5, median_age: 42.3, fertility_rate: 1.8, urban_pct: 82, life_expectancy: 82.7 },
  { country: "Ethiopia", region: "East Africa", year: 2024, population_m: 129.7, median_age: 19.1, fertility_rate: 3.9, urban_pct: 23, life_expectancy: 66.5 },
  { country: "South Korea", region: "East Asia", year: 2024, population_m: 51.7, median_age: 45.1, fertility_rate: 0.8, urban_pct: 81, life_expectancy: 83.7 },
];

const compactNumber = new Intl.NumberFormat("en", {
  maximumFractionDigits: 2,
  notation: "compact",
});

const CHART_COLORS = [
  "#2563eb",
  "#7c3aed",
  "#0f766e",
  "#b45309",
  "#be123c",
  "#475569",
  "#0891b2",
  "#9333ea",
  "#15803d",
  "#c2410c",
  "#64748b",
  "#4f46e5",
];

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
  const pairs = getNumericPairs(rows, xColumn, yColumn);

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

function getNumericPairs(rows: DataRow[], xColumn: string, yColumn: string) {
  return rows
    .map((row) => [row[xColumn], row[yColumn]])
    .filter((pair): pair is [number, number] => typeof pair[0] === "number" && typeof pair[1] === "number");
}

function linearRegression(rows: DataRow[], xColumn: string, yColumn: string) {
  const pairs = getNumericPairs(rows, xColumn, yColumn);

  if (pairs.length < 3) {
    return null;
  }

  const xs = pairs.map(([x]) => x);
  const ys = pairs.map(([, y]) => y);
  const xMean = mean(xs);
  const yMean = mean(ys);
  const numerator = pairs.reduce((total, [x, y]) => total + (x - xMean) * (y - yMean), 0);
  const denominator = pairs.reduce((total, [x]) => total + (x - xMean) ** 2, 0);

  if (denominator === 0) {
    return null;
  }

  const slope = numerator / denominator;
  const intercept = yMean - slope * xMean;
  const totalSquares = ys.reduce((total, y) => total + (y - yMean) ** 2, 0);
  const residualSquares = pairs.reduce((total, [x, y]) => total + (y - (slope * x + intercept)) ** 2, 0);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);

  return {
    intercept,
    lineX: [minX, maxX],
    lineY: [slope * minX + intercept, slope * maxX + intercept],
    r2: totalSquares === 0 ? 1 : 1 - residualSquares / totalSquares,
    slope,
  };
}

function formatValue(value: CellValue) {
  if (typeof value === "number") {
    return Number.isInteger(value) ? value.toLocaleString() : value.toLocaleString(undefined, { maximumFractionDigits: 3 });
  }

  return value ?? "-";
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

function EmptyState({ title, detail, action }: { title: string; detail: string; action?: React.ReactNode }) {
  return (
    <div className="empty-state">
      <div>
        <p className="text-lg font-semibold text-white">{title}</p>
        <p className="mt-2 max-w-xl text-sm leading-6 text-slate-400">{detail}</p>
        {action && <div className="mt-5">{action}</div>}
      </div>
    </div>
  );
}

function MiniBars({ values }: { values: number[] }) {
  const sample = values.slice(0, 32);

  if (sample.length === 0) {
    return <div className="h-16 rounded-md bg-white/[0.04]" />;
  }

  const min = Math.min(...sample);
  const max = Math.max(...sample);
  const span = max - min || 1;

  return (
    <div className="flex h-16 items-end gap-1">
      {sample.map((value, index) => (
        <div
          className="min-w-1 flex-1 rounded-t bg-blue-500/75"
          key={`${value}-${index}`}
          style={{ height: `${18 + ((value - min) / span) * 82}%` }}
          title={formatValue(value)}
        />
      ))}
    </div>
  );
}

function getGroupedRows(rows: DataRow[], groupColumn: string) {
  if (!groupColumn) {
    return [["All rows", rows]] as [string, DataRow[]][];
  }

  return Array.from(
    rows.reduce((groups, row) => {
      const key = String(row[groupColumn] ?? "Missing");
      groups.set(key, [...(groups.get(key) ?? []), row]);
      return groups;
    }, new Map<string, DataRow[]>()),
  ).slice(0, 12);
}

function buildPlotData({
  chartType,
  groupColumn,
  regression,
  rows,
  xColumn,
  yColumn,
}: {
  chartType: ChartType;
  groupColumn: string;
  regression: ReturnType<typeof linearRegression>;
  rows: DataRow[];
  xColumn: string;
  yColumn: string;
}) {
  if (!xColumn || (chartType !== "histogram" && !yColumn)) {
    return [];
  }

  const traces: PlotTrace[] = getGroupedRows(rows, groupColumn).map(([group, groupRows], index) => {
    const color = CHART_COLORS[index % CHART_COLORS.length];
    const chartRows =
      chartType === "histogram"
        ? groupRows.filter((row) => typeof row[xColumn] === "number")
        : groupRows.filter((row) => row[xColumn] !== null && row[yColumn] !== null);
    const x = chartRows.map((row) => row[xColumn]).filter((value): value is string | number => value !== null && value !== undefined);
    const y = chartRows.map((row) => row[yColumn]).filter((value): value is string | number => value !== null && value !== undefined);
    const baseTrace = {
      marker: { line: { color: "rgba(15, 23, 42, 0.12)", width: 1 }, opacity: 0.82 },
      name: group,
    };

    if (chartType === "histogram") {
      return {
        ...baseTrace,
        marker: { color },
        type: "histogram",
        x,
      };
    }

    if (chartType === "bar") {
      return {
        ...baseTrace,
        marker: { color },
        type: "bar",
        x,
        y,
      };
    }

    return {
      ...baseTrace,
      line: { color, width: 2.5 },
      marker: { color, size: chartType === "scatter" ? 8 : 6 },
      mode: chartType === "line" ? "lines+markers" : "markers",
      type: "scatter",
      x,
      y,
    };
  });

  if (chartType === "scatter" && regression) {
    traces.push({
      line: { color: "#334155", dash: "dot", width: 2.5 },
      mode: "lines",
      name: "Linear trend",
      type: "scatter",
      x: regression.lineX,
      y: regression.lineY,
    });
  }

  return traces;
}

function FilterPanel({
  columns,
  filters,
  isOpen,
  numericColumns,
  onClear,
  onToggle,
  onUpdate,
}: {
  columns: string[];
  filters: FilterState;
  isOpen: boolean;
  numericColumns: string[];
  onClear: () => void;
  onToggle: () => void;
  onUpdate: (column: string, key: "query" | "min" | "max", value: string) => void;
}) {
  const activeCount = Object.values(filters).filter((filter) => filter.query || filter.min || filter.max).length;

  return (
    <GlassPanel className="p-4 sm:p-5">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h2 className="text-lg font-semibold text-white">Filters</h2>
          <p className="mt-1 text-sm text-slate-400">{activeCount ? `${activeCount} active filter groups` : "Collapsed until you need to focus the dataset."}</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button className="quiet-button" onClick={onClear} type="button">
            Clear
          </button>
          <button className="quiet-button" onClick={onToggle} type="button">
            {isOpen ? "Hide filters" : "Show filters"}
          </button>
        </div>
      </div>

      {isOpen && (
        <div className="mt-5 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
          {columns.map((column) => {
            const numeric = numericColumns.includes(column);
            const filter = filters[column] ?? { query: "", min: "", max: "" };

            return (
              <div className="rounded-md border border-white/10 bg-white/[0.04] p-3" key={column}>
                <label className="text-xs font-semibold uppercase tracking-[0.16em] text-slate-400">{column}</label>
                <input
                  className="filter-input mt-2"
                  onChange={(event) => onUpdate(column, "query", event.target.value)}
                  placeholder="Contains"
                  value={filter.query}
                />
                {numeric && (
                  <div className="mt-2 grid grid-cols-2 gap-2">
                    <input className="filter-input" onChange={(event) => onUpdate(column, "min", event.target.value)} placeholder="Min" type="number" value={filter.min} />
                    <input className="filter-input" onChange={(event) => onUpdate(column, "max", event.target.value)} placeholder="Max" type="number" value={filter.max} />
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </GlassPanel>
  );
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [rows, setRows] = useState<DataRow[]>([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [filters, setFilters] = useState<FilterState>({});
  const [sort, setSort] = useState<SortState>(null);
  const [activeTab, setActiveTab] = useState<TabKey>("overview");
  const [chartType, setChartType] = useState<ChartType>("scatter");
  const [xColumn, setXColumn] = useState("");
  const [yColumn, setYColumn] = useState("");
  const [groupColumn, setGroupColumn] = useState("");
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [dragging, setDragging] = useState(false);

  const columns = useMemo(() => getColumns(rows), [rows]);
  const numericColumns = useMemo(() => columns.filter((column) => isNumericColumn(rows, column)), [columns, rows]);
  const textColumns = useMemo(() => columns.filter((column) => !numericColumns.includes(column)), [columns, numericColumns]);

  const filteredRows = useMemo(() => {
    const nextRows = rows.filter((row) =>
      columns.every((column) => {
        const filter = filters[column];
        if (!filter) {
          return true;
        }

        const value = row[column];
        const textMatch = !filter.query || String(value ?? "").toLowerCase().includes(filter.query.trim().toLowerCase());
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

  const selectedX = xColumn || numericColumns[0] || "";
  const selectedY = yColumn || numericColumns.find((column) => column !== selectedX) || numericColumns[0] || "";
  const regression = useMemo(() => linearRegression(filteredRows, selectedX, selectedY), [filteredRows, selectedX, selectedY]);
  const plotData = useMemo(
    () =>
      buildPlotData({
        chartType,
        groupColumn,
        regression,
        rows: filteredRows,
        xColumn: selectedX,
        yColumn: selectedY,
      }),
    [chartType, filteredRows, groupColumn, regression, selectedX, selectedY],
  );

  async function loadFile(file: File) {
    setError("");
    setFileName(file.name);

    try {
      const extension = file.name.toLowerCase().split(".").pop();

      if (extension === "csv") {
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
      } else if (extension === "xlsx" || extension === "xls") {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        if (!worksheet) {
          throw new Error("The workbook does not contain a readable first sheet.");
        }

        const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { defval: null });
        setRows(normalizeRows(json));
      } else {
        throw new Error("Upload a CSV, XLS, or XLSX file.");
      }

      setActiveTab("overview");
      setFilters({});
      setFiltersOpen(false);
      setGroupColumn("");
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
    setFiltersOpen(false);
    setGroupColumn("region");
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
    if (filteredRows.length === 0) {
      setError("There are no filtered rows to export.");
      return;
    }

    const blob = new Blob([toCsv(filteredRows, columns)], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `demography-export-${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  const hasData = rows.length > 0;
  const chartReady = numericColumns.length >= (chartType === "histogram" ? 1 : 2) && filteredRows.length > 0;

  return (
    <main className="relative min-h-screen overflow-hidden px-3 py-6 text-slate-100 sm:px-6 lg:px-10">
      <div className="cosmic-grid" />
      <div className="star-field" />
      <div className="mx-auto flex w-full max-w-7xl flex-col gap-5 sm:gap-6">
        <header className="flex flex-col items-center gap-5 py-6 text-center sm:py-8">
          <div className="rounded-full border border-blue-200/25 bg-white/10 px-4 py-2 text-xs font-semibold uppercase tracking-[0.24em] text-blue-100">
            Demography Research Studio
          </div>
          <div className="max-w-4xl">
            <h1 className="text-4xl font-semibold tracking-normal text-white sm:text-6xl">Demography Explorer</h1>
            <p className="mt-4 text-base leading-7 text-slate-300 sm:text-lg">
              Upload population, migration, age, fertility, income, or region data and explore patterns in a clear client-side workspace.
            </p>
          </div>
        </header>

        <GlassPanel className="mx-auto w-full max-w-4xl p-4 sm:p-6">
          <div
            className={`flex min-h-56 flex-col items-center justify-center rounded-lg border border-dashed p-6 text-center transition sm:min-h-64 sm:p-8 ${
              dragging ? "border-blue-300 bg-blue-50/80" : "border-slate-300 bg-slate-50/70"
            }`}
            onDragLeave={() => setDragging(false)}
            onDragOver={(event) => {
              event.preventDefault();
              setDragging(true);
            }}
            onDrop={handleDrop}
          >
            <input accept=".csv,.xlsx,.xls" className="hidden" onChange={handleFileChange} ref={inputRef} type="file" />
            <div className="mb-5 flex h-14 w-14 items-center justify-center rounded-full border border-blue-200 bg-blue-50 text-2xl text-blue-700 sm:h-16 sm:w-16">
              ↑
            </div>
            <h2 className="text-2xl font-semibold text-white">Drop a CSV or Excel file</h2>
            <p className="mt-2 max-w-xl text-sm leading-6 text-slate-300">
              Data stays in your browser. The first sheet is loaded for Excel workbooks, and numeric-looking values are detected automatically.
            </p>
            <div className="mt-6 flex flex-wrap items-center justify-center gap-3">
              <button className="primary-button" onClick={() => inputRef.current?.click()} type="button">
                Select file
              </button>
              <button className="quiet-button" onClick={loadSample} type="button">
                Load sample
              </button>
            </div>
            {fileName && <p className="mt-4 text-sm font-medium text-blue-700">Loaded: {fileName}</p>}
            {error && (
              <div className="mt-4 rounded-md border border-rose-300/30 bg-rose-400/10 px-4 py-3 text-sm text-rose-100">
                {error}
              </div>
            )}
          </div>
        </GlassPanel>

        {!hasData && (
          <EmptyState
            title="Your dashboard is waiting for a dataset"
            detail="Use the uploader above or load the sample to unlock tabs, charts, correlations, filters, preview, and export tools."
          />
        )}

        {hasData && (
          <>
            <nav aria-label="Explorer sections" className="tab-shell">
              {TABS.map((tab) => (
                <button
                  aria-current={activeTab === tab.key ? "page" : undefined}
                  className={`tab-button ${activeTab === tab.key ? "tab-button-active" : ""}`}
                  key={tab.key}
                  onClick={() => setActiveTab(tab.key)}
                  type="button"
                >
                  {tab.label}
                </button>
              ))}
            </nav>

            <FilterPanel
              columns={columns}
              filters={filters}
              isOpen={filtersOpen}
              numericColumns={numericColumns}
              onClear={() => setFilters({})}
              onToggle={() => setFiltersOpen((current) => !current)}
              onUpdate={updateFilter}
            />

            {filteredRows.length === 0 && (
              <EmptyState
                title="No rows match the current filters"
                detail="Relax one or more filter values, then the cards, charts, correlations, preview, and export will update instantly."
                action={
                  <button className="quiet-button" onClick={() => setFilters({})} type="button">
                    Clear filters
                  </button>
                }
              />
            )}

            {activeTab === "overview" && (
              <div className="space-y-5">
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
                    <p className="mt-2 text-3xl font-semibold text-white">{correlations[0] ? correlations[0].value.toFixed(2) : "-"}</p>
                    <p className="mt-1 truncate text-xs text-slate-500">{correlations[0] ? `${correlations[0].left} x ${correlations[0].right}` : "Need two numeric fields"}</p>
                  </GlassPanel>
                  <GlassPanel className="p-5">
                    <p className="text-sm text-slate-400">Filtered export</p>
                    <button className="mt-3 w-full rounded-md border border-blue-200 bg-blue-50 px-4 py-3 text-sm font-semibold text-blue-800 transition hover:bg-blue-100" onClick={exportCsv} type="button">
                      Download CSV
                    </button>
                  </GlassPanel>
                </div>

                <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                  {primaryStats.length > 0 ? (
                    primaryStats.map((stat) => (
                      <GlassPanel className="p-5" key={stat.column}>
                        <div className="flex items-start justify-between gap-3">
                          <div>
                            <p className="text-sm text-slate-400">{stat.column}</p>
                            <p className="mt-1 text-2xl font-semibold text-white">{compactNumber.format(stat.average)}</p>
                          </div>
                          <span className="rounded-full bg-blue-50 px-2 py-1 text-xs font-medium text-blue-700">avg</span>
                        </div>
                        <div className="mt-4">
                          <MiniBars values={stat.values} />
                        </div>
                        <p className="mt-3 text-xs text-slate-500">
                          {compactNumber.format(stat.min)} min / {compactNumber.format(stat.max)} max
                        </p>
                      </GlassPanel>
                    ))
                  ) : (
                    <GlassPanel className="p-5 md:col-span-2 xl:col-span-4">
                      <EmptyState title="No numeric summary yet" detail="This file loaded successfully, but no columns look numeric enough for summary cards." />
                    </GlassPanel>
                  )}
                </div>
              </div>
            )}

            {activeTab === "visualization" && (
              <GlassPanel className="p-4 sm:p-5">
                <div className="flex flex-col gap-4 xl:flex-row xl:items-start xl:justify-between">
                  <div className="max-w-2xl">
                    <h2 className="text-xl font-semibold text-white">Visualization</h2>
                    <p className="mt-1 text-sm leading-6 text-slate-400">
                      Build a readable chart from the filtered rows. Start with scatter for relationships, line for ordered measures, bar for category comparisons, or histogram for a single numeric distribution.
                    </p>
                    <div className="mt-4 rounded-md border border-blue-200 bg-blue-50 px-4 py-3 text-sm leading-6 text-blue-900">
                      Tip: the color field is optional. Use it for regions, cohorts, or other groups when you want separate series.
                    </div>
                  </div>
                  <div className="chart-controls">
                    <label>
                      <span>Chart type</span>
                      <select className="field-select" onChange={(event) => setChartType(event.target.value as ChartType)} value={chartType}>
                        <option value="scatter">Scatter</option>
                        <option value="bar">Bar</option>
                        <option value="line">Line</option>
                        <option value="histogram">Histogram</option>
                      </select>
                    </label>
                    <label>
                      <span>X column</span>
                      <select className="field-select" onChange={(event) => setXColumn(event.target.value)} value={selectedX}>
                        {numericColumns.map((column) => (
                          <option key={column}>{column}</option>
                        ))}
                      </select>
                    </label>
                    <label>
                      <span>Y column</span>
                      <select className="field-select" disabled={chartType === "histogram"} onChange={(event) => setYColumn(event.target.value)} value={selectedY}>
                        {numericColumns.map((column) => (
                          <option key={column}>{column}</option>
                        ))}
                      </select>
                    </label>
                    <label>
                      <span>Color group</span>
                      <select className="field-select" onChange={(event) => setGroupColumn(event.target.value)} value={groupColumn}>
                        <option value="">No grouping</option>
                        {[...textColumns, ...numericColumns].map((column) => (
                          <option key={column} value={column}>
                            {column}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>
                </div>

                <div className="mt-6 chart-frame">
                  {chartReady ? (
                    <Plot
                      config={{ displaylogo: false, responsive: true }}
                      data={plotData}
                      layout={{
                        autosize: true,
                        bargap: 0.18,
                        barmode: groupColumn ? "group" : "overlay",
                        font: { color: "#334155", family: "Arial, sans-serif" },
                        height: 440,
                        legend: { font: { color: "#475569" }, orientation: "h", x: 0, y: 1.14 },
                        margin: { b: 64, l: 58, r: 24, t: 36 },
                        paper_bgcolor: "rgba(0,0,0,0)",
                        plot_bgcolor: "#f8fafc",
                        xaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: selectedX, zerolinecolor: "#cbd5e1" },
                        yaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: chartType === "histogram" ? "Count" : selectedY, zerolinecolor: "#cbd5e1" },
                      }}
                      style={{ height: "100%", width: "100%" }}
                      useResizeHandler
                    />
                  ) : (
                    <EmptyState title="Chart needs numeric data" detail="Choose a dataset with at least one numeric field for histograms, or two numeric fields for scatter, bar, and line charts." />
                  )}
                </div>
              </GlassPanel>
            )}

            {activeTab === "correlation" && (
              <div className="grid gap-6 xl:grid-cols-[0.95fr_1.05fr]">
                <GlassPanel className="p-5">
                  <h2 className="text-xl font-semibold text-white">Correlation analysis</h2>
                  <p className="mt-1 text-sm text-slate-400">Pearson coefficients ranked by absolute strength.</p>
                  <div className="mt-5 space-y-3">
                    {correlations.slice(0, 12).map((item) => (
                      <div className="rounded-md border border-white/10 bg-white/[0.04] p-3" key={`${item.left}-${item.right}`}>
                        <div className="flex items-center justify-between gap-3 text-sm">
                          <span className="truncate text-slate-200">{item.left} x {item.right}</span>
                          <span className={item.value >= 0 ? "text-blue-700" : "text-violet-700"}>{item.value.toFixed(3)}</span>
                        </div>
                        <div className="mt-3 h-2 overflow-hidden rounded-full bg-white/10">
                          <div className={item.value >= 0 ? "h-full bg-blue-500" : "h-full bg-violet-400"} style={{ width: `${Math.abs(item.value) * 100}%` }} />
                        </div>
                      </div>
                    ))}
                    {correlations.length === 0 && <EmptyState title="No correlations yet" detail="Upload at least two numeric columns with three paired values, or loosen filters until paired values reappear." />}
                  </div>
                </GlassPanel>

                <GlassPanel className="p-5">
                  <h2 className="text-xl font-semibold text-white">Linear regression</h2>
                  <p className="mt-1 text-sm text-slate-400">Client-side ordinary least squares for the selected X and Y fields.</p>
                  {regression ? (
                    <div className="mt-5 grid gap-4 sm:grid-cols-3">
                      <div className="metric-tile">
                        <span>Slope</span>
                        <strong>{regression.slope.toFixed(4)}</strong>
                      </div>
                      <div className="metric-tile">
                        <span>Intercept</span>
                        <strong>{regression.intercept.toFixed(4)}</strong>
                      </div>
                      <div className="metric-tile">
                        <span>R squared</span>
                        <strong>{regression.r2.toFixed(4)}</strong>
                      </div>
                    </div>
                  ) : (
                    <EmptyState title="Regression needs paired numeric values" detail="Select numeric X and Y fields in Visualization and keep at least three matching rows after filters." />
                  )}
                  <div className="mt-5 rounded-md border border-blue-200 bg-blue-50 p-4 text-sm leading-6 text-blue-900">
                    Current model: {selectedY || "Y"} = {regression ? regression.slope.toFixed(4) : "m"} * {selectedX || "X"} + {regression ? regression.intercept.toFixed(4) : "b"}
                  </div>
                </GlassPanel>
              </div>
            )}

            {activeTab === "preview" && (
              <GlassPanel className="p-4 sm:p-5">
                <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
                  <div>
                    <h2 className="text-xl font-semibold text-white">Data preview</h2>
                    <p className="mt-1 text-sm text-slate-400">Sort headers and inspect the filtered view without leaving the browser.</p>
                  </div>
                  <button className="quiet-button self-start lg:self-auto" onClick={() => setFiltersOpen((current) => !current)} type="button">
                    {filtersOpen ? "Hide filters" : "Show filters"}
                  </button>
                </div>

                <div className="mt-6 overflow-hidden rounded-lg border border-white/10">
                  <div className="max-h-[560px] overflow-auto">
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
                                <span className="text-blue-600">{sort?.column === column ? (sort.direction === "asc" ? "↑" : "↓") : "↕"}</span>
                              </button>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {filteredRows.slice(0, 150).map((row, rowIndex) => (
                          <tr className="odd:bg-slate-50 hover:bg-blue-50" key={rowIndex}>
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
                <p className="mt-3 text-xs text-slate-500">Showing up to 150 rows in preview.</p>
              </GlassPanel>
            )}

            {activeTab === "export" && (
              <GlassPanel className="p-5">
                <div className="grid gap-5 lg:grid-cols-[1fr_0.8fr]">
                  <div>
                    <h2 className="text-xl font-semibold text-white">Export filtered data</h2>
                    <p className="mt-2 max-w-2xl text-sm leading-6 text-slate-400">
                      Export uses the current filters and sort order. No backend is involved; the CSV is generated directly in this browser session.
                    </p>
                    <div className="mt-6 flex flex-wrap gap-3">
                      <button className="primary-button" onClick={exportCsv} type="button">
                        Download filtered CSV
                      </button>
                      <button className="quiet-button" onClick={() => setActiveTab("preview")} type="button">
                        Inspect preview
                      </button>
                    </div>
                  </div>
                  <div className="grid gap-3 sm:grid-cols-3 lg:grid-cols-1">
                    <div className="metric-tile">
                      <span>Rows exported</span>
                      <strong>{filteredRows.length.toLocaleString()}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Columns</span>
                      <strong>{columns.length}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Source</span>
                      <strong className="truncate text-base">{fileName || "Untitled"}</strong>
                    </div>
                  </div>
                </div>
              </GlassPanel>
            )}
          </>
        )}
      </div>
    </main>
  );
}
