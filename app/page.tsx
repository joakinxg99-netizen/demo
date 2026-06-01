"use client";

import dynamic from "next/dynamic";
import { ChangeEvent, DragEvent, useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

type CellValue = string | number | null;
type DataRow = Record<string, CellValue>;
type FilterState = Record<string, { query: string; min: string; max: string }>;
type SortState = { column: string; direction: "asc" | "desc" } | null;
type TabKey = "report" | "overview" | "population" | "income" | "labour" | "insights" | "quality" | "visualization" | "preview" | "export";
type ChartType = "scatter" | "bar" | "line" | "histogram";
type VisualizationMode =
  | "populationPyramid"
  | "ageDistribution"
  | "incomeDistribution"
  | "employmentStructure"
  | "educationProfile"
  | "regionalComparison"
  | "customRelationship";
type PlotTrace = Record<string, unknown>;
type SurveyVariableKey =
  | "age"
  | "sex"
  | "individualIncome"
  | "householdIncome"
  | "employment"
  | "occupation"
  | "education"
  | "region"
  | "householdSize"
  | "children"
  | "year"
  | "quarter"
  | "householdId"
  | "personId";
type SurveyDetection = {
  column: string;
  concept: SurveyVariableKey;
  confidence: number;
  friendlyLabel: string;
  method: string;
};
type ColumnProfile = {
  average: number | null;
  column: string;
  completeness: number;
  frequencies: { count: number; label: string; percent: number }[];
  max: number | null;
  median: number | null;
  min: number | null;
  missing: number;
  numericCount: number;
  textCount: number;
  type: "empty" | "mixed" | "numeric" | "text";
  unique: number;
};

const Plot = dynamic(() => import("react-plotly.js"), {
  ssr: false,
  loading: () => <div className="empty-state">Charging the chart renderer...</div>,
});

const TABS: { key: TabKey; label: string }[] = [
  { key: "report", label: "Dataset Report" },
  { key: "overview", label: "Overview" },
  { key: "population", label: "Population" },
  { key: "income", label: "Income" },
  { key: "labour", label: "Labour" },
  { key: "insights", label: "Insights" },
  { key: "quality", label: "Data Quality" },
  { key: "visualization", label: "Visualization" },
  { key: "preview", label: "Data Preview" },
  { key: "export", label: "Export" },
];

const REPORT_DETECTION_GROUPS: { concepts: SurveyVariableKey[]; title: string }[] = [
  { title: "Population", concepts: ["age", "sex", "householdSize", "children"] },
  { title: "Income", concepts: ["individualIncome", "householdIncome"] },
  { title: "Labour", concepts: ["employment", "occupation"] },
  { title: "Context", concepts: ["region", "year", "quarter"] },
  { title: "Identifiers", concepts: ["householdId", "personId"] },
];

const SAMPLE_ROWS: DataRow[] = [
  { respondent_id: 1, region: "North", age: 22, sex: "Female", education: "University", employment_status: "Student", monthly_income: 950, household_size: 3, children: 0, housing_tenure: "Rent" },
  { respondent_id: 2, region: "North", age: 41, sex: "Male", education: "Secondary", employment_status: "Employed", monthly_income: 3200, household_size: 4, children: 2, housing_tenure: "Own" },
  { respondent_id: 3, region: "South", age: 35, sex: "Female", education: "University", employment_status: "Employed", monthly_income: 4100, household_size: 2, children: 1, housing_tenure: "Own" },
  { respondent_id: 4, region: "South", age: 67, sex: "Female", education: "Primary", employment_status: "Retired", monthly_income: 1800, household_size: 1, children: 3, housing_tenure: "Own" },
  { respondent_id: 5, region: "East", age: 29, sex: "Male", education: "Technical", employment_status: "Self-employed", monthly_income: 2700, household_size: 5, children: 2, housing_tenure: "Rent" },
  { respondent_id: 6, region: "East", age: 48, sex: "Female", education: "Secondary", employment_status: "Unemployed", monthly_income: 620, household_size: 4, children: 2, housing_tenure: "Rent" },
  { respondent_id: 7, region: "West", age: 53, sex: "Male", education: "University", employment_status: "Employed", monthly_income: 5300, household_size: 3, children: 1, housing_tenure: "Own" },
  { respondent_id: 8, region: "West", age: 19, sex: "Female", education: "Secondary", employment_status: "Student", monthly_income: 420, household_size: 4, children: 0, housing_tenure: "Family" },
  { respondent_id: 9, region: "Central", age: 38, sex: "Male", education: "Technical", employment_status: "Employed", monthly_income: 3600, household_size: 2, children: 0, housing_tenure: "Rent" },
  { respondent_id: 10, region: "Central", age: 44, sex: "Female", education: "University", employment_status: "Self-employed", monthly_income: 3900, household_size: 3, children: 1, housing_tenure: "Own" },
  { respondent_id: 11, region: "North", age: 61, sex: "Male", education: "Primary", employment_status: "Retired", monthly_income: 1550, household_size: 2, children: 3, housing_tenure: "Own" },
  { respondent_id: 12, region: "South", age: 26, sex: "Female", education: "University", employment_status: "Employed", monthly_income: 2800, household_size: 1, children: 0, housing_tenure: "Rent" },
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

const SURVEY_CONCEPTS: Record<
  SurveyVariableKey,
  {
    aliases: string[];
    ephAliases?: string[];
    friendlyLabel: string;
    keywords: string[];
    partials: string[];
  }
> = {
  age: {
    aliases: ["age", "respondent_age", "person_age", "edad", "idade", "anos", "edad_anios", "edad_anos"],
    ephAliases: ["ch06"],
    friendlyLabel: "Age",
    keywords: ["age", "edad", "idade"],
    partials: ["respondent_age", "person_age", "age_year", "edad", "idade"],
  },
  sex: {
    aliases: ["sex", "gender", "sexo", "genero", "gênero", "male_female", "female_male"],
    ephAliases: ["ch04"],
    friendlyLabel: "Sex / Gender",
    keywords: ["sex", "gender", "sexo", "genero"],
    partials: ["gender", "sexo", "genero", "male_female"],
  },
  individualIncome: {
    aliases: ["income", "individual_income", "personal_income", "earnings", "wage", "salary", "pay", "ingreso", "ingresos", "salario", "renda", "rendimento"],
    ephAliases: ["p47t"],
    friendlyLabel: "Individual income",
    keywords: ["income", "earnings", "wage", "salary", "pay", "ingreso", "ingresos", "salario", "renda", "rendimento"],
    partials: ["individual_income", "personal_income", "monthly_income", "labor_income", "labour_income", "ingreso", "renda", "rendimento"],
  },
  householdIncome: {
    aliases: ["household_income", "family_income", "household_per_capita_income", "per_capita_income", "ipcf", "ingreso_hogar", "renda_domiciliar", "renda_familiar"],
    ephAliases: ["ipcf"],
    friendlyLabel: "Household income",
    keywords: ["household_income", "family_income", "per_capita", "ingreso_hogar", "renda_domiciliar", "renda_familiar"],
    partials: ["household_income", "family_income", "per_capita_income", "ipcf", "ingreso_hogar", "renda_domiciliar"],
  },
  employment: {
    aliases: ["employment", "employment_status", "labour_status", "labor_status", "work_status", "job_status", "empleo", "ocupado", "emprego", "estado"],
    ephAliases: ["estado"],
    friendlyLabel: "Employment status",
    keywords: ["employment", "labour", "labor", "work", "job", "empleo", "emprego", "estado"],
    partials: ["employment_status", "labour_status", "labor_status", "work_status", "job_status", "estado"],
  },
  occupation: {
    aliases: ["occupation", "occupational_category", "occupation_category", "job_type", "ocupacion", "ocupação", "ocupacao", "cat_ocup"],
    ephAliases: ["cat_ocup"],
    friendlyLabel: "Occupation",
    keywords: ["occupation", "ocupacion", "ocupacao", "ocupação", "job_type"],
    partials: ["occupation", "occupational", "ocupacion", "ocupacao", "cat_ocup"],
  },
  education: {
    aliases: ["education", "education_level", "educ", "schooling", "qualification", "educacion", "educación", "educacao", "educação", "nivel_ed"],
    ephAliases: ["nivel_ed"],
    friendlyLabel: "Education",
    keywords: ["education", "educ", "school", "educacion", "educacao", "nivel"],
    partials: ["education_level", "schooling", "qualification", "educacion", "educacao", "nivel_ed"],
  },
  region: {
    aliases: ["region", "state", "province", "county", "district", "area", "location", "provincia", "regiao", "região", "uf", "estado", "aglomerado"],
    ephAliases: ["aglomerado"],
    friendlyLabel: "Region / geography",
    keywords: ["region", "state", "province", "provincia", "regiao", "uf", "district", "area", "location", "aglomerado"],
    partials: ["region", "province", "provincia", "regiao", "location", "aglomerado"],
  },
  householdSize: {
    aliases: ["household_size", "hh_size", "family_size", "household_members", "tam_hogar", "tamanho_domicilio", "domicilio", "hogar"],
    friendlyLabel: "Household size",
    keywords: ["household", "family", "hogar", "domicilio", "domicílio"],
    partials: ["household_size", "hh_size", "family_size", "household_members", "tam_hogar", "tamanho_domicilio"],
  },
  children: {
    aliases: ["children", "num_children", "kids", "dependents", "hijos", "filhos", "menores", "dependientes", "dependentes"],
    friendlyLabel: "Children / dependents",
    keywords: ["children", "kids", "dependents", "hijos", "filhos", "dependientes", "dependentes"],
    partials: ["num_children", "children", "dependents", "hijos", "filhos"],
  },
  year: {
    aliases: ["year", "survey_year", "ano", "anio", "año", "ano4"],
    ephAliases: ["ano4"],
    friendlyLabel: "Year",
    keywords: ["year", "ano", "anio"],
    partials: ["survey_year", "year", "ano4"],
  },
  quarter: {
    aliases: ["quarter", "trimester", "trimestre", "period", "periodo"],
    ephAliases: ["trimestre"],
    friendlyLabel: "Quarter / period",
    keywords: ["quarter", "trimester", "trimestre", "period", "periodo"],
    partials: ["quarter", "trimester", "trimestre", "period"],
  },
  householdId: {
    aliases: ["household_id", "hh_id", "household_sample_id", "sample_id", "codusu", "nro_hogar", "id_hogar", "id_domicilio"],
    ephAliases: ["nro_hogar", "codusu"],
    friendlyLabel: "Household ID",
    keywords: ["household_id", "hh_id", "sample_id", "codusu", "hogar", "domicilio"],
    partials: ["household_id", "hh_id", "sample_id", "codusu", "nro_hogar", "id_hogar"],
  },
  personId: {
    aliases: ["person_id", "individual_id", "respondent_id", "component", "componente", "id_persona", "id_individuo"],
    ephAliases: ["componente"],
    friendlyLabel: "Person ID",
    keywords: ["person_id", "individual_id", "respondent_id", "persona", "individuo", "component"],
    partials: ["person_id", "individual_id", "respondent_id", "componente", "component"],
  },
};

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

function getColumnProfiles(rows: DataRow[], columns: string[]): ColumnProfile[] {
  return columns.map((column) => {
    const values = rows.map((row) => row[column] ?? null);
    const present = values.filter((value) => value !== null);
    const numericValues = present.filter((value): value is number => typeof value === "number" && Number.isFinite(value));
    const sortedNumeric = [...numericValues].sort((a, b) => a - b);
    const mid = Math.floor(sortedNumeric.length / 2);
    const median =
      sortedNumeric.length === 0
        ? null
        : sortedNumeric.length % 2
          ? sortedNumeric[mid]
          : (sortedNumeric[mid - 1] + sortedNumeric[mid]) / 2;
    const counts = present.reduce((map, value) => {
      const key = String(value);
      map.set(key, (map.get(key) ?? 0) + 1);
      return map;
    }, new Map<string, number>());
    const frequencies = Array.from(counts.entries())
      .map(([label, count]) => ({ count, label, percent: present.length ? (count / present.length) * 100 : 0 }))
      .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
    const textCount = present.filter((value) => typeof value === "string").length;
    const type =
      present.length === 0
        ? "empty"
        : numericValues.length === present.length
          ? "numeric"
          : textCount === present.length
            ? "text"
            : "mixed";

    return {
      average: numericValues.length ? mean(numericValues) : null,
      column,
      completeness: rows.length ? (present.length / rows.length) * 100 : 0,
      frequencies,
      max: numericValues.length ? Math.max(...numericValues) : null,
      median,
      min: numericValues.length ? Math.min(...numericValues) : null,
      missing: values.length - present.length,
      numericCount: numericValues.length,
      textCount,
      type,
      unique: new Set(present.map((value) => String(value))).size,
    };
  });
}

function getDuplicateCount(rows: DataRow[], columns: string[]) {
  const counts = new Map<string, number>();

  rows.forEach((row) => {
    const key = JSON.stringify(columns.map((column) => row[column] ?? null));
    counts.set(key, (counts.get(key) ?? 0) + 1);
  });

  return Array.from(counts.values()).reduce((total, count) => total + Math.max(0, count - 1), 0);
}

function getDataQuality(rows: DataRow[], columns: string[], profiles: ColumnProfile[]) {
  const totalCells = rows.length * columns.length;
  const missingValues = profiles.reduce((total, profile) => total + profile.missing, 0);

  return {
    completeness: totalCells ? ((totalCells - missingValues) / totalCells) * 100 : 0,
    duplicateRows: getDuplicateCount(rows, columns),
    missingValues,
    totalCells,
  };
}

function normalizeColumnName(column: string) {
  return column
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function scoreColumnForConcept(column: string, concept: SurveyVariableKey): SurveyDetection | null {
  const definition = SURVEY_CONCEPTS[concept];
  const normalized = normalizeColumnName(column);
  const aliases = definition.aliases.map(normalizeColumnName);
  const ephAliases = (definition.ephAliases ?? []).map(normalizeColumnName);
  const partials = definition.partials.map(normalizeColumnName);
  const keywords = definition.keywords.map(normalizeColumnName);

  if (ephAliases.includes(normalized)) {
    return { column, concept, confidence: 100, friendlyLabel: definition.friendlyLabel, method: "EPH alias" };
  }

  if (aliases.includes(normalized)) {
    return { column, concept, confidence: 96, friendlyLabel: definition.friendlyLabel, method: "Exact alias" };
  }

  if (partials.some((partial) => normalized.includes(partial) || partial.includes(normalized))) {
    return { column, concept, confidence: 84, friendlyLabel: definition.friendlyLabel, method: "Partial string match" };
  }

  if (keywords.some((keyword) => normalized.split("_").includes(keyword) || normalized.includes(`_${keyword}_`) || normalized.startsWith(`${keyword}_`) || normalized.endsWith(`_${keyword}`))) {
    return { column, concept, confidence: 72, friendlyLabel: definition.friendlyLabel, method: "Survey keyword" };
  }

  return null;
}

function detectSurveyVariables(columns: string[]) {
  return (Object.keys(SURVEY_CONCEPTS) as SurveyVariableKey[]).reduce(
    (detections, concept) => {
      const candidates = columns
        .map((column) => scoreColumnForConcept(column, concept))
        .filter((detection): detection is SurveyDetection => detection !== null)
        .sort((a, b) => b.confidence - a.confidence);

      detections[concept] =
        candidates[0] ?? {
          column: "",
          concept,
          confidence: 0,
          friendlyLabel: SURVEY_CONCEPTS[concept].friendlyLabel,
          method: "Not detected",
        };

      return detections;
    },
    {} as Record<SurveyVariableKey, SurveyDetection>,
  );
}

function getSurveyVariables(detections: Record<SurveyVariableKey, SurveyDetection>) {
  return {
    age: detections.age.column,
    children: detections.children.column,
    education: detections.education.column,
    employment: detections.employment.column,
    household: detections.householdSize.column,
    income: detections.individualIncome.column || detections.householdIncome.column,
    occupation: detections.occupation.column,
    region: detections.region.column,
    sex: detections.sex.column,
  };
}

function frequency(rows: DataRow[], column: string) {
  if (!column) {
    return [];
  }

  const counts = rows.reduce((map, row) => {
    const value = row[column];
    if (value === null || value === undefined) {
      return map;
    }
    const key = String(value);
    map.set(key, (map.get(key) ?? 0) + 1);
    return map;
  }, new Map<string, number>());

  return Array.from(counts.entries())
    .map(([label, count]) => ({ count, label, percent: rows.length ? (count / rows.length) * 100 : 0 }))
    .sort((a, b) => b.count - a.count);
}

function numericSummary(rows: DataRow[], column: string) {
  const values = column ? getNumericValues(rows, column) : [];

  if (!values.length) {
    return null;
  }

  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  const median = sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;

  return {
    average: mean(values),
    count: values.length,
    max: Math.max(...values),
    median,
    min: Math.min(...values),
  };
}

function percentile(values: number[], percentileValue: number) {
  if (!values.length) {
    return null;
  }

  const sorted = [...values].sort((a, b) => a - b);
  const index = (sorted.length - 1) * percentileValue;
  const lower = Math.floor(index);
  const upper = Math.ceil(index);
  const weight = index - lower;

  return lower === upper ? sorted[lower] : sorted[lower] * (1 - weight) + sorted[upper] * weight;
}

function getPercentiles(rows: DataRow[], column: string) {
  const values = column ? getNumericValues(rows, column) : [];

  return {
    p10: percentile(values, 0.1),
    p25: percentile(values, 0.25),
    p50: percentile(values, 0.5),
    p75: percentile(values, 0.75),
    p90: percentile(values, 0.9),
  };
}

function averageByGroup(rows: DataRow[], groupColumn: string, valueColumn: string) {
  if (!groupColumn || !valueColumn) {
    return [];
  }

  const groups = rows.reduce((map, row) => {
    const group = row[groupColumn];
    const value = row[valueColumn];

    if (group === null || group === undefined || typeof value !== "number") {
      return map;
    }

    const key = String(group);
    const current = map.get(key) ?? { count: 0, total: 0 };
    map.set(key, { count: current.count + 1, total: current.total + value });
    return map;
  }, new Map<string, { count: number; total: number }>());

  return Array.from(groups.entries())
    .map(([label, item]) => ({ average: item.total / item.count, count: item.count, label }))
    .sort((a, b) => b.average - a.average);
}

function ageBand(age: number) {
  if (age < 18) return "Under 18";
  if (age < 30) return "18-29";
  if (age < 45) return "30-44";
  if (age < 60) return "45-59";
  return "60+";
}

function ageDistribution(rows: DataRow[], ageColumn: string) {
  if (!ageColumn) {
    return [];
  }

  const counts = rows.reduce((map, row) => {
    const age = row[ageColumn];
    if (typeof age !== "number") {
      return map;
    }
    const band = ageBand(age);
    map.set(band, (map.get(band) ?? 0) + 1);
    return map;
  }, new Map<string, number>());

  return ["Under 18", "18-29", "30-44", "45-59", "60+"]
    .map((label) => ({ count: counts.get(label) ?? 0, label, percent: rows.length ? ((counts.get(label) ?? 0) / rows.length) * 100 : 0 }))
    .filter((item) => item.count > 0);
}

function populationPyramid(rows: DataRow[], ageColumn: string, sexColumn: string) {
  if (!ageColumn || !sexColumn) {
    return [];
  }

  const bands = ["Under 18", "18-29", "30-44", "45-59", "60+"];
  const totals = rows.reduce((map, row) => {
    const age = row[ageColumn];
    const sex = row[sexColumn];

    if (typeof age !== "number" || sex === null || sex === undefined) {
      return map;
    }

    const band = ageBand(age);
    const sexKey = String(sex);
    const current = map.get(band) ?? new Map<string, number>();
    current.set(sexKey, (current.get(sexKey) ?? 0) + 1);
    map.set(band, current);
    return map;
  }, new Map<string, Map<string, number>>());
  const sexLabels = Array.from(new Set(rows.map((row) => row[sexColumn]).filter((value) => value !== null).map(String))).slice(0, 2);
  const maxCount = Math.max(1, ...Array.from(totals.values()).flatMap((item) => Array.from(item.values())));

  return bands.map((band) => ({
    band,
    left: { count: totals.get(band)?.get(sexLabels[0] ?? "") ?? 0, label: sexLabels[0] ?? "Group A", width: ((totals.get(band)?.get(sexLabels[0] ?? "") ?? 0) / maxCount) * 100 },
    right: { count: totals.get(band)?.get(sexLabels[1] ?? "") ?? 0, label: sexLabels[1] ?? "Group B", width: ((totals.get(band)?.get(sexLabels[1] ?? "") ?? 0) / maxCount) * 100 },
  }));
}

function rowMatchesSearch(row: DataRow, columns: string[], query: string) {
  const normalized = query.trim().toLowerCase();

  if (!normalized) {
    return true;
  }

  return columns.some((column) => String(row[column] ?? "").toLowerCase().includes(normalized));
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

function InsightCard({ detail, label, tone = "default" }: { detail: string; label: string; tone?: "default" | "warning" }) {
  return (
    <div className={`insight-card ${tone === "warning" ? "insight-card-warning" : ""}`}>
      <span>{label}</span>
      <p>{detail}</p>
    </div>
  );
}

function FrequencyList({ items, title }: { items: { count: number; label: string; percent: number }[]; title: string }) {
  return (
    <div className="domain-card">
      <h3>{title}</h3>
      <div className="mt-4 space-y-3">
        {items.length ? (
          items.slice(0, 6).map((item) => (
            <div key={item.label}>
              <div className="flex items-center justify-between gap-3 text-sm">
                <span className="font-medium text-slate-700">{item.label}</span>
                <span className="text-slate-500">{item.count.toLocaleString()} ({item.percent.toFixed(0)}%)</span>
              </div>
              <div className="quality-bar mt-2">
                <div style={{ width: `${Math.min(100, item.percent)}%` }} />
              </div>
            </div>
          ))
        ) : (
          <p className="text-sm text-slate-500">No matching variable detected.</p>
        )}
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

function TinyDistribution({ profile }: { profile: ColumnProfile }) {
  if (profile.type === "numeric") {
    const bars = profile.frequencies
      .map((item) => ({ ...item, numericLabel: Number(item.label) }))
      .filter((item) => Number.isFinite(item.numericLabel))
      .sort((a, b) => a.numericLabel - b.numericLabel)
      .slice(0, 32);
    const maxCount = Math.max(1, ...bars.map((item) => item.count));

    return (
      <div className="variable-mini-bars">
        {bars.map((item) => (
          <i
            key={item.label}
            style={{ height: `${18 + (item.count / maxCount) * 82}%` }}
            title={`${item.label}: ${item.count.toLocaleString()}`}
          />
        ))}
      </div>
    );
  }

  return (
    <div className="tiny-frequency">
      {profile.frequencies.slice(0, 5).map((item) => (
        <div key={item.label}>
          <span title={item.label}>{item.label}</span>
          <div>
            <i style={{ width: `${Math.max(4, item.percent)}%` }} />
          </div>
          <em>{item.count}</em>
        </div>
      ))}
      {profile.frequencies.length === 0 && <p>No values</p>}
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
  const [activeTab, setActiveTab] = useState<TabKey>("report");
  const [chartType, setChartType] = useState<ChartType>("scatter");
  const [visualizationMode, setVisualizationMode] = useState<VisualizationMode | "auto">("auto");
  const [globalSearch, setGlobalSearch] = useState("");
  const [xColumn, setXColumn] = useState("");
  const [yColumn, setYColumn] = useState("");
  const [groupColumn, setGroupColumn] = useState("");
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [dragging, setDragging] = useState(false);

  const columns = useMemo(() => getColumns(rows), [rows]);
  const surveyDetections = useMemo(() => detectSurveyVariables(columns), [columns]);
  const surveyDetectionList = useMemo(() => (Object.keys(SURVEY_CONCEPTS) as SurveyVariableKey[]).map((concept) => surveyDetections[concept]), [surveyDetections]);
  const surveyVariables = useMemo(() => getSurveyVariables(surveyDetections), [surveyDetections]);
  const numericColumns = useMemo(() => columns.filter((column) => isNumericColumn(rows, column)), [columns, rows]);
  const textColumns = useMemo(() => columns.filter((column) => !numericColumns.includes(column)), [columns, numericColumns]);
  const columnProfiles = useMemo(() => getColumnProfiles(rows, columns), [columns, rows]);
  const dataQuality = useMemo(() => getDataQuality(rows, columns, columnProfiles), [columnProfiles, columns, rows]);
  const searchedRows = useMemo(
    () => rows.filter((row) => rowMatchesSearch(row, columns, globalSearch)),
    [columns, globalSearch, rows],
  );

  const filteredRows = useMemo(() => {
    const nextRows = searchedRows.filter((row) =>
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
  }, [columns, filters, searchedRows, sort]);

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

  const notableProfiles = useMemo(
    () =>
      columnProfiles
        .filter((profile) => profile.type === "numeric" && profile.average !== null)
        .slice(0, 8),
    [columnProfiles],
  );

  const missingProfiles = useMemo(
    () => columnProfiles.filter((profile) => profile.missing > 0).sort((a, b) => b.missing - a.missing),
    [columnProfiles],
  );
  const typeCounts = useMemo(
    () =>
      columnProfiles.reduce(
        (counts, profile) => ({ ...counts, [profile.type]: counts[profile.type] + 1 }),
        { empty: 0, mixed: 0, numeric: 0, text: 0 },
      ),
    [columnProfiles],
  );
  const suspiciousColumns = useMemo(
    () =>
      columnProfiles
        .map((profile) => {
          const reasons = [
            profile.missing > 0 && profile.completeness < 90 ? `${profile.missing.toLocaleString()} missing values` : "",
            profile.unique <= 1 && rows.length > 1 ? "only one observed value" : "",
            profile.type === "mixed" ? "mixed numeric/text values" : "",
            profile.unique === rows.length && rows.length > 8 ? "unique in every row; likely an identifier" : "",
            profile.column === surveyVariables.age && profile.min !== null && profile.min < 0 ? "age contains negative values" : "",
            profile.column === surveyVariables.age && profile.max !== null && profile.max > 120 ? "age contains values above 120" : "",
            profile.column === surveyVariables.income && profile.min !== null && profile.min < 0 ? "income contains negative values" : "",
            profile.column === surveyVariables.household && profile.min !== null && profile.min < 1 ? "household size below 1" : "",
            profile.column === surveyVariables.children && profile.min !== null && profile.min < 0 ? "children/dependents contains negative values" : "",
          ].filter(Boolean);

          return { ...profile, reasons };
        })
        .filter((profile) => profile.reasons.length > 0)
        .sort((a, b) => a.completeness - b.completeness)
        .slice(0, 8),
    [columnProfiles, rows.length, surveyVariables],
  );
  const recommendedAnalyses = useMemo(() => {
    const recommendations = [
      surveyVariables.age || surveyVariables.sex || surveyVariables.household
        ? "Population Profile: inspect age structure, sex/gender composition, household size, dependents, and geography."
        : "",
      surveyVariables.income
        ? "Income Analysis: compare income distribution, medians, ranges, and group averages. Check missingness before interpreting inequality."
        : "",
      surveyVariables.employment
        ? "Labour Analysis: describe employment status and relate labour-market groups to education or income."
        : "",
      correlations.length > 0
        ? `Correlation Analysis: review the strongest numeric relationship, ${correlations[0].left} x ${correlations[0].right} (${correlations[0].value.toFixed(3)}).`
        : numericColumns.length >= 2
          ? "Correlation Analysis: compare numeric variables to identify relationships worth investigating."
          : "",
      numericColumns.length > 0
        ? "Visualization Opportunities: use histograms for distributions, bar charts for group comparisons, and scatterplots for relationships."
        : "",
      surveyDetections.year.column || surveyDetections.quarter.column
        ? "If this file combines periods, compare results by year or quarter before pooling observations."
        : "",
      dataQuality.completeness < 95
        ? "Review columns with missing values in Data Quality before estimating relationships or reporting totals."
        : "",
      dataQuality.duplicateRows > 0
        ? "Investigate duplicate rows before analysis; duplicates may inflate subgroup counts."
        : "",
    ].filter(Boolean);

    return recommendations.length
      ? recommendations
      : ["Inspect Data Quality first, then use Visualization to choose relevant numeric and categorical variables for exploratory analysis."];
  }, [correlations, dataQuality.completeness, dataQuality.duplicateRows, numericColumns.length, surveyDetections.quarter.column, surveyDetections.year.column, surveyVariables]);

  const executiveSummary = useMemo(() => {
    const detectedCore = [
      surveyVariables.age ? "age" : "",
      surveyVariables.sex ? "sex/gender" : "",
      surveyVariables.income ? "income" : "",
      surveyVariables.employment ? "employment status" : "",
      surveyVariables.region ? "region/geography" : "",
    ].filter(Boolean);
    const qualitySentence =
      missingProfiles.length > 0
        ? `${missingProfiles.slice(0, 2).length === 1 ? "One variable has" : "Two variables have"} substantial missingness and should be reviewed before analysis.`
        : "No missing values were detected in the loaded cells.";
    const suitability =
      surveyVariables.age || surveyVariables.sex || surveyVariables.income || surveyVariables.employment
        ? "The dataset appears suitable for demographic and social survey analysis."
        : "The dataset can be explored, but core survey concepts were only partially detected.";

    return `This dataset contains ${rows.length.toLocaleString()} observations and ${columns.length.toLocaleString()} variables. Completeness is ${dataQuality.completeness.toFixed(1)}%. ${
      detectedCore.length ? `${detectedCore.join(", ")} ${detectedCore.length === 1 ? "was" : "were"} successfully detected.` : "No core demographic concepts were confidently detected."
    } ${suitability} ${qualitySentence}`;
  }, [columns.length, dataQuality.completeness, missingProfiles, rows.length, surveyVariables]);

  const selectedX = xColumn || numericColumns[0] || "";
  const selectedY = yColumn || numericColumns.find((column) => column !== selectedX) || numericColumns[0] || "";
  const ageSummary = useMemo(() => numericSummary(filteredRows, surveyVariables.age), [filteredRows, surveyVariables.age]);
  const householdSummary = useMemo(() => numericSummary(filteredRows, surveyVariables.household), [filteredRows, surveyVariables.household]);
  const incomeSummary = useMemo(() => numericSummary(filteredRows, surveyVariables.income), [filteredRows, surveyVariables.income]);
  const incomePercentiles = useMemo(() => getPercentiles(filteredRows, surveyVariables.income), [filteredRows, surveyVariables.income]);
  const ageBands = useMemo(() => ageDistribution(filteredRows, surveyVariables.age), [filteredRows, surveyVariables.age]);
  const pyramidRows = useMemo(() => populationPyramid(filteredRows, surveyVariables.age, surveyVariables.sex), [filteredRows, surveyVariables.age, surveyVariables.sex]);
  const sexDistribution = useMemo(() => frequency(filteredRows, surveyVariables.sex), [filteredRows, surveyVariables.sex]);
  const employmentDistribution = useMemo(() => frequency(filteredRows, surveyVariables.employment), [filteredRows, surveyVariables.employment]);
  const householdDistribution = useMemo(() => frequency(filteredRows, surveyVariables.household), [filteredRows, surveyVariables.household]);
  const occupationDistribution = useMemo(() => frequency(filteredRows, surveyVariables.occupation), [filteredRows, surveyVariables.occupation]);
  const educationDistribution = useMemo(() => frequency(filteredRows, surveyVariables.education), [filteredRows, surveyVariables.education]);
  const regionDistribution = useMemo(() => frequency(filteredRows, surveyVariables.region), [filteredRows, surveyVariables.region]);
  const incomeBySex = useMemo(() => averageByGroup(filteredRows, surveyVariables.sex, surveyVariables.income), [filteredRows, surveyVariables.income, surveyVariables.sex]);
  const incomeByEmployment = useMemo(
    () => averageByGroup(filteredRows, surveyVariables.employment, surveyVariables.income),
    [filteredRows, surveyVariables.employment, surveyVariables.income],
  );
  const topBottomIncome = useMemo(() => {
    const values = surveyVariables.income ? getNumericValues(filteredRows, surveyVariables.income).sort((a, b) => a - b) : [];
    const groupSize = Math.max(1, Math.ceil(values.length * 0.2));
    const bottom = values.slice(0, groupSize);
    const top = values.slice(-groupSize);

    return {
      bottomAverage: bottom.length ? mean(bottom) : null,
      topAverage: top.length ? mean(top) : null,
    };
  }, [filteredRows, surveyVariables.income]);
  const overviewInterpretation = useMemo(() => {
    const detectedCount = surveyDetectionList.filter((detection) => detection.column).length;
    const quality = dataQuality.completeness >= 95 ? "high completeness" : dataQuality.completeness >= 85 ? "moderate completeness" : "notable missingness";
    return `The current filtered view contains ${filteredRows.length.toLocaleString()} records across ${columns.length.toLocaleString()} variables. ${detectedCount.toLocaleString()} survey concepts were detected, and the dataset has ${quality}, so it is ready for a first-pass profile before deeper modelling.`;
  }, [columns.length, dataQuality.completeness, filteredRows.length, surveyDetectionList]);
  const populationInterpretation = useMemo(() => {
    if (!ageSummary && !sexDistribution.length && !householdDistribution.length) {
      return "Population interpretation is limited because age, sex/gender, and household variables were not detected. Rename or map these fields to unlock demographic profiling.";
    }

    const ageText = ageSummary
      ? ageSummary.median !== null && ageSummary.median < 35
        ? `The sample is relatively young, with a median age of ${formatValue(ageSummary.median)}.`
        : ageSummary.median !== null && ageSummary.median > 55
          ? `The sample skews older, with a median age of ${formatValue(ageSummary.median)}.`
          : `The sample has a balanced adult age structure, with a median age of ${formatValue(ageSummary?.median)}.`
      : "Age structure could not be interpreted because no age variable was detected.";
    const sexText = sexDistribution[0] ? `The largest sex/gender group is ${sexDistribution[0].label}, representing ${sexDistribution[0].percent.toFixed(0)}% of analysed records.` : "";
    const householdText = householdSummary ? `Average household size is ${formatValue(householdSummary.average)}, which helps contextualize dependents and living arrangements.` : "";

    return [ageText, sexText, householdText].filter(Boolean).join(" ");
  }, [ageSummary, householdDistribution.length, householdSummary, sexDistribution]);
  const incomeInterpretation = useMemo(() => {
    if (!incomeSummary) {
      return "Income interpretation is unavailable until an individual or household income variable is detected.";
    }

    const skewText =
      incomeSummary.average !== null && incomeSummary.median !== null && incomeSummary.average > incomeSummary.median * 1.25
        ? "The income distribution is right-skewed, with a smaller group earning substantially above the median."
        : "The mean and median are relatively close, suggesting the income distribution is not extremely skewed in the filtered data.";
    const spreadText =
      incomePercentiles.p90 !== null && incomePercentiles.p10 !== null
        ? `The distance between P10 (${formatValue(incomePercentiles.p10)}) and P90 (${formatValue(incomePercentiles.p90)}) gives a quick view of income spread.`
        : "";

    return `${skewText} Median income is ${formatValue(incomeSummary.median)}. ${spreadText}`;
  }, [incomePercentiles.p10, incomePercentiles.p90, incomeSummary]);
  const labourInterpretation = useMemo(() => {
    if (!employmentDistribution.length) {
      return "Labour interpretation is limited because no employment-status variable was detected.";
    }

    const occupationText = occupationDistribution.length
      ? `Occupation is distributed across ${occupationDistribution.length.toLocaleString()} categories, led by ${occupationDistribution[0].label}.`
      : "Occupation was not detected, so the labour profile focuses on employment status only.";

    return `Employment is concentrated in ${employmentDistribution[0].label}, which accounts for ${employmentDistribution[0].percent.toFixed(0)}% of analysed records. ${occupationText}`;
  }, [employmentDistribution, occupationDistribution]);
  const dataQualityInterpretation = useMemo(() => {
    const missingWarning = missingProfiles.filter((profile) => profile.completeness < 90);
    const duplicateText =
      dataQuality.duplicateRows > 0
        ? `${dataQuality.duplicateRows.toLocaleString()} duplicate row${dataQuality.duplicateRows === 1 ? "" : "s"} should be reviewed before estimating counts or shares.`
        : "No duplicate rows were detected.";
    const missingText =
      missingWarning.length > 0
        ? `${missingWarning.length.toLocaleString()} variable${missingWarning.length === 1 ? " contains" : "s contain"} more than 10% missing values and should be reviewed before analysis.`
        : "No variables exceed the 10% missingness warning threshold.";

    return `${missingText} ${duplicateText}`;
  }, [dataQuality.duplicateRows, missingProfiles]);
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
  const customChartReady = numericColumns.length >= (chartType === "histogram" ? 1 : 2) && filteredRows.length > 0;
  const customVisualizationInterpretation = useMemo(() => {
    if (!customChartReady) {
      return "Choose numeric variables to generate a custom relationship chart.";
    }

    if (chartType === "histogram") {
      return `This distribution view helps assess ${selectedX}, including concentration, spread, and potential outliers.`;
    }

    if (chartType === "scatter" && regression) {
      return `The relationship view compares ${selectedX} and ${selectedY}. The linear trend is ${regression.slope >= 0 ? "positive" : "negative"}, with R² of ${regression.r2.toFixed(2)}.`;
    }

    return `This custom view compares ${selectedX} and ${selectedY} across the current filtered dataset. Use grouping to check whether patterns differ by region, cohort, or social category.`;
  }, [chartType, customChartReady, regression, selectedX, selectedY]);
  const recommendedVisualizations = useMemo(
    () => [
      {
        detail: "Best first view when both age and sex/gender are detected.",
        enabled: Boolean(surveyVariables.age && surveyVariables.sex),
        key: "populationPyramid" as VisualizationMode,
        title: "Population Pyramid",
      },
      {
        detail: "Understand whether the sample is young, older, or balanced across age bands.",
        enabled: Boolean(surveyVariables.age),
        key: "ageDistribution" as VisualizationMode,
        title: "Age Distribution",
      },
      {
        detail: "Inspect skew, spread, and outliers in detected income variables.",
        enabled: Boolean(surveyVariables.income),
        key: "incomeDistribution" as VisualizationMode,
        title: "Income Distribution",
      },
      {
        detail: "Summarize employment, unemployment, student, retired, or inactive groups.",
        enabled: Boolean(surveyVariables.employment),
        key: "employmentStructure" as VisualizationMode,
        title: "Employment Structure",
      },
      {
        detail: "Profile educational attainment categories in the survey population.",
        enabled: Boolean(surveyVariables.education),
        key: "educationProfile" as VisualizationMode,
        title: "Education Profile",
      },
      {
        detail: "Compare regions by income when available, otherwise by sample size.",
        enabled: Boolean(surveyVariables.region),
        key: "regionalComparison" as VisualizationMode,
        title: "Regional Comparison",
      },
      {
        detail: "Explore a custom relationship between selected numeric variables.",
        enabled: customChartReady,
        key: "customRelationship" as VisualizationMode,
        title: "Custom Relationship",
      },
    ],
    [customChartReady, surveyVariables.age, surveyVariables.education, surveyVariables.employment, surveyVariables.income, surveyVariables.region, surveyVariables.sex],
  );
  const activeVisualizationMode = useMemo(() => {
    if (visualizationMode !== "auto") {
      return visualizationMode;
    }

    return recommendedVisualizations.find((item) => item.enabled)?.key ?? "customRelationship";
  }, [recommendedVisualizations, visualizationMode]);
  const socialPlot = useMemo(() => {
    const baseLayout = {
      autosize: true,
      font: { color: "#334155", family: "Arial, sans-serif" },
      height: 460,
      legend: { font: { color: "#475569" }, orientation: "h", x: 0, y: 1.14 },
      margin: { b: 76, l: 72, r: 28, t: 44 },
      paper_bgcolor: "rgba(0,0,0,0)",
      plot_bgcolor: "#f8fafc",
    };

    if (activeVisualizationMode === "populationPyramid") {
      const ready = pyramidRows.some((row) => row.left.count || row.right.count);
      return {
        data: [
          {
            marker: { color: "#60a5fa" },
            name: pyramidRows[0]?.left.label ?? "Group A",
            orientation: "h",
            type: "bar",
            x: pyramidRows.map((row) => -row.left.count),
            y: pyramidRows.map((row) => row.band),
          },
          {
            marker: { color: "#a78bfa" },
            name: pyramidRows[0]?.right.label ?? "Group B",
            orientation: "h",
            type: "bar",
            x: pyramidRows.map((row) => row.right.count),
            y: pyramidRows.map((row) => row.band),
          },
        ] as PlotTrace[],
        interpretation: ready ? populationInterpretation : "Detect age and sex/gender variables to build a population pyramid.",
        layout: {
          ...baseLayout,
          barmode: "relative",
          title: { font: { color: "#0f172a", size: 17 }, text: "Population Pyramid" },
          xaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: "Respondents", zeroline: true, zerolinecolor: "#94a3b8" },
          yaxis: { autorange: "reversed", gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: "Age band" },
        },
        nextAnalyses: ["Compare income by age band", "Review household structure by life stage", "Check whether age-sex cells are weighted or unweighted"],
        ready,
        title: "Population Pyramid",
      };
    }

    if (activeVisualizationMode === "ageDistribution") {
      return {
        data: [{ marker: { color: "#2563eb" }, type: "bar", x: ageBands.map((item) => item.label), y: ageBands.map((item) => item.count) }] as PlotTrace[],
        interpretation: ageBands.length ? populationInterpretation : "Detect an age variable to explain the age structure.",
        layout: {
          ...baseLayout,
          title: { font: { color: "#0f172a", size: 17 }, text: "Age Distribution" },
          xaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: "Age band" },
          yaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: "Respondents" },
        },
        nextAnalyses: ["Compare age bands by sex/gender", "Inspect employment status by age", "Review whether children or older adults are represented"],
        ready: ageBands.length > 0,
        title: "Age Distribution",
      };
    }

    if (activeVisualizationMode === "incomeDistribution") {
      const values = getNumericValues(filteredRows, surveyVariables.income);
      return {
        data: [{ marker: { color: "#0f766e" }, type: "histogram", x: values }] as PlotTrace[],
        interpretation: values.length ? incomeInterpretation : "Detect an income variable to interpret the income distribution.",
        layout: {
          ...baseLayout,
          bargap: 0.08,
          title: { font: { color: "#0f172a", size: 17 }, text: "Income Distribution" },
          xaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: surveyVariables.income || "Income" },
          yaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: "Respondents" },
        },
        nextAnalyses: ["Compare median income by sex/gender", "Review top and bottom income groups", "Check whether zero or missing income values need recoding"],
        ready: values.length > 0,
        title: "Income Distribution",
      };
    }

    if (activeVisualizationMode === "employmentStructure") {
      return {
        data: [{ marker: { color: "#7c3aed" }, type: "bar", x: employmentDistribution.map((item) => item.label), y: employmentDistribution.map((item) => item.count) }] as PlotTrace[],
        interpretation: employmentDistribution.length ? labourInterpretation : "Detect an employment-status variable to explain labour-market structure.",
        layout: {
          ...baseLayout,
          title: { font: { color: "#0f172a", size: 17 }, text: "Employment Structure" },
          xaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: "Employment status" },
          yaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: "Respondents" },
        },
        nextAnalyses: ["Compare income by labour status", "Explore education within employment groups", "Review inactive or missing labour categories"],
        ready: employmentDistribution.length > 0,
        title: "Employment Structure",
      };
    }

    if (activeVisualizationMode === "educationProfile") {
      return {
        data: [{ marker: { color: "#b45309" }, type: "bar", x: educationDistribution.map((item) => item.label), y: educationDistribution.map((item) => item.count) }] as PlotTrace[],
        interpretation: educationDistribution[0]
          ? `Education is concentrated in ${educationDistribution[0].label}, representing ${educationDistribution[0].percent.toFixed(0)}% of analysed records.`
          : "Detect an education variable to interpret attainment patterns.",
        layout: {
          ...baseLayout,
          title: { font: { color: "#0f172a", size: 17 }, text: "Education Profile" },
          xaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: "Education level" },
          yaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: "Respondents" },
        },
        nextAnalyses: ["Compare education by age cohort", "Relate education to income", "Check whether education codes need labels"],
        ready: educationDistribution.length > 0,
        title: "Education Profile",
      };
    }

    if (activeVisualizationMode === "regionalComparison") {
      const incomeByRegion = averageByGroup(filteredRows, surveyVariables.region, surveyVariables.income);
      const regionalItems = incomeByRegion.length ? incomeByRegion.map((item) => ({ label: item.label, value: item.average })) : regionDistribution.map((item) => ({ label: item.label, value: item.count }));
      return {
        data: [{ marker: { color: "#475569" }, type: "bar", x: regionalItems.map((item) => item.label), y: regionalItems.map((item) => item.value) }] as PlotTrace[],
        interpretation: regionalItems[0]
          ? incomeByRegion.length
            ? `${regionalItems[0].label} has the highest observed average income among detected regions.`
            : `${regionalItems[0].label} is the largest regional group in the filtered dataset.`
          : "Detect a region or geography variable to compare places.",
        layout: {
          ...baseLayout,
          title: { font: { color: "#0f172a", size: 17 }, text: incomeByRegion.length ? "Regional Income Comparison" : "Regional Sample Comparison" },
          xaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: surveyVariables.region || "Region" },
          yaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: incomeByRegion.length ? "Average income" : "Respondents" },
        },
        nextAnalyses: ["Compare demographic structure by region", "Check sample sizes before interpreting regional differences", "Explore labour status by geography"],
        ready: regionalItems.length > 0,
        title: incomeByRegion.length ? "Regional Income Comparison" : "Regional Sample Comparison",
      };
    }

    return {
      data: plotData,
      interpretation: customChartReady
        ? customVisualizationInterpretation
        : "Choose numeric variables to generate a custom relationship chart.",
      layout: {
        ...baseLayout,
        bargap: 0.18,
        barmode: groupColumn ? "group" : "overlay",
        title: { font: { color: "#0f172a", size: 17 }, text: "Custom Relationship" },
        xaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: selectedX, zerolinecolor: "#cbd5e1" },
        yaxis: { gridcolor: "#e2e8f0", linecolor: "#cbd5e1", title: chartType === "histogram" ? "Count" : selectedY, zerolinecolor: "#cbd5e1" },
      },
      nextAnalyses: ["Try grouping by region or sex/gender", "Use histograms before interpreting outliers", "Move promising relationships into the Correlation or Insights views"],
      ready: customChartReady,
      title: "Custom Relationship",
    };
  }, [
    activeVisualizationMode,
    ageBands,
    chartType,
    customChartReady,
    educationDistribution,
    employmentDistribution,
    filteredRows,
    groupColumn,
    incomeInterpretation,
    labourInterpretation,
    plotData,
    populationInterpretation,
    pyramidRows,
    regionDistribution,
    selectedX,
    selectedY,
    surveyVariables.income,
    surveyVariables.region,
    customVisualizationInterpretation,
  ]);

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

      setActiveTab("report");
      setFilters({});
      setFiltersOpen(false);
      setGlobalSearch("");
      setGroupColumn("");
      setSort(null);
      setVisualizationMode("auto");
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
    setFileName("sample-social-survey.csv");
    setActiveTab("report");
    setFilters({});
    setFiltersOpen(false);
    setGlobalSearch("");
    setGroupColumn("region");
    setSort(null);
    setVisualizationMode("auto");
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
    link.download = `social-data-export-${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  const hasData = rows.length > 0;

  return (
    <main className="relative min-h-screen overflow-hidden px-3 py-6 text-slate-100 sm:px-6 lg:px-10">
      <div className="cosmic-grid" />
      <div className="star-field" />
      <div className="mx-auto flex w-full max-w-7xl flex-col gap-5 sm:gap-6">
        <header className="flex flex-col items-center gap-3 py-4 text-center sm:py-5">
          <div className="rounded-full border border-blue-200/25 bg-white/10 px-4 py-2 text-xs font-semibold uppercase tracking-[0.24em] text-blue-100">
            Social Data Studio
          </div>
          <div className="max-w-4xl">
            <h1 className="text-3xl font-semibold tracking-normal text-white sm:text-5xl">Social Data Explorer</h1>
            <p className="mt-3 text-base leading-7 text-slate-300 sm:text-lg">
              Explore, profile and understand social survey datasets.
            </p>
          </div>
        </header>

        <GlassPanel className="mx-auto w-full max-w-4xl p-4 sm:p-6">
          <div
            className={`flex min-h-40 flex-col items-center justify-center rounded-lg border border-dashed p-5 text-center transition sm:min-h-48 sm:p-6 ${
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
            <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-full border border-blue-200 bg-blue-50 text-2xl text-blue-700 sm:h-14 sm:w-14">
              ↑
            </div>
            <h2 className="text-2xl font-semibold text-white">Drop a survey CSV or Excel file</h2>
            <p className="mt-2 max-w-xl text-sm leading-6 text-slate-300">
              Data stays in your browser. Variables like age, sex, income, household size, employment, education, and region are detected automatically.
            </p>
            <div className="mt-5 flex flex-wrap items-center justify-center gap-3">
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
            detail="Use the uploader above or load the sample social survey to unlock demographic profiles, income and labour analysis, data quality checks, insights, charts, preview, and export tools."
          />
        )}

        {hasData && (
          <>
            <GlassPanel className="p-4 sm:p-5">
              <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
                <div>
                  <h2 className="text-lg font-semibold text-white">Dataset search</h2>
                  <p className="mt-1 text-sm text-slate-400">
                    Search across every survey variable before profiles, insights, charts, preview, and export are calculated.
                  </p>
                </div>
                <div className="w-full lg:max-w-md">
                  <input
                    className="search-input"
                    onChange={(event) => setGlobalSearch(event.target.value)}
                    placeholder="Search countries, regions, years, values..."
                    value={globalSearch}
                  />
                  <p className="mt-2 text-xs text-slate-500">
                    {searchedRows.length.toLocaleString()} of {rows.length.toLocaleString()} rows match search.
                  </p>
                </div>
              </div>
            </GlassPanel>

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

            <GlassPanel className="p-4 sm:p-5">
              <div className="flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
                <div>
                  <h2 className="text-lg font-semibold text-white">Detected survey variables</h2>
                  <p className="mt-1 text-sm text-slate-400">
                    Semantic detection combines EPH aliases, exact aliases, international survey names, partial matching, and keyword matching.
                  </p>
                </div>
                <p className="text-sm font-semibold text-blue-700">
                  {surveyDetectionList.filter((detection) => detection.column).length} / {surveyDetectionList.length} concepts detected
                </p>
              </div>
              <div className="detected-grid mt-5">
                {surveyDetectionList.map((detection) => (
                  <div className={`detected-variable ${detection.column ? "" : "detected-variable-empty"}`} key={detection.concept}>
                    <div>
                      <span>{detection.friendlyLabel}</span>
                      <strong>{detection.column || "Not detected"}</strong>
                    </div>
                    <div className="text-right">
                      <span>{detection.method}</span>
                      <strong>{detection.confidence}%</strong>
                    </div>
                  </div>
                ))}
              </div>
            </GlassPanel>

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

            {activeTab === "report" && (
              <div className="dataset-report space-y-5">
                <GlassPanel className="report-hero p-5">
                  <div className="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between">
                    <div>
                      <p className="report-kicker">Flagship pre-analysis report</p>
                      <h2 className="text-2xl font-semibold text-white">Dataset Report</h2>
                      <p className="mt-1 max-w-3xl text-sm leading-6 text-slate-400">
                        Understand the dataset before analysing it. Inspired by dfSummary(), adapted into a modern research dashboard for survey and demographic datasets.
                      </p>
                    </div>
                    <div className="rounded-md border border-blue-200 bg-blue-50 px-4 py-3 text-sm font-semibold text-blue-900">
                      {fileName || "Untitled dataset"}
                    </div>
                  </div>
                </GlassPanel>

                <GlassPanel className="p-5">
                  <p className="report-section-label">Section 6</p>
                  <h3 className="text-lg font-semibold text-white">Executive Summary</h3>
                  <p className="mt-3 max-w-5xl text-base leading-7 text-slate-700">{executiveSummary}</p>
                </GlassPanel>

                <GlassPanel className="p-5">
                  <p className="report-section-label">Research Assistant</p>
                  <h3 className="text-lg font-semibold text-white">Suggested Analyses</h3>
                  <div className="suggested-grid mt-4">
                    <button onClick={() => setActiveTab("income")} type="button">Compare income by sex</button>
                    <button onClick={() => setActiveTab("population")} type="button">Analyse age structure</button>
                    <button onClick={() => setActiveTab("labour")} type="button">Explore labour categories</button>
                    <button onClick={() => setActiveTab("quality")} type="button">Review missing values</button>
                  </div>
                </GlassPanel>

                <GlassPanel className="p-5">
                  <p className="report-section-label">Section 1</p>
                  <h3 className="text-lg font-semibold text-white">Dataset Overview</h3>
                  <div className="mt-5 grid gap-3 sm:grid-cols-2 lg:grid-cols-6">
                    <div className="metric-tile">
                      <span>Rows</span>
                      <strong>{rows.length.toLocaleString()}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Columns</span>
                      <strong>{columns.length.toLocaleString()}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Total cells</span>
                      <strong>{dataQuality.totalCells.toLocaleString()}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Missing values</span>
                      <strong>{dataQuality.missingValues.toLocaleString()}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Duplicate rows</span>
                      <strong>{dataQuality.duplicateRows.toLocaleString()}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Completeness</span>
                      <strong>{dataQuality.completeness.toFixed(1)}%</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Numeric vars</span>
                      <strong>{typeCounts.numeric}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Categorical vars</span>
                      <strong>{typeCounts.text}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Mixed vars</span>
                      <strong>{typeCounts.mixed}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Empty vars</span>
                      <strong>{typeCounts.empty}</strong>
                    </div>
                  </div>
                  <div className="quality-bar mt-5">
                    <div style={{ width: `${dataQuality.completeness}%` }} />
                  </div>
                </GlassPanel>

                <div className="grid gap-5 xl:grid-cols-[0.8fr_1.2fr]">
                  <GlassPanel className="p-5">
                    <p className="report-section-label">Section 2</p>
                    <h3 className="text-lg font-semibold text-white">Variable Inventory</h3>
                    <div className="mt-4 grid gap-3 sm:grid-cols-2">
                      <div className="metric-tile">
                        <span>Numeric</span>
                        <strong>{typeCounts.numeric}</strong>
                      </div>
                      <div className="metric-tile">
                        <span>Categorical</span>
                        <strong>{typeCounts.text}</strong>
                      </div>
                      <div className="metric-tile">
                        <span>Mixed</span>
                        <strong>{typeCounts.mixed}</strong>
                      </div>
                      <div className="metric-tile">
                        <span>Empty</span>
                        <strong>{typeCounts.empty}</strong>
                      </div>
                    </div>
                  </GlassPanel>

                  <GlassPanel className="p-5">
                    <p className="report-section-label">Section 2</p>
                    <h3 className="text-lg font-semibold text-white">Detected Survey Variables</h3>
                    <div className="report-detection-groups mt-4">
                      {REPORT_DETECTION_GROUPS.map((group) => (
                        <div className="report-detection-group" key={group.title}>
                          <h4>{group.title}</h4>
                          <div className="report-variable-list mt-3">
                            {group.concepts.map((concept) => {
                              const detection = surveyDetections[concept];
                              return (
                                <div key={concept}>
                                  <span>{detection.friendlyLabel}</span>
                                  <strong>{detection.column || "Not detected"}</strong>
                                  <em>{detection.confidence}% / {detection.method}</em>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      ))}
                    </div>
                  </GlassPanel>
                </div>

                <GlassPanel className="p-5">
                  <p className="report-section-label">Section 3</p>
                  <h3 className="text-lg font-semibold text-white">Variable Profiles</h3>
                  <p className="mt-1 text-sm text-slate-400">
                    A compact profile for each variable, including completeness, unique values, numeric statistics, and mini distributions.
                  </p>
                  <div className="variable-profile-grid mt-5">
                    {columnProfiles.map((profile) => (
                      <div className="variable-profile" key={profile.column}>
                        <div className="flex items-start justify-between gap-3">
                          <div className="min-w-0">
                            <h4 title={profile.column}>{profile.column}</h4>
                            <p>{profile.type} / {profile.unique.toLocaleString()} unique</p>
                          </div>
                          <span>{profile.completeness.toFixed(0)}%</span>
                        </div>
                        <div className="variable-profile-stats">
                          <div>
                            <span>Missing</span>
                            <strong>{profile.missing.toLocaleString()}</strong>
                          </div>
                          <div>
                            <span>Min</span>
                            <strong>{formatValue(profile.min)}</strong>
                          </div>
                          <div>
                            <span>Max</span>
                            <strong>{formatValue(profile.max)}</strong>
                          </div>
                          <div>
                            <span>Avg</span>
                            <strong>{formatValue(profile.average)}</strong>
                          </div>
                          <div>
                            <span>Median</span>
                            <strong>{formatValue(profile.median)}</strong>
                          </div>
                        </div>
                        <TinyDistribution profile={profile} />
                      </div>
                    ))}
                  </div>
                </GlassPanel>

                <div className="grid gap-5 xl:grid-cols-2">
                  <GlassPanel className="p-5">
                    <p className="report-section-label">Section 4</p>
                    <h3 className="text-lg font-semibold text-white">Data Quality Warnings</h3>
                    <p className="mt-1 text-sm text-slate-400">Columns with the highest missingness and other checks that may affect interpretation.</p>
                    <div className="mt-4 space-y-3">
                      {missingProfiles.length ? (
                        missingProfiles.slice(0, 8).map((profile) => (
                          <InsightCard
                            detail={`${profile.missing.toLocaleString()} missing values; ${profile.completeness.toFixed(1)}% complete.`}
                            key={profile.column}
                            label={profile.column}
                            tone={profile.completeness < 90 ? "warning" : "default"}
                          />
                        ))
                      ) : (
                        <InsightCard detail="No missing values detected across loaded cells." label="Missingness" />
                      )}
                    </div>
                  </GlassPanel>

                  <GlassPanel className="p-5">
                    <p className="report-section-label">Section 4</p>
                    <h3 className="text-lg font-semibold text-white">Suspicious columns</h3>
                    <div className="mt-4 grid gap-3 sm:grid-cols-3">
                      <InsightCard
                        detail={
                          dataQuality.duplicateRows > 0
                            ? `${dataQuality.duplicateRows.toLocaleString()} duplicate row${dataQuality.duplicateRows === 1 ? "" : "s"} detected.`
                            : "No duplicate rows detected."
                        }
                        label="Duplicates"
                        tone={dataQuality.duplicateRows > 0 ? "warning" : "default"}
                      />
                      <InsightCard
                        detail={
                          typeCounts.mixed > 0
                            ? `${typeCounts.mixed.toLocaleString()} mixed-type variable${typeCounts.mixed === 1 ? "" : "s"} detected.`
                            : "No mixed-type variables detected."
                        }
                        label="Mixed values"
                        tone={typeCounts.mixed > 0 ? "warning" : "default"}
                      />
                      <InsightCard
                        detail={
                          suspiciousColumns.some((profile) => profile.reasons.some((reason) => reason.includes("identifier")))
                            ? "Some columns are unique in every row and may be identifiers."
                            : "No likely identifier columns flagged by uniqueness checks."
                        }
                        label="Potential IDs"
                      />
                    </div>
                    <div className="mt-4 space-y-3">
                      {suspiciousColumns.length ? (
                        suspiciousColumns.map((profile) => (
                          <InsightCard
                            detail={profile.reasons.join("; ")}
                            key={profile.column}
                            label={profile.column}
                            tone="warning"
                          />
                        ))
                      ) : (
                        <InsightCard detail="No obvious structural issues detected from missingness, uniqueness, or mixed values." label="Column checks" />
                      )}
                    </div>
                  </GlassPanel>
                </div>

                <GlassPanel className="p-5">
                  <p className="report-section-label">Section 5</p>
                  <h3 className="text-lg font-semibold text-white">Recommended Analyses</h3>
                  <div className="mt-4 grid gap-3 lg:grid-cols-2">
                    {recommendedAnalyses.map((recommendation) => (
                      <InsightCard detail={recommendation} key={recommendation} label="Next step" />
                    ))}
                  </div>
                </GlassPanel>
              </div>
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

                <div className="grid gap-4 lg:grid-cols-[1.1fr_0.9fr]">
                  <GlassPanel className="p-5">
                    <h2 className="text-xl font-semibold text-white">Dashboard summary</h2>
                    <div className="mt-4">
                      <InsightCard detail={overviewInterpretation} label="Interpretation" />
                    </div>
                    <div className="mt-4 grid gap-3 sm:grid-cols-2">
                      <InsightCard
                        detail={
                          ageSummary
                            ? `This looks like respondent-level data with an average age of ${formatValue(ageSummary.average)} and ${filteredRows.length.toLocaleString()} analysed records.`
                            : `This dataset has ${filteredRows.length.toLocaleString()} analysed records. Add an age variable to unlock a fuller population profile.`
                        }
                        label="Survey profile"
                      />
                      <InsightCard
                        detail={
                          incomeSummary
                            ? `Average observed income is ${formatValue(incomeSummary.average)} with a median of ${formatValue(incomeSummary.median)}.`
                            : "No income variable was detected. Name income-like columns with terms such as income, earnings, wage, salary, or pay."
                        }
                        label="Income profile"
                      />
                    </div>
                  </GlassPanel>

                  <GlassPanel className="p-5">
                    <div className="flex items-start justify-between gap-4">
                      <div>
                        <h2 className="text-xl font-semibold text-white">Data quality</h2>
                        <p className="mt-1 text-sm text-slate-400">Completeness across all loaded cells.</p>
                      </div>
                      <strong className="text-3xl text-white">{dataQuality.completeness.toFixed(0)}%</strong>
                    </div>
                    <div className="quality-bar mt-5">
                      <div style={{ width: `${dataQuality.completeness}%` }} />
                    </div>
                    <div className="mt-4 grid gap-3 sm:grid-cols-2">
                      <div className="metric-tile">
                        <span>Missing values</span>
                        <strong>{dataQuality.missingValues.toLocaleString()}</strong>
                      </div>
                      <div className="metric-tile">
                        <span>Duplicate rows</span>
                        <strong>{dataQuality.duplicateRows.toLocaleString()}</strong>
                      </div>
                    </div>
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

            {activeTab === "population" && (
              <div className="space-y-5">
                <GlassPanel className="p-5">
                  <h2 className="text-xl font-semibold text-white">Population profile</h2>
                  <p className="mt-1 text-sm text-slate-400">
                    The app scans survey variables for age, sex or gender, region, education, household size, and children to describe the population under study.
                  </p>
                  <div className="mt-4">
                    <InsightCard detail={populationInterpretation} label="Interpretation" />
                  </div>
                  <div className="mt-5 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                    <div className="metric-tile">
                      <span>Analysed records</span>
                      <strong>{filteredRows.length.toLocaleString()}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Average age</span>
                      <strong>{ageSummary ? formatValue(ageSummary.average) : "-"}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Median age</span>
                      <strong>{ageSummary ? formatValue(ageSummary.median) : "-"}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Avg household size</span>
                      <strong>{householdSummary ? formatValue(householdSummary.average) : "-"}</strong>
                    </div>
                  </div>
                </GlassPanel>

                <div className="grid gap-5 xl:grid-cols-2">
                  <GlassPanel className="p-5">
                    <h3 className="text-lg font-semibold text-white">Population Pyramid</h3>
                    <p className="mt-1 text-sm text-slate-400">Age bands split by detected sex/gender variable.</p>
                    <div className="population-pyramid mt-5">
                      {pyramidRows.some((row) => row.left.count || row.right.count) ? (
                        pyramidRows.map((row) => (
                          <div className="pyramid-row" key={row.band}>
                            <div className="pyramid-side left"><i style={{ width: `${row.left.width}%` }} /><span>{row.left.count || ""}</span></div>
                            <strong>{row.band}</strong>
                            <div className="pyramid-side right"><i style={{ width: `${row.right.width}%` }} /><span>{row.right.count || ""}</span></div>
                          </div>
                        ))
                      ) : (
                        <p className="text-sm text-slate-500">Detect age and sex/gender variables to render a pyramid.</p>
                      )}
                    </div>
                    {pyramidRows.length > 0 && (
                      <p className="mt-3 text-xs text-slate-500">
                        Left: {pyramidRows[0]?.left.label} / Right: {pyramidRows[0]?.right.label}
                      </p>
                    )}
                  </GlassPanel>
                  <FrequencyList items={ageBands} title={`Age structure${surveyVariables.age ? ` (${surveyVariables.age})` : ""}`} />
                  <FrequencyList items={sexDistribution} title={`Sex / gender${surveyVariables.sex ? ` (${surveyVariables.sex})` : ""}`} />
                  <FrequencyList items={regionDistribution} title={`Geography${surveyVariables.region ? ` (${surveyVariables.region})` : ""}`} />
                  <FrequencyList items={householdDistribution} title={`Household structure${surveyVariables.household ? ` (${surveyVariables.household})` : ""}`} />
                  <FrequencyList items={educationDistribution} title={`Education${surveyVariables.education ? ` (${surveyVariables.education})` : ""}`} />
                </div>
              </div>
            )}

            {activeTab === "income" && (
              <div className="space-y-5">
                <GlassPanel className="p-5">
                  <h2 className="text-xl font-semibold text-white">Income analysis</h2>
                  <p className="mt-1 text-sm text-slate-400">
                    Income variables are detected from labels such as income, earnings, wage, salary, or pay. Group comparisons use available sex/gender and labour variables.
                  </p>
                  <div className="mt-4">
                    <InsightCard detail={incomeInterpretation} label="Interpretation" />
                  </div>
                  <div className="mt-5 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                    <div className="metric-tile">
                      <span>Income variable</span>
                      <strong className="truncate text-base">{surveyVariables.income || "Not detected"}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Average income</span>
                      <strong>{incomeSummary ? formatValue(incomeSummary.average) : "-"}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Median income</span>
                      <strong>{incomeSummary ? formatValue(incomeSummary.median) : "-"}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Income range</span>
                      <strong className="text-base">{incomeSummary ? `${formatValue(incomeSummary.min)} - ${formatValue(incomeSummary.max)}` : "-"}</strong>
                    </div>
                  </div>
                </GlassPanel>

                <div className="grid gap-5 xl:grid-cols-[1fr_1fr]">
                  <GlassPanel className="p-5">
                    <h3 className="text-lg font-semibold text-white">Income Distribution</h3>
                    {surveyVariables.income ? (
                      <div className="mt-4">
                        <MiniBars values={getNumericValues(filteredRows, surveyVariables.income)} />
                      </div>
                    ) : (
                      <EmptyState title="No income distribution yet" detail="Detect an income variable to render a distribution." />
                    )}
                  </GlassPanel>
                  <GlassPanel className="p-5">
                    <h3 className="text-lg font-semibold text-white">Income Percentiles</h3>
                    <div className="mt-4 grid gap-3 sm:grid-cols-5">
                      {[
                        ["P10", incomePercentiles.p10],
                        ["P25", incomePercentiles.p25],
                        ["P50", incomePercentiles.p50],
                        ["P75", incomePercentiles.p75],
                        ["P90", incomePercentiles.p90],
                      ].map(([label, value]) => (
                        <div className="metric-tile" key={String(label)}>
                          <span>{label}</span>
                          <strong>{formatValue(value as CellValue)}</strong>
                        </div>
                      ))}
                    </div>
                    <div className="mt-4 grid gap-3 sm:grid-cols-2">
                      <InsightCard detail={`Bottom group average: ${formatValue(topBottomIncome.bottomAverage)}`} label="Bottom 20%" />
                      <InsightCard detail={`Top group average: ${formatValue(topBottomIncome.topAverage)}`} label="Top 20%" />
                    </div>
                  </GlassPanel>
                </div>

                <div className="grid gap-5 xl:grid-cols-2">
                  <GlassPanel className="p-5">
                    <h3 className="text-lg font-semibold text-white">Average income by sex / gender</h3>
                    <div className="mt-4 space-y-3">
                      {incomeBySex.length ? (
                        incomeBySex.map((item) => (
                          <InsightCard detail={`${item.count.toLocaleString()} records average ${formatValue(item.average)}.`} key={item.label} label={item.label} />
                        ))
                      ) : (
                        <EmptyState title="No group comparison yet" detail="Add both income and sex/gender variables to compare average income by group." />
                      )}
                    </div>
                  </GlassPanel>

                  <GlassPanel className="p-5">
                    <h3 className="text-lg font-semibold text-white">Average income by labour status</h3>
                    <div className="mt-4 space-y-3">
                      {incomeByEmployment.length ? (
                        incomeByEmployment.map((item) => (
                          <InsightCard detail={`${item.count.toLocaleString()} records average ${formatValue(item.average)}.`} key={item.label} label={item.label} />
                        ))
                      ) : (
                        <EmptyState title="No labour comparison yet" detail="Add income and employment/labour variables to compare earnings by labour-market status." />
                      )}
                    </div>
                  </GlassPanel>
                </div>
              </div>
            )}

            {activeTab === "labour" && (
              <div className="space-y-5">
                <GlassPanel className="p-5">
                  <h2 className="text-xl font-semibold text-white">Labour profile</h2>
                  <p className="mt-1 text-sm text-slate-400">
                    Labour variables are detected from employment, work status, labour, occupation, job status, and similar field names.
                  </p>
                  <div className="mt-4">
                    <InsightCard detail={labourInterpretation} label="Interpretation" />
                  </div>
                  <div className="mt-5 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                    <div className="metric-tile">
                      <span>Labour variable</span>
                      <strong className="truncate text-base">{surveyVariables.employment || "Not detected"}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Statuses</span>
                      <strong>{employmentDistribution.length || "-"}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Largest group</span>
                      <strong className="truncate text-base">{employmentDistribution[0]?.label ?? "-"}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Largest share</span>
                      <strong>{employmentDistribution[0] ? `${employmentDistribution[0].percent.toFixed(0)}%` : "-"}</strong>
                    </div>
                  </div>
                </GlassPanel>
                <div className="grid gap-5 xl:grid-cols-2">
                  <FrequencyList items={employmentDistribution} title="Employment status distribution" />
                  <FrequencyList items={occupationDistribution} title="Occupational categories" />
                  <FrequencyList items={educationDistribution} title="Education profile for labour context" />
                  <GlassPanel className="p-5">
                    <h3 className="text-lg font-semibold text-white">Labour Market Profile</h3>
                    <div className="mt-4 grid gap-3">
                      <InsightCard
                        detail={
                          employmentDistribution[0]
                            ? `${employmentDistribution[0].label} is the largest labour category (${employmentDistribution[0].percent.toFixed(0)}% of analysed records).`
                            : "No labour status variable detected."
                        }
                        label="Largest status"
                      />
                      <InsightCard
                        detail={
                          surveyVariables.occupation
                            ? `${occupationDistribution.length.toLocaleString()} occupational categories are represented.`
                            : "No occupation variable detected."
                        }
                        label="Occupation coverage"
                      />
                    </div>
                  </GlassPanel>
                </div>
              </div>
            )}

            {activeTab === "insights" && (
              <div className="space-y-5">
                <div className="grid gap-5 xl:grid-cols-[1fr_1fr]">
                  <GlassPanel className="p-5">
                    <h2 className="text-xl font-semibold text-white">Automatic insights</h2>
                    <p className="mt-1 text-sm text-slate-400">
                      Generated from the current search and filters where appropriate, with quality checks from the full loaded dataset.
                    </p>
                    <div className="mt-5 grid gap-3">
                      <InsightCard
                        detail={
                          ageSummary
                            ? `The analysed population has an average age of ${formatValue(ageSummary.average)} and a median age of ${formatValue(ageSummary.median)}.`
                            : "No age variable was detected, so the population age profile is limited."
                        }
                        label="Population structure"
                      />
                      <InsightCard
                        detail={
                          employmentDistribution[0]
                            ? `${employmentDistribution[0].label} is the largest labour-market category (${employmentDistribution[0].percent.toFixed(0)}% of analysed records).`
                            : "No employment or labour-market status variable was detected."
                        }
                        label="Labour profile"
                      />
                      <InsightCard
                        detail={
                          incomeSummary
                            ? `Observed income has a median of ${formatValue(incomeSummary.median)} and ranges from ${formatValue(incomeSummary.min)} to ${formatValue(incomeSummary.max)}.`
                            : "No income-like variable was detected."
                        }
                        label="Income distribution"
                      />
                      <InsightCard
                        detail={
                          correlations[0]
                            ? `Among numeric variables, the strongest relationship is ${correlations[0].left} x ${correlations[0].right} at ${correlations[0].value.toFixed(3)}.`
                            : "No numeric relationship can be calculated yet. At least two numeric variables with paired values are needed."
                        }
                        label="Strongest numeric relationship"
                      />
                      {notableProfiles.slice(0, 4).map((profile) => (
                        <InsightCard
                          detail={`${profile.column}: average ${formatValue(profile.average)}, low ${formatValue(profile.min)}, high ${formatValue(profile.max)}.`}
                          key={profile.column}
                          label="Numeric summary"
                        />
                      ))}
                      {missingProfiles.length > 0 ? (
                        <InsightCard
                          detail={`${missingProfiles[0].column} has ${missingProfiles[0].missing.toLocaleString()} missing values. Review this before drawing conclusions from that field.`}
                          label="Missing value warning"
                          tone="warning"
                        />
                      ) : (
                        <InsightCard detail="No missing values were detected across the loaded cells." label="Missing value warning" />
                      )}
                      {dataQuality.duplicateRows > 0 ? (
                        <InsightCard
                          detail={`${dataQuality.duplicateRows.toLocaleString()} duplicate row${dataQuality.duplicateRows === 1 ? "" : "s"} detected based on all columns.`}
                          label="Duplicate warning"
                          tone="warning"
                        />
                      ) : (
                        <InsightCard detail="No duplicate rows detected using all columns." label="Duplicate warning" />
                      )}
                    </div>
                  </GlassPanel>

                  <GlassPanel className="p-5">
                    <h2 className="text-xl font-semibold text-white">Data Quality</h2>
                    <p className="mt-1 text-sm text-slate-400">Completeness, duplicate checks, and type detection for the loaded dataset.</p>
                    <div className="mt-5">
                      <div className="flex items-end justify-between gap-4">
                        <div>
                          <p className="text-sm font-semibold text-slate-400">Completeness score</p>
                          <p className="mt-1 text-4xl font-semibold text-white">{dataQuality.completeness.toFixed(1)}%</p>
                        </div>
                        <p className="text-right text-sm text-slate-500">
                          {dataQuality.missingValues.toLocaleString()} missing of {dataQuality.totalCells.toLocaleString()} cells
                        </p>
                      </div>
                      <div className="quality-bar mt-4">
                        <div style={{ width: `${dataQuality.completeness}%` }} />
                      </div>
                    </div>
                    <div className="mt-5 grid gap-3 sm:grid-cols-2">
                      <div className="metric-tile">
                        <span>Missing values</span>
                        <strong>{dataQuality.missingValues.toLocaleString()}</strong>
                      </div>
                      <div className="metric-tile">
                        <span>Duplicate rows</span>
                        <strong>{dataQuality.duplicateRows.toLocaleString()}</strong>
                      </div>
                    </div>
                  </GlassPanel>
                </div>

                <GlassPanel className="p-5">
                  <h2 className="text-xl font-semibold text-white">Column type detection</h2>
                  <div className="mt-5 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                    {columnProfiles.map((profile) => (
                      <div className="column-profile" key={profile.column}>
                        <div className="flex items-center justify-between gap-3">
                          <strong>{profile.column}</strong>
                          <span>{profile.type}</span>
                        </div>
                        <p className="mt-2 text-sm text-slate-500">
                          {profile.unique.toLocaleString()} unique / {profile.missing.toLocaleString()} missing / {profile.completeness.toFixed(0)}% complete
                        </p>
                        {profile.type === "numeric" && (
                          <p className="mt-2 text-sm text-slate-500">
                            Avg {formatValue(profile.average)} / Min {formatValue(profile.min)} / Max {formatValue(profile.max)}
                          </p>
                        )}
                      </div>
                    ))}
                  </div>
                </GlassPanel>
              </div>
            )}

            {activeTab === "quality" && (
              <div className="space-y-5">
                <GlassPanel className="p-5">
                  <h2 className="text-xl font-semibold text-white">Data Quality</h2>
                  <p className="mt-1 text-sm text-slate-400">
                    Review whether the survey file is ready for demographic analysis before interpreting substantive patterns.
                  </p>
                  <div className="mt-4">
                    <InsightCard detail={dataQualityInterpretation} label="Interpretation" tone={dataQuality.completeness < 90 || dataQuality.duplicateRows > 0 ? "warning" : "default"} />
                  </div>
                  <div className="mt-5">
                    <div className="flex items-end justify-between gap-4">
                      <div>
                        <p className="text-sm font-semibold text-slate-400">Completeness score</p>
                        <p className="mt-1 text-4xl font-semibold text-white">{dataQuality.completeness.toFixed(1)}%</p>
                      </div>
                      <p className="text-right text-sm text-slate-500">
                        {dataQuality.missingValues.toLocaleString()} missing of {dataQuality.totalCells.toLocaleString()} cells
                      </p>
                    </div>
                    <div className="quality-bar mt-4">
                      <div style={{ width: `${dataQuality.completeness}%` }} />
                    </div>
                  </div>
                  <div className="mt-5 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                    <div className="metric-tile">
                      <span>Missing values</span>
                      <strong>{dataQuality.missingValues.toLocaleString()}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Duplicate rows</span>
                      <strong>{dataQuality.duplicateRows.toLocaleString()}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Variables</span>
                      <strong>{columns.length}</strong>
                    </div>
                    <div className="metric-tile">
                      <span>Detected survey vars</span>
                      <strong>{surveyDetectionList.filter((detection) => detection.column).length}</strong>
                    </div>
                  </div>
                </GlassPanel>

                <GlassPanel className="p-5">
                  <h2 className="text-xl font-semibold text-white">Variable readiness</h2>
                  <div className="mt-5 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                    {columnProfiles.map((profile) => (
                      <div className="column-profile" key={profile.column}>
                        <div className="flex items-center justify-between gap-3">
                          <strong>{profile.column}</strong>
                          <span>{profile.type}</span>
                        </div>
                        <p className="mt-2 text-sm text-slate-500">
                          {profile.unique.toLocaleString()} unique / {profile.missing.toLocaleString()} missing / {profile.completeness.toFixed(0)}% complete
                        </p>
                        {profile.type === "numeric" && (
                          <p className="mt-2 text-sm text-slate-500">
                            Avg {formatValue(profile.average)} / Min {formatValue(profile.min)} / Max {formatValue(profile.max)}
                          </p>
                        )}
                      </div>
                    ))}
                  </div>
                </GlassPanel>
              </div>
            )}

            {activeTab === "visualization" && (
              <div className="space-y-5">
                <GlassPanel className="p-4 sm:p-5">
                  <div className="flex flex-col gap-4 xl:flex-row xl:items-start xl:justify-between">
                    <div className="max-w-3xl">
                      <h2 className="text-xl font-semibold text-white">Survey visualizations</h2>
                      <p className="mt-1 text-sm leading-6 text-slate-400">
                        Choose an analysis view based on detected social survey variables. The app prioritizes demographic, income, labour, education, and regional questions before generic chart settings.
                      </p>
                    </div>
                    <div className="rounded-md border border-blue-200 bg-blue-50 px-4 py-3 text-sm leading-6 text-blue-900">
                      {activeVisualizationMode === "populationPyramid" ? "Recommended first: age and sex/gender were detected." : `Current view: ${socialPlot.title}`}
                    </div>
                  </div>

                  <h3 className="mt-5 text-sm font-semibold uppercase text-slate-500">Recommended Visualizations</h3>
                  <div className="visualization-recommendations mt-3">
                    <button className={`visualization-card ${visualizationMode === "auto" ? "visualization-card-active" : ""}`} onClick={() => setVisualizationMode("auto")} type="button">
                      <span>Auto Recommended</span>
                      <strong>{recommendedVisualizations.find((item) => item.enabled)?.title ?? "Custom Relationship"}</strong>
                      <em>Let the app choose the strongest survey chart.</em>
                    </button>
                    {recommendedVisualizations.map((item) => (
                      <button
                        className={`visualization-card ${activeVisualizationMode === item.key && visualizationMode !== "auto" ? "visualization-card-active" : ""} ${item.enabled ? "" : "visualization-card-disabled"}`}
                        key={item.key}
                        onClick={() => setVisualizationMode(item.key)}
                        type="button"
                      >
                        <span>{item.enabled ? "Available" : "Not detected"}</span>
                        <strong>{item.title}</strong>
                        <em>{item.detail}</em>
                      </button>
                    ))}
                  </div>

                  {activeVisualizationMode === "customRelationship" && (
                    <div className="chart-controls mt-5">
                      <label>
                        <span>Relationship view</span>
                        <select className="field-select" onChange={(event) => setChartType(event.target.value as ChartType)} value={chartType}>
                          <option value="scatter">Relationship plot</option>
                          <option value="bar">Grouped comparison</option>
                          <option value="line">Ordered trend</option>
                          <option value="histogram">Distribution</option>
                        </select>
                      </label>
                      <label>
                        <span>First measure</span>
                        <select className="field-select" onChange={(event) => setXColumn(event.target.value)} value={selectedX}>
                          {numericColumns.map((column) => (
                            <option key={column}>{column}</option>
                          ))}
                        </select>
                      </label>
                      <label>
                        <span>Second measure</span>
                        <select className="field-select" disabled={chartType === "histogram"} onChange={(event) => setYColumn(event.target.value)} value={selectedY}>
                          {numericColumns.map((column) => (
                            <option key={column}>{column}</option>
                          ))}
                        </select>
                      </label>
                      <label>
                        <span>Compare groups</span>
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
                  )}
                </GlassPanel>

                <GlassPanel className="p-4 sm:p-5">
                  <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
                    <div>
                      <h3 className="text-lg font-semibold text-white">{socialPlot.title}</h3>
                      <p className="mt-1 text-sm text-slate-400">Generated from the current search and filters.</p>
                    </div>
                    <div className="max-w-2xl">
                      <InsightCard detail={socialPlot.interpretation} label="Automatic interpretation" />
                    </div>
                  </div>

                  <div className="mt-6 chart-frame">
                    {socialPlot.ready ? (
                      <Plot
                        config={{ displaylogo: false, responsive: true }}
                        data={socialPlot.data}
                        layout={socialPlot.layout}
                        style={{ height: "100%", width: "100%" }}
                        useResizeHandler
                      />
                    ) : (
                      <EmptyState title="This survey view needs detected variables" detail="Use the Dataset Report to inspect variable detection, or choose another recommended visualization that is available for this dataset." />
                    )}
                  </div>

                  <div className="mt-5">
                    <h4 className="text-sm font-semibold uppercase text-slate-500">Suggested next analyses</h4>
                    <div className="mt-3 grid gap-3 md:grid-cols-3">
                      {socialPlot.nextAnalyses.map((analysis) => (
                        <InsightCard detail={analysis} key={analysis} label="Next step" />
                      ))}
                    </div>
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
                    <div className="mt-4 max-w-2xl">
                      <InsightCard
                        detail={`The preview shows the first ${Math.min(filteredRows.length, 150).toLocaleString()} matching records, helping verify coding, labels, missing values, and unexpected categories before analysis.`}
                        label="Interpretation"
                      />
                    </div>
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
                    <div className="mt-4 max-w-2xl">
                      <InsightCard
                        detail={`Export will include ${filteredRows.length.toLocaleString()} records and ${columns.length.toLocaleString()} variables after the current search, filters, and sort are applied.`}
                        label="Interpretation"
                      />
                    </div>
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
