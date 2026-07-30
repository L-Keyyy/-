import type {
  PerformanceRecord,
  RouteProperty,
  RouteRow,
} from "../types";

export const REGION_OPTIONS = [
  { code: "NE", name: "东北区", source: "东北区" },
  { code: "GL", name: "大湖区", source: "大湖区" },
  { code: "FL", name: "佛州区", source: "佛州区" },
  { code: "WE", name: "美西大区", source: "美西大区" },
  { code: "MS", name: "中南大区", source: "中南大区" },
  { code: "TX", name: "德州大区", source: "德州大区" },
] as const;

export const sum = (values: number[]) =>
  values.reduce((total, value) => total + (Number(value) || 0), 0);

export function percentile(values: number[], p: number) {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (!sorted.length) return 0;
  const index = (sorted.length - 1) * p;
  const lower = Math.floor(index);
  const upper = Math.ceil(index);
  if (lower === upper) return sorted[lower];
  return sorted[lower] + (sorted[upper] - sorted[lower]) * (index - lower);
}

export function median(values: number[]) {
  return percentile(values, 0.5);
}

export function pph(attempted: number, hours: number) {
  return hours > 0 ? attempted / hours : 0;
}

export function failRate(delivered: number, attempted: number) {
  return attempted > 0 ? Math.max(0, attempted - delivered) / attempted : 0;
}

export function sortWeeks(weeks: string[]) {
  return [...new Set(weeks)].sort((a, b) => {
    const aNumber = Number(a.replace(/\D/g, ""));
    const bNumber = Number(b.replace(/\D/g, ""));
    return aNumber - bNumber || a.localeCompare(b);
  });
}

export function buildPropertyMap(properties: RouteProperty[]) {
  return new Map(properties.map((item) => [item.route, item]));
}

export function cleanPerformanceRecords(records: PerformanceRecord[]) {
  const reasonCounts = {
    invalidNumber: 0,
    negativeValue: 0,
    invalidVolume: 0,
    emptyWork: 0,
    lowVolume: 0,
    timeMismatch: 0,
    extremeEfficiency: 0,
  };

  const structurallyValid = records.filter((row) => {
    const numericValues = [
      row.delivered,
      row.attempted,
      row.sortHours,
      row.transitHours,
      row.deliveryHours,
      row.totalHours,
    ];
    if (numericValues.some((value) => !Number.isFinite(value))) {
      reasonCounts.invalidNumber += 1;
      return false;
    }
    if (numericValues.some((value) => value < 0)) {
      reasonCounts.negativeValue += 1;
      return false;
    }
    if (row.delivered > row.attempted) {
      reasonCounts.invalidVolume += 1;
      return false;
    }
    if (row.attempted <= 0 || row.totalHours <= 0) {
      reasonCounts.emptyWork += 1;
      return false;
    }
    if (row.delivered < 100) {
      reasonCounts.lowVolume += 1;
      return false;
    }

    const componentHours =
      row.sortHours + row.transitHours + row.deliveryHours;
    if (
      componentHours > 0 &&
      Math.abs(componentHours - row.totalHours) >
        Math.max(2, row.totalHours * 0.15)
    ) {
      reasonCounts.timeMismatch += 1;
      return false;
    }
    return true;
  });

  // Detect only implausibly high efficiency points. Low PPH remains visible
  // because it is an operational signal rather than a data-quality error.
  const efficiencyGroups = new Map<string, number[]>();
  structurallyValid.forEach((row) => {
    const key = `${row.region}¦${row.week}`;
    const values = efficiencyGroups.get(key) ?? [];
    values.push(Math.log(pph(row.attempted, row.totalHours)));
    efficiencyGroups.set(key, values);
  });

  const upperFences = new Map<string, number>();
  efficiencyGroups.forEach((values, key) => {
    if (values.length < 8) return;
    const q1 = percentile(values, 0.25);
    const q3 = percentile(values, 0.75);
    const iqr = q3 - q1;
    if (iqr > 0) upperFences.set(key, q3 + iqr * 4);
  });

  const cleaned = structurallyValid.filter((row) => {
    const upperFence = upperFences.get(`${row.region}¦${row.week}`);
    if (
      upperFence !== undefined &&
      Math.log(pph(row.attempted, row.totalHours)) > upperFence
    ) {
      reasonCounts.extremeEfficiency += 1;
      return false;
    }
    return true;
  });

  return {
    records: cleaned,
    excluded: records.length - cleaned.length,
    reasonCounts,
  };
}

export function aggregatePerformance(records: PerformanceRecord[]) {
  const aggregate = records.reduce(
    (current, row) => {
      current.delivered += row.delivered || 0;
      current.attempted += row.attempted || 0;
      current.sortHours += row.sortHours || 0;
      current.transitHours += row.transitHours || 0;
      current.deliveryHours += row.deliveryHours || 0;
      current.totalHours += row.totalHours || 0;
      return current;
    },
    {
      delivered: 0,
      attempted: 0,
      sortHours: 0,
      transitHours: 0,
      deliveryHours: 0,
      totalHours: 0,
    },
  );

  return {
    ...aggregate,
    operationPph: pph(aggregate.attempted, aggregate.totalHours),
    successPph: pph(aggregate.delivered, aggregate.totalHours),
    failRate: failRate(aggregate.delivered, aggregate.attempted),
  };
}

export function aggregateRouteRows(
  records: PerformanceRecord[],
  previousRecords: PerformanceRecord[],
  properties: RouteProperty[],
) {
  const propertyMap = buildPropertyMap(properties);
  const previousMap = groupRows(previousRecords);
  const currentMap = groupRows(records);
  const rawRows = [...currentMap.values()].map((row) => {
    const metrics = aggregatePerformance([row]);
    const previous = previousMap.get(rowKey(row));
    const previousPph = previous
      ? pph(previous.attempted, previous.totalHours)
      : 0;
    const wow =
      previousPph > 0
        ? (metrics.operationPph - previousPph) / previousPph
        : null;
    const property = propertyMap.get(row.route);
    return {
      ...row,
      ...metrics,
      wow,
      percentile: "",
      businessMode: property?.businessMode || "未标注",
      isNew: property?.isNew || "未标注",
    } satisfies RouteRow;
  });

  const p25 = percentile(
    rawRows.map((row) => row.operationPph).filter((value) => value > 0),
    0.25,
  );
  const p50 = percentile(
    rawRows.map((row) => row.operationPph).filter((value) => value > 0),
    0.5,
  );
  const p75 = percentile(
    rawRows.map((row) => row.operationPph).filter((value) => value > 0),
    0.75,
  );

  return rawRows.map((row) => ({
    ...row,
    percentile:
      row.operationPph < p25
        ? "< P25"
        : row.operationPph < p50
          ? "P25–P50"
          : row.operationPph < p75
            ? "P50–P75"
            : "≥ P75",
  }));
}

export function rowKey(row: Pick<RouteRow, "route" | "site" | "dsp">) {
  return `${row.route}¦${row.site}¦${row.dsp}`;
}

function groupRows(records: PerformanceRecord[]) {
  const grouped = new Map<string, PerformanceRecord>();
  for (const row of records) {
    const key = rowKey(row);
    const current = grouped.get(key);
    if (!current) {
      grouped.set(key, { ...row });
      continue;
    }
    current.delivered += row.delivered || 0;
    current.attempted += row.attempted || 0;
    current.sortHours += row.sortHours || 0;
    current.transitHours += row.transitHours || 0;
    current.deliveryHours += row.deliveryHours || 0;
    current.totalHours += row.totalHours || 0;
  }
  return grouped;
}

export function formatNumber(value: number, digits = 0) {
  const normalized = Number.isFinite(value) ? value : 0;
  const absolute = Math.abs(normalized);

  // Counts and hours use compact Chinese units to avoid clipped dashboard values.
  // Rate-like metrics explicitly pass a precision and keep their useful decimals.
  if (digits === 0 && absolute >= 100_000_000) {
    return `${new Intl.NumberFormat("zh-CN", {
      maximumFractionDigits: 1,
    }).format(normalized / 100_000_000)}亿`;
  }
  if (digits === 0 && absolute >= 10_000) {
    return `${new Intl.NumberFormat("zh-CN", {
      maximumFractionDigits: 1,
    }).format(normalized / 10_000)}万`;
  }

  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits,
  }).format(normalized);
}

export function formatPercent(value: number, digits = 1) {
  return `${formatNumber(value * 100, digits)}%`;
}

export function addressMixItems(value: string) {
  if (!value) return [];
  return value
    .split(",")
    .map((item) => {
      const match = item.trim().match(/^(.*?)(\d+(?:\.\d+)?)%$/);
      return match
        ? { name: match[1].trim(), value: Number(match[2]) }
        : { name: item.trim(), value: 0 };
    })
    .filter((item) => item.name);
}

export function csvEscape(value: string | number) {
  const text = String(value ?? "");
  return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}
