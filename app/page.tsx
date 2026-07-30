"use client";

import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type ComponentType,
} from "react";
import type { EChartsOption } from "echarts";
import * as XLSX from "xlsx";
import {
  Activity,
  AlertTriangle,
  ArrowDownRight,
  ArrowLeft,
  ArrowRight,
  ArrowUpDown,
  ArrowUpRight,
  BarChart3,
  BellRing,
  Building2,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  CircleDot,
  Clock3,
  Database,
  Download,
  FileSpreadsheet,
  Filter,
  Gauge,
  Layers3,
  LayoutDashboard,
  MapPinned,
  Menu,
  PackageCheck,
  PanelLeftClose,
  RefreshCw,
  Route,
  Search,
  ShieldCheck,
  Sparkles,
  TrendingDown,
  Truck,
  Upload,
  UsersRound,
  X,
  Zap,
} from "lucide-react";
import Chart from "./components/Chart";
import {
  REGION_OPTIONS,
  addressMixItems,
  aggregatePerformance,
  aggregateRouteRows,
  buildPropertyMap,
  cleanPerformanceRecords,
  csvEscape,
  failRate,
  formatNumber,
  formatPercent,
  median,
  percentile,
  pph,
  rowKey,
  sortWeeks,
  sum,
} from "./lib/analytics";
import type {
  InitialData,
  PerformanceRecord,
  RouteProperty,
  RouteRow,
} from "./types";

type RegionCode = "ALL" | (typeof REGION_OPTIONS)[number]["code"];
type SortKey =
  | "attempted"
  | "operationPph"
  | "successPph"
  | "failRate"
  | "wow"
  | "route";

const NAV_ITEMS = [
  { id: "overview", label: "全国总览", icon: LayoutDashboard },
  ...REGION_OPTIONS.map((region) => ({
    id: region.code,
    label: `${region.code} · ${region.name}周报`,
    icon: MapPinned,
  })),
  { id: "exceptions", label: "重点异常", icon: AlertTriangle },
  { id: "data", label: "全量数据", icon: Database },
];

const PALETTE = {
  navy: "#0B2451",
  blue: "#2563EB",
  cyan: "#0891B2",
  green: "#16A36A",
  red: "#DC4C56",
  orange: "#E9852D",
  slate: "#64748B",
  grid: "#E9EEF5",
};

function normalizeText(value: unknown) {
  return String(value ?? "").trim();
}

function normalizeNumber(value: unknown) {
  const result = Number(value ?? 0);
  return Number.isFinite(result) ? result : 0;
}

function normalizeDate(value: unknown) {
  if (value instanceof Date) return value.toISOString();
  return normalizeText(value);
}

function normalizeRate(value: unknown) {
  if (typeof value === "string" && value.trim().endsWith("%")) {
    return (Number.parseFloat(value) || 0) / 100;
  }
  return normalizeNumber(value);
}

async function readWorkbookRows(file: File) {
  const isCsv = file.name.toLowerCase().endsWith(".csv");
  const workbook = isCsv
    ? XLSX.read(await file.text(), { type: "string", cellDates: true })
    : XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: null,
  });
}

function normalizePerformanceRows(rows: Record<string, unknown>[]) {
  return rows
    .map(
      (row): PerformanceRecord => ({
        week: normalizeText(row["周数"] ?? row["周次"]),
        weekStart: normalizeDate(row["本周开始日期"] ?? row["日期"]),
        route: normalizeText(row["路区"] ?? row["路区名称"]),
        dsp: normalizeText(row["DSP"] ?? row["车队名称"]),
        site: normalizeText(row["站点"] ?? row["转运站点"]),
        region: normalizeText(row["大区"]),
        delivered: normalizeNumber(row["配送量"] ?? row["成功配送量"]),
        attempted: normalizeNumber(
          row["配送量（加派送失败）"] ??
            row["配送量(加派送失败)"] ??
            row["配送量（含派送失败）"],
        ),
        sortHours: normalizeNumber(row["分拣耗时"]),
        transitHours: normalizeNumber(row["在途耗时"]),
        deliveryHours: normalizeNumber(row["配送耗时"]),
        totalHours: normalizeNumber(row["总时长"]),
      }),
    )
    .filter((row) => row.week && row.route && row.region);
}

function normalizePropertyRows(rows: Record<string, unknown>[]) {
  return rows
    .map(
      (row): RouteProperty => ({
        route: normalizeText(row["路区名称"] ?? row["路区"]),
        businessMode: normalizeText(row["业务模式"]),
        sortCode: normalizeText(row["分拣码"]),
        transferSite: normalizeText(row["转运站点"]),
        fleet: normalizeText(row["车队名称"]),
        status: normalizeText(row["状态"]),
        postalCodes: normalizeText(row["邮编"]),
        addressMix: normalizeText(row["收件地址类型占比"]),
        safety: normalizeText(row["安全度"]),
        landArea: normalizeNumber(row["陆地面积（mi²）"]),
        populationDensity: normalizeNumber(row["人口密度（人/mi²）"]),
        isNew: normalizeText(row["是否新开"]),
        difficulty: normalizeText(row["路区难易度"]),
        firstMile: normalizeNumber(row["首单里程（mi）"]),
        expertPph: normalizeNumber(row["熟手PPH（件）"]),
        deliveryExceptionRate: normalizeRate(row["派送异常率"]),
        dnrRate: normalizeRate(row["DNR率"]),
      }),
    )
    .filter((row) => row.route);
}

function mergeRouteProperties(
  current: RouteProperty[],
  incoming: RouteProperty[],
) {
  const merged = new Map(current.map((item) => [item.route, item]));
  incoming.forEach((item) => {
    const previous = merged.get(item.route);
    if (!previous) {
      merged.set(item.route, item);
      return;
    }
    merged.set(
      item.route,
      Object.fromEntries(
        Object.entries({ ...previous, ...item }).map(([key, value]) => [
          key,
          value === "" || value === 0
            ? previous[key as keyof RouteProperty] ?? value
            : value,
        ]),
      ) as RouteProperty,
    );
  });
  return [...merged.values()];
}

function groupComparison(rows: RouteRow[], field: "site" | "dsp") {
  const groups = new Map<string, PerformanceRecord[]>();
  rows.forEach((row) => {
    const label = row[field] || "未标注";
    groups.set(label, [...(groups.get(label) ?? []), row]);
  });
  return [...groups.entries()]
    .map(([name, records]) => ({
      name,
      ...aggregatePerformance(records),
    }))
    .filter((row) => row.totalHours > 0)
    .sort((a, b) => b.operationPph - a.operationPph)
    .slice(0, 12);
}

function addressMixSummary(value?: string) {
  if (!value) return "未标注";
  const items = addressMixItems(value)
    .sort((a, b) => b.value - a.value)
    .slice(0, 2);
  return items.length
    ? items
        .map((item) => `${item.name} ${formatNumber(item.value, 1)}%`)
        .join(" · ")
    : "未标注";
}

function MetricCard({
  label,
  value,
  detail,
  icon: Icon,
  tone = "blue",
  change,
}: {
  label: string;
  value: string;
  detail: string;
  icon: ComponentType<{ size?: number; strokeWidth?: number }>;
  tone?: "blue" | "green" | "red" | "orange" | "slate";
  change?: number | null;
}) {
  const toneClasses = {
    blue: "metric-icon-blue",
    green: "metric-icon-green",
    red: "metric-icon-red",
    orange: "metric-icon-orange",
    slate: "metric-icon-slate",
  };
  return (
    <article className="metric-card">
      <div className={`metric-icon ${toneClasses[tone]}`}>
        <Icon size={18} strokeWidth={2} />
      </div>
      <div className="metric-copy">
        <div className="metric-label">{label}</div>
        <div className="metric-value-row">
          <strong>{value}</strong>
          {change !== undefined && change !== null ? (
            <span
              className={`change-pill ${
                change >= 0 ? "change-positive" : "change-negative"
              }`}
            >
              {change >= 0 ? (
                <ArrowUpRight size={13} />
              ) : (
                <ArrowDownRight size={13} />
              )}
              {formatPercent(Math.abs(change))}
            </span>
          ) : null}
        </div>
        <p>{detail}</p>
      </div>
    </article>
  );
}

function SectionHeader({
  eyebrow,
  title,
  description,
  right,
}: {
  eyebrow?: string;
  title: string;
  description?: string;
  right?: React.ReactNode;
}) {
  return (
    <div className="section-header">
      <div>
        {eyebrow ? <div className="section-eyebrow">{eyebrow}</div> : null}
        <h2>{title}</h2>
        {description ? <p>{description}</p> : null}
      </div>
      {right ? <div className="section-header-right">{right}</div> : null}
    </div>
  );
}

function EmptyState({ text }: { text: string }) {
  return (
    <div className="empty-state">
      <CircleDot size={22} />
      <span>{text}</span>
    </div>
  );
}

export default function Home() {
  const [records, setRecords] = useState<PerformanceRecord[]>([]);
  const [properties, setProperties] = useState<RouteProperty[]>([]);
  const [sourceMeta, setSourceMeta] = useState<InitialData["meta"] | null>(null);
  const [excludedCount, setExcludedCount] = useState(0);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState<"performance" | "property" | null>(
    null,
  );
  const [notice, setNotice] = useState("");
  const [activeRegion, setActiveRegion] = useState<RegionCode>("WE");
  const [activeNav, setActiveNav] = useState("WE");
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [selectedWeek, setSelectedWeek] = useState("");
  const [siteFilter, setSiteFilter] = useState("全部站点");
  const [dspFilter, setDspFilter] = useState("全部DSP");
  const [newFilter, setNewFilter] = useState("全部");
  const [businessFilter, setBusinessFilter] = useState("全部模式");
  const [comparisonMode, setComparisonMode] = useState<"site" | "dsp">("site");
  const [selectedRoute, setSelectedRoute] = useState<RouteRow | null>(null);
  const [expandedWatchlists, setExpandedWatchlists] = useState<
    Record<string, boolean>
  >({});
  const [tableSearch, setTableSearch] = useState("");
  const [tableSort, setTableSort] = useState<SortKey>("attempted");
  const [sortDirection, setSortDirection] = useState<"asc" | "desc">("desc");
  const [page, setPage] = useState(1);
  const performanceInput = useRef<HTMLInputElement>(null);
  const propertyInput = useRef<HTMLInputElement>(null);

  useEffect(() => {
    fetch("/data/initial.json")
      .then((response) => {
        if (!response.ok) throw new Error("Initial data is not available.");
        return response.json() as Promise<InitialData>;
      })
      .then((data) => {
        const cleaned = cleanPerformanceRecords(data.records);
        setRecords(cleaned.records);
        setProperties(data.properties);
        setExcludedCount(cleaned.excluded);
        setSourceMeta({
          ...data.meta,
          aggregatedRows: cleaned.records.length,
        });
        const weeks = sortWeeks(cleaned.records.map((row) => row.week));
        setSelectedWeek(weeks.at(-1) ?? "");
      })
      .catch(() => {
        setNotice("请上传运营明细与路区属性文件开始分析。");
      })
      .finally(() => setLoading(false));

  }, []);

  const propertyMap = useMemo(() => buildPropertyMap(properties), [properties]);
  const weeks = useMemo(
    () => sortWeeks(records.map((row) => row.week)),
    [records],
  );
  const currentRegionName = useMemo(() => {
    if (activeRegion === "ALL") return "全国";
    return (
      REGION_OPTIONS.find((region) => region.code === activeRegion)?.name ??
      "全国"
    );
  }, [activeRegion]);
  const currentRegionSource = useMemo(() => {
    if (activeRegion === "ALL") return "";
    return (
      REGION_OPTIONS.find((region) => region.code === activeRegion)?.source ?? ""
    );
  }, [activeRegion]);

  const recordMatches = useCallback(
    (record: PerformanceRecord, ignoreWeek = false) => {
      if (currentRegionSource && record.region !== currentRegionSource)
        return false;
      if (!ignoreWeek && selectedWeek && record.week !== selectedWeek)
        return false;
      if (siteFilter !== "全部站点" && record.site !== siteFilter) return false;
      if (dspFilter !== "全部DSP" && record.dsp !== dspFilter) return false;
      const property = propertyMap.get(record.route);
      const businessMode = property?.businessMode || "未标注";
      const isNew = property?.isNew || "未标注";
      if (
        businessFilter !== "全部模式" &&
        businessMode !== businessFilter
      )
        return false;
      if (newFilter !== "全部" && isNew !== newFilter) return false;
      return true;
    },
    [
      businessFilter,
      currentRegionSource,
      dspFilter,
      newFilter,
      propertyMap,
      selectedWeek,
      siteFilter,
    ],
  );

  const regionRecords = useMemo(
    () =>
      records.filter(
        (row) => !currentRegionSource || row.region === currentRegionSource,
      ),
    [currentRegionSource, records],
  );
  const siteOptions = useMemo(
    () =>
      [...new Set(regionRecords.map((row) => row.site).filter(Boolean))].sort(),
    [regionRecords],
  );
  const dspOptions = useMemo(
    () =>
      [
        ...new Set(
          regionRecords
            .filter(
              (row) => siteFilter === "全部站点" || row.site === siteFilter,
            )
            .map((row) => row.dsp)
            .filter(Boolean),
        ),
      ].sort(),
    [regionRecords, siteFilter],
  );
  const businessOptions = useMemo(
    () =>
      [
        ...new Set(
          properties.map((item) => item.businessMode || "未标注"),
        ),
      ].sort(),
    [properties],
  );
  const newOptions = useMemo(
    () =>
      [...new Set(properties.map((item) => item.isNew || "未标注"))].sort(),
    [properties],
  );

  const currentRecords = useMemo(
    () => records.filter((row) => recordMatches(row)),
    [recordMatches, records],
  );
  const currentWeekIndex = weeks.indexOf(selectedWeek);
  const previousWeek =
    currentWeekIndex > 0 ? weeks[currentWeekIndex - 1] : undefined;
  const previousRecords = useMemo(
    () =>
      previousWeek
        ? records.filter(
            (row) => row.week === previousWeek && recordMatches(row, true),
          )
        : [],
    [previousWeek, recordMatches, records],
  );
  const routeRows = useMemo(
    () => aggregateRouteRows(currentRecords, previousRecords, properties),
    [currentRecords, previousRecords, properties],
  );
  const currentMetrics = useMemo(
    () => aggregatePerformance(currentRecords),
    [currentRecords],
  );
  const previousMetrics = useMemo(
    () => aggregatePerformance(previousRecords),
    [previousRecords],
  );
  const weeklyTrend = useMemo(() => {
    const eligibleWeeks = weeks
      .filter((week) => weeks.indexOf(week) <= currentWeekIndex)
      .slice(-8);
    return eligibleWeeks.map((week) => ({
      week,
      ...aggregatePerformance(
        records.filter(
          (row) => row.week === week && recordMatches(row, true),
        ),
      ),
    }));
  }, [currentWeekIndex, recordMatches, records, weeks]);
  const recentMedian = useMemo(
    () =>
      median(
        weeklyTrend
          .slice(-4)
          .map((item) => item.operationPph)
          .filter((value) => value > 0),
      ),
    [weeklyTrend],
  );
  const wow =
    previousMetrics.operationPph > 0
      ? (currentMetrics.operationPph - previousMetrics.operationPph) /
        previousMetrics.operationPph
      : null;
  const routePphValues = routeRows
    .map((row) => row.operationPph)
    .filter((value) => value > 0);
  const quantiles = {
    p25: percentile(routePphValues, 0.25),
    p50: percentile(routePphValues, 0.5),
    p75: percentile(routePphValues, 0.75),
  };
  const volumeP75 = percentile(
    routeRows.map((row) => row.attempted),
    0.75,
  );
  const failP75 = percentile(
    routeRows.map((row) => row.failRate),
    0.75,
  );

  const comparisonRows = useMemo(
    () => groupComparison(routeRows, comparisonMode),
    [comparisonMode, routeRows],
  );
  const bubbleRows = useMemo(
    () =>
      routeRows
        .map((row) => ({
          row,
          property: propertyMap.get(row.route),
        }))
        .filter(
          ({ property }) =>
            property && Number(property.populationDensity) > 0,
        )
        .slice(0, 200),
    [propertyMap, routeRows],
  );

  const weeklyRouteMap = useMemo(() => {
    const map = new Map<string, RouteRow[]>();
    weeks.forEach((week, index) => {
      if (index > currentWeekIndex) return;
      const weekRecords = records.filter(
        (row) => row.week === week && recordMatches(row, true),
      );
      const previous =
        index > 0
          ? records.filter(
              (row) =>
                row.week === weeks[index - 1] && recordMatches(row, true),
            )
          : [];
      map.set(
        week,
        aggregateRouteRows(weekRecords, previous, properties),
      );
    });
    return map;
  }, [
    currentWeekIndex,
    properties,
    recordMatches,
    records,
    weeks,
  ]);

  const priorityLists = useMemo(() => {
    const drops = routeRows
      .filter((row) => row.wow !== null && row.wow < 0)
      .sort((a, b) => (a.wow ?? 0) - (b.wow ?? 0));
    const highVolumeLow = routeRows
      .filter(
        (row) =>
          row.attempted >= volumeP75 && row.operationPph < quantiles.p25,
      )
      .sort((a, b) => b.attempted - a.attempted);
    const highFailure = routeRows
      .filter((row) => row.failRate >= failP75 && row.failRate > 0)
      .sort((a, b) => b.failRate - a.failRate);
    const newRoutes = routeRows
      .filter(
        (row) =>
          row.isNew.includes("新开") && !row.isNew.includes("非新开"),
      )
      .sort((a, b) => b.attempted - a.attempted);
    const lastThree = weeks
      .filter((week) => weeks.indexOf(week) <= currentWeekIndex)
      .slice(-3);
    const continuous =
      lastThree.length < 3
        ? []
        : routeRows
            .filter((row) =>
              lastThree.every((week) => {
                const rows = weeklyRouteMap.get(week) ?? [];
                const match = rows.find(
                  (candidate) => rowKey(candidate) === rowKey(row),
                );
                const weekP25 = percentile(
                  rows
                    .map((candidate) => candidate.operationPph)
                    .filter((value) => value > 0),
                  0.25,
                );
                return Boolean(match && match.operationPph < weekP25);
              }),
            )
            .sort((a, b) => a.operationPph - b.operationPph);
    return { drops, highVolumeLow, highFailure, newRoutes, continuous };
  }, [
    currentWeekIndex,
    failP75,
    quantiles.p25,
    routeRows,
    volumeP75,
    weeklyRouteMap,
    weeks,
  ]);

  const trendOption = useMemo(
    () =>
      ({
        color: [PALETTE.blue, PALETTE.green],
        tooltip: {
          trigger: "axis",
          backgroundColor: "#0B1E3A",
          borderWidth: 0,
          textStyle: { color: "#fff" },
        },
        legend: {
          data: ["作业PPH", "成功PPH"],
          right: 8,
          top: 0,
          icon: "circle",
          itemWidth: 8,
          textStyle: { color: "#526176" },
        },
        grid: { left: 42, right: 18, top: 46, bottom: 30 },
        xAxis: {
          type: "category",
          data: weeklyTrend.map((item) => item.week),
          boundaryGap: false,
          axisLine: { lineStyle: { color: "#CBD5E1" } },
          axisTick: { show: false },
          axisLabel: { color: "#718096" },
        },
        yAxis: {
          type: "value",
          axisLabel: { color: "#718096" },
          splitLine: { lineStyle: { color: PALETTE.grid } },
        },
        series: [
          {
            name: "作业PPH",
            type: "line",
            smooth: 0.35,
            symbolSize: 7,
            lineStyle: { width: 3 },
            areaStyle: {
              color: {
                type: "linear",
                x: 0,
                y: 0,
                x2: 0,
                y2: 1,
                colorStops: [
                  { offset: 0, color: "rgba(37,99,235,.18)" },
                  { offset: 1, color: "rgba(37,99,235,0)" },
                ],
              },
            },
            data: weeklyTrend.map((item) =>
              Number(item.operationPph.toFixed(2)),
            ),
          },
          {
            name: "成功PPH",
            type: "line",
            smooth: 0.35,
            symbolSize: 7,
            lineStyle: { width: 3 },
            data: weeklyTrend.map((item) =>
              Number(item.successPph.toFixed(2)),
            ),
          },
        ],
      }) as EChartsOption,
    [weeklyTrend],
  );

  const comparisonOption = useMemo(
    () =>
      ({
        color: [PALETTE.blue, "#90B5FF"],
        tooltip: {
          trigger: "axis",
          axisPointer: { type: "shadow" },
          backgroundColor: "#0B1E3A",
          borderWidth: 0,
          textStyle: { color: "#fff" },
        },
        legend: {
          data: ["作业PPH", "成功PPH"],
          right: 0,
          top: 0,
          icon: "circle",
          itemWidth: 8,
          textStyle: { color: "#526176" },
        },
        grid: { left: 42, right: 12, top: 44, bottom: 64 },
        xAxis: {
          type: "category",
          data: comparisonRows.map((row) => row.name),
          axisLabel: {
            color: "#718096",
            rotate: comparisonRows.length > 7 ? 32 : 0,
            overflow: "truncate",
            width: 80,
          },
          axisTick: { show: false },
          axisLine: { lineStyle: { color: "#CBD5E1" } },
        },
        yAxis: {
          type: "value",
          axisLabel: { color: "#718096" },
          splitLine: { lineStyle: { color: PALETTE.grid } },
        },
        series: [
          {
            name: "作业PPH",
            type: "bar",
            barMaxWidth: 18,
            itemStyle: { borderRadius: [5, 5, 0, 0] },
            data: comparisonRows.map((row) =>
              Number(row.operationPph.toFixed(2)),
            ),
          },
          {
            name: "成功PPH",
            type: "bar",
            barMaxWidth: 18,
            itemStyle: { borderRadius: [5, 5, 0, 0] },
            data: comparisonRows.map((row) =>
              Number(row.successPph.toFixed(2)),
            ),
          },
        ],
      }) as EChartsOption,
    [comparisonRows],
  );

  const bubbleOption = useMemo(
    () =>
      ({
        color: [PALETTE.cyan],
        tooltip: {
          trigger: "item",
          backgroundColor: "#0B1E3A",
          borderWidth: 0,
          textStyle: { color: "#fff" },
          formatter: (params: unknown) => {
            const item = params as { data?: (string | number)[] };
            const data = item.data ?? [];
            return `${data[3] ?? ""}<br/>人口密度：${formatNumber(Number(data[0]))}<br/>作业PPH：${formatNumber(Number(data[1]), 2)}<br/>配送量：${formatNumber(Number(data[2]))}`;
          },
        },
        grid: { left: 56, right: 22, top: 24, bottom: 46 },
        xAxis: {
          type: "value",
          name: "人口密度（人/mi²）",
          nameLocation: "middle",
          nameGap: 30,
          axisLabel: { color: "#718096" },
          splitLine: { lineStyle: { color: PALETTE.grid } },
        },
        yAxis: {
          type: "value",
          name: "作业PPH",
          nameLocation: "middle",
          nameGap: 38,
          axisLabel: { color: "#718096" },
          splitLine: { lineStyle: { color: PALETTE.grid } },
        },
        series: [
          {
            type: "scatter",
            data: bubbleRows.map(({ row, property }) => [
              property?.populationDensity ?? 0,
              Number(row.operationPph.toFixed(2)),
              row.attempted,
              row.route,
            ]),
            symbolSize: (value: unknown) => {
              const data = value as number[];
              return Math.max(10, Math.min(44, Math.sqrt(data[2] || 0) / 2.5));
            },
            itemStyle: {
              opacity: 0.78,
              borderColor: "#fff",
              borderWidth: 2,
            },
          },
        ],
      }) as EChartsOption,
    [bubbleRows],
  );

  const timeOption = useMemo(() => {
    const timeTotal =
      currentMetrics.sortHours +
      currentMetrics.transitHours +
      currentMetrics.deliveryHours;
    return {
      color: [PALETTE.blue, PALETTE.orange, PALETTE.cyan],
      tooltip: {
        trigger: "item",
        formatter: "{b}<br/>{c} h · {d}%",
        backgroundColor: "#0B1E3A",
        borderWidth: 0,
        textStyle: { color: "#fff" },
      },
      legend: {
        orient: "vertical",
        right: 8,
        top: "center",
        icon: "circle",
        itemWidth: 9,
        textStyle: { color: "#526176" },
      },
      series: [
        {
          type: "pie",
          radius: ["48%", "72%"],
          center: ["38%", "52%"],
          avoidLabelOverlap: true,
          itemStyle: { borderColor: "#fff", borderWidth: 3 },
          label: { show: false },
          data: [
            {
              name: "分拣",
              value: Number(currentMetrics.sortHours.toFixed(2)),
            },
            {
              name: "在途",
              value: Number(currentMetrics.transitHours.toFixed(2)),
            },
            {
              name: "配送",
              value: Number(currentMetrics.deliveryHours.toFixed(2)),
            },
          ],
        },
      ],
      graphic:
        timeTotal > 0
          ? [
              {
                type: "text",
                left: "29%",
                top: "44%",
                style: {
                  text: `${formatNumber(currentMetrics.totalHours)}h`,
                  fill: PALETTE.navy,
                  fontSize: 18,
                  fontWeight: 700,
                },
              },
              {
                type: "text",
                left: "30%",
                top: "54%",
                style: {
                  text: "总工时",
                  fill: PALETTE.slate,
                  fontSize: 11,
                },
              },
            ]
          : [],
    } as EChartsOption;
  }, [currentMetrics]);

  const handleUpload = async (
    event: ChangeEvent<HTMLInputElement>,
    type: "performance" | "property",
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setUploading(type);
    setNotice("");
    try {
      const rows = await readWorkbookRows(file);
      if (type === "performance") {
        const nextRecords = normalizePerformanceRows(rows);
        if (!nextRecords.length)
          throw new Error("未识别到周数、路区、大区与配送字段。");
        const cleaned = cleanPerformanceRecords(nextRecords);
        setRecords(cleaned.records);
        setExcludedCount(cleaned.excluded);
        const nextWeeks = sortWeeks(cleaned.records.map((row) => row.week));
        setSelectedWeek(nextWeeks.at(-1) ?? "");
        setSourceMeta({
          sourceRows: rows.length,
          aggregatedRows: cleaned.records.length,
          propertyRows: properties.length,
          generatedAt: new Date().toISOString(),
        });
        setNotice(
          `运营明细已更新：保留 ${formatNumber(cleaned.records.length)} 条，自动剔除 ${formatNumber(cleaned.excluded)} 条未达门槛、异常或空载记录。`,
        );
      } else {
        const nextProperties = normalizePropertyRows(rows);
        if (!nextProperties.length)
          throw new Error("未识别到“路区名称”字段。");
        const mergedProperties = mergeRouteProperties(
          properties,
          nextProperties,
        );
        setProperties(mergedProperties);
        setSourceMeta((current) =>
          current
            ? { ...current, propertyRows: mergedProperties.length }
            : {
                sourceRows: 0,
                aggregatedRows: records.length,
                propertyRows: mergedProperties.length,
                generatedAt: new Date().toISOString(),
              },
        );
        setNotice(
          `地址与难易度数据已合并：本次 ${formatNumber(nextProperties.length)} 条，累计 ${formatNumber(mergedProperties.length)} 条。`,
        );
      }
    } catch (error) {
      setNotice(
        error instanceof Error ? error.message : "文件解析异常，请检查字段。",
      );
    } finally {
      setUploading(null);
      event.target.value = "";
    }
  };

  const navigate = (id: string) => {
    setActiveNav(id);
    setSidebarOpen(false);
    if (id === "overview") {
      setActiveRegion("ALL");
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }
    const region = REGION_OPTIONS.find((item) => item.code === id);
    if (region) {
      setActiveRegion(region.code);
      setSiteFilter("全部站点");
      setDspFilter("全部DSP");
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }
    document.getElementById(id)?.scrollIntoView({ behavior: "smooth" });
  };

  const sortedTableRows = useMemo(() => {
    const search = tableSearch.trim().toLowerCase();
    return routeRows
      .filter((row) =>
        search
          ? [row.route, row.site, row.dsp, row.businessMode]
              .join(" ")
              .toLowerCase()
              .includes(search)
          : true,
      )
      .sort((a, b) => {
        const aValue = a[tableSort] ?? -Infinity;
        const bValue = b[tableSort] ?? -Infinity;
        const result =
          typeof aValue === "string"
            ? aValue.localeCompare(String(bValue))
            : Number(aValue) - Number(bValue);
        return sortDirection === "asc" ? result : -result;
      });
  }, [routeRows, sortDirection, tableSearch, tableSort]);
  const pageSize = 15;
  const pageCount = Math.max(1, Math.ceil(sortedTableRows.length / pageSize));
  const visibleRows = sortedTableRows.slice(
    (page - 1) * pageSize,
    page * pageSize,
  );

  useEffect(() => {
    setPage(1);
  }, [
    activeRegion,
    businessFilter,
    dspFilter,
    newFilter,
    selectedWeek,
    siteFilter,
    tableSearch,
  ]);

  const toggleSort = (key: SortKey) => {
    if (tableSort === key) {
      setSortDirection((current) => (current === "asc" ? "desc" : "asc"));
    } else {
      setTableSort(key);
      setSortDirection(key === "route" ? "asc" : "desc");
    }
  };

  const exportRows = () => {
    const headers = [
      "周次",
      "大区",
      "路区",
      "站点",
      "DSP",
      "成功配送量",
      "配送量（含派送失败）",
      "总工时",
      "作业PPH",
      "成功PPH",
      "派送失败率",
      "周环比",
      "本区位置",
      "业务模式",
      "是否新开",
      "路区难易度",
      "首单里程（mi）",
      "熟手PPH（件）",
      "派送异常率",
      "DNR率",
      "收件地址类型占比",
    ];
    const lines = sortedTableRows.map((row) => {
      const property = propertyMap.get(row.route);
      return [
        row.week,
        row.region,
        row.route,
        row.site,
        row.dsp,
        row.delivered,
        row.attempted,
        row.totalHours.toFixed(2),
        row.operationPph.toFixed(2),
        row.successPph.toFixed(2),
        row.failRate.toFixed(4),
        row.wow === null ? "" : row.wow.toFixed(4),
        row.percentile,
        row.businessMode,
        row.isNew,
        property?.difficulty ?? "",
        property?.firstMile ?? "",
        property?.expertPph ?? "",
        property?.deliveryExceptionRate ?? "",
        property?.dnrRate ?? "",
        property?.addressMix ?? "",
      ]
        .map(csvEscape)
        .join(",");
    });
    const blob = new Blob(
      ["\uFEFF", headers.map(csvEscape).join(","), "\n", lines.join("\n")],
      { type: "text/csv;charset=utf-8" },
    );
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `PPH_${currentRegionName}_${selectedWeek}.csv`;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const selectedProperty = selectedRoute
    ? propertyMap.get(selectedRoute.route)
    : undefined;
  const selectedAnomalies = selectedRoute
    ? [
        selectedRoute.operationPph < quantiles.p25
          ? `作业PPH低于${currentRegionName}P25（${formatNumber(quantiles.p25, 2)}）`
          : "",
        selectedRoute.wow !== null && selectedRoute.wow < 0
          ? `较上周下降${formatPercent(Math.abs(selectedRoute.wow))}`
          : "",
        selectedRoute.failRate >= failP75 && selectedRoute.failRate > 0
          ? `派送失败率处于本区高位（${formatPercent(selectedRoute.failRate)}）`
          : "",
        selectedRoute.attempted >= volumeP75 &&
        selectedRoute.operationPph < quantiles.p25
          ? "高单量且效率落入本区P25以下"
          : "",
        selectedRoute.isNew.includes("新开") &&
        !selectedRoute.isNew.includes("非新开")
          ? "新开路区，纳入爬坡观察"
          : "",
      ].filter(Boolean)
    : [];
  const similarRoutes = useMemo(() => {
    if (!selectedRoute) return [];
    const sourceProperty = propertyMap.get(selectedRoute.route);
    return routeRows
      .filter((row) => rowKey(row) !== rowKey(selectedRoute))
      .map((row) => {
        const property = propertyMap.get(row.route);
        const volumeGap =
          Math.abs(row.attempted - selectedRoute.attempted) /
          Math.max(1, selectedRoute.attempted);
        const densityGap =
          sourceProperty?.populationDensity && property?.populationDensity
            ? Math.abs(
                property.populationDensity - sourceProperty.populationDensity,
              ) / sourceProperty.populationDensity
            : 0.5;
        const modeGap =
          sourceProperty?.businessMode &&
          property?.businessMode === sourceProperty.businessMode
            ? 0
            : 0.2;
        return { row, score: volumeGap + densityGap + modeGap };
      })
      .sort((a, b) => a.score - b.score)
      .slice(0, 3);
  }, [propertyMap, routeRows, selectedRoute]);

  const activeWeekStart = currentRecords.find(
    (row) => row.week === selectedWeek,
  )?.weekStart;
  const displayDate = activeWeekStart
    ? new Intl.DateTimeFormat("zh-CN", {
        month: "long",
        day: "numeric",
      }).format(new Date(activeWeekStart))
    : "";
  const availableWeekText =
    weeklyTrend.length < 8
      ? `当前文件含 ${weeklyTrend.length} 周`
      : "近 8 周";

  if (loading) {
    return (
      <main className="loading-screen">
        <div className="loading-mark">PPH</div>
        <div>
          <strong>正在汇总路区表现</strong>
          <span>计算大区分位数与异常名单…</span>
        </div>
      </main>
    );
  }

  return (
    <div className="app-shell">
      <aside className={`sidebar ${sidebarOpen ? "sidebar-open" : ""}`}>
        <div className="brand">
          <div className="brand-mark">P</div>
          <div>
            <strong>PPH周报</strong>
            <span>运营效能系统</span>
          </div>
          <button
            className="sidebar-close"
            onClick={() => setSidebarOpen(false)}
            aria-label="关闭导航"
          >
            <PanelLeftClose size={18} />
          </button>
        </div>
        <div className="nav-label">周报导航</div>
        <nav className="main-nav" aria-label="主导航">
          {NAV_ITEMS.map((item, index) => {
            const Icon = item.icon;
            const regionBoundary = index === 1;
            const actionBoundary = item.id === "exceptions";
            return (
              <div key={item.id}>
                {regionBoundary ? (
                  <div className="nav-group-label">大区周报</div>
                ) : null}
                {actionBoundary ? (
                  <div className="nav-group-label">运营管理</div>
                ) : null}
                <button
                  className={activeNav === item.id ? "nav-active" : ""}
                  onClick={() => navigate(item.id)}
                >
                  <Icon size={17} />
                  <span>{item.label}</span>
                  {REGION_OPTIONS.some(
                    (region) => region.code === item.id,
                  ) ? (
                    <ChevronRight size={14} className="nav-chevron" />
                  ) : null}
                </button>
              </div>
            );
          })}
        </nav>
        <div className="sidebar-source">
          <div className="source-status">
            <span className="status-dot" />
            数据已就绪
          </div>
          <strong>
            {formatNumber(sourceMeta?.sourceRows ?? records.length)} 条运营记录
          </strong>
          <span>
            {formatNumber(properties.length)} 条属性 · {weeks.length} 个周次
          </span>
        </div>
      </aside>

      {sidebarOpen ? (
        <button
          className="sidebar-backdrop"
          onClick={() => setSidebarOpen(false)}
          aria-label="关闭导航遮罩"
        />
      ) : null}

      <main className="main-content">
        <header className="topbar">
          <div className="topbar-left">
            <button
              className="menu-button"
              onClick={() => setSidebarOpen(true)}
              aria-label="打开导航"
            >
              <Menu size={20} />
            </button>
            <div>
              <div className="breadcrumb">
                PPH周报系统 <ChevronRight size={13} />{" "}
                <span>{currentRegionName}</span>
              </div>
              <h1>
                {activeRegion === "ALL"
                  ? "全国配送效率总览"
                  : `${currentRegionName} PPH周报`}
              </h1>
            </div>
          </div>
          <div className="topbar-actions">
            <div className="week-badge">
              <span>数据周期</span>
              <strong>
                {selectedWeek} {displayDate ? `· ${displayDate}起` : ""}
              </strong>
            </div>
            <button className="icon-button" aria-label="异常提醒">
              <BellRing size={18} />
              <span className="notification-dot" />
            </button>
            <div className="avatar">运</div>
          </div>
        </header>

        <div className="dashboard-content">
          <section className="data-toolbar">
            <div className="toolbar-title">
              <div className="toolbar-icon">
                <FileSpreadsheet size={20} />
              </div>
              <div>
                <strong>数据工作区</strong>
                <span>
                  按“路区名称”自动关联 · 仅保留配送量≥100 · 已剔除{" "}
                  {formatNumber(excludedCount)} 条未达门槛、异常或空载记录
                </span>
              </div>
            </div>
            <div className="upload-actions">
              <input
                ref={performanceInput}
                type="file"
                accept=".xlsx,.xls,.csv"
                hidden
                onChange={(event) => handleUpload(event, "performance")}
              />
              <input
                ref={propertyInput}
                type="file"
                accept=".xlsx,.xls,.csv"
                hidden
                onChange={(event) => handleUpload(event, "property")}
              />
              <button
                className="secondary-button"
                onClick={() => performanceInput.current?.click()}
                disabled={uploading !== null}
              >
                {uploading === "performance" ? (
                  <RefreshCw className="spin" size={16} />
                ) : (
                  <Upload size={16} />
                )}
                上传运营数据
              </button>
              <button
                className="primary-button"
                onClick={() => propertyInput.current?.click()}
                disabled={uploading !== null}
              >
                {uploading === "property" ? (
                  <RefreshCw className="spin" size={16} />
                ) : (
                  <Layers3 size={16} />
                )}
                上传地址/难易度
              </button>
            </div>
            {notice ? (
              <div className="toolbar-notice">
                <CheckCircle2 size={15} />
                {notice}
                <button onClick={() => setNotice("")} aria-label="关闭提示">
                  <X size={14} />
                </button>
              </div>
            ) : null}
          </section>

          <section className="filter-bar" aria-label="周报筛选">
            <div className="filter-lead">
              <Filter size={17} />
              <span>筛选</span>
            </div>
            <label>
              <span>周次</span>
              <div className="select-wrap">
                <select
                  value={selectedWeek}
                  onChange={(event) => setSelectedWeek(event.target.value)}
                >
                  {weeks.map((week) => (
                    <option key={week}>{week}</option>
                  ))}
                </select>
                <ChevronDown size={14} />
              </div>
            </label>
            <label>
              <span>站点</span>
              <div className="select-wrap">
                <select
                  value={siteFilter}
                  onChange={(event) => {
                    setSiteFilter(event.target.value);
                    setDspFilter("全部DSP");
                  }}
                >
                  <option>全部站点</option>
                  {siteOptions.map((site) => (
                    <option key={site}>{site}</option>
                  ))}
                </select>
                <ChevronDown size={14} />
              </div>
            </label>
            <label>
              <span>DSP</span>
              <div className="select-wrap">
                <select
                  value={dspFilter}
                  onChange={(event) => setDspFilter(event.target.value)}
                >
                  <option>全部DSP</option>
                  {dspOptions.map((dsp) => (
                    <option key={dsp}>{dsp}</option>
                  ))}
                </select>
                <ChevronDown size={14} />
              </div>
            </label>
            <label>
              <span>是否新开</span>
              <div className="select-wrap">
                <select
                  value={newFilter}
                  onChange={(event) => setNewFilter(event.target.value)}
                >
                  <option>全部</option>
                  {newOptions.map((option) => (
                    <option key={option}>{option}</option>
                  ))}
                </select>
                <ChevronDown size={14} />
              </div>
            </label>
            <label>
              <span>业务模式</span>
              <div className="select-wrap">
                <select
                  value={businessFilter}
                  onChange={(event) => setBusinessFilter(event.target.value)}
                >
                  <option>全部模式</option>
                  {businessOptions.map((option) => (
                    <option key={option}>{option}</option>
                  ))}
                </select>
                <ChevronDown size={14} />
              </div>
            </label>
          </section>

          <section className="scope-summary">
            <div>
              <span className="scope-kicker">
                <span />
                CURRENT SCOPE
              </span>
              <strong>
                {currentRegionName} · {selectedWeek}
              </strong>
              <p>
                {siteFilter} · {dspFilter} · {businessFilter}
              </p>
            </div>
            <div className="scope-counts">
              <div>
                <strong>{formatNumber(routeRows.length)}</strong>
                <span>路区组合</span>
              </div>
              <div>
                <strong>
                  {formatNumber(
                    new Set(currentRecords.map((row) => row.site)).size,
                  )}
                </strong>
                <span>站点</span>
              </div>
              <div>
                <strong>
                  {formatNumber(
                    new Set(currentRecords.map((row) => row.dsp)).size,
                  )}
                </strong>
                <span>DSP</span>
              </div>
            </div>
          </section>

          <section className="metrics-grid" aria-label="核心指标">
            <MetricCard
              label="作业PPH"
              value={formatNumber(currentMetrics.operationPph, 2)}
              detail="配送量（含失败）÷ 总时长"
              icon={Gauge}
              change={wow}
              tone={wow === null ? "blue" : wow >= 0 ? "green" : "red"}
            />
            <MetricCard
              label="成功PPH"
              value={formatNumber(currentMetrics.successPph, 2)}
              detail="成功配送量 ÷ 总时长"
              icon={PackageCheck}
              tone="green"
            />
            <MetricCard
              label="派送失败率"
              value={formatPercent(currentMetrics.failRate)}
              detail={`${formatNumber(Math.max(0, currentMetrics.attempted - currentMetrics.delivered))} 单失败`}
              icon={AlertTriangle}
              tone={currentMetrics.failRate > failP75 ? "orange" : "slate"}
            />
            <MetricCard
              label="近4周中位数"
              value={formatNumber(recentMedian, 2)}
              detail={`${weeklyTrend.slice(-4).length} 个可用周次的作业PPH`}
              icon={Activity}
              tone="blue"
            />
            <MetricCard
              label={`${currentRegionName}分位数`}
              value={`P50 ${formatNumber(quantiles.p50, 2)}`}
              detail={`P25 ${formatNumber(quantiles.p25, 2)} · P75 ${formatNumber(quantiles.p75, 2)}`}
              icon={BarChart3}
              tone="orange"
            />
            <MetricCard
              label="总配送量"
              value={formatNumber(currentMetrics.attempted)}
              detail={`成功 ${formatNumber(currentMetrics.delivered)} 单`}
              icon={Truck}
              tone="blue"
            />
            <MetricCard
              label="总工时"
              value={`${formatNumber(currentMetrics.totalHours)} h`}
              detail={`配送 ${formatNumber(currentMetrics.deliveryHours)} h`}
              icon={Clock3}
              tone="slate"
            />
          </section>

          <section className="chart-grid">
            <article className="panel panel-wide">
              <SectionHeader
                eyebrow="EFFICIENCY TREND"
                title="作业与成功 PPH 趋势"
                description={`${availableWeekText}实际数据；不设置目标线`}
                right={
                  <div className="quantile-legend">
                    本周处于
                    <strong>
                      {currentMetrics.operationPph < quantiles.p25
                        ? "P25以下"
                        : currentMetrics.operationPph < quantiles.p50
                          ? "P25–P50"
                          : currentMetrics.operationPph < quantiles.p75
                            ? "P50–P75"
                            : "P75以上"}
                    </strong>
                  </div>
                }
              />
              <Chart
                option={trendOption}
                height={315}
                ariaLabel="近八周作业PPH和成功PPH趋势图"
              />
            </article>
            <article className="panel">
              <SectionHeader
                eyebrow="TIME MIX"
                title="耗时结构"
                description="分拣、在途与配送耗时占比"
              />
              <Chart
                option={timeOption}
                height={315}
                ariaLabel="分拣在途配送耗时占比环形图"
              />
            </article>
            <article className="panel panel-wide">
              <SectionHeader
                eyebrow="NETWORK BENCHMARK"
                title={`${comparisonMode === "site" ? "站点" : "DSP"}效率对比`}
                description="按当前大区和筛选条件单独汇总"
                right={
                  <div className="segmented">
                    <button
                      className={comparisonMode === "site" ? "active" : ""}
                      onClick={() => setComparisonMode("site")}
                    >
                      站点
                    </button>
                    <button
                      className={comparisonMode === "dsp" ? "active" : ""}
                      onClick={() => setComparisonMode("dsp")}
                    >
                      DSP
                    </button>
                  </div>
                }
              />
              {comparisonRows.length ? (
                <Chart
                  option={comparisonOption}
                  height={330}
                  ariaLabel="各站点或DSP作业效率对比柱状图"
                />
              ) : (
                <EmptyState text="当前筛选下暂无站点或DSP数据" />
              )}
            </article>
            <article className="panel">
              <SectionHeader
                eyebrow="ROUTE CONTEXT"
                title="人口密度 × 作业PPH"
                description={`气泡大小代表配送量 · 已匹配 ${bubbleRows.length} 个路区组合`}
              />
              {bubbleRows.length ? (
                <Chart
                  option={bubbleOption}
                  height={330}
                  ariaLabel="人口密度与作业PPH气泡图"
                />
              ) : (
                <EmptyState text="上传含人口密度的路区属性后显示气泡图" />
              )}
            </article>
          </section>

          <section id="exceptions" className="exceptions-section">
            <SectionHeader
              eyebrow="PRIORITY WATCHLIST"
              title={`${currentRegionName}重点名单`}
              description="基于自身历史、本区分位数与同量级路区自动识别"
              right={
                <span className="method-tag">
                  <ShieldCheck size={14} /> 无目标值判定
                </span>
              }
            />
            <div className="watchlist-grid">
              {[
                {
                  id: "drops",
                  title: "本周PPH下降最多",
                  icon: TrendingDown,
                  tone: "red",
                  rows: priorityLists.drops,
                  value: (row: RouteRow) =>
                    row.wow === null ? "—" : `-${formatPercent(Math.abs(row.wow))}`,
                  caption: "较上周",
                },
                {
                  id: "high-volume-low",
                  title: "高单量低PPH",
                  icon: Truck,
                  tone: "orange",
                  rows: priorityLists.highVolumeLow,
                  value: (row: RouteRow) =>
                    formatNumber(row.operationPph, 2),
                  caption: "作业PPH",
                },
                {
                  id: "continuous",
                  title: "连续3周低于同类P25",
                  icon: Activity,
                  tone: "red",
                  rows: priorityLists.continuous,
                  value: (row: RouteRow) =>
                    formatNumber(row.operationPph, 2),
                  caption: "本周PPH",
                },
                {
                  id: "high-failure",
                  title: "高派送失败率",
                  icon: AlertTriangle,
                  tone: "orange",
                  rows: priorityLists.highFailure,
                  value: (row: RouteRow) => formatPercent(row.failRate),
                  caption: "失败率",
                },
                {
                  id: "new-routes",
                  title: "新开路区观察",
                  icon: Sparkles,
                  tone: "blue",
                  rows: priorityLists.newRoutes,
                  value: (row: RouteRow) =>
                    formatNumber(row.operationPph, 2),
                  caption: "作业PPH",
                },
              ].map((list) => {
                const Icon = list.icon;
                const isExpanded = Boolean(expandedWatchlists[list.id]);
                const hasMore = list.rows.length > 5;
                const visibleRows = isExpanded
                  ? list.rows
                  : list.rows.slice(0, 5);
                const toggleExpanded = () => {
                  if (!hasMore) return;
                  setExpandedWatchlists((current) => ({
                    ...current,
                    [list.id]: !current[list.id],
                  }));
                };
                return (
                  <article
                    className={`watch-card${isExpanded ? " watch-card-expanded" : ""}`}
                    key={list.id}
                  >
                    <button
                      type="button"
                      className="watch-card-head"
                      onClick={toggleExpanded}
                      aria-expanded={isExpanded}
                      aria-controls={`watchlist-${list.id}`}
                      disabled={!hasMore}
                    >
                      <div className={`watch-icon watch-${list.tone}`}>
                        <Icon size={16} />
                      </div>
                      <div className="watch-card-title">
                        <strong>{list.title}</strong>
                        <span>{list.rows.length} 个路区</span>
                      </div>
                      {hasMore ? (
                        <ChevronDown
                          className={`watch-card-chevron${isExpanded ? " expanded" : ""}`}
                          size={16}
                        />
                      ) : null}
                    </button>
                    <div
                      className="watch-rows"
                      id={`watchlist-${list.id}`}
                    >
                      {list.rows.length ? (
                        visibleRows.map((row) => (
                          <button
                            key={rowKey(row)}
                            onClick={() => setSelectedRoute(row)}
                          >
                            <div>
                              <strong>{row.route}</strong>
                              <span>
                                {row.site} · {row.dsp}
                              </span>
                            </div>
                            <div className="watch-value">
                              <strong>{list.value(row)}</strong>
                              <span>{list.caption}</span>
                            </div>
                          </button>
                        ))
                      ) : (
                        <div className="watch-empty">
                          <CheckCircle2 size={18} />
                          当前条件下无命中路区
                        </div>
                      )}
                    </div>
                    {hasMore ? (
                      <button
                        type="button"
                        className="watch-card-toggle"
                        onClick={toggleExpanded}
                        aria-expanded={isExpanded}
                        aria-controls={`watchlist-${list.id}`}
                      >
                        {isExpanded
                          ? "收起名单"
                          : `查看全部 ${list.rows.length} 个路区`}
                        <ChevronDown
                          className={isExpanded ? "expanded" : ""}
                          size={14}
                        />
                      </button>
                    ) : null}
                  </article>
                );
              })}
            </div>
          </section>

          <section id="data" className="panel data-panel">
            <SectionHeader
              eyebrow="FULL DATASET"
              title={`${currentRegionName}全量路区数据`}
              description={`当前筛选共 ${formatNumber(sortedTableRows.length)} 条，可排序、搜索并导出`}
              right={
                <button className="secondary-button" onClick={exportRows}>
                  <Download size={16} />
                  导出当前结果
                </button>
              }
            />
            <div className="table-toolbar">
              <label className="search-box">
                <Search size={16} />
                <input
                  value={tableSearch}
                  onChange={(event) => setTableSearch(event.target.value)}
                  placeholder="搜索路区、站点、DSP或模式"
                />
                {tableSearch ? (
                  <button
                    onClick={() => setTableSearch("")}
                    aria-label="清空搜索"
                  >
                    <X size={14} />
                  </button>
                ) : null}
              </label>
              <div className="table-legend">
                <span>
                  <i className="legend-dot legend-green" /> 改善
                </span>
                <span>
                  <i className="legend-dot legend-red" /> 下降
                </span>
                <span>
                  <i className="legend-dot legend-orange" /> 关注
                </span>
              </div>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>
                      <button onClick={() => toggleSort("route")}>
                        路区 <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>站点 / DSP</th>
                    <th>
                      <button onClick={() => toggleSort("attempted")}>
                        配送量 <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>
                      <button onClick={() => toggleSort("operationPph")}>
                        作业PPH <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>
                      <button onClick={() => toggleSort("successPph")}>
                        成功PPH <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>
                      <button onClick={() => toggleSort("failRate")}>
                        失败率 <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>
                      <button onClick={() => toggleSort("wow")}>
                        周环比 <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>本区位置</th>
                    <th>路区难易度</th>
                    <th>收件地址占比</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((row) => (
                    <tr key={rowKey(row)}>
                      <td>
                        <button
                          className="route-link"
                          onClick={() => setSelectedRoute(row)}
                        >
                          <Route size={15} />
                          <span>
                            <strong>{row.route}</strong>
                            <small>{row.region}</small>
                          </span>
                        </button>
                      </td>
                      <td>
                        <strong>{row.site}</strong>
                        <small>{row.dsp}</small>
                      </td>
                      <td>
                        <strong>{formatNumber(row.attempted)}</strong>
                        <small>成功 {formatNumber(row.delivered)}</small>
                      </td>
                      <td>
                        <strong>{formatNumber(row.operationPph, 2)}</strong>
                      </td>
                      <td>{formatNumber(row.successPph, 2)}</td>
                      <td>
                        <span
                          className={
                            row.failRate >= failP75
                              ? "cell-warning"
                              : undefined
                          }
                        >
                          {formatPercent(row.failRate)}
                        </span>
                      </td>
                      <td>
                        {row.wow === null ? (
                          <span className="cell-muted">—</span>
                        ) : (
                          <span
                            className={
                              row.wow >= 0
                                ? "cell-positive"
                                : "cell-negative"
                            }
                          >
                            {row.wow >= 0 ? (
                              <ArrowUpRight size={13} />
                            ) : (
                              <ArrowDownRight size={13} />
                            )}
                            {formatPercent(Math.abs(row.wow))}
                          </span>
                        )}
                      </td>
                      <td>
                        <span
                          className={`percentile-pill ${
                            row.percentile === "< P25"
                              ? "percentile-low"
                              : row.percentile === "≥ P75"
                                ? "percentile-high"
                                : ""
                          }`}
                        >
                          {row.percentile}
                        </span>
                      </td>
                      <td>
                        <strong>
                          {propertyMap.get(row.route)?.difficulty || "未标注"}
                        </strong>
                        <small>
                          {propertyMap.get(row.route)?.firstMile
                            ? `首单 ${formatNumber(
                                propertyMap.get(row.route)?.firstMile ?? 0,
                                1,
                              )} mi`
                            : `${row.businessMode} · ${row.isNew}`}
                        </small>
                      </td>
                      <td className="address-cell">
                        {addressMixSummary(
                          propertyMap.get(row.route)?.addressMix,
                        )}
                      </td>
                      <td>
                        <button
                          className="row-open"
                          onClick={() => setSelectedRoute(row)}
                          aria-label={`查看${row.route}详情`}
                        >
                          <ChevronRight size={16} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!visibleRows.length ? (
                <EmptyState text="当前筛选下没有可显示的数据" />
              ) : null}
            </div>
            <div className="pagination">
              <span>
                第 {page} / {pageCount} 页
              </span>
              <div>
                <button
                  onClick={() => setPage((current) => Math.max(1, current - 1))}
                  disabled={page === 1}
                  aria-label="上一页"
                >
                  <ArrowLeft size={15} />
                </button>
                <button
                  onClick={() =>
                    setPage((current) => Math.min(pageCount, current + 1))
                  }
                  disabled={page === pageCount}
                  aria-label="下一页"
                >
                  <ArrowRight size={15} />
                </button>
              </div>
            </div>
          </section>

          <footer className="dashboard-footer">
            <span>PPH周报系统 · 指标按当前大区独立计算</span>
            <span>异常依据：自身历史 · 大区分位数 · 相似路区</span>
          </footer>
        </div>
      </main>

      {selectedRoute ? (
        <div className="drawer-layer">
          <button
            className="drawer-backdrop"
            onClick={() => setSelectedRoute(null)}
            aria-label="关闭路区详情"
          />
          <aside className="route-drawer" aria-label="路区详情">
            <div className="drawer-head">
              <div>
                <span>{selectedRoute.region} · {selectedRoute.week}</span>
                <h2>{selectedRoute.route}</h2>
                <p>
                  {selectedRoute.site} · {selectedRoute.dsp}
                </p>
              </div>
              <button
                onClick={() => setSelectedRoute(null)}
                aria-label="关闭路区详情"
              >
                <X size={19} />
              </button>
            </div>
            <div className="drawer-scroll">
              <div className="drawer-metrics">
                <div>
                  <span>作业PPH</span>
                  <strong>{formatNumber(selectedRoute.operationPph, 2)}</strong>
                  <small>{selectedRoute.percentile}</small>
                </div>
                <div>
                  <span>成功PPH</span>
                  <strong>{formatNumber(selectedRoute.successPph, 2)}</strong>
                  <small>{formatNumber(selectedRoute.delivered)} 单</small>
                </div>
                <div>
                  <span>失败率</span>
                  <strong>{formatPercent(selectedRoute.failRate)}</strong>
                  <small>
                    {formatNumber(
                      selectedRoute.attempted - selectedRoute.delivered,
                    )}{" "}
                    单
                  </small>
                </div>
              </div>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <AlertTriangle size={16} />
                  <strong>异常原因</strong>
                </div>
                {selectedAnomalies.length ? (
                  <ul className="reason-list">
                    {selectedAnomalies.map((reason) => (
                      <li key={reason}>
                        <span />
                        {reason}
                      </li>
                    ))}
                  </ul>
                ) : (
                  <div className="drawer-good">
                    <CheckCircle2 size={17} />
                    当前规则下未识别到重点异常
                  </div>
                )}
              </section>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Building2 size={16} />
                  <strong>路区难易度</strong>
                </div>
                <div className="property-grid">
                  <div>
                    <span>难易度</span>
                    <strong>{selectedProperty?.difficulty || "未标注"}</strong>
                  </div>
                  <div>
                    <span>首单里程</span>
                    <strong>
                      {selectedProperty?.firstMile
                        ? `${formatNumber(selectedProperty.firstMile, 1)} mi`
                        : "未标注"}
                    </strong>
                  </div>
                  <div>
                    <span>熟手PPH</span>
                    <strong>
                      {selectedProperty?.expertPph
                        ? `${formatNumber(selectedProperty.expertPph, 1)} 件`
                        : "未标注"}
                    </strong>
                  </div>
                  <div>
                    <span>派送异常率</span>
                    <strong>
                      {selectedProperty?.deliveryExceptionRate
                        ? formatPercent(
                            selectedProperty.deliveryExceptionRate,
                            2,
                          )
                        : "未标注"}
                    </strong>
                  </div>
                  <div>
                    <span>DNR率</span>
                    <strong>
                      {selectedProperty?.dnrRate
                        ? formatPercent(selectedProperty.dnrRate, 2)
                        : "未标注"}
                    </strong>
                  </div>
                  <div>
                    <span>安全度</span>
                    <strong>{selectedProperty?.safety || "未标注"}</strong>
                  </div>
                </div>
              </section>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <UsersRound size={16} />
                  <strong>收件地址类型占比</strong>
                </div>
                {selectedProperty?.addressMix ? (
                  <div className="address-bars">
                    {addressMixItems(selectedProperty.addressMix)
                      .sort((a, b) => b.value - a.value)
                      .slice(0, 7)
                      .map((item) => (
                        <div key={item.name}>
                          <div>
                            <span>{item.name}</span>
                            <strong>{formatNumber(item.value, 1)}%</strong>
                          </div>
                          <div className="address-track">
                            <span
                              style={{
                                width: `${Math.min(100, item.value)}%`,
                              }}
                            />
                          </div>
                        </div>
                      ))}
                  </div>
                ) : (
                  <p className="missing-copy">
                    当前属性文件未包含该路区的地址类型明细。
                  </p>
                )}
              </section>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Zap size={16} />
                  <strong>相似路区对比</strong>
                </div>
                <div className="similar-list">
                  {similarRoutes.map(({ row }) => (
                    <button
                      key={rowKey(row)}
                      onClick={() => setSelectedRoute(row)}
                    >
                      <div>
                        <strong>{row.route}</strong>
                        <span>
                          {row.site} · 配送量 {formatNumber(row.attempted)}
                        </span>
                      </div>
                      <div>
                        <strong>{formatNumber(row.operationPph, 2)}</strong>
                        <span>作业PPH</span>
                      </div>
                    </button>
                  ))}
                </div>
              </section>
            </div>
          </aside>
        </div>
      ) : null}
    </div>
  );
}
