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
import * as XLSX from "xlsx";
import ReactECharts from "echarts-for-react";
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
  CalendarRange,
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
  Truck,
  Upload,
  X,
  Zap,
} from "lucide-react";
import {
  REGION_OPTIONS,
  addressMixItems,
  aggregatePerformance,
  aggregatePostalRows,
  aggregateRouteRows,
  buildPropertyMap,
  cleanPerformanceRecords,
  cleanPostalPerformanceRecords,
  csvEscape,
  formatNumber,
  formatPercent,
  imputeTransitHours,
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
  PostalCost,
  PostalPerformanceRecord,
  PostalProperty,
  PostalRow,
  RouteProperty,
  RouteRow,
} from "./types";

declare global {
  interface Window {
    __PPH_INITIAL_DATA__?: InitialData;
    __PPH_STANDALONE__?: boolean;
    __PPH_LOCKED_REGION__?: RegionCode;
  }
}

type RegionCode = "ALL" | (typeof REGION_OPTIONS)[number]["code"];
type SortKey =
  | "attempted"
  | "operationPph"
  | "successPph"
  | "failRate"
  | "wow"
  | "route";
type PostalSortKey =
  | "attempted"
  | "operationPph"
  | "successPph"
  | "failRate"
  | "totalHours"
  | "postalCode";

type RouteWatchContext = {
  id: string;
  title: string;
  caption: string;
};

const postalRowKey = (
  row: Pick<PostalRow, "postalCode" | "site" | "dsp" | "route">,
) => `${row.postalCode}¦${row.site}¦${row.dsp}¦${row.route ?? ""}`;

const FUNCTION_NAV_ITEMS = [
  { id: "monthly", label: "PPH月报系统", icon: CalendarRange },
  { id: "route-watchlist", label: "路区重点名单", icon: Route },
  { id: "postal-watchlist", label: "邮编重点名单", icon: MapPinned },
  { id: "data", label: "路区全量数据", icon: Database },
  { id: "postal", label: "邮编全量数据", icon: FileSpreadsheet },
];

const REGION_NAV_ITEMS = [
  { id: "overview", label: "全国总览", icon: LayoutDashboard },
  ...REGION_OPTIONS.map((region) => ({
    id: region.code,
    label: `${region.code} · ${region.name}周报`,
    icon: MapPinned,
  })),
  { id: "exceptions", label: "重点异常", icon: AlertTriangle },
  { id: "postal", label: "邮编数据", icon: MapPinned },
  { id: "data", label: "全量数据", icon: Database },
];

const DIFFICULT_ADDRESS_KEYWORDS = [
  "商业",
  "公寓",
  "学校",
  "山区",
  "山地",
  "农场",
  "工业园区",
  "仓库",
  "医院",
  "政府机构",
];

function addressDistributionDifference(left: string, right: string) {
  const leftItems = addressMixItems(left);
  const rightItems = addressMixItems(right);
  if (!leftItems.length || !rightItems.length) return 0.5;
  const leftMap = new Map(leftItems.map((item) => [item.name, item.value]));
  const rightMap = new Map(rightItems.map((item) => [item.name, item.value]));
  const names = new Set([...leftMap.keys(), ...rightMap.keys()]);
  const weightedDifference = [...names].reduce((total, name) => {
    const isDifficult = DIFFICULT_ADDRESS_KEYWORDS.some((keyword) =>
      name.includes(keyword),
    );
    const weight = isDifficult ? 1.5 : 1;
    return (
      total +
      Math.abs((leftMap.get(name) ?? 0) - (rightMap.get(name) ?? 0)) *
        weight
    );
  }, 0);
  return weightedDifference / 200;
}

function buildDrawerTrendOption(
  history: Array<{ week: string; attempted: number; operationPph: number }>,
) {
  const averagePph = history.length
    ? sum(history.map((item) => item.operationPph)) / history.length
    : 0;
  return {
    animationDuration: 520,
    animationEasing: "cubicOut",
    color: ["#9fc2ff", "#2563eb"],
    grid: { top: 44, right: 48, bottom: 32, left: 42 },
    legend: {
      top: 4,
      right: 2,
      itemWidth: 10,
      itemHeight: 7,
      textStyle: { color: "#64748b", fontSize: 8 },
      data: ["妥投量", "妥投PPH"],
    },
    tooltip: {
      trigger: "axis",
      backgroundColor: "rgba(7, 26, 58, 0.94)",
      borderWidth: 0,
      padding: [8, 10],
      textStyle: { color: "#fff", fontSize: 9 },
      axisPointer: {
        type: "line",
        lineStyle: { color: "#9eb0c9", width: 1, type: "dashed" },
      },
      formatter: (
        params: Array<{
          axisValue: string;
          marker: string;
          seriesName: string;
          value: number;
        }>,
      ) => {
        if (!params.length) return "";
        return [
          `<strong>${params[0].axisValue}</strong>`,
          ...params.map(
            (item) =>
              `${item.marker}${item.seriesName}：<strong>${
                item.seriesName === "妥投PPH"
                  ? formatNumber(item.value, 2)
                  : `${formatNumber(item.value)} 单`
              }</strong>`,
          ),
        ].join("<br/>");
      },
    },
    xAxis: {
      type: "category",
      data: history.map((item) => item.week),
      axisTick: { show: false },
      axisLine: { lineStyle: { color: "#dbe4f0" } },
      axisLabel: { color: "#64748b", fontSize: 8, margin: 10 },
    },
    yAxis: [
      {
        type: "value",
        name: "PPH",
        interval: 0.5,
        min: ({ min }: { min: number }) =>
          Math.max(0, Math.floor((min - 0.5) * 2) / 2),
        max: ({ max }: { max: number }) =>
          Math.ceil((max + 0.5) * 2) / 2,
        nameTextStyle: { color: "#94a3b8", fontSize: 8 },
        splitLine: { lineStyle: { color: "#edf1f6", type: "dashed" } },
        axisLabel: { color: "#94a3b8", fontSize: 8 },
      },
      {
        type: "value",
        name: "妥投量",
        nameTextStyle: { color: "#94a3b8", fontSize: 8 },
        splitLine: { show: false },
        axisLabel: {
          color: "#94a3b8",
          fontSize: 8,
          formatter: (value: number) => formatNumber(value),
        },
      },
    ],
    series: [
      {
        name: "妥投量",
        type: "bar",
        yAxisIndex: 1,
        data: history.map((item) => item.attempted),
        barMaxWidth: 24,
        itemStyle: {
          color: "rgba(102, 159, 255, 0.4)",
          borderColor: "rgba(72, 132, 238, 0.32)",
          borderWidth: 1,
          borderRadius: [4, 4, 0, 0],
        },
        emphasis: { itemStyle: { color: "rgba(79, 143, 255, 0.65)" } },
      },
      {
        name: "妥投PPH",
        type: "line",
        data: history.map((item) => Number(item.operationPph.toFixed(2))),
        smooth: 0.28,
        symbol: "circle",
        symbolSize: 7,
        lineStyle: { width: 2.5, color: "#2563eb" },
        itemStyle: { color: "#fff", borderColor: "#2563eb", borderWidth: 2.5 },
        areaStyle: {
          color: {
            type: "linear",
            x: 0,
            y: 0,
            x2: 0,
            y2: 1,
            colorStops: [
              { offset: 0, color: "rgba(37, 99, 235, 0.14)" },
              { offset: 1, color: "rgba(37, 99, 235, 0.01)" },
            ],
          },
        },
        markLine: averagePph
          ? {
              silent: true,
              symbol: "none",
              label: {
                formatter: `均值 ${formatNumber(averagePph, 2)}`,
                color: "#64748b",
                fontSize: 7,
                position: "insideEndTop",
              },
              lineStyle: { color: "#94a3b8", type: "dashed", width: 1 },
              data: [{ yAxis: Number(averagePph.toFixed(2)) }],
            }
          : undefined,
        z: 3,
      },
    ],
  };
}

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

function normalizePostalCode(value: unknown) {
  const text = normalizeText(value);
  return /^\d+$/.test(text) ? text.padStart(5, "0") : text;
}

function expandWorksheetRange(sheet: XLSX.WorkSheet) {
  const addresses = Object.keys(sheet).filter((key) => !key.startsWith("!"));
  if (!addresses.length) return;
  const range = addresses.reduce(
    (current, address) => {
      const cell = XLSX.utils.decode_cell(address);
      current.s.r = Math.min(current.s.r, cell.r);
      current.s.c = Math.min(current.s.c, cell.c);
      current.e.r = Math.max(current.e.r, cell.r);
      current.e.c = Math.max(current.e.c, cell.c);
      return current;
    },
    {
      s: { r: Number.POSITIVE_INFINITY, c: Number.POSITIVE_INFINITY },
      e: { r: 0, c: 0 },
    },
  );
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function expandWorksheetMerges(sheet: XLSX.WorkSheet) {
  for (const merge of sheet["!merges"] ?? []) {
    const source = sheet[XLSX.utils.encode_cell(merge.s)];
    if (!source) continue;
    for (let row = merge.s.r; row <= merge.e.r; row += 1) {
      for (let column = merge.s.c; column <= merge.e.c; column += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: column });
        if (!sheet[address]) sheet[address] = { ...source };
      }
    }
  }
}

async function readWorkbookRows(file: File) {
  const isCsv = file.name.toLowerCase().endsWith(".csv");
  const workbook = isCsv
    ? XLSX.read(await file.text(), {
        type: "string",
        cellDates: true,
        raw: true,
      })
    : XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  expandWorksheetMerges(sheet);
  expandWorksheetRange(sheet);
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: null,
  });
}

function deriveProfileHours(
  row: Record<string, unknown>,
  attempted: number,
) {
  const sourcePph = normalizeNumber(row["PPH"]);
  const perHundredDeliveryMinutes = normalizeNumber(
    row["百单派件时长(min)"],
  );
  const perHundredSortMinutes = normalizeNumber(row["百单分拣时长(min)"]);
  const averageDeliveryHours = normalizeNumber(row["司机派件时长(h)"]);
  const averageSortHours = normalizeNumber(row["司机分拣时长(h)"]);
  const averageTransitHours = normalizeNumber(row["司机行驶时长(h)"]);
  const deliveryHours =
    perHundredDeliveryMinutes > 0
      ? (perHundredDeliveryMinutes * attempted) / 6000
      : sourcePph > 0
        ? attempted / sourcePph
        : 0;
  const shiftScale =
    deliveryHours > 0 && averageDeliveryHours > 0
      ? deliveryHours / averageDeliveryHours
      : 0;
  const sortHours =
    shiftScale > 0
      ? averageSortHours * shiftScale
      : perHundredSortMinutes > 0
        ? (perHundredSortMinutes * attempted) / 6000
        : 0;
  const transitHours = shiftScale > 0 ? averageTransitHours * shiftScale : 0;
  return {
    sortHours,
    transitHours,
    deliveryHours,
    totalHours: sortHours + transitHours + deliveryHours,
  };
}

function isWeeklyProfileRows(rows: Record<string, unknown>[]) {
  const sample = rows.find((row) => normalizeText(row["路区名称"]));
  return Boolean(
    sample &&
      "大区编码" in sample &&
      "站点名称" in sample &&
      "DSP名称" in sample &&
      "邮编" in sample &&
      "妥投量" in sample,
  );
}

function normalizeWeeklyProfileRows(
  rows: Record<string, unknown>[],
  week: string,
  weekStart: string,
) {
  const records: PerformanceRecord[] = [];
  const postalRecords: PostalPerformanceRecord[] = [];
  rows.forEach((row) => {
    const regionCode = normalizeText(row["大区编码"]).toUpperCase();
    if (!regionCode || regionCode === "总和") return;
    const region = REGION_OPTIONS.find(
      (option) => option.code === regionCode,
    )?.source;
    const site = normalizeText(row["站点名称"]);
    const dsp = normalizeText(row["DSP名称"]);
    const route = normalizeText(row["路区名称"]);
    const postalCode = normalizePostalCode(row["邮编"]);
    if (!region || !site || !dsp || !route || !postalCode) return;
    const delivered = normalizeNumber(row["妥投量"]);
    const attempted = delivered;
    const hours = deriveProfileHours(row, attempted);
    const base = {
      week,
      weekStart,
      dsp,
      site,
      delivered,
      attempted,
      ...hours,
    };
    records.push({ ...base, route, region });
    postalRecords.push({
      ...base,
      postalCode,
      route,
      region: regionCode,
    });
  });
  return { records, postalRecords };
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
        delivered: normalizeNumber(
          row["妥投量"] ?? row["配送量"] ?? row["成功配送量"],
        ),
        attempted: normalizeNumber(
          row["妥投量"] ?? row["配送量"] ?? row["成功配送量"],
        ),
        sortHours: normalizeNumber(row["分拣耗时"]),
        transitHours: normalizeNumber(row["在途耗时"]),
        deliveryHours: normalizeNumber(row["配送耗时"]),
        totalHours: normalizeNumber(row["总时长"]),
      }),
    )
    .filter((row) => row.week && row.route && row.region);
}

function normalizePostalRows(rows: Record<string, unknown>[]) {
  return rows
    .map(
      (row): PostalPerformanceRecord => ({
        week: normalizeText(row["周数"] ?? row["周次"]),
        weekStart: normalizeDate(
          row["周开始日期"] ?? row["本周开始日期"] ?? row["日期"],
        ),
        postalCode: normalizePostalCode(row["邮编"] ?? row["收件邮编"]),
        dsp: normalizeText(row["DSP"] ?? row["车队名称"]),
        site: normalizeText(row["站点"] ?? row["转运站点"]),
        region: normalizeText(row["大区"]),
        delivered: normalizeNumber(
          row["妥投量"] ?? row["配送量"] ?? row["成功配送量"],
        ),
        attempted: normalizeNumber(
          row["妥投量"] ?? row["配送量"] ?? row["成功配送量"],
        ),
        sortHours: normalizeNumber(row["分拣耗时"]),
        transitHours: normalizeNumber(row["在途耗时"]),
        deliveryHours: normalizeNumber(row["配送耗时"]),
        totalHours: normalizeNumber(row["总时长"]),
      }),
    )
    .filter((row) => row.week && row.postalCode && row.region);
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
        routeUnitPrice: normalizeNumber(row["路区单均票价"]),
        routeHourlyWage: normalizeNumber(row["路区时薪"]),
        amazonHourlyMedian: normalizeNumber(
          row["Amazon Flex"] ?? row["亚马逊时薪"] ?? row["Amazon时薪"],
        ),
        salaryCity: normalizeText(row["调研城市名称"]),
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

type HomeProps = {
  reportVariant?: "weekly" | "monthly";
};

const PPH_SAVED_UPLOADS_KEY = "pph-dashboard-uploaded-data-v1";

type SavedUploads = {
  records?: PerformanceRecord[];
  postalRecords?: PostalPerformanceRecord[];
  properties?: RouteProperty[];
  updatedAt: string;
};

export default function Home({ reportVariant = "weekly" }: HomeProps = {}) {
  const reportLabel = reportVariant === "monthly" ? "月报" : "周报";
  const reportSystemName = `PPH${reportLabel}系统`;
  const lockedRegion =
    typeof window !== "undefined" ? window.__PPH_LOCKED_REGION__ : undefined;
  const standaloneMode =
    typeof window !== "undefined" && Boolean(window.__PPH_STANDALONE__);
  const initialRegion = lockedRegion ?? "WE";
  const [records, setRecords] = useState<PerformanceRecord[]>([]);
  const [postalRecords, setPostalRecords] = useState<
    PostalPerformanceRecord[]
  >([]);
  const [postalProperties, setPostalProperties] = useState<PostalProperty[]>(
    [],
  );
  const [postalCosts, setPostalCosts] = useState<PostalCost[]>([]);
  const [properties, setProperties] = useState<RouteProperty[]>([]);
  const [sourceMeta, setSourceMeta] = useState<InitialData["meta"] | null>(null);
  const [excludedCount, setExcludedCount] = useState(0);
  const [excludedPostalCount, setExcludedPostalCount] = useState(0);
  const [estimatedTransitCount, setEstimatedTransitCount] = useState(0);
  const [estimatedPostalTransitCount, setEstimatedPostalTransitCount] =
    useState(0);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState<
    "performance" | "property" | "postal" | null
  >(null);
  const [exportingHtml, setExportingHtml] = useState(false);
  const [notice, setNotice] = useState("");
  const [activeRegion, setActiveRegion] =
    useState<RegionCode>(initialRegion);
  const [activeNav, setActiveNav] = useState<string>("route-watchlist");
  const [sidebarMode, setSidebarMode] = useState<"function" | "region">(
    "function",
  );
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [selectedWeek, setSelectedWeek] = useState("");
  const [siteFilter, setSiteFilter] = useState("");
  const [routeFilter, setRouteFilter] = useState("");
  const [postalFilter, setPostalFilter] = useState("");
  const [dspFilter, setDspFilter] = useState("");
  const [newFilter, setNewFilter] = useState("全部");
  const [businessFilter, setBusinessFilter] = useState("全部模式");
  const [selectedRoute, setSelectedRoute] = useState<RouteRow | null>(null);
  const [selectedRouteContext, setSelectedRouteContext] =
    useState<RouteWatchContext | null>(null);
  const [selectedPostal, setSelectedPostal] = useState<PostalRow | null>(null);
  const [selectedPostalContext, setSelectedPostalContext] =
    useState<RouteWatchContext | null>(null);
  const [postalParentRoute, setPostalParentRoute] =
    useState<RouteRow | null>(null);
  const [postalParentRouteContext, setPostalParentRouteContext] =
    useState<RouteWatchContext | null>(null);
  const [expandedWatchlists, setExpandedWatchlists] = useState<
    Record<string, boolean>
  >({});
  const [tableSearch, setTableSearch] = useState("");
  const [tableSort, setTableSort] = useState<SortKey>("attempted");
  const [sortDirection, setSortDirection] = useState<"asc" | "desc">("desc");
  const [page, setPage] = useState(1);
  const [postalSearch, setPostalSearch] = useState("");
  const [postalSort, setPostalSort] =
    useState<PostalSortKey>("attempted");
  const [postalSortDirection, setPostalSortDirection] = useState<
    "asc" | "desc"
  >("desc");
  const [postalPage, setPostalPage] = useState(1);
  const performanceInput = useRef<HTMLInputElement>(null);
  const propertyInput = useRef<HTMLInputElement>(null);
  const postalInput = useRef<HTMLInputElement>(null);

  const saveUploadedData = (data: Omit<SavedUploads, "updatedAt">) => {
    const current = JSON.parse(
      localStorage.getItem(PPH_SAVED_UPLOADS_KEY) ?? "{}",
    ) as Partial<SavedUploads>;
    localStorage.setItem(
      PPH_SAVED_UPLOADS_KEY,
      JSON.stringify({ ...current, ...data, updatedAt: new Date().toISOString() }),
    );
  };

  useEffect(() => {
    const applyInitialData = (data: InitialData, saved?: SavedUploads) => {
      const effectiveRecords = saved?.records?.length
        ? saved.records
        : data.records;
      const effectivePostalRecords = saved?.postalRecords?.length
        ? saved.postalRecords
        : data.postalRecords ?? [];
      const effectiveProperties = saved?.properties?.length
        ? saved.properties
        : data.properties;
      const transitFilled = imputeTransitHours(effectiveRecords);
      const postalTransitFilled = imputeTransitHours(
        effectivePostalRecords,
      );
      const cleaned = cleanPerformanceRecords(transitFilled.records);
      const cleanedPostal = cleanPostalPerformanceRecords(
        postalTransitFilled.records,
      );
      setRecords(cleaned.records);
      setPostalRecords(cleanedPostal.records);
      setPostalProperties(data.postalProperties ?? []);
      setPostalCosts(data.postalCosts ?? []);
      setProperties(effectiveProperties);
      setExcludedCount(cleaned.excluded);
      setExcludedPostalCount(cleanedPostal.excluded);
      setEstimatedTransitCount(
        cleaned.records.filter((row) => row.transitHoursEstimated).length,
      );
      setEstimatedPostalTransitCount(
        cleanedPostal.records.filter((row) => row.transitHoursEstimated).length,
      );
      setSourceMeta({
        ...data.meta,
        aggregatedRows: cleaned.records.length,
        postalRows: cleanedPostal.records.length,
        estimatedTransitRows: cleaned.records.filter(
          (row) => row.transitHoursEstimated,
        ).length,
        estimatedPostalTransitRows: cleanedPostal.records.filter(
          (row) => row.transitHoursEstimated,
        ).length,
      });
      const dataWeeks = sortWeeks(cleaned.records.map((row) => row.week));
      setSelectedWeek(dataWeeks.at(-1) ?? "");
    };

    let savedUploads: SavedUploads | undefined;
    try {
      savedUploads = JSON.parse(
        localStorage.getItem(PPH_SAVED_UPLOADS_KEY) ?? "null",
      ) as SavedUploads | undefined;
    } catch {
      localStorage.removeItem(PPH_SAVED_UPLOADS_KEY);
    }

    if (window.__PPH_INITIAL_DATA__) {
      applyInitialData(window.__PPH_INITIAL_DATA__, savedUploads);
      setLoading(false);
      return;
    }

    Promise.all([
      fetch("/data/initial.json"),
      fetch("/data/postal-records.json"),
    ])
      .then(async ([initialResponse, postalResponse]) => {
        if (!initialResponse.ok || !postalResponse.ok)
          throw new Error("Initial data is not available.");
        const [data, postalData] = await Promise.all([
          initialResponse.json() as Promise<InitialData>,
          postalResponse.json() as Promise<{
            postalRecords: PostalPerformanceRecord[];
          }>,
        ]);
        return { ...data, postalRecords: postalData.postalRecords };
      })
      .then((data) => applyInitialData(data, savedUploads))
      .catch(() => {
        setNotice("请上传运营明细与路区属性文件开始分析。");
      })
      .finally(() => setLoading(false));

  }, []);

  useEffect(() => {
    const isFullPageCapture =
      new URLSearchParams(window.location.search).get("capture") === "full";
    if (!isFullPageCapture) return;
    document.documentElement.classList.add("capture-full");
    return () => document.documentElement.classList.remove("capture-full");
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
    (
      record: PerformanceRecord,
      ignoreWeek = false,
      ignoreRoute = false,
    ) => {
      if (currentRegionSource && record.region !== currentRegionSource)
        return false;
      if (!ignoreWeek && selectedWeek && record.week !== selectedWeek)
        return false;
      const siteQuery = siteFilter.trim().toLowerCase();
      const routeQuery = routeFilter.trim().toLowerCase();
      const dspQuery = dspFilter.trim().toLowerCase();
      if (siteQuery && !record.site.toLowerCase().includes(siteQuery))
        return false;
      if (
        !ignoreRoute &&
        routeQuery &&
        !record.route.toLowerCase().includes(routeQuery)
      )
        return false;
      if (dspQuery && !record.dsp.toLowerCase().includes(dspQuery))
        return false;
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
      routeFilter,
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
              (row) =>
                !siteFilter.trim() ||
                row.site
                  .toLowerCase()
                  .includes(siteFilter.trim().toLowerCase()),
            )
            .map((row) => row.dsp)
            .filter(Boolean),
        ),
      ].sort(),
    [regionRecords, siteFilter],
  );
  const routeOptions = useMemo(
    () =>
      [
        ...new Set(
          regionRecords
            .filter((row) => {
              const siteQuery = siteFilter.trim().toLowerCase();
              const dspQuery = dspFilter.trim().toLowerCase();
              if (selectedWeek && row.week !== selectedWeek) return false;
              if (
                siteQuery &&
                !row.site.toLowerCase().includes(siteQuery)
              )
                return false;
              if (dspQuery && !row.dsp.toLowerCase().includes(dspQuery))
                return false;
              return true;
            })
            .map((row) => row.route)
            .filter(Boolean),
        ),
      ].sort(),
    [dspFilter, regionRecords, selectedWeek, siteFilter],
  );
  const postalCodeOptions = useMemo(() => {
    const siteQuery = siteFilter.trim().toLowerCase();
    const routeQuery = routeFilter.trim().toLowerCase();
    const dspQuery = dspFilter.trim().toLowerCase();
    return [
      ...new Set(
        postalRecords
          .filter((row) => {
            if (activeRegion !== "ALL" && row.region !== activeRegion)
              return false;
            if (siteQuery && !row.site.toLowerCase().includes(siteQuery))
              return false;
            if (
              routeQuery &&
              !(row.route ?? "").toLowerCase().includes(routeQuery)
            )
              return false;
            if (dspQuery && !row.dsp.toLowerCase().includes(dspQuery))
              return false;
            return true;
          })
          .map((row) => row.postalCode)
          .filter(Boolean),
      ),
    ].sort();
  }, [activeRegion, dspFilter, postalRecords, routeFilter, siteFilter]);
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
  const routeSearchRows = useMemo(() => {
    const current = records.filter((row) => recordMatches(row, false, true));
    const previous = previousWeek
      ? records.filter(
          (row) =>
            row.week === previousWeek && recordMatches(row, true, true),
        )
      : [];
    return aggregateRouteRows(current, previous, properties);
  }, [previousWeek, properties, recordMatches, records]);
  const currentMetrics = useMemo(
    () => aggregatePerformance(currentRecords),
    [currentRecords],
  );
  const previousMetrics = useMemo(
    () => aggregatePerformance(previousRecords),
    [previousRecords],
  );
  const postalRecordMatches = useCallback(
    (row: PostalPerformanceRecord, ignoreWeek = false) => {
    const postalQuery = postalSearch.trim().toLowerCase();
    const postalScopeQuery = postalFilter.trim().toLowerCase();
    const siteQuery = siteFilter.trim().toLowerCase();
    const routeQuery = routeFilter.trim().toLowerCase();
    const dspQuery = dspFilter.trim().toLowerCase();
      if (activeRegion !== "ALL" && row.region !== activeRegion) return false;
      if (!ignoreWeek && selectedWeek && row.week !== selectedWeek) return false;
      if (siteQuery && !row.site.toLowerCase().includes(siteQuery)) return false;
      if (
        routeQuery &&
        !(row.route ?? "").toLowerCase().includes(routeQuery)
      )
        return false;
      if (dspQuery && !row.dsp.toLowerCase().includes(dspQuery)) return false;
      if (
        postalScopeQuery &&
        !row.postalCode.toLowerCase().includes(postalScopeQuery)
      )
        return false;
      if (
        postalQuery &&
        ![row.postalCode, row.site, row.dsp, row.route]
          .join(" ")
          .toLowerCase()
          .includes(postalQuery)
      )
        return false;
      return true;
    },
    [
      activeRegion,
      dspFilter,
      postalSearch,
      postalFilter,
      routeFilter,
      selectedWeek,
      siteFilter,
    ],
  );
  const currentPostalRecords = useMemo(
    () => postalRecords.filter((row) => postalRecordMatches(row)),
    [postalRecordMatches, postalRecords],
  );
  const postalRows = useMemo(
    () => aggregatePostalRows(currentPostalRecords),
    [currentPostalRecords],
  );
  const postalMetrics = useMemo(
    () =>
      aggregatePerformance(
        currentPostalRecords.map((row) => ({
          ...row,
          route: row.postalCode,
        })),
      ),
    [currentPostalRecords],
  );
  const postalCodeCount = useMemo(
    () => new Set(postalRows.map((row) => row.postalCode)).size,
    [postalRows],
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
  const scopeFocus = postalFilter.trim()
    ? "postal"
    : routeFilter.trim()
      ? "route"
      : siteFilter.trim()
        ? "site"
        : null;
  const scopeFocusLabel =
    scopeFocus === "postal"
      ? postalFilter.trim()
      : scopeFocus === "route"
        ? routeFilter.trim()
        : scopeFocus === "site"
          ? siteFilter.trim()
          : currentRegionName;
  const scopeTrend = useMemo(() => {
    if (!scopeFocus) return [];
    const siteQuery = siteFilter.trim().toLowerCase();
    const routeQuery = routeFilter.trim().toLowerCase();
    const postalQuery = postalFilter.trim().toLowerCase();
    const dspQuery = dspFilter.trim().toLowerCase();
    const trendWeeks = weeks
      .filter((week) => weeks.indexOf(week) <= currentWeekIndex)
      .slice(-4);

    return trendWeeks.map((week) => {
      if (scopeFocus === "postal") {
        const scopedPostalRecords = postalRecords.filter((row) => {
          if (row.week !== week) return false;
          if (activeRegion !== "ALL" && row.region !== activeRegion)
            return false;
          if (siteQuery && !row.site.toLowerCase().includes(siteQuery))
            return false;
          if (
            routeQuery &&
            !(row.route ?? "").toLowerCase().includes(routeQuery)
          )
            return false;
          if (dspQuery && !row.dsp.toLowerCase().includes(dspQuery))
            return false;
          if (
            postalQuery &&
            !row.postalCode.toLowerCase().includes(postalQuery)
          )
            return false;
          return true;
        });
        return {
          week,
          ...aggregatePerformance(
            scopedPostalRecords.map((row) => ({
              ...row,
              route: row.route ?? row.postalCode,
            })),
          ),
        };
      }

      const scopedRecords = records.filter((row) => {
        if (row.week !== week) return false;
        if (currentRegionSource && row.region !== currentRegionSource)
          return false;
        if (siteQuery && !row.site.toLowerCase().includes(siteQuery))
          return false;
        if (routeQuery && !row.route.toLowerCase().includes(routeQuery))
          return false;
        if (dspQuery && !row.dsp.toLowerCase().includes(dspQuery))
          return false;
        const property = propertyMap.get(row.route);
        if (
          businessFilter !== "全部模式" &&
          (property?.businessMode || "未标注") !== businessFilter
        )
          return false;
        if (
          newFilter !== "全部" &&
          (property?.isNew || "未标注") !== newFilter
        )
          return false;
        return true;
      });
      return { week, ...aggregatePerformance(scopedRecords) };
    });
  }, [
    activeRegion,
    businessFilter,
    currentRegionSource,
    currentWeekIndex,
    dspFilter,
    newFilter,
    postalFilter,
    postalRecords,
    propertyMap,
    records,
    routeFilter,
    scopeFocus,
    siteFilter,
    weeks,
  ]);
  const scopeTrendOption = useMemo(() => {
    const averagePph = scopeTrend.length
      ? sum(scopeTrend.map((item) => item.operationPph)) / scopeTrend.length
      : 0;
    return {
      animationDuration: 650,
      animationEasing: "cubicOut",
      color: ["#8bb6ff", "#2563eb"],
      grid: { top: 58, right: 62, bottom: 38, left: 56 },
      legend: {
        top: 8,
        right: 10,
        itemWidth: 12,
        itemHeight: 8,
        textStyle: { color: "#64748b", fontSize: 11 },
        data: ["妥投量", "妥投PPH"],
      },
      tooltip: {
        trigger: "axis",
        backgroundColor: "rgba(7, 26, 58, 0.94)",
        borderWidth: 0,
        padding: [10, 12],
        textStyle: { color: "#fff", fontSize: 11 },
        axisPointer: {
          type: "line",
          lineStyle: {
            color: "#b7c7df",
            width: 1,
            type: "dashed",
          },
        },
        formatter: (
          params: Array<{
            axisValue: string;
            marker: string;
            seriesName: string;
            value: number;
          }>,
        ) => {
          if (!params.length) return "";
          return [
            `<strong>${params[0].axisValue}</strong>`,
            ...params.map(
              (item) =>
                `${item.marker}${item.seriesName}：<strong>${
                  item.seriesName === "妥投PPH"
                    ? formatNumber(item.value, 2)
                    : formatNumber(item.value)
                }</strong>`,
            ),
          ].join("<br/>");
        },
      },
      xAxis: {
        type: "category",
        data: scopeTrend.map((item) => item.week),
        axisTick: { show: false },
        axisLine: { lineStyle: { color: "#dbe4f0" } },
        axisLabel: { color: "#64748b", fontSize: 11, margin: 13 },
      },
      yAxis: [
        {
          type: "value",
          name: "PPH",
          interval: 0.5,
          min: ({ min }: { min: number }) =>
            Math.max(0, Math.floor((min - 0.5) * 2) / 2),
          max: ({ max }: { max: number }) =>
            Math.ceil((max + 0.5) * 2) / 2,
          nameTextStyle: { color: "#94a3b8", fontSize: 10, padding: [0, 0, 4, 0] },
          splitLine: { lineStyle: { color: "#edf1f6", type: "dashed" } },
          axisLabel: { color: "#94a3b8", fontSize: 10 },
        },
        {
          type: "value",
          name: "妥投量",
          nameTextStyle: { color: "#94a3b8", fontSize: 10, padding: [0, 0, 4, 0] },
          splitLine: { show: false },
          axisLabel: {
            color: "#94a3b8",
            fontSize: 10,
            formatter: (value: number) => formatNumber(value),
          },
        },
      ],
      series: [
        {
          name: "妥投量",
          type: "bar",
          yAxisIndex: 1,
          data: scopeTrend.map((item) => item.attempted),
          barMaxWidth: 34,
          itemStyle: {
            color: "rgba(102, 159, 255, 0.42)",
            borderColor: "rgba(72, 132, 238, 0.3)",
            borderWidth: 1,
            borderRadius: [5, 5, 0, 0],
          },
          emphasis: { itemStyle: { color: "rgba(79, 143, 255, 0.68)" } },
        },
        {
          name: "妥投PPH",
          type: "line",
          data: scopeTrend.map((item) => Number(item.operationPph.toFixed(2))),
          smooth: 0.28,
          symbol: "circle",
          symbolSize: 9,
          showSymbol: true,
          lineStyle: { width: 3, color: "#2563eb" },
          itemStyle: { color: "#fff", borderColor: "#2563eb", borderWidth: 3 },
          areaStyle: {
            color: {
              type: "linear",
              x: 0,
              y: 0,
              x2: 0,
              y2: 1,
              colorStops: [
                { offset: 0, color: "rgba(37, 99, 235, 0.18)" },
                { offset: 1, color: "rgba(37, 99, 235, 0.01)" },
              ],
            },
          },
          markLine: averagePph
            ? {
                silent: true,
                symbol: "none",
                label: {
                  formatter: `四周均值 ${formatNumber(averagePph, 2)}`,
                  color: "#64748b",
                  fontSize: 9,
                  position: "insideEndTop",
                },
                lineStyle: { color: "#94a3b8", type: "dashed", width: 1 },
                data: [{ yAxis: Number(averagePph.toFixed(2)) }],
              }
            : undefined,
          z: 3,
        },
      ],
    };
  }, [scopeTrend]);
  const scopeSummaryInsight = useMemo(() => {
    const toneForChange = (change: number | null) =>
      change === null ? "blue" : change >= 0 ? "positive" : "negative";
    const signedPercent = (change: number | null) =>
      change === null
        ? "新进入"
        : `${change >= 0 ? "+" : ""}${formatPercent(change)}`;
    const fallback = (label: string) => [
      {
        label,
        value: "—",
        meta: "本周暂无可比数据",
        tone: "blue",
      },
    ];

    if (scopeFocus === "postal") {
      const latest = scopeTrend.at(-1);
      const previous = scopeTrend.at(-2);
      const pphChange =
        latest && previous && previous.operationPph > 0
          ? (latest.operationPph - previous.operationPph) /
            previous.operationPph
          : null;
      const volumeChange =
        latest && previous && previous.attempted > 0
          ? (latest.attempted - previous.attempted) / previous.attempted
          : null;
      return {
        title: "邮编运营概况",
        subtitle: `${scopeFocusLabel} · 本周表现与环比`,
        items: latest
          ? [
              {
                label: "本周妥投PPH",
                value: formatNumber(latest.operationPph, 2),
                meta: `${formatNumber(latest.attempted)} 单`,
                tone: "blue",
              },
              {
                label: "PPH环比",
                value: signedPercent(pphChange),
                meta: previous
                  ? `上周 ${formatNumber(previous.operationPph, 2)}`
                  : "首个可用周次",
                tone: toneForChange(pphChange),
              },
              {
                label: "妥投量环比",
                value: signedPercent(volumeChange),
                meta: previous
                  ? `上周 ${formatNumber(previous.attempted)} 单`
                  : "首个可用周次",
                tone: toneForChange(volumeChange),
              },
            ]
          : fallback("暂无邮编数据"),
      };
    }

    const siteQuery = siteFilter.trim().toLowerCase();
    const routeQuery = routeFilter.trim().toLowerCase();
    const dspQuery = dspFilter.trim().toLowerCase();
    if (scopeFocus === "route") {
      const matchesPostal = (row: PostalPerformanceRecord, week?: string) => {
        if (week && row.week !== week) return false;
        if (activeRegion !== "ALL" && row.region !== activeRegion) return false;
        if (siteQuery && !row.site.toLowerCase().includes(siteQuery)) return false;
        if (
          routeQuery &&
          !(row.route ?? "").toLowerCase().includes(routeQuery)
        )
          return false;
        if (dspQuery && !row.dsp.toLowerCase().includes(dspQuery)) return false;
        return true;
      };
      const current = postalRecords.filter((row) =>
        matchesPostal(row, selectedWeek),
      );
      const previous = previousWeek
        ? postalRecords.filter((row) => matchesPostal(row, previousWeek))
        : [];
      const codes = [...new Set(current.map((row) => row.postalCode))];
      const ranked = codes
        .map((code) => {
          const currentValue = aggregatePerformance(
            current
              .filter((row) => row.postalCode === code)
              .map((row) => ({ ...row, route: row.route ?? row.postalCode })),
          );
          const previousValue = aggregatePerformance(
            previous
              .filter((row) => row.postalCode === code)
              .map((row) => ({ ...row, route: row.route ?? row.postalCode })),
          );
          const change =
            previousValue.operationPph > 0
              ? (currentValue.operationPph - previousValue.operationPph) /
                previousValue.operationPph
              : null;
          const dailyExtra = Math.max(
            0,
            Math.round(
              ((currentValue.operationPph - previousValue.operationPph) *
                currentValue.totalHours) /
                7,
            ),
          );
          return { code, currentValue, change, dailyExtra };
        })
        .sort((a, b) => (b.change ?? -Infinity) - (a.change ?? -Infinity));
      const rising = ranked.filter((item) => (item.change ?? 0) > 0).slice(0, 3);
      return {
        title: "路区下邮编涨幅",
        subtitle: `${scopeFocusLabel} · 较上周提升最快`,
        items: rising.length
          ? rising.map((item) => ({
              label: item.code,
              value: signedPercent(item.change),
              meta: `${formatNumber(item.currentValue.operationPph, 2)} PPH`,
              dailyExtra: item.dailyExtra,
              tone: toneForChange(item.change),
            }))
          : fallback("暂无上涨邮编"),
      };
    }

    if (scopeFocus === "site") {
      const matchesRecord = (row: PerformanceRecord, week?: string) => {
        if (week && row.week !== week) return false;
        if (currentRegionSource && row.region !== currentRegionSource)
          return false;
        if (siteQuery && !row.site.toLowerCase().includes(siteQuery)) return false;
        if (dspQuery && !row.dsp.toLowerCase().includes(dspQuery)) return false;
        const property = propertyMap.get(row.route);
        if (
          businessFilter !== "全部模式" &&
          (property?.businessMode || "未标注") !== businessFilter
        )
          return false;
        if (
          newFilter !== "全部" &&
          (property?.isNew || "未标注") !== newFilter
        )
          return false;
        return true;
      };
      const current = records.filter((row) => matchesRecord(row, selectedWeek));
      const previous = previousWeek
        ? records.filter((row) => matchesRecord(row, previousWeek))
        : [];
      const routes = [...new Set(current.map((row) => row.route))];
      const ranked = routes
        .map((route) => {
          const currentValue = aggregatePerformance(
            current.filter((row) => row.route === route),
          );
          const previousValue = aggregatePerformance(
            previous.filter((row) => row.route === route),
          );
          const change =
            previousValue.operationPph > 0
              ? (currentValue.operationPph - previousValue.operationPph) /
                previousValue.operationPph
              : null;
          const dailyExtra = Math.max(
            0,
            Math.round(
              ((currentValue.operationPph - previousValue.operationPph) *
                currentValue.totalHours) /
                7,
            ),
          );
          return { route, currentValue, change, dailyExtra };
        })
        .sort((a, b) => (b.change ?? -Infinity) - (a.change ?? -Infinity));
      const rising = ranked.filter((item) => (item.change ?? 0) > 0).slice(0, 3);
      return {
        title: "站点内路区涨幅",
        subtitle: `${scopeFocusLabel} · 较上周提升最快`,
        items: rising.length
          ? rising.map((item) => ({
              label: item.route,
              value: signedPercent(item.change),
              meta: `${formatNumber(item.currentValue.operationPph, 2)} PPH`,
              dailyExtra: item.dailyExtra,
              tone: toneForChange(item.change),
            }))
          : fallback("暂无上涨路区"),
      };
    }

    return {
      title: "本周运营概况",
      subtitle: `${currentRegionName} · ${selectedWeek}`,
      items: [
        {
          label: "妥投PPH",
          value: formatNumber(currentMetrics.operationPph, 2),
          meta: "当前筛选口径",
          tone: "blue",
        },
        {
          label: "PPH环比",
          value: signedPercent(wow),
          meta: previousWeek ? `对比 ${previousWeek}` : "首个可用周次",
          tone: toneForChange(wow),
        },
        {
          label: "妥投单量",
          value: formatNumber(currentMetrics.delivered),
          meta: "当前筛选口径",
          tone: "blue",
        },
      ],
    };
  }, [
    activeRegion,
    businessFilter,
    currentMetrics.delivered,
    currentMetrics.operationPph,
    currentRegionName,
    currentRegionSource,
    dspFilter,
    newFilter,
    postalRecords,
    previousWeek,
    propertyMap,
    records,
    routeFilter,
    scopeFocus,
    scopeFocusLabel,
    scopeTrend,
    selectedWeek,
    siteFilter,
    wow,
  ]);
  const routePphValues = routeRows
    .map((row) => row.operationPph)
    .filter((value) => value > 0);
  const quantiles = {
    p25: percentile(routePphValues, 0.25),
    p50: percentile(routePphValues, 0.5),
    p75: percentile(routePphValues, 0.75),
  };
  const p75WatchRows = routeRows
    .filter((row) => row.operationPph >= quantiles.p75)
    .sort((a, b) => b.operationPph - a.operationPph);
  const p25WatchRows = routeRows
    .filter((row) => row.operationPph < quantiles.p25)
    .sort((a, b) => a.operationPph - b.operationPph);
  const monthlyDifficultyRows = useMemo(
    () =>
      properties
        .filter((item) => item.difficulty)
        .filter((item) => {
          const query = tableSearch.trim().toLowerCase();
          return query
            ? [item.route, item.transferSite, item.fleet, item.difficulty]
                .join(" ")
                .toLowerCase()
                .includes(query)
            : true;
        })
        .slice(0, 80),
    [properties, tableSearch],
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
    const buildConsecutiveIncrease = (weekCount: number) => {
      const increaseWeeks = weeks
        .filter((week) => weeks.indexOf(week) <= currentWeekIndex)
        .slice(-weekCount);
      const changeMap = new Map<string, number>();
      if (increaseWeeks.length === weekCount) {
        [...new Set(routeRows.map((row) => row.route))].forEach((route) => {
          const values = increaseWeeks.map((week) => {
            const matches = (weeklyRouteMap.get(week) ?? []).filter(
              (row) => row.route === route,
            );
            return pph(
              sum(matches.map((row) => row.attempted)),
              sum(matches.map((row) => row.totalHours)),
            );
          });
          const risesEveryWeek = values.every(
            (value, index) =>
              value > 0 &&
              (index === 0 ||
                (value - values[index - 1]) / values[index - 1] >= 0.001),
          );
          if (risesEveryWeek) {
            changeMap.set(
              route,
              (values.at(-1)! - values[0]) / values[0],
            );
          }
        });
      }
      const rows = routeRows
        .filter(
          (row, index, candidates) =>
            changeMap.has(row.route) &&
            candidates.findIndex(
              (candidate) => candidate.route === row.route,
            ) === index,
        )
        .sort(
          (a, b) =>
            (changeMap.get(b.route) ?? 0) -
            (changeMap.get(a.route) ?? 0),
        );
      return { rows, changeMap };
    };

    const consecutiveTwo = buildConsecutiveIncrease(2);
    const consecutiveThree = buildConsecutiveIncrease(3);
    const consecutiveFour = buildConsecutiveIncrease(4);

    const previousRouteRows = previousWeek
      ? weeklyRouteMap.get(previousWeek) ?? []
      : [];
    const volumeUpPphFlatMap = new Map<
      string,
      {
        currentVolume: number;
        volumeIncrease: number;
        volumeChange: number;
        pphChange: number;
      }
    >();
    [...new Set(routeRows.map((row) => row.route))].forEach((route) => {
      const currentMatches = routeRows.filter((row) => row.route === route);
      const previousMatches = previousRouteRows.filter(
        (row) => row.route === route,
      );
      const currentVolume = sum(
        currentMatches.map((row) => row.attempted),
      );
      const previousVolume = sum(
        previousMatches.map((row) => row.attempted),
      );
      const currentPph = pph(
        currentVolume,
        sum(currentMatches.map((row) => row.totalHours)),
      );
      const previousPph = pph(
        previousVolume,
        sum(previousMatches.map((row) => row.totalHours)),
      );
      const pphChange =
        previousPph > 0 ? (currentPph - previousPph) / previousPph : 0;
      if (
        previousVolume > 0 &&
        previousPph > 0 &&
        currentVolume > previousVolume &&
        pphChange <= 0.01
      ) {
        volumeUpPphFlatMap.set(route, {
          currentVolume,
          volumeIncrease: currentVolume - previousVolume,
          volumeChange: (currentVolume - previousVolume) / previousVolume,
          pphChange,
        });
      }
    });
    const volumeUpPphFlat = routeRows
      .filter(
        (row, index, rows) =>
          volumeUpPphFlatMap.has(row.route) &&
          rows.findIndex((candidate) => candidate.route === row.route) ===
            index,
      )
      .sort(
        (a, b) =>
          (volumeUpPphFlatMap.get(b.route)?.volumeIncrease ?? 0) -
          (volumeUpPphFlatMap.get(a.route)?.volumeIncrease ?? 0),
      );

    return {
      consecutiveTwo,
      consecutiveThree,
      consecutiveFour,
      volumeUpPphFlat,
      volumeUpPphFlatMap,
    };
  }, [
    currentWeekIndex,
    previousWeek,
    routeRows,
    weeklyRouteMap,
    weeks,
  ]);

  const weeklyPostalMap = useMemo(() => {
    const map = new Map<string, PostalRow[]>();
    weeks.forEach((week, index) => {
      if (index > currentWeekIndex) return;
      map.set(
        week,
        aggregatePostalRows(
          postalRecords.filter(
            (row) =>
              row.week === week && postalRecordMatches(row, true),
          ),
        ),
      );
    });
    return map;
  }, [
    currentWeekIndex,
    postalRecordMatches,
    postalRecords,
    weeks,
  ]);

  const postalPriorityLists = useMemo(() => {
    const currentByKey = new Map(
      postalRows.map((row) => [postalRowKey(row), row]),
    );
    const weekMaps = new Map(
      [...weeklyPostalMap.entries()].map(([week, rows]) => [
        week,
        new Map(rows.map((row) => [postalRowKey(row), row])),
      ]),
    );

    const buildConsecutiveIncrease = (weekCount: number) => {
      const increaseWeeks = weeks
        .filter((week) => weeks.indexOf(week) <= currentWeekIndex)
        .slice(-weekCount);
      const changeMap = new Map<string, number>();
      if (increaseWeeks.length === weekCount) {
        currentByKey.forEach((_row, key) => {
          const values = increaseWeeks.map(
            (week) => weekMaps.get(week)?.get(key)?.operationPph ?? 0,
          );
          const risesEveryWeek = values.every(
            (value, index) =>
              value > 0 &&
              (index === 0 ||
                (value - values[index - 1]) / values[index - 1] >= 0.001),
          );
          if (risesEveryWeek) {
            changeMap.set(
              key,
              (values.at(-1)! - values[0]) / values[0],
            );
          }
        });
      }
      const rows = postalRows
        .filter((row) => changeMap.has(postalRowKey(row)))
        .sort(
          (a, b) =>
            (changeMap.get(postalRowKey(b)) ?? 0) -
            (changeMap.get(postalRowKey(a)) ?? 0),
        );
      return { rows, changeMap };
    };

    const postalP75 = percentile(
      postalRows
        .map((row) => row.operationPph)
        .filter((value) => value > 0),
      0.75,
    );
    const p75Rows = postalRows
      .filter((row) => row.operationPph >= postalP75)
      .sort((a, b) => b.operationPph - a.operationPph);
    const previousRows = previousWeek
      ? weeklyPostalMap.get(previousWeek) ?? []
      : [];
    const previousByKey = new Map(
      previousRows.map((row) => [postalRowKey(row), row]),
    );
    const volumeUpPphFlatMap = new Map<
      string,
      { volumeIncrease: number; pphChange: number }
    >();
    postalRows.forEach((row) => {
      const key = postalRowKey(row);
      const previous = previousByKey.get(key);
      if (!previous || previous.attempted <= 0 || previous.operationPph <= 0)
        return;
      const pphChange =
        (row.operationPph - previous.operationPph) /
        previous.operationPph;
      if (row.attempted > previous.attempted && pphChange <= 0.01) {
        volumeUpPphFlatMap.set(key, {
          volumeIncrease: row.attempted - previous.attempted,
          pphChange,
        });
      }
    });
    const volumeUpPphFlat = postalRows
      .filter((row) => volumeUpPphFlatMap.has(postalRowKey(row)))
      .sort(
        (a, b) =>
          (volumeUpPphFlatMap.get(postalRowKey(b))?.volumeIncrease ?? 0) -
          (volumeUpPphFlatMap.get(postalRowKey(a))?.volumeIncrease ?? 0),
      );

    return {
      consecutiveTwo: buildConsecutiveIncrease(2),
      consecutiveThree: buildConsecutiveIncrease(3),
      consecutiveFour: buildConsecutiveIncrease(4),
      p75: postalP75,
      p75Rows,
      volumeUpPphFlat,
      volumeUpPphFlatMap,
    };
  }, [
    currentWeekIndex,
    postalRows,
    previousWeek,
    weeklyPostalMap,
    weeks,
  ]);

  const handleUpload = async (
    event: ChangeEvent<HTMLInputElement>,
    type: "performance" | "property" | "postal",
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setUploading(type);
    setNotice("");
    try {
      const rows = await readWorkbookRows(file);
      if (type === "performance") {
        if (isWeeklyProfileRows(rows)) {
          const fileWeek = file.name.match(/W\d+/i)?.[0]?.toUpperCase();
          const latestWeek = weeks.at(-1) ?? "W0";
          const latestWeekNumber = Number.parseInt(latestWeek.slice(1), 10) || 0;
          const targetWeek = fileWeek ?? `W${latestWeekNumber + 1}`;
          const targetWeekNumber =
            Number.parseInt(targetWeek.slice(1), 10) || latestWeekNumber + 1;
          const existingWeekStart = records.find(
            (row) => row.week === targetWeek,
          )?.weekStart;
          const latestWeekStart = records.find(
            (row) => row.week === latestWeek,
          )?.weekStart;
          const inferredStart = latestWeekStart
            ? new Date(latestWeekStart)
            : new Date();
          if (!existingWeekStart) {
            inferredStart.setUTCDate(
              inferredStart.getUTCDate() +
                Math.max(1, targetWeekNumber - latestWeekNumber) * 7,
            );
          }
          const weekStart =
            existingWeekStart ?? inferredStart.toISOString();
          const normalized = normalizeWeeklyProfileRows(
            rows,
            targetWeek,
            weekStart,
          );
          if (!normalized.records.length)
            throw new Error("未识别到有效的路区画像明细。");
          const cleaned = cleanPerformanceRecords(normalized.records);
          const cleanedPostal = cleanPostalPerformanceRecords(
            normalized.postalRecords,
          );
          const combinedRecords = [
            ...records.filter((row) => row.week !== targetWeek),
            ...cleaned.records,
          ];
          const combinedPostalRecords = [
            ...postalRecords.filter((row) => row.week !== targetWeek),
            ...cleanedPostal.records,
          ];
          setRecords(combinedRecords);
          setPostalRecords(combinedPostalRecords);
          saveUploadedData({
            records: combinedRecords,
            postalRecords: combinedPostalRecords,
          });
          setSelectedWeek(targetWeek);
          setExcludedCount((current) => current + cleaned.excluded);
          setExcludedPostalCount(
            (current) => current + cleanedPostal.excluded,
          );
          setPostalSearch("");
          setPage(1);
          setPostalPage(1);
          setSourceMeta({
            sourceRows: combinedRecords.length,
            aggregatedRows: combinedRecords.length,
            propertyRows: properties.length,
            postalRows: combinedPostalRecords.length,
            postalPropertyRows: postalProperties.length,
            postalCostRows: postalCosts.length,
            estimatedTransitRows: estimatedTransitCount,
            estimatedPostalTransitRows: estimatedPostalTransitCount,
            generatedAt: new Date().toISOString(),
          });
          const startDate = new Date(weekStart);
          const endDate = new Date(weekStart);
          endDate.setUTCDate(endDate.getUTCDate() + 6);
          const dateFormatter = new Intl.DateTimeFormat("zh-CN", {
            month: "numeric",
            day: "numeric",
            timeZone: "UTC",
          });
          setNotice(
            `${targetWeek} 已更新（${dateFormatter.format(startDate)}—${dateFormatter.format(endDate)}）：路区 ${formatNumber(cleaned.records.length)} 条、邮编 ${formatNumber(cleanedPostal.records.length)} 条，已自动切换到新周次。`,
          );
        } else {
          const nextRecords = normalizePerformanceRows(rows);
          if (!nextRecords.length)
            throw new Error("未识别到周数、路区、大区与配送字段。");
          const transitFilled = imputeTransitHours(nextRecords);
          const cleaned = cleanPerformanceRecords(transitFilled.records);
          const retainedEstimates = cleaned.records.filter(
            (row) => row.transitHoursEstimated,
          ).length;
          setRecords(cleaned.records);
          saveUploadedData({ records: cleaned.records });
          setExcludedCount(cleaned.excluded);
          setEstimatedTransitCount(retainedEstimates);
          const nextWeeks = sortWeeks(cleaned.records.map((row) => row.week));
          setSelectedWeek(nextWeeks.at(-1) ?? "");
          setSourceMeta({
            sourceRows: rows.length,
            aggregatedRows: cleaned.records.length,
            propertyRows: properties.length,
            postalRows: postalRecords.length,
            postalPropertyRows: postalProperties.length,
            postalCostRows: postalCosts.length,
            estimatedTransitRows: retainedEstimates,
            estimatedPostalTransitRows: estimatedPostalTransitCount,
            generatedAt: new Date().toISOString(),
          });
          setNotice(
            `运营明细已更新：保留 ${formatNumber(cleaned.records.length)} 条，补齐 ${formatNumber(retainedEstimates)} 条均值在途时长，自动剔除 ${formatNumber(cleaned.excluded)} 条异常或空载记录。`,
          );
        }
      } else if (type === "postal") {
        const nextPostalRecords = normalizePostalRows(rows);
        if (!nextPostalRecords.length)
          throw new Error("未识别到周数、邮编、大区与配送字段。");
        const transitFilled = imputeTransitHours(nextPostalRecords);
        const cleanedPostal = cleanPostalPerformanceRecords(
          transitFilled.records,
        );
        const retainedEstimates = cleanedPostal.records.filter(
          (row) => row.transitHoursEstimated,
        ).length;
        setPostalRecords(cleanedPostal.records);
        saveUploadedData({ postalRecords: cleanedPostal.records });
        setExcludedPostalCount(cleanedPostal.excluded);
        setEstimatedPostalTransitCount(retainedEstimates);
        setPostalSearch("");
        setPostalPage(1);
        setSourceMeta((current) =>
          current
            ? {
                ...current,
                postalRows: cleanedPostal.records.length,
                estimatedPostalTransitRows: retainedEstimates,
              }
            : {
                sourceRows: records.length,
                aggregatedRows: records.length,
                propertyRows: properties.length,
                postalRows: cleanedPostal.records.length,
                postalPropertyRows: postalProperties.length,
                postalCostRows: postalCosts.length,
                estimatedTransitRows: estimatedTransitCount,
                estimatedPostalTransitRows: retainedEstimates,
                generatedAt: new Date().toISOString(),
              },
        );
        setNotice(
          `邮编数据已更新：保留 ${formatNumber(cleanedPostal.records.length)} 条，补齐 ${formatNumber(retainedEstimates)} 条均值在途时长，自动剔除 ${formatNumber(cleanedPostal.excluded)} 条异常或空载记录。`,
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
        saveUploadedData({ properties: mergedProperties });
        setSourceMeta((current) =>
          current
            ? { ...current, propertyRows: mergedProperties.length }
            : {
                sourceRows: 0,
                aggregatedRows: records.length,
                propertyRows: mergedProperties.length,
                postalRows: postalRecords.length,
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
    if (id === "monthly") {
      window.location.href = reportVariant === "monthly" ? "/" : "/monthly";
      return;
    }
    const requestedRegion =
      id === "overview"
        ? "ALL"
        : REGION_OPTIONS.find((item) => item.code === id)?.code;
    if (
      lockedRegion &&
      requestedRegion &&
      requestedRegion !== lockedRegion
    ) {
      return;
    }
    setActiveNav(id);
    setSidebarOpen(false);
    if (requestedRegion) {
      setActiveRegion(requestedRegion as RegionCode);
      setSiteFilter("");
      setRouteFilter("");
      setPostalFilter("");
      setDspFilter("");
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }
    document.getElementById(id)?.scrollIntoView({ behavior: "smooth" });
  };
  const toggleSidebarMode = () => {
    if (standaloneMode) return;
    const next = sidebarMode === "function" ? "region" : "function";
    setSidebarMode(next);
    setActiveNav(
      next === "region"
        ? activeRegion === "ALL"
          ? "overview"
          : activeRegion
        : "route-watchlist",
    );
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
  const sortedPostalRows = useMemo(
    () =>
      [...postalRows].sort((a, b) => {
        const aValue = a[postalSort];
        const bValue = b[postalSort];
        const result =
          typeof aValue === "string"
            ? aValue.localeCompare(String(bValue))
            : Number(aValue) - Number(bValue);
        return postalSortDirection === "asc" ? result : -result;
      }),
    [postalRows, postalSort, postalSortDirection],
  );
  const postalPageSize = 15;
  const postalPageCount = Math.max(
    1,
    Math.ceil(sortedPostalRows.length / postalPageSize),
  );
  const visiblePostalRows = sortedPostalRows.slice(
    (postalPage - 1) * postalPageSize,
    postalPage * postalPageSize,
  );

  useEffect(() => {
    setPage(1);
  }, [
    activeRegion,
    businessFilter,
    dspFilter,
    newFilter,
    postalFilter,
    routeFilter,
    selectedWeek,
    siteFilter,
    tableSearch,
  ]);

  useEffect(() => {
    setPostalPage(1);
  }, [
    activeRegion,
    dspFilter,
    postalFilter,
    postalSearch,
    routeFilter,
    selectedWeek,
    siteFilter,
  ]);

  const toggleSort = (key: SortKey) => {
    if (tableSort === key) {
      setSortDirection((current) => (current === "asc" ? "desc" : "asc"));
    } else {
      setTableSort(key);
      setSortDirection(key === "route" ? "asc" : "desc");
    }
  };

  const togglePostalSort = (key: PostalSortKey) => {
    if (postalSort === key) {
      setPostalSortDirection((current) =>
        current === "asc" ? "desc" : "asc",
      );
    } else {
      setPostalSort(key);
      setPostalSortDirection(key === "postalCode" ? "asc" : "desc");
    }
  };

  const exportRows = () => {
    const headers = [
      "周次",
      "大区",
      "路区",
      "站点",
      "DSP",
      "妥投量",
      "分拣耗时",
      "在途耗时",
      "在途时长口径",
      "配送耗时",
      "总工时",
      "妥投PPH",
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
        row.attempted,
        row.sortHours.toFixed(2),
        row.transitHours.toFixed(2),
        row.estimatedTransitRows
          ? row.transitHoursAverageBasis || "均值补齐"
          : "实际值",
        row.deliveryHours.toFixed(2),
        row.totalHours.toFixed(2),
        row.operationPph.toFixed(2),
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

  const exportPostalRows = () => {
    const headers = [
      "周次",
      "大区",
      "邮编",
      "路区",
      "站点",
      "DSP",
      "妥投量",
      "分拣耗时",
      "在途耗时",
      "在途时长口径",
      "配送耗时",
      "总工时",
      "妥投PPH",
    ];
    const lines = sortedPostalRows.map((row) =>
      [
        row.week,
        row.region,
        row.postalCode,
        row.route ?? "",
        row.site,
        row.dsp,
        row.attempted,
        row.sortHours.toFixed(2),
        row.transitHours.toFixed(2),
        row.estimatedTransitRows
          ? row.transitHoursAverageBasis || "均值补齐"
          : "实际值",
        row.deliveryHours.toFixed(2),
        row.totalHours.toFixed(2),
        row.operationPph.toFixed(2),
      ]
        .map(csvEscape)
        .join(","),
    );
    const blob = new Blob(
      ["\uFEFF", headers.map(csvEscape).join(","), "\n", lines.join("\n")],
      { type: "text/csv;charset=utf-8" },
    );
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `邮编维度_${currentRegionName}_${selectedWeek}.csv`;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const downloadWatchlist = (
    title: string,
    rows: RouteRow[],
    value: (row: RouteRow) => string,
    caption: string,
  ) => {
    if (!rows.length) {
      setNotice(`${title}当前没有可下载的路区。`);
      return;
    }
    const headers = [
      "周次",
      "大区",
      "名单类型",
      "路区",
      "站点",
      "DSP",
      "妥投量",
      "妥投PPH",
      "周环比",
      "名单指标",
      "指标说明",
      "本区位置",
      "路区难易度",
      "是否新开",
    ];
    const lines = rows.map((row) => {
      const property = propertyMap.get(row.route);
      return [
        row.week,
        row.region,
        title,
        row.route,
        row.site,
        row.dsp,
        row.attempted,
        row.operationPph.toFixed(2),
        row.wow === null ? "" : row.wow.toFixed(4),
        value(row),
        caption,
        row.percentile,
        property?.difficulty ?? "",
        row.isNew,
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
    const safeTitle = title.replace(/[\\/:*?"<>|]/g, "");
    anchor.href = url;
    anchor.download = `${currentRegionName}_${selectedWeek}_${safeTitle}.csv`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    setNotice(`${title}已下载，共 ${formatNumber(rows.length)} 个路区。`);
  };

  const downloadPostalWatchlist = (
    title: string,
    rows: PostalRow[],
    value: (row: PostalRow) => string,
    caption: string,
  ) => {
    if (!rows.length) {
      setNotice(`${title}当前没有可下载的邮编。`);
      return;
    }
    const headers = [
      "周次",
      "大区",
      "名单类型",
      "邮编",
      "站点",
      "DSP",
      "妥投量",
      "总工时",
      "妥投PPH",
      "名单指标",
      "指标说明",
    ];
    const lines = rows.map((row) =>
      [
        row.week,
        row.region,
        title,
        row.postalCode,
        row.site,
        row.dsp,
        row.attempted,
        row.totalHours.toFixed(2),
        row.operationPph.toFixed(2),
        value(row),
        caption,
      ]
        .map(csvEscape)
        .join(","),
    );
    const blob = new Blob(
      ["\uFEFF", headers.map(csvEscape).join(","), "\n", lines.join("\n")],
      { type: "text/csv;charset=utf-8" },
    );
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    const safeTitle = title.replace(/[\\/:*?"<>|]/g, "");
    anchor.href = url;
    anchor.download = `${currentRegionName}_${selectedWeek}_邮编_${safeTitle}.csv`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    setNotice(`${title}已下载，共 ${formatNumber(rows.length)} 个邮编组合。`);
  };

  const downloadHtmlReport = async () => {
    setExportingHtml(true);
    try {
      let html: string;
      if (window.__PPH_STANDALONE__) {
        html = `<!doctype html>\n${document.documentElement.outerHTML}`;
      } else {
        const response = await fetch("/PPH周报系统-交互版.html", {
          cache: "no-store",
        });
        if (!response.ok) throw new Error("Interactive HTML is unavailable.");
        html = await response.text();
      }
      const regionSource =
        activeRegion === "ALL"
          ? null
          : REGION_OPTIONS.find((region) => region.code === activeRegion)
              ?.source;
      const scopedRecords = regionSource
        ? records.filter((row) => row.region === regionSource)
        : records;
      const scopedRoutes = new Set(scopedRecords.map((row) => row.route));
      const scopedProperties = regionSource
        ? properties.filter((row) => scopedRoutes.has(row.route))
        : properties;
      const scopedPostalRecords =
        activeRegion === "ALL"
          ? postalRecords
          : postalRecords.filter((row) => row.region === activeRegion);
      const scopedPostalCodes = new Set(
        scopedPostalRecords.map((row) => row.postalCode),
      );
      const scopedPostalProperties = postalProperties.filter((row) =>
        scopedPostalCodes.has(row.postalCode),
      );
      const scopedPostalCosts = postalCosts.filter(
        (row) =>
          scopedPostalCodes.has(row.postalCode) &&
          (activeRegion === "ALL" || row.region === activeRegion),
      );
      const scopedData: InitialData = {
        meta: {
          sourceRows: scopedRecords.length,
          aggregatedRows: scopedRecords.length,
          propertyRows: scopedProperties.length,
          postalRows: scopedPostalRecords.length,
          postalPropertyRows: scopedPostalProperties.length,
          postalCostRows: scopedPostalCosts.length,
          estimatedTransitRows: scopedRecords.filter(
            (row) => row.transitHoursEstimated,
          ).length,
          estimatedPostalTransitRows: scopedPostalRecords.filter(
            (row) => row.transitHoursEstimated,
          ).length,
          generatedAt: new Date().toISOString(),
        },
        records: scopedRecords,
        properties: scopedProperties,
        postalRecords: scopedPostalRecords,
        postalProperties: scopedPostalProperties,
        postalCosts: scopedPostalCosts,
      };
      const serializedData = JSON.stringify(scopedData)
        .replace(/&/g, "\\u0026")
        .replace(/</g, "\\u003c")
        .replace(/>/g, "\\u003e")
        .replace(/\u2028/g, "\\u2028")
        .replace(/\u2029/g, "\\u2029");
      const initialDataScript = `<script id="pph-initial-data">window.__PPH_LOCKED_REGION__=${JSON.stringify(activeRegion)};window.__PPH_INITIAL_DATA__=${serializedData};document.title=${JSON.stringify(`${currentRegionName}PPH${reportLabel}`)};</script>`;
      const initialDataMarker = '<script id="pph-initial-data">';
      const initialDataStart = html.lastIndexOf(initialDataMarker);
      const initialDataEnd =
        initialDataStart >= 0
          ? html.indexOf("</script>", initialDataStart)
          : -1;
      html =
        initialDataStart >= 0 && initialDataEnd >= 0
          ? `${html.slice(0, initialDataStart)}${initialDataScript}${html.slice(initialDataEnd + "</script>".length)}`
          : html.replace(/<\/head>/i, `${initialDataScript}\n</head>`);
      const blob = new Blob([html], { type: "text/html;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${currentRegionName}PPH${reportLabel}.html`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      setNotice(
        `${currentRegionName}独立版HTML已生成，仅包含本大区数据。`,
      );
    } catch {
      setNotice(`HTML${reportLabel}生成异常，请刷新页面后重试。`);
    } finally {
      setExportingHtml(false);
    }
  };

  const similarRoutes = useMemo(() => {
    if (!selectedRoute) return [];
    const sourceProperty = propertyMap.get(selectedRoute.route);
    const selectedRouteMetrics = aggregatePerformance(
      routeSearchRows.filter((row) => row.route === selectedRoute.route),
    );
    return routeSearchRows
      .filter((row) => row.route !== selectedRoute.route)
      .map((row) => {
        const property = propertyMap.get(row.route);
        const volumeGap =
          Math.abs(row.attempted - selectedRouteMetrics.attempted) /
          Math.max(1, selectedRouteMetrics.attempted);
        const addressGap = addressDistributionDifference(
          sourceProperty?.addressMix ?? "",
          property?.addressMix ?? "",
        );
        const pphGap =
          Math.abs(row.operationPph - selectedRouteMetrics.operationPph) /
          Math.max(0.01, selectedRouteMetrics.operationPph);
        return {
          row,
          volumeGap,
          addressGap,
          pphGap,
          score: volumeGap + addressGap,
        };
      })
      .sort((a, b) => a.score - b.score)
      .slice(0, 3);
  }, [propertyMap, routeSearchRows, selectedRoute]);
  const selectedRouteHistory = useMemo(() => {
    if (!selectedRoute) return [];
    return weeks
      .filter((week) => weeks.indexOf(week) <= currentWeekIndex)
      .map((week) => {
        const matches = (weeklyRouteMap.get(week) ?? []).filter(
          (row) => row.route === selectedRoute.route,
        );
        if (!matches.length) return undefined;
        const metrics = aggregatePerformance(matches);
        const sites = [...new Set(matches.map((row) => row.site))];
        const dsps = [...new Set(matches.map((row) => row.dsp))];
        const selectedCombination = matches.find(
          (row) => rowKey(row) === rowKey(selectedRoute),
        );
        return {
          ...matches[0],
          ...metrics,
          site: sites.join(" / "),
          dsp: dsps.join(" / "),
          percentile:
            selectedCombination?.percentile ?? matches[0].percentile,
          estimatedTransitRows: sum(
            matches.map((row) => row.estimatedTransitRows ?? 0),
          ),
          transitHoursAverageBasis: matches.some(
            (row) => row.estimatedTransitRows,
          )
            ? "路区内DSP汇总（部分按均值补齐）"
            : undefined,
        } satisfies RouteRow;
      })
      .filter((row): row is RouteRow => Boolean(row))
      .slice(-4);
  }, [currentWeekIndex, selectedRoute, weeklyRouteMap, weeks]);
  const selectedRouteDetail = selectedRouteHistory.at(-1) ?? selectedRoute;
  const selectedRouteProperty = selectedRoute
    ? propertyMap.get(selectedRoute.route)
    : undefined;
  const selectedRouteReasons = useMemo(() => {
    if (!selectedRouteDetail) return [];
    const reasons: string[] = [];
    if (selectedRouteDetail.operationPph < quantiles.p25) {
      reasons.push(
        `妥投PPH低于本区P25（${formatNumber(quantiles.p25, 2)}）`,
      );
    }
    if ((selectedRouteDetail.wow ?? 0) < 0) {
      reasons.push(`较上周下降 ${formatPercent(Math.abs(selectedRouteDetail.wow ?? 0))}`);
    }
    const medianVolume = median(routeRows.map((row) => row.attempted));
    if (
      selectedRouteDetail.attempted >= medianVolume &&
      selectedRouteDetail.operationPph < quantiles.p25
    ) {
      reasons.push("高单量且效率落入本区P25以下");
    }
    return reasons;
  }, [quantiles.p25, routeRows, selectedRouteDetail]);
  const selectedRouteTrendOption = useMemo(
    () => buildDrawerTrendOption(selectedRouteHistory),
    [selectedRouteHistory],
  );
  const selectedRouteChangeSummary = useMemo(() => {
    if (!selectedRoute || selectedRouteHistory.length < 2) return null;
    const contextId = selectedRouteContext?.id ?? "";
    const requestedWeekCount = contextId.includes("consecutive-four")
      ? 4
      : contextId.includes("consecutive-three")
        ? 3
        : 2;
    const history = selectedRouteHistory.slice(-requestedWeekCount);
    if (history.length < 2) return null;
    const from = history[0];
    const to = history.at(-1)!;
    const isVolume = contextId === "volume-up-pph-flat";
    const fromValue = isVolume ? from.attempted : from.operationPph;
    const toValue = isVolume ? to.attempted : to.operationPph;
    return {
      title: selectedRouteContext?.title || "较上周变化",
      metric: isVolume ? "妥投量" : "妥投PPH",
      unit: isVolume ? "单" : "",
      fromWeek: from.week,
      toWeek: to.week,
      fromValue,
      toValue,
      change: fromValue > 0 ? (toValue - fromValue) / fromValue : null,
    };
  }, [selectedRoute, selectedRouteContext, selectedRouteHistory]);

  const postalP25 = percentile(
    postalRows.map((row) => row.operationPph).filter((value) => value > 0),
    0.25,
  );
  const selectedRoutePostalRelations = useMemo(
    () =>
      selectedRoute
        ? postalProperties.filter((row) => row.route === selectedRoute.route)
        : [],
    [postalProperties, selectedRoute],
  );
  const selectedRouteVolumeSignal = selectedRoute
    ? priorityLists.volumeUpPphFlatMap.get(selectedRoute.route)
    : undefined;
  const routePostalImpactRows = useMemo(() => {
    if (!selectedRoute || !selectedRoutePostalRelations.length) return [];
    const relationMatches = (row: PostalRow) =>
      selectedRoutePostalRelations.some(
        (relation) =>
          relation.postalCode === row.postalCode &&
          (!row.route || relation.route === row.route) &&
          (!relation.site || relation.site === row.site) &&
          (!relation.dsp || relation.dsp === row.dsp),
      );
    let sourceRows: PostalRow[];
    switch (selectedRouteContext?.id) {
      case "consecutive-two":
        sourceRows = postalPriorityLists.consecutiveTwo.rows;
        break;
      case "consecutive-three":
        sourceRows = postalPriorityLists.consecutiveThree.rows;
        break;
      case "consecutive-four":
        sourceRows = postalPriorityLists.consecutiveFour.rows;
        break;
      case "p75-high-pph":
        sourceRows = postalPriorityLists.p75Rows;
        break;
      case "p25-low-pph":
        sourceRows = postalRows.filter(
          (row) => row.operationPph < postalP25,
        );
        break;
      case "volume-up-pph-flat":
        // 展示全部关联邮编，并将未跟随路区单量同步增长的邮编标为异常。
        sourceRows = postalRows;
        break;
      default:
        sourceRows = postalRows;
    }
    const matchedKeys = new Set(sourceRows.map((row) => postalRowKey(row)));
    const previousRows = previousWeek
      ? weeklyPostalMap.get(previousWeek) ?? []
      : [];
    const previousMap = new Map(
      previousRows.map((row) => [postalRowKey(row), row]),
    );
    const rows = postalRows.filter(relationMatches).map((row) => {
      const key = postalRowKey(row);
      const previous = previousMap.get(key);
      const volumeIncrease = previous
        ? row.attempted - previous.attempted
        : 0;
      const pphChange =
        previous && previous.operationPph > 0
          ? (row.operationPph - previous.operationPph) /
            previous.operationPph
          : null;
      const cumulativeChange =
        selectedRouteContext?.id === "consecutive-two"
          ? postalPriorityLists.consecutiveTwo.changeMap.get(key) ?? 0
          : selectedRouteContext?.id === "consecutive-three"
            ? postalPriorityLists.consecutiveThree.changeMap.get(key) ?? 0
            : selectedRouteContext?.id === "consecutive-four"
              ? postalPriorityLists.consecutiveFour.changeMap.get(key) ?? 0
              : 0;
      const metric =
        selectedRouteContext?.id === "volume-up-pph-flat"
          ? volumeIncrease
          : selectedRouteContext?.id?.startsWith("consecutive-")
            ? cumulativeChange
            : row.operationPph;
      const isMatch = selectedRouteContext
        ? selectedRouteContext.id === "volume-up-pph-flat"
          ? false
          : matchedKeys.has(key)
        : false;
      return {
        row,
        previous,
        volumeIncrease,
        pphChange,
        cumulativeChange,
        metric,
        isMatch,
      };
    });
    return rows.sort((a, b) => {
      if (a.isMatch !== b.isMatch) return Number(b.isMatch) - Number(a.isMatch);
      if (selectedRouteContext?.id === "p25-low-pph") {
        return a.metric - b.metric;
      }
      return b.metric - a.metric;
    });
  }, [
    postalP25,
    postalPriorityLists,
    postalRows,
    previousWeek,
    selectedRoute,
    selectedRouteContext,
    selectedRoutePostalRelations,
    weeklyPostalMap,
  ]);
  const routePostalMatchedCount = routePostalImpactRows.filter(
    (item) => item.isMatch,
  ).length;
  const isRoutePostalMismatchContext =
    selectedRouteContext?.id === "volume-up-pph-flat";
  const routeVolumeIncrease = selectedRouteVolumeSignal?.volumeIncrease ?? 0;
  const routePostalReconciliation = useMemo(() => {
    if (!selectedRoute || !previousWeek) {
      return { currentVolume: 0, previousVolume: 0, volumeIncrease: 0 };
    }
    const regionCode =
      activeRegion === "ALL"
        ? REGION_OPTIONS.find(
            (region) => region.source === selectedRoute.region,
          )?.code
        : activeRegion;
    const siteQuery = siteFilter.trim().toLowerCase();
    const dspQuery = dspFilter.trim().toLowerCase();
    const relevantRows = postalRecords.filter(
      (row) =>
        row.route === selectedRoute.route &&
        (!regionCode || row.region === regionCode) &&
        (!siteQuery || row.site.toLowerCase().includes(siteQuery)) &&
        (!dspQuery || row.dsp.toLowerCase().includes(dspQuery)) &&
        (row.week === selectedWeek || row.week === previousWeek),
    );
    const currentVolume = sum(
      relevantRows
        .filter((row) => row.week === selectedWeek)
        .map((row) => row.attempted),
    );
    const previousVolume = sum(
      relevantRows
        .filter((row) => row.week === previousWeek)
        .map((row) => row.attempted),
    );
    return {
      currentVolume,
      previousVolume,
      volumeIncrease: currentVolume - previousVolume,
    };
  }, [
    activeRegion,
    dspFilter,
    postalRecords,
    previousWeek,
    selectedRoute,
    selectedWeek,
    siteFilter,
  ]);
  const mappedPostalVolumeIncrease = routePostalReconciliation.volumeIncrease;
  const routePostalVolumeGap =
    routeVolumeIncrease - mappedPostalVolumeIncrease;
  const routePostalMismatchTolerance = Math.max(
    10,
    Math.abs(routeVolumeIncrease) * 0.01,
  );
  const isRoutePostalVolumeMismatch = Boolean(
    isRoutePostalMismatchContext &&
      routeVolumeIncrease >= 1000 &&
      Math.abs(routePostalVolumeGap) > routePostalMismatchTolerance,
  );

  const downloadRoutePostalImpactExcel = () => {
    if (!selectedRoute || !routePostalImpactRows.length) {
      setNotice("当前路区没有可下载的邮编变化明细。");
      return;
    }
    const rows = routePostalImpactRows.map(
      ({
        row,
        previous,
        volumeIncrease,
        pphChange,
        cumulativeChange,
        isMatch,
      }) => {
        const property = selectedRoutePostalRelations.find(
          (item) => item.postalCode === row.postalCode,
        );
        const costs = postalCosts.filter(
          (item) =>
            item.postalCode === row.postalCode &&
            item.region === row.region &&
            item.site === row.site,
        );
        const costVolume = sum(costs.map((item) => item.shipmentVolume));
        const bookedCost = sum(costs.map((item) => item.bookedCost));
        return {
          名单类型: selectedRouteContext?.title || "关联邮编明细",
          识别状态: selectedRouteContext
            ? isRoutePostalMismatchContext
              ? isRoutePostalVolumeMismatch
                ? "路区与邮编汇总不一致"
                : "路区与邮编汇总已对齐"
              : isMatch
                ? "命中"
                : "未命中"
            : "参考",
          周次: row.week,
          大区: row.region,
          站点: row.site,
          路区: selectedRoute.route,
          邮编: row.postalCode,
          DSP: row.dsp,
          本周妥投量: row.attempted,
          上周妥投量: previous?.attempted ?? "",
          单量增加: previous ? volumeIncrease : "",
          路区单量增加: isRoutePostalMismatchContext
            ? routeVolumeIncrease
            : "",
          邮编汇总单量增加: isRoutePostalMismatchContext
            ? mappedPostalVolumeIncrease
            : "",
          路区与邮编汇总差额: isRoutePostalMismatchContext
            ? routePostalVolumeGap
            : "",
          单量环比: previous?.attempted
            ? volumeIncrease / previous.attempted
            : "",
          本周妥投PPH: Number(row.operationPph.toFixed(2)),
          上周妥投PPH: previous
            ? Number(previous.operationPph.toFixed(2))
            : "",
          PPH环比: pphChange ?? "",
          连续周期累计变化: cumulativeChange || "",
          在途时长: Number(row.transitHours.toFixed(2)),
          在途时长口径: row.estimatedTransitRows
            ? row.transitHoursAverageBasis || "均值补齐"
            : "实际值",
          邮编难易度: property?.difficulty || "",
          是否新邮编: property?.isNew || "",
          熟手PPH: property?.expertPph || "",
          DSP件均成本: costVolume > 0 ? bookedCost / costVolume : "",
        };
      },
    );
    const sheet = XLSX.utils.json_to_sheet(rows);
    sheet["!cols"] = Object.keys(rows[0]).map((header) => ({
      wch: Math.min(24, Math.max(10, header.length * 2 + 2)),
    }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "路区全部邮编");
    const safeTitle = (selectedRouteContext?.title || "邮编变化明细").replace(
      /[\\/:*?"<>|]/g,
      "",
    );
    XLSX.writeFile(
      workbook,
      `${selectedRoute.route}_${selectedWeek}_${safeTitle}.xlsx`,
    );
    setNotice(
      isRoutePostalMismatchContext
        ? `${selectedRoute.route}邮编Excel已生成，共 ${formatNumber(rows.length)} 个邮编，路区与邮编汇总${isRoutePostalVolumeMismatch ? "存在差异" : "已对齐"}。`
        : `${selectedRoute.route}邮编Excel已生成，共 ${formatNumber(rows.length)} 个邮编，${formatNumber(routePostalMatchedCount)} 个命中当前变化。`,
    );
  };

  const selectedPostalProperties = useMemo(() => {
    if (!selectedPostal) return [];
    return postalProperties
      .filter((row) => row.postalCode === selectedPostal.postalCode)
      .sort((a, b) => {
        const aMatch =
          Number(Boolean(selectedPostal.route) && a.route === selectedPostal.route) * 4 +
          Number(a.site === selectedPostal.site) * 2 +
          Number(Boolean(a.dsp) && a.dsp === selectedPostal.dsp);
        const bMatch =
          Number(Boolean(selectedPostal.route) && b.route === selectedPostal.route) * 4 +
          Number(b.site === selectedPostal.site) * 2 +
          Number(Boolean(b.dsp) && b.dsp === selectedPostal.dsp);
        return bMatch - aMatch;
      });
  }, [postalProperties, selectedPostal]);
  const selectedPostalCosts = useMemo(() => {
    if (!selectedPostal) return [];
    const matches = postalCosts.filter(
      (row) =>
        row.postalCode === selectedPostal.postalCode &&
        row.region === selectedPostal.region,
    );
    const siteMatches = matches.filter(
      (row) => row.site === selectedPostal.site,
    );
    return siteMatches.length ? siteMatches : matches;
  }, [postalCosts, selectedPostal]);
  const selectedPostalRoutes = useMemo(
    () =>
      [
        ...new Set(
          [
            ...selectedPostalProperties.map((row) => row.route),
            ...selectedPostalCosts.map((row) => row.route),
          ].filter(Boolean),
        ),
      ].sort(),
    [selectedPostalCosts, selectedPostalProperties],
  );
  const selectedPostalHistory = useMemo(() => {
    if (!selectedPostal) return [];
    const key = postalRowKey(selectedPostal);
    return weeks
      .filter((week) => weeks.indexOf(week) <= currentWeekIndex)
      .map((week) => weeklyPostalMap.get(week)?.find((row) => postalRowKey(row) === key))
      .filter((row): row is PostalRow => Boolean(row))
      .slice(-4);
  }, [currentWeekIndex, selectedPostal, weeklyPostalMap, weeks]);
  const selectedPostalTrendOption = useMemo(
    () => buildDrawerTrendOption(selectedPostalHistory),
    [selectedPostalHistory],
  );
  const selectedPostalChangeSummary = useMemo(() => {
    if (!selectedPostal || selectedPostalHistory.length < 2) return null;
    const contextId = selectedPostalContext?.id ?? "";
    const requestedWeekCount = contextId.includes("consecutive-four")
      ? 4
      : contextId.includes("consecutive-three")
        ? 3
        : 2;
    const history = selectedPostalHistory.slice(-requestedWeekCount);
    if (history.length < 2) return null;
    const from = history[0];
    const to = history.at(-1)!;
    const isVolume = contextId === "postal-volume-up-pph-flat";
    const fromValue = isVolume ? from.attempted : from.operationPph;
    const toValue = isVolume ? to.attempted : to.operationPph;
    return {
      title: selectedPostalContext?.title || "较上周变化",
      metric: isVolume ? "妥投量" : "妥投PPH",
      unit: isVolume ? "单" : "",
      fromWeek: from.week,
      toWeek: to.week,
      fromValue,
      toValue,
      change: fromValue > 0 ? (toValue - fromValue) / fromValue : null,
    };
  }, [selectedPostal, selectedPostalContext, selectedPostalHistory]);
  const similarPostals = useMemo(() => {
    if (!selectedPostal) return [];
    return postalRows
      .filter((row) => postalRowKey(row) !== postalRowKey(selectedPostal))
      .map((row) => {
        const volumeGap =
          Math.abs(row.attempted - selectedPostal.attempted) /
          Math.max(1, selectedPostal.attempted);
        const siteGap = row.site === selectedPostal.site ? 0 : 0.2;
        return { row, score: volumeGap + siteGap };
      })
      .sort((a, b) => a.score - b.score)
      .slice(0, 3);
  }, [postalRows, selectedPostal]);

  const openPostalDetails = (
    row: PostalRow,
    context: RouteWatchContext | null = null,
  ) => {
    if (selectedRoute) {
      setPostalParentRoute(selectedRoute);
      setPostalParentRouteContext(selectedRouteContext);
    } else if (!postalParentRoute) {
      setPostalParentRouteContext(null);
    }
    setSelectedRoute(null);
    setSelectedRouteContext(null);
    setSelectedPostalContext(context);
    setSelectedPostal(row);
  };
  const openRouteDetails = (
    row: RouteRow,
    context: RouteWatchContext | null = null,
  ) => {
    setSelectedPostal(null);
    setSelectedPostalContext(null);
    setPostalParentRoute(null);
    setPostalParentRouteContext(null);
    setSelectedRouteContext(context);
    setSelectedRoute(row);
  };
  const closePostalDetails = () => {
    setSelectedPostal(null);
    setSelectedPostalContext(null);
    setPostalParentRoute(null);
    setPostalParentRouteContext(null);
  };
  const returnToParentRoute = () => {
    if (!postalParentRoute) return;
    setSelectedPostal(null);
    setSelectedPostalContext(null);
    setSelectedRoute(postalParentRoute);
    setSelectedRouteContext(postalParentRouteContext);
    setPostalParentRoute(null);
    setPostalParentRouteContext(null);
  };
  const findSearchedRoute = (value: string, allowUniquePartial = false) => {
    const query = value.trim().toLowerCase();
    if (!query) return null;
    const exactRows = routeSearchRows
      .filter((row) => row.route.toLowerCase() === query)
      .sort((a, b) => b.attempted - a.attempted);
    if (exactRows.length) return exactRows[0];
    if (!allowUniquePartial) return null;
    const partialRows = routeSearchRows.filter((row) =>
      row.route.toLowerCase().includes(query),
    );
    const routeNames = [...new Set(partialRows.map((row) => row.route))];
    if (routeNames.length !== 1) return null;
    return partialRows
      .filter((row) => row.route === routeNames[0])
      .sort((a, b) => b.attempted - a.attempted)[0];
  };
  const handleRouteSearchChange = (value: string) => {
    setRouteFilter(value);
  };
  const handleRouteSearchSubmit = () => {
    const matchedRoute = findSearchedRoute(routeFilter, true);
    if (!matchedRoute) {
      setNotice("当前大区及筛选条件下未找到唯一的路区，请输入完整路区名称。");
      return;
    }
    setRouteFilter(matchedRoute.route);
    openRouteDetails(matchedRoute);
  };
  const closeRouteDetails = () => {
    setSelectedRoute(null);
    setSelectedRouteContext(null);
  };
  const activeWeekStart = currentRecords.find(
    (row) => row.week === selectedWeek,
  )?.weekStart;
  const displayDate = activeWeekStart
    ? new Intl.DateTimeFormat("zh-CN", {
        month: "long",
        day: "numeric",
      }).format(new Date(activeWeekStart))
    : "";
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
          {standaloneMode ? (
            <div className="brand-mark" aria-label={`PPH${reportLabel}`}>
              P
            </div>
          ) : (
            <button
              type="button"
              className="brand-mark brand-mark-button"
              onClick={toggleSidebarMode}
              aria-label={
                sidebarMode === "function"
                  ? `打开大区${reportLabel}导航`
                  : "返回功能导航"
              }
              title={
                sidebarMode === "function"
                  ? `打开原大区${reportLabel}导航`
                  : "返回功能导航"
              }
            >
              P
            </button>
          )}
          <div>
            <strong>PPH{reportLabel}</strong>
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
        <div className="nav-label">
          {sidebarMode === "function" ? "功能导航" : `原大区${reportLabel}导航`}
        </div>
        <nav className="main-nav" aria-label="主导航">
          {sidebarMode === "function"
            ? FUNCTION_NAV_ITEMS.map((item) => {
                const Icon = item.icon;
                return (
                  <div key={item.id}>
                    <button
                      className={activeNav === item.id ? "nav-active" : ""}
                      onClick={() => navigate(item.id)}
                    >
                      <Icon size={17} />
                      <span>
                        {item.id === "monthly" && reportVariant === "monthly"
                          ? "PPH周报系统"
                          : item.label.replaceAll("周报", reportLabel)}
                      </span>
                    </button>
                  </div>
                );
              })
            : REGION_NAV_ITEMS.map((item, index) => {
                const Icon = item.icon;
                const regionBoundary = index === 1;
                const actionBoundary = item.id === "exceptions";
                return (
                  <div key={item.id}>
                    {regionBoundary ? (
                      <div className="nav-group-label">大区{reportLabel}</div>
                    ) : null}
                    {actionBoundary ? (
                      <div className="nav-group-label">运营管理</div>
                    ) : null}
                    <button
                      className={activeNav === item.id ? "nav-active" : ""}
                      onClick={() => navigate(item.id)}
                    >
                      <Icon size={17} />
                      <span>{item.label.replaceAll("周报", reportLabel)}</span>
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
            {formatNumber(properties.length)} 条属性 ·{" "}
            {formatNumber(postalRecords.length)} 条邮编记录 · {weeks.length} 个周次
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
                {reportSystemName} <ChevronRight size={13} />{" "}
                <span>{currentRegionName}</span>
              </div>
              <h1>
                {activeRegion === "ALL"
                  ? "全国配送效率总览"
                  : `${currentRegionName} PPH${reportLabel}`}
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
                  周度路区画像表可直接上传，自动更新下一周及邮编数据 ·
                  按“路区名称”自动关联 · 保留全部有效正单量记录 · 已剔除{" "}
                  {formatNumber(excludedCount)} 条异常或空载记录 ·
                  在途为0按同站点/大区同周单均时长均值补齐：路区{" "}
                  {formatNumber(estimatedTransitCount)} 条、邮编{" "}
                  {formatNumber(estimatedPostalTransitCount)} 条 · 邮编维度已加载{" "}
                  {formatNumber(postalRecords.length)} 条
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
              <input
                ref={postalInput}
                type="file"
                accept=".xlsx,.xls,.csv"
                hidden
                onChange={(event) => handleUpload(event, "postal")}
              />
              <button
                className="secondary-button html-export-button"
                onClick={downloadHtmlReport}
                disabled={exportingHtml}
              >
                {exportingHtml ? (
                  <RefreshCw className="spin" size={16} />
                ) : (
                  <Download size={16} />
                )}
                下载HTML{reportLabel}
              </button>
              <button
                className="primary-button"
                onClick={() => performanceInput.current?.click()}
                disabled={uploading !== null}
                title="支持包含大区编码、站点名称、DSP名称、路区名称和邮编的周度路区画像表"
              >
                {uploading === "performance" ? (
                  <RefreshCw className="spin" size={16} />
                ) : (
                  <Upload size={16} />
                )}
                上传周度画像表
              </button>
              <button
                className="secondary-button"
                onClick={() => postalInput.current?.click()}
                disabled={uploading !== null}
              >
                {uploading === "postal" ? (
                  <RefreshCw className="spin" size={16} />
                ) : (
                  <MapPinned size={16} />
                )}
                上传邮编数据
              </button>
              <button
                className="secondary-button"
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

          <section className="filter-bar" aria-label={`${reportLabel}筛选`}>
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
              <div className="text-filter-wrap">
                <Search size={14} />
                <input
                  list="site-filter-options"
                  value={siteFilter}
                  onChange={(event) => {
                    setSiteFilter(event.target.value);
                    setRouteFilter("");
                    setPostalFilter("");
                  }}
                  placeholder="输入站点名称"
                  aria-label="手动输入站点名称"
                />
                <datalist id="site-filter-options">
                  {siteOptions.map((site) => (
                    <option key={site} value={site} />
                  ))}
                </datalist>
              </div>
            </label>
            <label>
              <span>路区</span>
              <div className="text-filter-wrap">
                <Search size={14} />
                <input
                  list="route-filter-options"
                  value={routeFilter}
                  onChange={(event) =>
                    handleRouteSearchChange(event.target.value)
                  }
                  onKeyDown={(event) => {
                    if (event.key !== "Enter") return;
                    event.preventDefault();
                    handleRouteSearchSubmit();
                  }}
                  placeholder="输入路区名称"
                  aria-label="手动输入路区名称"
                  title="输入完整路区名称或按回车查看路区详情"
                />
                <datalist id="route-filter-options">
                  {routeOptions.map((route) => (
                    <option key={route} value={route} />
                  ))}
                </datalist>
              </div>
            </label>
            <label>
              <span>邮编</span>
              <div className="text-filter-wrap">
                <MapPinned size={14} />
                <input
                  list="postal-filter-options"
                  value={postalFilter}
                  onChange={(event) => setPostalFilter(event.target.value)}
                  placeholder="输入邮编"
                  aria-label="手动输入邮编"
                />
                <datalist id="postal-filter-options">
                  {postalCodeOptions.map((postalCode) => (
                    <option key={postalCode} value={postalCode} />
                  ))}
                </datalist>
              </div>
            </label>
            <label>
              <span>DSP</span>
              <div className="text-filter-wrap">
                <Search size={14} />
                <input
                  list="dsp-filter-options"
                  value={dspFilter}
                  onChange={(event) => setDspFilter(event.target.value)}
                  placeholder="输入DSP名称"
                  aria-label="手动输入DSP名称"
                />
                <datalist id="dsp-filter-options">
                  {dspOptions.map((dsp) => (
                    <option key={dsp} value={dsp} />
                  ))}
                </datalist>
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
                {siteFilter || "全部站点"} · {routeFilter || "全部路区"} ·{" "}
                {postalFilter || "全部邮编"} · {dspFilter || "全部DSP"} ·{" "}
                {businessFilter}
              </p>
            </div>
            <div className="scope-insight" aria-live="polite">
              <div className="scope-insight-heading">
                <span>{scopeSummaryInsight.title}</span>
                <small>{scopeSummaryInsight.subtitle}</small>
              </div>
              <div className="scope-insight-items">
                {scopeSummaryInsight.items.map((item) => {
                  const dailyExtra =
                    "dailyExtra" in item &&
                    typeof item.dailyExtra === "number"
                      ? item.dailyExtra
                      : 0;
                  return (
                    <div className="scope-insight-item" key={item.label}>
                      <span>{item.label}</span>
                      <strong className={`scope-value-${item.tone}`}>
                        {item.value}
                      </strong>
                      <small>{item.meta}</small>
                      {dailyExtra > 0 ? (
                        <em
                          className="scope-daily-extra"
                          title="按本周总工时与PPH提升值测算"
                        >
                          日均可多送
                          <strong>+{formatNumber(dailyExtra)} 单</strong>
                        </em>
                      ) : null}
                    </div>
                  );
                })}
              </div>
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

          {scopeFocus ? (
            <section className="scope-trend-panel" aria-label="PPH与妥投量四周趋势">
              <div className="scope-trend-head">
                <div>
                  <span className="scope-trend-kicker">
                    <Activity size={13} /> 4-WEEK PERFORMANCE
                  </span>
                  <h2>{scopeFocusLabel} · PPH与妥投量趋势</h2>
                  <p>折线为妥投PPH，柱状为妥投量；均按当前全部筛选条件汇总</p>
                </div>
                <div className="scope-trend-kpis">
                  <div>
                    <span>本周PPH</span>
                    <strong>
                      {formatNumber(scopeTrend.at(-1)?.operationPph ?? 0, 2)}
                    </strong>
                  </div>
                  <div>
                    <span>本周妥投量</span>
                    <strong>
                      {formatNumber(scopeTrend.at(-1)?.attempted ?? 0)}
                    </strong>
                  </div>
                  <span className={`scope-level scope-level-${scopeFocus}`}>
                    {scopeFocus === "site"
                      ? "站点"
                      : scopeFocus === "route"
                        ? "路区"
                        : "邮编"}
                  </span>
                </div>
              </div>
              {scopeTrend.some((item) => item.attempted > 0) ? (
                <ReactECharts
                  option={scopeTrendOption}
                  notMerge
                  lazyUpdate
                  className="scope-trend-chart"
                  style={{ height: 310, width: "100%" }}
                />
              ) : (
                <div className="scope-trend-empty">
                  <Search size={20} />
                  <strong>没有匹配到趋势数据</strong>
                  <span>可尝试从下拉建议中选择完整的站点、路区或邮编</span>
                </div>
              )}
            </section>
          ) : null}

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
            <div
              id="route-watchlist"
              className="watchlist-dimension-header nav-anchor"
            >
              <div className="watchlist-dimension-icon route-dimension-icon">
                <Route size={16} />
              </div>
              <div>
                <h3>路区重点名单</h3>
                <p>以路区为对象识别连续改善、高效与妥投量效率错配</p>
              </div>
            </div>
            <div className="watchlist-grid">
              {[
                {
                  id: "consecutive-two",
                  title: "连续两周PPH上升",
                  icon: ArrowUpRight,
                  tone: "green",
                  rows: priorityLists.consecutiveTwo.rows,
                  value: (row: RouteRow) =>
                    `+${formatPercent(
                      priorityLists.consecutiveTwo.changeMap.get(row.route) ??
                        0,
                    )}`,
                  caption: "2周累计增幅",
                },
                {
                  id: "consecutive-three",
                  title: "连续三周PPH上升",
                  icon: ArrowUpRight,
                  tone: "green",
                  rows: priorityLists.consecutiveThree.rows,
                  value: (row: RouteRow) =>
                    `+${formatPercent(
                      priorityLists.consecutiveThree.changeMap.get(
                        row.route,
                      ) ?? 0,
                    )}`,
                  caption: "3周累计增幅",
                },
                {
                  id: "consecutive-four",
                  title: "连续四周PPH上升",
                  icon: ArrowUpRight,
                  tone: "green",
                  rows: priorityLists.consecutiveFour.rows,
                  value: (row: RouteRow) =>
                    `+${formatPercent(
                      priorityLists.consecutiveFour.changeMap.get(row.route) ??
                        0,
                    )}`,
                  caption: "4周累计增幅",
                },
                {
                  id: "p75-high-pph",
                  title: "高PPH值的路区（P75）",
                  icon: BarChart3,
                  tone: "green",
                  rows: p75WatchRows,
                  value: (row: RouteRow) =>
                    formatNumber(row.operationPph, 2),
                  caption: "妥投PPH",
                },
                {
                  id: "volume-up-pph-flat",
                  title: "妥投量上升但PPH未上升",
                  icon: ArrowUpRight,
                  tone: "orange",
                  rows: priorityLists.volumeUpPphFlat,
                  value: (row: RouteRow) =>
                    `+${formatPercent(
                      priorityLists.volumeUpPphFlatMap.get(row.route)
                        ?.volumeChange ?? 0,
                    )}`,
                  caption: "较上周妥投量涨幅 · PPH涨幅≤1%",
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
                    <div className="watch-card-head">
                      <button
                        type="button"
                        className="watch-card-summary"
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
                      {list.id === "volume-up-pph-flat" ? (
                        <label className="watch-card-week-select">
                          <span>周次</span>
                          <select
                            aria-label="选择妥投量观察周次"
                            value={selectedWeek}
                            onChange={(event) =>
                              setSelectedWeek(event.target.value)
                            }
                          >
                            {weeks.map((week) => (
                              <option key={week} value={week}>
                                {week}
                              </option>
                            ))}
                          </select>
                          <ChevronDown size={12} aria-hidden="true" />
                        </label>
                      ) : null}
                      <button
                        type="button"
                        className="watch-card-download"
                        onClick={() =>
                          downloadWatchlist(
                            list.title,
                            list.rows,
                            list.value,
                            list.caption,
                          )
                        }
                        disabled={!list.rows.length}
                        aria-label={`下载${list.title}`}
                        title={`下载${list.title}`}
                      >
                        <Download size={15} />
                      </button>
                    </div>
                    <div
                      className="watch-rows"
                      id={`watchlist-${list.id}`}
                    >
                      {list.rows.length ? (
                        visibleRows.map((row) => (
                          <button
                            key={rowKey(row)}
                            onClick={() =>
                              openRouteDetails(row, {
                                id: list.id,
                                title: list.title,
                                caption: list.caption,
                              })
                            }
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

            <div
              id="postal-watchlist"
              className="watchlist-dimension-header postal-dimension-header nav-anchor"
            >
              <div className="watchlist-dimension-icon postal-dimension-icon">
                <MapPinned size={16} />
              </div>
              <div>
                <h3>邮编重点名单</h3>
                <p>
                  以邮编、站点与DSP组合为对象独立计算，当前P75为{" "}
                  {formatNumber(postalPriorityLists.p75, 2)}
                </p>
              </div>
            </div>
            <div className="watchlist-grid">
              {[
                {
                  id: "postal-consecutive-two",
                  title: "连续两周PPH上升",
                  icon: ArrowUpRight,
                  tone: "green",
                  rows: postalPriorityLists.consecutiveTwo.rows,
                  value: (row: PostalRow) =>
                    `+${formatPercent(
                      postalPriorityLists.consecutiveTwo.changeMap.get(
                        postalRowKey(row),
                      ) ?? 0,
                    )}`,
                  caption: "2周累计增幅",
                },
                {
                  id: "postal-consecutive-three",
                  title: "连续三周PPH上升",
                  icon: ArrowUpRight,
                  tone: "green",
                  rows: postalPriorityLists.consecutiveThree.rows,
                  value: (row: PostalRow) =>
                    `+${formatPercent(
                      postalPriorityLists.consecutiveThree.changeMap.get(
                        postalRowKey(row),
                      ) ?? 0,
                    )}`,
                  caption: "3周累计增幅",
                },
                {
                  id: "postal-consecutive-four",
                  title: "连续四周PPH上升",
                  icon: ArrowUpRight,
                  tone: "green",
                  rows: postalPriorityLists.consecutiveFour.rows,
                  value: (row: PostalRow) =>
                    `+${formatPercent(
                      postalPriorityLists.consecutiveFour.changeMap.get(
                        postalRowKey(row),
                      ) ?? 0,
                    )}`,
                  caption: "4周累计增幅",
                },
                {
                  id: "postal-p75-high-pph",
                  title: "高PPH值的邮编（P75）",
                  icon: BarChart3,
                  tone: "green",
                  rows: postalPriorityLists.p75Rows,
                  value: (row: PostalRow) =>
                    formatNumber(row.operationPph, 2),
                  caption: "妥投PPH",
                },
                {
                  id: "postal-volume-up-pph-flat",
                  title: "妥投量上升但PPH未上升",
                  icon: ArrowUpRight,
                  tone: "orange",
                  rows: postalPriorityLists.volumeUpPphFlat,
                  value: (row: PostalRow) =>
                    `+${formatNumber(
                      postalPriorityLists.volumeUpPphFlatMap.get(
                        postalRowKey(row),
                      )?.volumeIncrease ?? 0,
                    )} 单`,
                  caption: "较上周增加 · PPH涨幅≤1%",
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
                    <div className="watch-card-head">
                      <button
                        type="button"
                        className="watch-card-summary"
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
                          <span>{list.rows.length} 个邮编组合</span>
                        </div>
                        {hasMore ? (
                          <ChevronDown
                            className={`watch-card-chevron${isExpanded ? " expanded" : ""}`}
                            size={16}
                          />
                        ) : null}
                      </button>
                      <button
                        type="button"
                        className="watch-card-download"
                        onClick={() =>
                          downloadPostalWatchlist(
                            list.title,
                            list.rows,
                            list.value,
                            list.caption,
                          )
                        }
                        disabled={!list.rows.length}
                        aria-label={`下载邮编${list.title}`}
                        title={`下载邮编${list.title}`}
                      >
                        <Download size={15} />
                      </button>
                    </div>
                    <div className="watch-rows" id={`watchlist-${list.id}`}>
                      {list.rows.length ? (
                        visibleRows.map((row) => (
                          <button
                            key={postalRowKey(row)}
                            onClick={() =>
                              openPostalDetails(row, {
                                id: list.id,
                                title: list.title,
                                caption: list.caption,
                              })
                            }
                          >
                            <div>
                              <strong>{row.postalCode}</strong>
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
                          当前条件下无命中邮编
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
                          : `查看全部 ${list.rows.length} 个邮编`}
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

          <section className="metrics-grid" aria-label="核心指标">
            <MetricCard
              label="妥投PPH"
              value={formatNumber(currentMetrics.operationPph, 2)}
              detail="妥投量 ÷ 总时长"
              icon={Gauge}
              change={wow}
              tone={wow === null ? "blue" : wow >= 0 ? "green" : "red"}
            />
            <MetricCard
              label="日均妥投量"
              value={formatNumber(Math.round(currentMetrics.delivered / 7))}
              detail="本周妥投量 ÷ 7天"
              icon={PackageCheck}
              tone="green"
            />
            <MetricCard
              label="近4周中位数"
              value={formatNumber(recentMedian, 2)}
              detail={`${weeklyTrend.slice(-4).length} 个可用周次的妥投PPH`}
              icon={Activity}
              tone="blue"
            />
            <MetricCard
              label={`${currentRegionName} P50`}
              value={formatNumber(quantiles.p50, 2)}
              detail="本区妥投PPH中位数"
              icon={BarChart3}
              tone="blue"
            />
            <MetricCard
              label="妥投单量"
              value={formatNumber(currentMetrics.delivered)}
              detail="当前筛选口径"
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
            <article className="panel quantile-watch-panel">
              <SectionHeader
                eyebrow="QUANTILE WATCH"
                title="P75 / P25 路区观察"
                description="按当前大区妥投PPH分位数识别高效与关注路区"
              />
              <div className="quantile-watch-grid">
                {[
                  {
                    id: "p75",
                    title: "P75高效路区",
                    threshold: `≥ ${formatNumber(quantiles.p75, 2)}`,
                    rows: p75WatchRows,
                    tone: "high",
                  },
                  {
                    id: "p25",
                    title: "P25关注路区",
                    threshold: `< ${formatNumber(quantiles.p25, 2)}`,
                    rows: p25WatchRows,
                    tone: "low",
                  },
                ].map((group) => (
                  <section className="quantile-watch-column" key={group.id}>
                    <div className="quantile-watch-head">
                      <div>
                        <span
                          className={`quantile-watch-badge quantile-watch-${group.tone}`}
                        >
                          {group.id.toUpperCase()}
                        </span>
                        <div>
                          <strong>{group.title}</strong>
                          <small>
                            {formatNumber(group.rows.length)} 个 ·{" "}
                            {group.threshold}
                          </small>
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() =>
                          downloadWatchlist(
                            group.title,
                            group.rows,
                            (row) => formatNumber(row.operationPph, 2),
                            `妥投PPH ${group.threshold}`,
                          )
                        }
                        disabled={!group.rows.length}
                        aria-label={`下载${group.title}`}
                        title={`下载${group.title}`}
                      >
                        <Download size={14} />
                      </button>
                    </div>
                    <div className="quantile-watch-list">
                      {group.rows.length ? (
                        group.rows.map((row) => (
                          <button
                            type="button"
                            key={rowKey(row)}
                            onClick={() =>
                              openRouteDetails(row, {
                                id:
                                  group.id === "p75"
                                    ? "p75-high-pph"
                                    : "p25-low-pph",
                                title: group.title,
                                caption: `妥投PPH ${group.threshold}`,
                              })
                            }
                          >
                            <span>
                              <strong>{row.route}</strong>
                              <small>
                                {row.site} · {row.dsp}
                              </small>
                            </span>
                            <span>
                              <strong>
                                {formatNumber(row.operationPph, 2)}
                              </strong>
                              <small>{formatNumber(row.attempted)} 单</small>
                            </span>
                          </button>
                        ))
                      ) : (
                        <div className="quantile-watch-empty">
                          当前筛选下暂无路区
                        </div>
                      )}
                    </div>
                  </section>
                ))}
              </div>
            </article>
          </section>

          {reportVariant === "monthly" ? (
            <section className="panel monthly-reference-panel">
              <SectionHeader
                eyebrow="ROUTE DIFFICULTY"
                title="路区难易度数据"
                description={`现有月报模块基础上新增；共 ${formatNumber(properties.length)} 个路区，其中 ${formatNumber(properties.filter((item) => item.difficulty).length)} 个已标注难易度`}
                right={
                  <span className="method-tag">
                    <CheckCircle2 size={14} /> 全量已载入
                  </span>
                }
              />
              <div className="monthly-reference-table-wrap">
                <table className="monthly-reference-table">
                  <thead>
                    <tr>
                      <th>路区名称</th>
                      <th>路区难易度</th>
                      <th>首单里程</th>
                      <th>熟手PPH</th>
                      <th>派送异常率</th>
                      <th>安全度</th>
                      <th>路区时薪</th>
                      <th>Amazon时薪</th>
                    </tr>
                  </thead>
                  <tbody>
                    {monthlyDifficultyRows.map((item) => (
                      <tr key={item.route}>
                        <td><strong>{item.route}</strong><small>{item.transferSite || "未标注站点"}</small></td>
                        <td><span className="monthly-difficulty-value">{item.difficulty}</span></td>
                        <td>{item.firstMile ? `${formatNumber(item.firstMile, 1)} mi` : "未标注"}</td>
                        <td>{item.expertPph ? formatNumber(item.expertPph, 1) : "未标注"}</td>
                        <td>{item.deliveryExceptionRate ? formatPercent(item.deliveryExceptionRate, 2) : "未标注"}</td>
                        <td>{item.safety || "未标注"}</td>
                        <td>{item.routeHourlyWage ? `$${formatNumber(item.routeHourlyWage, 2)}/h` : "暂无数据"}</td>
                        <td>{item.amazonHourlyMedian ? `$${formatNumber(item.amazonHourlyMedian, 2)}/h` : "暂无数据"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          ) : null}

          <section id="postal" className="panel postal-panel nav-anchor">
            <SectionHeader
              eyebrow="POSTAL DIMENSION"
              title={`${currentRegionName}邮编维度`}
              description={`${selectedWeek} · 按邮编、路区、站点与DSP汇总；当前筛选共 ${formatNumber(sortedPostalRows.length)} 个组合`}
              right={
                <button
                  className="secondary-button"
                  onClick={exportPostalRows}
                  disabled={!sortedPostalRows.length}
                >
                  <Download size={16} />
                  导出邮编数据
                </button>
              }
            />

            <div className="postal-toolbar">
              <label className="search-box">
                <Search size={16} />
                <input
                  value={postalSearch}
                  onChange={(event) => setPostalSearch(event.target.value)}
                  placeholder="输入邮编、路区、站点或DSP"
                />
                {postalSearch ? (
                  <button
                    onClick={() => setPostalSearch("")}
                    aria-label="清空邮编搜索"
                  >
                    <X size={14} />
                  </button>
                ) : null}
              </label>
              <span>
                保留全部有效正单量记录 · 已清洗 {formatNumber(excludedPostalCount)} 条 ·{" "}
                {formatNumber(estimatedPostalTransitCount)} 条在途时长按均值补齐
              </span>
            </div>

            <div className="postal-metrics-grid">
              <div className="postal-metric">
                <span>邮编数量</span>
                <strong>{formatNumber(postalCodeCount)}</strong>
                <small>当前大区独立统计</small>
              </div>
              <div className="postal-metric">
                <span>妥投量</span>
                <strong>{formatNumber(postalMetrics.attempted)}</strong>
                <small>当前筛选口径</small>
              </div>
              <div className="postal-metric">
                <span>妥投PPH</span>
                <strong>{formatNumber(postalMetrics.operationPph, 2)}</strong>
                <small>妥投量 ÷ 总工时</small>
              </div>
              <div className="postal-metric">
                <span>总工时</span>
                <strong>{formatNumber(postalMetrics.totalHours)} h</strong>
                <small>分拣、在途与配送</small>
              </div>
            </div>

            <div className="table-wrap postal-table-wrap">
              <table className="postal-table">
                <thead>
                  <tr>
                    <th>
                      <button onClick={() => togglePostalSort("postalCode")}>
                        邮编 <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>路区 / 站点 / DSP</th>
                    <th>
                      <button onClick={() => togglePostalSort("attempted")}>
                        妥投量 <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>在途时长</th>
                    <th>
                      <button onClick={() => togglePostalSort("totalHours")}>
                        总工时 <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>
                      <button
                        onClick={() => togglePostalSort("operationPph")}
                      >
                        妥投PPH <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {visiblePostalRows.map((row) => (
                    <tr key={postalRowKey(row)}>
                      <td>
                        <button
                          className="route-link postal-link"
                          onClick={() => openPostalDetails(row)}
                        >
                          <MapPinned size={15} />
                          <span>
                            <strong>{row.postalCode}</strong>
                            <small>{row.region}</small>
                          </span>
                        </button>
                      </td>
                      <td>
                        <button
                          className="postal-site-link"
                          onClick={() => openPostalDetails(row)}
                        >
                          <strong>{row.route || "未标注路区"}</strong>
                          <small>
                            {row.site || "未标注站点"} · {row.dsp || "未标注DSP"}
                          </small>
                        </button>
                      </td>
                      <td>{formatNumber(row.attempted)}</td>
                      <td>
                        <strong>{formatNumber(row.transitHours, 1)} h</strong>
                        {row.estimatedTransitRows ? (
                          <small className="average-value-tag">均值补齐</small>
                        ) : (
                          <small>实际值</small>
                        )}
                      </td>
                      <td>{formatNumber(row.totalHours)} h</td>
                      <td>
                        <strong>{formatNumber(row.operationPph, 2)}</strong>
                      </td>
                      <td>
                        <button
                          className="row-open"
                          onClick={() => openPostalDetails(row)}
                          aria-label={`查看邮编${row.postalCode}详情`}
                        >
                          <ChevronRight size={16} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!visiblePostalRows.length ? (
                <EmptyState text="当前筛选下没有可显示的邮编数据" />
              ) : null}
            </div>
            <div className="pagination">
              <span>
                第 {postalPage} / {postalPageCount} 页
              </span>
              <div>
                <button
                  onClick={() =>
                    setPostalPage((current) => Math.max(1, current - 1))
                  }
                  disabled={postalPage === 1}
                  aria-label="邮编上一页"
                >
                  <ArrowLeft size={15} />
                </button>
                <button
                  onClick={() =>
                    setPostalPage((current) =>
                      Math.min(postalPageCount, current + 1),
                    )
                  }
                  disabled={postalPage === postalPageCount}
                  aria-label="邮编下一页"
                >
                  <ArrowRight size={15} />
                </button>
              </div>
            </div>
          </section>

          <section id="data" className="panel data-panel nav-anchor">
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
                        妥投量 <ArrowUpDown size={12} />
                      </button>
                    </th>
                    <th>在途时长</th>
                    <th>
                      <button onClick={() => toggleSort("operationPph")}>
                        妥投PPH <ArrowUpDown size={12} />
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
                          onClick={() => openRouteDetails(row)}
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
                        <small>妥投单量</small>
                      </td>
                      <td>
                        <strong>{formatNumber(row.transitHours, 1)} h</strong>
                        {row.estimatedTransitRows ? (
                          <small className="average-value-tag">均值补齐</small>
                        ) : (
                          <small>实际值</small>
                        )}
                      </td>
                      <td>
                        <strong>{formatNumber(row.operationPph, 2)}</strong>
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
                          onClick={() => openRouteDetails(row)}
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
            <span>{reportSystemName} · 指标按当前大区独立计算</span>
            <span>异常依据：自身历史 · 大区分位数 · 相似路区</span>
          </footer>
        </div>
      </main>

      {selectedRoute && selectedRouteDetail ? (
        <div className="drawer-layer">
          <button
            className="drawer-backdrop"
            onClick={closeRouteDetails}
            aria-label="关闭路区详情"
          />
          <section className="route-trend-flyout" aria-label="路区四周PPH与妥投量趋势">
            <div className="drawer-route-trend-head">
              <div>
                <span>
                  <Activity size={12} /> 4-WEEK PERFORMANCE
                </span>
                <strong>{selectedRoute.route} · 四周PPH与妥投量趋势</strong>
                <small>折线为妥投PPH，柱状为妥投量</small>
              </div>
              <div>
                <span>本周</span>
                <strong>{formatNumber(selectedRouteDetail.operationPph, 2)}</strong>
                <small>PPH</small>
              </div>
            </div>
            {selectedRouteHistory.length ? (
              <div
                className="drawer-route-trend-chart"
                role="img"
                aria-label={`${selectedRoute.route}最近${selectedRouteHistory.length}周妥投PPH与妥投量趋势图`}
              >
                <ReactECharts
                  option={selectedRouteTrendOption}
                  notMerge
                  lazyUpdate
                  style={{ height: 300, width: "100%" }}
                />
              </div>
            ) : (
              <p className="missing-copy">当前路区暂无趋势数据。</p>
            )}
          </section>
          <aside className="route-drawer" aria-label="路区详情">
            <div className="drawer-head route-drawer-head">
              <div className="route-head-identity">
                <span>{selectedRoute.region} · {selectedRoute.week}</span>
                <h2>{selectedRoute.route}</h2>
                <p>
                  {selectedRouteDetail.site} · {selectedRouteDetail.dsp}
                </p>
                <div className="drawer-postal-impact-head">
                  <span className="drawer-postal-impact-label">
                    {selectedRouteContext?.title || "关联邮编变化"}
                  </span>
                  {isRoutePostalVolumeMismatch ? (
                    <span className="drawer-postal-warning">
                      <AlertTriangle size={12} />
                      邮编与路区单量增加异常
                    </span>
                  ) : null}
                  <strong>
                    {selectedRouteContext && !isRoutePostalMismatchContext
                      ? `${formatNumber(routePostalMatchedCount)} 命中 / `
                      : ""}
                    {formatNumber(routePostalImpactRows.length)} 个邮编
                  </strong>
                  <button
                    type="button"
                    className="drawer-postal-excel"
                    onClick={downloadRoutePostalImpactExcel}
                    disabled={!routePostalImpactRows.length}
                    aria-label="下载路区邮编变化Excel"
                  >
                    <FileSpreadsheet size={14} />
                    Excel
                  </button>
                </div>
              </div>
              {selectedRouteChangeSummary ? (
                <div className="route-change-summary">
                  <div className="route-change-summary-title">
                    <span>{selectedRouteChangeSummary.title}</span>
                    <strong>{selectedRouteChangeSummary.metric}</strong>
                    {selectedRouteChangeSummary.change !== null ? (
                      <em
                        className={
                          selectedRouteChangeSummary.change >= 0
                            ? "positive"
                            : "negative"
                        }
                      >
                        {selectedRouteChangeSummary.change >= 0 ? "+" : ""}
                        {formatPercent(selectedRouteChangeSummary.change)}
                      </em>
                    ) : null}
                  </div>
                  <div className="route-change-range">
                    <span>
                      <small>{selectedRouteChangeSummary.fromWeek} · 从</small>
                      <strong>
                        {formatNumber(
                          selectedRouteChangeSummary.fromValue,
                          selectedRouteChangeSummary.unit ? 0 : 2,
                        )}
                        {selectedRouteChangeSummary.unit
                          ? ` ${selectedRouteChangeSummary.unit}`
                          : ""}
                      </strong>
                    </span>
                    <ArrowRight size={14} />
                    <span>
                      <small>{selectedRouteChangeSummary.toWeek} · 到</small>
                      <strong>
                        {formatNumber(
                          selectedRouteChangeSummary.toValue,
                          selectedRouteChangeSummary.unit ? 0 : 2,
                        )}
                        {selectedRouteChangeSummary.unit
                          ? ` ${selectedRouteChangeSummary.unit}`
                          : ""}
                      </strong>
                    </span>
                  </div>
                </div>
              ) : null}
              <button
                onClick={closeRouteDetails}
                aria-label="关闭路区详情"
              >
                <X size={19} />
              </button>
            </div>
            <div className="drawer-scroll">
              <section className="drawer-postal-impact-panel">
                <div className="drawer-postal-impact-title">
                  <div>
                    <strong>该路区全部邮编</strong>
                    <span>
                      {isRoutePostalMismatchContext
                        ? `路区 ${routeVolumeIncrease >= 0 ? "+" : ""}${formatNumber(routeVolumeIncrease)} 单 · 邮编合计 ${mappedPostalVolumeIncrease >= 0 ? "+" : ""}${formatNumber(mappedPostalVolumeIncrease)} 单 · 差额 ${formatNumber(routePostalVolumeGap)} 单`
                        : selectedRouteContext
                          ? `框选 ${formatNumber(routePostalMatchedCount)} 个命中邮编，其余邮编正常展示`
                        : "展示当前路区关联邮编及周环比"}
                    </span>
                  </div>
                  <span>{formatNumber(routePostalImpactRows.length)} 个</span>
                </div>
                {routePostalImpactRows.length ? (
                  <div className="drawer-postal-impact-list">
                    {routePostalImpactRows.map(
                      ({
                        row,
                        volumeIncrease,
                        pphChange,
                        cumulativeChange,
                        isMatch,
                      }) => (
                        <button
                          key={postalRowKey(row)}
                          className={
                            isMatch && !isRoutePostalMismatchContext
                              ? "is-highlighted"
                              : undefined
                          }
                          onClick={() => openPostalDetails(row)}
                        >
                          <span>
                            <strong>{row.postalCode}</strong>
                            <small>
                              {row.site} · {row.dsp}
                            </small>
                          </span>
                          <span className="postal-impact-value">
                            {isMatch && !isRoutePostalMismatchContext ? (
                              <em>命中</em>
                            ) : null}
                            <strong>
                              {selectedRouteContext?.id ===
                              "volume-up-pph-flat"
                                ? `${volumeIncrease > 0 ? "+" : ""}${formatNumber(volumeIncrease)} 单`
                                : selectedRouteContext?.id?.startsWith(
                                      "consecutive-",
                                    )
                                  ? isMatch
                                    ? `${cumulativeChange > 0 ? "+" : ""}${formatPercent(cumulativeChange)}`
                                    : formatNumber(row.operationPph, 2)
                                  : formatNumber(row.operationPph, 2)}
                            </strong>
                            <small>
                              {pphChange === null
                                ? "妥投PPH"
                                : `PPH ${pphChange >= 0 ? "+" : ""}${formatPercent(pphChange)}`}
                            </small>
                          </span>
                        </button>
                      ),
                    )}
                  </div>
                ) : (
                  <p className="missing-copy">
                    当前路区没有可关联的邮编数据。
                  </p>
                )}
              </section>

              <div className="drawer-metrics route-drawer-metrics">
                <div>
                  <span>PPH值</span>
                  <strong>{formatNumber(selectedRouteDetail.operationPph, 2)}</strong>
                  <small>{selectedRouteDetail.percentile}</small>
                </div>
                <div>
                  <span>妥投量</span>
                  <strong>{formatNumber(selectedRouteDetail.attempted)} 单</strong>
                  <small>当前路区全部DSP</small>
                </div>
              </div>

              {reportVariant === "monthly" ? (
                <>
                  <section className="drawer-section">
                    <div className="drawer-section-title">
                      <Gauge size={16} />
                      <strong>时薪对比</strong>
                    </div>
                    <div className="property-grid salary-grid">
                      <div className="route-salary-card">
                        <span>路区时薪</span>
                        <strong>{selectedRouteProperty?.routeHourlyWage ? `$${formatNumber(selectedRouteProperty.routeHourlyWage, 2)}/h` : "暂无数据"}</strong>
                      </div>
                      <div className="amazon-salary-card">
                        <span>Amazon时薪（中位数）</span>
                        <strong>{selectedRouteProperty?.amazonHourlyMedian ? `$${formatNumber(selectedRouteProperty.amazonHourlyMedian, 2)}/h` : "暂无数据"}</strong>
                      </div>
                    </div>
                    {selectedRouteProperty?.salaryCity ? <p className="salary-source">调研城市：{selectedRouteProperty.salaryCity}</p> : null}
                  </section>

                  <section className="drawer-section">
                    <div className="drawer-section-title">
                      <AlertTriangle size={16} />
                      <strong>异常原因</strong>
                    </div>
                    {selectedRouteReasons.length ? (
                      <ul className="reason-list">
                        {selectedRouteReasons.map((reason) => <li key={reason}><span />{reason}</li>)}
                      </ul>
                    ) : (
                      <div className="drawer-good"><CheckCircle2 size={17} />当前规则下未识别到重点异常</div>
                    )}
                  </section>

                  <section className="drawer-section">
                    <div className="drawer-section-title">
                      <Route size={16} />
                      <strong>路区难易度</strong>
                    </div>
                    <div className="property-grid monthly-property-grid">
                      <div><span>难易度</span><strong>{selectedRouteProperty?.difficulty || "未标注"}</strong></div>
                      <div><span>首单里程</span><strong>{selectedRouteProperty?.firstMile ? `${formatNumber(selectedRouteProperty.firstMile, 1)} mi` : "未标注"}</strong></div>
                      <div><span>熟手PPH</span><strong>{selectedRouteProperty?.expertPph ? `${formatNumber(selectedRouteProperty.expertPph, 1)} 件` : "未标注"}</strong></div>
                      <div><span>派送异常率</span><strong>{selectedRouteProperty?.deliveryExceptionRate ? formatPercent(selectedRouteProperty.deliveryExceptionRate, 2) : "未标注"}</strong></div>
                      <div><span>DNR率</span><strong>{selectedRouteProperty?.dnrRate ? formatPercent(selectedRouteProperty.dnrRate, 2) : "未标注"}</strong></div>
                      <div><span>安全度</span><strong>{selectedRouteProperty?.safety || "未标注"}</strong></div>
                    </div>
                  </section>

                  <section className="drawer-section">
                    <div className="drawer-section-title">
                      <MapPinned size={16} />
                      <strong>收件地址类型占比</strong>
                    </div>
                    {selectedRouteProperty?.addressMix ? (
                      <div className="address-bars">
                        {addressMixItems(selectedRouteProperty.addressMix).sort((left, right) => right.value - left.value).slice(0, 7).map((item) => (
                          <div key={item.name}>
                            <div><span>{item.name}</span><strong>{formatNumber(item.value, 1)}%</strong></div>
                            <div className="address-track"><span style={{ width: `${Math.min(100, item.value)}%` }} /></div>
                          </div>
                        ))}
                      </div>
                    ) : <p className="missing-copy">当前属性文件未包含该路区的地址类型明细。</p>}
                  </section>
                </>
              ) : null}

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Clock3 size={16} />
                  <strong>耗时结构</strong>
                </div>
                <div className="postal-time-list">
                  {[
                    { label: "分拣", value: selectedRouteDetail.sortHours },
                    { label: "在途", value: selectedRouteDetail.transitHours },
                    { label: "配送", value: selectedRouteDetail.deliveryHours },
                  ].map((item) => {
                    const ratio =
                      selectedRouteDetail.totalHours > 0
                        ? item.value / selectedRouteDetail.totalHours
                        : 0;
                    const isEstimatedTransit =
                      item.label === "在途" &&
                      Boolean(selectedRouteDetail.estimatedTransitRows);
                    return (
                      <div key={item.label}>
                        <div>
                          <span>
                            {item.label}
                            {isEstimatedTransit ? "（均值）" : ""}
                          </span>
                          <strong>
                            {formatNumber(item.value, 1)} h ·{" "}
                            {formatPercent(ratio)}
                          </strong>
                        </div>
                        <div className="address-track">
                          <span
                            style={{ width: `${Math.min(100, ratio * 100)}%` }}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
                {selectedRouteDetail.estimatedTransitRows ? (
                  <p className="average-basis-note">
                    在途时长为0，已使用
                    {selectedRouteDetail.transitHoursAverageBasis || "同类记录均值"}
                    补齐。
                  </p>
                ) : null}
              </section>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Zap size={16} />
                  <strong>相似路区对比</strong>
                </div>
                <p className="similar-basis-note">
                  按单量差异率＋地址分布差异率匹配；商业、公寓、学校、山区等难送地址按1.5倍加权。PPH仅作为结果对比，不参与相似度排名。
                </p>
                <div className="similar-list">
                  {similarRoutes.map(({ row, volumeGap, addressGap, pphGap }) => (
                    <button
                      key={rowKey(row)}
                      onClick={() => openRouteDetails(row)}
                    >
                      <div>
                        <strong>{row.route}</strong>
                        <span>
                          单量差 {formatPercent(volumeGap)} · 地址差 {formatPercent(addressGap)}
                        </span>
                      </div>
                      <div>
                        <strong>{formatNumber(row.operationPph, 2)}</strong>
                        <span>PPH差 {formatPercent(pphGap)}</span>
                      </div>
                    </button>
                  ))}
                </div>
              </section>
            </div>
          </aside>
        </div>
      ) : null}

      {selectedPostal ? (
        <div className="drawer-layer">
          <button
            className="drawer-backdrop"
            onClick={closePostalDetails}
            aria-label="关闭邮编详情"
          />
          <section
            className="route-trend-flyout postal-trend-flyout"
            aria-label="邮编四周PPH与妥投量趋势"
          >
            <div className="drawer-route-trend-head">
              <div>
                <span>
                  <Activity size={12} /> 4-WEEK PERFORMANCE
                </span>
                <strong>
                  {selectedPostal.postalCode} · 四周PPH与妥投量趋势
                </strong>
                <small>折线为妥投PPH，柱状为妥投量</small>
              </div>
              <div>
                <span>本周</span>
                <strong>
                  {formatNumber(
                    selectedPostalHistory.at(-1)?.operationPph ??
                      selectedPostal.operationPph,
                    2,
                  )}
                </strong>
                <small>PPH</small>
              </div>
            </div>
            {selectedPostalHistory.length ? (
              <div
                className="drawer-route-trend-chart"
                role="img"
                aria-label={`邮编${selectedPostal.postalCode}最近${selectedPostalHistory.length}周妥投PPH与妥投量趋势图`}
              >
                <ReactECharts
                  option={selectedPostalTrendOption}
                  notMerge
                  lazyUpdate
                  style={{ height: 300, width: "100%" }}
                />
              </div>
            ) : (
              <p className="missing-copy">当前邮编暂无趋势数据。</p>
            )}
          </section>
          <aside className="route-drawer postal-drawer" aria-label="邮编详情">
            <div className="drawer-head postal-drawer-head">
              <div className="postal-head-identity">
                {postalParentRoute ? (
                  <button
                    type="button"
                    className="postal-back-button"
                    onClick={returnToParentRoute}
                    aria-label={`返回路区${postalParentRoute.route}`}
                  >
                    <ArrowLeft size={12} />
                    返回路区
                  </button>
                ) : null}
                <span>
                  {currentRegionName} · {selectedPostal.week}
                </span>
                <h2>{selectedPostal.postalCode}</h2>
                <p>
                  {selectedPostal.site || "未标注站点"} ·{" "}
                  {selectedPostal.dsp || "未标注DSP"} ·{" "}
                  {selectedPostal.route || "未标注路区"}
                </p>
              </div>
              {selectedPostalChangeSummary ? (
                <div className="postal-change-summary">
                  <div className="postal-change-summary-title">
                    <span>{selectedPostalChangeSummary.title}</span>
                    <strong>{selectedPostalChangeSummary.metric}</strong>
                    {selectedPostalChangeSummary.change !== null ? (
                      <em
                        className={
                          selectedPostalChangeSummary.change >= 0
                            ? "positive"
                            : "negative"
                        }
                      >
                        {selectedPostalChangeSummary.change >= 0 ? "+" : ""}
                        {formatPercent(selectedPostalChangeSummary.change)}
                      </em>
                    ) : null}
                  </div>
                  <div className="postal-change-range">
                    <span>
                      <small>{selectedPostalChangeSummary.fromWeek} · 从</small>
                      <strong>
                        {formatNumber(
                          selectedPostalChangeSummary.fromValue,
                          selectedPostalChangeSummary.unit ? 0 : 2,
                        )}
                        {selectedPostalChangeSummary.unit
                          ? ` ${selectedPostalChangeSummary.unit}`
                          : ""}
                      </strong>
                    </span>
                    <ArrowRight size={14} />
                    <span>
                      <small>{selectedPostalChangeSummary.toWeek} · 到</small>
                      <strong>
                        {formatNumber(
                          selectedPostalChangeSummary.toValue,
                          selectedPostalChangeSummary.unit ? 0 : 2,
                        )}
                        {selectedPostalChangeSummary.unit
                          ? ` ${selectedPostalChangeSummary.unit}`
                          : ""}
                      </strong>
                    </span>
                  </div>
                </div>
              ) : null}
              <button
                onClick={closePostalDetails}
                aria-label="关闭邮编详情"
              >
                <X size={19} />
              </button>
            </div>
            <div className="drawer-scroll">
              <div className="drawer-metrics route-drawer-metrics">
                <div>
                  <span>PPH值</span>
                  <strong>
                    {formatNumber(selectedPostal.operationPph, 2)}
                  </strong>
                  <small>
                    {selectedPostal.operationPph >= postalPriorityLists.p75
                      ? "≥ 邮编P75"
                      : selectedPostal.operationPph < postalP25
                        ? "< 邮编P25"
                        : "P25–P75"}
                  </small>
                </div>
                <div>
                  <span>妥投量</span>
                  <strong>{formatNumber(selectedPostal.delivered)} 单</strong>
                  <small>当前邮编</small>
                </div>
              </div>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Clock3 size={16} />
                  <strong>耗时结构</strong>
                </div>
                <div className="postal-time-list">
                  {[
                    { label: "分拣", value: selectedPostal.sortHours },
                    { label: "在途", value: selectedPostal.transitHours },
                    { label: "配送", value: selectedPostal.deliveryHours },
                  ].map((item) => {
                    const ratio =
                      selectedPostal.totalHours > 0
                        ? item.value / selectedPostal.totalHours
                        : 0;
                    const isEstimatedTransit =
                      item.label === "在途" &&
                      Boolean(selectedPostal.estimatedTransitRows);
                    return (
                      <div key={item.label}>
                        <div>
                          <span>
                            {item.label}
                            {isEstimatedTransit ? "（均值）" : ""}
                          </span>
                          <strong>
                            {formatNumber(item.value)} h ·{" "}
                            {formatPercent(ratio)}
                          </strong>
                        </div>
                        <div className="address-track">
                          <span style={{ width: `${Math.min(100, ratio * 100)}%` }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
                {selectedPostal.estimatedTransitRows ? (
                  <p className="average-basis-note">
                    在途时长为0，已使用
                    {selectedPostal.transitHoursAverageBasis ||
                      "同类记录均值"}
                    补齐。
                  </p>
                ) : null}
              </section>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Route size={16} />
                  <strong>关联路区</strong>
                </div>
                {selectedPostalRoutes.length ? (
                  <div className="postal-route-list">
                    {selectedPostalRoutes.map((route) => {
                      const routeMatch = routeRows.find(
                        (row) => row.route === route,
                      );
                      return (
                        <button
                          key={route}
                          disabled={!routeMatch}
                          onClick={() => {
                            if (!routeMatch) return;
                            openRouteDetails(routeMatch);
                          }}
                        >
                          <span>
                            <strong>{route}</strong>
                            <small>
                              {selectedPostalProperties.find(
                                (row) => row.route === route,
                              )?.site || selectedPostal.site}
                            </small>
                          </span>
                          <span>
                            {routeMatch
                              ? formatNumber(routeMatch.operationPph, 2)
                              : "当前筛选外"}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                ) : (
                  <p className="missing-copy">当前对应表中未找到关联路区。</p>
                )}
              </section>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Activity size={16} />
                  <strong>近4周表现</strong>
                </div>
                {selectedPostalHistory.length ? (
                  <div className="postal-history-list">
                    {selectedPostalHistory.map((row) => (
                      <div key={row.week}>
                        <span>{row.week}</span>
                        <strong>{formatNumber(row.operationPph, 2)}</strong>
                        <small>{formatNumber(row.attempted)} 单</small>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="missing-copy">暂无可用历史周数据。</p>
                )}
              </section>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Zap size={16} />
                  <strong>相似邮编对比</strong>
                </div>
                <div className="similar-list">
                  {similarPostals.map(({ row }) => (
                    <button
                      key={postalRowKey(row)}
                      onClick={() => openPostalDetails(row)}
                    >
                      <div>
                        <strong>{row.postalCode}</strong>
                        <span>
                          {row.site} · 妥投量 {formatNumber(row.attempted)}
                        </span>
                      </div>
                      <div>
                        <strong>{formatNumber(row.operationPph, 2)}</strong>
                        <span>妥投PPH</span>
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
