"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import ReactECharts from "echarts-for-react";
import {
  Activity,
  ArrowDownRight,
  ArrowLeft,
  ArrowRight,
  ArrowUpRight,
  CalendarRange,
  CircleDollarSign,
  Calculator,
  ChevronDown,
  Clock3,
  Download,
  FileSpreadsheet,
  Gauge,
  Layers3,
  MapPinned,
  Route,
  Search,
  SlidersHorizontal,
  TrendingUp,
  WalletCards,
  X,
  Zap,
} from "lucide-react";
import {
  REGION_OPTIONS,
  cleanPerformanceRecords,
  cleanPostalPerformanceRecords,
  formatNumber,
  formatPercent,
  imputeTransitHours,
  percentile,
  pph,
  sortWeeks,
  sum,
} from "../lib/analytics";
import type {
  InitialData,
  PerformanceRecord,
  PostalPerformanceRecord,
  RouteProperty,
} from "../types";

type EntityMode = "route" | "postal";

type PostalWeightCost = {
  postalCode: string;
  route: string;
  site: string;
  region: string;
  weightBand: string;
  priceType: string;
  shipmentVolume: number;
  bookedCost: number;
  averageDspCost: number;
};

type MonthlyRow = {
  key: string;
  route: string;
  postalCode?: string;
  region: string;
  site: string;
  dsp: string;
  weekPph: Array<number | null>;
  weekVolumes: number[];
  monthPph: number;
  monthVolume: number;
  cumulativeChange: number | null;
  consecutiveUp: number;
  property?: RouteProperty;
};

const monthKey = (value: string) => {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "未知月份";
  date.setUTCDate(date.getUTCDate() + 3);
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
};

const monthLabel = (value: string) => {
  const [year, month] = value.split("-");
  return value === "未知月份" ? value : `${year}年${Number(month)}月`;
};

const signedPercent = (value: number | null) =>
  value === null
    ? "—"
    : `${value >= 0 ? "+" : ""}${formatPercent(value)}`;

const salaryGap = (property?: RouteProperty) =>
  property
    ? property.routeHourlyWage - property.amazonHourlyMedian
    : 0;

function getConsecutiveUp(values: Array<number | null>) {
  let count = 0;
  for (let index = values.length - 1; index > 0; index -= 1) {
    const current = values[index];
    const previous = values[index - 1];
    if (current && previous && current > previous) count += 1;
    else break;
  }
  return count;
}

function buildMonthlyDrawerTrendOption(
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

function buildRows(
  mode: EntityMode,
  records: PerformanceRecord[] | PostalPerformanceRecord[],
  weeks: string[],
  propertyMap: Map<string, RouteProperty>,
) {
  const groups = new Map<string, (PerformanceRecord | PostalPerformanceRecord)[]>();
  records.forEach((record) => {
    const postal = "postalCode" in record ? record.postalCode : "";
    const routeName = record.route ?? "未标注路区";
    const key =
      mode === "route"
        ? `${routeName}¦${record.site}¦${record.dsp}`
        : `${postal}¦${routeName}¦${record.site}¦${record.dsp}`;
    const current = groups.get(key) ?? [];
    current.push(record);
    groups.set(key, current);
  });

  return [...groups.entries()].map(([key, rows]) => {
    const first = rows[0];
    const weekPph = weeks.map((week) => {
      const matches = rows.filter((row) => row.week === week);
      if (!matches.length) return null;
      return pph(
        sum(matches.map((row) => row.attempted)),
        sum(matches.map((row) => row.totalHours)),
      );
    });
    const weekVolumes = weeks.map((week) =>
      sum(rows.filter((row) => row.week === week).map((row) => row.attempted)),
    );
    const existingPph = weekPph.filter((value): value is number => value !== null);
    const cumulativeChange =
      existingPph.length >= 2 && existingPph[0] > 0
        ? (existingPph.at(-1)! - existingPph[0]) / existingPph[0]
        : null;
    const routeName = first.route ?? "未标注路区";
    return {
      key,
      route: routeName,
      postalCode: "postalCode" in first ? first.postalCode : undefined,
      region: first.region,
      site: first.site,
      dsp: first.dsp,
      weekPph,
      weekVolumes,
      monthPph: pph(
        sum(rows.map((row) => row.attempted)),
        sum(rows.map((row) => row.totalHours)),
      ),
      monthVolume: sum(rows.map((row) => row.attempted)),
      cumulativeChange,
      consecutiveUp: getConsecutiveUp(weekPph),
      property: propertyMap.get(routeName),
    } satisfies MonthlyRow;
  });
}

export default function MonthlyDashboard() {
  const [records, setRecords] = useState<PerformanceRecord[]>([]);
  const [postalRecords, setPostalRecords] = useState<PostalPerformanceRecord[]>([]);
  const [properties, setProperties] = useState<RouteProperty[]>([]);
  const [postalWeightCosts, setPostalWeightCosts] = useState<PostalWeightCost[]>([]);
  const [loading, setLoading] = useState(true);
  const [selectedMonth, setSelectedMonth] = useState("");
  const [regionFilter, setRegionFilter] = useState("全部大区");
  const [siteFilter, setSiteFilter] = useState("");
  const [dspFilter, setDspFilter] = useState("");
  const [difficultyFilter, setDifficultyFilter] = useState("全部难易度");
  const [businessFilter, setBusinessFilter] = useState("全部模式");
  const [search, setSearch] = useState("");
  const [entityMode, setEntityMode] = useState<EntityMode>("route");
  const [calculatorRoute, setCalculatorRoute] = useState("");
  const [targetHourlyWage, setTargetHourlyWage] = useState("35");
  const [calculatorPph, setCalculatorPph] = useState("");
  const [selectedDetailRow, setSelectedDetailRow] = useState<MonthlyRow | null>(null);

  useEffect(() => {
    setSelectedDetailRow(null);
  }, [selectedMonth, regionFilter, siteFilter, dspFilter, difficultyFilter, businessFilter]);

  useEffect(() => {
    Promise.all([
      fetch("/data/initial.json"),
      fetch("/data/postal-records.json"),
      fetch("/data/postal-weight-costs.json"),
    ])
      .then(async ([initialResponse, postalResponse, weightResponse]) => {
        if (!initialResponse.ok || !postalResponse.ok || !weightResponse.ok)
          throw new Error("data");
        const initial = (await initialResponse.json()) as InitialData;
        const postal = (await postalResponse.json()) as {
          postalRecords: PostalPerformanceRecord[];
        };
        const weightData = (await weightResponse.json()) as {
          postalWeightCosts: PostalWeightCost[];
        };
        const routeFilled = imputeTransitHours(initial.records);
        const postalFilled = imputeTransitHours(postal.postalRecords);
        const cleanedRoutes = cleanPerformanceRecords(routeFilled.records);
        const cleanedPostal = cleanPostalPerformanceRecords(postalFilled.records);
        setRecords(cleanedRoutes.records);
        setPostalRecords(cleanedPostal.records);
        setProperties(initial.properties);
        setPostalWeightCosts(weightData.postalWeightCosts);
        const availableMonths = [
          ...new Set(cleanedRoutes.records.map((row) => monthKey(row.weekStart))),
        ].sort();
        setSelectedMonth(availableMonths.at(-1) ?? "");
      })
      .finally(() => setLoading(false));
  }, []);

  const propertyMap = useMemo(
    () => new Map(properties.map((property) => [property.route, property])),
    [properties],
  );
  const months = useMemo(
    () => [...new Set(records.map((row) => monthKey(row.weekStart)))].sort(),
    [records],
  );
  const monthWeeks = useMemo(
    () =>
      sortWeeks(
        records
          .filter((row) => monthKey(row.weekStart) === selectedMonth)
          .map((row) => row.week),
      ),
    [records, selectedMonth],
  );

  const matchesSharedFilters = (record: PerformanceRecord | PostalPerformanceRecord) => {
    if (monthKey(record.weekStart) !== selectedMonth) return false;
    const regionSource = REGION_OPTIONS.find((item) => item.code === regionFilter)?.source;
    if (regionSource && record.region !== regionSource && record.region !== regionFilter)
      return false;
    if (siteFilter && record.site !== siteFilter) return false;
    if (dspFilter && record.dsp !== dspFilter) return false;
    const property = propertyMap.get(record.route ?? "");
    if (
      difficultyFilter !== "全部难易度" &&
      (property?.difficulty || "未标注") !== difficultyFilter
    )
      return false;
    if (
      businessFilter !== "全部模式" &&
      (property?.businessMode || "未标注") !== businessFilter
    )
      return false;
    return true;
  };

  const monthRecords = useMemo(
    () => records.filter(matchesSharedFilters),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [records, selectedMonth, regionFilter, siteFilter, dspFilter, difficultyFilter, businessFilter, propertyMap],
  );
  const monthPostalRecords = useMemo(
    () => postalRecords.filter(matchesSharedFilters),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [postalRecords, selectedMonth, regionFilter, siteFilter, dspFilter, difficultyFilter, businessFilter, propertyMap],
  );

  const routeRows = useMemo(
    () => buildRows("route", monthRecords, monthWeeks, propertyMap),
    [monthRecords, monthWeeks, propertyMap],
  );
  const postalRows = useMemo(
    () => buildRows("postal", monthPostalRecords, monthWeeks, propertyMap),
    [monthPostalRecords, monthWeeks, propertyMap],
  );

  const selectedRouteRecords = useMemo(() => {
    if (!selectedDetailRow) return [];
    return monthRecords.filter(
      (record) =>
        record.route === selectedDetailRow.route &&
        record.site === selectedDetailRow.site &&
        record.dsp === selectedDetailRow.dsp,
    );
  }, [monthRecords, selectedDetailRow]);

  const selectedRouteHistory = useMemo(() => {
    if (!selectedDetailRow) return [];
    return monthWeeks
      .map((week) => {
        const matches = selectedRouteRecords.filter((record) => record.week === week);
        const attempted = sum(matches.map((record) => record.attempted));
        const totalHours = sum(matches.map((record) => record.totalHours));
        return {
          week,
          attempted,
          operationPph: pph(attempted, totalHours),
        };
      })
      .filter((item) => item.attempted > 0);
  }, [monthWeeks, selectedDetailRow, selectedRouteRecords]);

  const selectedRouteTrendOption = useMemo(
    () => buildMonthlyDrawerTrendOption(selectedRouteHistory),
    [selectedRouteHistory],
  );

  const selectedRouteMetrics = useMemo(() => {
    const attempted = sum(selectedRouteRecords.map((record) => record.attempted));
    const totalHours = sum(selectedRouteRecords.map((record) => record.totalHours));
    return {
      attempted,
      totalHours,
      operationPph: pph(attempted, totalHours),
      sortHours: sum(selectedRouteRecords.map((record) => record.sortHours)),
      transitHours: sum(selectedRouteRecords.map((record) => record.transitHours)),
      deliveryHours: sum(selectedRouteRecords.map((record) => record.deliveryHours)),
    };
  }, [selectedRouteRecords]);

  const selectedRoutePostalRows = useMemo(() => {
    if (!selectedDetailRow) return [];
    return postalRows
      .filter(
        (row) =>
          row.route === selectedDetailRow.route &&
          row.site === selectedDetailRow.site &&
          row.dsp === selectedDetailRow.dsp,
      )
      .sort((left, right) => right.monthVolume - left.monthVolume);
  }, [postalRows, selectedDetailRow]);

  const selectedSimilarRows = useMemo(() => {
    if (!selectedDetailRow) return [];
    return routeRows
      .filter((row) => row.key !== selectedDetailRow.key)
      .map((row) => ({
        row,
        volumeGap:
          selectedDetailRow.monthVolume > 0
            ? Math.abs(row.monthVolume - selectedDetailRow.monthVolume) /
              selectedDetailRow.monthVolume
            : 0,
        pphGap:
          selectedDetailRow.monthPph > 0
            ? (row.monthPph - selectedDetailRow.monthPph) /
              selectedDetailRow.monthPph
            : 0,
        sameDifficulty:
          Boolean(selectedDetailRow.property?.difficulty) &&
          row.property?.difficulty === selectedDetailRow.property?.difficulty,
      }))
      .sort(
        (left, right) =>
          Number(right.sameDifficulty) - Number(left.sameDifficulty) ||
          left.volumeGap - right.volumeGap,
      )
      .slice(0, 5);
  }, [routeRows, selectedDetailRow]);

  const activeRows = entityMode === "route" ? routeRows : postalRows;
  const filteredRows = useMemo(() => {
    const query = search.trim().toLowerCase();
    return activeRows
      .filter((row) =>
        query
          ? [row.route, row.postalCode, row.site, row.dsp, row.property?.salaryCity]
              .join(" ")
              .toLowerCase()
              .includes(query)
          : true,
      )
      .sort(
        (left, right) =>
          right.consecutiveUp - left.consecutiveUp ||
          (right.cumulativeChange ?? -Infinity) -
            (left.cumulativeChange ?? -Infinity),
      );
  }, [activeRows, search]);

  const p75 = percentile(
    routeRows.map((row) => row.monthPph).filter((value) => value > 0),
    0.75,
  );
  const p75Rows = routeRows
    .filter((row) => row.monthPph >= p75)
    .sort((left, right) => right.monthPph - left.monthPph);
  const volumePressureRows = routeRows
    .filter((row) => {
      const volumes = row.weekVolumes;
      const pphValues = row.weekPph;
      if (volumes.length < 2 || pphValues.length < 2) return false;
      const previousVolume = volumes.at(-2) ?? 0;
      const currentVolume = volumes.at(-1) ?? 0;
      const previousPph = pphValues.at(-2);
      const currentPph = pphValues.at(-1);
      if (!previousPph || !currentPph || previousVolume <= 0) return false;
      return (
        currentVolume > previousVolume &&
        (currentPph - previousPph) / previousPph <= 0.01
      );
    })
    .sort((left, right) => {
      const leftIncrease = (left.weekVolumes.at(-1) ?? 0) - (left.weekVolumes.at(-2) ?? 0);
      const rightIncrease = (right.weekVolumes.at(-1) ?? 0) - (right.weekVolumes.at(-2) ?? 0);
      return rightIncrease - leftIncrease;
    });

  const metrics = useMemo(() => {
    const attempted = sum(monthRecords.map((row) => row.attempted));
    const hours = sum(monthRecords.map((row) => row.totalHours));
    return {
      pph: pph(attempted, hours),
      attempted,
      routes: new Set(monthRecords.map((row) => row.route)).size,
      rising: routeRows.filter((row) => row.consecutiveUp >= 2).length,
    };
  }, [monthRecords, routeRows]);

  const siteOptions = useMemo(
    () =>
      [...new Set(records.filter((row) => monthKey(row.weekStart) === selectedMonth).map((row) => row.site))].sort(),
    [records, selectedMonth],
  );
  const dspOptions = useMemo(
    () =>
      [...new Set(records.filter((row) => monthKey(row.weekStart) === selectedMonth && (!siteFilter || row.site === siteFilter)).map((row) => row.dsp))].sort(),
    [records, selectedMonth, siteFilter],
  );
  const difficultyOptions = useMemo(
    () => [...new Set(properties.map((row) => row.difficulty || "未标注"))].sort(),
    [properties],
  );
  const businessOptions = useMemo(
    () => [...new Set(properties.map((row) => row.businessMode || "未标注"))].sort(),
    [properties],
  );

  const salaryRows = useMemo(
    () =>
      routeRows
        .filter(
          (row) =>
            (row.property?.routeHourlyWage ?? 0) > 0 ||
            (row.property?.amazonHourlyMedian ?? 0) > 0,
        )
        .sort((left, right) => salaryGap(left.property) - salaryGap(right.property)),
    [routeRows],
  );
  const salaryAverage = salaryRows.length
    ? sum(salaryRows.map((row) => row.property?.routeHourlyWage ?? 0)) /
      salaryRows.length
    : 0;
  const competitorAverage = salaryRows.length
    ? sum(salaryRows.map((row) => row.property?.amazonHourlyMedian ?? 0)) /
      salaryRows.length
    : 0;
  const calculatorRow =
    routeRows.find((row) => row.route === calculatorRoute) ??
    salaryRows[0] ??
    routeRows[0];
  const calculatorProperty = calculatorRow?.property;
  const calculatorBasisPph =
    Number(calculatorPph) > 0
      ? Number(calculatorPph)
      : calculatorProperty?.expertPph || calculatorRow?.monthPph || 0;
  const targetHourlyValue = Number(targetHourlyWage) || 0;
  const targetUnitPrice =
    calculatorBasisPph > 0 && targetHourlyValue > 0
      ? targetHourlyValue / calculatorBasisPph
      : 0;
  const currentUnitPrice = calculatorProperty?.routeUnitPrice ?? 0;
  const currentCalculatedWage = currentUnitPrice * calculatorBasisPph;
  const unitPriceAdjustment =
    currentUnitPrice > 0 && targetUnitPrice > 0
      ? targetUnitPrice / currentUnitPrice - 1
      : null;
  const calculatorWeightRows = useMemo(() => {
    if (!calculatorRow) return [];
    const groups = new Map<
      string,
      { volume: number; cost: number; priceTypes: Set<string> }
    >();
    postalWeightCosts
      .filter(
        (row) =>
          row.route === calculatorRow.route &&
          (!calculatorRow.site || row.site === calculatorRow.site),
      )
      .forEach((row) => {
        const current = groups.get(row.weightBand) ?? {
          volume: 0,
          cost: 0,
          priceTypes: new Set<string>(),
        };
        current.volume += row.shipmentVolume;
        current.cost += row.bookedCost;
        current.priceTypes.add(row.priceType);
        groups.set(row.weightBand, current);
      });
    return [...groups.entries()]
      .map(([weightBand, item]) => ({
        weightBand,
        volume: item.volume,
        unitCost: item.volume > 0 ? item.cost / item.volume : 0,
        priceTypes: [...item.priceTypes].join(" / "),
      }))
      .sort((left, right) => right.volume - left.volume)
      .slice(0, 8);
  }, [calculatorRow, postalWeightCosts]);

  const exportMonthly = () => {
    const data = routeRows.map((row) => ({
      路区: row.route,
      站点: row.site,
      DSP: row.dsp,
      月度PPH: Number(row.monthPph.toFixed(2)),
      月度妥投量: row.monthVolume,
      累计PPH变化: signedPercent(row.cumulativeChange),
      连续上涨周数: row.consecutiveUp,
      路区难易度: row.property?.difficulty ?? "未标注",
      路区单价: row.property?.routeUnitPrice ?? "",
      路区时薪: row.property?.routeHourlyWage ?? "",
      Amazon时薪中位数: row.property?.amazonHourlyMedian ?? "",
      薪资差: salaryGap(row.property),
      薪资参考城市: row.property?.salaryCity ?? "",
      业务模式: row.property?.businessMode ?? "未标注",
      地址结构: row.property?.addressMix ?? "",
      首单里程: row.property?.firstMile ?? "",
      专家PPH: row.property?.expertPph ?? "",
      妥投异常率: row.property?.deliveryExceptionRate ?? "",
      DNR率: row.property?.dnrRate ?? "",
    }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data), "月报路区明细");
    XLSX.writeFile(workbook, `PPH月报-${selectedMonth}.xlsx`);
  };

  if (loading) {
    return (
      <main className="monthly-loading">
        <CalendarRange size={28} />
        <strong>正在生成PPH月报</strong>
        <span>汇总周趋势、薪资与路区画像…</span>
      </main>
    );
  }

  const goToSection = (id: string) =>
    document.getElementById(id)?.scrollIntoView({ behavior: "smooth" });

  return (
    <div className="app-shell monthly-shell weekly-style-monthly">
      <aside className="sidebar monthly-sidebar">
        <div className="brand">
          <div className="brand-mark">P</div>
          <div><strong>PPH月报</strong><span>运营效能系统</span></div>
        </div>
        <div className="nav-label">功能导航</div>
        <nav className="main-nav" aria-label="月报导航">
          <button onClick={() => goToSection("overview")}><Gauge size={17} /><span>月度概况</span></button>
          <button onClick={() => goToSection("trend")}><TrendingUp size={17} /><span>连续周变化</span></button>
          <button onClick={() => goToSection("p75")}><Route size={17} /><span>P75路区</span></button>
          <button onClick={() => goToSection("pressure")}><Layers3 size={17} /><span>量增效平</span></button>
          <button onClick={() => goToSection("salary-compare")}><CircleDollarSign size={17} /><span>竞对薪资</span></button>
          <button onClick={() => goToSection("salary-calculator")}><Calculator size={17} /><span>薪资计算器</span></button>
          <button onClick={() => goToSection("properties")}><WalletCards size={17} /><span>薪资与画像</span></button>
          <button onClick={() => { window.location.href = "/"; }}><ArrowLeft size={17} /><span>返回周报系统</span></button>
        </nav>
        <div className="sidebar-source">
          <div className="source-status"><span className="status-dot" />月报数据已就绪</div>
          <strong>{formatNumber(records.length)} 条运营记录</strong>
          <span>{formatNumber(properties.length)} 条路区属性 · {months.length} 个月份</span>
        </div>
      </aside>

      <main className="main-content monthly-main">
        <header className="topbar monthly-topbar">
          <div className="topbar-left">
            <div>
              <div className="breadcrumb">PPH月报系统 <span>月度综合经营分析</span></div>
              <h1>PPH月报系统</h1>
            </div>
          </div>
          <div className="topbar-actions">
            <div className="week-badge"><span>数据周期</span><strong>{monthLabel(selectedMonth)}</strong></div>
            <button className="secondary-button" onClick={exportMonthly}><Download size={16} />导出月报Excel</button>
            <div className="avatar">月</div>
          </div>
        </header>

        <div className="dashboard-content">
        <section className="data-toolbar monthly-data-toolbar">
          <div className="toolbar-title">
            <div className="toolbar-icon"><CalendarRange size={20} /></div>
            <div><strong>月度数据工作区</strong><span>沿用周报数据口径，自动汇总连续周PPH、P75路区、量效承压、薪资与路区难易度</span></div>
          </div>
          <div className="upload-actions"><a className="secondary-button" href="/">返回PPH周报</a></div>
        </section>

        <section className="filter-bar monthly-filters" aria-label="月报筛选">
          <div className="filter-lead monthly-filter-title"><SlidersHorizontal size={18} /><strong>筛选</strong></div>
          <label><span>月份</span><div><select value={selectedMonth} onChange={(event) => setSelectedMonth(event.target.value)}>{months.map((month) => <option key={month} value={month}>{monthLabel(month)}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>大区</span><div><select value={regionFilter} onChange={(event) => { setRegionFilter(event.target.value); setSiteFilter(""); setDspFilter(""); }}><option>全部大区</option>{REGION_OPTIONS.map((region) => <option key={region.code} value={region.code}>{region.name}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>站点</span><div><select value={siteFilter} onChange={(event) => { setSiteFilter(event.target.value); setDspFilter(""); }}><option value="">全部站点</option>{siteOptions.map((site) => <option key={site}>{site}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>DSP</span><div><select value={dspFilter} onChange={(event) => setDspFilter(event.target.value)}><option value="">全部DSP</option>{dspOptions.map((dsp) => <option key={dsp}>{dsp}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>难易度</span><div><select value={difficultyFilter} onChange={(event) => setDifficultyFilter(event.target.value)}><option>全部难易度</option>{difficultyOptions.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>业务模式</span><div><select value={businessFilter} onChange={(event) => setBusinessFilter(event.target.value)}><option>全部模式</option>{businessOptions.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown size={14} /></div></label>
        </section>

        <section id="overview" className="scope-summary monthly-hero nav-anchor">
          <div>
            <span>CURRENT MONTH</span>
            <h2>{monthLabel(selectedMonth)}</h2>
            <p>{monthWeeks.join(" · ")} · 月报由周数据自动汇总</p>
          </div>
          <div className="monthly-metric"><small>月度PPH</small><strong>{formatNumber(metrics.pph, 2)}</strong><span>妥投口径</span></div>
          <div className="monthly-metric"><small>月度妥投量</small><strong>{formatNumber(metrics.attempted)}</strong><span>累计单量</span></div>
          <div className="monthly-metric"><small>路区</small><strong>{formatNumber(metrics.routes)}</strong><span>当前筛选</span></div>
          <div className="monthly-metric"><small>连续上涨</small><strong>{formatNumber(metrics.rising)}</strong><span>至少连续2周</span></div>
        </section>

        <section id="trend" className="monthly-panel panel nav-anchor">
          <div className="monthly-section-head">
            <div><span>CONTINUOUS WEEKLY CHANGE</span><h2>连续周PPH变化</h2><p>同一张表查看周度PPH、累计变化、难易度及薪资水平</p></div>
            <div className="monthly-segment"><button className={entityMode === "route" ? "active" : ""} onClick={() => setEntityMode("route")}><Route size={15} />路区</button><button className={entityMode === "postal" ? "active" : ""} onClick={() => setEntityMode("postal")}><MapPinned size={15} />邮编</button></div>
          </div>
          <div className="monthly-table-tools"><label><Search size={15} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="搜索路区、邮编、站点、DSP或城市" /></label><span>共 {formatNumber(filteredRows.length)} 条</span></div>
          <div className="monthly-table-wrap">
            <table className="monthly-table monthly-trend-table"><thead><tr><th>{entityMode === "route" ? "路区" : "邮编 / 路区"}</th><th>站点 / DSP</th>{monthWeeks.map((week) => <th key={week}>{week} PPH</th>)}<th>连续上涨</th><th>累计变化</th><th>月妥投量</th><th>难易度</th><th>路区时薪</th><th>薪资差</th></tr></thead>
              <tbody>
                {filteredRows.slice(0, 120).map((row) => {
                  const detailRow =
                    routeRows.find(
                      (routeRow) =>
                        routeRow.route === row.route &&
                        routeRow.site === row.site &&
                        routeRow.dsp === row.dsp,
                    ) ?? row;
                  return <tr key={row.key}>
                    <td>
                      <button className="monthly-route-link" onClick={() => setSelectedDetailRow(detailRow)}>
                        <strong>{entityMode === "route" ? row.route : row.postalCode}</strong>
                        <small>{entityMode === "postal" ? row.route : row.region} · 点击查看趋势</small>
                      </button>
                    </td>
                    <td><strong>{row.site}</strong><small>{row.dsp}</small></td>
                    {row.weekPph.map((value, index) => <td key={`${row.key}-${monthWeeks[index]}`}><strong>{value === null ? "—" : formatNumber(value, 2)}</strong>{index > 0 && value !== null && row.weekPph[index - 1] ? <small className={value >= row.weekPph[index - 1]! ? "monthly-positive" : "monthly-negative"}>{value >= row.weekPph[index - 1]! ? <ArrowUpRight size={12} /> : <ArrowDownRight size={12} />}{formatPercent(Math.abs((value - row.weekPph[index - 1]!) / row.weekPph[index - 1]!))}</small> : null}</td>)}
                    <td><span className={row.consecutiveUp >= 2 ? "monthly-pill good" : "monthly-pill"}>{row.consecutiveUp ? `${row.consecutiveUp}周` : "—"}</span></td>
                    <td className={(row.cumulativeChange ?? 0) >= 0 ? "monthly-positive" : "monthly-negative"}>{signedPercent(row.cumulativeChange)}</td>
                    <td>{formatNumber(row.monthVolume)}</td>
                    <td><span className="monthly-pill difficulty">{row.property?.difficulty || "未标注"}</span></td>
                    <td>{row.property?.routeHourlyWage ? `$${formatNumber(row.property.routeHourlyWage, 2)}` : "—"}</td>
                    <td className={salaryGap(row.property) >= 0 ? "monthly-positive" : "monthly-negative"}>{row.property ? `${salaryGap(row.property) >= 0 ? "+" : ""}$${formatNumber(salaryGap(row.property), 2)}` : "—"}</td>
                  </tr>;
                })}
              </tbody></table>
          </div>
        </section>

        <section className="monthly-grid-two">
          <div id="p75" className="monthly-panel panel compact nav-anchor"><div className="monthly-section-head"><div><span>TOP QUARTILE</span><h2>P75高PPH路区</h2><p>当前P75标准：{formatNumber(p75, 2)} PPH</p></div><span className="monthly-count">{p75Rows.length}</span></div><div className="monthly-table-wrap"><table className="monthly-table"><thead><tr><th>路区</th><th>月PPH</th><th>高于P75</th><th>妥投量</th><th>难易度</th></tr></thead><tbody>{p75Rows.slice(0, 12).map((row) => <tr key={row.key}><td><button className="monthly-route-link" onClick={() => setSelectedDetailRow(row)}><strong>{row.route}</strong><small>{row.site} · {row.dsp} · 点击查看趋势</small></button></td><td><strong>{formatNumber(row.monthPph, 2)}</strong></td><td className="monthly-positive">+{formatNumber(row.monthPph - p75, 2)}</td><td>{formatNumber(row.monthVolume)}</td><td>{row.property?.difficulty || "未标注"}</td></tr>)}</tbody></table></div></div>
          <div id="pressure" className="monthly-panel panel compact nav-anchor"><div className="monthly-section-head"><div><span>VOLUME UP · PPH FLAT</span><h2>单量上升但PPH未升</h2><p>最近两周妥投量上涨，PPH涨幅≤1%</p></div><span className="monthly-count warning">{volumePressureRows.length}</span></div><div className="monthly-table-wrap"><table className="monthly-table"><thead><tr><th>路区</th><th>单量涨幅</th><th>增加单量</th><th>PPH变化</th><th>难易度</th></tr></thead><tbody>{volumePressureRows.slice(0, 12).map((row) => { const previousVolume = row.weekVolumes.at(-2) ?? 0; const currentVolume = row.weekVolumes.at(-1) ?? 0; const previousPph = row.weekPph.at(-2) ?? 0; const currentPph = row.weekPph.at(-1) ?? 0; const pphChange = previousPph > 0 ? (currentPph - previousPph) / previousPph : 0; return <tr key={row.key}><td><button className="monthly-route-link" onClick={() => setSelectedDetailRow(row)}><strong>{row.route}</strong><small>{row.site} · {row.dsp} · 点击查看趋势</small></button></td><td className="monthly-warning">+{formatPercent((currentVolume - previousVolume) / previousVolume)}</td><td>+{formatNumber(currentVolume - previousVolume)}</td><td className={pphChange >= 0 ? "monthly-positive" : "monthly-negative"}>{signedPercent(pphChange)}</td><td>{row.property?.difficulty || "未标注"}</td></tr>; })}</tbody></table></div></div>
        </section>

        <section id="salary-compare" className="monthly-panel panel nav-anchor salary-compare-panel">
          <div className="monthly-section-head">
            <div>
              <span>SALARY BENCHMARK</span>
              <h2>路区薪资与竞对薪资</h2>
              <p>路区时薪与Amazon同城市时薪中位数对比，优先识别薪资竞争力偏低的路区</p>
            </div>
            <div className="salary-summary-cards">
              <div><span>平均路区时薪</span><strong>${formatNumber(salaryAverage, 2)}/h</strong></div>
              <div><span>竞对平均时薪</span><strong>${formatNumber(competitorAverage, 2)}/h</strong></div>
              <div><span>低于竞对</span><strong>{formatNumber(salaryRows.filter((row) => salaryGap(row.property) < 0).length)}</strong></div>
            </div>
          </div>
          <div className="monthly-table-wrap">
            <table className="monthly-table salary-compare-table">
              <thead><tr><th>路区</th><th>站点 / DSP</th><th>难易度</th><th>月PPH</th><th>路区单价</th><th>路区时薪</th><th>Amazon时薪中位数</th><th>薪资差</th><th>竞争力</th><th>参考城市</th></tr></thead>
              <tbody>
                {salaryRows.slice(0, 120).map((row) => {
                  const gap = salaryGap(row.property);
                  const ratio = (row.property?.amazonHourlyMedian ?? 0) > 0
                    ? (row.property?.routeHourlyWage ?? 0) / (row.property?.amazonHourlyMedian ?? 1)
                    : 0;
                  return <tr key={row.key}>
                    <td><button className="monthly-route-link" onClick={() => setSelectedDetailRow(row)}><strong>{row.route}</strong><small>{row.region} · 点击查看趋势</small></button></td>
                    <td><strong>{row.site}</strong><small>{row.dsp}</small></td>
                    <td><span className="monthly-pill difficulty">{row.property?.difficulty || "未标注"}</span></td>
                    <td>{formatNumber(row.monthPph, 2)}</td>
                    <td>{row.property?.routeUnitPrice ? `$${formatNumber(row.property.routeUnitPrice, 2)}/单` : "—"}</td>
                    <td><strong>${formatNumber(row.property?.routeHourlyWage ?? 0, 2)}/h</strong></td>
                    <td>${formatNumber(row.property?.amazonHourlyMedian ?? 0, 2)}/h</td>
                    <td className={gap >= 0 ? "monthly-positive" : "monthly-negative"}>{gap >= 0 ? "+" : ""}${formatNumber(gap, 2)}</td>
                    <td><span className={`monthly-pill ${ratio >= 1 ? "good" : ratio >= .9 ? "attention" : "risk"}`}>{ratio >= 1 ? "高于竞对" : ratio >= .9 ? "接近竞对" : "低于竞对"}</span></td>
                    <td className="monthly-wide-cell">{row.property?.salaryCity || "—"}</td>
                  </tr>;
                })}
              </tbody>
            </table>
          </div>
        </section>

        <section id="salary-calculator" className="monthly-panel panel nav-anchor wage-calculator-panel">
          <div className="monthly-section-head">
            <div>
              <span>WAGE CALCULATOR</span>
              <h2>目标时薪与重量分段单价计算器</h2>
              <p>选择路区后，以月度PPH或自定义效率测算目标单价，并同步计算各重量段建议价</p>
            </div>
            <Calculator size={27} />
          </div>
          <div className="monthly-calculator-layout">
            <div className="calculator-controls-card">
              <label><span>选择路区</span><div className="select-wrap"><select value={calculatorRow?.route ?? ""} onChange={(event) => { setCalculatorRoute(event.target.value); setCalculatorPph(""); }}><option value="">请选择路区</option>{salaryRows.map((row) => <option key={row.key} value={row.route}>{row.route} · {row.site}</option>)}</select><ChevronDown size={14} /></div></label>
              <div className="calculator-input-grid">
                <label><span>效率计算基准</span><div className="money-input"><input type="number" min="0" step="0.01" value={calculatorPph} onChange={(event) => setCalculatorPph(event.target.value)} placeholder={formatNumber(calculatorProperty?.expertPph || calculatorRow?.monthPph || 0, 2)} /><em>PPH</em></div><small>留空时优先使用薪资文件效率基准</small></label>
                <label><span>目标时薪</span><div className="money-input"><b>$</b><input type="number" min="0" step="0.01" value={targetHourlyWage} onChange={(event) => setTargetHourlyWage(event.target.value)} /><em>/h</em></div><small>可直接输入期望达到的时薪</small></label>
              </div>
              <div className="calculator-quick-targets"><span>快速参照：</span><button onClick={() => setTargetHourlyWage(formatNumber(calculatorProperty?.amazonHourlyMedian ?? 0, 2))}>竞对时薪</button><button onClick={() => setTargetHourlyWage(formatNumber(calculatorProperty?.routeHourlyWage ?? 0, 2))}>当前时薪</button><button onClick={() => setTargetHourlyWage("35")}>$35/h</button></div>
              <div className="calculator-route-meta"><span>{calculatorProperty?.difficulty || "未标注难易度"}</span><span>{calculatorProperty?.businessMode || "未标注模式"}</span><span>{calculatorProperty?.salaryCity || "未标注城市"}</span></div>
            </div>
            <div className="calculator-results-card">
              <div><span>当前文件时薪</span><strong>${formatNumber(calculatorProperty?.routeHourlyWage ?? 0, 2)}/h</strong><small>薪资文件原值</small></div>
              <div><span>竞对时薪中位数</span><strong>${formatNumber(calculatorProperty?.amazonHourlyMedian ?? 0, 2)}/h</strong><small>Amazon同城市参考</small></div>
              <div><span>建议综合单价</span><strong>{targetUnitPrice ? `$${formatNumber(targetUnitPrice, 2)}/单` : "—"}</strong><small>目标时薪 ÷ 效率基准</small></div>
              <div><span>相对当前单价</span><strong className={(unitPriceAdjustment ?? 0) >= 0 ? "monthly-positive" : "monthly-negative"}>{unitPriceAdjustment === null ? "—" : signedPercent(unitPriceAdjustment)}</strong><small>当前单价 ${formatNumber(currentUnitPrice, 2)}/单</small></div>
              <div><span>当前单价测算时薪</span><strong>${formatNumber(currentCalculatedWage, 2)}/h</strong><small>当前单价 × {formatNumber(calculatorBasisPph, 2)} PPH</small></div>
              <div><span>目标与竞对差</span><strong className={targetHourlyValue >= (calculatorProperty?.amazonHourlyMedian ?? 0) ? "monthly-positive" : "monthly-negative"}>{targetHourlyValue - (calculatorProperty?.amazonHourlyMedian ?? 0) >= 0 ? "+" : ""}${formatNumber(targetHourlyValue - (calculatorProperty?.amazonHourlyMedian ?? 0), 2)}/h</strong><small>目标时薪－竞对时薪</small></div>
            </div>
          </div>
          <div className="calculator-formula-line"><CircleDollarSign size={15} /><span>建议综合单价 = 目标时薪 ÷ PPH效率基准；重量段建议价 = 当前重量段单价 ×（建议综合单价 ÷ 当前路区单价）。</span></div>
          <div className="monthly-table-wrap calculator-weight-table">
            <table className="monthly-table"><thead><tr><th>重量段</th><th>价格类型</th><th>样本单量</th><th>当前件均成本</th><th>建议重量段单价</th><th>调整幅度</th></tr></thead><tbody>{calculatorWeightRows.length ? calculatorWeightRows.map((row) => { const ratio = currentUnitPrice > 0 && targetUnitPrice > 0 ? targetUnitPrice / currentUnitPrice : 0; const suggested = ratio > 0 ? row.unitCost * ratio : 0; return <tr key={row.weightBand}><td><strong>{row.weightBand}</strong></td><td>{row.priceTypes}</td><td>{formatNumber(row.volume)}</td><td>${formatNumber(row.unitCost, 2)}</td><td className="monthly-positive">{suggested ? `$${formatNumber(suggested, 2)}` : "—"}</td><td>{ratio ? signedPercent(ratio - 1) : "—"}</td></tr>; }) : <tr><td colSpan={6} className="calculator-empty">当前路区暂无重量分段成本数据</td></tr>}</tbody></table>
          </div>
        </section>

        <section id="properties" className="monthly-panel panel nav-anchor">
          <div className="monthly-section-head"><div><span>SALARY & ROUTE PROFILE</span><h2>薪资、难易度与路区综合画像</h2><p>恢复周报中精简掉的属性字段，用于解释效率差异和制定站点动作</p></div><FileSpreadsheet size={28} /></div>
          <div className="monthly-table-wrap"><table className="monthly-table monthly-profile-table"><thead><tr><th>路区</th><th>难易度</th><th>业务模式</th><th>路区单价</th><th>路区时薪</th><th>Amazon时薪中位数</th><th>薪资差</th><th>参考城市</th><th>首单里程</th><th>专家PPH</th><th>妥投异常率</th><th>DNR率</th><th>面积</th><th>人口密度</th><th>地址结构</th></tr></thead><tbody>{routeRows.slice().sort((a, b) => b.monthVolume - a.monthVolume).slice(0, 120).map((row) => <tr key={row.key}><td><button className="monthly-route-link" onClick={() => setSelectedDetailRow(row)}><strong>{row.route}</strong><small>{row.site} · {row.dsp} · 点击查看趋势</small></button></td><td><span className="monthly-pill difficulty">{row.property?.difficulty || "未标注"}</span></td><td>{row.property?.businessMode || "未标注"}</td><td>{row.property?.routeUnitPrice ? `$${formatNumber(row.property.routeUnitPrice, 2)}` : "—"}</td><td>{row.property?.routeHourlyWage ? `$${formatNumber(row.property.routeHourlyWage, 2)}` : "—"}</td><td>{row.property?.amazonHourlyMedian ? `$${formatNumber(row.property.amazonHourlyMedian, 2)}` : "—"}</td><td className={salaryGap(row.property) >= 0 ? "monthly-positive" : "monthly-negative"}>{row.property ? `${salaryGap(row.property) >= 0 ? "+" : ""}$${formatNumber(salaryGap(row.property), 2)}` : "—"}</td><td className="monthly-wide-cell">{row.property?.salaryCity || "—"}</td><td>{row.property?.firstMile ? `${formatNumber(row.property.firstMile, 1)} mi` : "—"}</td><td>{row.property?.expertPph ? formatNumber(row.property.expertPph, 2) : "—"}</td><td>{row.property?.deliveryExceptionRate ? formatPercent(row.property.deliveryExceptionRate) : "—"}</td><td>{row.property?.dnrRate ? formatPercent(row.property.dnrRate) : "—"}</td><td>{row.property?.landArea ? formatNumber(row.property.landArea, 1) : "—"}</td><td>{row.property?.populationDensity ? formatNumber(row.property.populationDensity, 1) : "—"}</td><td className="monthly-address-cell">{row.property?.addressMix || "—"}</td></tr>)}</tbody></table></div>
        </section>
        <footer className="monthly-footer"><span>PPH月报系统 · 周数据自动汇总</span><span>单量统一采用妥投量口径</span></footer>
        </div>
      </main>

      {selectedDetailRow ? (
        <div className="drawer-layer">
          <button
            className="drawer-backdrop"
            onClick={() => setSelectedDetailRow(null)}
            aria-label="关闭路区详情"
          />
          <section className="route-trend-flyout" aria-label="路区月内PPH与妥投量趋势">
            <div className="drawer-route-trend-head">
              <div>
                <span><Activity size={12} /> MONTHLY WEEK PERFORMANCE</span>
                <strong>{selectedDetailRow.route} · PPH与妥投量趋势</strong>
                <small>折线为妥投PPH，柱状为妥投量；PPH刻度为0.5</small>
              </div>
              <div>
                <span>月度</span>
                <strong>{formatNumber(selectedRouteMetrics.operationPph, 2)}</strong>
                <small>PPH</small>
              </div>
            </div>
            {selectedRouteHistory.length ? (
              <div
                className="drawer-route-trend-chart"
                role="img"
                aria-label={`${selectedDetailRow.route}月内各周妥投PPH与妥投量趋势图`}
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

          <aside className="route-drawer" aria-label="月报路区详情">
            <div className="drawer-head route-drawer-head">
              <div className="route-head-identity">
                <span>{selectedDetailRow.region} · {monthLabel(selectedMonth)}</span>
                <h2>{selectedDetailRow.route}</h2>
                <p>{selectedDetailRow.site} · {selectedDetailRow.dsp}</p>
                <div className="drawer-postal-impact-head">
                  <span className="drawer-postal-impact-label">月度路区画像</span>
                  <strong>{formatNumber(selectedRoutePostalRows.length)} 个邮编</strong>
                </div>
              </div>
              <div className="route-change-summary">
                <div className="route-change-summary-title">
                  <span>月内PPH变化</span>
                  <strong>妥投PPH</strong>
                  {selectedDetailRow.cumulativeChange !== null ? (
                    <em className={selectedDetailRow.cumulativeChange >= 0 ? "positive" : "negative"}>
                      {signedPercent(selectedDetailRow.cumulativeChange)}
                    </em>
                  ) : null}
                </div>
                <div className="route-change-range">
                  <span>
                    <small>{selectedRouteHistory[0]?.week ?? "期初"} · 从</small>
                    <strong>{selectedRouteHistory[0] ? formatNumber(selectedRouteHistory[0].operationPph, 2) : "—"}</strong>
                  </span>
                  <ArrowRight size={14} />
                  <span>
                    <small>{selectedRouteHistory.at(-1)?.week ?? "期末"} · 到</small>
                    <strong>{selectedRouteHistory.length ? formatNumber(selectedRouteHistory.at(-1)!.operationPph, 2) : "—"}</strong>
                  </span>
                </div>
              </div>
              <button onClick={() => setSelectedDetailRow(null)} aria-label="关闭路区详情">
                <X size={19} />
              </button>
            </div>

            <div className="drawer-scroll">
              <section className="drawer-postal-impact-panel">
                <div className="drawer-postal-impact-title">
                  <div>
                    <strong>该路区全部邮编</strong>
                    <span>同步展示月度妥投PPH、妥投量与月内变化</span>
                  </div>
                  <span>{formatNumber(selectedRoutePostalRows.length)} 个</span>
                </div>
                {selectedRoutePostalRows.length ? (
                  <div className="drawer-postal-impact-list monthly-drawer-postal-list">
                    {selectedRoutePostalRows.map((row) => (
                      <div className="monthly-postal-impact-row" key={row.key}>
                        <span>
                          <strong>{row.postalCode}</strong>
                          <small>{row.site} · {row.dsp}</small>
                        </span>
                        <span>
                          <strong>{formatNumber(row.monthPph, 2)}</strong>
                          <small>{formatNumber(row.monthVolume)} 单 · {signedPercent(row.cumulativeChange)}</small>
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="missing-copy">当前路区没有可关联的邮编数据。</p>
                )}
              </section>

              <div className="drawer-metrics route-drawer-metrics">
                <div>
                  <span>PPH值</span>
                  <strong>{formatNumber(selectedRouteMetrics.operationPph, 2)}</strong>
                  <small>月度妥投口径</small>
                </div>
                <div>
                  <span>妥投量</span>
                  <strong>{formatNumber(selectedRouteMetrics.attempted)} 单</strong>
                  <small>月度累计</small>
                </div>
              </div>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Clock3 size={16} />
                  <strong>耗时结构</strong>
                </div>
                <div className="postal-time-list">
                  {[
                    { label: "分拣", value: selectedRouteMetrics.sortHours },
                    { label: "在途", value: selectedRouteMetrics.transitHours },
                    { label: "配送", value: selectedRouteMetrics.deliveryHours },
                  ].map((item) => {
                    const ratio = selectedRouteMetrics.totalHours > 0
                      ? item.value / selectedRouteMetrics.totalHours
                      : 0;
                    return <div key={item.label}>
                      <div><span>{item.label}</span><strong>{formatNumber(item.value, 1)} h · {formatPercent(ratio)}</strong></div>
                      <div className="address-track"><span style={{ width: `${Math.min(100, ratio * 100)}%` }} /></div>
                    </div>;
                  })}
                </div>
              </section>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <CircleDollarSign size={16} />
                  <strong>薪资、难易度与路区画像</strong>
                </div>
                <div className="monthly-drawer-profile">
                  <div><span>难易度</span><strong>{selectedDetailRow.property?.difficulty || "未标注"}</strong></div>
                  <div><span>业务模式</span><strong>{selectedDetailRow.property?.businessMode || "未标注"}</strong></div>
                  <div><span>路区单价</span><strong>{selectedDetailRow.property?.routeUnitPrice ? `$${formatNumber(selectedDetailRow.property.routeUnitPrice, 2)}/单` : "—"}</strong></div>
                  <div><span>路区时薪</span><strong>{selectedDetailRow.property?.routeHourlyWage ? `$${formatNumber(selectedDetailRow.property.routeHourlyWage, 2)}/h` : "—"}</strong></div>
                  <div><span>竞对时薪</span><strong>{selectedDetailRow.property?.amazonHourlyMedian ? `$${formatNumber(selectedDetailRow.property.amazonHourlyMedian, 2)}/h` : "—"}</strong></div>
                  <div><span>薪资差</span><strong className={salaryGap(selectedDetailRow.property) >= 0 ? "monthly-positive" : "monthly-negative"}>{selectedDetailRow.property ? `${salaryGap(selectedDetailRow.property) >= 0 ? "+" : ""}$${formatNumber(salaryGap(selectedDetailRow.property), 2)}/h` : "—"}</strong></div>
                  <div><span>首单里程</span><strong>{selectedDetailRow.property?.firstMile ? `${formatNumber(selectedDetailRow.property.firstMile, 1)} mi` : "—"}</strong></div>
                  <div><span>专家PPH</span><strong>{selectedDetailRow.property?.expertPph ? formatNumber(selectedDetailRow.property.expertPph, 2) : "—"}</strong></div>
                </div>
                <p className="monthly-address-profile"><strong>地址结构：</strong>{selectedDetailRow.property?.addressMix || "暂无地址结构信息"}</p>
              </section>

              <section className="drawer-section">
                <div className="drawer-section-title">
                  <Zap size={16} />
                  <strong>相似路区对比</strong>
                </div>
                <p className="similar-basis-note">优先匹配相同难易度，再按月度妥投量接近程度排序；PPH仅作为结果对比。</p>
                <div className="similar-list">
                  {selectedSimilarRows.map(({ row, volumeGap, pphGap }) => (
                    <button key={row.key} onClick={() => setSelectedDetailRow(row)}>
                      <div><strong>{row.route}</strong><span>{row.site} · 单量差 {formatPercent(volumeGap)}</span></div>
                      <div><strong>{formatNumber(row.monthPph, 2)}</strong><span>PPH差 {pphGap >= 0 ? "+" : ""}{formatPercent(pphGap)}</span></div>
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
