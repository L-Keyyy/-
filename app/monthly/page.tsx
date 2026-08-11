"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  ArrowDownRight,
  ArrowLeft,
  ArrowUpRight,
  CalendarRange,
  ChevronDown,
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
  const [loading, setLoading] = useState(true);
  const [selectedMonth, setSelectedMonth] = useState("");
  const [regionFilter, setRegionFilter] = useState("全部大区");
  const [siteFilter, setSiteFilter] = useState("");
  const [dspFilter, setDspFilter] = useState("");
  const [difficultyFilter, setDifficultyFilter] = useState("全部难易度");
  const [businessFilter, setBusinessFilter] = useState("全部模式");
  const [search, setSearch] = useState("");
  const [entityMode, setEntityMode] = useState<EntityMode>("route");

  useEffect(() => {
    Promise.all([fetch("/data/initial.json"), fetch("/data/postal-records.json")])
      .then(async ([initialResponse, postalResponse]) => {
        if (!initialResponse.ok || !postalResponse.ok) throw new Error("data");
        const initial = (await initialResponse.json()) as InitialData;
        const postal = (await postalResponse.json()) as {
          postalRecords: PostalPerformanceRecord[];
        };
        const routeFilled = imputeTransitHours(initial.records);
        const postalFilled = imputeTransitHours(postal.postalRecords);
        const cleanedRoutes = cleanPerformanceRecords(routeFilled.records);
        const cleanedPostal = cleanPostalPerformanceRecords(postalFilled.records);
        setRecords(cleanedRoutes.records);
        setPostalRecords(cleanedPostal.records);
        setProperties(initial.properties);
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

  return (
    <div className="monthly-shell">
      <aside className="monthly-sidebar">
        <a className="monthly-brand" href="/monthly">
          <span>M</span>
          <div><strong>PPH月报</strong><small>综合经营分析</small></div>
        </a>
        <nav>
          <a href="#overview"><Gauge size={17} />月度概况</a>
          <a href="#trend"><TrendingUp size={17} />连续周变化</a>
          <a href="#p75"><Route size={17} />P75路区</a>
          <a href="#pressure"><Layers3 size={17} />量增效平</a>
          <a href="#properties"><WalletCards size={17} />薪资与画像</a>
        </nav>
        <a className="monthly-back" href="/"><ArrowLeft size={16} />返回周报系统</a>
      </aside>

      <main className="monthly-main">
        <header className="monthly-topbar">
          <div>
            <span>MONTHLY PERFORMANCE REVIEW</span>
            <h1>PPH月报系统</h1>
            <p>连续周效率、P75标杆、量效承压与薪资难易度统一分析</p>
          </div>
          <button onClick={exportMonthly}><Download size={16} />导出月报Excel</button>
        </header>

        <section className="monthly-filters" aria-label="月报筛选">
          <div className="monthly-filter-title"><SlidersHorizontal size={18} /><strong>筛选</strong></div>
          <label><span>月份</span><div><select value={selectedMonth} onChange={(event) => setSelectedMonth(event.target.value)}>{months.map((month) => <option key={month} value={month}>{monthLabel(month)}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>大区</span><div><select value={regionFilter} onChange={(event) => { setRegionFilter(event.target.value); setSiteFilter(""); setDspFilter(""); }}><option>全部大区</option>{REGION_OPTIONS.map((region) => <option key={region.code} value={region.code}>{region.name}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>站点</span><div><select value={siteFilter} onChange={(event) => { setSiteFilter(event.target.value); setDspFilter(""); }}><option value="">全部站点</option>{siteOptions.map((site) => <option key={site}>{site}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>DSP</span><div><select value={dspFilter} onChange={(event) => setDspFilter(event.target.value)}><option value="">全部DSP</option>{dspOptions.map((dsp) => <option key={dsp}>{dsp}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>难易度</span><div><select value={difficultyFilter} onChange={(event) => setDifficultyFilter(event.target.value)}><option>全部难易度</option>{difficultyOptions.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown size={14} /></div></label>
          <label><span>业务模式</span><div><select value={businessFilter} onChange={(event) => setBusinessFilter(event.target.value)}><option>全部模式</option>{businessOptions.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown size={14} /></div></label>
        </section>

        <section id="overview" className="monthly-hero">
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

        <section id="trend" className="monthly-panel">
          <div className="monthly-section-head">
            <div><span>CONTINUOUS WEEKLY CHANGE</span><h2>连续周PPH变化</h2><p>同一张表查看周度PPH、累计变化、难易度及薪资水平</p></div>
            <div className="monthly-segment"><button className={entityMode === "route" ? "active" : ""} onClick={() => setEntityMode("route")}><Route size={15} />路区</button><button className={entityMode === "postal" ? "active" : ""} onClick={() => setEntityMode("postal")}><MapPinned size={15} />邮编</button></div>
          </div>
          <div className="monthly-table-tools"><label><Search size={15} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="搜索路区、邮编、站点、DSP或城市" /></label><span>共 {formatNumber(filteredRows.length)} 条</span></div>
          <div className="monthly-table-wrap">
            <table className="monthly-table monthly-trend-table"><thead><tr><th>{entityMode === "route" ? "路区" : "邮编 / 路区"}</th><th>站点 / DSP</th>{monthWeeks.map((week) => <th key={week}>{week} PPH</th>)}<th>连续上涨</th><th>累计变化</th><th>月妥投量</th><th>难易度</th><th>路区时薪</th><th>薪资差</th></tr></thead>
              <tbody>{filteredRows.slice(0, 120).map((row) => <tr key={row.key}><td><strong>{entityMode === "route" ? row.route : row.postalCode}</strong><small>{entityMode === "postal" ? row.route : row.region}</small></td><td><strong>{row.site}</strong><small>{row.dsp}</small></td>{row.weekPph.map((value, index) => <td key={`${row.key}-${monthWeeks[index]}`}><strong>{value === null ? "—" : formatNumber(value, 2)}</strong>{index > 0 && value !== null && row.weekPph[index - 1] ? <small className={value >= row.weekPph[index - 1]! ? "monthly-positive" : "monthly-negative"}>{value >= row.weekPph[index - 1]! ? <ArrowUpRight size={12} /> : <ArrowDownRight size={12} />}{formatPercent(Math.abs((value - row.weekPph[index - 1]!) / row.weekPph[index - 1]!))}</small> : null}</td>)}<td><span className={row.consecutiveUp >= 2 ? "monthly-pill good" : "monthly-pill"}>{row.consecutiveUp ? `${row.consecutiveUp}周` : "—"}</span></td><td className={(row.cumulativeChange ?? 0) >= 0 ? "monthly-positive" : "monthly-negative"}>{signedPercent(row.cumulativeChange)}</td><td>{formatNumber(row.monthVolume)}</td><td><span className="monthly-pill difficulty">{row.property?.difficulty || "未标注"}</span></td><td>{row.property?.routeHourlyWage ? `$${formatNumber(row.property.routeHourlyWage, 2)}` : "—"}</td><td className={salaryGap(row.property) >= 0 ? "monthly-positive" : "monthly-negative"}>{row.property ? `${salaryGap(row.property) >= 0 ? "+" : ""}$${formatNumber(salaryGap(row.property), 2)}` : "—"}</td></tr>)}</tbody></table>
          </div>
        </section>

        <section className="monthly-grid-two">
          <div id="p75" className="monthly-panel compact"><div className="monthly-section-head"><div><span>TOP QUARTILE</span><h2>P75高PPH路区</h2><p>当前P75标准：{formatNumber(p75, 2)} PPH</p></div><span className="monthly-count">{p75Rows.length}</span></div><div className="monthly-table-wrap"><table className="monthly-table"><thead><tr><th>路区</th><th>月PPH</th><th>高于P75</th><th>妥投量</th><th>难易度</th></tr></thead><tbody>{p75Rows.slice(0, 12).map((row) => <tr key={row.key}><td><strong>{row.route}</strong><small>{row.site} · {row.dsp}</small></td><td><strong>{formatNumber(row.monthPph, 2)}</strong></td><td className="monthly-positive">+{formatNumber(row.monthPph - p75, 2)}</td><td>{formatNumber(row.monthVolume)}</td><td>{row.property?.difficulty || "未标注"}</td></tr>)}</tbody></table></div></div>
          <div id="pressure" className="monthly-panel compact"><div className="monthly-section-head"><div><span>VOLUME UP · PPH FLAT</span><h2>单量上升但PPH未升</h2><p>最近两周妥投量上涨，PPH涨幅≤1%</p></div><span className="monthly-count warning">{volumePressureRows.length}</span></div><div className="monthly-table-wrap"><table className="monthly-table"><thead><tr><th>路区</th><th>单量涨幅</th><th>增加单量</th><th>PPH变化</th><th>难易度</th></tr></thead><tbody>{volumePressureRows.slice(0, 12).map((row) => { const previousVolume = row.weekVolumes.at(-2) ?? 0; const currentVolume = row.weekVolumes.at(-1) ?? 0; const previousPph = row.weekPph.at(-2) ?? 0; const currentPph = row.weekPph.at(-1) ?? 0; const pphChange = previousPph > 0 ? (currentPph - previousPph) / previousPph : 0; return <tr key={row.key}><td><strong>{row.route}</strong><small>{row.site} · {row.dsp}</small></td><td className="monthly-warning">+{formatPercent((currentVolume - previousVolume) / previousVolume)}</td><td>+{formatNumber(currentVolume - previousVolume)}</td><td className={pphChange >= 0 ? "monthly-positive" : "monthly-negative"}>{signedPercent(pphChange)}</td><td>{row.property?.difficulty || "未标注"}</td></tr>; })}</tbody></table></div></div>
        </section>

        <section id="properties" className="monthly-panel">
          <div className="monthly-section-head"><div><span>SALARY & ROUTE PROFILE</span><h2>薪资、难易度与路区综合画像</h2><p>恢复周报中精简掉的属性字段，用于解释效率差异和制定站点动作</p></div><FileSpreadsheet size={28} /></div>
          <div className="monthly-table-wrap"><table className="monthly-table monthly-profile-table"><thead><tr><th>路区</th><th>难易度</th><th>业务模式</th><th>路区单价</th><th>路区时薪</th><th>Amazon时薪中位数</th><th>薪资差</th><th>参考城市</th><th>首单里程</th><th>专家PPH</th><th>妥投异常率</th><th>DNR率</th><th>面积</th><th>人口密度</th><th>地址结构</th></tr></thead><tbody>{routeRows.slice().sort((a, b) => b.monthVolume - a.monthVolume).slice(0, 120).map((row) => <tr key={row.key}><td><strong>{row.route}</strong><small>{row.site} · {row.dsp}</small></td><td><span className="monthly-pill difficulty">{row.property?.difficulty || "未标注"}</span></td><td>{row.property?.businessMode || "未标注"}</td><td>{row.property?.routeUnitPrice ? `$${formatNumber(row.property.routeUnitPrice, 2)}` : "—"}</td><td>{row.property?.routeHourlyWage ? `$${formatNumber(row.property.routeHourlyWage, 2)}` : "—"}</td><td>{row.property?.amazonHourlyMedian ? `$${formatNumber(row.property.amazonHourlyMedian, 2)}` : "—"}</td><td className={salaryGap(row.property) >= 0 ? "monthly-positive" : "monthly-negative"}>{row.property ? `${salaryGap(row.property) >= 0 ? "+" : ""}$${formatNumber(salaryGap(row.property), 2)}` : "—"}</td><td className="monthly-wide-cell">{row.property?.salaryCity || "—"}</td><td>{row.property?.firstMile ? `${formatNumber(row.property.firstMile, 1)} mi` : "—"}</td><td>{row.property?.expertPph ? formatNumber(row.property.expertPph, 2) : "—"}</td><td>{row.property?.deliveryExceptionRate ? formatPercent(row.property.deliveryExceptionRate) : "—"}</td><td>{row.property?.dnrRate ? formatPercent(row.property.dnrRate) : "—"}</td><td>{row.property?.landArea ? formatNumber(row.property.landArea, 1) : "—"}</td><td>{row.property?.populationDensity ? formatNumber(row.property.populationDensity, 1) : "—"}</td><td className="monthly-address-cell">{row.property?.addressMix || "—"}</td></tr>)}</tbody></table></div>
        </section>
        <footer className="monthly-footer"><span>PPH月报系统 · 周数据自动汇总</span><span>单量统一采用妥投量口径</span></footer>
      </main>
    </div>
  );
}
