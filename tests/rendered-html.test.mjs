import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request("http://localhost/", {
      headers: { accept: "text/html" },
    }),
    {
      ASSETS: {
        fetch: async () => new Response("Not found", { status: 404 }),
      },
    },
    {
      waitUntil() {},
      passThroughOnException() {},
    },
  );
}

test("server-renders the PPH weekly dashboard shell", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<html[^>]*lang="zh-CN"/i);
  assert.match(html, /<title>PPH周报系统<\/title>/i);
  assert.match(html, /正在汇总路区表现/);
  assert.doesNotMatch(html, /codex-preview/i);
  assert.doesNotMatch(html, /Your site is taking shape/i);
});

test("ships the associated starter dataset and removes disposable preview code", async () => {
  const [page, analytics, layout, packageJson, initialData] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/lib/analytics.ts", import.meta.url), "utf8"),
    readFile(new URL("../app/layout.tsx", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
    readFile(new URL("../public/data/initial.json", import.meta.url), "utf8"),
  ]);

  assert.match(page, /路区名称/);
  assert.match(page, /作业PPH/);
  assert.match(page, /人口密度/);
  assert.match(page, /导出当前结果/);
  assert.match(page, /连续两周PPH上升/);
  assert.match(page, /连续三周PPH上升/);
  assert.match(page, /连续四周PPH上升/);
  assert.match(page, /高PPH值的路区（P75）/);
  assert.match(page, /单量上升但PPH未上升/);
  assert.match(page, /路区时薪/);
  assert.match(page, /手动输入站点名称/);
  assert.match(page, /手动输入路区名称/);
  assert.match(page, /handleRouteSearchChange/);
  assert.match(page, /handleRouteSearchSubmit/);
  assert.match(page, /输入完整路区名称或按回车查看路区详情/);
  assert.match(page, /手动输入DSP名称/);
  assert.match(page, /路区重点名单/);
  assert.match(page, /邮编重点名单/);
  assert.match(page, /路区-邮编搜索/);
  assert.match(page, /路区全量数据/);
  assert.match(page, /邮编全量数据/);
  assert.doesNotMatch(page, /label: "路区难易度全量数据"/);
  assert.match(page, /邮编维度/);
  assert.match(page, /上传邮编数据/);
  assert.match(page, /导出邮编数据/);
  assert.match(page, /postalPriorityLists/);
  assert.match(page, /downloadPostalWatchlist/);
  assert.match(page, /邮编详情/);
  assert.match(page, /selectedPostalChangeSummary/);
  assert.match(page, /postal-change-summary/);
  assert.match(page, /· 从/);
  assert.match(page, /· 到/);
  assert.doesNotMatch(page, /<strong>邮编DSP成本<\/strong>/);
  assert.doesNotMatch(page, /<strong>邮编难易度<\/strong>/);
  assert.match(page, /关联路区/);
  assert.match(page, /openPostalDetails/);
  assert.match(page, /在途时长为0/);
  assert.match(page, /均值补齐/);
  assert.match(page, /estimatedTransitCount/);
  assert.match(page, /路区邮编时薪/);
  assert.match(page, /重量分段单价/);
  assert.match(page, /目标时薪单价计算器/);
  assert.match(page, /建议综合单价/);
  assert.match(page, /重量段建议价/);
  assert.match(page, /routePostalSalaryRows/);
  assert.match(page, /postalWeightCosts/);
  assert.match(page, /selectedRouteContext/);
  assert.match(page, /该路区全部邮编/);
  assert.match(page, /识别状态/);
  assert.match(page, /未同步增长异常/);
  assert.match(page, /邮编与路区单量增加异常/);
  assert.match(page, /drawer-postal-warning/);
  assert.match(page, /expectedVolumeIncrease/);
  assert.match(page, /路区单量增加/);
  assert.match(page, /is-highlighted/);
  assert.match(page, /下载路区邮编变化Excel/);
  assert.match(page, /downloadRoutePostalImpactExcel/);
  assert.match(page, /XLSX\.writeFile/);
  assert.match(analytics, /imputeTransitHours/);
  assert.match(analytics, /同站点同周单均在途时长均值/);
  assert.match(page, /volumeIncrease/);
  assert.match(page, /pphChange <= 0\.01/);
  assert.match(page, /downloadWatchlist/);
  assert.match(page, /currentRegionName}PPH周报\.html/);
  assert.match(page, /__PPH_LOCKED_REGION__/);
  assert.match(page, /scopedRecords/);
  assert.match(page, /lastIndexOf\(initialDataMarker\)/);
  assert.match(page, /P75 \/ P25 路区观察/);
  assert.match(page, /返回路区/);
  assert.match(page, /route-change-summary/);
  assert.match(page, /不使用当前周PPH重新计算/);
  assert.match(page, /薪资文件效率基准/);
  assert.match(page, /addressDistributionDifference/);
  assert.match(page, /商业、公寓、学校、山区等难送地址按1\.5倍加权/);
  assert.match(page, /PPH仅作为结果对比，不参与相似度排名/);
  assert.doesNotMatch(page, /<strong>时薪对比<\/strong>/);
  assert.doesNotMatch(page, /路区时薪（薪资文件）/);
  assert.doesNotMatch(page, /<strong>路区难易度<\/strong>/);
  assert.doesNotMatch(page, /<strong>收件地址类型占比<\/strong>/);
  assert.doesNotMatch(page, /<strong>异常原因<\/strong>/);
  assert.doesNotMatch(page, /件均DSP成本 × 邮编作业PPH/);
  assert.doesNotMatch(page, /人口密度 × 作业PPH/);
  assert.doesNotMatch(page, /label="派送失败率"/);
  assert.doesNotMatch(page, /本区低位分界值/);
  assert.doesNotMatch(page, /本区高位分界值/);
  assert.doesNotMatch(page, /作业与成功 PPH 趋势/);
  assert.doesNotMatch(page, /NETWORK BENCHMARK/);
  assert.doesNotMatch(page, /站点效率对比/);
  assert.doesNotMatch(
    page,
    /eyebrow="TIME MIX"[\s\S]{0,200}title="耗时结构"/,
  );
  assert.match(layout, /PPH周报系统/);
  assert.doesNotMatch(packageJson, /react-loading-skeleton/);
  assert.doesNotMatch(page, /_sites-preview|SkeletonPreview/);
  const parsedInitialData = JSON.parse(initialData);
  assert.ok(parsedInitialData.records.length > 1000);
  assert.ok(parsedInitialData.postalRecords.length > 1000);
  assert.ok(parsedInitialData.postalProperties.length > 10000);
  assert.ok(parsedInitialData.postalCosts.length > 9000);
  assert.ok(parsedInitialData.postalWeightCosts.length > 80000);
  const performancePostalCodes = new Set(
    parsedInitialData.postalRecords.map((row) => row.postalCode),
  );
  const propertyPostalCodes = new Set(
    parsedInitialData.postalProperties.map((row) => row.postalCode),
  );
  assert.ok(
    [...performancePostalCodes].every((postalCode) =>
      propertyPostalCodes.has(postalCode),
    ),
  );
  assert.ok(
    parsedInitialData.postalRecords.some((row) =>
      /^0\d{4}$/.test(row.postalCode),
    ),
  );
  assert.ok(
    parsedInitialData.properties.some((row) => row.routeHourlyWage > 0),
  );
  assert.ok(
    parsedInitialData.properties.some((row) => row.routeUnitPrice > 0),
  );
  assert.ok(
    parsedInitialData.properties.some((row) => row.amazonHourlyMedian > 0),
  );
  await assert.rejects(access(new URL("../app/_sites-preview", import.meta.url)));
});

test("ships a self-contained interactive HTML dashboard", async () => {
  const html = await readFile(
    new URL("../public/PPH周报系统-交互版.html", import.meta.url),
    "utf8",
  );
  const bundleStart = html.indexOf('<script type="module">');
  const bundleEnd = html.indexOf("</script>", bundleStart);
  const initialDataStart = html.lastIndexOf(
    '<script id="pph-initial-data">',
  );
  const documentShell = html.replace(
    /<script type="module">[\s\S]*?<\/script>/,
    "",
  );

  assert.ok(html.length > 5_000_000);
  assert.ok(bundleStart > 0);
  assert.ok(initialDataStart > bundleEnd);
  assert.equal((html.match(/<\/script>/gi) ?? []).length, 2);
  assert.match(html, /id="pph-initial-data"/);
  assert.match(html, /__PPH_LOCKED_REGION__="WE"/);
  assert.match(documentShell, /<style>[\s\S]+<\/style>/);
  assert.doesNotMatch(documentShell, /<script[^>]+src=/i);
  assert.doesNotMatch(documentShell, /<link[^>]+rel="stylesheet"/i);
  assert.match(html, /LAX02-025/);
  assert.match(html, /下载HTML周报/);
  assert.match(html, /standalone-report/);
});
