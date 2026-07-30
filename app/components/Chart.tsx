"use client";

import { useEffect, useRef } from "react";
import type { EChartsOption } from "echarts";

type ChartProps = {
  option: EChartsOption;
  height?: number;
  ariaLabel: string;
};

export default function Chart({
  option,
  height = 300,
  ariaLabel,
}: ChartProps) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    let disposed = false;
    let cleanup = () => {};

    async function renderChart() {
      const echarts = await import("echarts");
      if (disposed || !containerRef.current) return;

      const chart = echarts.init(containerRef.current, undefined, {
        renderer: "canvas",
      });
      chart.setOption(option, { notMerge: true });

      const resizeObserver = new ResizeObserver(() => chart.resize());
      resizeObserver.observe(containerRef.current);
      cleanup = () => {
        resizeObserver.disconnect();
        chart.dispose();
      };
    }

    renderChart();
    return () => {
      disposed = true;
      cleanup();
    };
  }, [option]);

  return (
    <div
      ref={containerRef}
      role="img"
      aria-label={ariaLabel}
      style={{ width: "100%", height }}
    />
  );
}
