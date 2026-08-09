export type PerformanceRecord = {
  week: string;
  weekStart: string;
  route: string;
  dsp: string;
  site: string;
  region: string;
  delivered: number;
  attempted: number;
  sortHours: number;
  transitHours: number;
  deliveryHours: number;
  totalHours: number;
  transitHoursEstimated?: boolean;
  transitHoursAverageBasis?: string;
  estimatedTransitRows?: number;
};

export type PostalPerformanceRecord = Omit<PerformanceRecord, "route"> & {
  postalCode: string;
  /** 统一周数据中的真实所属路区，用于精确核对路区与邮编口径。 */
  route?: string;
};

export type PostalProperty = {
  postalCode: string;
  route: string;
  site: string;
  dsp: string;
  businessMode: string;
  sortCode: string;
  status: string;
  isNew: string;
  difficulty: string;
  firstMile: number;
  expertPph: number;
  deliveryExceptionRate: number;
  dnrRate: number;
  safety: string;
  source: "难易度文件" | "路区邮编对应表" | "统一周数据";
};

export type PostalCost = {
  postalCode: string;
  route: string;
  site: string;
  region: string;
  shipmentVolume: number;
  bookedCost: number;
  averageDspCost: number;
};

export type PostalWeightCost = PostalCost & {
  weightBand: string;
  priceType: string;
};

export type RouteProperty = {
  route: string;
  businessMode: string;
  sortCode: string;
  transferSite: string;
  fleet: string;
  status: string;
  postalCodes: string;
  addressMix: string;
  safety: string;
  landArea: number;
  populationDensity: number;
  isNew: string;
  difficulty: string;
  firstMile: number;
  expertPph: number;
  deliveryExceptionRate: number;
  dnrRate: number;
  routeUnitPrice?: number;
  routeHourlyWage: number;
  amazonHourlyMedian: number;
  salaryCity: string;
};

export type InitialData = {
  meta: {
    sourceRows: number;
    aggregatedRows: number;
    propertyRows: number;
    postalRows?: number;
    postalPropertyRows?: number;
    postalCostRows?: number;
    postalWeightCostRows?: number;
    estimatedTransitRows?: number;
    estimatedPostalTransitRows?: number;
    generatedAt: string;
  };
  records: PerformanceRecord[];
  properties: RouteProperty[];
  postalRecords?: PostalPerformanceRecord[];
  postalProperties?: PostalProperty[];
  postalCosts?: PostalCost[];
  postalWeightCosts?: PostalWeightCost[];
};

export type RouteRow = PerformanceRecord & {
  operationPph: number;
  successPph: number;
  failRate: number;
  wow: number | null;
  percentile: string;
  businessMode: string;
  isNew: string;
};

export type PostalRow = PostalPerformanceRecord & {
  operationPph: number;
  successPph: number;
  failRate: number;
};
