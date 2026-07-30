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
};

export type InitialData = {
  meta: {
    sourceRows: number;
    aggregatedRows: number;
    propertyRows: number;
    generatedAt: string;
  };
  records: PerformanceRecord[];
  properties: RouteProperty[];
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
