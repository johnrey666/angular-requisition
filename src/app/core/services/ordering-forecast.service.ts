import { Injectable } from '@angular/core';
import { DatabaseService } from './database.service';

export interface ForecastSku {
  skuCode: string;
  skuName: string;
  historicalAvgPerMonth: number;
  projectedCount: number;
}

export interface ForecastMaterial {
  name: string;
  unit: string;
  historicalAvgPerMonth: number;
  projectedQty: number;
}

export interface OrderingForecastResult {
  topSkus: ForecastSku[];
  topMaterials: ForecastMaterial[];
  hasEnoughData: boolean;
  forecastDays: number;
}

const HISTORY_DAYS = 90;
const MIN_WEEKS_FOR_FORECAST = 3;
const TOP_COUNT = 5;

@Injectable({ providedIn: 'root' })
export class OrderingForecastService {
  constructor(private db: DatabaseService) {}

  computeForecast(
    requisitions: any[],
    skuCache: Map<string, any[]>,
    forecastDays: number,
    getRequisitionDate: (req: any) => Date | null
  ): OrderingForecastResult {
    const empty: OrderingForecastResult = {
      topSkus: [],
      topMaterials: [],
      hasEnoughData: false,
      forecastDays,
    };

    const historical = this.filterHistorical(requisitions, getRequisitionDate);
    if (historical.length === 0) {
      return empty;
    }

    const weekKeys = this.collectWeekKeys(historical, getRequisitionDate);
    if (weekKeys.size < MIN_WEEKS_FOR_FORECAST) {
      return empty;
    }

    const trendFactor = this.computeTrendFactor(historical, getRequisitionDate);
    const weeksInHistory = weekKeys.size;
    const monthsInHistory = HISTORY_DAYS / 30;

    const topSkus = this.forecastSkus(
      historical,
      getRequisitionDate,
      weeksInHistory,
      monthsInHistory,
      forecastDays,
      trendFactor
    );

    const topMaterials = this.forecastMaterials(
      historical,
      skuCache,
      getRequisitionDate,
      weeksInHistory,
      monthsInHistory,
      forecastDays,
      trendFactor
    );

    return {
      topSkus,
      topMaterials,
      hasEnoughData: topSkus.length > 0 || topMaterials.length > 0,
      forecastDays,
    };
  }

  private filterHistorical(
    requisitions: any[],
    getRequisitionDate: (req: any) => Date | null
  ): any[] {
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - HISTORY_DAYS);
    cutoff.setHours(0, 0, 0, 0);

    return requisitions.filter(req => {
      if ((req.status || '') === 'Removed') return false;
      const date = getRequisitionDate(req);
      return date !== null && date >= cutoff;
    });
  }

  private collectWeekKeys(
    requisitions: any[],
    getRequisitionDate: (req: any) => Date | null
  ): Set<string> {
    const keys = new Set<string>();
    for (const req of requisitions) {
      const date = getRequisitionDate(req);
      if (date) keys.add(this.getIsoWeekKey(date));
    }
    return keys;
  }

  private computeTrendFactor(
    requisitions: any[],
    getRequisitionDate: (req: any) => Date | null
  ): number {
    const weeklyTotals = new Map<string, number>();

    for (const req of requisitions) {
      const date = getRequisitionDate(req);
      if (!date) continue;
      const key = this.getIsoWeekKey(date);
      weeklyTotals.set(key, (weeklyTotals.get(key) || 0) + 1);
    }

    const sortedWeeks = [...weeklyTotals.keys()].sort();
    if (sortedWeeks.length < 4) return 1;

    const recentWeeks = sortedWeeks.slice(-4);
    const priorWeeks = sortedWeeks.slice(-8, -4);

    if (priorWeeks.length === 0) return 1;

    const recentAvg = recentWeeks.reduce((sum, w) => sum + (weeklyTotals.get(w) || 0), 0) / recentWeeks.length;
    const priorAvg = priorWeeks.reduce((sum, w) => sum + (weeklyTotals.get(w) || 0), 0) / priorWeeks.length;

    if (priorAvg === 0) return 1;

    return Math.min(1.3, Math.max(0.7, recentAvg / priorAvg));
  }

  private forecastSkus(
    historical: any[],
    getRequisitionDate: (req: any) => Date | null,
    weeksInHistory: number,
    monthsInHistory: number,
    forecastDays: number,
    trendFactor: number
  ): ForecastSku[] {
    const skuWeekly = new Map<string, Map<string, number>>();
    const skuNames = new Map<string, string>();

    for (const req of historical) {
      const code = this.db.normalizeSkuCode(req.sku_code || req.skuCode || '');
      if (!code) continue;

      const date = getRequisitionDate(req);
      if (!date) continue;

      const weekKey = this.getIsoWeekKey(date);
      if (!skuWeekly.has(code)) skuWeekly.set(code, new Map());
      const weeks = skuWeekly.get(code)!;
      weeks.set(weekKey, (weeks.get(weekKey) || 0) + 1);

      if (!skuNames.has(code)) {
        skuNames.set(code, String(req.skuName || req.sku_name || code).trim());
      }
    }

    const results: ForecastSku[] = [];

    for (const [code, weeks] of skuWeekly) {
      const totalCount = [...weeks.values()].reduce((sum, c) => sum + c, 0);
      const weeklyAvg = totalCount / weeksInHistory;
      const projectedCount = Math.round(weeklyAvg * (forecastDays / 7) * trendFactor * 10) / 10;
      const historicalAvgPerMonth = Math.round((totalCount / monthsInHistory) * 10) / 10;

      results.push({
        skuCode: code,
        skuName: skuNames.get(code) || code,
        historicalAvgPerMonth,
        projectedCount,
      });
    }

    return results
      .sort((a, b) => b.projectedCount - a.projectedCount)
      .slice(0, TOP_COUNT);
  }

  private forecastMaterials(
    historical: any[],
    skuCache: Map<string, any[]>,
    getRequisitionDate: (req: any) => Date | null,
    weeksInHistory: number,
    monthsInHistory: number,
    forecastDays: number,
    trendFactor: number
  ): ForecastMaterial[] {
    const materialWeekly = new Map<string, Map<string, number>>();
    const materialUnits = new Map<string, string>();

    for (const req of historical) {
      const skuCode = this.db.normalizeSkuCode(req.sku_code || req.skuCode || '');
      if (!skuCode) continue;

      const date = getRequisitionDate(req);
      if (!date) continue;

      const weekKey = this.getIsoWeekKey(date);
      const materials = skuCache.get(skuCode) || [];
      const qtyRequired = Number(req.qty_needed ?? req.quantity ?? 0);

      for (const mat of materials) {
        if (!mat.raw_material) continue;
        const qty = Number(mat.quantity_per_batch ?? 0) * qtyRequired;
        const name = mat.raw_material;

        if (!materialWeekly.has(name)) materialWeekly.set(name, new Map());
        const weeks = materialWeekly.get(name)!;
        weeks.set(weekKey, (weeks.get(weekKey) || 0) + qty);

        if (!materialUnits.has(name)) {
          materialUnits.set(name, mat.unit || '');
        }
      }
    }

    const results: ForecastMaterial[] = [];

    for (const [name, weeks] of materialWeekly) {
      const totalQty = [...weeks.values()].reduce((sum, q) => sum + q, 0);
      const weeklyAvg = totalQty / weeksInHistory;
      const projectedQty = Math.round(weeklyAvg * (forecastDays / 7) * trendFactor * 100) / 100;
      const historicalAvgPerMonth = Math.round((totalQty / monthsInHistory) * 100) / 100;

      results.push({
        name,
        unit: materialUnits.get(name) || '',
        historicalAvgPerMonth,
        projectedQty,
      });
    }

    return results
      .sort((a, b) => b.projectedQty - a.projectedQty)
      .slice(0, TOP_COUNT);
  }

  private getIsoWeekKey(date: Date): string {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
    return `${d.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
  }
}
