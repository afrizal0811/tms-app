// File: src/lib/reportGenerators/rangkumanReport.js
'use client';

import { getCachedHubs } from '@/lib/localStorageHandler';
import { formatDateUniversal } from '@/lib/utils';
import * as XLSX from 'xlsx-js-style';
import { getLocalStorage } from '../../localStorageHandler';
import {
  calculateDistanceSummaryData,
  calculatePendingReasonData,
  calculateTimeDriverData,
  calculateTruckDetailData,
  calculateTruckUsageData,
  generateDistanceSummarySheet,
  generatePendingReasonSheet,
  generateRoutingTimeSheet,
  generateServiceLevelSheet,
  generateTaskSummarySheet,
  generateTimeDriverSheet,
  generateTruckDetailSheet,
  generateTruckUsageSheet,
} from './sheets';

export async function generateSummaryDataPreview(
  driverData,
  taskData,
  resultsData,
  locationHistoryData,
  startDateStr,
  endDateStr,
  localeCode,
  activeHubCoords,
  masterTruckData
) {
  const { summaryData, monthTotals } = calculateDistanceSummaryData(
    resultsData,
    startDateStr,
    endDateStr,
    localeCode,
    driverData,
    taskData,
    locationHistoryData
  );

  const truckUsageData = await calculateTruckUsageData(
    resultsData,
    startDateStr,
    endDateStr,
    taskData,
    masterTruckData
  );

  const truckDetailRaw = calculateTruckDetailData(
    driverData,
    resultsData,
    taskData,
    startDateStr,
    endDateStr,
    localeCode,
    activeHubCoords
  );

  const timeDriverRaw = calculateTimeDriverData(
    driverData,
    locationHistoryData,
    startDateStr,
    endDateStr,
    localeCode
  );

  const pendingReasonData = calculatePendingReasonData(
    driverData,
    taskData,
    startDateStr,
    endDateStr
  );

  return {
    startDateStr,
    endDateStr,
    distanceSummaryData: summaryData,
    monthTotals: monthTotals,
    truckUsageData: truckUsageData,
    truckDetailData: { ...truckDetailRaw, driverMap: Object.fromEntries(truckDetailRaw.driverMap) },
    timeDriverData: { ...timeDriverRaw, driverMap: Object.fromEntries(timeDriverRaw.driverMap) },
    pendingReasonsData: pendingReasonData,
  };
}

export async function generateSummaryWorkbook(
  driverData,
  taskData,
  resultsData,
  locationHistoryData,
  startDateStr,
  endDateStr,
  taskSummaryMetrics,
  masterTruckData,
  translate,
  localeCode,
  hasPendingGR,
  pendingDetails,
  isDailyReport = false,
  isManualMode = false
) {
  const wb = XLSX.utils.book_new();
  const hubsList = getCachedHubs() || [];
  const activeHub = hubsList.activeHub;
  activeHub?.lat && activeHub?.lng ? `${activeHub.lat},${activeHub.lng}` : null;
  generateRoutingTimeSheet(wb, taskData, startDateStr, endDateStr, translate, localeCode);
  generateTaskSummarySheet(
    wb,
    taskSummaryMetrics,
    startDateStr,
    endDateStr,
    masterTruckData,
    translate,
    isDailyReport,
    isManualMode
  );

  generatePendingReasonSheet(
    wb,
    driverData,
    taskData,
    translate,
    startDateStr,
    endDateStr,
    hasPendingGR,
    pendingDetails,
    isDailyReport
  );

  generateTimeDriverSheet(
    wb,
    driverData,
    locationHistoryData,
    startDateStr,
    endDateStr,
    translate,
    localeCode,
    taskData,
    resultsData
  );

  generateTruckDetailSheet(
    wb,
    driverData,
    resultsData,
    taskData,
    startDateStr,
    endDateStr,
    translate,
    localeCode
  );

  await generateTruckUsageSheet(
    wb,
    resultsData,
    startDateStr,
    endDateStr,
    translate,
    localeCode,
    taskData,
    masterTruckData,
    isDailyReport
  );
  generateDistanceSummarySheet(
    wb,
    resultsData,
    startDateStr,
    endDateStr,
    translate,
    localeCode,
    driverData,
    taskData,
    locationHistoryData,
    isDailyReport
  );
  if (!isDailyReport) {
    generateServiceLevelSheet(wb, taskData, startDateStr, endDateStr, translate);
  }
  const formattedStart = formatDateUniversal(startDateStr, 'DD.MM.YYYY');
  const formattedEnd = formatDateUniversal(endDateStr, 'DD.MM.YYYY');
  const { storedLocationAcronym: locationName } = getLocalStorage() || '-';
  const excelFileName = `${translate('common.summary')} - (${formattedStart} - ${formattedEnd}) - ${locationName}.xlsx`;

  return { wb, excelFileName };
}
