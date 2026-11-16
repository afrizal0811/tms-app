import {
  getLocationHistories,
  getResults,
  getTasks,
  getVehicleTypes,
  getPendingDetails,
} from '@/lib/api/mileapp';
import { getCachedHubs, getLocalStorage } from '@/lib/localStorageHandler';
import { convertLocationHistories } from '@/lib/reportGenerators/helper';
import {
  parseDeliveryToTasks,
  parseRoutingToResults,
} from '@/lib/reportGenerators/reports/daily/parseManualToApi';
import { generateSummaryWorkbook } from '@/lib/reportGenerators/summary/summaryReport';
import { calculateTaskSummaryMetrics } from '@/lib/reportGenerators/summary/taskSummaryMetrics';
import { applyPendingDetailsToTasks } from '@/lib/deliveryStatus';
import { toastError, toastSuccess } from '@/lib/toast';
import {
  calculateStartFinishDates,
  formatDateUniversal,
  isEmpty,
  toApiDateString,
} from '@/lib/utils';
import * as XLSX from 'xlsx-js-style';

export const getManualDate = (headerName, deliveryBuffers, fallbackDate) => {
  try {
    const dates = [];

    for (const buf of deliveryBuffers) {
      const wbInput = XLSX.read(buf, { type: 'array' });
      const sheetName =
        wbInput.SheetNames.find((s) => s.toLowerCase() === 'main') || wbInput.SheetNames[0];
      const rawRows = XLSX.utils.sheet_to_json(wbInput.Sheets[sheetName], { header: 1 });

      const headIdx = rawRows.findIndex((row) => {
        if (!Array.isArray(row)) return false;
        const rStr = row.map((c) => String(c).toLowerCase().trim());
        return rStr.includes('_id') && rStr.includes('flow');
      });

      if (headIdx === -1) continue;

      const headers = rawRows[headIdx].map((h) => String(h).toLowerCase().trim());
      const idxStart = headers.indexOf(headerName);

      if (idxStart === -1) continue;

      for (let i = headIdx + 1; i < rawRows.length; i++) {
        const dPart = rawRows[i]?.[idxStart]
          ? String(rawRows[i][idxStart]).split(/[T\s]/)[0]
          : null;
        if (dPart) dates.push(dPart);
      }
    }

    if (dates.length === 0) return fallbackDate;

    const majorityDate = Object.entries(
      dates.reduce((acc, d) => ({ ...acc, [d]: (acc[d] || 0) + 1 }), {})
    ).sort(([, a], [, b]) => b - a)[0][0];

    const parts = majorityDate.split(/[-/]/);
    if (parts.length !== 3) return majorityDate;

    if (parts[0].length === 4)
      return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    if (parts[2].length === 4)
      return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;

    return majorityDate;
  } catch (e) {
    console.error(e);
    return fallbackDate;
  }
};

export const handleSingleDownload = async ({
  hubName,
  selectedDate,
  selectedDateString,
  isCustomRouting,
  routingDate,
  driverData,
  setIsLoading,
  t,
  isIndonesian,
}) => {
  const localeCode = isIndonesian ? 'id' : 'en';
  try {
    setIsLoading(true);

    const dateStr = selectedDateString || (selectedDate ? formatDateUniversal(selectedDate) : null);
    if (!dateStr) throw new Error(t('common.invalid_date'));
    const routingDeliveryDate =
      isCustomRouting && routingDate ? formatDateUniversal(new Date(routingDate)) : dateStr;

    const timeFromTasks = new Date(`${dateStr}T00:00:00`).toISOString();
    const timeToTasks = new Date(`${dateStr}T23:59:59`).toISOString();

    const allTasks = await getTasks({
      status: 'DONE,ONGOING',
      timeFrom: timeFromTasks,
      timeTo: timeToTasks,
    });

    if (isEmpty(allTasks)) {
      throw new Error(t('common.no_data'));
    }

    const locStartObj = new Date(dateStr);
    locStartObj.setDate(locStartObj.getDate() - 3);
    locStartObj.setHours(0, 0, 0, 0);

    const locEndObj = new Date(dateStr);
    locEndObj.setDate(locEndObj.getDate() + 2);
    locEndObj.setHours(23, 59, 59, 999);

    const { storedLocationAcronym } = getLocalStorage();
    const [filteredResults, hubsData, locationHistoriesRes] = await Promise.all([
      getResults({
        deliveryDate: routingDeliveryDate,
      }),
      getCachedHubs(),
      getLocationHistories({
        date: routingDeliveryDate,
        timeFrom: toApiDateString(locStartObj),
        timeTo: toApiDateString(locEndObj),
      }).catch(() => []),
    ]);

    const singleDateHistories = (locationHistoriesRes || []).filter((item) => {
      return item.startTime?.startsWith(dateStr);
    });
    const { timeDataObjects } = convertLocationHistories(singleDateHistories, driverData);
    const filteredTimeData = timeDataObjects.filter(
      (item) => !isEmpty(item.startTimeFmt) && !isEmpty(item.finishTimeFmt)
    );

    if (isEmpty(filteredResults) && isEmpty(allTasks) && isEmpty(filteredTimeData)) {
      throw new Error(t('common.no_data'));
    }

    const hasPendingGR = hubsData.activeHub ? hubsData.activeHub.hasPendingGR : false;
    const hubLabel = storedLocationAcronym || hubName;

    const [pendingDetails, masterTruckData] = await Promise.all([
      getPendingDetails(dateStr, dateStr).catch(() => []),
      getVehicleTypes(),
    ]);
    const effectiveTasks = applyPendingDetailsToTasks(allTasks, pendingDetails);
    const taskSummaryMetrics = await calculateTaskSummaryMetrics({
      allTasks: effectiveTasks,
      allResults: filteredResults,
      fetchedDrivers: driverData,
      hasPendingGR,
      t,
    });

    const { wb } = await generateSummaryWorkbook(
      driverData,
      effectiveTasks,
      filteredResults,
      singleDateHistories,
      dateStr,
      dateStr,
      taskSummaryMetrics,
      masterTruckData,
      t,
      localeCode,
      hasPendingGR,
      pendingDetails,
      true
    );

    const baseDailyReport = t('common.daily_report');
    const formattedDate = formatDateUniversal(dateStr, 'DD.MM.YYYY');
    const singleExcelFileName = `${baseDailyReport} - ${formattedDate} - ${hubLabel}.xlsx`;

    XLSX.writeFile(wb, singleExcelFileName);
    toastSuccess(t('common.toast.success'));
  } catch (err) {
    toastError(t('common.toast.error', { err: err.message }), err);
  } finally {
    setIsLoading(false);
  }
};

export const handleBulkDownload = async ({
  hubName,
  startDate,
  endDate,
  driverData,
  setIsLoading,
  t,
  isIndonesian,
}) => {
  const localeCode = isIndonesian ? 'id' : 'en';
  try {
    setIsLoading(true);

    const startStr = formatDateUniversal(startDate);
    const endStr = formatDateUniversal(endDate);

    const taskStartObj = new Date(startDate);
    taskStartObj.setDate(taskStartObj.getDate() - 4);
    taskStartObj.setHours(0, 0, 0, 0);

    const taskEndObj = new Date(endDate);
    taskEndObj.setDate(taskEndObj.getDate() + 4);
    taskEndObj.setHours(23, 59, 59, 999);


    const locStartObj = new Date(startDate);
    locStartObj.setDate(locStartObj.getDate() - 3);
    locStartObj.setHours(0, 0, 0, 0);

    const locEndObj = new Date(endDate);
    locEndObj.setDate(locEndObj.getDate() + 2);
    locEndObj.setHours(23, 59, 59, 999);

    const createDateChunks = (start, end, maxDays) => {
      const chunks = [];
      let curr = new Date(start);
      while (curr <= end) {
        let next = new Date(curr);
        next.setDate(next.getDate() + maxDays - 1);
        next.setHours(23, 59, 59, 999);
        if (next > end) next = new Date(end);
        chunks.push({ from: toApiDateString(curr), to: toApiDateString(next) });
        curr = new Date(next);
        curr.setDate(curr.getDate() + 1);
        curr.setHours(0, 0, 0, 0);
      }
      return chunks;
    };

    const taskRanges = createDateChunks(taskStartObj, taskEndObj, 5);
    const historyRanges = createDateChunks(locStartObj, locEndObj, 7);

    const mergeResults = (resArray) => {
      let merged = [];
      resArray.forEach((res) => {
        if (Array.isArray(res)) merged = [...merged, ...res];
        else if (res?.data) merged = [...merged, ...res.data];
        else if (res?.tasks?.data) merged = [...merged, ...res.tasks.data];
      });
      return merged;
    };

    const pTasks = (async () => {
      const rawResults = [];
      for (const range of taskRanges) {
        const res = await getTasks({
          status: 'ONGOING,DONE',
          timeFrom: range.from,
          timeTo: range.to,
        });
        rawResults.push(res);
      }
      return mergeResults(rawResults);
    })();

    const pRouting = (async () => {
      const res = await getResults({
        dateFrom: startStr,
        dateTo: endStr,
      });
      return Array.isArray(res) ? res : res?.data || [];
    })();

    const pHistory = (async () => {
      const rawResults = [];
      for (const range of historyRanges) {
        const res = await getLocationHistories({ timeFrom: range.from, timeTo: range.to });
        rawResults.push(res);
      }
      return mergeResults(rawResults);
    })();

    const [tasksRes, resultsRes, locRes, hubsDB, masterRes, pendingDetails] = await Promise.all([
      pTasks,
      pRouting,
      pHistory,
      getCachedHubs(),
      getVehicleTypes(),
      getPendingDetails(startStr, endStr).catch(() => []),
    ]);

    const hasPendingGR = hubsDB?.activeHub?.hasPendingGR || false;

    if (isEmpty(tasksRes) && isEmpty(resultsRes)) {
      throw new Error(t('common.no_data'));
    }

    const effectiveTasks = applyPendingDetailsToTasks(tasksRes, pendingDetails);

    const taskSummaryMetrics = await calculateTaskSummaryMetrics({
      allTasks: effectiveTasks,
      allResults: resultsRes,
      fetchedDrivers: driverData,
      hasPendingGR,
      t,
    });

    const { wb } = await generateSummaryWorkbook(
      driverData,
      effectiveTasks,
      resultsRes,
      locRes,
      startStr,
      endStr,
      taskSummaryMetrics,
      masterRes,
      t,
      localeCode,
      hasPendingGR,
      pendingDetails,
      true
    );

    const { storedLocationAcronym } = getLocalStorage();
    const hubLabel = storedLocationAcronym || hubName;
    const baseDailyReport = t('common.daily_report');
    const bulkLabel = t('common.bulk');
    const formattedStart = formatDateUniversal(startStr, 'DD.MM.YYYY');
    const formattedEnd = formatDateUniversal(endStr, 'DD.MM.YYYY');

    const bulkExcelFileName = `${baseDailyReport} (${bulkLabel})- (${formattedStart} - ${formattedEnd}) - ${hubLabel}.xlsx`;
    XLSX.writeFile(wb, bulkExcelFileName);
    toastSuccess(t('common.toast.success'));
  } catch (err) {
    toastError(t('common.toast.error', { err: err.message }), err);
  } finally {
    setIsLoading(false);
  }
};

export const handleManualDownload = async ({
  hubName,
  selectedDateString,
  selectedRoutingFiles,
  selectedDeliveryFiles,
  driverData,
  setIsLoading,
  setIsModalOpen,
  setSelectedRoutingFiles,
  setSelectedDeliveryFiles,
  t,
  isIndonesian,
}) => {
  const localeCode = isIndonesian ? 'id' : 'en';
  try {
    setIsLoading(true);

    const { storedLocationAcronym } = getLocalStorage();
    const hubLabel = storedLocationAcronym || hubName;

    const routingBuffers = await Promise.all(
      selectedRoutingFiles.map((file) => file.arrayBuffer())
    );
    const deliveryBuffers = await Promise.all(
      selectedDeliveryFiles.map((file) => file.arrayBuffer())
    );

    const extractedStartDate = getManualDate('starttime', deliveryBuffers, selectedDateString);
    const { timeFrom, timeTo } = calculateStartFinishDates(extractedStartDate);
    const [vehicleTypes, [hubsData, locationHistoriesRes]] = await Promise.all([
      getVehicleTypes(),
      Promise.all([
        getCachedHubs(),
        getLocationHistories({
          date: extractedStartDate,
          timeFrom,
          timeTo,
        }),
      ]),
    ]);

    const singleDateHistories = (locationHistoriesRes || []).filter((item) => {
      return item.startTime?.startsWith(extractedStartDate);
    });

    const allTasks = await parseDeliveryToTasks(deliveryBuffers);
    const filteredResults = await parseRoutingToResults(
      routingBuffers,
      extractedStartDate
    );
    const hasPendingGR = hubsData.activeHub ? hubsData.activeHub.hasPendingGR : false;

    const pendingDetails = await getPendingDetails(extractedStartDate, extractedStartDate).catch(
      () => []
    );

    const effectiveTasks = applyPendingDetailsToTasks(allTasks, pendingDetails);

    const taskSummaryMetrics = await calculateTaskSummaryMetrics({
      allTasks: effectiveTasks,
      allResults: filteredResults,
      fetchedDrivers: driverData,
      hasPendingGR,
      t,
      isManualMode: true,
    });

    const { wb } = await generateSummaryWorkbook(
      driverData,
      effectiveTasks,
      filteredResults,
      singleDateHistories,
      extractedStartDate,
      extractedStartDate,
      taskSummaryMetrics,
      vehicleTypes,
      t,
      localeCode,
      hasPendingGR,
      pendingDetails,
      true, // isDailyReport
      true // isManualMode
    );

    const baseDailyReport = t('common.daily_report');
    const formattedDate = formatDateUniversal(extractedStartDate, 'DD.MM.YYYY');
    const excelFileName = `${baseDailyReport} (Manual) - ${formattedDate} - ${hubLabel}.xlsx`;

    XLSX.writeFile(wb, excelFileName);
    toastSuccess(t('common.toast.success'));
    setIsModalOpen(false);
    setSelectedRoutingFiles([]);
    setSelectedDeliveryFiles([]);
  } catch (err) {
    toastError(t('common.toast.error', { err: err.message }), err);
  } finally {
    setIsLoading(false);
  }
};
