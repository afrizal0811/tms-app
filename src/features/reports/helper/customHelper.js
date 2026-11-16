import { getLocationHistories, getResult, getResults, getTasks } from '@/lib/api/mileapp';
import { getCachedHubs } from '@/lib/localStorageHandler';
import { convertLocationHistories } from '@/lib/reportGenerators/helper';
import {
  buildDriverMap,
  buildRoutingMap,
  buildSyncTimeMap,
  extractBreadRows,
  generateBreadWorkbook,
  generateTaskDetailWorkbook,
  generateTaskManualDetailWorkbook,
  groupTasksByDriver,
} from '@/lib/reportGenerators/reports';
import { toastError, toastSuccess } from '@/lib/toast';
import {
  calculateStartFinishDates,
  formatDateUniversal,
  isEmpty,
  parseCustomerString,
  toApiDateString,
} from '@/lib/utils';
import * as XLSX from 'xlsx-js-style';
import {
  serviceLevelHeaders,
  taskHeaders,
  taskKeyMapping,
  taskManualHeaders,
  taskManualKeyMapping,
  tripActivityHeaders,
} from './constants';
import { bulkExcelDownloader, bulkZipDownloader, getDatesInRange } from './help';

const normalizeTasksData = (tasks) =>
  !isEmpty(tasks) && Array.isArray(tasks) ? tasks : tasks?.data || [];

const getReportDates = (start, end) => {
  const localStart = new Date(start);
  localStart.setHours(0, 0, 0, 0);

  const validEndDate = end ? new Date(end) : new Date(start);
  validEndDate.setHours(23, 59, 59, 999);

  const timeFromUtc = toApiDateString(localStart);
  const timeToUtc = toApiDateString(validEndDate);

  const startString = formatDateUniversal(localStart);
  const endString = formatDateUniversal(validEndDate);

  const { timeFrom: locTimeFrom } = calculateStartFinishDates(startString);
  const { timeTo: locTimeTo } = calculateStartFinishDates(endString);

  return { timeFromUtc, timeToUtc, locTimeFrom, locTimeTo, selectedDateString: startString };
};

export const handleCustomDownload = async ({
  isBulkMode,
  bulkFormat = 'zip',
  startDate,
  endDate,
  singleDate,
  driverData,
  hubAcronym,
  hubName,
  t,
  setIsLoading,
  reportType,
}) => {
  if (isEmpty(driverData)) {
    toastError(t('common.no_driver'));
    return;
  }
  setIsLoading(true);
  try {
    const locationName = hubAcronym || hubName;

    const reportTypeConfig = {
      bread: { process: processBreadReport, title: t('report.custom.bread_items') },
      detail: { process: processTaskRoutingReport, title: t('report.custom.task_routing') },
      manual: { process: processTaskManualReport, title: t('report.custom.task_manual') },
      service_level: {
        process: processServiceLevelReport,
        title: t('report.custom.service_level'),
      },
      trip_activity: {
        process: processTripActivityReport,
        title: t('report.custom.trip_activity'),
      },
    };

    const config = reportTypeConfig[reportType];
    if (!config) throw new Error('Invalid report type');

    if (isBulkMode) {
      if (!startDate || !endDate) throw new Error(t('common.invalid_date'));
      if (endDate < startDate)
        throw new Error('Tanggal akhir tidak boleh kurang dari tanggal awal.');

      if (bulkFormat === 'xls') {
        await bulkExcelDownloader({
          startDate,
          endDate,
          locationName,
          title: config.title,
          setIsLoading,
          t,
          fetchFilesCallback: async () => {
            const datesToProcess = getDatesInRange(startDate, endDate || startDate);
            return await config.process({ datesToProcess, locationName, t, driverData });
          },
        });
      } else {
        await bulkZipDownloader({
          startDate,
          endDate,
          driverData,
          zipPrefix: `${config.title} (${t('common.bulk')})`,
          setIsLoading,
          processDateCallback: async ({ dateObj }) => {
            const files = await config.process({
              datesToProcess: [dateObj],
              locationName,
              t,
              driverData,
            });
            if (files.length === 0) return { error: true };
            return { wb: files[0].wb, excelFileName: files[0].fileName };
          },
          t,
        });
      }
    } else {
      const files = await config.process({
        datesToProcess: [singleDate],
        locationName,
        t,
        driverData,
      });
      if (files.length === 0) throw new Error(t('common.no_data'));
      XLSX.writeFile(files[0].wb, files[0].fileName);
      toastSuccess(t('common.toast.success'));
    }
  } catch (error) {
    toastError(t('common.toast.error', { err: error.message }), error);
  } finally {
    if (!isBulkMode) setIsLoading(false);
  }
};

export const processBreadReport = async ({ datesToProcess, locationName, t, driverData }) => {
  const driverMap = buildDriverMap(driverData);
  const generatedFiles = [];

  for (const date of datesToProcess) {
    const { timeFromUtc, timeToUtc } = getReportDates(date, date);

    const tasks = await getTasks({
      status: 'DONE,ONGOING',
      timeFrom: timeFromUtc,
      timeTo: timeToUtc,
    });

    const tasksData = normalizeTasksData(tasks);
    if (isEmpty(tasksData)) continue;

    const dateStr = formatDateUniversal(date, 'DD-MM-YYYY');
    const rows = extractBreadRows(tasksData, driverMap, dateStr);
    if (rows.length === 0) continue;

    const wb = generateBreadWorkbook(rows, t);
    const fileDateStr = formatDateUniversal(date, 'DD.MM.YYYY');
    const fileName = `${t('report.custom.bread_items')} - ${fileDateStr} - ${locationName}.xlsx`;
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    generatedFiles.push({ fileName, wb, wbout });
  }

  return generatedFiles;
};

export const processTaskRoutingReport = async ({ datesToProcess, locationName, t, driverData }) => {
  const hubsList = getCachedHubs() || [];
  const activeHub = hubsList.activeHub;
  const hubCoordsStr =
    activeHub?.lat && activeHub?.lng ? `${activeHub.lat},${activeHub.lng}` : null;

  const generatedFiles = [];

  for (const date of datesToProcess) {
    const { timeFromUtc, timeToUtc, locTimeFrom, locTimeTo, selectedDateString } = getReportDates(
      date,
      date
    );

    const [tasks, locHistories] = await Promise.all([
      getTasks({
        status: 'DONE,ONGOING',
        timeFrom: timeFromUtc,
        timeTo: timeToUtc,
        isNeedFields: false,
        source: 'vendor',
      }),
      getLocationHistories({
        date: selectedDateString,
        timeFrom: locTimeFrom,
        timeTo: locTimeTo,
      }),
    ]);

    const tasksData = normalizeTasksData(tasks);
    if (isEmpty(tasksData)) continue;

    const uniqueRoutingIds = [
      ...new Set(tasksData.map((task) => task.routingResultId).filter(Boolean)),
    ];

    const routingResults = await Promise.all(uniqueRoutingIds.map((id) => getResult(id)));
    const routingMap = buildRoutingMap(routingResults, driverData);

    const timeMap = buildSyncTimeMap(locHistories, driverData, selectedDateString);
    const groupedData = groupTasksByDriver(tasksData);

    const wb = generateTaskDetailWorkbook(
      groupedData,
      timeMap,
      routingMap,
      hubCoordsStr,
      taskHeaders,
      taskKeyMapping
    );

    const dateStr = formatDateUniversal(date, 'DD.MM.YYYY');
    const fileName = `${t('report.custom.task_routing')} - ${dateStr} - ${locationName}.xlsx`;
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    generatedFiles.push({ fileName, wb, wbout });
  }

  return generatedFiles;
};

export const processTaskManualReport = async ({ datesToProcess, locationName, t }) => {
  const generatedFiles = [];

  for (const date of datesToProcess) {
    const { timeFromUtc, timeToUtc } = getReportDates(date, date);

    const tasks = await getTasks({
      status: 'DONE,ONGOING',
      timeFrom: timeFromUtc,
      timeTo: timeToUtc,
      isNeedFields: false,
      source: 'vendor',
    });

    const tasksData = normalizeTasksData(tasks);
    if (isEmpty(tasksData)) continue;

    const uniqueRoutingIds = [
      ...new Set(tasksData.map((task) => task.routingResultId).filter(Boolean)),
    ];

    const deliveryDateStr = formatDateUniversal(date, 'YYYY-MM-DD');
    let routingResults = await getResults({ deliveryDate: deliveryDateStr });
    if (isEmpty(routingResults) && uniqueRoutingIds.length > 0) {
      routingResults = await Promise.all(
        uniqueRoutingIds.map((id) => getResult(id).catch(() => null))
      );
      routingResults = routingResults.filter(Boolean);
    }

    const parseHistories = (histories) => {
      if (!histories) return null;
      if (typeof histories === 'string') {
        try {
          return JSON.parse(histories);
        } catch {
          return null;
        }
      }
      return histories;
    };

    const overrideTaskMap = new Map();
    (routingResults || []).forEach((res) => {
      const histData = parseHistories(res.histories);
      const historyObj = histData?.history;
      if (!historyObj) return;
      const h = Array.isArray(historyObj) ? historyObj[0] : historyObj;
      if (!h) return;

      ['move', 'dropped', 'switch', 'change', 'manual'].forEach((key) => {
        const items = Array.isArray(h[key]) ? h[key] : h[key]?.data || [];
        items.forEach((m) => {
          const ver = Number(m.version) || 0;

          let targetVisits = m.visits || [];
          if (key === 'change' && m.description) {
            if (m.description.toLowerCase().startsWith('visit ')) {
              const specific = targetVisits.filter(
                (v) => v.visitName && m.description.includes(v.visitName)
              );
              if (specific.length > 0) targetVisits = specific;
            }
          }

          targetVisits.forEach((v) => {
            if (v.visitId && v.visitId.includes('-')) {
              const tId = v.visitId.substring(v.visitId.indexOf('-') + 1);
              if (!overrideTaskMap.has(tId)) overrideTaskMap.set(tId, []);
              overrideTaskMap.get(tId).push({
                action: key === 'manual' ? 'manual' : (m.action || key).toLowerCase(),
                version: ver,
              });
            }
          });
        });
      });
    });

    tasksData.forEach((task) => {
      const taskId = String(task._id || task.id);
      const isUnassigned =
        (!task.eta || task.eta === '-' || task.eta === '') &&
        (!task.etd || task.etd === '-' || task.etd === '') &&
        (!task.routePlannedOrder ||
          task.routePlannedOrder === '-' ||
          task.routePlannedOrder === '');

      if (overrideTaskMap.has(taskId)) {
        task.isManual = true;
        task.manualType = 'Forced Assign';
        const sortedActions = overrideTaskMap
          .get(taskId)
          .sort((a, b) => a.version - b.version)
          .map((item) => item.action);
        task.action = [...new Set(sortedActions)].join(', ');
      } else if (isUnassigned) {
        task.isManual = true;
        task.manualType = 'Manual Assign';
        task.action = '-';
      } else {
        task.isManual = false;
        task.manualType = null;
        task.action = null;
      }
    });

    const groupedData = groupTasksByDriver(tasksData);

    const wb = generateTaskManualDetailWorkbook(
      groupedData,
      taskManualHeaders,
      taskManualKeyMapping,
      routingResults
    );

    const dateStr = formatDateUniversal(date, 'DD.MM.YYYY');
    const fileName = `${t('report.custom.task_manual')} - ${dateStr} - ${locationName}.xlsx`;
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    generatedFiles.push({ fileName, wb, wbout });
  }

  return generatedFiles;
};

export const processServiceLevelReport = async ({
  datesToProcess,
  locationName,
  t,
  driverData,
}) => {
  const generatedFiles = [];

  for (const date of datesToProcess) {
    const { timeFromUtc, timeToUtc, locTimeFrom, locTimeTo, selectedDateString } = getReportDates(
      date,
      date
    );

    const [tasks, locHistories] = await Promise.all([
      getTasks({
        status: 'DONE,ONGOING',
        timeFrom: timeFromUtc,
        timeTo: timeToUtc,
      }),
      getLocationHistories({
        dateFrom: selectedDateString,
        dateTo: selectedDateString,
        timeFrom: locTimeFrom,
        timeTo: locTimeTo,
      }),
    ]);

    const locationHistoryByDate = (locHistories || []).filter((item) => {
      return item.startTime?.startsWith(selectedDateString);
    });

    const { timeDataObjects } = convertLocationHistories(locationHistoryByDate, driverData);
    const filteredTimeData = timeDataObjects.filter(
      (item) => !isEmpty(item.startTimeFmt) && !isEmpty(item.finishTimeFmt)
    );
    const tasksData = normalizeTasksData(tasks);
    const sheetData = [serviceLevelHeaders];
    const parsedRows = tasksData.map((task) => {
      const { name, id, location, invoiceNumber } = parseCustomerString(task.customerOrder);
      const arrivalSource = task.klikJikaSudahSampai || task.klikJikaAndaSudahSampai;
      const doneSource = task.page3DoneTime || task.doneTime;
      const flow = task.flow || '-';
      const statusDelivery = task.statusDelivery || '-';
      const created = task.createdTime || '-';
      const arrived = arrivalSource || '-';
      const assigned = task.assignedTime || '-';
      const completed = doneSource || '-';
      let serviceLevel = '-';
      let startTrip = null;
      const assigneeEmail = Array.isArray(task.assignee)
        ? task.assignee[0] || ''
        : task.assignee || '';
      const driverName = task.driverName || assigneeEmail || '-';
      const licenseNumber = task.basePlat || '-';

      const createdTime = new Date(created);
      const doneDate = new Date(doneSource);

      if (!isNaN(createdTime) && !isNaN(doneDate)) {
        const diffMs = doneDate.getTime() - createdTime.getTime();

        if (diffMs > 0) {
          serviceLevel = `${Math.ceil(diffMs / 86400000)}`;
        }
      }

      if (assigneeEmail) {
        const driverHistory = filteredTimeData.find((item) => item.email === assigneeEmail);
        const timeDriver =
          driverHistory?.startDate && driverHistory?.startTimeFmt
            ? `${driverHistory.startDate} ${driverHistory.startTimeFmt}`
            : '-';

        startTrip = !isEmpty(timeDriver) ? timeDriver : '-';
      }

      return {
        row: [
          flow,
          driverName,
          licenseNumber,
          name || '-',
          id || '-',
          location || '-',
          invoiceNumber || '-',
          statusDelivery || '-',
          created,
          assigned,
          startTrip,
          arrived,
          completed,
          serviceLevel,
        ],
        driverName,
        rawCompleted: task.doneTime || '',
      };
    });

    parsedRows.sort((a, b) => {
      const driverCmp = a.driverName.localeCompare(b.driverName);
      if (driverCmp !== 0) return driverCmp;
      return a.rawCompleted.localeCompare(b.rawCompleted);
    });
    parsedRows.forEach((item) => sheetData.push(item.row));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 0; R <= range.e.r; ++R) {
      for (let C = 0; C <= range.e.c; ++C) {
        const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
        if (ws[cell_address]) {
          if (R === 0) {
            ws[cell_address].s = {
              font: { bold: true, color: { rgb: 'FFFFFF' } },
              fill: { patternType: 'solid', fgColor: { rgb: '0369A1' } },
              alignment: { horizontal: 'center', vertical: 'center' },
            };

            if (C === 13) {
              ws[cell_address].c = [{ a: 'System', t: 'Completed Time - Created Time' }];
            }
          } else {
            const isLeft = C === 1 || C === 3 || C === 6;
            ws[cell_address].s = {
              alignment: { horizontal: isLeft ? 'left' : 'center', vertical: 'center' },
            };
          }
        }
      }
    }
    ws['!cols'] = serviceLevelHeaders.map(() => ({ wch: 20 }));

    XLSX.utils.book_append_sheet(wb, ws, 'Task Date');

    const dateStr = formatDateUniversal(date, 'DD.MM.YYYY');
    const fileName = `${t('report.custom.service_level')} - ${dateStr} - ${locationName}.xlsx`;
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    generatedFiles.push({ fileName, wb, wbout });
  }

  return generatedFiles;
};

export const processTripActivityReport = async ({ datesToProcess, locationName, t }) => {
  const generatedFiles = [];

  for (const date of datesToProcess) {
    const { locTimeFrom, locTimeTo, selectedDateString } = getReportDates(date, date);

    const locHistories = await getLocationHistories({
      date: selectedDateString,
      timeFrom: locTimeFrom,
      timeTo: locTimeTo,
    });

    const trips = locHistories || [];
    if (isEmpty(trips)) continue;
    const rows = [];

    trips.forEach((trip) => {
      const email = trip.email;
      const d = trips.find((driver) => driver.email === email);
      const start = formatDateUniversal(trip.startTime);
      if (isEmpty(d) || start !== selectedDateString) return;

      const startTrip = trip.startTime
        ? formatDateUniversal(trip.startTime, 'DD/MM/YYYY HH:mm:ss')
        : '-';
      const endTrip = trip.finish?.finishTime
        ? formatDateUniversal(trip.finish?.finishTime, 'DD/MM/YYYY HH:mm:ss')
        : '-';

      const assignedVehicleId = d.vehicleId || '-';
      const username = d.driverName || '-';
      const assignedVehicle = d.basePlat || '-';

      const startTripCoordinate = trip.lat && trip.lon ? `${trip.lat}, ${trip.lon}` : '-';
      const endTripCoordinate =
        trip.finish?.lat && trip.finish?.lon ? `${trip.finish.lat}, ${trip.finish.lon}` : '-';

      const totalDistance = trip.finish?.totalDistance ?? '-';
      const totalDuration = trip.durationHour ?? '-';

      rows.push({
        rowData: [
          assignedVehicleId,
          username,
          assignedVehicle,
          startTrip,
          startTripCoordinate,
          endTrip,
          endTripCoordinate,
          totalDistance,
          totalDuration,
        ],
        username,
        rawStartTime: trip.startTime || '',
      });
    });

    rows.sort((a, b) => {
      const userCmp = String(a.username).localeCompare(String(b.username));
      if (userCmp !== 0) return userCmp;
      return String(a.rawStartTime).localeCompare(String(b.rawStartTime));
    });

    const sheetData = [tripActivityHeaders, ...rows.map((r) => r.rowData)];

    if (sheetData.length === 1) continue;

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let C = 0; C <= range.e.c; ++C) {
      const cell_address = XLSX.utils.encode_cell({ r: 0, c: C });
      if (ws[cell_address]) {
        ws[cell_address].s = {
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          fill: { patternType: 'solid', fgColor: { rgb: '0369A1' } },
          alignment: { horizontal: 'center', vertical: 'center' },
        };
      }
    }
    ws['!cols'] = tripActivityHeaders.map(() => ({ wch: 20 }));

    XLSX.utils.book_append_sheet(wb, ws, 'Trip Activity');

    const dateStr = formatDateUniversal(date, 'DD.MM.YYYY');
    const fileName = `${t('report.custom.trip_activity')} - ${dateStr} - ${locationName}.xlsx`;
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    generatedFiles.push({ fileName, wb, wbout });
  }

  return generatedFiles;
};
