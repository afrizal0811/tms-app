import { getVehicles, getVehicleExternals } from '@/lib/api/mileapp';
import {
  formatLongDate,
  formatDateUniversal,
  getBasePlate,
  getDeliveryDateFromRouting,
  getStorageType,
  isEmpty,
  isPastDate,
} from '@/lib/utils';
import * as XLSX from 'xlsx-js-style';
import { BASE_STYLES, BORDERS, FILL_STYLES, FONT_STYLES, HEADER_STYLES } from './reportStyles';

const normalizePlate = (plate) => (plate || '').replace(/\s+/g, '').toLowerCase();

function extractTaskEmail(task) {
  if (task.assignee) {
    return typeof task.assignee === 'string' ? task.assignee.split(',')[0].trim() : task.assignee;
  }
  if (task.assignedTo?.email) return task.assignedTo.email;
  if (task.doneBy) return task.doneBy;
  return null;
}

function isTaskVehicleFrozen(driverInfo, taskTypeStorage) {
  if (driverInfo) {
    const driverDataArr = [driverInfo.name, driverInfo.plat, driverInfo.rawType];
    let isFrozen = getStorageType(driverDataArr) === 'Frozen';
    const typeUpper = (driverInfo.rawType || '').toUpperCase();
    if (typeUpper.includes('FROZEN') || typeUpper.includes('FRZ')) isFrozen = true;
    if (typeUpper.includes('DRY')) isFrozen = false;
    return isFrozen;
  }

  let isFrozen = getStorageType(taskTypeStorage || '') === 'Frozen';
  const tStorage = (taskTypeStorage || '').toUpperCase();
  if (tStorage.includes('FROZEN') || tStorage.includes('FRZ')) isFrozen = true;
  if (tStorage.includes('DRY')) isFrozen = false;
  return isFrozen;
}

function extractVehicleType(dType, vehicleTypes) {
  if (!dType) return 'Lainnya';
  const cleanUpper = dType.replace(/["'\\]/g, '').trim().toUpperCase();
  if (vehicleTypes.includes(cleanUpper)) return cleanUpper;

  const withoutStorage = cleanUpper.replace(/^(DRY|FROZEN|FRZ)-/, '');
  if (vehicleTypes.includes(withoutStorage)) return withoutStorage;

  return withoutStorage.replace(/-\d+$/, '');
}

function buildUsageStat(tms, manual, vehicleCount, workingDays) {
  const V = vehicleCount || 0;
  const TV = V * workingDays;
  const TVU = tms + manual;
  const PctTVU = TV > 0 ? TVU / TV : 0;
  const PctTMS = TV > 0 ? tms / TV : 0;
  const PctManual = TV > 0 ? manual / TV : 0;
  const VU = Math.ceil(PctTVU * V);
  const IV = Math.max(0, V - VU);
  const PctIV = V > 0 ? IV / V : 0;
  return { TMS: tms, Manual: manual, TVU, TV, PctTVU, PctTMS, PctManual, PctIV, V, VU, IV };
}

export function calculateUsageSummary(dateMap, dateKeys, hubMasterData, vehicleTypes) {
  const summary = { Dry: { types: {}, total: {} }, Frozen: { types: {}, total: {} }, OTV: {} };

  const workingDays = dateKeys.filter((d) => {
    if (d.isSunday) return false;
    return (dateMap[d.str]?.OTV || 0) > 0 || (dateMap[d.str]?.OTVManual || 0) > 0;
  }).length;

  const categories = ['Dry', 'Frozen'];

  categories.forEach((cat) => {
    let grpTMS = 0;
    let grpManual = 0;

    vehicleTypes.forEach((type) => {
      let totalTMS = 0;
      let totalManual = 0;
      dateKeys.forEach((d) => {
        totalTMS += dateMap[d.str][cat][type] || 0;
        totalManual += dateMap[d.str][`${cat}Manual`][type]?.count || 0;
      });

      const V_Type = hubMasterData?.[cat]?.[type] || 0;

      summary[cat].types[type] = buildUsageStat(totalTMS, totalManual, V_Type, workingDays);
      grpTMS += totalTMS;
      grpManual += totalManual;
    });

    let interbranchTMS = 0;
    let interbranchManual = 0;
    dateKeys.forEach((d) => {
      interbranchTMS += dateMap[d.str][cat]['Interbranch'] || 0;
      interbranchManual += dateMap[d.str][`${cat}Manual`]['Interbranch']?.count || 0;
    });

    const netTMS = grpTMS - interbranchTMS;
    const netManual = grpManual - interbranchManual;
    const V_Total = hubMasterData?.[cat]?.Total || 0;

    summary[cat].total = buildUsageStat(netTMS, netManual, V_Total, workingDays);
  });

  let otvTMS = 0;
  let otvManual = 0;
  dateKeys.forEach((d) => {
    otvTMS += dateMap[d.str].OTV || 0;
    otvManual += dateMap[d.str].OTVManual || 0;
  });

  const V_OTV = (hubMasterData?.Dry?.Total || 0) + (hubMasterData?.Frozen?.Total || 0);
  summary.OTV = buildUsageStat(otvTMS, otvManual, V_OTV, workingDays);

  return summary;
}

function findDriverInfoByPlate(masterDriversDB, canonicalPlate) {
  if (!canonicalPlate || !Array.isArray(masterDriversDB)) return undefined;
  const targetNorm = normalizePlate(canonicalPlate);
  const platMatch = masterDriversDB.find((d) => {
    if (!d) return false;
    const baseNorm = normalizePlate(d.basePlat);
    const platNorm = normalizePlate(d.plat);
    return (baseNorm && baseNorm === targetNorm) || (platNorm && platNorm === targetNorm);
  });
  if (!platMatch) return undefined;
  return {
    name: platMatch.name,
    storage: (platMatch.storage || '').toUpperCase(),
    rawType: platMatch._rawType || platMatch.type,
    plat: platMatch.basePlat || platMatch.plat,
  };
}

export async function calculateTruckUsageData(
  resultsData,
  startDateStr,
  endDateStr,
  taskData,
  masterTruckData
) {
  const taskPresence = {};
  if (taskData && Array.isArray(taskData)) {
    taskData.forEach((t) => {
      const d = t.startTime ? formatDateUniversal(t.startTime) : null;
      if (d) taskPresence[d] = true;
    });
  }
  const [allDriversDB, manualUsageDB] = await Promise.all([
    getVehicles(),
    getVehicleExternals({ startDate: startDateStr, endDate: endDateStr }),
  ]);

  const activeTypes = masterTruckData?.activeTypes || [];
  const hubMasterData = masterTruckData?.masterData;

  const groupedByEmail = {};
  (allDriversDB || []).forEach((d) => {
    const email = (d.email || '').toLowerCase().trim();
    if (email) {
      if (!groupedByEmail[email]) groupedByEmail[email] = [];
      groupedByEmail[email].push(d);
    }
  });

  const masterDriversDB = [];
  const conditionalPlates = new Set();

  (allDriversDB || []).forEach((d) => {
    const email = (d.email || '').toLowerCase().trim();
    let isConditional = false;

    const isSewa = (d.plat || '').toUpperCase().includes('SEWA');

    if (!isSewa && email && email !== '-' && groupedByEmail[email]) {
      const group = groupedByEmail[email];
      if (group.length > 1) {
        const spaceCount = (d.plat || '').trim().split(' ').length - 1;
        const minSpaces = Math.min(
          ...group.map((v) => (v.plat || '').trim().split(' ').length - 1)
        );

        if (spaceCount > minSpaces && spaceCount > 2) {
          isConditional = true;
        }
      }
    }

    if (isConditional) {
      if (d.plat) conditionalPlates.add(normalizePlate(d.plat));
    } else {
      masterDriversDB.push(d);
    }
  });

  const masterVehicleList = {
    Dry: { Gabungan: [] },
    Frozen: { Gabungan: [] },
    OTV: { Gabungan: [] },
  };

  activeTypes.forEach((type) => {
    masterVehicleList.Dry[type] = [];
    masterVehicleList.Frozen[type] = [];
  });

  const activeDrivers = masterDriversDB.filter((d) => {
    const plat = d.plat || '-';
    return !isEmpty(plat) && !plat.toUpperCase().includes('DEMO');
  });

  activeDrivers.forEach((d) => {
    const type = extractVehicleType(d.type, activeTypes);
    const firstTag = getStorageType(d.tags || d.vehicleTags || d.userTags);

    let isFrozen = firstTag === 'Frozen';
    const platUpper = (d.plat || '').toUpperCase();
    const nameUpper = (d.name || '').toUpperCase();
    const typeUpper = (d.type || '').toUpperCase();

    if (
      platUpper.includes('FRZ') ||
      nameUpper.includes('FRZ') ||
      typeUpper.includes('FRZ') ||
      typeUpper.includes('FROZEN')
    ) {
      isFrozen = true;
    }
    if (typeUpper.includes('DRY')) {
      isFrozen = false;
    }
    const storage = isFrozen ? 'Frozen' : 'Dry';
    const vInfo = { plate: d.plat, driver: d.name, type: type };

    if (!masterVehicleList[storage][type]) masterVehicleList[storage][type] = [];
    masterVehicleList[storage][type].push(vInfo);
    masterVehicleList[storage].Gabungan.push(vInfo);
    masterVehicleList.OTV.Gabungan.push(vInfo);
  });

  const dateMap = {};
  const dateKeys = [];
  const currentIterDate = new Date(startDateStr);
  const endDateObj = new Date(endDateStr);

  while (currentIterDate <= endDateObj) {
    const dateStr = formatDateUniversal(currentIterDate);
    const dayNum = currentIterDate.getDate();
    const isSunday = currentIterDate.getDay() === 0;
    dateKeys.push({ str: dateStr, day: dayNum, isSunday });

    dateMap[dateStr] = {
      Dry: { Interbranch: 0, Interbranch_details: [] },
      Frozen: { Interbranch: 0, Interbranch_details: [] },
      DryTotal: 0,
      FrozenTotal: 0,
      OTV: 0,
      DryManual: { Interbranch: {} },
      FrozenManual: { Interbranch: {} },
      DryTotalManual: 0,
      FrozenTotalManual: 0,
      OTVManual: 0,
      routingNames: new Set(),
    };

    activeTypes.forEach((type) => {
      dateMap[dateStr].Dry[type] = 0;
      dateMap[dateStr].Dry[`${type}_details`] = [];
      dateMap[dateStr].Frozen[type] = 0;
      dateMap[dateStr].Frozen[`${type}_details`] = [];
    });
    currentIterDate.setDate(currentIterDate.getDate() + 1);
  }

  if (Array.isArray(manualUsageDB)) {
    manualUsageDB.forEach((item) => {
      const dStr = item.date;
      if (dateMap[dStr]) {
        const st = item.storageType;
        const vt = item.vehicleType;
        dateMap[dStr][`${st}Manual`][vt] = {
          count: item.count,
          desc: item.description,
          id: item.id,
        };

        if (vt === 'Interbranch') {
          dateMap[dStr][`${st}TotalManual`] -= item.count;
          dateMap[dStr].OTVManual -= item.count;
        } else {
          dateMap[dStr][`${st}TotalManual`] += item.count;
          dateMap[dStr].OTVManual += item.count;
        }
      }
    });
  }

  if (resultsData && Array.isArray(resultsData)) {
    const driverMapHash = new Map();
    masterDriversDB.forEach((d) => {
      if (d.email) {
        driverMapHash.set(d.email.toLowerCase().trim(), {
          name: d.name,
          storage: d.storage.toUpperCase(),
          plat: d.basePlat,
          rawType: d.type,
        });
      }
    });

    const usedVehiclesPerDay = new Map();

    resultsData.forEach((res) => {
      if (res.dispatchStatus?.toLowerCase() !== 'done') return;
      const dateKey = res.deliveryDate || getDeliveryDateFromRouting(res.createdTime);
      if (!dateKey || !dateMap[dateKey]) return;
      if (res.name) {
        dateMap[dateKey].routingNames.add(res.name);
      }

      const plates = Array.from(
        new Set([
          ...Object.keys(res.totalDistance || {}),
          ...Object.keys(res.time || {}),
          ...Object.keys(res.hubTimes || {}),
        ])
      );

      plates.forEach((plat) => {
        if (!usedVehiclesPerDay.has(dateKey)) usedVehiclesPerDay.set(dateKey, new Map());
        const dailyVehicles = usedVehiclesPerDay.get(dateKey);

        const driverInfo =
          findDriverInfoByPlate(masterDriversDB, plat) || findDriverInfoByPlate(allDriversDB, plat);

        const canonicalPlate = plat || '-';

        if (!dailyVehicles.has(canonicalPlate)) {
          const rawSt = driverInfo ? getStorageType(driverInfo.name || driverInfo.rawType) : '-';
          dailyVehicles.set(canonicalPlate, {
            storageType: rawSt === 'Frozen' ? 'Frozen' : 'Dry',
            firstTag: driverInfo?.rawType || driverInfo?.storage,
            plate: plat || '-',
            driverName: driverInfo?.name || '-',
          });
        }
      });
    });

    if (taskData && Array.isArray(taskData)) {
      taskData.forEach((task) => {
        const dateKey = task.startTime ? formatDateUniversal(task.startTime) : null;
        if (!dateKey || !dateMap[dateKey]) return;

        if (!usedVehiclesPerDay.has(dateKey)) usedVehiclesPerDay.set(dateKey, new Map());
        const dailyVehicles = usedVehiclesPerDay.get(dateKey);

        const rawEmail = extractTaskEmail(task);

        const rawPlate = task.basePlat || task.vehicleName || task.plat || '';
        const rawCanonical = normalizePlate(rawPlate);
        if (conditionalPlates.has(rawCanonical)) return;

        const plat = rawPlate.replace(/\s*\([^)]*\)/g, '').trim();
        const canonicalPlate = normalizePlate(plat) || rawEmail || `unknown-${Math.random()}`;

        let driverInfo = null;
        if (rawPlate) {
          driverInfo = findDriverInfoByPlate(masterDriversDB, canonicalPlate);
        }
        if (!driverInfo && rawEmail) {
          driverInfo = driverMapHash.get(rawEmail.toLowerCase().trim());
        }

        const firstTag = driverInfo?.storage || task.typeStorage || '';
        const isFrozen = isTaskVehicleFrozen(driverInfo, task.typeStorage);

        if (!dailyVehicles.has(canonicalPlate)) {
          dailyVehicles.set(canonicalPlate, {
            storageType: isFrozen ? 'Frozen' : 'Dry',
            firstTag: driverInfo?.rawType || firstTag,
            plate: plat || driverInfo?.plat || '-',
            driverName: driverInfo?.name || rawEmail || '-',
          });
        }
      });
    }

    if (taskData && Array.isArray(taskData)) {
      taskData.forEach((task) => {
        const dateKey = task.startTime
          ? formatDateUniversal(task.startTime, 'YYYY-MM-DD HH:mm')
          : null;
        if (!dateKey || !dateMap[dateKey]) return;
        if (usedVehiclesPerDay.has(dateKey) && usedVehiclesPerDay.get(dateKey).size > 0) return;

        if (!usedVehiclesPerDay.has(dateKey)) usedVehiclesPerDay.set(dateKey, new Map());
        const dailyVehicles = usedVehiclesPerDay.get(dateKey);

        const rawEmail = extractTaskEmail(task);

        const emailClean = (rawEmail || '').toLowerCase().trim();
        const rawPlate = task.basePlat || task.vehicleName || task.plat || '';
        const rawCanonical = normalizePlate(rawPlate);
        if (conditionalPlates.has(rawCanonical)) return;

        const plat = rawPlate.replace(/\s*\([^)]*\)/g, '').trim();
        let driverInfo = null;

        if (rawPlate) {
          driverInfo = findDriverInfoByPlate(masterDriversDB, normalizePlate(plat));
        }
        if (!driverInfo && emailClean) {
          driverInfo = driverMapHash.get(emailClean);
        }

        let plateForCanonical = plat;
        if (!plateForCanonical && driverInfo && driverInfo.plat) {
          plateForCanonical = driverInfo.plat;
        }

        const canonicalPlate =
          normalizePlate(plateForCanonical) || emailClean || `unknown-task-${Math.random()}`;

        const firstTag = driverInfo?.storage || task.typeStorage || '';
        const isFrozen = isTaskVehicleFrozen(driverInfo, task.typeStorage);

        if (!dailyVehicles.has(canonicalPlate)) {
          dailyVehicles.set(canonicalPlate, {
            storageType: isFrozen ? 'Frozen' : 'Dry',
            firstTag: driverInfo?.rawType || firstTag,
            plate: plat || driverInfo?.plat || '-',
            driverName: driverInfo?.name || emailClean || '-',
          });
        } else {
          const existing = dailyVehicles.get(canonicalPlate);
          if (!existing.plate || existing.plate === '-') {
            existing.plate = plat || driverInfo?.plat || '-';
          }
          if (!existing.driverName || existing.driverName === '-') {
            existing.driverName = driverInfo?.name || emailClean || '-';
          }
        }
      });
    }

    usedVehiclesPerDay.forEach((dailyVehicles, dateKey) => {
      dailyVehicles.forEach((vh) => {
        const type = extractVehicleType(vh.firstTag, activeTypes);
        const storage = vh.storageType;
        if (dateMap[dateKey][storage] && dateMap[dateKey][storage][type] !== undefined) {
          const detailsList = dateMap[dateKey][storage][`${type}_details`];
          const vhBasePlate = normalizePlate(getBasePlate(vh.plate) || vh.plate);
          const vhDriver = (vh.driverName || '').toLowerCase().trim();

          const isExist = detailsList.some((d) => {
            const dBasePlate = normalizePlate(getBasePlate(d.plate) || d.plate);
            const dDriver = (d.driver || '').toLowerCase().trim();
            return (
              (vhBasePlate && dBasePlate && vhBasePlate === dBasePlate) ||
              (vhDriver && dDriver && vhDriver !== '-' && vhDriver === dDriver)
            );
          });

          if (!isExist) {
            dateMap[dateKey][storage][type]++;
            dateMap[dateKey][`${storage}Total`]++;
            dateMap[dateKey].OTV++;
            detailsList.push({
              plate: vh.plate,
              driver: vh.driverName,
              type: type,
            });
          }
        }
      });
    });
    const LOOKBACK_LIMIT = 3;
    dateKeys.forEach((dk) => {
      const currDateKey = dk.str;
      const currDm = dateMap[currDateKey];
      const currHasTasks = taskPresence[currDateKey];

      if (currHasTasks && currDm.OTV === 0) {
        for (let back = 1; back <= LOOKBACK_LIMIT; back++) {
          const d = new Date(currDateKey);
          d.setDate(d.getDate() - back);
          const prevDateKey = formatDateUniversal(d);
          const prevDm = dateMap[prevDateKey];
          const prevHasTasks = taskPresence[prevDateKey];

          if (prevDm && prevDm.OTV > 0 && !prevHasTasks) {
            activeTypes.forEach((type) => {
              currDm.Dry[type] = prevDm.Dry[type];
              currDm.Frozen[type] = prevDm.Frozen[type];
              currDm.Dry[`${type}_details`] = [...prevDm.Dry[`${type}_details`]];
              currDm.Frozen[`${type}_details`] = [...prevDm.Frozen[`${type}_details`]];

              prevDm.Dry[type] = 0;
              prevDm.Frozen[type] = 0;
              prevDm.Dry[`${type}_details`] = [];
              prevDm.Frozen[`${type}_details`] = [];
            });

            currDm.Dry[`Interbranch`] = prevDm.Dry[`Interbranch`];
            currDm.Frozen[`Interbranch`] = prevDm.Frozen[`Interbranch`];
            currDm.Dry[`Interbranch_details`] = [...prevDm.Dry[`Interbranch_details`]];
            currDm.Frozen[`Interbranch_details`] = [...prevDm.Frozen[`Interbranch_details`]];

            prevDm.Dry[`Interbranch`] = 0;
            prevDm.Frozen[`Interbranch`] = 0;
            prevDm.Dry[`Interbranch_details`] = [];
            prevDm.Frozen[`Interbranch_details`] = [];

            currDm.DryTotal = prevDm.DryTotal;
            currDm.FrozenTotal = prevDm.FrozenTotal;
            currDm.OTV = prevDm.OTV;

            prevDm.DryTotal = 0;
            prevDm.FrozenTotal = 0;
            prevDm.OTV = 0;

            prevDm.routingNames.forEach((name) => currDm.routingNames.add(name));
            prevDm.routingNames.clear();
            break;
          }
        }
      }
    });
  }

  dateKeys.forEach((dk) => {
    const dm = dateMap[dk.str];
    const isZero =
      dm.DryTotal === 0 &&
      dm.FrozenTotal === 0 &&
      dm.OTV === 0 &&
      dm.DryTotalManual === 0 &&
      dm.FrozenTotalManual === 0 &&
      dm.OTVManual === 0;

    dk.isPast = isPastDate(dk.str);
    dk.isDynamicHoliday = !dk.isSunday && dk.isPast && isZero;
    dk.routingNames = Array.from(dm.routingNames || []);
  });

  const summaryData = calculateUsageSummary(dateMap, dateKeys, hubMasterData, activeTypes);
  return { dateMap, dateKeys, activeTypes, hubMasterData, summaryData, masterVehicleList };
}

export async function generateTruckUsageSheet(
  wb,
  resultsData,
  startDateStr,
  endDateStr,
  translate,
  localeCode,
  taskData,
  masterTruckData,
  isDailyReport = false
) {
  const hidePercentageTable = Boolean(isDailyReport);
  const { dateMap, dateKeys, activeTypes, hubMasterData, summaryData } =
    await calculateTruckUsageData(resultsData, startDateStr, endDateStr, taskData, masterTruckData);
  const dryTypes = activeTypes.filter((type) => (hubMasterData?.Dry?.[type] || 0) > 0);
  const frozenTypes = activeTypes.filter((type) => (hubMasterData?.Frozen?.[type] || 0) > 0);
  const typesByCat = { Dry: dryTypes, Frozen: frozenTypes };

  const monthName = formatLongDate(startDateStr, localeCode).split(' ').slice(1).join(' ');
  const excelData = [];
  const merges = [];

  const getPctFill = (val) => {
    if (val > 1) return FILL_STYLES.alertRed;
    if (val >= 0.75) return { patternType: 'solid', fgColor: { rgb: 'B7E1CD' } };
    if (val >= 0.5) return { patternType: 'solid', fgColor: { rgb: 'F1C232' } };
    return { patternType: 'solid', fgColor: { rgb: 'F4CCCC' } };
  };

  let summaryCountEndRow = 0;
  let summaryPctStartRow = 0;
  let summaryPctEndRow = 0;

  const isSingleDay = startDateStr === endDateStr;

  if (!isSingleDay) {
    excelData.push([
      `${translate('summary.tabs.truck_usage.subtitle_1')} - ${monthName}`,
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
    ]);
    excelData.push([
      translate('common.vehicle_type'),
      'TMS',
      'Non TMS',
      'TVU',
      'TV',
      '% TVU',
      'V',
      'VU',
      'IV',
    ]);

    const addSummarySection = (cat, isPercentage = false) => {
      typesByCat[cat].forEach((type) => {
        const d = summaryData[cat].types[type];
        if (isPercentage) {
          excelData.push([type, d.PctTMS, d.PctManual, d.PctTVU]);
        } else {
          excelData.push([
            type,
            d.TMS || 0,
            d.Manual || 0,
            d.TVU || 0,
            d.TV || 0,
            d.PctTVU,
            null,
            null,
            null,
          ]);
        }
      });

      const t = summaryData[cat].total;
      if (isPercentage) {
        excelData.push([
          translate('summary.tabs.truck_usage.total_used'),
          t.PctTMS,
          t.PctManual,
          t.PctTVU,
        ]);
      } else {
        excelData.push([
          translate('summary.tabs.truck_usage.total_used'),
          t.TMS,
          t.Manual,
          t.TVU,
          t.TV,
          t.PctTVU,
          t.V,
          t.VU,
          t.IV,
        ]);
      }
    };

    addSummarySection('Dry', false);
    addSummarySection('Frozen', false);

    const otv = summaryData.OTV;
    excelData.push([
      'OTV',
      otv.TMS,
      otv.Manual,
      otv.TVU,
      otv.TV,
      otv.PctTVU,
      otv.V,
      otv.VU,
      otv.IV,
    ]);

    summaryCountEndRow = excelData.length;
    excelData.push([]);

    if (!hidePercentageTable) {
      summaryPctStartRow = excelData.length;

      excelData.push([
        `${translate('summary.tabs.truck_usage.subtitle_2')} - ${monthName}`,
        '',
        '',
        '',
      ]);
      excelData.push([translate('common.vehicle_type'), 'TMS', 'Non TMS', 'TVU']);

      addSummarySection('Dry', true);
      addSummarySection('Frozen', true);
      excelData.push(['OTV', otv.PctTMS, otv.PctManual, otv.PctTVU]);

      summaryPctEndRow = excelData.length;
      excelData.push([]);
    }
  }

  const table1StartRow = excelData.length;

  const buildTableData = (isPercentage, startRowIndex) => {
    const tableRows = [];
    const row1 = [isPercentage ? `${monthName} (%)` : monthName, 'Date', 'Total'];
    dateKeys.forEach((d) => row1.push(d.day, '', ''));
    tableRows.push(row1);

    const row2 = [translate('common.storage_type'), translate('common.vehicle_type'), ''];
    dateKeys.forEach(() => row2.push('TMS', 'Non TMS', 'TVU'));
    tableRows.push(row2);

    const rowMasterTotals = {};

    const createRow = (label1, label2, category, relativeRowIdx) => {
      const row = [label1, label2];
      let totalVal = null;
      if (category === 'Dry') totalVal = hubMasterData?.Dry?.[label2];
      else if (category === 'Frozen') totalVal = hubMasterData?.Frozen?.[label2];
      else if (category === 'DryTotal') totalVal = hubMasterData?.Dry?.Total;
      else if (category === 'FrozenTotal') totalVal = hubMasterData?.Frozen?.Total;
      else if (category === 'OTV')
        totalVal = (hubMasterData?.Dry?.Total || 0) + (hubMasterData?.Frozen?.Total || 0);

      rowMasterTotals[relativeRowIdx] = totalVal || 0;
      row.push(totalVal || null);

      dateKeys.forEach((d) => {
        let valRaw = 0;
        let manualRaw = 0;
        const isHoliday = d.isSunday || d.isDynamicHoliday;
        if (isHoliday) {
          if (relativeRowIdx === 2) {
            const text = d.isSunday
              ? translate('common.holiday_sunday')
              : translate('common.holiday');
            row.push(text, null, null);
          } else {
            row.push(null, null, null);
          }
          return;
        }
        if (category === 'Dry' || category === 'Frozen') {
          valRaw = dateMap[d.str][category][label2] || 0;
          manualRaw = dateMap[d.str][`${category}Manual`][label2]?.count || 0;
        } else if (category === 'DryTotal' || category === 'FrozenTotal') {
          valRaw = dateMap[d.str][category] || 0;
          manualRaw = dateMap[d.str][`${category}Manual`] || 0;
        } else if (category === 'OTV') {
          valRaw = dateMap[d.str].OTV || 0;
          manualRaw = dateMap[d.str].OTVManual || 0;
        }

        const tmsDisp = valRaw > 0 ? valRaw : null;
        const manualDisp = manualRaw > 0 ? manualRaw : null;
        let tvuDisp = null;
        const sumVal = valRaw + manualRaw;
        if (sumVal > 0) tvuDisp = sumVal;

        if (isPercentage) {
          if (totalVal > 0) {
            row.push(
              tmsDisp !== null ? tmsDisp / totalVal : null,
              manualDisp !== null ? manualDisp / totalVal : null,
              tvuDisp !== null ? tvuDisp / totalVal : null
            );
          } else {
            row.push(null, null, null);
          }
        } else {
          row.push(tmsDisp, manualDisp, tvuDisp);
        }
      });
      return row;
    };

    let rIdx = 2;
    dryTypes.forEach((type, idx) =>
      tableRows.push(createRow(idx === 0 ? 'Dry' : '', type, 'Dry', rIdx++))
    );
    tableRows.push(
      createRow(translate('summary.tabs.truck_usage.interbranch'), 'Interbranch', 'Dry', rIdx++)
    );
    tableRows.push(
      createRow(translate('summary.tabs.truck_usage.total_used'), '', 'DryTotal', rIdx++)
    );

    frozenTypes.forEach((type, idx) =>
      tableRows.push(createRow(idx === 0 ? 'Frozen' : '', type, 'Frozen', rIdx++))
    );
    tableRows.push(
      createRow(translate('summary.tabs.truck_usage.interbranch'), 'Interbranch', 'Frozen', rIdx++)
    );
    tableRows.push(
      createRow(translate('summary.tabs.truck_usage.total_used'), '', 'FrozenTotal', rIdx++)
    );
    tableRows.push(createRow('OTV', '', 'OTV', rIdx++));

    const H1 = startRowIndex;
    const H2 = startRowIndex + 1;
    let colIdx = 3;
    const totalRows = dryTypes.length + frozenTypes.length + 5;
    dateKeys.forEach((d) => {
      merges.push({ s: { r: H1, c: colIdx }, e: { r: H1, c: colIdx + 2 } });
      const isHoliday = d.isSunday || d.isDynamicHoliday;
      if (isHoliday) {
        merges.push({
          s: { r: H1 + 2, c: colIdx },
          e: { r: H1 + 2 + totalRows - 1, c: colIdx + 2 },
        });
      }
      colIdx += 3;
    });
    merges.push({ s: { r: H1, c: 2 }, e: { r: H2, c: 2 } });

    const dryStart = startRowIndex + 2;
    const dryInter = dryStart + dryTypes.length;
    const dryTot = dryInter + 1;
    const frzStart = dryTot + 1;
    const frzInter = frzStart + frozenTypes.length;
    const frzTot = frzInter + 1;
    const otvRow = frzTot + 1;

    if (dryTypes.length > 1) {
      merges.push({ s: { r: dryStart, c: 0 }, e: { r: dryInter - 1, c: 0 } });
    }
    if (frozenTypes.length > 1) {
      merges.push({ s: { r: frzStart, c: 0 }, e: { r: frzInter - 1, c: 0 } });
    }
    merges.push({ s: { r: dryInter, c: 0 }, e: { r: dryInter, c: 1 } });
    merges.push({ s: { r: dryTot, c: 0 }, e: { r: dryTot, c: 1 } });
    merges.push({ s: { r: frzInter, c: 0 }, e: { r: frzInter, c: 1 } });
    merges.push({ s: { r: frzTot, c: 0 }, e: { r: frzTot, c: 1 } });
    merges.push({ s: { r: otvRow, c: 0 }, e: { r: otvRow, c: 1 } });

    return { tableRows, rowMasterTotals };
  };

  const table1 = buildTableData(false, table1StartRow);
  excelData.push(...table1.tableRows);
  excelData.push([]);

  let table2StartRow = -1;
  let table2 = null;
  if (!hidePercentageTable) {
    table2StartRow = excelData.length;
    table2 = buildTableData(true, table2StartRow);
    excelData.push(...table2.tableRows);
    excelData.push([]);
  }
  excelData.push([translate('common.explanation')]);

  const legendTitleRow = excelData.length - 1;
  merges.push({ s: { r: legendTitleRow, c: 0 }, e: { r: legendTitleRow, c: 2 } });

  const legendItems = [
    { key: 'TMS', desc: translate('summary.tabs.truck_usage.tms') },
    { key: 'Non TMS', desc: translate('summary.tabs.truck_usage.non_tms') },
    { key: 'TVU', desc: translate('summary.tabs.truck_usage.tvu') },
    { key: 'TV', desc: translate('summary.tabs.truck_usage.tv') },
    { key: 'V', desc: translate('common.vehicle') },
    { key: 'VU', desc: translate('summary.tabs.truck_usage.vu') },
    { key: 'IV', desc: translate('summary.tabs.truck_usage.iv') },
    { key: 'OTV', desc: translate('summary.tabs.truck_usage.otv') },
  ];

  const legendItemStartRow = excelData.length;
  legendItems.forEach((item, idx) => {
    excelData.push([item.key, item.desc]);
    const currentRow = legendItemStartRow + idx;
    merges.push({ s: { r: currentRow, c: 1 }, e: { r: currentRow, c: 4 } });
  });

  const ws = XLSX.utils.aoa_to_sheet(excelData);

  if (!isSingleDay) {
    merges.push({ s: { r: 0, c: 1 }, e: { r: 0, c: 8 } });
    merges.push({ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } });
    if (!hidePercentageTable && summaryPctStartRow > 0) {
      merges.push({ s: { r: summaryPctStartRow, c: 1 }, e: { r: summaryPctStartRow, c: 3 } });
      merges.push({ s: { r: summaryPctStartRow, c: 0 }, e: { r: summaryPctStartRow + 1, c: 0 } });
    }
  }

  ws['!merges'] = merges;
  ws['!views'] = [{ state: 'frozen', xSplit: 3, ySplit: table1StartRow + 2 }];

  const range = XLSX.utils.decode_range(ws['!ref']);
  const tableHeight = 2 + dryTypes.length + 1 + 1 + frozenTypes.length + 1 + 1 + 1;
  const sumDryTot = 2 + dryTypes.length;
  const sumFrzStart = sumDryTot + 1;
  const sumFrzTot = sumFrzStart + frozenTypes.length;
  const sumOTV = sumFrzTot + 1;

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };
      const cell = ws[cellRef];

      if (R >= legendTitleRow) {
        if (R === legendTitleRow && C === 0) {
          cell.s = { font: { bold: true, underline: true }, alignment: { horizontal: 'left' } };
        } else if (R >= legendItemStartRow && C === 0) {
          cell.s = { alignment: { horizontal: 'left', wrapText: true, vertical: 'center' } };
        }
        continue;
      }

      const isSum1 = R < summaryCountEndRow;
      const isSum2 =
        !hidePercentageTable &&
        summaryPctStartRow > 0 &&
        R >= summaryPctStartRow &&
        R < summaryPctEndRow;

      if (isSum1 || isSum2) {
        const startRow = isSum1 ? 0 : summaryPctStartRow;
        const relR = R - startRow;
        const maxCol = isSum1 ? 8 : 3;

        if (C > maxCol) {
          cell.v = '';
          cell.s = {};
          continue;
        }

        if (relR === 0 || relR === 1) {
          cell.s = { ...HEADER_STYLES.main };
        } else {
          cell.s = { ...BASE_STYLES.center, border: BORDERS.thin };

          if (relR >= 2 && relR < sumDryTot) cell.s.fill = FILL_STYLES.dry;
          else if (relR === sumDryTot) {
            cell.s.fill = FILL_STYLES.dryTotal;
            cell.s.font = FONT_STYLES.bold;
          } else if (relR >= sumFrzStart && relR < sumFrzTot) cell.s.fill = FILL_STYLES.frozen;
          else if (relR === sumFrzTot) {
            cell.s.fill = FILL_STYLES.frozenTotal;
            cell.s.font = FONT_STYLES.bold;
          } else if (relR === sumOTV) {
            cell.s.fill = FILL_STYLES.otv;
            cell.s.font = FONT_STYLES.bold;
          }

          if (C === 0) cell.s.alignment = { horizontal: 'left', indent: 1 };

          if (isSum1 && C === 5) {
            cell.t = 'n';
            cell.s.numFmt = '0.00%';
          }
          if (isSum2 && C >= 1) {
            cell.t = 'n';
            cell.s.numFmt = '0.00%';
            if (cell.v > 0) cell.s.fill = getPctFill(cell.v);
          }
        }
        continue;
      }

      if (
        !isSingleDay &&
        (R === summaryCountEndRow ||
          (!hidePercentageTable && summaryPctEndRow > 0 && R === summaryPctEndRow))
      )
        continue;

      let isTable1 = false,
        isTable2 = false,
        relR = -1;

      if (R >= table1StartRow && R < table1StartRow + tableHeight) {
        isTable1 = true;
        relR = R - table1StartRow;
      } else if (table2StartRow !== -1 && R >= table2StartRow && R < table2StartRow + tableHeight) {
        isTable2 = true;
        relR = R - table2StartRow;
      }

      if (relR !== -1) {
        const currentMasterTotals = isTable1
          ? table1.rowMasterTotals
          : table2?.rowMasterTotals || {};
        const startDry = 2;
        const dryInter = startDry + dryTypes.length;
        const dryTot = dryInter + 1;
        const frzStart = dryTot + 1;
        const frzInter = frzStart + frozenTypes.length;
        const frzTot = frzInter + 1;
        const otvRow = frzTot + 1;

        cell.s = { ...BASE_STYLES.center };

        if (relR === 0 || relR === 1) {
          cell.s = { ...HEADER_STYLES.main };
          if (C === 3) cell.s.border.left = BORDERS.medium;
          if (C > 2 && (C - 3) % 3 === 2) cell.s.border.right = BORDERS.medium;
          if (C === 2) cell.s.border.right = BORDERS.medium;
        } else {
          let rowFill = null;
          if (relR >= startDry && relR <= dryInter) rowFill = FILL_STYLES.dry;
          else if (relR === dryTot) rowFill = FILL_STYLES.dryTotal;
          else if (relR >= frzStart && relR <= frzInter) rowFill = FILL_STYLES.frozen;
          else if (relR === frzTot) rowFill = FILL_STYLES.frozenTotal;
          else if (relR === otvRow) rowFill = FILL_STYLES.otv;

          if (isTable2 && C > 2) {
            if (typeof cell.v === 'number') {
              cell.t = 'n';
              cell.s.numFmt = '0%';
            }
          }

          if (C <= 2) {
            if (C <= 1) {
              const isMerged = [dryInter, dryTot, frzInter, frzTot, otvRow].includes(relR);
              if (isMerged && C === 0) cell.s.alignment = { ...BASE_STYLES.left.alignment };
              else if (C === 0) cell.s.alignment = { ...BASE_STYLES.center.alignment };
              else cell.s.alignment = { ...BASE_STYLES.left.alignment };

              if (rowFill) cell.s.fill = rowFill;
              cell.s.border = undefined;
            } else {
              cell.s.border = { left: BORDERS.thin, right: BORDERS.medium };
              if (rowFill) cell.s.fill = rowFill;
              cell.s.font = FONT_STYLES.bold;
            }
          } else {
            if (C === 3) cell.s.border = { left: BORDERS.medium };
            else cell.s.border = {};
            if ((C - 3) % 3 === 2) cell.s.border.right = BORDERS.medium;

            let isSundayCol = false;
            let isDynamicCol = false;
            let isTMSCol = false;
            let isTVUCol = false;

            if (C > 2) {
              const relativeIdx = (C - 3) % 3;
              if (relativeIdx === 0) isTMSCol = true;
              if (relativeIdx === 2) isTVUCol = true;
              const dateIndex = Math.floor((C - 3) / 3);
              if (dateKeys[dateIndex]) {
                if (dateKeys[dateIndex].isSunday) isSundayCol = true;
                if (dateKeys[dateIndex].isDynamicHoliday) isDynamicCol = true;
              }
            }

            let finalFill = rowFill;
            const isDetailRow = ![dryInter, dryTot, frzInter, frzTot, otvRow].includes(relR);
            const masterTotal = currentMasterTotals[relR] || 0;
            const val = cell.v;

            if (isTable1) {
              if (isDetailRow && masterTotal > 0) {
                const baseIdx = C - ((C - 3) % 3);
                const tmsVal = ws[XLSX.utils.encode_cell({ r: R, c: baseIdx })]?.v || 0;
                const manualVal = ws[XLSX.utils.encode_cell({ r: R, c: baseIdx + 1 })]?.v || 0;

                if (isTMSCol && val > masterTotal) finalFill = FILL_STYLES.alertRed;
                else if (!isTMSCol && !isTVUCol && tmsVal + manualVal > masterTotal)
                  finalFill = FILL_STYLES.alertRed;
                else if (isTVUCol && val > masterTotal) finalFill = FILL_STYLES.alertRed;
              }
              if ((isSundayCol || isDynamicCol) && finalFill !== FILL_STYLES.alertRed)
                finalFill = FILL_STYLES.red;
            } else {
              if (isSundayCol || isDynamicCol) {
                finalFill = FILL_STYLES.red;
                if (relR === 2 && isTMSCol) {
                  cell.t = 's';
                  cell.s.alignment = { horizontal: 'center', vertical: 'center' };
                  cell.s.font = { ...FONT_STYLES.bold, color: { rgb: '9C0006' } };
                }
              } else if ((isTMSCol || isTVUCol) && typeof val === 'number' && val > 0) {
                finalFill = getPctFill(val);
              }
            }

            if (finalFill) cell.s.fill = finalFill;
            if (
              (isTable2 || finalFill === FILL_STYLES.alertRed) &&
              finalFill === FILL_STYLES.alertRed
            ) {
              cell.s.font = FONT_STYLES.whiteBold;
            }
          }
          if ([dryInter, dryTot, frzInter, frzTot, otvRow].includes(relR)) {
            cell.s.font = FONT_STYLES.bold;
          }
        }
      }
    }
  }

  const cols = [{ wch: 15 }, { wch: 25 }, { wch: 8 }];
  for (let i = 0; i < dateKeys.length * 3; i++) cols.push({ wch: 8 });
  ws['!cols'] = cols;

  XLSX.utils.book_append_sheet(wb, ws, translate('summary.tabs.truck_usage.title'));
}
