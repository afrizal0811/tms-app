'use client';

import Tooltip from '@/components/Tooltip';
import { toastError, toastSuccess } from '@/lib/toast';
import * as XLSX from 'xlsx-js-style';
import {
  formatDeliveryStatus,
  deliveryTooltip,
  normalizeDeliveryStatus,
} from '../../../lib/deliveryStatus';
import {
  formatDateUniversal,
  getBasePlate,
  getStorageType,
  isEmpty,
  normalizeEmail,
  parseCustomerString,
  sortRows,
} from '../../../lib/utils';

export function routingActual({ tasks, drivers, dateStr }) {
  if (!tasks || !drivers) return [];

  const emailPlatMap = new Map();
  const emailFallbackMap = new Map();
  for (const d of drivers) {
    const normEmail = normalizeEmail(d.email);
    const bPlat = d.basePlat || '';
    if (normEmail) {
      emailFallbackMap.set(normEmail, { plat: d.plat || null, name: d.name });
      if (bPlat)
        emailPlatMap.set(`${normEmail}_${bPlat.toLowerCase()}`, { plat: d.plat, name: d.name });
    }
  }

  const processed = [];

  for (const t of tasks) {
    const flow = t.flow || '';

    const taskPlat = t.basePlat || t.vehicleName || t.vehicleId || t.plat || t.licensePlate || null;

    let emailStr = null;
    if (t.assignee) {
      emailStr = typeof t.assignee === 'string' ? t.assignee.split(',')[0].trim() : t.assignee;
    }
    const driverEmail = normalizeEmail(emailStr);
    const taskBasePlat = taskPlat || '';

    let driverInfo = null;
    if (driverEmail && taskBasePlat) {
      driverInfo = emailPlatMap.get(`${driverEmail}_${taskBasePlat.toLowerCase()}`);
    }
    if (!driverInfo && driverEmail) {
      driverInfo = emailFallbackMap.get(driverEmail);
    }

    const driverName = driverInfo?.name || t.assignedTo?.name || driverEmail || 'N/A';
    const finalPlat = taskPlat || driverInfo?.plat || '-';
    const basePlat = getBasePlate(finalPlat) || finalPlat;
    const groupKey = `${driverName}_${basePlat}`;

    let statusLabel = null;
    if (t.statusDelivery) {
      statusLabel = normalizeDeliveryStatus(t.statusDelivery);
    }

    if (flow === 'Pickup') {
      statusLabel = t.status
        ? normalizeDeliveryStatus(t.status) || t.status.toUpperCase()
        : statusLabel;
    }
    if (flow === 'Pickup' && (statusLabel === 'DONE' || statusLabel === 'SUCCESS'))
      statusLabel = 'SUCCESS';
    if (t.status !== 'ONGOING' && flow !== 'Pickup') statusLabel = statusLabel || '-';

    const rawCustStr = t.customerOrder || t.customerName || '';
    const {
      name: cName,
      id: cId,
      location: cLoc,
      fullCustomerName,
    } = parseCustomerString(rawCustStr);

    const isGrOrPickup = flow.toUpperCase().includes('GR') || flow.toUpperCase().includes('PICKUP');
    const actualArr = isGrOrPickup
      ? t.page1DoneTime
      : t.klikJikaSudahSampai || t.klikJikaAndaSudahSampai;
    const actualDep = isGrOrPickup ? t.page1DoneTime : t.page3DoneTime;

    const arrDate = new Date(actualArr);
    const depDate = new Date(actualDep);

    const actualArrTimestamp = isNaN(arrDate.getTime()) ? null : arrDate.getTime();
    const actualDepTimestamp = isNaN(depDate.getTime()) ? null : depDate.getTime();

    const actualArrVal = actualArrTimestamp ? formatDateUniversal(arrDate, 'HH:mm') : '-';
    const actualDepVal = actualDepTimestamp ? formatDateUniversal(depDate, 'HH:mm') : '-';

    const openTimeVal =
      t.openTime || formatDateUniversal(`${dateStr} ${t.openTime}`, 'HH:mm') || '-';
    const closeTimeVal =
      t.closeTime || formatDateUniversal(`${dateStr} ${t.closeTime}`, 'HH:mm') || '-';

    let hoursStatus = null;
    if (actualArrVal !== '-' && openTimeVal !== '-' && closeTimeVal !== '-') {
      const isInside =
        openTimeVal > closeTimeVal
          ? actualArrVal >= openTimeVal || actualArrVal <= closeTimeVal
          : actualArrVal >= openTimeVal && actualArrVal <= closeTimeVal;
      hoursStatus = isInside ? 'yes' : actualArrVal < openTimeVal ? 'early' : 'no';
    }

    let actualVisitMinutes = '-';
    if (actualArrTimestamp && actualDepTimestamp) {
      const arrH = arrDate.getHours();
      const arrM = arrDate.getMinutes();
      const depH = depDate.getHours();
      const depM = depDate.getMinutes();
      actualVisitMinutes = depH * 60 + depM - (arrH * 60 + arrM);
      if (actualVisitMinutes < 0) actualVisitMinutes += 24 * 60;
    }

    processed.push({
      groupKey,
      basePlat,
      driver: driverName,
      driverEmail,
      plat: finalPlat,
      actualArrivalTimestamp: actualArrTimestamp,
      actualDepartureTimestamp: actualDepTimestamp,
      roSequence: t.routePlannedOrder || 0,
      statusLabel,
      isModified: Boolean(
        t.isModified ||
          (t.pendingDetail?.new_status &&
            t.originalStatus &&
            normalizeDeliveryStatus(t.pendingDetail.new_status) !==
              normalizeDeliveryStatus(t.originalStatus))
      ),
      originalStatus: t.originalStatus || null,
      pendingDetail: t.pendingDetail || null,
      flow,
      customerName: cName || fullCustomerName || t.customerName,
      originalCustomerString: rawCustStr,
      customerId: cId,
      locationId: cLoc,
      openTime: openTimeVal,
      closeTime: closeTimeVal,
      eta: t.eta || '-',
      etd: t.etd || '-',
      actualArrival: actualArrVal,
      actualDeparture: actualDepVal,
      visitTime: t.visitTime || '-',
      actualVisitTime: actualVisitMinutes,
      isManualAssign: !t.routePlannedOrder || t.routePlannedOrder === 0,
      isWithinHoursStatus: hoursStatus,
      reason: t.reason || '',
      orderId: t.orderId || '',
      temperature: getStorageType(driverName),
      _id: t._id,
      rawTask: t,
    });
  }

  processed.sort((a, b) => {
    if (a.groupKey !== b.groupKey) return a.groupKey.localeCompare(b.groupKey);
    return (a.actualDepartureTimestamp || Infinity) - (b.actualDepartureTimestamp || Infinity);
  });

  let currGroup = null,
    rank = 1;
  for (const row of processed) {
    if (row.groupKey !== currGroup) {
      currGroup = row.groupKey;
      rank = 1;
    }
    row.realSequence = row.actualDepartureTimestamp ? rank++ : null;
  }

  return processed;
}

export const processRoutingVsActualData = ({ tasks, results, drivers, searchQuery, date }) => {
  if (!tasks || !drivers) return [];

  const hubTimesMap = new Map();
  const hubTimesFallbackMap = new Map();

  const driverByPlat = new Map();
  tasks.forEach((t) => {
    const p = getBasePlate(t.basePlat || t.vehicleName || t.plat);
    if (p && t.driverName) driverByPlat.set(p, t.driverName);
  });
  drivers.forEach((d) => {
    const p = getBasePlate(d.plat);
    if (p && d.name) driverByPlat.set(p, d.name);
  });

  if (results) {
    for (const result of results) {
      const rawHubTimes = result.hubTimes || result.hub_times;
      if (rawHubTimes) {
        const hubTimesData =
          typeof rawHubTimes === 'string' ? JSON.parse(rawHubTimes) : rawHubTimes;
        const hubLongLat =
          result.hub?.lat && result.hub?.lng ? `${result.hub.lat},${result.hub.lng}` : null;

        for (const [rawPlat, trips] of Object.entries(hubTimesData || {})) {
          if (!Array.isArray(trips) || isEmpty(trips)) continue;
          const basePlat = getBasePlate(rawPlat) || rawPlat;
          const driverName = driverByPlat.get(basePlat) || 'N/A';

          const firstHub = trips.find((t) => t.category === 'first') || trips[0];
          const lastHub = trips.find((t) => t.category === 'last') || trips[trips.length - 1];
          const middleHubs = trips.filter(
            (t) => t.category === 'middle' || (t.category && t.category.startsWith('middle'))
          );

          const timesObj = {
            hubETD: firstHub?.etd
              ? formatDateUniversal(`${date} ${firstHub.etd}`, 'HH:mm') || '-'
              : '-',
            hubETA: lastHub?.eta
              ? formatDateUniversal(`${date} ${lastHub.eta}`, 'HH:mm') || '-'
              : '-',
            hubLongLat,
            middleHubs: middleHubs.map((h) => ({
              order: h.order,
              eta: h.eta ? formatDateUniversal(`${date} ${h.eta}`, 'HH:mm') || '-' : '-',
              etd: h.etd ? formatDateUniversal(`${date} ${h.etd}`, 'HH:mm') || '-' : '-',
              longlat: hubLongLat,
            })),
          };

          if (driverName !== 'N/A') {
            hubTimesFallbackMap.set(driverName, timesObj);
            hubTimesMap.set(`${driverName}_${basePlat}`, timesObj);
          }
          hubTimesMap.set(basePlat, timesObj);
          hubTimesMap.set(rawPlat, timesObj);
        }
      }
    }
  }

  const allTaskData = routingActual({ tasks, drivers, dateStr: date });
  const tasksByGroupMap = new Map();
  const groupStats = new Map();

  for (const task of allTaskData) {
    if (task.driver === 'N/A' || task.rawTask?.status === 'UNASSIGNED') continue;

    const gKey = task.groupKey;
    if (!tasksByGroupMap.has(gKey)) tasksByGroupMap.set(gKey, []);
    tasksByGroupMap.get(gKey).push(task);
    if (!groupStats.has(gKey)) {
      groupStats.set(gKey, { driver: task.driver, plat: task.plat, basePlat: task.basePlat });
    }
  }

  const driverList = Array.from(groupStats.entries()).map(([gKey, stats]) => ({
    gKey,
    plat: stats.plat,
    driver: stats.driver,
  }));
  const sortDrivers = sortRows(driverList, 'plat', 'driver');
  const finalRows = [];
  const query = (searchQuery || '').toLowerCase();

  for (const driverRow of sortDrivers) {
    const driverName = driverRow.driver;
    const driverPlat = driverRow.plat;
    const basePlat = getBasePlate(driverPlat);
    const driverTasks = tasksByGroupMap.get(driverRow.gKey) || [];
    const hubTimes = hubTimesMap.get(driverRow.gKey) ||
      hubTimesMap.get(basePlat) ||
      hubTimesMap.get(driverPlat) ||
      hubTimesFallbackMap.get(driverName) || { hubETD: '-', hubETA: '-', hubLongLat: null };

    const isDriverMatch =
      driverName.toLowerCase().includes(query) ||
      (driverPlat && driverPlat.toLowerCase().includes(query));

    const matchingTasks = driverTasks.filter((t) => {
      if (isDriverMatch) return true;
      return t.customerName && t.customerName.toLowerCase().includes(query);
    });

    if (isEmpty(matchingTasks) && !isDriverMatch) continue;

    finalRows.push({
      type: 'HUB_START',
      driver: driverName,
      plat: driverPlat,
      time: hubTimes.hubETD,
      longlat: hubTimes.hubLongLat,
      customerName: 'HUB',
    });

    const mHubs = (hubTimes.middleHubs || []).map((h) => ({
      type: 'HUB_MIDDLE',
      driver: driverName,
      plat: driverPlat,
      eta: h.eta,
      etd: h.etd,
      roSequence: h.order,
      customerName: 'HUB',
      longlat: h.longlat,
    }));

    const combined = [...matchingTasks, ...mHubs].sort(
      (a, b) => (a.roSequence || 0) - (b.roSequence || 0)
    );
    combined.forEach((t) => finalRows.push(t.type === 'HUB_MIDDLE' ? t : { type: 'TASK', ...t }));

    finalRows.push({
      type: 'HUB_END',
      driver: driverName,
      plat: driverPlat,
      time: hubTimes.hubETA,
      longlat: hubTimes.hubLongLat,
      customerName: 'HUB',
    });

    finalRows.push({ type: 'SPACER' });
  }

  return finalRows;
};

const getHoursStatusUI = (status, t) => {
  if (status === 'yes')
    return {
      text: t('common.button.btn_yes'),
      color: 'text-[#16A34A] dark:text-[#86EFAC]',
      hex: '16A34A',
    };
  if (status === 'early')
    return {
      text: t('dashboard.tab.routing_actual.early'),
      color: 'text-[#F59E0B] dark:text-[#FCD34D]',
      hex: 'F59E0B',
    };
  if (status === 'no')
    return {
      text: t('common.button.btn_no'),
      color: 'text-[#DC2626] dark:text-[#FCA5A5]',
      hex: 'DC2626',
    };
  return { text: '-', color: 'text-[#F1F5F9]', hex: null };
};

export const getRoutingActualColumns = (t) => {
  const colors = {
    green: { normal: 'DCFCE7', dark: '14532D' },
    orange: { normal: 'FFEDD5', dark: '7C2D12' },
    yellow: { normal: 'FEF3C7', dark: '713F12' },
    pink: { normal: 'FCE7F3', dark: '831843' },
    blue: { normal: 'DBEAFE', dark: '1E3A8A' },
  };

  const theme = {
    greenClass: 'bg-[#DCFCE7] dark:bg-[#14532D]/40',
    orangeClass: 'bg-[#FFEDD5] dark:bg-[#7C2D12]/40',
    yellowClass: 'bg-[#FEF3C7] dark:bg-[#713F12]/40',
    pinkClass: 'bg-[#FCE7F3] dark:bg-[#831843]/40',
    blueClass: 'bg-[#DBEAFE] dark:bg-[#1E3A8A]/40',
  };

  return [
    {
      id: 'flow',
      header: t('common.flow'),
      align: 'center',
      width: 'min-w-[90px]',
      excelWidth: 12,
      getValue: (row) => row.flow || '-',
    },
    {
      id: 'plat',
      header: t('common.license_number'),
      align: 'center',
      width: 'min-w-[120px]',
      excelWidth: 15,
      highlight: true,
      getValue: (row) => getBasePlate(row.plat) || '-',
    },
    {
      id: 'driver',
      header: t('common.driver'),
      align: 'left',
      width: 'min-w-[140px]',
      className: 'font-medium',
      excelWidth: 25,
      highlight: true,
      getValue: (row) => row.driver || '-',
    },
    {
      id: 'customerName',
      header: t('common.customer_name'),
      align: 'left',
      width: 'min-w-[260px]',
      excelWidth: 35,
      highlight: true,
      getValue: (row) => row.customerName || '-',
    },
    {
      id: 'statusLabel',
      header: 'Status',
      align: 'center',
      excelWidth: 15,
      getValue: (row) => formatDeliveryStatus(row.statusLabel, t),
      getUI: (val, row) => {
        const isModified =
          row.isModified ||
          Boolean(
            row.pendingDetail?.new_status &&
              row.originalStatus &&
              normalizeDeliveryStatus(row.pendingDetail.new_status) !==
                normalizeDeliveryStatus(row.originalStatus)
          );
        if (!isModified) return val;
        const tooltip = deliveryTooltip({
          originalStatus: row.originalStatus,
          pendingDetail: row.pendingDetail,
          formatStatus: (s) => formatDeliveryStatus(s, t),
          t,
        });
        return (
          <Tooltip tooltipContent={tooltip} zIndex="z-[100002]">
            <span
              className="inline-block cursor-help font-bold text-red-600 dark:text-red-400 border-b border-dotted border-red-500"
              onClick={(e) => {
                e.stopPropagation();
              }}
            >
              {val}
            </span>
          </Tooltip>
        );
      },
    },

    {
      id: 'openTime',
      header: t('common.open_time'),
      align: 'center',
      excelWidth: 12,
      className: theme.greenClass,
      excelBg: colors.green.normal,
      getValue: (row) => row.openTime || '-',
    },
    {
      id: 'closeTime',
      header: t('common.close_time'),
      align: 'center',
      excelWidth: 12,
      className: theme.greenClass,
      excelBg: colors.green.normal,
      getValue: (row) => row.closeTime || '-',
    },

    {
      id: 'eta',
      header: t('common.eta'),
      align: 'center',
      excelWidth: 12,
      className: theme.orangeClass,
      excelBg: colors.orange.normal,
      getValue: (row) => row.eta || '-',
    },
    {
      id: 'actualArrival',
      header: t('common.actual_arrival'),
      align: 'center',
      excelWidth: 15,
      className: theme.orangeClass,
      excelBg: colors.orange.normal,
      getValue: (row) => row.actualArrival || '-',
    },

    {
      id: 'etd',
      header: t('common.etd'),
      align: 'center',
      excelWidth: 12,
      className: theme.yellowClass,
      excelBg: colors.yellow.normal,
      getValue: (row) => row.etd || '-',
    },
    {
      id: 'actualDeparture',
      header: t('common.actual_departure'),
      align: 'center',
      excelWidth: 15,
      className: theme.yellowClass,
      excelBg: colors.yellow.normal,
      getValue: (row) => row.actualDeparture || '-',
    },

    {
      id: 'visitTime',
      header: t('common.plan_visit'),
      align: 'center',
      excelWidth: 12,
      className: theme.pinkClass,
      excelBg: colors.pink.normal,
      getValue: (row) =>
        row.visitTime !== undefined && row.visitTime !== null && row.visitTime !== ''
          ? row.visitTime
          : '-',
    },
    {
      id: 'actualVisitTime',
      header: t('common.actual_visit'),
      align: 'center',
      excelWidth: 12,
      className: theme.pinkClass,
      getCellClassName: (row) =>
        row.actualVisitTime === 0 ? 'bg-red-500 text-white dark:bg-red-600' : theme.pinkClass,
      excelBg: colors.pink.normal,
      getValue: (row) => (row.actualVisitTime !== '-' ? row.actualVisitTime : '-'),
    },

    {
      id: 'roSequence',
      header: t('common.plan_seq'),
      align: 'center',
      excelWidth: 10,
      className: `font-semibold ${theme.blueClass}`,
      excelBg: colors.blue.normal,
      getValue: (row) => (isEmpty(row.roSequence) || row.roSequence === 0 ? '-' : row.roSequence),
    },
    {
      id: 'realSequence',
      header: t('common.actual_seq'),
      align: 'center',
      excelWidth: 10,
      className: `font-semibold ${theme.blueClass}`,
      excelBg: colors.blue.normal,
      getValue: (row) =>
        isEmpty(row.realSequence) || row.realSequence === 0 ? '-' : row.realSequence,
    },

    {
      id: 'isMatch',
      header: t('dashboard.tab.routing_actual.is_match'),
      align: 'center',
      excelWidth: 15,
      tooltip: t('dashboard.tab.routing_actual.tooltip.exp_is_same'),
      getValue: (row) => {
        const isRoSeqNull = isEmpty(row.roSequence) || row.roSequence === 0;
        const isRealSeqNull = isEmpty(row.realSequence) || row.realSequence === 0;
        if (isRealSeqNull) return '-';
        const roSeq = isRoSeqNull ? '-' : row.roSequence;
        const realSeq = isRealSeqNull ? '-' : row.realSequence;
        return roSeq == realSeq ? t('common.status.match') : t('common.status.mismatch');
      },
      getUI: (val, row, t) => {
        if (val === '-') return <span className="font-bold text-[#F1F5F9]">-</span>;
        const isMatch = val === t('common.status.match');
        return (
          <span
            className={`font-bold ${isMatch ? 'text-[#16A34A] dark:text-[#86EFAC]' : 'text-[#DC2626] dark:text-[#FCA5A5]'}`}
          >
            {val}
          </span>
        );
      },
      getExcelColor: (val, t) => {
        if (val === t('common.status.match')) return '16A34A';
        if (val === t('common.status.mismatch')) return 'DC2626';
        return null;
      },
    },
    {
      id: 'isWithinHoursStatus',
      header: t('dashboard.tab.routing_actual.is_within_hours'),
      align: 'center',
      excelWidth: 20,
      tooltip: t('dashboard.tab.routing_actual.tooltip.exp_within_hours'),
      tooltipWidth: 'w-40',
      getValue: (row) => row.isWithinHoursStatus,
      getUI: (val, row, t) => {
        const statusUI = getHoursStatusUI(val, t);
        return <span className={`font-bold ${statusUI.color}`}>{statusUI.text}</span>;
      },
      getExcelValue: (val, row, t) => getHoursStatusUI(val, t).text,
      getExcelColor: (val, t) => {
        if (val === t('common.button.btn_yes')) return '16A34A';
        if (val === t('dashboard.tab.routing_actual.early')) return 'F59E0B';
        if (val === t('common.button.btn_no')) return 'DC2626';
        return null;
      },
    },
  ];
};

export function routingActualSheet(wb, data, t) {
  if (!data || !Array.isArray(data) || data.length === 0) return;

  const columns = getRoutingActualColumns(t);
  const headers = columns.map((c) => c.header);
  const sheetData = [headers];
  const manualAssignRows = new Set();
  let lastDriver = null;

  data.forEach((row) => {
    if (row.type === 'SPACER') return;

    const currentDriver = row.driver || 'Unknown';
    const isHubStart = row.type === 'HUB_START';
    const isHubEnd = row.type === 'HUB_END';
    const isHubMiddle = row.type === 'HUB_MIDDLE';
    const isHub = isHubStart || isHubEnd || isHubMiddle;

    if (lastDriver !== null && currentDriver !== lastDriver) {
      sheetData.push(Array(columns.length).fill(''));
    }
    lastDriver = currentDriver;

    if (!isHub && row.isManualAssign) {
      manualAssignRows.add(sheetData.length);
    }

    const rowData = columns.map((col) => {
      if (isHub) {
        if (col.id === 'customerName') return 'HUB';
        if (col.id === 'etd' && isHubStart) return row.time;
        if (col.id === 'eta' && isHubEnd) return row.time;
        if (isHubMiddle) {
          if (col.id === 'eta') return row.eta;
          if (col.id === 'etd') return row.etd;
          if (col.id === 'roSequence') return row.roSequence;
        }
        return null;
      }
      const rawVal = col.getValue(row);
      return col.getExcelValue ? col.getExcelValue(rawVal, row, t) : rawVal;
    });

    sheetData.push(rowData);
  });

  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  ws['!view'] = { state: 'frozen', ySplit: 1 };
  ws['!cols'] = columns.map((col, i) => {
    let maxLen = 0;
    sheetData.forEach((r) => {
      if (r[i] !== null && r[i] !== undefined) maxLen = Math.max(maxLen, String(r[i]).length);
    });
    return { wch: Math.min(Math.max(maxLen + 2, col.excelWidth || 10), 45) };
  });

  const STYLES = {
    center: { alignment: { horizontal: 'center', vertical: 'center' } },
    left: { alignment: { horizontal: 'left', vertical: 'center' } },
    hubRed: {
      alignment: { horizontal: 'center', vertical: 'center' },
      font: { bold: true, color: { rgb: 'FF0000' } },
    },
    hubViolet: {
      alignment: { horizontal: 'center', vertical: 'center' },
      font: { bold: true, color: { rgb: '7C3AED' } },
    },
  };

  for (let R = 0; R < sheetData.length; ++R) {
    const isHubRow = ws[XLSX.utils.encode_cell({ r: R, c: 3 })]?.v === 'HUB';
    const isMiddleHubRow = isHubRow && ws[XLSX.utils.encode_cell({ r: R, c: 13 })]?.v;

    for (let C = 0; C < columns.length; ++C) {
      const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };

      const col = columns[C];

      if (R === 0) {
        ws[cellRef].s = {
          ...STYLES.center,
          font: { bold: true, color: { rgb: '000000' } },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
          },
          fill: { fgColor: { rgb: col.excelBg || 'EFEFEF' }, patternType: 'solid' },
          ...(col.id === 'visitTime' ||
          col.id === 'actualVisitTime' ||
          col.id === 'roSequence' ||
          col.id === 'realSequence'
            ? { alignment: { horizontal: 'center', vertical: 'center', wrapText: true } }
            : {}),
        };
      } else if (isHubRow) {
        if (isMiddleHubRow) {
          if (
            col.id === 'customerName' ||
            col.id === 'eta' ||
            col.id === 'etd' ||
            col.id === 'roSequence'
          )
            ws[cellRef].s = STYLES.hubViolet;
          else ws[cellRef].s = { font: { color: { rgb: '7C3AED' } } };
        } else {
          if (col.id === 'customerName' || col.id === 'eta' || col.id === 'etd')
            ws[cellRef].s = STYLES.hubRed;
          else ws[cellRef].s = { font: { color: { rgb: 'FF0000' } } };
        }
      } else {
        const isSpacerRow =
          isEmpty(ws[XLSX.utils.encode_cell({ r: R, c: 0 })]?.v) &&
          isEmpty(ws[XLSX.utils.encode_cell({ r: R, c: 2 })]?.v);

        if (isSpacerRow) {
          ws[cellRef].s = { fill: { fgColor: { rgb: 'E2E8F0' }, patternType: 'solid' } };
          continue;
        }

        const baseStyle = col.align === 'center' ? STYLES.center : STYLES.left;
        const bgStyle = col.excelBg
          ? { fill: { fgColor: { rgb: col.excelBg }, patternType: 'solid' } }
          : {};

        if (manualAssignRows.has(R)) {
          ws[cellRef].s = {
            ...baseStyle,
            fill: { fgColor: { rgb: 'FEE2E2' }, patternType: 'solid' },
          };
        } else {
          ws[cellRef].s = { ...baseStyle, ...bgStyle };
        }

        if (col.id === 'actualVisitTime' && ws[cellRef].v === 0) {
          ws[cellRef].s = {
            ...baseStyle,
            fill: { fgColor: { rgb: 'EF4444' }, patternType: 'solid' },
            font: { bold: true, color: { rgb: 'FFFFFF' } },
          };
        }

        if (typeof ws[cellRef].v === 'number') ws[cellRef].t = 'n';

        if (col.getExcelColor) {
          const color = col.getExcelColor(ws[cellRef].v, t);
          if (color)
            ws[cellRef].s = { ...ws[cellRef].s, font: { bold: true, color: { rgb: color } } };
        }
      }
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, t('common.routing_vs_actual'));
}

export const downloadRoutingActualExcel = (data, t, selectedDate, hubLabel) => {
  try {
    const wb = XLSX.utils.book_new();
    routingActualSheet(wb, data, t);
    const dateStr = formatDateUniversal(selectedDate || new Date(), 'DD.MM.YYYY');
    const safeHubLabel = hubLabel ? ` - ${hubLabel}` : '';
    XLSX.writeFile(wb, `${t('common.routing_vs_actual')} - ${dateStr}${safeHubLabel}.xlsx`);
    toastSuccess(t('common.toast.success'));
  } catch (e) {
    toastError(t('common.toast.error', { err: e.message }), e);
  }
};
