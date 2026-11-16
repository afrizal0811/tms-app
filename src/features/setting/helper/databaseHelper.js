import { formatDateUniversal, formatDurationHHMMSS, parseCustomerString } from '@/lib/utils';
import JSZip from 'jszip';
import * as XLSX from 'xlsx-js-style';

const headerStyle = {
  font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
  fill: { fgColor: { rgb: '334155' } },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: {
    top: { style: 'thin', color: { rgb: 'CBD5E1' } },
    bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
    left: { style: 'thin', color: { rgb: 'CBD5E1' } },
    right: { style: 'thin', color: { rgb: 'CBD5E1' } },
  },
};

const cellCenter = {
  font: { sz: 9 },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: {
    top: { style: 'thin', color: { rgb: 'E2E8F0' } },
    bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
    left: { style: 'thin', color: { rgb: 'E2E8F0' } },
    right: { style: 'thin', color: { rgb: 'E2E8F0' } },
  },
};

const cellLeft = {
  font: { sz: 9 },
  alignment: { horizontal: 'left', vertical: 'center' },
  border: {
    top: { style: 'thin', color: { rgb: 'E2E8F0' } },
    bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
    left: { style: 'thin', color: { rgb: 'E2E8F0' } },
    right: { style: 'thin', color: { rgb: 'E2E8F0' } },
  },
};

function createStyledSheet(aoa, colWidths = [], centerIndices = []) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const numRows = aoa.length;
  const numCols = aoa[0] ? aoa[0].length : 0;

  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      if (!ws[cellRef]) continue;
      if (r === 0) {
        ws[cellRef].s = headerStyle;
      } else {
        ws[cellRef].s = centerIndices.includes(c) ? cellCenter : cellLeft;
      }
    }
  }

  if (colWidths.length > 0) {
    ws['!cols'] = colWidths.map((w) => ({ wch: w }));
  }

  return ws;
}

function formatAuditDate(val) {
  if (!val) return '';
  const formatted = formatDateUniversal(val, 'DD-MM-YYYY HH:mm:ss');
  return formatted === '-' ? '' : formatted;
}

export function buildDateWorkbook({
  dateStr,
  routing = [],
  tasks = [],
  dailyFleet = [],
  locations = [],
  kpi = null,
  translate,
}) {
  const routingHeader = [
    'No',
    translate('common.routing_id'),
    translate('common.routing_name'),
    translate('common.delivery_date'),
    translate('common.dropped_task'),
    translate('common.histories'),
    translate('common.created_by'),
    translate('common.created_time'),
    translate('common.updated_by'),
    translate('common.updated_time'),
  ];
  const routingRows = routing.map((r, i) => {
    const ru = r.route_user || r.routeUser;
    const su = r.sync_user || r.syncUser;
    const totalHistories =
      typeof r.histories === 'object' && r.histories !== null
        ? (r.histories.total ?? 0)
        : typeof r.histories === 'string'
          ? (() => {
              try {
                return JSON.parse(r.histories)?.total ?? 0;
              } catch {
                return 0;
              }
            })()
          : 0;
    return [
      i + 1,
      r.id || '',
      r.name || '',
      r.deliveryDate || dateStr,
      typeof r.dropped === 'object' && r.dropped !== null
        ? (r.dropped.total ?? 0)
        : (r.dropped ?? 0),
      totalHistories,
      ru?.created_by || r.createdBy || '',
      formatAuditDate(ru?.created_at || r.createdAt),
      su?.updated_by || su?.created_by || r.updatedBy || '',
      formatAuditDate(su?.updated_at || su?.created_at || r.updatedAt),
    ];
  });
  const wsRouting = createStyledSheet(
    [routingHeader, ...routingRows],
    [6, 25, 30, 18, 12, 10, 18, 20, 18, 20],
    [0, 3, 4, 5, 7, 9]
  );

  const taskHeader = [
    'No',
    translate('common.task_id'),
    translate('common.invoice_number'),
    translate('common.customer_name'),
    translate('common.address'),
    translate('common.status.delivery_status'),
    'Status',
    translate('common.reason'),
    translate('common.storage_type'),
    translate('common.flow'),
    translate('common.driver'),
    translate('common.license_number'),
    translate('common.vehicle_type'),
    translate('common.plan_seq'),
    `${translate('common.weight')} (kg)`,
    `${translate('common.volume')} (cbm)`,
    translate('common.distance'),
    translate('common.start_time'),
    translate('common.finish_time'),
    translate('common.actual_arrival'),
    'ETA',
    'ETD',
    translate('common.created_by'),
    translate('common.created_time'),
    translate('common.updated_by'),
    translate('common.updated_time'),
  ];
  const taskRows = tasks.map((t, i) => {
    const cust = parseCustomerString(t.customerOrder || t.customer_order);
    const tu = t.task_user || t.taskUser;
    return [
      i + 1,
      t.id || t._id || '',
      cust.invoiceNumber || t.orderId || '',
      cust.name || t.customerName || '',
      t.address || '',
      t.statusDelivery || t.status_delivery || '',
      t.status || '',
      t.reason || '',
      t.typeStorage || t.type_storage || '',
      t.flow || '',
      t.driverName || t.driver_name || '',
      t.plat || '',
      t.vehicleType || t.vehicle_type || '',
      t.routePlannedOrder ?? t.route_planned_order ?? '',
      t.weightKg ?? t.weight_kg ?? 0,
      t.volumeCbm ?? t.volume_cbm ?? 0,
      t.distance ?? 0,
      formatAuditDate(t.startTime || t.start_time),
      formatAuditDate(t.endTime || t.end_time),
      formatAuditDate(t.arrivalTime || t.arrival_time),
      t.eta || '',
      t.etd || '',
      tu?.created_by || t.createdBy || '',
      formatAuditDate(tu?.created_at || t.createdAt),
      tu?.updated_by || t.updatedBy || '',
      formatAuditDate(tu?.updated_at || t.updatedAt),
    ];
  });
  const wsTask = createStyledSheet(
    [taskHeader, ...taskRows],
    [
      6, 25, 18, 25, 35, 16, 12, 20, 14, 12, 20, 14, 14, 12, 12, 12, 12, 20, 20, 20, 10, 10, 18, 20,
      18, 20,
    ],
    [0, 5, 6, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 23, 25]
  );

  const vehicleDetailsHeader = [
    'No',
    translate('common.license_number'),
    translate('common.driver'),
    translate('common.vehicle_type'),
    translate('common.storage_type'),
    `${translate('common.actual')} ${translate('common.weight')} (kg)`,
    `${translate('common.weight_max')} (kg)`,
    `${translate('common.actual')} ${translate('common.volume')} (cbm)`,
    `${translate('common.volume_max')} (cbm)`,
    translate('common.distance'),
    translate('common.total_task'),
    translate('common.status.success'),
    translate('common.status.pending'),
    translate('common.status.partial'),
    translate('common.status.cancel'),
    translate('common.status.pending_gr'),
    translate('common.sync_by'),
    translate('common.updated_time'),
  ];
  const vehicleDetailsRows = dailyFleet.map((df, i) => {
    const td = df.task_data || {};
    const su = df.sync_user || df.syncUser;
    return [
      i + 1,
      df.plat || '',
      df.driver_name || '',
      df.vehicle_type || '',
      df.storage_type || '',
      df.act_weight ?? 0,
      df.max_weight ?? 0,
      df.act_volume ?? 0,
      df.max_volume ?? 0,
      df.total_distance ?? 0,
      td.total ?? 0,
      td.success ?? 0,
      td.pending ?? 0,
      td.partial ?? 0,
      td.cancel ?? 0,
      td.pending_gr ?? 0,
      su?.updated_by || su?.created_by || '',
      formatAuditDate(su?.updated_at || su?.created_at),
    ];
  });
  const wsVehicleDetails = createStyledSheet(
    [vehicleDetailsHeader, ...vehicleDetailsRows],
    [6, 16, 22, 16, 12, 18, 20, 20, 22, 12, 12, 10, 10, 10, 10, 12, 18, 20],
    [0, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17]
  );

  const locHeader = [
    'No',
    'Email',
    translate('common.driver'),
    translate('common.date'),
    translate('common.start_time'),
    translate('common.finish_time'),
    translate('common.duration'),
    translate('common.distance'),
    translate('common.start_location'),
    translate('common.finish_location'),
    translate('common.updated_by'),
    translate('common.updated_time'),
  ];
  const locRows = locations.map((l, i) => {
    const finish = l.finish || {};
    const finishTime = finish.finishTime || l.finish_time || l.finishTime || '';
    const totalDuration = finish.totalDuration ?? l.total_duration ?? l.totalDuration ?? '';
    const totalDistance = finish.totalDistance ?? l.total_distance ?? l.totalDistance;
    const startLat = l.startLat ?? l.start_lat ?? l.lat ?? '';
    const startLon = l.startLon ?? l.start_lon ?? l.lon ?? '';
    const finishLat = finish.lat ?? l.finish_lat ?? '';
    const finishLon = finish.lon ?? l.finish_lon ?? '';
    const startLocation = startLat !== '' && startLon !== '' ? `${startLat}, ${startLon}` : '';
    const finishLocation = finishLat !== '' && finishLon !== '' ? `${finishLat}, ${finishLon}` : '';

    return [
      i + 1,
      l.email || '',
      l.driver_name || l.driverName || '',
      l.date || dateStr,
      formatAuditDate(l.start_time || l.startTime),
      formatAuditDate(finishTime),
      formatDurationHHMMSS(totalDuration),
      totalDistance != null && totalDistance !== '' ? Number(totalDistance).toFixed(2) : '',
      startLocation,
      finishLocation,
      l.sync_user?.updated_by || l.syncUser?.updated_by || '',
      formatAuditDate(l.sync_user?.updated_at || l.syncUser?.updated_at),
    ];
  });
  const wsLoc = createStyledSheet(
    [locHeader, ...locRows],
    [6, 25, 22, 14, 20, 20, 14, 12, 24, 24, 18, 20],
    [0, 3, 4, 5, 6, 7, 11]
  );

  const kpiItem = Array.isArray(kpi) ? kpi[0] : kpi;
  const kpiHeader = ['Data', 'Total'];
  const kpiRows = [
    [`${translate('common.vehicle')} (Dry)`, kpiItem?.truck_dry ?? 0],
    [`${translate('common.vehicle')} (Frozen)`, kpiItem?.truck_frz ?? 0],
    [`${translate('common.vehicle')} Total`, kpiItem?.total_truck ?? 0],
    [`${translate('common.task')} (Dry)`, kpiItem?.task_dry ?? 0],
    [`${translate('common.task')} (Frozen)`, kpiItem?.task_frz ?? 0],
    [translate('common.total_task'), (kpiItem?.task_dry ?? 0) + (kpiItem?.task_frz ?? 0)],
    ['Pending (Dry)', kpiItem?.pending_dry ?? 0],
    ['Pending (Frozen)', kpiItem?.pending_frz ?? 0],
    [translate('common.routing'), kpiItem?.pending_routing ?? 0],
    [translate('common.driver'), kpiItem?.pending_driver ?? 0],
    ['Sales', kpiItem?.pending_sales ?? 0],
    [translate('common.customer'), kpiItem?.pending_customer ?? 0],
    [translate('common.others'), kpiItem?.pending_other ?? 0],
    [
      `${translate('common.estimate')} ${translate('common.duration')} (Dry)`,
      kpiItem?.est_hour_dry ?? 0,
    ],
    [
      `${translate('common.estimate')} ${translate('common.duration')} (Frozen)`,
      kpiItem?.est_hour_frz ?? 0,
    ],
    [
      `${translate('common.actual')} ${translate('common.duration')} (Dry)`,
      kpiItem?.act_hour_dry ?? 0,
    ],
    [
      `${translate('common.actual')} ${translate('common.duration')} (Frozen)`,
      kpiItem?.act_hour_frz ?? 0,
    ],
    [
      `${translate('common.estimate')} ${translate('common.distance')} (Dry)`,
      kpiItem?.est_dist_dry ?? 0,
    ],
    [
      `${translate('common.estimate')} ${translate('common.distance')} (Frozen)`,
      kpiItem?.est_dist_frz ?? 0,
    ],
    [
      `${translate('common.actual')} ${translate('common.distance')} (Dry)`,
      kpiItem?.act_dist_dry ?? 0,
    ],
    [
      `${translate('common.actual')} ${translate('common.distance')} (Frozen)`,
      kpiItem?.act_dist_frz ?? 0,
    ],
    [
      `${translate('common.weight')} (Dry)`,
      `${kpiItem?.total_kg_dry ?? 0} / ${kpiItem?.capacity_kg_dry ?? 0}`,
    ],
    [
      `${translate('common.volume')} (Dry)`,
      `${kpiItem?.total_cbm_dry ?? 0} / ${kpiItem?.capacity_cbm_dry ?? 0}`,
    ],
    [
      `${translate('common.weight')} (Frozen)`,
      `${kpiItem?.total_kg_frz ?? 0} / ${kpiItem?.capacity_kg_frz ?? 0}`,
    ],
    [
      `${translate('common.volume')} (Frozen)`,
      `${kpiItem?.total_cbm_frz ?? 0} / ${kpiItem?.capacity_cbm_frz ?? 0}`,
    ],
  ];
  const wsKpi = createStyledSheet([kpiHeader, ...kpiRows], [35, 20], [1]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsRouting, translate('common.routing'));
  XLSX.utils.book_append_sheet(wb, wsTask, translate('common.task'));
  XLSX.utils.book_append_sheet(wb, wsVehicleDetails, 'Vehicle Details');
  XLSX.utils.book_append_sheet(wb, wsLoc, translate('common.driver_location'));
  XLSX.utils.book_append_sheet(wb, wsKpi, 'KPI Data');

  return wb;
}

export function downloadSingleDateExcel({ dateStr, data, hubAcronym = '', translate }) {
  const wb = buildDateWorkbook({ dateStr, ...data, translate });
  const dFormatted = formatDateUniversal(dateStr, 'DD.MM.YYYY');
  const fileName = hubAcronym
    ? `Database - ${dFormatted} - ${hubAcronym}.xlsx`
    : `Database - ${dFormatted}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

export async function downloadDateRangeZip({
  datesList,
  fetchDateData,
  hubAcronym = '',
  translate,
}) {
  const zip = new JSZip();

  for (const dateStr of datesList) {
    const data = await fetchDateData(dateStr);
    const wb = buildDateWorkbook({ dateStr, ...data, translate });
    const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const dFormatted = formatDateUniversal(dateStr, 'DD.MM.YYYY');
    zip.file(`${dFormatted}.xlsx`, buffer);
  }

  const startD = formatDateUniversal(datesList[0], 'DD.MM.YYYY');
  const endD = formatDateUniversal(datesList[datesList.length - 1], 'DD.MM.YYYY');
  const zipName = hubAcronym
    ? `Database (Bulk) - (${startD} - ${endD}) - ${hubAcronym}.zip`
    : `Database (Bulk) - (${startD} - ${endD}).zip`;

  const blob = await zip.generateAsync({ type: 'blob' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = zipName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);
}
