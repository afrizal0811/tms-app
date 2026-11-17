import prisma from '@/lib/prisma';
import { formatListProductsForApi, getBasePlate, parseCustomerString } from '@/lib/utils';

export const TASK_TIMESTAMP_FIELDS = [
  'startTime',
  'endTime',
  'assignedTime',
  'doneTime',
  'createdTime',
  'updatedTime',
  'klikJikaSudahSampai',
  'page1DoneTime',
  'page2DoneTime',
  'page3DoneTime',
];

export function parseJsonField(val, fallback = {}) {
  if (typeof val === 'string') {
    try {
      return JSON.parse(val || '{}');
    } catch {
      return fallback;
    }
  }
  return val || fallback;
}

export async function buildDriversByEmailMap(hubId) {
  const driverWhere = hubId && hubId !== 'all' ? { hub_id: hubId } : {};
  const vehicleWhere = hubId && hubId !== 'all' ? { hubs_id: hubId } : {};

  const [driverList, vehicleList] = await Promise.all([
    prisma.driverList.findMany({ where: driverWhere }),
    prisma.vehicle.findMany({ where: vehicleWhere }),
  ]);

  const vehiclesByDriverId = new Map(
    vehicleList.filter((v) => v.driver_id).map((v) => [v.driver_id, v])
  );

  const driversByEmail = {};
  driverList.forEach((d) => {
    const email = (d.email || '').toLowerCase().trim();
    if (email && email !== '-') {
      const v = vehiclesByDriverId.get(d.id);
      const mappedTypeStr = v?.type ? (v.storage ? `${v.storage}-${v.type}` : v.type) : null;
      driversByEmail[email] = {
        id: d.id,
        name: d.name,
        email: d.email,
        plat: v?.license_number || null,
        basePlat: v ? getBasePlate(v.license_number) : null,
        type: mappedTypeStr,
        assignedVehicleId: v?.vehicle_id || null,
      };
    }
  });

  if (driverList.length === 0 && vehicleList.length === 0) {
    const oldWhere = hubId && hubId !== 'all' ? { hubs: { some: { id: hubId } } } : {};
    const [rawDrivers, mappingsDB] = await Promise.all([
      prisma.driver.findMany({ where: oldWhere }),
      prisma.vehicleMapping.findMany(),
    ]);

    const mappingsObj = mappingsDB.reduce((acc, curr) => {
      acc[curr.plat] = curr.mapped_type ?? curr.mappedType;
      acc[getBasePlate(curr.plat)] = curr.mapped_type ?? curr.mappedType;
      return acc;
    }, {});

    rawDrivers.forEach((d) => {
      const basePlat = getBasePlate(d.plat);
      const mappedType = mappingsObj[basePlat] || mappingsObj[d.plat];
      let mappedTypeStr = d.type;
      if (mappedType) {
        mappedTypeStr = d.storage ? `${d.storage}-${mappedType}` : mappedType;
      }

      const email = (d.email || '').toLowerCase().trim();
      if (email && email !== '-' && !driversByEmail[email]) {
        driversByEmail[email] = {
          ...d,
          type: mappedTypeStr,
          basePlat,
        };
      }
    });
  }

  return driversByEmail;
}

export function parseDurationToMinutes(val) {
  if (val == null || val === '') return null;
  if (typeof val === 'number') return Number.isNaN(val) ? null : val;
  const str = String(val).trim();
  if (!str) return null;
  if (str.includes(':')) {
    const parts = str.split(':').map(Number);
    if (parts.length >= 2 && !Number.isNaN(parts[0]) && !Number.isNaN(parts[1])) {
      return parts[0] * 60 + parts[1] + (parts[2] && !Number.isNaN(parts[2]) ? parts[2] / 60 : 0);
    }
  }
  const num = Number(str);
  return Number.isNaN(num) ? null : num;
}

export function formatPendingTask(p) {
  if (!p || typeof p !== 'object' || !p.id) return null;
  return {
    ...p,
    taskId: p.task_id,
    internalExternal: p.internal_external,
    detailReason: p.detail_reason,
    groupReason: p.group_reason,
    createdAt: p.created_at,
    updatedAt: p.updated_at,
  };
}

export function formatTask(dbTask) {
  const cust = parseCustomerString(dbTask.customer_order);
  const others = parseJsonField(dbTask.others);
  const listProduct = formatListProductsForApi(dbTask.list_products);
  const syncUser = parseJsonField(dbTask.sync_user);
  const taskUser = parseJsonField(dbTask.task_user);

  return {
    _id: dbTask.id,
    id: dbTask.id,
    hubId: dbTask.hub_id,
    routingResultId: dbTask.routing_result_id,
    customerOrder: dbTask.customer_order,
    customerName: cust.fullCustomerName || cust.name || '',
    orderId: cust.invoiceNumber || '',
    address: dbTask.address,
    flow: dbTask.flow,
    typeStorage: dbTask.type_storage,
    status: dbTask.status,
    statusDelivery: dbTask.status_delivery,
    reason: dbTask.reason,
    deliveryDate: dbTask.delivery_date,
    startTime: dbTask.start_time,
    endTime: dbTask.end_time,
    assignedTime: dbTask.assigned_time,
    doneTime: dbTask.done_time,
    arrivalTime: dbTask.arrival_time,
    arrival_time: dbTask.arrival_time,
    klikJikaSudahSampai: dbTask.arrival_time,
    assignedVehicleId: dbTask.assigned_vehicle_id,
    assignee: dbTask.assignee,
    driverName: dbTask.driver_name,
    plat: dbTask.plat,
    basePlat: getBasePlate(dbTask.plat),
    vehicleType: dbTask.vehicle_type,
    routePlannedOrder: dbTask.route_planned_order,
    weightKg: dbTask.weight_kg,
    volumeCbm: dbTask.volume_cbm,
    distance: dbTask.distance,
    travelDistance: dbTask.travel_distance,
    travelDuration: dbTask.travel_duration,
    eta: dbTask.eta,
    etd: dbTask.etd,
    expectedCoordinate: dbTask.expected_coordinate,
    longlat: dbTask.expected_coordinate,
    doneCoordinate: dbTask.done_coordinate,
    gpsSesuai: dbTask.new_coordinate ? 'TIDAK' : 'YA',
    newCoordinate: dbTask.new_coordinate,
    klikLokasiClient: dbTask.new_coordinate,
    listProduct,
    ...others,
    createdFrom: dbTask.created_from || null,
    syncUser,
    taskUser,
    sync_user: syncUser,
    task_user: taskUser,
    createdTime: taskUser.created_at || null,
    createdBy: taskUser.created_by || null,
    createdByName: taskUser.created_by || null,
    updatedBy: taskUser.updated_by || null,
    updatedByName: taskUser.updated_by || null,
    updatedTime: taskUser.updated_at || null,
    createdAt: taskUser.created_at || null,
    updatedAt: taskUser.updated_at || null,
  };
}
