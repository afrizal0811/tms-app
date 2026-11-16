import prisma from '@/lib/prisma';
import { DELIVERY_STATUS_KEYS, normalizeDeliveryStatus } from '@/lib/deliveryStatus';
import { getBasePlate } from '@/lib/utils';

export function categorizeDeliveryStatus(rawStatus) {
  const norm = normalizeDeliveryStatus(rawStatus);
  if (norm === DELIVERY_STATUS_KEYS.SUCCESS || norm === DELIVERY_STATUS_KEYS.DONE) {
    return 'success';
  }
  if (norm === DELIVERY_STATUS_KEYS.PARTIAL) {
    return 'partial';
  }
  if (norm === DELIVERY_STATUS_KEYS.CANCEL) {
    return 'cancel';
  }
  if (norm === DELIVERY_STATUS_KEYS.PENDING_GR) {
    return 'pending_gr';
  }
  if (norm === DELIVERY_STATUS_KEYS.PENDING) {
    return 'pending';
  }
  return null;
}

export async function syncDailyFleetForDate(hubId, deliveryDate, userName = 'system') {
  if (!deliveryDate) return 0;

  const taskWhere = { delivery_date: deliveryDate };
  if (hubId && hubId !== 'all') {
    taskWhere.hub_id = hubId;
  }

  const tasks = await prisma.task.findMany({
    where: taskWhere,
    select: {
      id: true,
      hub_id: true,
      delivery_date: true,
      routing_result_id: true,
      plat: true,
      driver_name: true,
      assignee: true,
      vehicle_type: true,
      type_storage: true,
      status: true,
      status_delivery: true,
      weight_kg: true,
      volume_cbm: true,
      distance: true,
      travel_distance: true,
    },
  });

  if (!tasks || tasks.length === 0) return 0;

  const taskIds = tasks.map((t) => t.id);
  const existingFleetWhere = { date: deliveryDate };
  if (hubId && hubId !== 'all') {
    existingFleetWhere.hub_id = hubId;
  }

  const [pendingDetails, vehicles, drivers, existingFleets] = await Promise.all([
    prisma.pendingDetail.findMany({
      where: { task_id: { in: taskIds } },
      select: { task_id: true, new_status: true },
    }),
    prisma.vehicle.findMany({
      select: {
        license_number: true,
        base_plate: true,
        max_weight: true,
        max_volume: true,
        type: true,
        storage: true,
      },
    }),
    prisma.driver.findMany({
      select: {
        email: true,
        name: true,
        plat: true,
        max_weight: true,
        max_volume: true,
        type: true,
        storage: true,
      },
    }),
    prisma.dailyFleet.findMany({
      where: existingFleetWhere,
      select: { hub_id: true, date: true, plat: true, sync_user: true },
    }),
  ]);

  const existingSyncMap = new Map(
    existingFleets.map((f) => [`${f.hub_id}_${f.date}_${f.plat}`, f.sync_user])
  );

  const pendingMap = new Map(pendingDetails.map((p) => [p.task_id, p.new_status]));

  const vehicleByPlate = new Map();
  vehicles.forEach((v) => {
    if (v.license_number) vehicleByPlate.set(v.license_number.toUpperCase().trim(), v);
    if (v.base_plate) vehicleByPlate.set(v.base_plate.toUpperCase().trim(), v);
  });

  const driverByEmail = new Map();
  drivers.forEach((d) => {
    if (d.email) driverByEmail.set(d.email.toLowerCase().trim(), d);
  });

  const groups = new Map();

  tasks.forEach((task) => {
    const rawPlat = (task.plat || '').trim();
    if (!rawPlat || rawPlat === '-' || rawPlat === 'Unknown') return;

    const basePlat = getBasePlate(rawPlat) || rawPlat.toUpperCase();
    const groupKey = `${task.hub_id}_${deliveryDate}_${basePlat}`;

    let group = groups.get(groupKey);
    if (!group) {
      const email = (task.assignee || '').toLowerCase().trim();
      const matchedDriver = driverByEmail.get(email);
      const matchedVehicle = vehicleByPlate.get(basePlat) || vehicleByPlate.get(rawPlat.toUpperCase());

      const maxWeight = matchedVehicle?.max_weight || matchedDriver?.max_weight || 0;
      const maxVolume = matchedVehicle?.max_volume || matchedDriver?.max_volume || 0;
      const rawVehicleType = task.vehicle_type || matchedVehicle?.type || matchedDriver?.type || null;
      const vehicleType = rawVehicleType
        ? String(rawVehicleType).replace(/^(dry|frozen)[\s\-_]*/i, '').trim() || null
        : null;
      const storageType = task.type_storage || matchedVehicle?.storage || matchedDriver?.storage || null;

      group = {
        hub_id: task.hub_id,
        date: deliveryDate,
        _routingSet: new Set(),
        plat: basePlat,
        driver_name: task.driver_name || matchedDriver?.name || 'Driver',
        driver_email: email || null,
        vehicle_type: vehicleType,
        storage_type: storageType,
        act_weight: 0,
        max_weight: maxWeight,
        act_volume: 0,
        max_volume: maxVolume,
        task_data: {
          total: 0,
          success: 0,
          pending: 0,
          cancel: 0,
          partial: 0,
          pending_gr: 0,
        },
        task_ids: [],
        total_distance: 0,
      };
      groups.set(groupKey, group);
    }

    group.act_weight = Number((group.act_weight + (task.weight_kg || 0)).toFixed(2));
    group.act_volume = Number((group.act_volume + (task.volume_cbm || 0)).toFixed(2));

    const dist = task.travel_distance || task.distance || 0;
    group.total_distance = Number((group.total_distance + dist).toFixed(2));

    if (task.routing_result_id) {
      group._routingSet.add(task.routing_result_id);
    }

    group.task_data.total += 1;
    group.task_ids.push(task.id);

    const effectiveStatus = pendingMap.get(task.id) || task.status_delivery || task.status;
    const cat = categorizeDeliveryStatus(effectiveStatus);
    if (cat && group.task_data[cat] !== undefined) {
      group.task_data[cat] += 1;
    }
  });

  const nowIso = new Date().toISOString();
  const resolvedUser = userName || 'system';

  const upsertOps = Array.from(groups.values()).map((item) => {
    const existingSync = existingSyncMap.get(`${item.hub_id}_${item.date}_${item.plat}`);
    const syncUser = {
      created_by: existingSync?.created_by || resolvedUser,
      created_at: existingSync?.created_at || nowIso,
      updated_by: resolvedUser,
      updated_at: nowIso,
    };

    const payload = {
      routing_result_ids: Array.from(item._routingSet),
      driver_name: item.driver_name,
      driver_email: item.driver_email,
      vehicle_type: item.vehicle_type,
      storage_type: item.storage_type,
      act_weight: item.act_weight,
      max_weight: item.max_weight,
      act_volume: item.act_volume,
      max_volume: item.max_volume,
      task_data: item.task_data,
      task_ids: item.task_ids,
      total_distance: item.total_distance,
      sync_user: syncUser,
    };

    return prisma.dailyFleet.upsert({
      where: {
        hub_id_date_plat: {
          hub_id: item.hub_id,
          date: item.date,
          plat: item.plat,
        },
      },
      update: payload,
      create: {
        hub_id: item.hub_id,
        date: item.date,
        plat: item.plat,
        ...payload,
      },
    });
  });

  if (upsertOps.length > 0) {
    await prisma.$transaction(upsertOps);
  }

  return upsertOps.length;
}

export async function recalculateDailyFleetForTask(taskId, userName = 'system') {
  if (!taskId) return;
  const task = await prisma.task.findUnique({
    where: { id: taskId },
    select: { hub_id: true, delivery_date: true },
  });
  if (task?.delivery_date) {
    await syncDailyFleetForDate(task.hub_id, task.delivery_date, userName);
  }
}

