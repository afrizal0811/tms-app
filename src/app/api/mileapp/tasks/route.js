import prisma from '@/lib/prisma';
import {
  formatDateUniversal,
  formatMinutesToHHMM,
  formatUTC7,
  getBasePlate,
  parseCustomerString,
  tomorrowDate,
} from '@/lib/utils';
import { resolveUserName } from '../help';
import { NextResponse } from 'next/server';

const TASK_TIMESTAMP_FIELDS = [
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

const STRIPPED_TASK_FIELDS = ['assignedVehicle', 'parentId', 'subId', 'label', 'taskType'];

function parseJsonField(val, fallback = {}) {
  if (typeof val === 'string') return JSON.parse(val || '{}');
  return val || fallback;
}

function extractDatePart(val) {
  if (!val) return '';
  return val.split(/[ T]/)[0] || '';
}

function resolveTaskDeliveryDate(t, existingTask, routingDateMap) {
  const routingResultId = t.routingResultId || existingTask?.routing_result_id || null;
  const routingDeliveryDate = routingResultId ? routingDateMap.get(routingResultId) : null;
  if (routingDeliveryDate) return routingDeliveryDate;

  const baseDate = extractDatePart(t.startTime || t.createdTime || '');
  if (!baseDate) return '';

  const d = new Date(`${baseDate}T00:00:00Z`);
  if (d.getUTCDay() !== 0) return baseDate;

  const actualDate = extractDatePart(
    t.klikJikaSudahSampai ||
      t.arrival_time ||
      t.doneTime ||
      existingTask?.arrival_time ||
      existingTask?.done_time
  );
  if (actualDate && new Date(`${actualDate}T00:00:00Z`).getUTCDay() !== 0) {
    return actualDate;
  }

  const flow = String(t.flow || existingTask?.flow || '').toLowerCase();
  const createdBy = String(t.createdBy || existingTask?.created_by || '').toLowerCase();
  const isReDelivery = flow.includes('re delivery');
  const isAutomation = createdBy.includes('system') || createdBy.includes('automation');

  if (isReDelivery || isAutomation) {
    d.setUTCDate(d.getUTCDate() + 1);
    return d.toISOString().split('T')[0];
  }

  return baseDate;
}

function getUtcRange(dateStr, offsetHours = 7) {
  if (!dateStr) return { timeFrom: null, timeTo: null };
  const [y, m, d] = dateStr.split('-').map(Number);
  const startUtcMs = Date.UTC(y, m - 1, d, 0, 0, 0) - offsetHours * 3600 * 1000;
  const endUtcMs = Date.UTC(y, m - 1, d, 23, 59, 59, 999) - offsetHours * 3600 * 1000;

  const timeFrom = new Date(startUtcMs).toISOString().slice(0, 19).replace('T', ' ');
  const timeTo = new Date(endUtcMs).toISOString().slice(0, 19).replace('T', ' ');
  return { timeFrom, timeTo };
}

function getLocalDateFromUtcString(utcStr, offsetHours = 7) {
  if (!utcStr) return null;
  const cleanStr =
    utcStr.includes('T') && utcStr.endsWith('Z') ? utcStr : utcStr.replace(' ', 'T') + 'Z';
  const dateObj = new Date(cleanStr);
  if (isNaN(dateObj.getTime())) {
    return utcStr.split(/[ T]/)[0];
  }
  const shiftedMs = dateObj.getTime() + offsetHours * 3600 * 1000;
  return new Date(shiftedMs).toISOString().split('T')[0];
}

function formatDbTask(dbTask) {
  const cust = parseCustomerString(dbTask.customer_order);
  const others = parseJsonField(dbTask.others);
  const listProduct =
    typeof dbTask.list_product === 'string'
      ? JSON.parse(dbTask.list_product || '[]')
      : dbTask.list_product || [];
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

async function saveTasks(tasks, targetHubId, currentUser = 'System') {
  if (!Array.isArray(tasks) || tasks.length === 0) {
    return { savedCount: 0, insertedCount: 0, updatedCount: 0, skippedCount: 0 };
  }

  const uniqueTasksMap = new Map();
  for (const t of tasks) {
    const id = String(t._id || t.id || '');
    if (id) {
      uniqueTasksMap.set(id, t);
    }
  }
  const uniqueTasks = Array.from(uniqueTasksMap.values());
  const tasksToProcess =
    targetHubId && targetHubId !== 'all'
      ? uniqueTasks.filter((t) => !t.hubId || String(t.hubId) === String(targetHubId))
      : uniqueTasks;
  const hubIds = Array.from(
    new Set(tasksToProcess.map((t) => t.hubId || targetHubId).filter(Boolean))
  );
  const taskIds = tasksToProcess.map((t) => String(t._id || t.id)).filter(Boolean);

  const [existingHubs, existingTasks] = await Promise.all([
    hubIds.length > 0
      ? prisma.hub.findMany({
          where: { id: { in: hubIds } },
          select: { id: true },
        })
      : [],
    taskIds.length > 0
      ? prisma.task.findMany({
          where: { id: { in: taskIds } },
          select: {
            id: true,
            status: true,
            status_delivery: true,
            delivery_date: true,
            flow: true,
            routing_result_id: true,
            driver_name: true,
            plat: true,
            assigned_vehicle_id: true,
            arrival_time: true,
            done_time: true,
            reason: true,
            new_coordinate: true,
            done_coordinate: true,
            task_user: true,
            sync_user: true,
            created_from: true,
          },
        })
      : [],
  ]);

  const validHubSet = new Set(existingHubs.map((h) => h.id));
  const existingTaskMap = new Map(existingTasks.map((t) => [t.id, t]));

  const routingResultIds = Array.from(
    new Set(
      tasksToProcess
        .map(
          (t) => t.routingResultId || existingTaskMap.get(String(t._id || t.id))?.routing_result_id
        )
        .filter(Boolean)
    )
  );

  const routingResults =
    routingResultIds.length > 0
      ? await prisma.routingResult.findMany({
          where: { id: { in: routingResultIds } },
          select: { id: true, delivery_date: true },
        })
      : [];
  const routingDateMap = new Map(routingResults.map((r) => [r.id, r.delivery_date]));

  const nowIso = formatUTC7(new Date(), 'YYYY-MM-DDTHH:mm:ss');
  let insertedCount = 0;
  let updatedCount = 0;
  let skippedCount = 0;

  const uniqueEmails = new Set();
  if (currentUser && currentUser.includes('@')) uniqueEmails.add(currentUser);
  tasksToProcess.forEach((t) => {
    if (t.createdBy && t.createdBy.includes('@')) uniqueEmails.add(t.createdBy);
    if (t.updatedBy && t.updatedBy.includes('@')) uniqueEmails.add(t.updatedBy);
  });
  existingTasks.forEach((et) => {
    const su = parseJsonField(et.sync_user);
    if (su.created_by && su.created_by.includes('@')) uniqueEmails.add(su.created_by);
    if (su.updated_by && su.updated_by.includes('@')) uniqueEmails.add(su.updated_by);
  });

  const userNameMap = new Map();
  await Promise.all(
    Array.from(uniqueEmails).map(async (email) => {
      const name = await resolveUserName(email);
      userNameMap.set(email.toLowerCase().trim(), name);
    })
  );

  const resolveName = (val) => {
    if (!val) return val;
    return userNameMap.get(String(val).toLowerCase().trim()) || val;
  };

  const resolvedCurrentUser = resolveName(currentUser) || 'System';

  const CHUNK_SIZE = 25;
  for (let i = 0; i < tasksToProcess.length; i += CHUNK_SIZE) {
    const chunk = tasksToProcess.slice(i, i + CHUNK_SIZE);
    const operations = chunk.map((t) => {
      const id = String(t._id || t.id);
      const existingTask = existingTaskMap.get(id);

      const resolvedHubId =
        targetHubId && targetHubId !== 'all'
          ? targetHubId
          : validHubSet.has(t.hubId)
            ? t.hubId
            : existingHubs[0]?.id;

      if (!resolvedHubId) return Promise.resolve(null);

      const routingResultId = t.routingResultId || existingTask?.routing_result_id || null;
      const deliveryDate = resolveTaskDeliveryDate(t, existingTask, routingDateMap);
      const taskReason = t.reason || t.alasan || null;

      const others = {
        doneOrder: t.doneOrder ?? null,
        isSplitTask: t.isSplitTask,
        openTime: t.openTime,
        closeTime: t.closeTime,
        visitTime: t.visitTime,
        maksimumVehicleType: t.maksimumVehicleType,
        page1DoneTime: t.page1DoneTime,
        page2DoneTime: t.page2DoneTime,
        page3DoneTime: t.page3DoneTime,
        maxWeight: t.maxWeight,
        maxVolume: t.maxVolume,
        weightPct: t.weightPct,
        volumePct: t.volumePct,
      };

      const taskUser = {
        created_by: resolveName(t.createdBy) || null,
        created_at: t.createdTime || null,
        updated_by: resolveName(t.updatedBy) || null,
        updated_at: t.updatedTime || null,
      };

      if (existingTask) {
        const existingTaskUser = parseJsonField(existingTask.task_user);

        const sameRouting = routingResultId === (existingTask.routing_result_id || null);

        const isFieldChanged =
          deliveryDate !== (existingTask.delivery_date || '') ||
          (t.status || null) !== (existingTask.status || null) ||
          (t.statusDelivery || null) !== (existingTask.status_delivery || null) ||
          !sameRouting ||
          (t.driverName || null) !== (existingTask.driver_name || null) ||
          (t.plat || null) !== (existingTask.plat || null) ||
          (t.assignedVehicleId || null) !== (existingTask.assigned_vehicle_id || null) ||
          (t.klikJikaSudahSampai || null) !== (existingTask.arrival_time || null) ||
          (t.doneTime || null) !== (existingTask.done_time || null) ||
          taskReason !== (existingTask.reason || null) ||
          (t.klikLokasiClient || null) !== (existingTask.new_coordinate || null) ||
          (t.doneCoordinate || null) !== (existingTask.done_coordinate || null);

        if (!isFieldChanged && existingTaskUser.updated_at) {
          return Promise.resolve({ type: 'skip' });
        }
      }

      const existingSyncUser = existingTask?.sync_user
        ? parseJsonField(existingTask.sync_user, null)
        : null;

      const syncUser = {
        created_by: resolveName(existingSyncUser?.created_by) || resolvedCurrentUser,
        created_at: existingSyncUser?.created_at || nowIso,
        updated_by: resolvedCurrentUser,
        updated_at: nowIso,
      };

      const data = {
        hub_id: resolvedHubId,
        routing_result_id: routingResultId,
        customer_order: t.customerOrder || null,
        address: t.address || null,
        flow: t.flow || null,
        type_storage: t.typeStorage || null,
        status: t.status || null,
        status_delivery: t.statusDelivery || null,
        reason: taskReason,
        delivery_date: deliveryDate,
        start_time: t.startTime || null,
        end_time: t.endTime || null,
        assigned_time: t.assignedTime || null,
        done_time: t.doneTime || null,
        arrival_time: t.klikJikaSudahSampai || null,
        assigned_vehicle_id: t.assignedVehicleId || null,
        assignee: t.assignee || null,
        driver_name: t.driverName || null,
        plat: t.plat || null,
        vehicle_type: t.vehicleType || null,
        route_planned_order: t.routePlannedOrder != null ? Number(t.routePlannedOrder) : null,
        weight_kg: t.weightKg != null ? Number(t.weightKg) : null,
        volume_cbm: t.volumeCbm != null ? Number(t.volumeCbm) : null,
        distance: t.distance != null ? Number(t.distance) : null,
        travel_distance: t.travelDistance != null ? Number(t.travelDistance) : null,
        travel_duration: t.travelDuration || null,
        eta: t.eta || null,
        etd: t.etd || null,
        expected_coordinate: t.expectedCoordinate || t.longlat || null,
        done_coordinate: t.doneCoordinate || null,
        new_coordinate: t.klikLokasiClient || null,
        list_product: t.listProduct || [],
        others,
        created_from: t.createdFrom || 'API',
        sync_user: syncUser,
        task_user: taskUser,
      };

      if (existingTask) {
        return prisma.task
          .update({
            where: { id },
            data,
          })
          .then(() => ({ type: 'update' }));
      }

      return prisma.task
        .upsert({
          where: { id },
          update: data,
          create: {
            id,
            ...data,
          },
        })
        .then(() => ({ type: 'insert' }))
        .catch(async (err) => {
          if (err?.code === 'P2002') {
            await prisma.task.update({ where: { id }, data });
            return { type: 'update' };
          }
          throw err;
        });
    });

    const results = await Promise.all(operations);
    insertedCount += results.filter((r) => r?.type === 'insert').length;
    updatedCount += results.filter((r) => r?.type === 'update').length;
    skippedCount += results.filter((r) => r?.type === 'skip').length;
  }

  return {
    savedCount: insertedCount + updatedCount,
    insertedCount,
    updatedCount,
    skippedCount,
  };
}

async function fetchAndEnrichTasks({
  hubId,
  timeFrom,
  timeTo,
  status = 'DONE,ONGOING,UNASSIGNED',
  timeBy = 'startTime',
  limit = 1000,
  fields,
}) {
  const apiUrl = process.env.NEXT_PUBLIC_API_URL;
  const apiToken = process.env.API_TOKEN;

  if (!apiUrl || !apiToken) {
    throw new Error('Variabel API tidak diatur di server.');
  }

  const fetchTasksPage = async (pageNumber) => {
    const externalUrl = new URL(`${apiUrl}/tasks`);
    if (hubId && hubId !== 'all') {
      externalUrl.searchParams.append('hubId', hubId);
    }
    externalUrl.searchParams.append('timeFrom', timeFrom);
    externalUrl.searchParams.append('timeTo', timeTo);
    externalUrl.searchParams.append('status', status);
    externalUrl.searchParams.append('timeBy', timeBy);
    externalUrl.searchParams.append('limit', limit);
    externalUrl.searchParams.append('page', pageNumber);
    if (fields) {
      externalUrl.searchParams.append('fields', fields);
    }

    const res = await fetch(externalUrl.toString(), {
      headers: {
        Authorization: `Bearer ${apiToken}`,
        'Content-Type': 'application/json',
      },
    });

    if (!res.ok) {
      throw new Error(`Error fetching page ${pageNumber}: ${res.statusText}`);
    }
    return res.json();
  };

  const firstPageData = await fetchTasksPage(1);

  if (!firstPageData || !firstPageData.tasks || !Array.isArray(firstPageData.tasks.data)) {
    return [];
  }
  let allTasks = [...firstPageData.tasks.data];
  const lastPage = firstPageData.tasks.last_page || 1;

  if (lastPage > 1) {
    const promises = [];
    for (let page = 2; page <= lastPage; page++) {
      promises.push(fetchTasksPage(page));
    }
    const remainingPages = await Promise.all(promises);

    remainingPages.forEach((pageData) => {
      if (pageData.tasks && Array.isArray(pageData.tasks.data)) {
        allTasks = allTasks.concat(pageData.tasks.data);
      }
    });
  }

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
    const where = hubId && hubId !== 'all' ? { hubs: { some: { id: hubId } } } : {};
    const [rawDrivers, mappingsDB] = await Promise.all([
      prisma.driver.findMany({ where }),
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

  allTasks.sort((a, b) => new Date(a.createdTime || 0) - new Date(b.createdTime || 0));

  allTasks.forEach((task) => {
    for (const field of TASK_TIMESTAMP_FIELDS) {
      if (task[field]) task[field] = formatUTC7(task[field], 'YYYY-MM-DDTHH:mm:ss');
    }

    if (task.volumeCbm != null) task.volumeCbm = Number(Number(task.volumeCbm).toFixed(2));
    if (task.weightKg != null) task.weightKg = Number(Number(task.weightKg).toFixed(2));

    if (task.distance != null) task.distance = Number((task.distance / 1000).toFixed(2));
    if (task.travelDistance != null)
      task.travelDistance = Number((task.travelDistance / 1000).toFixed(2));

    if (task.travelDuration != null) task.travelDuration = formatMinutesToHHMM(task.travelDuration);

    if (task.eta && typeof task.eta === 'string') task.eta = task.eta.substring(0, 5);
    if (task.etd && typeof task.etd === 'string') task.etd = task.etd.substring(0, 5);

    for (const field of STRIPPED_TASK_FIELDS) {
      delete task[field];
    }

    const assigneeEmail =
      task.assignee && task.assignee.length > 0 ? task.assignee[0].toLowerCase().trim() : null;
    const driverData = assigneeEmail ? driversByEmail[assigneeEmail] : null;

    if (driverData) {
      task.driverName = driverData.name;
      task.plat = driverData.plat;
      task.basePlat = driverData.basePlat;
      task.vehicleType = driverData.type;
      task.maxVolume = driverData.maxVolume;
      task.maxWeight = driverData.maxWeight;
    }

    task.weightPct =
      task.maxWeight && task.weightKg != null
        ? `${((task.weightKg / task.maxWeight) * 100).toFixed(2)}%`
        : '0%';
    task.volumePct =
      task.maxVolume && task.volumeCbm != null
        ? `${((task.volumeCbm / task.maxVolume) * 100).toFixed(2)}%`
        : '0%';

    if (Array.isArray(task.statusDelivery)) task.statusDelivery = task.statusDelivery.join(', ');
    if (Array.isArray(task.assignee)) task.assignee = task.assignee.join(', ');
  });

  return allTasks;
}

export async function GET(request) {
  try {
    const { searchParams } = new URL(request.url);
    const hubId = searchParams.get('hubId');
    const timeFrom = searchParams.get('timeFrom');
    const timeTo = searchParams.get('timeTo');
    const dateFrom = searchParams.get('dateFrom');
    const dateTo = searchParams.get('dateTo');
    const status = searchParams.get('status');
    const deliveryDate = searchParams.get('deliveryDate');

    let fromDate = null;
    let toDate = null;

    if (deliveryDate) {
      fromDate = deliveryDate;
      toDate = deliveryDate;
    } else if (dateFrom && dateTo) {
      fromDate = dateFrom.split(/[ T]/)[0];
      toDate = dateTo.split(/[ T]/)[0];
    } else if (dateFrom) {
      fromDate = dateFrom.split(/[ T]/)[0];
      toDate = fromDate;
    } else if (timeFrom && timeTo) {
      fromDate = getLocalDateFromUtcString(timeFrom);
      toDate = getLocalDateFromUtcString(timeTo);
    }

    if (!fromDate && !deliveryDate) {
      return NextResponse.json(
        { error: 'Missing required parameters: deliveryDate, dateRange, or timeFrom & timeTo' },
        { status: 400 }
      );
    }

    const where = {};
    if (fromDate && toDate) {
      where.delivery_date = fromDate === toDate ? fromDate : { gte: fromDate, lte: toDate };
    } else if (fromDate) {
      where.delivery_date = fromDate;
    }

    if (hubId && hubId !== 'all') {
      where.hub_id = hubId;
    }
    if (status && status !== 'all') {
      const statusList = status
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
      if (statusList.length === 1) {
        where.status = statusList[0];
      } else if (statusList.length > 1) {
        where.status = { in: statusList };
      }
    }

    const existingDbTasks = await prisma.task.findMany({
      where,
      orderBy: [{ route_planned_order: 'asc' }, { start_time: 'asc' }],
    });

    return NextResponse.json(existingDbTasks.map(formatDbTask));
  } catch (error) {
    console.error('Error Tasks:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      {
        error: 'Gagal mengambil data tasks dari database',
        detail: errorMessage,
      },
      { status: 500 }
    );
  }
}

export async function POST(request) {
  try {
    const body = await request.json();
    const {
      autoUpdate,
      deliveryDate,
      hubId,
      timeFrom,
      timeTo,
      status = 'DONE,ONGOING,UNASSIGNED',
      timeBy = 'startTime',
      created_by,
      updated_by,
    } = body;

    const user = updated_by || created_by || 'System';

    if (autoUpdate) {
      const targetWhere = {
        status: { in: ['UNASSIGNED', 'ONGOING'] },
      };
      if (hubId && hubId !== 'all') {
        targetWhere.hub_id = hubId;
      }

      const pendingTasks = await prisma.task.findMany({
        where: targetWhere,
        select: {
          delivery_date: true,
          start_time: true,
        },
      });

      const datesToSync = new Set();
      pendingTasks.forEach((t) => {
        if (t.delivery_date) {
          datesToSync.add(t.delivery_date);
        } else if (t.start_time) {
          datesToSync.add(t.start_time.split(/[ T]/)[0]);
        }
      });

      const todayStr = formatDateUniversal(new Date(), 'YYYY-MM-DD');
      const tomorrowStr = formatDateUniversal(tomorrowDate(true), 'YYYY-MM-DD');
      datesToSync.add(todayStr);
      datesToSync.add(tomorrowStr);

      let totalUpdated = 0;
      for (const dStr of datesToSync) {
        try {
          const { timeFrom: tFrom, timeTo: tTo } = getUtcRange(dStr);
          const tasks = await fetchAndEnrichTasks({
            hubId,
            timeFrom: tFrom,
            timeTo: tTo,
            status: 'DONE,ONGOING,UNASSIGNED',
            timeBy: 'startTime',
          });

          if (tasks.length > 0) {
            const stats = await saveTasks(tasks, hubId, user);
            const count = typeof stats === 'object' ? stats.savedCount || 0 : stats || 0;
            totalUpdated += count;
          }
        } catch (err) {
          console.warn(`[AutoUpdate] Gagal sync tanggal ${dStr}:`, err.message);
        }
      }

      return NextResponse.json({
        message: `Auto-update selesai. ${totalUpdated} tasks diproses.`,
        updatedCount: totalUpdated,
        dates: Array.from(datesToSync),
      });
    }

    let targetDate = deliveryDate;
    let tFrom = timeFrom;
    let tTo = timeTo;

    if (targetDate) {
      const utcRange = getUtcRange(targetDate);
      tFrom = tFrom || utcRange.timeFrom;
      tTo = tTo || utcRange.timeTo;
    } else if (timeFrom && timeTo) {
      targetDate = getLocalDateFromUtcString(timeFrom);
    }

    if (!tFrom || !tTo) {
      return NextResponse.json(
        { error: 'Parameter deliveryDate atau timeFrom & timeTo harus disertakan.' },
        { status: 400 }
      );
    }

    const allTasks = await fetchAndEnrichTasks({
      hubId,
      timeFrom: tFrom,
      timeTo: tTo,
      status,
      timeBy,
    });

    const stats = await saveTasks(allTasks, hubId, user);
    const savedCount = typeof stats === 'object' ? stats.savedCount : stats;
    const insertedCount = stats?.insertedCount || 0;
    const updatedCount = stats?.updatedCount || 0;
    const skippedCount = stats?.skippedCount || 0;

    return NextResponse.json({
      message: `Berhasil sinkronisasi tasks: ${insertedCount} baru, ${updatedCount} diperbarui, ${skippedCount} dilewati (tidak berubah).`,
      syncedCount: savedCount,
      insertedCount,
      updatedCount,
      skippedCount,
      deliveryDate: targetDate,
    });
  } catch (error) {
    console.error('Error Sync Tasks:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      {
        error: 'Gagal sinkronisasi data tasks',
        detail: errorMessage,
      },
      { status: 500 }
    );
  }
}
