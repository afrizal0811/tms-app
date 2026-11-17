import prisma from '@/lib/prisma';
import { formatUTC7, getBasePlate, getDeliveryDateFromRouting, isEmpty } from '@/lib/utils';
import { NextResponse } from 'next/server';

const STRIPPED_TRIP_FIELDS = [
  'etaStr',
  'etdStr',
  'allowOddEven',
  'timeWindow',
  'visitGroup',
  'visitGroupPriority',
  'tags',
];

const STRIPPED_ROUTE_FIELDS = [
  'workingTime',
  'breakTime',
  'vehicleTags',
  'oddEven',
  'isOddEven',
  'speed',
  'fixedCost',
  'totalNodes',
  'totalTrips',
  'finishTime',
];

const parseJsonObject = (val, fallback = {}) => {
  if (typeof val === 'string') return JSON.parse(val || '{}');
  return val || fallback;
};

const normalizeHubId = (val) =>
  val && String(val).trim() !== '' && val !== 'undefined' ? String(val).trim() : null;

const toKm = (m) => parseFloat(((m || 0) / 1000).toFixed(2));

const toPct = (val, max) => parseFloat((max ? ((val || 0) / max) * 100 : 0).toFixed(2));

const toHr = (mins) => {
  const h = Math.floor((mins || 0) / 60);
  const m = Math.floor((mins || 0) % 60);
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
};

const formatRoutingHistories = (rawHistory) => {
  const grouped = {
    move: [],
    dropped: [],
    switch: [],
    manual: [],
    change: [],
  };

  (rawHistory || []).forEach((item) => {
    const act = (item.action || '').toLowerCase();
    const vFrom = (item.vehicleFrom || '').toLowerCase();

    if (act === 'move' && vFrom === 'dropped') {
      grouped.manual.push(item);
      return;
    }

    if (act === 'move') {
      grouped.move.push(item);
    } else if (act === 'dropped' || act === 'drop') {
      grouped.dropped.push(item);
    } else if (act === 'switch') {
      grouped.switch.push(item);
    } else if (act === 'change') {
      grouped.change.push(item);
    }
    delete item.organizationId;
    delete item.createdType;
    delete item.createdFrom;
    delete item.updatedTime;
    delete item.resultId;
  });

  const total =
    grouped.move.length +
    grouped.dropped.length +
    grouped.switch.length +
    grouped.manual.length +
    grouped.change.length;

  return {
    total,
    history: grouped,
  };
};

export async function GET(request) {
  try {
    const { searchParams } = new URL(request.url);
    const deliveryDate = searchParams.get('deliveryDate');
    const dateFrom = searchParams.get('dateFrom');
    const dateTo = searchParams.get('dateTo');
    const hubId = searchParams.get('hubId');

    const where = {};
    if (deliveryDate) {
      where.delivery_date = deliveryDate;
    } else if (dateFrom && dateTo) {
      where.delivery_date = {
        gte: dateFrom.split(/[ T]/)[0],
        lte: dateTo.split(/[ T]/)[0],
      };
    } else if (dateFrom) {
      where.delivery_date = { gte: dateFrom.split(/[ T]/)[0] };
    }

    if (hubId && hubId !== 'all') {
      where.hub_id = hubId;
    }

    const routingResults = await prisma.routingResult.findMany({
      where,
      include: {
        hub: true,
      },
      orderBy: [{ delivery_date: 'asc' }, { id: 'asc' }],
    });

    const deliveryDates = Array.from(
      new Set(routingResults.map((r) => r.delivery_date).filter(Boolean))
    );
    const hubIds = Array.from(new Set(routingResults.map((r) => r.hub_id).filter(Boolean)));

    const taskCountMap = new Map();

    if (deliveryDates.length > 0) {
      const taskCounts = await prisma.task.groupBy({
        by: ['hub_id', 'delivery_date'],
        where: {
          delivery_date: { in: deliveryDates },
          ...(hubIds.length > 0 ? { hub_id: { in: hubIds } } : {}),
        },
        _count: {
          id: true,
        },
      });

      taskCounts.forEach((tc) => {
        taskCountMap.set(`${tc.hub_id}_${tc.delivery_date}`, tc._count.id);
        const prev = taskCountMap.get(tc.delivery_date) || 0;
        taskCountMap.set(tc.delivery_date, prev + tc._count.id);
      });
    }

    const formattedData = routingResults.map((item) => {
      const hubTimes = parseJsonObject(item.hub_times);
      const totalDistance = parseJsonObject(item.total_distance);
      const timeData = parseJsonObject(item.time);

      const formattedTime = {};
      for (const [p, t] of Object.entries(timeData)) {
        formattedTime[p] = {
          spentTime: Number(t.spentTime ?? t.spent_time) || 0,
          visitTime: Number(t.visitTime ?? t.visit_time) || 0,
          travelTime: Number(t.travelTime ?? t.travel_time) || 0,
          waitingTime: Number(t.waitingTime ?? t.waiting_time) || 0,
        };
      }

      const taskCount =
        taskCountMap.get(`${item.hub_id}_${item.delivery_date}`) ??
        taskCountMap.get(item.delivery_date) ??
        0;

      const syncUser = parseJsonObject(item.sync_user);
      const routeUser = parseJsonObject(item.route_user);

      return {
        id: item.id,
        _id: item.id,
        hubId: item.hub_id,
        name: item.name,
        deliveryDate: item.delivery_date,
        routeUser,
        route_user: routeUser,
        syncUser,
        sync_user: syncUser,
        createdBy: syncUser.created_by || null,
        createdAt: syncUser.created_at || null,
        updatedBy: syncUser.updated_by || null,
        updatedAt: syncUser.updated_at || null,
        createdTime: routeUser.created_at || item.delivery_date,
        taskCount,
        dispatchStatus: 'done',
        dropped: item.dropped,
        histories: item.histories,
        totalDistance,
        time: formattedTime,
        hubTimes,
        hub: item.hub,
      };
    });

    return NextResponse.json(formattedData);
  } catch (error) {
    console.error('Error Results:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      {
        error: 'Gagal mengambil data results dari database',
        detail: errorMessage,
      },
      { status: 500 }
    );
  }
}

const getDeliveryDate = async (item, apiUrl, apiToken) => {
  const customerTaskIds = [];
  for (const r of item.result?.routing || []) {
    for (const t of r.trips || []) {
      if (!t.isHub && t.visitId) {
        const id = String(t.visitId)
          .replace(/^taskId-/, '')
          .trim();
        if (id) customerTaskIds.push(id);
      }
    }
  }

  if (customerTaskIds.length === 0) {
    return getDeliveryDateFromRouting(item?.createdTime);
  }

  const sampleIds = [];
  if (customerTaskIds.length <= 15) {
    sampleIds.push(...customerTaskIds);
  } else {
    const step = customerTaskIds.length / 15;
    for (let i = 0; i < 15; i++) {
      sampleIds.push(customerTaskIds[Math.floor(i * step)]);
    }
  }

  const taskDates = await Promise.all(
    sampleIds.map(async (id) => {
      try {
        const res = await fetch(`${apiUrl}/task/${id}`, {
          headers: {
            Authorization: `Bearer ${apiToken}`,
            'Content-Type': 'application/json',
          },
          cache: 'no-store',
        });
        if (!res.ok) return null;
        const json = await res.json();
        const task = json?.task || json?.data || json;
        const rawTime = task?.startTime || task?.taskDate || task?.date;
        return rawTime ? formatUTC7(rawTime, 'YYYY-MM-DD') : null;
      } catch {
        return null;
      }
    })
  );

  const validDates = taskDates.filter(Boolean);
  if (validDates.length === 0) {
    return getDeliveryDateFromRouting(item?.createdTime);
  }

  const counts = {};
  for (const d of validDates) {
    counts[d] = (counts[d] || 0) + 1;
  }

  let majorityDate = null;
  let maxCount = -1;
  for (const [dateStr, count] of Object.entries(counts)) {
    if (count > maxCount) {
      maxCount = count;
      majorityDate = dateStr;
    }
  }

  return majorityDate;
};

async function buildDriversMap(hubId) {
  const hubFilter = hubId && hubId !== 'all' ? { hub_id: String(hubId) } : {};
  const vehicleHubFilter = hubId && hubId !== 'all' ? { hubs_id: String(hubId) } : {};
  const [rawDrivers, vehiclesList] = await Promise.all([
    prisma.driverList.findMany({ where: hubFilter }),
    prisma.vehicle.findMany({ where: vehicleHubFilter }),
  ]);

  const vehicleByDriverId = new Map(
    vehiclesList.filter((v) => v.driver_id).map((v) => [v.driver_id, v])
  );
  const driversMap = {};
  rawDrivers.forEach((d) => {
    const email = (d.email || '').toLowerCase().trim();
    if (email && email !== '-') {
      const v = vehicleByDriverId.get(d.id);
      driversMap[email] = { name: d.name, plat: v ? getBasePlate(v.license_number) : null };
    }
  });

  if (rawDrivers.length === 0) {
    const fallbackDrivers = await prisma.driver.findMany({
      where: hubId && hubId !== 'all' ? { hubs: { some: { id: hubId } } } : {},
    });
    fallbackDrivers.forEach((d) => {
      const email = (d.email || '').toLowerCase().trim();
      if (email && email !== '-' && !driversMap[email]) {
        driversMap[email] = { name: d.name, plat: getBasePlate(d.plat) };
      }
    });
  }

  return driversMap;
}

function cleanRoutingEntries(routing, driversMap) {
  return (routing || []).map((r) => {
    let tW = 0,
      tV = 0,
      tD = 0,
      tTr = 0,
      tVi = 0,
      tWa = 0;

    if (r.trips && r.trips.length > 0) {
      const lastHubWait = Number(r.trips[r.trips.length - 1].waitingTime) || 0;
      if ((Number(r.trips[0].waitingTime) || 0) <= 0) {
        r.trips[0].waitingTime = lastHubWait;
      }

      r.trips = r.trips.map((t, i) => {
        if (t.isHub === false && t.visitId?.startsWith('taskId-')) {
          t.visitId = t.visitId.replace('taskId-', '');
        }

        const tripTravel = Number(t.travelTime) || 0;
        const tripVisit = Number(t.visitTime) || 0;
        const tripWait = Number(t.waitingTime) || 0;

        if (!t.isHub) {
          tW += t.weight || 0;
          tV += t.volume || 0;
        }
        tD += t.distance || 0;
        tTr += tripTravel;
        tVi += tripVisit;
        if (i !== r.trips.length - 1) {
          tWa += tripWait;
        }

        t.travelTime = tripTravel;
        t.visitTime = tripVisit;
        t.waitingTime = tripWait;
        t.spentTime = tripTravel + tripVisit + tripWait;

        for (const field of STRIPPED_TRIP_FIELDS) {
          delete t[field];
        }
        return t;
      });
    }

    const spent = tTr + tVi + tWa;
    r.totalWeight = isEmpty(tW) ? 0 : Number(tW).toFixed(2);
    r.totalVolume = isEmpty(tV) ? 0 : Number(tV).toFixed(2);
    r.weightPercentage = toPct(tW, r.vehicleMaxWeight);
    r.volumePercentage = toPct(tV, r.vehicleMaxVolume);
    r.totalDistance = tD;
    r.distanceKm = toKm(tD);
    r.totalTravelTime = tTr;
    r.travelHour = toHr(tTr);
    r.totalVisitTime = tVi;
    r.visitHour = toHr(tVi);
    r.totalWaitingTime = tWa;
    r.waitingHour = toHr(tWa);
    r.totalSpentTime = spent;
    r.spentHour = toHr(spent);

    const email = (r.assignee || '').toLowerCase().trim();
    r.driverName = driversMap[email]?.name || null;
    if (!r.basePlat && !r.vehicleName) {
      r.basePlat = driversMap[email]?.plat || null;
    } else {
      r.basePlat = getBasePlate(r.basePlat || r.vehicleName);
    }

    for (const field of STRIPPED_ROUTE_FIELDS) {
      delete r[field];
    }

    return r;
  });
}

function buildTripMaps(activeVehicles, seenVehicles) {
  const totalDistanceMap = {};
  const timeMap = {};
  const hubTimesMap = {};

  for (const r of activeVehicles) {
    const taskIds = (r.trips || [])
      .filter((t) => !t.isHub && t.visitId)
      .map((t) => String(t.visitId));

    const vehicleName = r.vehicleName || r.basePlat || '-';
    const vehicleKey = (r.vehicleId ? String(r.vehicleId).trim() : vehicleName)
      .toLowerCase()
      .trim();
    const tripSignature = [...taskIds].sort().join(',');

    const existingTrips = seenVehicles.get(vehicleKey);
    if (existingTrips && existingTrips.has(tripSignature)) {
      continue;
    }

    if (!seenVehicles.has(vehicleKey)) {
      seenVehicles.set(vehicleKey, new Set());
    }
    seenVehicles.get(vehicleKey).add(tripSignature);

    const plateKey = r.basePlat || getBasePlate(r.vehicleName) || vehicleName;

    totalDistanceMap[plateKey] = Number(r.totalDistance) || 0;
    timeMap[plateKey] = {
      spent_time: Number(r.totalSpentTime) || 0,
      visit_time: Number(r.totalVisitTime) || 0,
      travel_time: Number(r.totalTravelTime) || 0,
      waiting_time: Number(r.totalWaitingTime) || 0,
    };

    const rawHubTrips = (r.trips || []).filter((t) => t.isHub === true);
    const middleCount = rawHubTrips.length - 2;
    let middleIndex = 0;

    hubTimesMap[plateKey] = rawHubTrips.map((t, idx) => {
      let category;
      if (idx === 0) {
        category = 'first';
      } else if (idx === rawHubTrips.length - 1) {
        category = 'last';
      } else {
        middleIndex++;
        category = middleCount === 1 ? 'middle' : `middle_${middleIndex}`;
      }

      return {
        visitId: t.visitId,
        isHub: true,
        order: t.order ?? idx,
        eta: t.eta || null,
        etd: t.etd || null,
        category,
      };
    });
  }

  return { totalDistanceMap, timeMap, hubTimesMap };
}

async function fetchRoutingHistories(resultId, apiUrl, apiToken) {
  const emptyHistory = {
    total: 0,
    history: { move: [], dropped: [], switch: [], manual: [], change: [] },
  };
  try {
    const histResponse = await fetch(`${apiUrl}/result/${resultId}/history`, {
      headers: {
        Authorization: `Bearer ${apiToken}`,
        'Content-Type': 'application/json',
      },
      cache: 'no-store',
    });
    if (!histResponse.ok) return emptyHistory;
    const histJson = await histResponse.json();
    return formatRoutingHistories(histJson.data || []);
  } catch (histError) {
    console.error(`Error fetching history for ${resultId}:`, histError);
    return emptyHistory;
  }
}

function formatDroppedData(rawDropped) {
  const list = Array.isArray(rawDropped) ? rawDropped : [];
  return {
    total: list.length,
    data: list.map((d) => ({
      id: d.refId || (typeof d.id === 'string' ? d.id.replace(/_\d+$/, '') : d.id),
      name: d.name || d.visitName || '',
      tags: d.tags || [],
      reason: d.reason || '',
      address: d.address || '',
      capacity: d.capacity ?? null,
      timeWindow: d.timeWindow ?? null,
    })),
  };
}

export async function POST(request) {
  try {
    const body = await request.json();
    const { deliveryDate, hubId, force = false, created_by, updated_by, userName } = body || {};
    const currentUser = updated_by || created_by || userName || null;

    if (!deliveryDate) {
      return NextResponse.json(
        { error: 'Parameter "deliveryDate" (format YYYY-MM-DD) dibutuhkan.' },
        { status: 400 }
      );
    }

    const apiUrl = process.env.NEXT_PUBLIC_API_URL;
    const apiToken = process.env.API_TOKEN;

    if (!apiUrl || !apiToken) {
      return NextResponse.json(
        {
          error:
            'Konfigurasi server API hilang: NEXT_PUBLIC_API_URL atau API_TOKEN tidak ditemukan.',
        },
        { status: 500 }
      );
    }

    const deliveryDateObj = new Date(`${deliveryDate}T00:00:00Z`);
    const searchStartObj = new Date(deliveryDateObj);
    searchStartObj.setUTCDate(searchStartObj.getUTCDate() - 7);
    const searchEndObj = new Date(deliveryDateObj);

    const pad = (n) => String(n).padStart(2, '0');
    const dateFromStr = `${searchStartObj.getUTCFullYear()}-${pad(searchStartObj.getUTCMonth() + 1)}-${pad(searchStartObj.getUTCDate())} 00:00:00`;
    const dateToStr = `${searchEndObj.getUTCFullYear()}-${pad(searchEndObj.getUTCMonth() + 1)}-${pad(searchEndObj.getUTCDate())} 23:59:59`;

    const externalUrl = new URL(`${apiUrl}/results`);
    externalUrl.searchParams.append('dateFrom', dateFromStr);
    externalUrl.searchParams.append('dateTo', dateToStr);
    externalUrl.searchParams.append('limit', '1000');
    if (hubId && hubId !== 'all') {
      externalUrl.searchParams.append('hubId', hubId);
    }

    const externalResponse = await fetch(externalUrl.toString(), {
      headers: {
        Authorization: `Bearer ${apiToken}`,
        'Content-Type': 'application/json',
      },
      cache: 'no-store',
    });

    const data = await externalResponse.json();
    if (!externalResponse.ok) {
      return NextResponse.json(
        { error: 'Gagal mengambil data results dari vendor API', details: data },
        { status: externalResponse.status }
      );
    }

    const routingData = data?.data?.data || (Array.isArray(data?.data) ? data.data : []);

    const candidateResults = routingData.filter((item) => {
      const isDone = item?.dispatchStatus?.toLowerCase() === 'done';
      const message = item?.dispatchMessage;
      return isDone && !isEmpty(message);
    });

    const targetResultsWithDate = (
      await Promise.all(
        candidateResults.map(async (item) => {
          const itemDeliveryDate = await getDeliveryDate(item, apiUrl, apiToken);
          return itemDeliveryDate === deliveryDate ? item : null;
        })
      )
    ).filter(Boolean);

    const targetResults = targetResultsWithDate.sort(
      (a, b) => new Date(a.createdTime || 0) - new Date(b.createdTime || 0)
    );

    if (targetResults.length === 0) {
      return NextResponse.json({
        message: `Tidak ditemukan routing berstatus done untuk tanggal pengiriman ${deliveryDate}.`,
        syncedCount: 0,
      });
    }

    const driversMap = await buildDriversMap(hubId);

    const targetResultIds = targetResults.map((t) => String(t._id));
    if (force) {
      await prisma.routingResult.deleteMany({
        where: {
          delivery_date: deliveryDate,
          ...(hubId && hubId !== 'all' ? { hub_id: String(hubId) } : {}),
          id: { notIn: targetResultIds },
        },
      });
    }

    const seenVehicles = new Map();
    const savedRecords = [];
    const updatedRecords = [];
    const skippedRecords = [];

    for (const item of targetResults) {
      const cleanedRouting = cleanRoutingEntries(item.result?.routing, driversMap);
      const activeVehicles = cleanedRouting.filter((r) =>
        (r.trips || []).some((t) => !t.isHub && t.visitId)
      );

      const rawUser = typeof item.user === 'object' ? item.user?.name : item.user || null;
      const routeUser = {
        created_by: rawUser || item.createdBy || null,
        created_at: item.createdTime ? formatUTC7(item.createdTime, 'YYYY-MM-DDTHH:mm:ss') : null,
      };

      const { totalDistanceMap, timeMap, hubTimesMap } = buildTripMaps(
        activeVehicles,
        seenVehicles
      );

      const targetHubId =
        normalizeHubId(item.hubId) ||
        normalizeHubId(hubId) ||
        item.result?.routing?.[0]?.trips
          ?.find((t) => t.isHub && t.visitId?.startsWith('hub-'))
          ?.visitId?.replace('hub-', '') ||
        null;

      if (!targetHubId) {
        throw new Error(`Hub ID tidak ditemukan untuk data routing "${item.name || item._id}".`);
      }

      const historiesData = await fetchRoutingHistories(item._id, apiUrl, apiToken);
      const droppedData = formatDroppedData(item.result?.dropped);

      const nowIso = formatUTC7(new Date(), 'YYYY-MM-DDTHH:mm:ss');
      const existingRouting = await prisma.routingResult.findUnique({
        where: { id: String(item._id) },
        select: {
          id: true,
          hub_id: true,
          name: true,
          dropped: true,
          total_distance: true,
          time: true,
          hub_times: true,
          histories: true,
          sync_user: true,
        },
      });

      const isSame =
        existingRouting &&
        existingRouting.hub_id === targetHubId &&
        existingRouting.name === (item.name || '') &&
        JSON.stringify(existingRouting.dropped) === JSON.stringify(droppedData) &&
        JSON.stringify(existingRouting.total_distance) === JSON.stringify(totalDistanceMap) &&
        JSON.stringify(existingRouting.time) === JSON.stringify(timeMap) &&
        JSON.stringify(existingRouting.hub_times) === JSON.stringify(hubTimesMap) &&
        JSON.stringify(existingRouting.histories) === JSON.stringify(historiesData);

      if (isSame) {
        skippedRecords.push(item._id);
        continue;
      }

      const existingSyncUser = existingRouting?.sync_user
        ? parseJsonObject(existingRouting.sync_user, null)
        : null;

      const syncUser = {
        created_by: existingSyncUser?.created_by || currentUser,
        created_at: existingSyncUser?.created_at || nowIso,
        updated_by: currentUser,
        updated_at: nowIso,
      };

      const routingData = {
        hub_id: targetHubId,
        name: item.name || '',
        delivery_date: deliveryDate,
        route_user: routeUser,
        total_distance: totalDistanceMap,
        time: timeMap,
        hub_times: hubTimesMap,
        dropped: droppedData,
        histories: historiesData,
        sync_user: syncUser,
      };

      if (existingRouting) {
        await prisma.routingResult.update({
          where: { id: String(item._id) },
          data: routingData,
        });
        updatedRecords.push(item._id);
      } else {
        await prisma.routingResult.create({
          data: {
            id: String(item._id),
            ...routingData,
          },
        });
        savedRecords.push(item._id);
      }
    }

    return NextResponse.json({
      message: `Berhasil sinkronisasi routing: ${savedRecords.length} baru, ${updatedRecords.length} diperbarui, ${skippedRecords.length} dilewati (tidak berubah).`,
      syncedCount: savedRecords.length + updatedRecords.length,
      insertedCount: savedRecords.length,
      updatedCount: updatedRecords.length,
      skippedCount: skippedRecords.length,
      ids: [...savedRecords, ...updatedRecords],
    });
  } catch (error) {
    console.error('Error Sync Routing Results:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      { error: 'Gagal sinkronisasi data routing ke database', detail: errorMessage },
      { status: 500 }
    );
  }
}
