import prisma from '@/lib/prisma';
import { formatMinutesToHHMM, formatUTC7, getBasePlate } from '@/lib/utils';
import { NextResponse } from 'next/server';

let userCache = {
  data: {},
  expiresAt: 0,
};

async function getMileAppUserMap(apiUrl, apiToken) {
  const now = Date.now();
  if (userCache.expiresAt > now && Object.keys(userCache.data).length > 0) {
    return userCache.data;
  }

  try {
    const usersUrl = new URL(`${apiUrl}/users`);
    usersUrl.searchParams.append('status', 'active');
    usersUrl.searchParams.append('limit', '1000');

    const res = await fetch(usersUrl.toString(), {
      headers: {
        Authorization: `Bearer ${apiToken}`,
        'Content-Type': 'application/json',
      },
    });

    if (res.ok) {
      const json = await res.json();
      const list = json?.data || [];
      const map = {};
      list.forEach((u) => {
        if (u.email) {
          map[u.email.toLowerCase().trim()] = u.name;
        }
      });
      userCache = {
        data: map,
        expiresAt: now + 60 * 60 * 1000,
      };
      return map;
    }
  } catch (err) {
    console.error('Gagal mengambil data users:', err);
  }

  return userCache.data || {};
}

export async function GET(request, { params }) {
  try {
    const { searchParams } = new URL(request.url);
    const resolvedParams = await params;
    const id = resolvedParams?.id;
    const hubId = searchParams.get('hubId');

    if (!id) {
      return NextResponse.json(
        { error: 'Parameter "id" sangat dibutuhkan untuk mengambil data spesifik' },
        { status: 400 }
      );
    }

    const apiUrl = process.env.NEXT_PUBLIC_API_URL;
    const apiToken = process.env.API_TOKEN;

    if (!apiUrl || !apiToken) {
      return NextResponse.json({ error: 'Kesalahan konfigurasi server internal' }, { status: 500 });
    }

    const externalUrl = new URL(`${apiUrl}/task/${id}`);
    const driverWhere = hubId && hubId !== 'all' ? { hub_id: hubId } : {};
    const vehicleWhere = hubId && hubId !== 'all' ? { hubs_id: hubId } : {};

    const [externalResponse, userMap, driverList, vehicleList] = await Promise.all([
      fetch(externalUrl.toString(), {
        headers: {
          Authorization: `Bearer ${apiToken}`,
          'Content-Type': 'application/json',
        },
      }),
      getMileAppUserMap(apiUrl, apiToken),
      prisma.driverList.findMany({ where: driverWhere }),
      prisma.vehicle.findMany({ where: vehicleWhere }),
    ]);

    const fetchData = await externalResponse.json();
    let data = fetchData ? fetchData.task : fetchData;

    if (!externalResponse.ok) {
      return NextResponse.json(
        {
          error: `Gagal mengambil data task dengan ID ${id}`,
          details: data,
        },
        {
          status: externalResponse.status,
        }
      );
    }

    if (!data) {
      return NextResponse.json(
        { error: `Data task dengan ID ${id} tidak ditemukan` },
        { status: 404 }
      );
    }

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

    const mapUserName = (email) => {
      if (!email) return null;
      const cleanEmail = String(email).toLowerCase().trim();
      return driversByEmail[cleanEmail]?.name || userMap[cleanEmail] || email;
    };

    data.createdByName = mapUserName(data.createdBy);
    data.updatedByName = mapUserName(data.updatedBy);
    data.assignedByName = mapUserName(data.assignedBy);
    data.doneByName = mapUserName(data.doneBy);

    if (data.startTime) data.startTime = formatUTC7(data.startTime, 'YYYY-MM-DDTHH:mm:ss');
    if (data.endTime) data.endTime = formatUTC7(data.endTime, 'YYYY-MM-DDTHH:mm:ss');
    if (data.assignedTime) data.assignedTime = formatUTC7(data.assignedTime, 'YYYY-MM-DDTHH:mm:ss');
    if (data.updatedTime) data.updatedTime = formatUTC7(data.updatedTime, 'YYYY-MM-DDTHH:mm:ss');
    if (data.doneTime) data.doneTime = formatUTC7(data.doneTime, 'YYYY-MM-DDTHH:mm:ss');
    if (data.createdTime) data.createdTime = formatUTC7(data.createdTime, 'YYYY-MM-DDTHH:mm:ss');
    if (data.klikJikaSudahSampai)
      data.klikJikaSudahSampai = formatUTC7(data.klikJikaSudahSampai, 'YYYY-MM-DDTHH:mm:ss');
    if (data.page1DoneTime)
      data.page1DoneTime = formatUTC7(data.page1DoneTime, 'YYYY-MM-DDTHH:mm:ss');
    if (data.page2DoneTime)
      data.page2DoneTime = formatUTC7(data.page2DoneTime, 'YYYY-MM-DDTHH:mm:ss');
    if (data.page3DoneTime)
      data.page3DoneTime = formatUTC7(data.page3DoneTime, 'YYYY-MM-DDTHH:mm:ss');

    if (data.volumeCbm != null) data.volumeCbm = Number(Number(data.volumeCbm).toFixed(2));
    if (data.weightKg != null) data.weightKg = Number(Number(data.weightKg).toFixed(2));

    if (data.distance != null) data.distance = Number((data.distance / 1000).toFixed(2));
    if (data.travelDistance != null)
      data.travelDistance = Number((data.travelDistance / 1000).toFixed(2));

    if (data.travelDuration != null)
      data.travelDurationHour = formatMinutesToHHMM(Number(data.travelDuration), true);

    if (data.eta && typeof data.eta === 'string') data.eta = data.eta.substring(0, 5);
    if (data.etd && typeof data.etd === 'string') data.etd = data.etd.substring(0, 5);
    data.travelDuration = data.travelDuration ? Number(data.travelDuration.toFixed(2)) : null;

    delete data.assignedVehicle;
    delete data.parentId;
    delete data.subId;
    delete data.priority;
    delete data.totalSo;
    delete data.originalWeight;
    delete data.originalVolume;
    delete data.expectedVehicle;
    delete data.vehicleCapacity;
    delete data.groupVisit;
    delete data.orderIndex;
    delete data.organization;
    delete data.content;
    delete data.title;
    delete data.flowId;
    delete data.organizationId;
    delete data.hub;
    delete data.isDeleted;
    delete data.taskType;
    delete data.assignedTo;
    delete data.visitGroup;
    delete data.visitGroupPriority;
    delete data.appVersion;
    delete data.deviceInformation;
    delete data.createdCoordinate;
    delete data.doneFrom;
    delete data.expectedDoneTime;
    delete data.geoLock;
    delete data.geoLockTrigger;
    delete data.idLocal;
    delete data.idleDuration;
    delete data.inLocation;
    delete data.isLastDoneOrder;
    delete data.label;
    delete data.namaPenerimaPic;
    delete data.openTaskTime;
    delete data.outLocation;
    delete data.photoDelivery_0_id;
    delete data.photoDelivery_0_path;
    delete data.photoDelivery_0_url;
    delete data.tandaTanganPenerima_path;
    delete data.tandaTanganPenerima_url;
    delete data.useGeoLock;
    delete data.workflow;

    const assigneeEmail =
      data.assignee && data.assignee.length > 0 ? data.assignee[0].toLowerCase().trim() : null;
    const driverData = assigneeEmail ? driversByEmail[assigneeEmail] : null;

    if (driverData) {
      data.driverName = driverData.name;
      data.basePlat = driverData.basePlat;
      data.vehicleType = driverData.type;
      data.maxVolume = driverData.maxVolume;
      data.maxWeight = driverData.maxWeight;
    }

    data.weightPct =
      data.maxWeight && data.weightKg != null
        ? `${((data.weightKg / data.maxWeight) * 100).toFixed(2)}%`
        : '0%';
    data.volumePct =
      data.maxVolume && data.volumeCbm != null
        ? `${((data.volumeCbm / data.maxVolume) * 100).toFixed(2)}%`
        : '0%';

    if (Array.isArray(data.gpsSesuai)) data.gpsSesuai = data.gpsSesuai.join(', ');
    if (Array.isArray(data.statusDelivery)) data.statusDelivery = data.statusDelivery.join(', ');
    if (Array.isArray(data.assignee)) data.assignee = data.assignee.join(', ');
    if (!data.reason && data.alasan) data.reason = data.alasan;

    let plannedTravelDuration = null;
    if (data.routingResultId && data.routePlannedOrder != null && data.eta) {
      const plateKey = data.basePlat || getBasePlate(data.plat) || data.plat;
      if (data.routePlannedOrder > 1) {
        const prevTask = await prisma.task.findFirst({
          where: {
            routing_result_id: data.routingResultId,
            route_planned_order: data.routePlannedOrder - 1,
            ...(plateKey ? { plat: { contains: plateKey } } : {}),
          },
          select: { etd: true },
        });
        if (prevTask?.etd) {
          const [pH, pM] = prevTask.etd.split(':').map(Number);
          const [eH, eM] = data.eta.split(':').map(Number);
          if (!isNaN(pH) && !isNaN(eH)) {
            let diff = eH * 60 + eM - (pH * 60 + pM);
            if (diff < 0) diff += 24 * 60;
            plannedTravelDuration = diff;
          }
        }
      } else if (data.routePlannedOrder === 1) {
        const routingRes = await prisma.routingResult.findUnique({
          where: { id: data.routingResultId },
          select: { hub_times: true },
        });
        const hubTimes =
          typeof routingRes?.hub_times === 'string'
            ? JSON.parse(routingRes.hub_times)
            : routingRes?.hub_times || {};
        const firstHub = (hubTimes[plateKey] || []).find(
          (h) => h.category === 'first' || h.order === 0
        );
        if (firstHub?.etd) {
          const [hH, hM] = firstHub.etd.split(':').map(Number);
          const [eH, eM] = data.eta.split(':').map(Number);
          if (!isNaN(hH) && !isNaN(eH)) {
            let diff = eH * 60 + eM - (hH * 60 + hM);
            if (diff < 0) diff += 24 * 60;
            plannedTravelDuration = diff;
          }
        }
      }
    }

    if (plannedTravelDuration != null) {
      data.plannedTravelDuration = plannedTravelDuration;
      data.plannedTravelDurationHour = formatMinutesToHHMM(plannedTravelDuration);
    }

    return NextResponse.json(data);
  } catch (error) {
    console.error('Error Single Task:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      {
        error: 'Gagal mengambil atau memproses data spesifik dari API eksternal',
        detail: errorMessage,
      },
      { status: 500 }
    );
  }
}
