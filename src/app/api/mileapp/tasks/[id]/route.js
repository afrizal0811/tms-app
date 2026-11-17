import prisma from '@/lib/prisma';
import { formatMinutesToHHMM, formatUTC7, getBasePlate } from '@/lib/utils';
import { getMileAppUserMap } from '../../help';
import {
  buildDriversByEmailMap,
  formatPendingTask,
  formatTask,
  parseDurationToMinutes,
  TASK_TIMESTAMP_FIELDS,
} from '../taskFormatter';
import { NextResponse } from 'next/server';

const STRIPPED_VENDOR_FIELDS = [
  'assignedVehicle',
  'parentId',
  'subId',
  'priority',
  'totalSo',
  'originalWeight',
  'originalVolume',
  'expectedVehicle',
  'vehicleCapacity',
  'groupVisit',
  'orderIndex',
  'organization',
  'content',
  'title',
  'flowId',
  'organizationId',
  'hub',
  'isDeleted',
  'taskType',
  'assignedTo',
  'visitGroup',
  'visitGroupPriority',
  'appVersion',
  'deviceInformation',
  'createdCoordinate',
  'doneFrom',
  'expectedDoneTime',
  'geoLock',
  'geoLockTrigger',
  'idLocal',
  'idleDuration',
  'inLocation',
  'isLastDoneOrder',
  'label',
  'namaPenerimaPic',
  'openTaskTime',
  'outLocation',
  'photoDelivery_0_id',
  'photoDelivery_0_path',
  'photoDelivery_0_url',
  'tandaTanganPenerima_path',
  'tandaTanganPenerima_url',
  'useGeoLock',
  'workflow',
];

async function calculatePlanTravelDuration(data) {
  if (!data?.routingResultId || data.routePlannedOrder == null || !data.eta) return null;

  const plateKey = data.basePlat || getBasePlate(data.plat) || data.plat;
  let refEtd = null;

  if (data.routePlannedOrder > 1) {
    const prevTask = await prisma.task.findFirst({
      where: {
        routing_result_id: data.routingResultId,
        route_planned_order: data.routePlannedOrder - 1,
        ...(plateKey ? { plat: { contains: plateKey } } : {}),
      },
      select: { etd: true },
    });
    refEtd = prevTask?.etd || null;
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
    refEtd = firstHub?.etd || null;
  }

  if (!refEtd) return null;
  const [pH, pM] = refEtd.split(':').map(Number);
  const [eH, eM] = data.eta.split(':').map(Number);
  if (Number.isNaN(pH) || Number.isNaN(pM) || Number.isNaN(eH) || Number.isNaN(eM)) return null;

  let diff = eH * 60 + eM - (pH * 60 + pM);
  if (diff < 0) diff += 24 * 60;
  return diff;
}

export async function GET(request, { params }) {
  try {
    const { searchParams } = new URL(request.url);
    const resolvedParams = await params;
    const id = resolvedParams?.id;
    const hubId = searchParams.get('hubId');
    const vendorOnly = searchParams.get('vendorOnly') === '1';

    if (!id) {
      return NextResponse.json(
        { error: 'Parameter "id" sangat dibutuhkan untuk mengambil data spesifik' },
        { status: 400 }
      );
    }

    const apiUrl = process.env.NEXT_PUBLIC_API_URL;
    const apiToken = process.env.API_TOKEN;

    if (vendorOnly) {
      if (!apiUrl || !apiToken) {
        return NextResponse.json(
          { error: 'Kesalahan konfigurasi server internal' },
          { status: 500 }
        );
      }
      const externalRes = await fetch(`${apiUrl}/task/${id}`, {
        headers: {
          Authorization: `Bearer ${apiToken}`,
          'Content-Type': 'application/json',
        },
        cache: 'no-store',
      });
      const json = await externalRes.json();
      const vendorTask = json ? json.task : json;
      if (!externalRes.ok || !vendorTask) {
        return NextResponse.json(
          { error: `Gagal mengambil data vendor untuk task ${id}`, details: vendorTask },
          { status: externalRes.ok ? 404 : externalRes.status }
        );
      }
      return NextResponse.json({
        histories: Array.isArray(vendorTask.histories) ? vendorTask.histories : [],
        waitingTime: vendorTask.waitingTime != null ? Number(vendorTask.waitingTime) : null,
        source: 'vendor',
      });
    }

    const dbTask = await prisma.task.findUnique({
      where: { id },
      include: {
        list_products: { orderBy: { id: 'asc' } },
        pending_detail: true,
      },
    });

    if (dbTask) {
      const localData = formatTask(dbTask);
      const travelMins = parseDurationToMinutes(dbTask.travel_duration);
      if (travelMins != null) {
        localData.travelDuration = Number(travelMins.toFixed(2));
        localData.travelDurationHour = formatMinutesToHHMM(travelMins, true);
      } else {
        localData.travelDuration = null;
        localData.travelDurationHour = null;
      }

      if (dbTask.pending_detail) {
        localData.pendingDetail = formatPendingTask(dbTask.pending_detail);
      }

      const plannedTravelDuration = await calculatePlanTravelDuration(localData);
      if (plannedTravelDuration != null) {
        localData.plannedTravelDuration = plannedTravelDuration;
        localData.plannedTravelDurationHour = formatMinutesToHHMM(plannedTravelDuration);
      }

      localData.source = 'db';
      return NextResponse.json(localData);
    }

    if (!apiUrl || !apiToken) {
      return NextResponse.json({ error: 'Kesalahan konfigurasi server internal' }, { status: 500 });
    }

    const externalUrl = new URL(`${apiUrl}/task/${id}`);

    const [externalResponse, userMap, driversByEmail] = await Promise.all([
      fetch(externalUrl.toString(), {
        headers: {
          Authorization: `Bearer ${apiToken}`,
          'Content-Type': 'application/json',
        },
      }),
      getMileAppUserMap(),
      buildDriversByEmailMap(hubId),
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

    const mapUserName = (email) => {
      if (!email) return null;
      const cleanEmail = String(email).toLowerCase().trim();
      return driversByEmail[cleanEmail]?.name || userMap.get(cleanEmail) || email;
    };

    data.createdByName = mapUserName(data.createdBy);
    data.updatedByName = mapUserName(data.updatedBy);
    data.assignedByName = mapUserName(data.assignedBy);
    data.doneByName = mapUserName(data.doneBy);

    for (const field of TASK_TIMESTAMP_FIELDS) {
      if (data[field]) data[field] = formatUTC7(data[field], 'YYYY-MM-DDTHH:mm:ss');
    }

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

    for (const field of STRIPPED_VENDOR_FIELDS) {
      delete data[field];
    }

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

    const plannedTravelDuration = await calculatePlanTravelDuration(data);
    if (plannedTravelDuration != null) {
      data.plannedTravelDuration = plannedTravelDuration;
      data.plannedTravelDurationHour = formatMinutesToHHMM(plannedTravelDuration);
    }

    data.source = 'vendor';
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
