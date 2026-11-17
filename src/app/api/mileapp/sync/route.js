import { POST as syncResults } from '@/app/api/mileapp/routing-results/route';
import { POST as syncTasks } from '@/app/api/mileapp/tasks/route';
import { POST as syncLocationHistories } from '@/app/api/mileapp/drivers/location-histories/route';
import prisma from '@/lib/prisma';
import { syncVehicleUsageForDate } from '@/lib/services/vehicleUsageService';
import { NextResponse } from 'next/server';

const MS_PER_DAY = 1000 * 60 * 60 * 24;

const ENTITY_SYNC_STEPS = [
  {
    key: 'results',
    match: (entities) => entities.includes('results'),
    handler: syncResults,
  },
  {
    key: 'locationHistories',
    match: (entities) =>
      entities.includes('location-histories') || entities.includes('locationHistories'),
    handler: syncLocationHistories,
  },
  {
    key: 'tasks',
    match: (entities) => entities.includes('tasks'),
    handler: syncTasks,
  },
];

function getDaysFromToday(dateStr) {
  const [y, m, d] = String(dateStr).split('-').map(Number);
  if (!y || !m || !d) return null;

  const now = new Date();
  now.setHours(0, 0, 0, 0);

  const target = new Date(y, m - 1, d);
  target.setHours(0, 0, 0, 0);

  return Math.round((now.getTime() - target.getTime()) / MS_PER_DAY);
}

async function resolveSkipDecision({ deliveryDate, hubId, force }) {
  if (force) return null;

  const diffDays = getDaysFromToday(deliveryDate);
  if (diffDays === null || diffDays <= 2) return null;

  const isGlobal = !hubId || hubId === 'all';

  if (isGlobal) {
    const [activeHubs, tasks, kpiRecords] = await Promise.all([
      prisma.hub.findMany({
        where: { is_active: true },
        select: { id: true },
      }),
      prisma.task.findMany({
        where: { delivery_date: deliveryDate },
        select: { hub_id: true, status: true },
      }),
      prisma.kpiData.findMany({
        where: { date: deliveryDate },
        select: { hub_id: true },
      }),
    ]);

    const hasOngoing = tasks.some((t) => t.status === 'ONGOING');
    if (hasOngoing) return null;

    const recordedHubIds = new Set([
      ...tasks.map((t) => t.hub_id),
      ...kpiRecords.map((k) => k.hub_id),
    ]);
    const allHubsPresent =
      activeHubs.length > 0 && activeHubs.every((h) => recordedHubIds.has(h.id));

    if (!allHubsPresent) return null;

    return {
      message: `Sinkronisasi tanggal ${deliveryDate} dilewati karena seluruh cabang aktif sudah lengkap dan berstatus selesai (DONE).`,
      hubId: 'all',
    };
  }

  const tasks = await prisma.task.findMany({
    where: { delivery_date: deliveryDate, hub_id: hubId },
    select: { status: true },
  });

  const hasOngoing = tasks.some((t) => t.status === 'ONGOING');
  if (tasks.length === 0 || hasOngoing) return null;

  return {
    message: `Sinkronisasi tanggal ${deliveryDate} dilewati karena seluruh tasks cabang ini sudah selesai (DONE).`,
    hubId,
  };
}

async function executeEntitySync(handler, createReq) {
  try {
    const res = await handler(createReq());
    const data = await res.json();
    return {
      ok: res.ok,
      result: {
        status: res.status,
        success: res.ok,
        ...data,
      },
    };
  } catch (err) {
    return {
      ok: false,
      result: {
        status: 500,
        success: false,
        error: err.message,
      },
    };
  }
}

export async function GET() {
  return NextResponse.json({
    message:
      'Endpoint sinkronisasi operasional MileApp. Gunakan POST untuk melakukan sinkronisasi.',
  });
}

export async function POST(request) {
  try {
    const body = await request.json().catch(() => ({}));
    const { deliveryDate, hubId, force = false, created_by, updated_by, userName } = body || {};

    if (!deliveryDate) {
      return NextResponse.json(
        { error: 'Parameter "deliveryDate" (format YYYY-MM-DD) dibutuhkan.' },
        { status: 400 }
      );
    }

    const skipDecision = await resolveSkipDecision({ deliveryDate, hubId, force });
    if (skipDecision) {
      return NextResponse.json({
        message: skipDecision.message,
        success: true,
        skipped: true,
        deliveryDate,
        hubId: skipDecision.hubId,
      });
    }

    const payload = {
      deliveryDate,
      hubId,
      force,
      created_by: created_by || userName,
      updated_by: updated_by || userName,
      userName: userName || updated_by || created_by,
    };

    const createReq = () =>
      new Request('http://localhost/api/mileapp/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

    const entities =
      Array.isArray(body.entities) && body.entities.length > 0
        ? body.entities
        : ['results', 'location-histories', 'tasks'];

    const summary = {
      results: null,
      locationHistories: null,
      tasks: null,
    };

    let hasError = false;
    let successCount = 0;

    for (const step of ENTITY_SYNC_STEPS) {
      if (!step.match(entities)) continue;
      const { ok, result } = await executeEntitySync(step.handler, createReq);
      summary[step.key] = result;
      if (ok) successCount++;
      else hasError = true;
    }

    if (summary.tasks?.success !== false) {
      try {
        await syncVehicleUsageForDate(hubId, deliveryDate, payload.userName || 'system');
      } catch (vhErr) {
        console.warn(`[Sync] Gagal generate daily vehicle usage untuk ${deliveryDate}:`, vhErr.message);
      }
    }

    const totalEntities = Object.values(summary).filter(Boolean).length;
    const allFailed = totalEntities > 0 && successCount === 0;

    return NextResponse.json(
      {
        message: allFailed
          ? 'Gagal sinkronisasi data operasional.'
          : hasError
            ? 'Sinkronisasi selesai dengan beberapa catatan.'
            : 'Berhasil sinkronisasi seluruh data operasional.',
        success: !allFailed,
        partial: hasError && !allFailed,
        deliveryDate,
        hubId: hubId || null,
        ...summary,
      },
      { status: allFailed ? 500 : 200 }
    );
  } catch (error) {
    console.error('Error Master Sync:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      {
        error: 'Gagal memproses sinkronisasi operasional',
        detail: errorMessage,
      },
      { status: 500 }
    );
  }
}
