import { POST as syncOperational } from '@/app/api/mileapp/sync/route';
import { processKpiForDates } from '@/lib/services/kpiService';
import { NextResponse } from 'next/server';

export const dynamic = 'force-dynamic';
export const maxDuration = 60;

function getJakartaDateString(offsetDays = 0) {
  const jakartaNow = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Jakarta',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(new Date());

  const [year, month, day] = jakartaNow.split('-').map(Number);
  const target = new Date(Date.UTC(year, month - 1, day + offsetDays));
  return target.toISOString().slice(0, 10);
}

function isAuthorized(request) {
  const secret = process.env.CRON_SECRET;
  if (!secret) return { ok: false, status: 500, error: 'CRON_SECRET belum dikonfigurasi' };

  const authHeader = request.headers.get('authorization') || '';
  if (authHeader !== `Bearer ${secret}`) {
    return { ok: false, status: 401, error: 'Unauthorized' };
  }

  return { ok: true };
}

async function handleCronSync(request) {
  const auth = isAuthorized(request);
  if (!auth.ok) {
    return NextResponse.json({ error: auth.error }, { status: auth.status });
  }

  try {
    const { searchParams } = new URL(request.url);
    const hubId = searchParams.get('hubId') || 'all';
    const range = searchParams.get('range') || 'full';
    const force = searchParams.get('force') === 'true';

    const offsets = range === 'light' ? [0, 1] : [-3, -2, -1, 0, 1];
    const dates = offsets.map((offset) => getJakartaDateString(offset));
    const pastDates = dates.filter((d) => d < getJakartaDateString(0));

    const syncResults = [];
    for (const deliveryDate of dates) {
      const req = new Request('http://localhost/api/mileapp/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          deliveryDate,
          hubId,
          force,
          created_by: 'SYSTEM',
          updated_by: 'SYSTEM',
        }),
      });

      const res = await syncOperational(req);
      const data = await res.json().catch(() => ({}));
      syncResults.push({
        deliveryDate,
        status: res.status,
        success: res.ok && data?.success !== false,
        skipped: Boolean(data?.skipped),
        message: data?.message || data?.error || null,
      });
    }

    let kpiResult = null;
    if (pastDates.length > 0) {
      kpiResult = await processKpiForDates({
        hubId,
        deliveryDates: pastDates,
        userName: 'SYSTEM',
      }).catch((err) => ({ error: err.message }));
    }

    const hasFailure = syncResults.some((r) => !r.success);
    return NextResponse.json(
      {
        success: !hasFailure,
        hubId,
        range,
        dates,
        syncResults,
        kpiResult,
      },
      { status: hasFailure ? 207 : 200 }
    );
  } catch (error) {
    console.error('[Cron Sync Error]:', error);
    return NextResponse.json(
      {
        success: false,
        error: 'Gagal menjalankan cron sync',
        detail: error instanceof Error ? error.message : 'Unknown error',
      },
      { status: 500 }
    );
  }
}

export async function GET(request) {
  return handleCronSync(request);
}

export async function POST(request) {
  return handleCronSync(request);
}
