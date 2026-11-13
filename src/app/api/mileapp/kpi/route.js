import prisma from '@/lib/prisma';
import { processKpiForDates } from '@/lib/services/kpiService';
import { NextResponse } from 'next/server';

export const dynamic = 'force-dynamic';

export async function GET(request) {
  try {
    const { searchParams } = new URL(request.url);
    const hubId = searchParams.get('hubId');
    const date = searchParams.get('date') || searchParams.get('deliveryDate');
    const dateFrom = searchParams.get('dateFrom') || searchParams.get('startDate');
    const dateTo = searchParams.get('dateTo') || searchParams.get('endDate');

    const where = {};
    if (hubId && hubId !== 'all') {
      where.hub_id = hubId;
    }
    if (date) {
      where.date = date;
    } else if (dateFrom && dateTo) {
      where.date = { gte: dateFrom, lte: dateTo };
    } else if (dateFrom) {
      where.date = { gte: dateFrom };
    }

    const data = await prisma.kpiData.findMany({
      where,
      orderBy: [{ date: 'asc' }, { hub_id: 'asc' }],
    });

    return NextResponse.json(data, { status: 200 });
  } catch (error) {
    return NextResponse.json(
      { error: 'Gagal mengambil data KPI', detail: error.message },
      { status: 500 }
    );
  }
}

export async function POST(request) {
  try {
    const body = await request.json().catch(() => ({}));
    const { hubId, deliveryDate, deliveryDates, userName } = body || {};

    const dates = deliveryDates || (deliveryDate ? [deliveryDate] : []);
    if (dates.length === 0) {
      return NextResponse.json(
        { error: 'Parameter deliveryDate atau deliveryDates dibutuhkan.' },
        { status: 400 }
      );
    }

    const result = await processKpiForDates({
      hubId: hubId || 'all',
      deliveryDates: dates,
      userName: userName || 'system',
    });

    return NextResponse.json(result, { status: 200 });
  } catch (error) {
    return NextResponse.json(
      { error: 'Gagal memproses data KPI', detail: error.message },
      { status: 500 }
    );
  }
}

