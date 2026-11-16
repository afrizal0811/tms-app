import prisma from '@/lib/prisma';
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

    const data = await prisma.dailyFleet.findMany({
      where,
      orderBy: [{ date: 'asc' }, { driver_name: 'asc' }, { plat: 'asc' }],
    });

    return NextResponse.json(data, { status: 200 });
  } catch (error) {
    return NextResponse.json(
      { error: 'Gagal mengambil data Daily Fleet', detail: error.message },
      { status: 500 }
    );
  }
}
