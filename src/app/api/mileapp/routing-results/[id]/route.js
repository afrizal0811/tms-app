import prisma from '@/lib/prisma';
import { NextResponse } from 'next/server';

export async function GET(request, { params }) {
  try {
    const resolvedParams = await params;
    const id = resolvedParams?.id;

    if (!id) {
      return NextResponse.json(
        { error: 'Parameter "id" sangat dibutuhkan untuk mengambil data spesifik' },
        { status: 400 }
      );
    }

    const item = await prisma.routingResult.findUnique({
      where: { id },
      include: { hub: true },
    });

    if (!item) {
      return NextResponse.json(
        { error: `Data routing dengan ID ${id} tidak ditemukan di database` },
        { status: 404 }
      );
    }

    const hubTimes =
      typeof item.hub_times === 'string'
        ? JSON.parse(item.hub_times || '{}')
        : item.hub_times || {};
    const totalDistance =
      typeof item.total_distance === 'string'
        ? JSON.parse(item.total_distance || '{}')
        : item.total_distance || {};
    const timeData =
      typeof item.time === 'string'
        ? JSON.parse(item.time || '{}')
        : item.time || {};

    const formattedTime = {};
    for (const [p, t] of Object.entries(timeData)) {
      formattedTime[p] = {
        spentTime: Number(t.spentTime ?? t.spent_time) || 0,
        visitTime: Number(t.visitTime ?? t.visit_time) || 0,
        travelTime: Number(t.travelTime ?? t.travel_time) || 0,
        waitingTime: Number(t.waitingTime ?? t.waiting_time) || 0,
      };
    }

    const syncUser =
      typeof item.sync_user === 'string'
        ? JSON.parse(item.sync_user || '{}')
        : item.sync_user || {};
    const routeUser =
      typeof item.route_user === 'string'
        ? JSON.parse(item.route_user || '{}')
        : item.route_user || {};

    const formatted = {
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
      dispatchStatus: 'done',
      dropped: item.dropped,
      histories: item.histories,
      totalDistance,
      time: formattedTime,
      hubTimes,
      hub: item.hub,
    };

    return NextResponse.json(formatted);
  } catch (error) {
    console.error('Error Single Result:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      {
        error: 'Gagal mengambil data result dari database',
        detail: errorMessage,
      },
      { status: 500 }
    );
  }
}

export async function PATCH(request, { params }) {
  try {
    const resolvedParams = await params;
    const id = resolvedParams?.id;

    if (!id) {
      return NextResponse.json(
        { error: 'Parameter "id" dibutuhkan untuk memperbarui data routing' },
        { status: 400 }
      );
    }

    const body = await request.json();
    const { deliveryDate, updatedBy } = body || {};

    if (!deliveryDate) {
      return NextResponse.json(
        { error: 'Parameter "deliveryDate" dibutuhkan.' },
        { status: 400 }
      );
    }

    const existingRouting = await prisma.routingResult.findUnique({
      where: { id },
      select: { sync_user: true },
    });

    const existingSyncUser = existingRouting?.sync_user
      ? typeof existingRouting.sync_user === 'string'
        ? JSON.parse(existingRouting.sync_user)
        : existingRouting.sync_user
      : null;

    const nowIso = new Date().toISOString();
    const syncUser = {
      created_by: existingSyncUser?.created_by || updatedBy || 'System',
      created_at: existingSyncUser?.created_at || nowIso,
      updated_by: updatedBy || 'System',
      updated_at: nowIso,
    };

    const updateData = {
      delivery_date: deliveryDate,
      sync_user: syncUser,
    };

    const updated = await prisma.routingResult.update({
      where: { id },
      data: updateData,
    });

    return NextResponse.json({
      message: 'Berhasil memperbarui data routing',
      data: {
        id: updated.id,
        name: updated.name,
        deliveryDate: updated.delivery_date,
        syncUser,
        sync_user: syncUser,
      },
    });
  } catch (error) {
    console.error('Error updating routing result:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      {
        error: 'Gagal memperbarui data routing',
        detail: errorMessage,
      },
      { status: 500 }
    );
  }
}
