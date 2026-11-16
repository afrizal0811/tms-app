import prisma from '@/lib/prisma';
import { recalculateDailyFleetForTask } from '@/lib/services/dailyFleetService';
import { NextResponse } from 'next/server';
import { clearPendingCheckCache } from './check/route';

export const dynamic = 'force-dynamic';

function formatPending(p) {
  if (!p || typeof p !== 'object' || !p.id) return p;
  return {
    ...p,
    taskId: p.task_id,
    internalExternal: p.internal_external,
    detailReason: p.detail_reason,
    groupReason: p.group_reason,
    createdAt: p.created_at,
    updatedAt: p.updated_at,
  };
}

export async function GET(request) {
  const { searchParams } = new URL(request.url);
  const taskId = searchParams.get('taskId');
  const startDate = searchParams.get('startDate');
  const endDate = searchParams.get('endDate');

  try {
    if (taskId) {
      const item = await prisma.pendingDetail.findUnique({
        where: { task_id: taskId },
      });
      return NextResponse.json(item ? formatPending(item) : {}, { status: 200 });
    }

    const data = await prisma.pendingDetail.findMany({
      where: {
        date: { gte: startDate, lte: endDate },
      },
    });
    return NextResponse.json(data.map(formatPending), { status: 200 });
  } catch (error) {
    console.error('Error Get Pending Details:', error);
    return NextResponse.json(
      { error: 'Gagal mengambil data detail pending', detail: error.message },
      { status: 500 }
    );
  }
}

export async function POST(request) {
  try {
    const body = await request.json();
    const isBatch = Array.isArray(body) || Array.isArray(body?.items);
    const rawList = Array.isArray(body)
      ? body
      : Array.isArray(body?.items)
        ? body.items
        : [body];

    const taskIds = rawList.map((item) => item.taskId).filter(Boolean);
    const existingList =
      taskIds.length > 0
        ? await prisma.pendingDetail.findMany({
            where: { task_id: { in: taskIds } },
          })
        : [];
    const existingMap = new Map(existingList.map((e) => [e.task_id, e]));

    const upsertOperations = rawList.map((entry) => {
      const {
        taskId,
        date,
        new_status,
        internalExternal,
        detailReason,
        groupReason,
        pic,
        userName,
        created_by,
        updated_by,
      } = entry;

      const currentUser = updated_by || created_by || userName || null;
      const existing = existingMap.get(taskId);

      const updateData = {
        internal_external: internalExternal ?? null,
        detail_reason: detailReason ?? null,
        group_reason: groupReason ?? null,
        pic: pic ?? null,
        updated_by: currentUser,
      };

      if (new_status !== undefined) {
        updateData.new_status = new_status;
      }

      if (existing && !existing.created_by && currentUser) {
        updateData.created_by = currentUser;
      }

      return prisma.pendingDetail.upsert({
        where: { task_id: taskId },
        update: updateData,
        create: {
          task_id: taskId,
          date,
          new_status: new_status || null,
          internal_external: internalExternal || null,
          detail_reason: detailReason || null,
          group_reason: groupReason || null,
          pic: pic || null,
          created_by: currentUser,
          updated_by: currentUser,
        },
      });
    });

    const results = await prisma.$transaction(upsertOperations);
    clearPendingCheckCache();

    const batchUser = rawList.find((e) => e.updated_by || e.created_by || e.userName);
    const resolvedUser = batchUser?.updated_by || batchUser?.created_by || batchUser?.userName || 'system';
    Promise.all(taskIds.map((tId) => recalculateDailyFleetForTask(tId, resolvedUser))).catch((err) =>
      console.warn('[PendingDetails] Gagal recalculate vehicle usage:', err.message)
    );

    return NextResponse.json(
      { message: 'Success', data: isBatch ? results.map(formatPending) : formatPending(results[0]) },
      { status: 200 }
    );
  } catch (error) {
    console.error('Error Upsert Pending Details:', error);
    return NextResponse.json(
      { error: 'Gagal menyimpan data detail pending', detail: error.message },
      { status: 500 }
    );
  }
}

export async function DELETE(request) {
  const { searchParams } = new URL(request.url);
  const taskId = searchParams.get('taskId');

  try {
    if (!taskId) return NextResponse.json({ error: 'taskId dibutuhkan' }, { status: 400 });

    await prisma.pendingDetail.delete({
      where: { task_id: taskId },
    });
    clearPendingCheckCache();

    recalculateDailyFleetForTask(taskId).catch((err) =>
      console.warn('[PendingDetails] Gagal recalculate vehicle usage saat delete:', err.message)
    );

    return NextResponse.json({ message: 'Detail pending berhasil dihapus' }, { status: 200 });
  } catch (error) {
    console.error('Error Delete Pending Details:', error);
    return NextResponse.json(
      { error: 'Gagal menghapus data detail pending', detail: error.message },
      { status: 500 }
    );
  }
}
