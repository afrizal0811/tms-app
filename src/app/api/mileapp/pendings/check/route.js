import { DELIVERY_STATUS_KEYS, normalizeDeliveryStatus } from '@/lib/deliveryStatus';
import prisma from '@/lib/prisma';
import { formatDateUniversal, normalizeEmail } from '@/lib/utils';
import { NextResponse } from 'next/server';

export const dynamic = 'force-dynamic';

const TARGET_STATUSES = [
  'BATAL',
  'CANCEL',
  'TERIMA SEBAGIAN',
  'PARTIAL',
  'PENDING',
  'PENDING GR',
  'PENDING_GR',
];

const checkCache = new Map();
const CACHE_TTL_MS = 5 * 60 * 1000;

export function clearPendingCheckCache() {
  checkCache.clear();
}

function getActiveMonthDateRange() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);
  return [start, end];
}

export async function GET(request) {
  try {
    const { searchParams } = new URL(request.url);
    const hubId = searchParams.get('hubId');
    const customStart = searchParams.get('startDate');
    const customEnd = searchParams.get('endDate');

    let startDate, endDate;
    if (customStart && customEnd) {
      startDate = new Date(customStart);
      endDate = new Date(customEnd);
      endDate.setHours(23, 59, 59, 999);
    } else {
      [startDate, endDate] = getActiveMonthDateRange();
    }

    const startStr = formatDateUniversal(startDate, 'YYYY-MM-DD');
    const endStr = formatDateUniversal(endDate, 'YYYY-MM-DD');

    const cacheKey = `${hubId || 'all'}_${startStr}_${endStr}`;
    const cached = checkCache.get(cacheKey);
    if (cached && Date.now() - cached.timestamp < CACHE_TTL_MS) {
      return NextResponse.json({ count: cached.count }, { status: 200 });
    }

    const driverListWhere = hubId && hubId !== 'all' ? { hub_id: hubId } : {};
    let rawDrivers = await prisma.driverList.findMany({ where: driverListWhere });
    if (rawDrivers.length === 0) {
      const oldDriverWhere = hubId && hubId !== 'all' ? { hubs: { some: { id: hubId } } } : {};
      rawDrivers = await prisma.driver.findMany({ where: oldDriverWhere });
    }

    const validDriverEmails = new Set();
    rawDrivers.forEach((d) => {
      if (d.email) {
        validDriverEmails.add(normalizeEmail(d.email));
      }
    });

    if (validDriverEmails.size === 0) {
      checkCache.set(cacheKey, { count: 0, timestamp: Date.now() });
      return NextResponse.json({ count: 0 }, { status: 200 });
    }

    const taskWhere = {
      delivery_date: {
        gte: startStr,
        lte: endStr,
      },
    };
    if (hubId && hubId !== 'all') {
      taskWhere.hub_id = hubId;
    }

    const allTasks = await prisma.task.findMany({
      where: taskWhere,
      select: {
        id: true,
        delivery_date: true,
        status: true,
        status_delivery: true,
        assignee: true,
      },
    });

    const targetTaskIds = [];
    allTasks.forEach((task) => {
      let statusDeliveryRaw = task.status_delivery || task.status || '';
      const status = statusDeliveryRaw.split(',')[0].trim().toUpperCase();

      if (!TARGET_STATUSES.includes(status)) return;

      const email = normalizeEmail((task.assignee || '').split(',')[0]);
      if (!validDriverEmails.has(email)) return;

      targetTaskIds.push(task.id);
    });

    if (targetTaskIds.length === 0) {
      checkCache.set(cacheKey, { count: 0, timestamp: Date.now() });
      return NextResponse.json({ count: 0 }, { status: 200 });
    }

    const existingDetails = await prisma.pendingDetail.findMany({
      where: { task_id: { in: targetTaskIds } },
    });

    const detailsMap = new Map();
    existingDetails.forEach((pd) => {
      detailsMap.set(pd.task_id || pd.taskId, pd);
    });

    let uncompletedCount = 0;
    targetTaskIds.forEach((id) => {
      const pd = detailsMap.get(id);
      if (pd && normalizeDeliveryStatus(pd.new_status) === DELIVERY_STATUS_KEYS.SUCCESS) return;

      const hasAnyData =
        pd &&
        Boolean(
          ((pd.group_reason ?? pd.groupReason) &&
            (pd.group_reason ?? pd.groupReason).trim() !== '' &&
            (pd.group_reason ?? pd.groupReason) !== '-') ||
            ((pd.detail_reason ?? pd.detailReason) &&
              (pd.detail_reason ?? pd.detailReason).trim() !== '' &&
              (pd.detail_reason ?? pd.detailReason) !== '-') ||
            ((pd.internal_external ?? pd.internalExternal) &&
              (pd.internal_external ?? pd.internalExternal).trim() !== '' &&
              (pd.internal_external ?? pd.internalExternal) !== '-') ||
            (pd.pic && pd.pic.trim() !== '' && pd.pic !== '-')
        );

      if (!hasAnyData) {
        uncompletedCount++;
      }
    });

    checkCache.set(cacheKey, { count: uncompletedCount, timestamp: Date.now() });
    return NextResponse.json({ count: uncompletedCount }, { status: 200 });
  } catch (error) {
    console.error('Error Get Pending Details Check:', error);
    return NextResponse.json(
      { error: 'Gagal mengecek pending details', detail: error.message },
      { status: 500 }
    );
  }
}
