import {
  DELIVERY_STATUS_KEYS,
  isTaskUnassigned,
  normalizeDeliveryStatus,
} from '@/lib/deliveryStatus';
import { getStorageType, normalizeEmail } from '@/lib/utils';
import { resolveRouteCategory } from '../../../helper/resolveRouteCategory';

export function calculateVolumeSummary(resultsData, tasksData, driverMap, driverData = []) {
  const dryTrucks = new Set();
  const frzTrucks = new Set();
  let ttDry = 0;
  let ttFrz = 0;
  const countedTaskIds = new Set();

  if (Array.isArray(resultsData)) {
    resultsData.forEach((item) => {
      const plates = Array.from(
        new Set([
          ...Object.keys(item.totalDistance || {}),
          ...Object.keys(item.time || {}),
          ...Object.keys(item.hubTimes || {}),
        ])
      );

      plates.forEach((plat) => {
        const truckId = plat || '-';
        const matchingTasks = Array.isArray(tasksData)
          ? tasksData.filter(
              (t) => !isTaskUnassigned(t) && (t.basePlat || t.vehicleName || t.plat) === truckId
            )
          : [];
        const taskCount = matchingTasks.length;
        if (taskCount === 0) return;
        const tripsCount = taskCount;

        matchingTasks.forEach((t) => {
          if (t.id || t._id) countedTaskIds.add(String(t.id || t._id));
        });

        const route = { basePlat: plat, vehicleName: plat };
        const category = resolveRouteCategory(route, item, tasksData, driverData, driverMap);

        if (truckId) {
          if (category === 'DRY') {
            dryTrucks.add(truckId);
            ttDry += tripsCount;
          } else if (category === 'FROZEN') {
            frzTrucks.add(truckId);
            ttFrz += tripsCount;
          }
        }
      });
    });
  }

  if (Array.isArray(tasksData)) {
    tasksData.forEach((task) => {
      if (isTaskUnassigned(task)) return;
      const taskId = String(task.id || task._id || '');
      if (taskId && countedTaskIds.has(taskId)) return;

      const truckId = task.basePlat || task.vehicleName || task.plat || task.driverName || null;
      let category = (task.typeStorage || task.storage || '').toUpperCase();
      if (!category.includes('DRY') && !category.includes('FRZ') && !category.includes('FROZEN')) {
        const route = { basePlat: truckId, vehicleName: truckId };
        category = resolveRouteCategory(route, null, tasksData, driverData, driverMap);
      }

      if (category.includes('DRY')) {
        if (truckId) dryTrucks.add(truckId);
        ttDry += 1;
        if (taskId) countedTaskIds.add(taskId);
      } else if (category.includes('FRZ') || category.includes('FROZEN')) {
        if (truckId) frzTrucks.add(truckId);
        ttFrz += 1;
        if (taskId) countedTaskIds.add(taskId);
      }
    });
  }

  const countDry = dryTrucks.size;
  const countFrz = frzTrucks.size;
  const countTotal = countDry + countFrz;

  let rtDry = 0;
  let rtFrz = 0;
  const pendingTasks = [];

  if (Array.isArray(tasksData)) {
    tasksData.forEach((task) => {
      if (isTaskUnassigned(task)) return;
      const pd = task.pendingDetail;
      const effectiveStatus = pd?.new_status || task.statusDelivery || task.status;
      const norm = normalizeDeliveryStatus(effectiveStatus);
      const isPending = norm === DELIVERY_STATUS_KEYS.PENDING;

      if (isPending) {
        pendingTasks.push(task);
        let driverName = task.driverName === '-' ? '' : task.driverName || '';
        let category = task.typeStorage === '-' ? '' : task.typeStorage || '';

        if (!driverName && task.assignee) {
          const rawAssignee =
            typeof task.assignee === 'string' ? task.assignee.split(',')[0].trim() : task.assignee;
          driverName = driverMap[normalizeEmail(rawAssignee)] || rawAssignee;
        }

        if (!category && driverName) {
          category = getStorageType(driverName).toUpperCase();
        }

        const catUp = (category || '').toUpperCase();
        if (catUp.includes('DRY')) {
          rtDry++;
        } else if (catUp.includes('FRZ') || catUp.includes('FROZEN')) {
          rtFrz++;
        }
      }
    });
  }

  let rtRouting = null;
  let rtDriver = null;
  let rtSales = null;
  let rtCustomer = null;
  let rtOther = null;

  const hasPendingDetails = pendingTasks.some(
    (t) =>
      t.pendingDetail &&
      Boolean(
        (t.pendingDetail.pic && t.pendingDetail.pic.trim() !== '' && t.pendingDetail.pic !== '-') ||
          (t.pendingDetail.groupReason &&
            t.pendingDetail.groupReason.trim() !== '' &&
            t.pendingDetail.groupReason !== '-') ||
          (t.pendingDetail.detailReason &&
            t.pendingDetail.detailReason.trim() !== '' &&
            t.pendingDetail.detailReason !== '-') ||
          (t.pendingDetail.internalExternal &&
            t.pendingDetail.internalExternal.trim() !== '' &&
            t.pendingDetail.internalExternal !== '-')
      )
  );

  if (hasPendingDetails) {
    rtRouting = 0;
    rtDriver = 0;
    rtSales = 0;
    rtCustomer = 0;
    rtOther = 0;

    pendingTasks.forEach((task) => {
      const rawPic = task.pendingDetail?.pic;
      const pic = typeof rawPic === 'string' ? rawPic.trim().toLowerCase() : '';

      if (pic === 'expedition') {
        rtRouting++;
      } else if (pic === 'driver') {
        rtDriver++;
      } else if (pic === 'sales') {
        rtSales++;
      } else if (pic === 'customer') {
        rtCustomer++;
      } else if (pic.includes('back office') || pic.includes('warehouse') || pic === 'other') {
        rtOther++;
      }
    });
  }

  const hasIncompletePendingDetail = pendingTasks.some((t) => {
    const pd = t.pendingDetail;
    const hasDetail = Boolean(
      pd &&
        ((pd.detailReason && pd.detailReason.trim() !== '' && pd.detailReason !== '-') ||
          (pd.groupReason && pd.groupReason.trim() !== '' && pd.groupReason !== '-') ||
          (pd.pic && pd.pic.trim() !== '' && pd.pic !== '-') ||
          (pd.internalExternal && pd.internalExternal.trim() !== '' && pd.internalExternal !== '-'))
    );
    return !hasDetail;
  });

  return {
    countDry,
    countFrz,
    countTotal,
    ttDry,
    ttFrz,
    rtDry,
    rtFrz,
    rtRouting,
    rtDriver,
    rtSales,
    rtCustomer,
    rtOther,
    hasIncompletePendingDetail,
  };
}
