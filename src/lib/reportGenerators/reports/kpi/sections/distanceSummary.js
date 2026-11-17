import { isTaskUnassigned } from '@/lib/deliveryStatus';
import { resolveRouteCategory } from '../../../helper/resolveRouteCategory';

export function calculateDistanceSummary(
  resultsData,
  historiesData,
  driverMap,
  driverData = [],
  tasksData = []
) {
  let totalActMinutes = 0;

  if (Array.isArray(historiesData)) {
    historiesData.forEach((history) => {
      if (history.startTime && history.finish?.finishTime) {
        const diffMins =
          (new Date(history.finish.finishTime) - new Date(history.startTime)) / 60000;
        if (diffMins > 0) totalActMinutes += diffMins;
      }
    });
  }

  let rawDistDryMeters = 0;
  let rawDistFrzMeters = 0;

  const processedPlates = new Set();
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
        if (Array.isArray(tasksData) && tasksData.length > 0) {
          const hasActiveTask = tasksData.some(
            (t) => !isTaskUnassigned(t) && (t.basePlat || t.vehicleName || t.plat) === plat
          );
          if (!hasActiveTask) return;
        }
        processedPlates.add(plat);
        const route = { basePlat: plat, vehicleName: plat };
        const category = resolveRouteCategory(route, item, tasksData, driverData, driverMap);
        const vehicleDistMeters = Number(item.totalDistance?.[plat] || 0);

        if (vehicleDistMeters > 0) {
          if (category === 'DRY') rawDistDryMeters += vehicleDistMeters;
          else if (category === 'FROZEN') rawDistFrzMeters += vehicleDistMeters;
        }
      });
    });
  }

  if (Array.isArray(tasksData)) {
    tasksData.forEach((task) => {
      const plat = task.basePlat || task.vehicleName || task.plat;
      if (plat && processedPlates.has(plat)) return;

      const taskDist = Number(task.distance ?? task.travelDistance ?? 0);
      if (taskDist > 0) {
        let category = (task.typeStorage || task.storage || '').toUpperCase();
        if (
          !category.includes('DRY') &&
          !category.includes('FRZ') &&
          !category.includes('FROZEN')
        ) {
          const route = { basePlat: plat, vehicleName: plat };
          category = resolveRouteCategory(route, null, tasksData, driverData, driverMap);
        }
        if (category.includes('DRY')) rawDistDryMeters += taskDist * 1000;
        else if (category.includes('FRZ') || category.includes('FROZEN'))
          rawDistFrzMeters += taskDist * 1000;
      }
    });
  }

  return {
    actOperatingHours: Math.round(totalActMinutes / 60),
    estDistanceDry: Number((rawDistDryMeters / 1000).toFixed(2)),
    estDistanceFrz: Number((rawDistFrzMeters / 1000).toFixed(2)),
  };
}
