'use client';

import SelectionLayout from '@/components/page/SelectionLayout';
import Spinner from '@/components/Spinner';
import VehicleModal from '@/components/modal/VehicleModal';
import PendingReasonModal from '@/features/summary/tabs/modals/PendingReasonModal';
import { useLanguage } from '@/context/LanguageContext';
import {
  getKpiData,
  getReasons,
  getVehicles,
  processKpiData,
  syncOperationalData,
} from '@/lib/api/mileapp';
import { getLocalStorage } from '@/lib/localStorageHandler';
import { toastWarning } from '@/lib/toast';
import { formatDateUniversal, getBasePlate, tomorrowDate } from '@/lib/utils';
import { useEffect, useRef, useState } from 'react';

function withTimeout(promise, ms = 15000) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`Timeout after ${ms}ms`)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

function buildSyncDates(now = new Date()) {
  const todayStr = formatDateUniversal(now, 'YYYY-MM-DD');

  const pastDates = [];
  for (let i = 3; i >= 1; i--) {
    const d = new Date(now);
    d.setDate(d.getDate() - i);
    pastDates.push(formatDateUniversal(d, 'YYYY-MM-DD'));
  }

  const tomorrow = new Date(now);
  tomorrow.setDate(now.getDate() + 1);

  return {
    todayStr,
    pastDates,
    yesterdayStr: pastDates[pastDates.length - 1],
    tomorrowStr: formatDateUniversal(tomorrow, 'YYYY-MM-DD'),
    nextDeliveryStr: formatDateUniversal(tomorrowDate(true), 'YYYY-MM-DD'),
  };
}

function buildEntityPlan(datesByEntity) {
  const plan = new Map();
  for (const [entity, dates] of Object.entries(datesByEntity)) {
    dates.forEach((dStr) => {
      if (!plan.has(dStr)) plan.set(dStr, []);
      const list = plan.get(dStr);
      if (!list.includes(entity)) list.push(entity);
    });
  }
  return Array.from(plan.entries());
}

async function fetchMissingPendingTasks(hubId, dateStr, user) {
  if (!hubId || hubId === 'all' || !dateStr) return [];
  try {
    const res = await processKpiData({
      hubId,
      deliveryDate: dateStr,
      userName: user?.name || 'system',
    });
    if (res?.status === 'MISSING_PENDING_REASONS' && res?.missingTasks?.length > 0) {
      return res.missingTasks.filter((task) => !task.hubId || String(task.hubId) === String(hubId));
    }
  } catch (err) {
    console.warn('[DataSync] KPI process error:', err.message);
  }
  return [];
}

export default function DataSync({ children }) {
  const { t } = useLanguage();
  const tRef = useRef(t);
  tRef.current = t;

  const isSyncingRef = useRef(false);
  const hasRunRef = useRef(false);
  const [syncStatus, setSyncStatus] = useState(null);
  const [syncStartTime, setSyncStartTime] = useState(null);
  const [unmappedVehicles, setUnmappedVehicles] = useState([]);
  const [uncompletedPendingTasks, setUncompletedPendingTasks] = useState([]);
  const [pendingReasonsList, setPendingReasonsList] = useState([]);
  const pendingDatesRef = useRef([]);

  const checkUncompletedVehicles = async (hubId) => {
    if (!hubId || hubId === 'all') return;
    try {
      const vehicles = await getVehicles(hubId);
      if (!Array.isArray(vehicles)) return;

      const unmappedMap = new Map();
      vehicles.forEach((v) => {
        const hasStorage = Boolean(v.storage && String(v.storage).trim() !== '');
        const rawType = v._rawType || (v.type ? v.type.replace(/^(Dry|Frozen)-/i, '') : '');
        const hasType = Boolean(rawType && String(rawType).trim() !== '');

        if (hasStorage && hasType) return;

        const base = v.base_plate || getBasePlate(v.license_number || v.plat);
        if (!base) return;

        const existing = unmappedMap.get(base);
        if (!existing) {
          unmappedMap.set(base, {
            base_plate: base,
            license_number: base,
            vehicle_ids: v.vehicle_id ? [v.vehicle_id] : [],
            storage: v.storage || '',
            type: rawType || '',
          });
        } else if (v.vehicle_id && !existing.vehicle_ids.includes(v.vehicle_id)) {
          existing.vehicle_ids.push(v.vehicle_id);
        }
      });

      setUnmappedVehicles(Array.from(unmappedMap.values()));
    } catch (err) {
      console.warn('[DataSync] Check uncompleted vehicles error:', err.message);
    }
  };

  useEffect(() => {
    const { storedLocation } = getLocalStorage();
    if (storedLocation && storedLocation !== 'all') {
      checkUncompletedVehicles(storedLocation);
    }
  }, []);

  useEffect(() => {
    const runSync = async () => {
      if (typeof window === 'undefined' || isSyncingRef.current || hasRunRef.current) return;

      const { storedUser, storedLocation } = getLocalStorage();
      if (!storedUser || !storedLocation) return;

      hasRunRef.current = true;
      isSyncingRef.current = true;

      const { todayStr, pastDates, yesterdayStr, tomorrowStr, nextDeliveryStr } = buildSyncDates();

      const targetHubId = storedLocation;
      const syncKey = `tms_last_sync_date_${targetHubId}`;
      const isInitialSync = localStorage.getItem(syncKey) !== todayStr;

      if (isInitialSync) {
        setSyncStartTime(Date.now());
        setSyncStatus(`${tRef.current('common.sync')}...`);
      }

      try {
        const entityPlan = buildEntityPlan({
          results: isInitialSync ? [yesterdayStr, todayStr, tomorrowStr] : [todayStr, tomorrowStr],
          'location-histories': [...pastDates, todayStr],
          tasks: isInitialSync ? [...pastDates, todayStr, nextDeliveryStr] : [todayStr, tomorrowStr],
        });

        await Promise.allSettled(
          entityPlan.map(([deliveryDate, entities]) =>
            withTimeout(
              syncOperationalData({ deliveryDate, hubId: targetHubId, entities }),
              entities.length * 15000
            ).catch((err) =>
              console.warn(`[DataSync] Sync error (${deliveryDate}):`, err.message)
            )
          )
        );

        localStorage.setItem(syncKey, todayStr);
        localStorage.setItem('tms_last_sync_date', todayStr);

        if (isInitialSync) {
          const existingKpiList = await getKpiData({
            hubId: targetHubId,
            dateFrom: pastDates[0],
            dateTo: yesterdayStr,
          }).catch(() => []);

          const existingDates = new Set(
            (existingKpiList || [])
              .filter(
                (k) => (k.total_truck || 0) > 0 || (k.task_dry || 0) > 0 || (k.task_frz || 0) > 0
              )
              .map((k) => k.date)
          );

          const allMissingTasks = [];
          const missingDates = [];

          for (const dStr of pastDates) {
            if (!existingDates.has(dStr)) {
              const tasksForDate = await fetchMissingPendingTasks(targetHubId, dStr, storedUser);
              if (tasksForDate.length > 0) {
                allMissingTasks.push(...tasksForDate);
                missingDates.push(dStr);
              }
            }
          }

          if (allMissingTasks.length > 0) {
            pendingDatesRef.current = missingDates;
            setUncompletedPendingTasks(allMissingTasks);
            const reasons = await getReasons().catch(() => []);
            setPendingReasonsList(Array.isArray(reasons) ? reasons : []);
            const template =
              tRef.current?.('summary.tabs.pending_reasons.toast_uncompleted', {
                count: allMissingTasks.length,
              }) || `Terdapat ${allMissingTasks.length} data pending yang belum memiliki alasan!`;
            toastWarning(template.replace('{count}', allMissingTasks.length));
          }
        }

        window.dispatchEvent(new CustomEvent('tms-pending-details-updated'));
        localStorage.setItem('tms_pending_details_last_update', Date.now().toString());
      } catch (err) {
        console.warn('[DataSync] Sync unexpected error:', err.message);
      } finally {
        if (isInitialSync) {
          setSyncStatus(null);
          setSyncStartTime(null);
        }
        isSyncingRef.current = false;
        await checkUncompletedVehicles(targetHubId);
      }
    };

    runSync();

    const handleSessionUpdate = () => {
      if (hasRunRef.current) return;
      const { storedUser, storedLocation } = getLocalStorage();
      if (storedUser && storedLocation) {
        runSync();
      }
    };

    const handleSyncVehiclesCompleted = () => {
      const { storedLocation } = getLocalStorage();
      if (storedLocation) checkUncompletedVehicles(storedLocation);
    };

    window.addEventListener('session_update', handleSessionUpdate);
    window.addEventListener('tms-sync-vehicles-completed', handleSyncVehiclesCompleted);

    return () => {
      window.removeEventListener('session_update', handleSessionUpdate);
      window.removeEventListener('tms-sync-vehicles-completed', handleSyncVehiclesCompleted);
    };
  }, []);

  if (syncStatus) {
    return (
      <SelectionLayout>
        <div className="flex flex-col items-center justify-center gap-4 py-16 select-none animate-in fade-in duration-200">
          <Spinner size="w-16 h-16" showTimer={true} startTime={syncStartTime} />
          <p className="text-sm font-semibold text-slate-400 dark:text-slate-200 italic">
            {syncStatus}
          </p>
        </div>
      </SelectionLayout>
    );
  }

  return (
    <>
      {unmappedVehicles.length > 0 && (
        <VehicleModal
          isOpen={true}
          unmappedData={unmappedVehicles}
          onCompleted={async () => {
            const { storedLocation } = getLocalStorage();
            await checkUncompletedVehicles(storedLocation);
            window.dispatchEvent(new CustomEvent('tms-vehicle-configured'));
          }}
          t={t}
        />
      )}
      {uncompletedPendingTasks.length > 0 && (
        <PendingReasonModal
          isOpen={true}
          noClose={true}
          data={uncompletedPendingTasks}
          reasons={pendingReasonsList}
          onClose={() => setUncompletedPendingTasks([])}
          onSuccess={async () => {
            const { storedLocation, storedUser } = getLocalStorage();
            const targetDates = pendingDatesRef.current;
            if (storedLocation && targetDates.length > 0) {
              try {
                await processKpiData({
                  hubId: storedLocation,
                  deliveryDates: targetDates,
                  userName: storedUser?.name || 'system',
                });
              } catch (err) {
                console.warn('[DataSync] Async KPI retry error:', err.message);
              }
            }
          }}
        />
      )}
      {children || null}
    </>
  );
}
