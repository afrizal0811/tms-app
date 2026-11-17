'use client';

import Button from '@/components/button/Button';
import CustomDatePicker from '@/components/CustomDatePicker';
import ConfirmModal from '@/components/modal/ConfirmModal';
import PendingReasonModal from '@/features/summary/tabs/modals/PendingReasonModal';
import { getUsers as getMceasyUsers, patchVehicle as patchMceasyVehicle } from '@/lib/api/mceasy';
import {
  getDrivers,
  getReasons,
  patchDriverMceasy,
  patchVehicleMceasy,
  postDriverList,
  postHubs,
  postRoles,
  postVehicles,
  processKpiData,
  syncOperationalData,
} from '@/lib/api/mileapp';
import { getLocalStorage } from '@/lib/localStorageHandler';
import { toastError, toastSuccess, toastWarning } from '@/lib/toast';
import { formatDateUniversal } from '@/lib/utils';
import { useCallback, useEffect, useMemo, useState } from 'react';
import Card from '../components/Card';
import SyncModal from '../components/SyncModal';

const isSameMismatch = (a, b) =>
  a.id === b.id ||
  (a.mcVehicleId && a.mcVehicleId === b.mcVehicleId) ||
  (a.maPlatBase && b.maPlatBase && a.maPlatBase === b.maPlatBase);

const isMatchSame = (a, b) => a.id === b.id || (a.vms_id && a.vms_id === b.vms_id);

function getDatesInRange(startDate, endDate) {
  if (!startDate || !endDate) return null;
  const start = new Date(startDate);
  const end = new Date(endDate);
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);

  const diffTime = Math.abs(end - start);
  const totalDays = Math.round(diffTime / (1000 * 60 * 60 * 24)) + 1;

  const dates = [];
  const curr = new Date(start);
  while (curr <= end) {
    dates.push(formatDateUniversal(curr, 'YYYY-MM-DD'));
    curr.setDate(curr.getDate() + 1);
  }

  return { totalDays, dates };
}

export default function SyncDataTab({
  hubs,
  lastUpdated,
  updatedBy = {},
  onRefresh,
  isReadOnly,
  translate,
  onOpenDetailTab,
}) {
  const [syncLoading, setSyncLoading] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [mismatchedData, setMismatchedData] = useState([]);
  const [matchedData, setMatchedData] = useState([]);
  const [mcEasyDrivers, setMcEasyDrivers] = useState([]);
  const [isDriverLoading, setIsDriverLoading] = useState(false);
  const [syncScope, setSyncScope] = useState({ mode: 'active' });
  const [routingDateRange, setRoutingDateRange] = useState(() => {
    const today = new Date();
    return [today, today];
  });
  const [confirmModalConfig, setConfirmModalConfig] = useState({
    isOpen: false,
    title: '',
    message: '',
    action: null,
  });
  const [pendingKpiTasks, setPendingKpiTasks] = useState([]);
  const [pendingKpiTarget, setPendingKpiTarget] = useState(null);
  const [kpiReasonsList, setKpiReasonsList] = useState([]);

  const hasCompleteRoutingRange = Boolean(routingDateRange?.[0] && routingDateRange?.[1]);
  const { storedLocation: activeHubId, storedLocationName } = getLocalStorage();

  const allHubs = hubs?.allHub || [];
  const activeHub = hubs?.activeHub || allHubs.find((h) => String(h._id) === String(activeHubId));
  const activeHasVms = Boolean(activeHub?.hasVms);

  const maxRoutingDate = useMemo(() => {
    const d = new Date();
    d.setDate(d.getDate() + 1);
    d.setHours(23, 59, 59, 999);
    return d;
  }, []);

  const handleSuccessReload = async () => {
    toastSuccess(translate('common.toast.success'));
    await onRefresh();
    window.location.reload();
  };

  const closeConfirmModal = () => {
    setConfirmModalConfig((prev) => ({ ...prev, isOpen: false, action: null }));
  };

  const handleConfirmAction = () => {
    const action = confirmModalConfig.action;
    closeConfirmModal();
    if (action) action();
  };

  const fetchMcEasyDrivers = useCallback(async () => {
    setIsDriverLoading(true);
    try {
      const [res, mileappData] = await Promise.all([
        getMceasyUsers({ 'position-name': 'Driver', show: 10000 }),
        getDrivers(syncScope.mode === 'global' ? 'all' : activeHubId),
      ]);

      const validEmails = new Set(
        mileappData.map((m) => m.email?.toLowerCase().trim()).filter(Boolean)
      );

      const rawUsers = res.data || res;
      const filteredDrivers = rawUsers.filter(
        (driver) => driver.email && validEmails.has(driver.email.toLowerCase().trim())
      );

      const seenLabels = new Set();
      const options = [];

      filteredDrivers.forEach((d) => {
        let label = d.fullname;
        if (seenLabels.has(label)) label = `${label} (${d.email})`;
        seenLabels.add(label);
        options.push({ label, value: d.userId });
      });

      options.unshift({ label: translate('setting.tab.modal.blank_driver'), value: '-' });
      setMcEasyDrivers(options);
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setIsDriverLoading(false);
    }
  }, [activeHubId, syncScope.mode, translate]);

  useEffect(() => {
    if (modalOpen) fetchMcEasyDrivers();
  }, [modalOpen, fetchMcEasyDrivers]);

  const handleMismatchedChange = (id, key, value) => {
    setMismatchedData((prev) =>
      prev.map((m) => (m.id === id ? { ...m, [key]: value, isUpdated: true } : m))
    );
  };

  const handleSaveMismatches = async () => {
    setSyncLoading(syncScope.mode === 'global' ? 'drivers-global' : 'drivers-active');
    try {
      const updates = mismatchedData.filter((m) => m.isUpdated);
      const targetDriverIds = new Set(
        updates
          .map((u) => u.updatedDriverId)
          .filter((id) => id && id !== '-')
          .map((id) => String(id))
      );

      const vehiclesToClear = new Set();
      updates.forEach((u) => {
        if (u.mcVehicleId && u.updatedDriverId !== undefined) {
          vehiclesToClear.add(u.mcVehicleId);
        }
      });
      mismatchedData.forEach((m) => {
        if (m.mcVehicleId && m.vmsDriverId && targetDriverIds.has(String(m.vmsDriverId))) {
          vehiclesToClear.add(m.mcVehicleId);
        }
      });

      for (const vehicleId of vehiclesToClear) {
        try {
          const clearPayload = new URLSearchParams();
          clearPayload.append('driver1Id', '-');
          await patchMceasyVehicle(vehicleId, clearPayload.toString());
        } catch (clearErr) {
          console.warn(
            `[Unassign Driver] Gagal melepas driver pada vehicle ${vehicleId}:`,
            clearErr.message
          );
        }
      }

      const seenVehicleIds = new Set();
      for (const update of updates) {
        if (!update.mcVehicleId || seenVehicleIds.has(update.mcVehicleId)) continue;
        seenVehicleIds.add(update.mcVehicleId);

        const payload = new URLSearchParams();
        if (update.updatedDriverId !== undefined) {
          payload.append(
            'driver1Id',
            update.updatedDriverId && update.updatedDriverId !== '-' ? update.updatedDriverId : '-'
          );
        }
        if (update.updatedPlat) {
          payload.append('licensePlate', update.updatedPlat);
        }

        try {
          await patchMceasyVehicle(update.mcVehicleId, payload.toString());
        } catch (err) {
          throw new Error(`[${update.mcPlat}] ${err.message}`);
        }
      }

      const resolvedMatches = updates.map((u) => ({
        id: u.id,
        imei: u.imei,
        vms_id: u.mcVehicleId,
        vms_driver_id:
          u.updatedDriverId !== undefined && u.updatedDriverId !== '-'
            ? u.updatedDriverId
            : u.vmsDriverId,
      }));

      const allMatchesToSave = [...matchedData, ...resolvedMatches];
      if (allMatchesToSave.length > 0) {
        await patchVehicleMceasy(null, null, allMatchesToSave);
      }

      setModalOpen(false);
      await handleSuccessReload();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setSyncLoading(null);
    }
  };

  const handleBypass = async () => {
    setSyncLoading(syncScope.mode === 'global' ? 'vehicles-global' : 'vehicles-active');
    try {
      if (matchedData.length > 0) {
        await patchVehicleMceasy(null, null, matchedData);
      }

      setModalOpen(false);
      await handleSuccessReload();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setSyncLoading(null);
    }
  };

  const executeDriverSyncActive = async () => {
    if (isReadOnly) return;
    setSyncLoading('drivers-active');
    const cleanActiveHubName = (storedLocationName || '').replace(/^Hub\s+/i, '').trim();
    setSyncScope({ mode: 'active', hubId: activeHubId, hubName: cleanActiveHubName });
    try {
      await Promise.all([postDriverList([activeHubId]), postVehicles([activeHubId])]);

      if (activeHasVms) {
        const result = await patchVehicleMceasy(activeHubId, storedLocationName);
        await patchDriverMceasy(activeHubId, storedLocationName).catch((err) =>
          console.warn('[Driver VMS Sync] Active hub error:', err.message)
        );

        if (!result.success && result.mismatched?.length > 0) {
          const dedupedMismatches = [];
          for (const m of result.mismatched) {
            const existing = dedupedMismatches.find((x) => isSameMismatch(x, m));
            if (!existing) {
              dedupedMismatches.push({
                ...m,
                hubName: cleanActiveHubName,
                hubId: activeHubId,
              });
            }
          }

          setMismatchedData(dedupedMismatches);
          setMatchedData(result.matched || []);
          setModalOpen(true);
          setSyncLoading(null);
          return;
        }

        if (result.matched?.length > 0) {
          await patchVehicleMceasy(null, null, result.matched);
        }
      }

      await handleSuccessReload();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setSyncLoading(null);
    }
  };

  const executeDriverSyncGlobal = async () => {
    if (isReadOnly) return;
    setSyncLoading('drivers-global');
    setSyncScope({ mode: 'global' });
    try {
      await Promise.all([postDriverList([]), postVehicles([])]);

      const vmsHubs = allHubs.filter((h) => h.hasVms);
      const allMismatches = [];
      const allMatches = [];

      for (const hub of vmsHubs) {
        const hubNameClean = (hub.name || '').replace(/^Hub\s+/i, '').trim();
        try {
          const result = await patchVehicleMceasy(hub._id, hubNameClean);
          await patchDriverMceasy(hub._id, hubNameClean).catch(() => {});
          if (result.mismatched?.length > 0) {
            for (const m of result.mismatched) {
              const existing = allMismatches.find((x) => isSameMismatch(x, m));
              if (existing) {
                const currentHubs = (existing.hubName || '')
                  .split(',')
                  .map((s) => s.trim())
                  .filter(Boolean);
                if (!currentHubs.includes(hubNameClean)) {
                  currentHubs.push(hubNameClean);
                  existing.hubName = currentHubs.join(', ');
                }
              } else {
                allMismatches.push({
                  ...m,
                  hubName: hubNameClean,
                  hubId: hub._id,
                });
              }
            }
          }
          if (result.matched?.length > 0) {
            for (const match of result.matched) {
              const alreadyMatched = allMatches.some((x) => isMatchSame(x, match));
              if (!alreadyMatched) {
                allMatches.push(match);
              }
            }
          }
        } catch (err) {
          console.warn(`[VMS Sync] Hub ${hub.name} dilewati karena error:`, err.message);
        }
      }

      const filteredMatches = allMatches.filter(
        (match) =>
          !allMismatches.some(
            (m) => m.id === match.id || (m.mcVehicleId && m.mcVehicleId === match.vms_id)
          )
      );

      if (allMismatches.length > 0) {
        setMismatchedData(allMismatches);
        setMatchedData(filteredMatches);
        setModalOpen(true);
        setSyncLoading(null);
        return;
      }

      if (filteredMatches.length > 0) {
        await patchVehicleMceasy(null, null, filteredMatches);
      }

      await handleSuccessReload();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setSyncLoading(null);
    }
  };

  const executeSync = async (type) => {
    if (isReadOnly) return;
    setSyncLoading(type);
    try {
      if (type === 'hubs') {
        const { storedUser } = getLocalStorage();
        await postHubs(storedUser?.name);
      } else if (type === 'roles') {
        await postRoles();
      }

      await handleSuccessReload();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setSyncLoading(null);
    }
  };

  const executeRoutingSync = async (mode = 'active', customDates = null) => {
    if (isReadOnly) return;
    const datesToSync =
      customDates || getDatesInRange(routingDateRange?.[0], routingDateRange?.[1])?.dates;
    if (!datesToSync || datesToSync.length === 0) return;

    const isGlobal = mode === 'global';
    const targetHubId = isGlobal
      ? 'all'
      : activeHub?._id || activeHubId || getLocalStorage().storedLocation;
    setSyncLoading(isGlobal ? 'routing-global' : 'routing-active');
    let hasError = false;
    let lastError = null;

    try {
      for (let i = 0; i < datesToSync.length; i++) {
        const dStr = datesToSync[i];
        try {
          const res = await syncOperationalData({
            deliveryDate: dStr,
            hubId: targetHubId,
          });
          if (res?.partial) {
            hasError = true;
            lastError = new Error(`Sebagian data gagal disinkronkan untuk tanggal ${dStr}`);
          }
        } catch (err) {
          hasError = true;
          lastError = err;
        }
      }

      if (onRefresh) await onRefresh();

      window.dispatchEvent(new CustomEvent('tms-pending-details-updated'));
      localStorage.setItem('tms_pending_details_last_update', Date.now().toString());

      if (hasError) {
        toastError(translate('common.toast.error', { err: lastError?.message }), lastError);
      } else {
        toastSuccess(translate('common.toast.success'));
      }

      const todayStr = formatDateUniversal(new Date(), 'YYYY-MM-DD');
      const hPlusOneDates = (datesToSync || []).filter((d) => d < todayStr);
      if (hPlusOneDates.length > 0) {
        try {
          const { storedUser } = getLocalStorage();
          const kpiRes = await processKpiData({
            hubId: targetHubId,
            deliveryDates: hPlusOneDates,
            userName: storedUser?.name || 'system',
          });
          if (kpiRes?.status === 'MISSING_PENDING_REASONS' && kpiRes?.missingTasks?.length > 0) {
            const tasksToShow =
              !isGlobal && targetHubId && targetHubId !== 'all'
                ? kpiRes.missingTasks.filter(
                    (t) => !t.hubId || String(t.hubId) === String(targetHubId)
                  )
                : kpiRes.missingTasks;

            if (tasksToShow.length > 0) {
              setPendingKpiTasks(tasksToShow);
              setPendingKpiTarget({ hubId: targetHubId, dates: hPlusOneDates });
              const reasons = await getReasons().catch(() => []);
              setKpiReasonsList(Array.isArray(reasons) ? reasons : []);
              const template =
                translate('summary.tabs.pending_reasons.toast_uncompleted', {
                  count: tasksToShow.length,
                }) || `Terdapat ${tasksToShow.length} data pending yang perlu dilengkapi!`;
              toastWarning(template.replace('{count}', tasksToShow.length));
            }
          }
        } catch (err) {
          console.warn('[SyncDataTab] KPI processing error:', err.message);
        }
      }
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setSyncLoading(null);
    }
  };

  const triggerRoutingSync = (mode = 'active') => {
    if (isReadOnly) return;
    const range = getDatesInRange(routingDateRange?.[0], routingDateRange?.[1]);
    if (!range) return;

    const isGlobal = mode === 'global';
    const limitDays = isGlobal ? 7 : 14;

    if (range.totalDays > limitDays) {
      setConfirmModalConfig({
        isOpen: true,
        title: translate('common.modal.data_load_title'),
        message: translate('common.modal.data_load_message', { days: limitDays }),
        action: () => executeRoutingSync(mode, range.dates),
      });
      return;
    }

    executeRoutingSync(mode, range.dates);
  };

  const driverGlobalOptions = [
    {
      id: 'global',
      label: `${translate('common.all')} ${translate('common.branch')}`,
      onClick: executeDriverSyncGlobal,
    },
  ];

  const routingOptions = [
    {
      id: 'global',
      label: `${translate('common.all')} ${translate('common.branch')}`,
      onClick: () => triggerRoutingSync('global'),
      disabled: !hasCompleteRoutingRange,
    },
    {
      id: 'detail',
      label: 'Database',
      onClick: () => {
        if (typeof onOpenDetailTab === 'function') {
          onOpenDetailTab({
            hubId: activeHubId,
            hubName: storedLocationName,
            dateRange: routingDateRange,
          });
        }
      },
      disabled: !hasCompleteRoutingRange,
    },
  ];

  const syncData = [
    {
      type: 'hubs',
      label: translate('setting.tab.sync_data.hubs'),
    },
    {
      type: 'roles',
      label: translate('setting.tab.sync_data.roles'),
    },
    {
      type: 'drivers',
      label: translate('setting.tab.sync_data.driver_vehicles'),
    },
    {
      type: 'mains',
      label: translate('setting.tab.sync_data.main_data'),
    },
  ];

  const renderCardAction = (type) => {
    if (isReadOnly) return null;

    if (type === 'drivers') {
      return (
        <Button
          type="split"
          onClick={executeDriverSyncActive}
          options={driverGlobalOptions}
          isLoading={syncLoading === 'drivers-active' || syncLoading === 'drivers-global'}
          disabled={syncLoading !== null}
          text={translate('setting.tab.button.btn_sync_hub', {
            hub: storedLocationName,
          })}
          size="md"
        />
      );
    }

    if (type === 'mains') {
      return (
        <Button
          type="split"
          onClick={() => triggerRoutingSync('active')}
          options={routingOptions}
          isLoading={syncLoading === 'routing-active' || syncLoading === 'routing-global'}
          disabled={syncLoading !== null || !hasCompleteRoutingRange}
          text={translate('setting.tab.button.btn_sync_hub', {
            hub: storedLocationName,
          })}
          size="md"
        />
      );
    }

    return (
      <Button
        onClick={() => executeSync(type)}
        isLoading={syncLoading === type}
        disabled={syncLoading !== null}
        text={translate('common.synchronize')}
        size="md"
      />
    );
  };

  const renderCardContent = (type, label) => {
    const isDriverType = type === 'drivers';
    const isRoutingType = type === 'mains';

    return (
      <div
        className={`flex h-full flex-col justify-between p-5 bg-slate-50 dark:bg-slate-800 border border-gray-200 dark:border-slate-700 rounded-lg shadow-md dark:shadow-slate-700/40 transition-colors ${
          isReadOnly ? '' : 'min-h-40'
        }`}
      >
        <div className="text-center md:text-left">
          <div className="flex items-center justify-center md:justify-start gap-2 flex-wrap">
            <span className="font-bold text-slate-800 dark:text-slate-200">{label}</span>
            {isDriverType &&
              (activeHasVms ? (
                <span className="px-1.5 py-0.5 text-[10px] font-semibold rounded bg-sky-100 text-sky-700 dark:bg-sky-950/40 dark:text-sky-300 border border-sky-200/80 dark:border-sky-800/80">
                  VMS
                </span>
              ) : (
                <span className="px-1.5 py-0.5 text-[10px] font-medium rounded bg-slate-200/70 text-slate-600 dark:bg-slate-700 dark:text-slate-400">
                  Non-VMS
                </span>
              ))}
          </div>

          {isRoutingType ? (
            <div className="mt-2 text-left">
              <label className="block text-xs text-slate-500 dark:text-slate-400 mb-1">
                {translate('common.delivery_date')}
              </label>
              <CustomDatePicker
                selectsRange={true}
                startDate={routingDateRange[0]}
                endDate={routingDateRange[1]}
                onChange={(dates) => setRoutingDateRange(dates)}
                useCustomRangeFormat={true}
                maxDate={maxRoutingDate}
                disableSunday={false}
                className="w-full text-xs cursor-pointer"
                wrapperClassName="w-full"
                disabled={syncLoading !== null || isReadOnly}
              />
            </div>
          ) : (
            <>
              <div className="text-xs text-slate-500 dark:text-slate-400 mt-1.5">
                {translate('setting.last_updated')} <br />
                <span className="font-medium text-slate-600 dark:text-slate-500">
                  {lastUpdated?.[type]}
                </span>
              </div>
              <div className="text-xs text-slate-500 dark:text-slate-400 mt-1.5">
                {translate('setting.updated_by')} <br />
                <span className="font-medium text-slate-600 dark:text-slate-500">
                  {updatedBy?.[type] || '-'}
                </span>
              </div>
            </>
          )}
        </div>

        {!isReadOnly && <div className="mt-4">{renderCardAction(type)}</div>}
      </div>
    );
  };

  return (
    <div className="w-full flex flex-col h-full">
      <Card className="grow">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 border-b border-gray-100 dark:border-slate-800 pb-4 gap-3">
          <div>
            <h2 className="text-lg font-bold text-gray-900 dark:text-slate-200">
              {translate('setting.tab.sync_data.sync_title')}
            </h2>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
              {translate('setting.tab.sync_data.subtitle')}
            </p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
          {syncData.map((item) => (
            <div key={item.type} className="col-span-1">
              {renderCardContent(item.type, item.label)}
            </div>
          ))}
        </div>
      </Card>

      <SyncModal
        isOpen={modalOpen}
        mismatchedData={mismatchedData}
        onMismatchedChange={handleMismatchedChange}
        onSave={handleSaveMismatches}
        onBypass={handleBypass}
        mcEasyDrivers={mcEasyDrivers}
        isDriverLoading={isDriverLoading}
        isSaving={syncLoading === 'drivers-active' || syncLoading === 'drivers-global'}
        translate={translate}
      />

      <ConfirmModal
        isOpen={confirmModalConfig.isOpen}
        title={confirmModalConfig.title}
        message={confirmModalConfig.message}
        onConfirm={handleConfirmAction}
        onCancel={closeConfirmModal}
      />

      {pendingKpiTasks.length > 0 && (
        <PendingReasonModal
          isOpen={true}
          noClose={true}
          data={pendingKpiTasks}
          reasons={kpiReasonsList}
          onClose={() => {
            setPendingKpiTasks([]);
            setPendingKpiTarget(null);
          }}
          onSuccess={async () => {
            if (pendingKpiTarget) {
              const { storedUser } = getLocalStorage();
              try {
                await processKpiData({
                  hubId: pendingKpiTarget.hubId,
                  deliveryDates: pendingKpiTarget.dates,
                  userName: storedUser?.name || 'system',
                });
              } catch (err) {
                console.warn('[SyncDataTab] Async KPI retry error:', err.message);
              }
            }
          }}
        />
      )}
    </div>
  );
}
