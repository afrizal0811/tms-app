import ConfirmModal from '@/components/modal/ConfirmModal';
import Modal from '@/components/modal/Modal';
import { deleteVehicleExternal, postVehicleExternal } from '@/lib/api/mileapp';
import { getLocalStorage } from '@/lib/localStorageHandler';
import { toastError, toastSuccess } from '@/lib/toast';
import { formatLongDate, getBasePlate } from '@/lib/utils';
import { useEffect, useMemo, useState } from 'react';

const normalizePlate = (plate) => (plate || '').replace(/\s+/g, '').toLowerCase();

const EmptyState = ({ translate }) => (
  <div className="text-center text-slate-400 dark:text-slate-500 text-sm py-6 italic border border-dashed border-gray-300 dark:border-slate-700 rounded-lg bg-gray-50 dark:bg-slate-800/50">
    {translate('common.no_data')}
  </div>
);

const sortVehicles = (vehicles, typeOrder = [], isGabungan = false) => {
  return [...vehicles].sort((a, b) => {
    if (isGabungan) {
      const indexA = typeOrder.indexOf(a.type);
      const indexB = typeOrder.indexOf(b.type);
      const orderA = indexA === -1 ? 999 : indexA;
      const orderB = indexB === -1 ? 999 : indexB;

      if (orderA !== orderB) return orderA - orderB;
    }
    return (a.driverName || a.driver || '').localeCompare(b.driverName || b.driver || '');
  });
};

export default function TruckUsageModal({
  isOpen,
  onClose,
  data,
  hubId,
  onSuccess,
  driverData,
  vehicleTypes,
  translate,
  localeCode,
  masterVehicleList,
}) {
  const [count, setCount] = useState('');
  const [desc, setDesc] = useState('');
  const [initialCount, setInitialCount] = useState('');
  const [initialDesc, setInitialDesc] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [showAll, setShowAll] = useState(false);

  useEffect(() => {
    if (isOpen) {
      setShowAll(false);
    }
  }, [isOpen]);

  useEffect(() => {
    if (data && !data.isTms && !data.isMaster) {
      const initC = data.manualCount > 0 ? String(data.manualCount) : '';
      const initD = data.description || '';

      setCount(initC);
      setInitialCount(initC);
      setDesc(initD);
      setInitialDesc(initD);
    }
  }, [data]);

  const sortedMasterVehicles = useMemo(() => {
    if (!data?.isMaster) return [];
    const vehicles = masterVehicleList?.[data.storage]?.[data.type] || [];
    const seenCombos = new Set();
    const uniqueVehicles = vehicles.filter((vh) => {
      const plateRaw = normalizePlate(vh.plate || vh.plat);
      if (seenCombos.has(plateRaw)) return false;
      seenCombos.add(plateRaw);
      return true;
    });

    return sortVehicles(uniqueVehicles);
  }, [data, masterVehicleList]);

  const tmsDetailsList = useMemo(() => {
    if (!data?.isTms) return [];

    const driverLookup = new Map();
    (driverData || []).forEach((d) => {
      if (d.email) driverLookup.set(d.email.toLowerCase().trim(), d.name);
    });

    const seenCombos = new Set();
    let details = (data.tmsDetails || []).reduce((acc, vh) => {
      const emailToMatch = (vh.driver || '').toLowerCase().trim();
      const driverName = driverLookup.get(emailToMatch) || vh.driver;
      const plateRaw = vh.plate || vh.plat || '';
      const comboKey = `${emailToMatch}_${plateRaw}`;
      if (!seenCombos.has(comboKey)) {
        seenCombos.add(comboKey);
        acc.push({ ...vh, driverName, plate: plateRaw });
      }
      return acc;
    }, []);

    if (showAll && masterVehicleList) {
      const masterCat = masterVehicleList[data.storage]?.[data.type] || [];
      const usedPlates = new Set(details.map((v) => normalizePlate(v.plate)));
      const unusedVehicles = masterCat
        .filter((v) => !usedPlates.has(normalizePlate(v.plate)))
        .map((v) => ({ ...v, isUnused: true, driverName: v.driver }));

      details = [...details, ...unusedVehicles];
    }

    return sortVehicles(details, vehicleTypes, data.type === 'Gabungan');
  }, [data, driverData, showAll, masterVehicleList, vehicleTypes]);

  if (!data) return null;

  const getActiveHub = () => {
    const active =
      hubId || (typeof window !== 'undefined' ? getLocalStorage()?.storedLocation : null);
    if (!active || active === 'all') {
      toastError(translate('common.toast.error', { err: 'Hub belum dipilih atau tidak valid' }));
      return null;
    }
    return active;
  };

  const handleSave = async () => {
    const activeHub = getActiveHub();
    if (!activeHub) return;

    setIsLoading(true);
    try {
      const resData = await postVehicleExternal({
        hubId: activeHub,
        date: data.date,
        storageType: data.storage,
        vehicleType: data.type,
        count,
        description: desc,
      });

      toastSuccess(translate('common.toast.success'));
      onSuccess(resData.data ? resData.data : resData);
      onClose();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setIsLoading(false);
    }
  };

  const handleDelete = async () => {
    if (!data?.id) return;
    const activeHub = getActiveHub();
    if (!activeHub) return;

    setIsConfirmOpen(false);
    setIsLoading(true);
    try {
      await deleteVehicleExternal({
        hubId: activeHub,
        date: data.date,
        storageType: data.storage,
        vehicleType: data.type,
      });

      toastSuccess(translate('common.toast.success'));
      onSuccess({ id: data.id, isDelete: true });
      onClose();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setIsLoading(false);
    }
  };

  const renderMasterVehicle = () => {
    const modalTitle = `Master - ${data.storage} (${data.type === 'Gabungan' ? 'Total' : data.type})`;

    return (
      <Modal isOpen={isOpen} onClose={onClose} maxWidth="max-w-md" title={modalTitle}>
        <div className="flex flex-col gap-3 pt-2 pb-2">
          <div className="bg-sky-50 dark:bg-sky-900/30 text-sky-700 dark:text-sky-400 text-sm px-3 py-2.5 rounded-md border border-sky-100 dark:border-sky-800 flex justify-between items-center">
            <span>Total</span>
            <span className="font-bold text-lg">{data.masterTotal}</span>
          </div>
          <div className="mt-1">
            <h4 className="text-xs font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-2">
              {translate('summary.tabs.truck_usage.modal.vehicle_list')}
            </h4>
            {sortedMasterVehicles.length > 0 ? (
              <div className="flex flex-col gap-2 max-h-[50vh] overflow-y-auto pr-1">
                {sortedMasterVehicles.map((vh, idx) => (
                  <div
                    key={idx}
                    className="flex flex-col p-3 border border-gray-200 dark:border-slate-700 rounded-lg bg-slate-50 dark:bg-slate-800 hover:bg-white dark:hover:bg-slate-700/50 transition-colors shadow-sm"
                  >
                    <div className="flex justify-between items-start mb-1">
                      <div className="font-bold text-slate-800 dark:text-slate-200 text-base">
                        {getBasePlate(vh.plate)}
                      </div>
                      {vh.type && (
                        <div className="text-[10px] font-bold px-2 py-0.5 bg-slate-200 dark:bg-slate-700 text-slate-600 dark:text-slate-300 rounded">
                          {vh.type}
                        </div>
                      )}
                    </div>
                    <div className="text-xs text-slate-500 dark:text-slate-400 mt-1 border-t border-gray-100 dark:border-slate-700/50 pt-1">
                      <span className="font-medium text-slate-700 dark:text-slate-300">
                        {vh.driver}
                      </span>
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <EmptyState translate={translate} />
            )}
          </div>
        </div>
      </Modal>
    );
  };

  const renderTmsVehicle = () => {
    let modalTitle = `TMS - ${data.storage} (${data.type})`;
    if (data.type === 'Gabungan') {
      const storageTitles = { Dry: 'Dry', Frozen: 'Frozen', OTV: 'OTV' };
      if (storageTitles[data.storage]) modalTitle = `TMS - Total (${storageTitles[data.storage]})`;
    }

    return (
      <Modal
        isOpen={isOpen}
        onClose={onClose}
        maxWidth="max-w-md"
        title={modalTitle}
        subtitle={formatLongDate(data.date, localeCode)}
      >
        <div className="flex flex-col gap-3 pt-2 pb-2">
          <div className="bg-sky-50 dark:bg-sky-900/30 text-sky-700 dark:text-sky-400 text-sm px-3 py-2.5 rounded-md border border-sky-100 dark:border-sky-800 flex justify-between items-center">
            <span>Total</span>
            <span className="font-bold text-lg">{data.tmsCount}</span>
          </div>
          <div className="mt-1">
            <div className="flex justify-between items-center mb-2">
              <h4 className="text-xs font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider">
                {translate('summary.tabs.truck_usage.modal.vehicle_list')}
              </h4>
              <label className="flex items-center gap-2 text-xs cursor-pointer text-slate-700 dark:text-slate-300 font-medium">
                <input
                  type="checkbox"
                  className="rounded border-gray-300 text-sky-600 focus:ring-sky-500"
                  checked={showAll}
                  onChange={(e) => setShowAll(e.target.checked)}
                />
                {translate('summary.tabs.truck_usage.modal.all_vehicles')}
              </label>
            </div>
            {tmsDetailsList.length > 0 ? (
              <div className="flex flex-col gap-2 max-h-[50vh] overflow-y-auto pr-1">
                {tmsDetailsList.map((vh, idx) => (
                  <div
                    key={idx}
                    className={`flex flex-col p-3 border rounded-lg transition-colors shadow-sm ${vh.isUnused ? 'bg-red-50 border-red-200 dark:bg-red-900/20 dark:border-red-800' : 'border-gray-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 hover:bg-white dark:hover:bg-slate-700/50'}`}
                  >
                    <div className="flex justify-between items-start mb-1">
                      <div
                        className={`font-bold text-base ${vh.isUnused ? 'text-red-700 dark:text-red-400' : 'text-slate-800 dark:text-slate-200'}`}
                      >
                        {getBasePlate(vh.plate)}
                      </div>
                      {vh.type && (
                        <div
                          className={`text-[10px] font-bold px-2 py-0.5 rounded ${vh.isUnused ? 'bg-red-200 dark:bg-red-900/50 text-red-700 dark:text-red-300' : 'bg-slate-200 dark:bg-slate-700 text-slate-600 dark:text-slate-300'}`}
                        >
                          {vh.type}
                        </div>
                      )}
                    </div>
                    <div
                      className={`text-xs mt-1 border-t pt-1 ${vh.isUnused ? 'text-red-500 dark:text-red-400 border-red-100 dark:border-red-800/50' : 'text-slate-500 dark:text-slate-400 border-gray-100 dark:border-slate-700/50'}`}
                    >
                      <span className="font-medium">{vh.driverName}</span>
                    </div>
                    {vh.isUnused && (
                      <div className="mt-1 text-[10px] text-red-500 font-semibold uppercase tracking-wider block">
                        {translate('summary.tabs.truck_usage.modal.unused')}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            ) : (
              <EmptyState translate={translate} />
            )}
          </div>
        </div>
      </Modal>
    );
  };

  const renderManualForm = () => {
    const totalInput = (parseInt(count, 10) || 0) + (data?.tmsCount || 0);
    const masterCapacity = data?.masterTotal || 0;
    const isOverLimit = data?.type !== 'Interbranch' && totalInput > masterCapacity;
    const isChanged = count !== initialCount || desc !== initialDesc;
    const isSaveDisabled = isLoading || count === '' || !desc.trim() || !isChanged;

    const msgParts = translate('common.modal.delete_message', { text: '|||' }).split('|||');

    return (
      <>
        <ConfirmModal
          isOpen={isConfirmOpen}
          onCancel={() => setIsConfirmOpen(false)}
          onConfirm={handleDelete}
          title={translate('common.modal.delete_title', { text: 'data' })}
          message={
            <span>
              {msgParts[0]}
              <strong>Non TMS data</strong>
              {msgParts[1]}
            </span>
          }
        />

        <Modal
          isOpen={isOpen && !isConfirmOpen}
          onClose={onClose}
          maxWidth="max-w-md"
          title={`Non TMS - ${data.storage} (${data.type})`}
          subtitle={formatLongDate(data.date, localeCode)}
          footer={
            <div className="flex justify-between items-center w-full">
              <div>
                {data.id && (
                  <button
                    disabled={isLoading}
                    onClick={() => setIsConfirmOpen(true)}
                    className="px-4 py-2 bg-red-100 dark:bg-red-900/40 text-red-700 dark:text-red-400 rounded-md hover:bg-red-200 dark:hover:bg-red-900/60 font-medium text-sm transition-colors cursor-pointer"
                  >
                    {translate('common.button.btn_delete')}
                  </button>
                )}
              </div>
              <div className="flex gap-2">
                <button
                  disabled={isLoading}
                  onClick={onClose}
                  className="px-4 py-2 bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-md hover:bg-slate-300 dark:hover:bg-slate-600 font-medium text-sm cursor-pointer transition-colors"
                >
                  {translate('common.button.btn_cancel')}
                </button>
                <button
                  disabled={isSaveDisabled}
                  onClick={handleSave}
                  className="px-4 py-2 bg-sky-600 text-white rounded-md hover:bg-sky-700 font-medium text-sm min-w-[90px] disabled:bg-slate-300 dark:disabled:bg-slate-700 dark:disabled:text-slate-500 disabled:cursor-not-allowed cursor-pointer transition-colors"
                >
                  {isLoading
                    ? translate('summary.tabs.truck_usage.modal.loading_text')
                    : translate('common.button.btn_save')}
                </button>
              </div>
            </div>
          }
        >
          <div className="flex flex-col gap-4 pt-2 pb-2 relative">
            <div>
              <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                {translate('summary.tabs.truck_usage.modal.manual_total')}{' '}
                <span className="text-red-500">*</span>
              </label>
              <input
                type="number"
                min="0"
                value={count}
                onChange={(e) => setCount(e.target.value)}
                placeholder={translate('summary.tabs.truck_usage.modal.manual_total_placeholder')}
                className="w-full px-3 py-2 border border-gray-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 rounded-md focus:outline-none focus:ring-1 focus:ring-sky-500"
              />
            </div>

            {isOverLimit && (
              <div className="bg-red-50 dark:bg-red-900/30 border border-red-200 dark:border-red-800 text-red-700 dark:text-red-400 px-3 py-2 rounded-md text-xs font-medium">
                ⚠️ {translate('summary.tabs.truck_usage.modal.over_limit')}
              </div>
            )}

            <div>
              <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                {translate('summary.tabs.truck_usage.modal.comment')}{' '}
                <span className="text-red-500">*</span>
              </label>
              <textarea
                rows="3"
                value={desc}
                onChange={(e) => setDesc(e.target.value)}
                placeholder={translate('summary.tabs.truck_usage.modal.comment_placeholder')}
                className="w-full px-3 py-2 border border-gray-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 rounded-md focus:outline-none focus:ring-1 focus:ring-sky-500 resize-none"
              ></textarea>
            </div>
          </div>
        </Modal>
      </>
    );
  };

  if (data.isMaster) return renderMasterVehicle();
  if (data.isTms) return renderTmsVehicle();
  return renderManualForm();
}
