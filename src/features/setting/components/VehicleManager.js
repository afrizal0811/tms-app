'use client';

import Dropdown from '@/components/dropdown/Dropdown';
import HighlightText from '@/components/HighlightText';
import ConfirmModal from '@/components/modal/ConfirmModal';
import VehicleModal from '@/components/modal/VehicleModal';
import SearchBar from '@/components/SearchBar';
import { deleteVehicle, getVehicles, updateVehicle } from '@/lib/api/mileapp';
import { getLocalStorage } from '@/lib/localStorageHandler';
import { toastError, toastSuccess } from '@/lib/toast';
import { getBasePlate } from '@/lib/utils';
import { useCallback, useEffect, useMemo, useState } from 'react';
import Card from './Card';
import CustomTable from './CustomTable';

const STORAGE_OPTIONS = [
  { label: 'Dry', value: 'Dry' },
  { label: 'Frozen', value: 'Frozen' },
];

export default function VehicleManager({ vehicleTypes, isReadOnly, translate }) {
  const [activeHub, setActiveHub] = useState({ hubId: '', hubName: '' });
  const [vehicles, setVehicles] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [deleteConfig, setDeleteConfig] = useState({ isOpen: false, item: null });
  const [reconfigureData, setReconfigureData] = useState(null);

  useEffect(() => {
    const { storedUser, storedLocation, storedLocationName } = getLocalStorage();
    const hubId = storedUser?.activeHub || storedLocation;
    setActiveHub({ hubId, hubName: storedLocationName });
  }, []);

  const loadVehicles = useCallback(async () => {
    if (!activeHub?.hubId) return;
    setIsLoading(true);
    try {
      const data = await getVehicles();
      setVehicles(data || []);
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setIsLoading(false);
    }
  }, [activeHub, translate]);

  useEffect(() => {
    if (activeHub?.hubId) loadVehicles();

    const handleVehiclesUpdated = () => {
      loadVehicles();
    };

    window.addEventListener('tms-sync-vehicles-completed', handleVehiclesUpdated);
    window.addEventListener('tms-vehicle-configured', handleVehiclesUpdated);

    return () => {
      window.removeEventListener('tms-sync-vehicles-completed', handleVehiclesUpdated);
      window.removeEventListener('tms-vehicle-configured', handleVehiclesUpdated);
    };
  }, [activeHub, loadVehicles]);

  const vehicleTypeOptions = useMemo(
    () => (vehicleTypes || []).map((v) => ({ label: v.name, value: v.name })),
    [vehicleTypes]
  );

  const filteredVehicles = useMemo(() => {
    const list = vehicles || [];
    const q = searchQuery.toLowerCase().trim();

    const masterMap = new Map();
    for (const v of list) {
      const cleanPlate = (v.license_number || v.plat || '').trim();
      const base = v.base_plate || getBasePlate(cleanPlate);
      if (!base) continue;

      const rawType = v._rawType || (v.type ? v.type.replace(/^(Dry|Frozen)-/i, '') : '') || null;
      const vehicleItem = {
        ...v,
        license_number: base,
        plat: base,
        type: rawType,
        _rawType: rawType,
      };

      const targetItem = masterMap.get(base);
      if (!targetItem) {
        masterMap.set(base, { ...vehicleItem, all_ids: [v.id] });
      } else {
        if (!targetItem.all_ids) targetItem.all_ids = [targetItem.id];
        if (!targetItem.all_ids.includes(v.id)) targetItem.all_ids.push(v.id);

        if (cleanPlate === base) {
          masterMap.set(base, {
            ...vehicleItem,
            all_ids: targetItem.all_ids,
            type: vehicleItem.type || targetItem.type,
            storage: vehicleItem.storage || targetItem.storage,
          });
        } else if (!targetItem.type || !targetItem.storage) {
          masterMap.set(base, {
            ...targetItem,
            type: targetItem.type || vehicleItem.type,
            storage: targetItem.storage || vehicleItem.storage,
          });
        }
      }
    }

    const uniqueList = Array.from(masterMap.values()).sort((a, b) => {
      const storageCompare = (a.storage || '').localeCompare(b.storage || '');
      if (storageCompare !== 0) return storageCompare;

      const typeCompare = (a.type || '').localeCompare(b.type || '');
      if (typeCompare !== 0) return typeCompare;

      return (a.license_number || '').localeCompare(b.license_number || '');
    });

    const filtered = !q
      ? uniqueList
      : uniqueList.filter((v) => {
          const plat = (v.license_number || '').toLowerCase();
          const type = (v.type || '').toLowerCase();
          const storage = (v.storage || '').toLowerCase();
          return plat.includes(q) || type.includes(q) || storage.includes(q);
        });

    return filtered.map((item, idx) => ({
      ...item,
      _index: idx + 1,
    }));
  }, [vehicles, searchQuery]);

  const handleUpdateVehicle = async (id, editValues, item) => {
    if (isReadOnly) return;
    try {
      await updateVehicle({
        id: item.id,
        vehicle_id: item.vehicle_id,
        type: editValues.type !== undefined ? editValues.type : item.type,
        storage: editValues.storage !== undefined ? editValues.storage : item.storage,
      });
      toastSuccess(translate('common.toast.success'));
      await loadVehicles();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
      throw e;
    }
  };

  const confirmDelete = async () => {
    const targetItem = deleteConfig.item;
    setDeleteConfig({ isOpen: false, item: null });
    if (!targetItem) return;

    setIsLoading(true);
    try {
      const res = await deleteVehicle({
        id: targetItem.id,
        ids: targetItem.all_ids || (targetItem.id ? [targetItem.id] : []),
        vehicle_id: targetItem.vehicle_id,
        hubId: activeHub.hubId,
      });

      if (res?.stillExists) {
        toastError(
          res.message ||
            'Kendaraan masih terdaftar di MileApp. Silakan tentukan tipe dan storagenya kembali.'
        );
        setReconfigureData([
          {
            id: targetItem.id,
            vehicle_id: targetItem.vehicle_id,
            license_number: targetItem.license_number || targetItem.plat,
            rawTag: targetItem.tags || '',
            storage: targetItem.storage,
            type: targetItem.type,
          },
        ]);
      } else {
        toastSuccess(translate('common.toast.success'));
        await loadVehicles();
      }
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    } finally {
      setIsLoading(false);
    }
  };

  const columns = useMemo(
    () => [
      {
        header: '#',
        field: '_index',
        headerClassName: 'w-10 sm:w-12 text-center',
        align: 'center',
        cellClassName: 'text-center',
        render: (item) => (
          <span className="text-[10px] md:text-xs font-medium text-slate-500 dark:text-slate-400">
            {item._index}
          </span>
        ),
      },
      {
        header: translate('common.license_number'),
        field: 'license_number',
        render: (item) => (
          <span className="font-semibold text-slate-700 dark:text-slate-200 text-[10px] md:text-sm">
            <HighlightText text={item.license_number} highlight={searchQuery} />
          </span>
        ),
      },
      {
        header: translate('common.storage_type'),
        field: 'storage',
        headerClassName: 'w-28 md:w-36',
        render: (item) => (
          <span className="text-[10px] md:text-sm font-medium text-slate-700 dark:text-slate-300">
            <HighlightText text={item.storage || '-'} highlight={searchQuery} />
          </span>
        ),
        renderEdit: (value, onChange) => (
          <Dropdown
            options={STORAGE_OPTIONS}
            value={value}
            onChange={onChange}
            getLabel={(val) => val || translate('common.select')}
            size="sm"
            className="w-full min-w-0"
          />
        ),
      },
      {
        header: translate('common.vehicle_type'),
        field: 'type',
        headerClassName: 'w-32 md:w-48',
        render: (item) => (
          <span className="text-[10px] md:text-sm font-medium text-sky-700 dark:text-sky-400">
            <HighlightText text={item.type || '-'} highlight={searchQuery} />
          </span>
        ),
        renderEdit: (value, onChange) => (
          <Dropdown
            options={vehicleTypeOptions}
            value={value}
            onChange={onChange}
            getLabel={(val) => val || translate('common.select')}
            size="sm"
            className="w-full min-w-0"
          />
        ),
      },
    ],
    [searchQuery, translate, vehicleTypeOptions]
  );

  const msgParts = translate('common.modal.delete_message', { text: '|||' }).split('|||');

  return (
    <Card>
      <ConfirmModal
        isOpen={deleteConfig.isOpen}
        onCancel={() => setDeleteConfig({ isOpen: false, item: null })}
        onConfirm={confirmDelete}
        title={translate('common.modal.delete_title', {
          text: translate('common.vehicle_data'),
        })}
        message={
          <span>
            {msgParts[0]}
            <strong>{deleteConfig.item?.license_number}</strong>
            {msgParts[1]}
          </span>
        }
      />

      {reconfigureData && (
        <VehicleModal
          unmappedData={reconfigureData}
          onCompleted={() => {
            setReconfigureData(null);
            loadVehicles();
          }}
          t={translate}
        />
      )}

      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-4 border-b border-gray-100 dark:border-slate-800 pb-3">
        <div>
          <h2 className="text-lg font-bold text-gray-900 dark:text-slate-200">
            {translate('common.vehicle_data')}{' '}
            {activeHub?.hubName ? `(${activeHub.hubName})` : ''}
          </h2>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
            {translate('setting.tab.general.vehicle_subtitle')}
          </p>
        </div>
        <div className="w-full sm:w-64 shrink-0">
          <SearchBar
            value={searchQuery}
            onChange={setSearchQuery}
            placeholder={translate('common.search')}
            size="md"
            width="w-full"
            tooltip={`${translate('common.license_number')}, ${translate('common.vehicle_type')}`}
          />
        </div>
      </div>

      <CustomTable
        data={filteredVehicles}
        columns={columns}
        isReadOnly={isReadOnly}
        isLoading={isLoading}
        containerHeight="h-[408px]"
        emptyMessage={translate('common.no_data')}
        translate={translate}
        onSave={handleUpdateVehicle}
        onDelete={(item) => setDeleteConfig({ isOpen: true, item })}
      />
    </Card>
  );
}
