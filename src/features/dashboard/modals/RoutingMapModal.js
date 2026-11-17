'use client';

import Button from '@/components/button/Button';
import Dropdown from '@/components/dropdown/Dropdown';
import Modal from '@/components/modal/Modal';
import { useLanguage } from '@/context/LanguageContext';
import { getLocalStorage } from '@/lib/localStorageHandler';
import {
  formatDateUniversal,
  getBasePlate,
  isEmpty,
  parseCoordinates,
  parseCustomerString,
} from '@/lib/utils';
import JSZip from 'jszip';
import { useCallback, useEffect, useMemo, useState } from 'react';
import MapViewSection, { assignPlanSequences } from '../components/MapViewSection';
import { getGroupKey, getGroupTasks, renderGroupToJpeg } from '../helper/help';

export default function RoutingMapModal({ isOpen, onClose, data, selectedDate }) {
  const { t } = useLanguage();
  const [userSelectedGroup, setUserSelectedGroup] = useState('');
  const [selectedCustomer, setSelectedCustomer] = useState('');
  const [taskFilterType, setTaskFilterType] = useState('plan');
  const [prevGroup, setPrevGroup] = useState(null);
  const [mapRo, setMapRo] = useState(null);
  const [mapReal, setMapReal] = useState(null);
  const [isExporting, setIsExporting] = useState(false);

  useEffect(() => {
    if (!mapRo || !mapReal) return;
    let isSyncing = false;
    const syncRoToReal = () => {
      if (isSyncing) return;
      isSyncing = true;
      mapReal.setView(mapRo.getCenter(), mapRo.getZoom(), { animate: false });
      isSyncing = false;
    };
    const syncRealToRo = () => {
      if (isSyncing) return;
      isSyncing = true;
      mapRo.setView(mapReal.getCenter(), mapReal.getZoom(), { animate: false });
      isSyncing = false;
    };
    mapRo.on('move', syncRoToReal);
    mapReal.on('move', syncRealToRo);
    return () => {
      mapRo.off('move', syncRoToReal);
      mapReal.off('move', syncRealToRo);
    };
  }, [mapRo, mapReal]);

  const groups = useMemo(() => {
    if (!data) return [];
    const map = new Map();
    data.forEach((d) => {
      const gKey = d.groupKey || getGroupKey(d);
      const driverName = d.driver && d.driver !== 'N/A' && d.driver !== 'Driver' ? d.driver : '';
      const platName = getBasePlate(d.plat || d.vehicleName) || d.plat || d.vehicleName || '';

      if (gKey && driverName && platName && !map.has(gKey)) {
        map.set(gKey, {
          value: gKey,
          label: `${driverName} (${platName})`,
        });
      }
    });
    return Array.from(map.values()).sort((a, b) => a.label.localeCompare(b.label));
  }, [data]);

  const formattedDateStr = useMemo(() => {
    return formatDateUniversal(selectedDate || new Date(), 'DD.MM.YYYY');
  }, [selectedDate]);

  const hubAcronym = useMemo(() => {
    const { storedLocationAcronym, storedLocationName } = getLocalStorage();
    return storedLocationAcronym || storedLocationName || 'HUB';
  }, []);

  const getPlateNumber = useCallback(
    (groupVal) => {
      const currentGroupObj = groups.find((g) => g.value === groupVal);
      const match = currentGroupObj?.label.match(/\(([^)]+)\)/);
      if (match && match[1]) return match[1].trim();
      const parts = (groupVal || '').split('_');
      return (parts[1] || groupVal || 'VEHICLE').trim();
    },
    [groups]
  );

  const selectedGroup = useMemo(() => {
    if (userSelectedGroup && groups.some((g) => g.value === userSelectedGroup)) {
      return userSelectedGroup;
    }
    return groups.length > 0 ? groups[0].value : '';
  }, [userSelectedGroup, groups]);

  if (selectedGroup !== prevGroup) {
    setPrevGroup(selectedGroup);
    const tasks = data ? data.filter((d) => (d.groupKey || getGroupKey(d)) === selectedGroup) : [];
    const hasHub = tasks.some((t) => t.type === 'HUB_START');
    setSelectedCustomer(hasHub ? 'HUB' : '');
  }

  const filteredTasks = useMemo(
    () => (selectedGroup ? getGroupTasks(data, selectedGroup) : []),
    [data, selectedGroup]
  );

  const nameFrequency = useMemo(() => {
    const counts = {};
    filteredTasks.forEach((t) => {
      if (t.customerName === 'HUB') return;
      const rawStr = t.originalCustomerString || t.customerName;
      const { name } = parseCustomerString(rawStr);
      if (name) counts[name] = (counts[name] || 0) + 1;
    });
    return counts;
  }, [filteredTasks]);

  const resolveDisplayName = useCallback(
    (originalCustomerString, flow) => {
      if (flow === 'Pickup') return 'Pickup';
      if (!originalCustomerString || originalCustomerString === 'HUB') return 'HUB';
      const { name, location } = parseCustomerString(originalCustomerString);
      if (!name) return '-';
      return nameFrequency[name] > 1 ? `${name} (${location || '?'})` : name;
    },
    [nameFrequency]
  );

  const customerOptions = useMemo(() => {
    if (!filteredTasks.length) return [];

    const seqKey = taskFilterType === 'plan' ? 'roSequence' : 'realSequence';
    const sorted = [...filteredTasks].sort((a, b) => (a[seqKey] || 999999) - (b[seqKey] || 999999));
    const sequencedTasks =
      taskFilterType === 'plan' ? assignPlanSequences(sorted, 'roSequence') : sorted;

    const options = [];
    const seen = new Set();

    sequencedTasks.forEach((t) => {
      if (!t.customerName || t.customerName === 'HUB') return;

      const isPending = isEmpty(t.realSequence);

      if (t.flow === 'Pickup') {
        if (!seen.has('Pickup')) {
          seen.add('Pickup');
          options.push({ value: 'Pickup', label: 'PICKUP', isPending: false });
        }
        return;
      }

      const rawStr = t.originalCustomerString || t.customerName;
      const displayName = resolveDisplayName(rawStr, t.flow);
      if (seen.has(displayName)) return;

      seen.add(displayName);
      const seq = t[seqKey];
      const seqPrefix = !isEmpty(seq) ? `${seq}. ` : '';
      const manualPrefix = t.isManualAssign ? '[M] ' : '';
      options.push({
        value: displayName,
        label: `${seqPrefix}${manualPrefix}${displayName}`,
        isPending,
      });
    });

    const hasHub = filteredTasks.some((t) => t.type === 'HUB_START');
    return hasHub ? [{ value: 'HUB', label: 'HUB', isPending: false }, ...options] : options;
  }, [filteredTasks, resolveDisplayName, taskFilterType]);

  const resetMapZoom = useCallback(async () => {
    const targetMap = mapRo || mapReal;
    if (!targetMap) return;
    const coords = filteredTasks
      .map((t) => parseCoordinates(t.longlat))
      .filter(Boolean)
      .map((c) => [c.lat, c.lon]);
    if (coords.length > 0) {
      const L = (await import('leaflet')).default;
      targetMap.fitBounds(L.latLngBounds(coords), { padding: [50, 50], maxZoom: 16 });
    }
  }, [filteredTasks, mapRo, mapReal]);

  const handleFocusCustomer = async (val) => {
    setSelectedCustomer(val);
    if (!val) {
      resetMapZoom();
      return;
    }
    const targetTask = filteredTasks.find((t) => {
      if (val === 'HUB' && t.type === 'HUB_START') return true;
      if (val === 'Pickup' && t.flow === 'Pickup') return true;
      const rawStr = t.originalCustomerString || t.customerName;
      return resolveDisplayName(rawStr, t.flow) === val;
    });

    if (targetTask) {
      const coords = parseCoordinates(targetTask.longlat);
      const targetMap = mapRo || mapReal;
      if (coords && targetMap) {
        const zoomLevel = 16;
        const isMobile = typeof window !== 'undefined' && window.innerWidth < 768;

        if (isMobile) {
          const L = (await import('leaflet')).default;
          const mapSize = targetMap.getSize();
          const targetPoint = targetMap.project([coords.lat, coords.lon], zoomLevel);
          const offsetY = mapSize.y * 0.25;
          const newCenterPoint = new L.Point(targetPoint.x, targetPoint.y + offsetY);
          const newCenterLatLng = targetMap.unproject(newCenterPoint, zoomLevel);
          targetMap.setView(newCenterLatLng, zoomLevel, { animate: true, duration: 1.0 });
        } else {
          targetMap.setView([coords.lat, coords.lon], zoomLevel, { animate: true, duration: 1.0 });
        }
      }
    }
  };

  const onMarkerClick = (task) => {
    if (task.type === 'HUB_START') {
      handleFocusCustomer('HUB');
      return;
    }
    if (task.flow === 'Pickup') {
      handleFocusCustomer('Pickup');
      return;
    }
    if (task) {
      const rawStr = task.originalCustomerString || task.customerName;
      handleFocusCustomer(resolveDisplayName(rawStr, task.flow));
    }
  };

  const handleCloseInfo = () => {
    setSelectedCustomer('');
    resetMapZoom();
  };
  const handleCloseModal = () => {
    if (isExporting) return;
    setUserSelectedGroup('');
    setSelectedCustomer('');
    setTaskFilterType('plan');
    setPrevGroup(null);
    onClose();
  };

  const downloadFile = (href, filename) => {
    const link = document.createElement('a');
    link.download = filename;
    link.href = href;
    link.click();
  };

  const handleDownloadActiveJpg = async () => {
    if (!filteredTasks.length || !selectedGroup) return;
    setIsExporting(true);
    try {
      const L = (await import('leaflet')).default;
      const { toJpeg } = await import('html-to-image');
      const dataUrl = await renderGroupToJpeg(selectedGroup, data, L, toJpeg);
      if (!dataUrl) throw new Error('No data');

      const plate = getPlateNumber(selectedGroup);
      downloadFile(dataUrl, `${plate} - ${formattedDateStr}.jpg`);
    } catch {
      // error silently handled
    } finally {
      setIsExporting(false);
    }
  };

  const handleDownloadAllZip = async () => {
    if (!groups.length) return;
    setIsExporting(true);
    try {
      const L = (await import('leaflet')).default;
      const { toJpeg } = await import('html-to-image');
      const zip = new JSZip();

      for (let i = 0; i < groups.length; i++) {
        const grp = groups[i];
        const dataUrl = await renderGroupToJpeg(grp.value, data, L, toJpeg);
        if (dataUrl) {
          const base64Data = dataUrl.split(',')[1];
          const plate = getPlateNumber(grp.value);
          zip.file(`${plate} - ${formattedDateStr}.jpg`, base64Data, { base64: true });
        }
      }

      const content = await zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      });
      const zipUrl = URL.createObjectURL(content);

      downloadFile(zipUrl, `Routing - ${formattedDateStr} - ${hubAcronym}.zip`);
      URL.revokeObjectURL(zipUrl);
    } catch {
      // error silently handled
    } finally {
      setIsExporting(false);
    }
  };

  const headerContent = (
    <div className="flex flex-col lg:flex-row items-end justify-between w-full">
      <div className="flex flex-row gap-2 w-full lg:w-auto items-end ml-auto">
        <div className="flex flex-col w-1/2 lg:w-64 gap-1">
          <label className="text-[10px] font-bold tracking-wide">{t('common.driver')}</label>
          <Dropdown
            options={groups}
            value={selectedGroup}
            onChange={(val) => setUserSelectedGroup(val)}
            getLabel={(val) => groups.find((g) => g.value === val)?.label || val}
            disabled={!groups.length || isExporting}
            size="sm"
            textSize="text-xs"
            isAutocomplete={true}
            className="w-full text-slate-700 font-semibold"
          />
        </div>

        <div className="flex flex-col w-1/2 lg:w-72 gap-1">
          <div className="flex items-center justify-between">
            <label className="text-[10px] font-bold tracking-wide">{t('common.task')}</label>
            <div className="flex items-center gap-1 bg-gray-100 border border-gray-300 p-0.5 rounded-md shadow-inner text-[11px] font-semibold">
              {[
                {
                  id: 'plan',
                  label: t('dashboard.map.plan'),
                  activeBg: 'bg-[#D97706]',
                  dotBg: 'bg-[#D97706]',
                },
                {
                  id: 'real',
                  label: t('common.actual_seq'),
                  activeBg: 'bg-green-700',
                  dotBg: 'bg-green-700',
                },
              ].map((mode) => {
                const isActive = taskFilterType === mode.id;
                return (
                  <button
                    key={mode.id}
                    type="button"
                    aria-pressed={isActive}
                    onClick={() => setTaskFilterType(mode.id)}
                    className={`flex items-center gap-1.5 px-2 py-0.5 rounded-[5px] cursor-pointer transition-all ${
                      isActive
                        ? `${mode.activeBg} text-white shadow-sm`
                        : 'text-gray-700 hover:bg-white hover:text-gray-900 hover:shadow-xs'
                    }`}
                  >
                    <span
                      className={`w-1.5 h-1.5 rounded-full ${isActive ? 'bg-white' : mode.dotBg}`}
                    />
                    {mode.label}
                  </button>
                );
              })}
            </div>
          </div>
          <select
            value={selectedCustomer}
            onChange={(e) => handleFocusCustomer(e.target.value)}
            onClick={(e) => e.stopPropagation()}
            disabled={!selectedGroup}
            className="w-full border border-gray-300 rounded-lg px-3 py-1.5 text-xs text-slate-700 outline-none bg-gray-50 cursor-pointer disabled:opacity-50 h-8"
          >
            {customerOptions.map((opt, i) => (
              <option
                key={`${opt.value}-${i}`}
                value={opt.value}
                className={opt.isPending ? 'text-red-500 font-medium' : 'text-slate-700'}
                style={opt.isPending ? { color: '#EF4444' } : {}}
              >
                {opt.label}
              </option>
            ))}
          </select>
        </div>

        <Button
          type="split"
          size="sm"
          text={t('common.download')}
          isLoading={isExporting}
          disabled={!selectedGroup || isExporting}
          onClick={handleDownloadAllZip}
          className="bg-sky-600 text-white hover:bg-sky-700 dark:bg-sky-600 border border-sky-800/60"
          width="w-auto"
          options={[
            {
              id: 'current',
              label: `${t('common.download')} (${getPlateNumber(selectedGroup)})`,
              onClick: handleDownloadActiveJpg,
            },
          ]}
        />
      </div>
    </div>
  );

  return (
    <Modal
      isOpen={isOpen}
      onClose={handleCloseModal}
      noClose={isExporting}
      title={t('dashboard.map.title')}
      subtitle={t('dashboard.map.subtitle')}
      headerContent={headerContent}
      maxWidth="max-w-6xl"
      contentClassName="h-[90vh]"
      bodyClassName="p-4 flex flex-col gap-4 overflow-hidden h-full"
    >
      {selectedGroup ? (
        <>
          <MapViewSection
            title={t('dashboard.map.plan')}
            tasks={filteredTasks}
            sequenceKey="roSequence"
            colorClass="bg-[#D97706]"
            lineColor="#FF0000"
            setMap={setMapRo}
            onMarkerClick={onMarkerClick}
            selectedCustomer={selectedCustomer}
            showInfoCard={false}
            isActualMap={false}
            resolveDisplayName={resolveDisplayName}
          />

          <MapViewSection
            title={t('common.actual_seq')}
            tasks={filteredTasks}
            sequenceKey="realSequence"
            colorClass="bg-[#10B981]"
            lineColor="#FF0000"
            setMap={setMapReal}
            onMarkerClick={onMarkerClick}
            selectedCustomer={selectedCustomer}
            showInfoCard={true}
            onCloseCard={handleCloseInfo}
            isActualMap={true}
            resolveDisplayName={resolveDisplayName}
          />
        </>
      ) : (
        <div className="h-full flex items-center justify-center text-gray-500">
          {t('dashboard.map.select_driver')}
        </div>
      )}
    </Modal>
  );
}
