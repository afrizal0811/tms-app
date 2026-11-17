'use client';

import Button from '@/components/button/Button';
import InformationButton from '@/components/button/InformationButton';
import ToggleButton from '@/components/button/ToggleButton';
import CustomDatePicker from '@/components/CustomDatePicker';
import StorageTypeFilter from '@/components/dropdown/StorageTypeFilter';
import VehicleTypeFilter from '@/components/dropdown/VehicleTypeFilter';
import TaskModal from '@/components/modal/TaskModal';
import PageTemplate from '@/components/page/PageTemplate';
import SearchBar from '@/components/SearchBar';
import Tooltip from '@/components/Tooltip';
import { useLanguage } from '@/context/LanguageContext';
import { getLocalStorage, setLocalStorage } from '@/lib/localStorageHandler';
import {
  calculateStartFinishDates,
  checkInvalidSoList,
  formatDateUniversal,
  getBaseVehicleType,
  isEmpty,
  normalizeEmail,
  parseCustomerString,
  sortRows,
  toApiDateString,
  tomorrowDate,
  getBasePlate,
} from '@/lib/utils';
import { useEffect, useMemo, useRef, useState } from 'react';
import { getVehicles, getHubs, getLocationHistories, getResults, getTasks } from '@/lib/api/mileapp';
import { toastError, toastWarning } from '@/lib/toast';
import CustomTable from './components/CustomTable';
import {
  driverTime,
  handleFullDeliveryFormDownload,
  handleFullDeliveryListDownload,
  handleFullRouteTransDownload,
  handlePartialDeliveryFormDownload,
  handlePartialDeliveryListDownload,
  handlePartialRouteTransDownload,
} from './helper';
import BunListModal from './modal/BunListModal';
import PartialRoutingModal from './modal/PartialRoutingModal';

const getStoragePrefix = (storageFilter) => {
  if (storageFilter.includes('DRY') && !storageFilter.includes('FROZEN')) return 'DRY';
  if (!storageFilter.includes('DRY') && storageFilter.includes('FROZEN')) return 'FRZ';
  return '';
};

const buildActivePlatsByEmail = (driverData = [], resultsData = []) => {
  const platesInResults = new Set();
  (resultsData || []).forEach((res) => {
    [
      ...Object.keys(res?.totalDistance || {}),
      ...Object.keys(res?.time || {}),
      ...Object.keys(res?.hubTimes || {}),
    ].forEach((p) => {
      const base = getBasePlate(p) || p;
      if (base) platesInResults.add(base);
    });
  });

  const activePlatsByEmail = new Map();
  (driverData || []).forEach((driver) => {
    if (!driver?.email || !driver?.plat) return;
    const normalizedEmail = normalizeEmail(driver.email);
    const candidatePlat = getBasePlate(driver.plat) || driver.plat;
    if (!normalizedEmail || !candidatePlat) return;
    const existingPlat = activePlatsByEmail.get(normalizedEmail);
    if (
      !existingPlat ||
      (!platesInResults.has(existingPlat) && platesInResults.has(candidatePlat))
    ) {
      activePlatsByEmail.set(normalizedEmail, candidatePlat);
    }
  });
  return activePlatsByEmail;
};

const persistDeliveryPageSetting = (key, value) => {
  const { storedSession } = getLocalStorage();
  if (storedSession) {
    setLocalStorage('data', {
      ...storedSession,
      deliveryPage: { ...(storedSession.deliveryPage || {}), [key]: value },
    });
  }
};

export default function DeliveryPage() {
  const { t } = useLanguage();
  const [activeVehicleId, setActiveVehicleId] = useState(null);
  const [allRoutes, setAllRoutes] = useState([]);
  const [bunSoList, setBunSoList] = useState([]);
  const [downloadType, setDownloadType] = useState(null);
  const [hubsData, setHubsData] = useState([]);
  const [isBunModalOpen, setIsBunModalOpen] = useState(false);
  const [isClient, setIsClient] = useState(false);
  const [isDetailView, setIsDetailView] = useState(false);
  const [isDownloadDropdownOpen, setIsDownloadDropdownOpen] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isNoBun, setIsNoBun] = useState(false);
  const [isSplitMultitrip, setIsSplitMultitrip] = useState(false);
  const [isSplitStorageType, setIsSplitStorageType] = useState(true);
  const [isRoutingModalOpen, setIsRoutingModalOpen] = useState(false);
  const [routingResults, setRoutingResults] = useState([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedDate, setSelectedDate] = useState('');
  const [sortConfig, setSortConfig] = useState({ key: 'no', direction: 'asc' });
  const [storageFilter, setStorageFilter] = useState(['DRY', 'FROZEN']);
  const [typeFilter, setTypeFilter] = useState('');
  const [masterVehicleTypes, setMasterVehicleTypes] = useState([]);
  const [timeMap, setTimeMap] = useState(new Map());
  const [isRouteSettingsOpen, setIsRouteSettingsOpen] = useState(false);
  const [selectedTaskId, setSelectedTaskId] = useState(null);
  const [isTaskModalOpen, setIsTaskModalOpen] = useState(false);
  const downloadDropdownRef = useRef(null);
  const lastWarnedPlates = useRef('');

  useEffect(() => {
    setIsClient(true);
    const date = new Date();
    setSelectedDate(formatDateUniversal(date, 'YYYY-MM-DD'));
    const { storedSession } = getLocalStorage();
    if (storedSession && storedSession.deliveryPage) {
      const dp = storedSession.deliveryPage;
      if (typeof dp.isDetailView === 'boolean') {
        setIsDetailView(dp.isDetailView);
      }
      if (typeof dp.isNoBun === 'boolean') {
        setIsNoBun(dp.isNoBun);
      }
      if (typeof dp.isSplitMultitrip === 'boolean') {
        setIsSplitMultitrip(dp.isSplitMultitrip);
      }
      if (typeof dp.isSplitStorageType === 'boolean') {
        setIsSplitStorageType(dp.isSplitStorageType);
      }
    }
  }, []);

  useEffect(() => {
    const fetchHubsData = async () => {
      try {
        const res = await getHubs();
        setHubsData(res);
      } catch (e) {
        toastError(t('common.toast.error', { err: e.message }), e);
      }
    };
    fetchHubsData();
  }, [t]);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (downloadDropdownRef.current && !downloadDropdownRef.current.contains(e.target)) {
        setIsDownloadDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleToggleView = (isDetail) => {
    setIsDetailView(isDetail);
    persistDeliveryPageSetting('isDetailView', isDetail);
  };

  const handleToggleNoBun = (isActive) => {
    setIsNoBun(isActive);
    persistDeliveryPageSetting('isNoBun', isActive);
  };

  const handleToggleSplitMultitrip = (isActive) => {
    setIsSplitMultitrip(isActive);
    persistDeliveryPageSetting('isSplitMultitrip', isActive);
  };

  const handleToggleSplitStorageType = (isActive) => {
    setIsSplitStorageType(isActive);
    persistDeliveryPageSetting('isSplitStorageType', isActive);
  };

  const handleRowClick = (taskId) => {
    if (!taskId || taskId === '-' || taskId.includes('hub')) return;
    setSelectedTaskId(taskId);
    setIsTaskModalOpen(true);
  };

  const runPartialDownload = (type, fileNamePrefix, excludeSoList = []) => {
    const baseProps = {
      routingResults,
      filteredVehicleRoutes,
      setIsDownloading,
      t,
      selectedDate,
      timeMap,
      isDetailView,
      fileNamePrefix,
      excludeSoList,
      sortConfig,
      isSplitMultitrip,
      isSplitStorageType,
    };

    if (type === 'routeTransaction') {
      handlePartialRouteTransDownload(baseProps);
    } else if (type === 'deliveryList') {
      handlePartialDeliveryListDownload(baseProps);
    } else if (type === 'deliveryForm') {
      handlePartialDeliveryFormDownload(baseProps);
    }
  };

  const runFullDownload = (type, fileNamePrefix, excludeSoList = []) => {
    const baseProps = {
      filteredVehicleRoutes,
      setIsDownloading,
      t,
      selectedDate,
      timeMap,
      isDetailView,
      excludeSoList,
      sortConfig,
      isSplitMultitrip,
      isSplitStorageType,
    };

    if (type === 'routeTransaction') {
      handleFullRouteTransDownload(baseProps);
    } else if (type === 'deliveryList') {
      handleFullDeliveryListDownload({ ...baseProps, fileNamePrefix });
    } else if (type === 'deliveryForm') {
      handleFullDeliveryFormDownload(baseProps);
    }
  };

  const handleDownloadTrigger = (type) => {
    if (isEmpty(filteredVehicleRoutes)) {
      toastError(t('common.toast.error', { err: t('common.no_data') }));
      return;
    }
    setIsDownloadDropdownOpen(false);

    if (['routeTransaction', 'deliveryList', 'deliveryForm'].includes(type)) {
      const activeHub = hubsData.activeHub;

      let excludeList = [];
      if (type === 'routeTransaction' && isNoBun) {
        excludeList = bunSoList.filter((b) => !b.hasNonBun).map((b) => b.so);
      }

      if (activeHub?.hasPartialRouting) {
        setDownloadType(type);
        setIsRoutingModalOpen(true);
      } else {
        runFullDownload(type, getStoragePrefix(storageFilter), excludeList);
      }
    }
  };

  const handleDownloadBunSpecific = (excludeList) => {
    setIsBunModalOpen(false);
    const activeHub = hubsData.activeHub;

    const baseProps = {
      setIsDownloading,
      t,
      selectedDate,
      excludeSoList: excludeList,
      isSplitMultitrip,
      isSplitStorageType,
    };

    if (activeHub?.hasPartialRouting) {
      handlePartialRouteTransDownload({ ...baseProps, routingResults });
    } else {
      handleFullRouteTransDownload({ ...baseProps, filteredVehicleRoutes });
    }
  };

  useEffect(() => {
    if (!selectedDate) return;
    const deliveryDateObj = new Date(selectedDate);
    if (deliveryDateObj.getDay() === 0) {
      setAllRoutes([]);
      setIsLoading(false);
      return;
    }

    const fetchData = async () => {
      setIsLoading(true);
      setAllRoutes([]);
      setActiveVehicleId(null);
      setTimeMap(new Map());
      try {
        let currentHubs = hubsData;
        if (currentHubs.length === 0) {
          currentHubs = await getHubs();
          setHubsData(currentHubs);
        }

        const dateStr = formatDateUniversal(deliveryDateObj, 'YYYY-MM-DD');

        const startD = new Date(selectedDate);
        startD.setHours(0, 0, 0, 0);
        const endD = new Date(selectedDate);
        endD.setHours(23, 59, 59, 999);

        const { timeFrom: historyFrom, timeTo: historyTo } =
          calculateStartFinishDates(selectedDate);

        const [resultsData, historyData, tasksResponse, rawDrivers] = await Promise.all([
          getResults({
            deliveryDate: dateStr,
          }),
          getLocationHistories({
            date: dateStr,
            timeFrom: historyFrom,
            timeTo: historyTo,
          }),
          getTasks({
            timeFrom: toApiDateString(startD),
            timeTo: toApiDateString(endD),
            status: 'DONE,ONGOING,UNASSIGNED',
          }),
          getVehicles(),
        ]);

        const driverData = Array.isArray(rawDrivers) ? rawDrivers : rawDrivers?.data || [];
        setRoutingResults(resultsData);
        const filteredTasks = tasksResponse.filter((t) => t?.assignee);

        const tempBunList = [];
        filteredTasks.forEach((task) => {
          const regexBun = /\bbun\b/i;
          const allItems = (task.listProduct || []).map((p) => p.title || '');
          const hasBun = allItems.some((title) => regexBun.test(title));

          if (hasBun && task.customerOrder) {
            const parsedCust = parseCustomerString(task.customerOrder);
            const sos = (parsedCust.invoiceNumber || task.orderId || '')
              .split(',')
              .map((s) => s.trim())
              .filter(Boolean);

            const hasNonBun = allItems.some((title) => !regexBun.test(title));

            sos.forEach((so) => {
              tempBunList.push({
                so,
                customer: parsedCust.name || '-',
                vehicle: task.basePlat || task.vehicleName || task.plat || '-',
                items: allItems,
                hasNonBun,
              });
            });
          }
        });
        tempBunList.sort((a, b) => a.so.localeCompare(b.so));
        setBunSoList(tempBunList);

        const soToWarehouseMap = new Map();
        filteredTasks.forEach((t) => {
          if (t.flow === 'Pickup' && t.orderId) {
            const wh = t['warehouseName-1'] || t.warehouseName || '';
            if (wh)
              t.orderId
                .split(',')
                .map((s) => s.trim())
                .filter(Boolean)
                .forEach((so) => soToWarehouseMap.set(so, wh));
          }
        });

        const activePlatsByEmail = buildActivePlatsByEmail(driverData, resultsData);

        const tasksByPlat = filteredTasks.reduce((groups, task) => {
          const email = normalizeEmail(task?.assignee);
          const plat = task.basePlat || activePlatsByEmail.get(email) || t('common.others');
          const groupKey = `${email}_${plat}`;

          if (!groups[groupKey])
            groups[groupKey] = {
              vehicleId: groupKey,
              plat,
              email,
              vehicleType: task.vehicleType || '',
              assigneeName: task.driverName || task.user?.name || task.courierName || email,
              tasks: [],
            };
          groups[groupKey].tasks.push(task);
          return groups;
        }, {});

        const resultHubsByPlat = new Map();
        resultsData.forEach((res) => {
          const hubTimes = res.hubTimes || {};
          for (const [rawPlat, hubs] of Object.entries(hubTimes)) {
            if (!Array.isArray(hubs) || hubs.length === 0) continue;
            const basePlat = getBasePlate(rawPlat) || rawPlat;
            const startHub = hubs.find((t) => t.category === 'first' || t.order === 0) || hubs[0];
            const endHub = hubs.find((t) => t.category === 'last') || hubs[hubs.length - 1];
            const middleHubs = hubs.filter(
              (t) => t.category === 'middle' || (t.category && t.category.startsWith('middle'))
            );
            const hubObj = { startHub, endHub, middleHubs };
            resultHubsByPlat.set(basePlat, hubObj);
            const driver = (driverData || []).find((d) => getBasePlate(d.plat) === basePlat);
            if (driver?.email) {
              resultHubsByPlat.set(`${normalizeEmail(driver.email)}_${basePlat}`, hubObj);
            }
          }
        });

        const finalRoutes = Object.values(tasksByPlat).map(
          ({ vehicleId, plat, tasks, email, assigneeName, vehicleType }) => {
            tasks.sort(
              (a, b) => (a.routePlannedOrder ?? Infinity) - (b.routePlannedOrder ?? Infinity)
            );

            const taskTrips = tasks.map((task) => {
              const sos = (task.orderId || '')
                .split(',')
                .map((s) => s.trim())
                .filter(Boolean);

              const startFormat = task.startTime ? formatDateUniversal(task.startTime) : null;
              const taskId = task._id || task.taskId || task.id;
              const effectiveOrder = task.routePlannedOrder;
              const isManual = effectiveOrder == null;

              return {
                visitId: taskId,
                routePlannedOrder: effectiveOrder,
                visitName: task.customerOrder || task.customerName || '',
                orderId: task.orderId,
                flow: task.flow,
                warehouseName:
                  task.flow === 'Pickup' ? task['warehouseName-1'] || task.warehouseName : '',
                locationName: task.locationName || null,
                openTime: task.openTime,
                closeTime: task.closeTime,
                eta: startFormat ? `${startFormat} ${task.eta}` : task.eta,
                etd: startFormat ? `${startFormat} ${task.etd}` : task.etd,
                isHub: false,
                isManual: isManual,
                isReDelivery: task.flow?.toLowerCase().includes('re delivery'),
                soWarehouseMapping: sos.map((so) => ({ so, wh: soToWarehouseMap.get(so) || '' })),
                typeStorage: task.typeStorage,
              };
            });

            const hubData =
              resultHubsByPlat.get(`${email}_${plat}`) ||
              resultHubsByPlat.get(plat) ||
              resultHubsByPlat.get(email);
            const hubTime =
              hubData?.endHub?.eta && hubData?.startHub?.etd
                ? {
                    eta: `${selectedDate} ${hubData.endHub.eta}`,
                    etd: `${selectedDate} ${hubData.startHub.etd}`,
                  }
                : null;
            const finalTrips = [];

            const middleHubTrips = (hubData?.middleHubs || []).map((h) => ({
              ...h,
              isHub: true,
              visitName: 'HUB',
              routePlannedOrder: h.order,
              eta: h.eta ? `${selectedDate} ${h.eta}` : null,
              etd: h.etd ? `${selectedDate} ${h.etd}` : null,
              isMiddleHub: true,
            }));

            const combinedTrips = [...taskTrips, ...middleHubTrips].sort(
              (a, b) => (a.routePlannedOrder ?? Infinity) - (b.routePlannedOrder ?? Infinity)
            );

            if (hubData?.startHub)
              finalTrips.push({ ...hubData.startHub, ...hubTime, isHub: true, visitName: 'HUB' });
            finalTrips.push(...combinedTrips);
            if (hubData?.endHub)
              finalTrips.push({ ...hubData.endHub, ...hubTime, isHub: true, visitName: 'HUB' });

            return {
              vehicleId,
              vehicleName: plat,
              assignee: email,
              assigneeName,
              basePlat: plat,
              vehicleType: vehicleType,
              trips: finalTrips,
            };
          }
        );

        sortRows(finalRoutes, 'vehicleName', 'vehicleName');

        const badPlates = new Set();
        finalRoutes.forEach((r) => {
          r.trips.forEach((t) => {
            if (t.isHub || !t.orderId || t.isReDelivery) return;
            const parsed = parseCustomerString(t.visitName);
            const isBadCust = isEmpty(parsed?.id) || isEmpty(parsed?.location);
            const bad = checkInvalidSoList(parsed.invoiceNumber || t.orderId, isBadCust);
            if (bad) badPlates.add(r.vehicleName || 'Vehicle');
          });
        });

        if (badPlates.size > 0) {
          const platesStr = Array.from(badPlates).join('\n');
          if (lastWarnedPlates.current !== platesStr) {
            toastWarning(`${t('delivery.toast.invalid_invoice')}\n${platesStr}`);
            lastWarnedPlates.current = platesStr;
          }
        } else {
          lastWarnedPlates.current = '';
        }

        setAllRoutes(finalRoutes);
        setActiveVehicleId(finalRoutes.length > 0 ? finalRoutes[0].vehicleId : null);
        setTimeMap(driverTime(historyData));
      } catch (err) {
        toastError(t('common.toast.error', { err: err.message }), err);
      } finally {
        setIsLoading(false);
      }
    };
    fetchData();
    // eslint-disable-next-line
  }, [selectedDate, t]);

  const enrichedRoutes = useMemo(() => {
    if (isEmpty(allRoutes)) return [];
    const soTracker = {};

    allRoutes.forEach((route) => {
      route.trips.forEach((trip) => {
        if (trip.isHub || !trip.orderId) return;
        trip.orderId
          .split(',')
          .map((s) => s.trim())
          .filter(Boolean)
          .forEach((so) => {
            if (!soTracker[so]) soTracker[so] = { Pickup: null, Delivery: null };
            soTracker[so][trip.flow] = route.vehicleName;
          });
      });
    });

    return allRoutes.map((route) => {
      const tripsWithSyncStatus = route.trips.map((trip) => {
        if (trip.isHub || !trip.orderId) {
          return {
            ...trip,
            isUnsync: false,
            partnerVehicle: null,
            syncDetails: {},
            hasAnyPartner: false,
            partnerSOs: [],
          };
        }

        const individualSOs = trip.orderId
          .split(',')
          .map((s) => s.trim())
          .filter(Boolean);
        let isUnsync = false;
        const partnerVehicles = new Set();
        const syncDetails = {};
        const partnerSOs = [];

        individualSOs.forEach((so) => {
          const pickupVehicle = soTracker[so]?.Pickup;
          const deliveryVehicle = soTracker[so]?.Delivery;
          if (pickupVehicle && deliveryVehicle) {
            partnerSOs.push(so);
            if (pickupVehicle !== deliveryVehicle) {
              isUnsync = true;
              const partner = trip.flow === 'Pickup' ? deliveryVehicle : pickupVehicle;
              if (partner) {
                partnerVehicles.add(partner);
                syncDetails[so] = partner;
              }
            }
          }
        });

        return {
          ...trip,
          isUnsync,
          partnerVehicle: partnerVehicles.size > 0 ? Array.from(partnerVehicles).join(', ') : null,
          syncDetails,
          isRedelivery: trip.flow?.toLowerCase().includes('re delivery'),
          hasAnyPartner: partnerSOs.length > 0,
          partnerSOs,
        };
      });

      let hasInvalidSo = false;
      tripsWithSyncStatus.forEach((t) => {
        if (t.isHub || !t.orderId || t.isReDelivery) return;
        const parsed = parseCustomerString(t.visitName);
        const isBadCust = isEmpty(parsed?.id) || isEmpty(parsed?.location);
        if (checkInvalidSoList(parsed.invoiceNumber || t.orderId, isBadCust)) {
          hasInvalidSo = true;
        }
      });

      return {
        ...route,
        trips: tripsWithSyncStatus,
        hasManual: tripsWithSyncStatus.some((t) => t.isManual),
        hasUnsync: tripsWithSyncStatus.some((t) => t.isUnsync),
        isRedelivery: tripsWithSyncStatus.some((t) => t.isRedelivery),
        hasInvalidSo,
      };
    });
  }, [allRoutes]);

  const filteredVehicleRoutes = useMemo(() => {
    let routes = enrichedRoutes;
    if (searchQuery) {
      const lower = searchQuery.toLowerCase();
      routes = routes
        .map((r) => {
          const dName = (r.basePlat || '').toLowerCase();
          if (
            dName.includes(lower) ||
            (r.vehicleName || '').toLowerCase().includes(lower) ||
            (r.assigneeName || '').toLowerCase().includes(lower)
          )
            return r;
          const matchingTrips = r.trips.filter(
            (t) =>
              (t.visitName || '').toLowerCase().includes(lower) ||
              (t.orderId || '').toLowerCase().includes(lower) ||
              (t.warehouseName || '').toLowerCase().includes(lower)
          );
          return matchingTrips.length > 0 ? { ...r, trips: matchingTrips } : null;
        })
        .filter(Boolean);
    }

    routes = routes.filter((route) => {
      if (storageFilter.length === 0) return false;
      let keep = true;
      if (storageFilter.length === 1) {
        const dName = route.assigneeName;
        const vType = route.vehicleType;
        keep =
          (storageFilter.includes('DRY') && (dName.includes("'DRY'") || vType.includes('DRY'))) ||
          (storageFilter.includes('FROZEN') &&
            (dName.includes("'FRZ'") || vType.includes('FROZEN')));
      }
      if (!keep) return false;

      if (typeFilter && typeFilter !== 'all') {
        if (getBaseVehicleType(route.vehicleType, masterVehicleTypes) !== typeFilter) return false;
      }
      return true;
    });

    return sortRows([...routes], 'vehicleName', 'vehicleName');
  }, [searchQuery, enrichedRoutes, storageFilter, typeFilter, masterVehicleTypes]);

  useEffect(() => {
    if (activeVehicleId && !filteredVehicleRoutes.some((r) => r.vehicleId === activeVehicleId)) {
      setActiveVehicleId(
        filteredVehicleRoutes.length > 0 ? filteredVehicleRoutes[0].vehicleId : null
      );
    } else if (!activeVehicleId && filteredVehicleRoutes.length > 0) {
      setActiveVehicleId(filteredVehicleRoutes[0].vehicleId);
    }
  }, [filteredVehicleRoutes, activeVehicleId]);

  const activeRoute = useMemo(
    () => filteredVehicleRoutes.find((r) => r.vehicleId === activeVehicleId) || null,
    [filteredVehicleRoutes, activeVehicleId]
  );

  const getModalTitle = () => {
    switch (downloadType) {
      case 'routeTransaction':
        return 'Route Transaction';
      case 'deliveryList':
        return 'Delivery List';
      case 'deliveryForm':
        return 'Delivery Form';
      default:
        return 'Download Options';
    }
  };
  const searchPlaceholder = `${t('common.license_number')}, ${t('common.customer_name')}, ${t('common.invoice_number')}`;
  const headerItems = [
    {
      label: t('common.search'),
      component: (
        <SearchBar
          disabled={isLoading || isDownloading}
          onChange={setSearchQuery}
          placeholder={t('common.search')}
          tooltip={searchPlaceholder}
          value={searchQuery}
          width="w-full xl:w-70"
        />
      ),
    },
    {
      label: t('common.delivery_date'),
      component: (
        <CustomDatePicker
          id="estimasiDate"
          isLoading={isLoading || isDownloading}
          onChange={(d) => d && setSelectedDate(formatDateUniversal(d, 'YYYY-MM-DD'))}
          selected={selectedDate ? new Date(selectedDate) : new Date()}
          maxDate={tomorrowDate(false)}
          className="w-full"
        />
      ),
    },
    {
      label: t('common.storage_type'),
      component: (
        <StorageTypeFilter
          disabled={isLoading || isDownloading}
          onApply={setStorageFilter}
          selectedTypes={storageFilter}
        />
      ),
    },
    {
      label: t('common.vehicle_type'),
      component: (
        <VehicleTypeFilter
          disabled={isLoading || isDownloading}
          onApply={setTypeFilter}
          onMasterTypesLoad={setMasterVehicleTypes}
          selectedType={typeFilter}
        />
      ),
    },
    {
      label: t('delivery.view'),
      component: (
        <ToggleButton
          options={[
            { label: t('delivery.view_summary'), value: false },
            { label: t('common.detail'), value: true },
          ]}
          value={isDetailView}
          onChange={handleToggleView}
          disabled={isLoading || isDownloading}
          className="w-full"
        />
      ),
    },
    {
      label: 'Export',
      hideLabel: true,
      component: (
        <div className="w-full z-50 relative" ref={downloadDropdownRef}>
          <Button
            disabled={isLoading || isDownloading || isEmpty(filteredVehicleRoutes)}
            isLoading={isDownloading}
            onClick={() => setIsDownloadDropdownOpen((prev) => !prev)}
            text={t('common.download')}
          />
          {isDownloadDropdownOpen && (
            <div className="absolute right-0 mt-2 bg-white dark:bg-slate-700 rounded-md shadow-xl border border-gray-200 dark:border-slate-600 z-100 py-1.5 w-max min-w-[260px] flex flex-col">
              <div className="flex items-center justify-between hover:bg-slate-50 dark:hover:bg-slate-800/30 transition-colors">
                <button
                  onClick={() => handleDownloadTrigger('routeTransaction')}
                  className="w-full text-left px-3 py-2 text-sm font-medium text-slate-700 dark:text-slate-200 cursor-pointer"
                >
                  Route Transaction
                </button>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setIsRouteSettingsOpen(!isRouteSettingsOpen);
                  }}
                  className="px-3 py-2 text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-200 cursor-pointer transition-transform"
                >
                  <svg
                    className="w-4 h-4 transition-transform duration-200"
                    style={{
                      transform: isRouteSettingsOpen ? 'rotate(180deg)' : 'rotate(0deg)',
                    }}
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M19 9l-7 7-7-7"
                    />
                  </svg>
                </button>
              </div>

              {isRouteSettingsOpen && (
                <div className="ml-3 pl-3 border-l-2 border-slate-200 dark:border-slate-600 mb-1.5 animate-in slide-in-from-top-1 duration-200">
                  <div className="flex flex-col gap-2 pr-3 py-1">
                    <div className="flex items-center justify-between">
                      <label className="flex items-center gap-1.5 text-xs text-slate-500 dark:text-slate-400 cursor-pointer select-none">
                        <input
                          type="checkbox"
                          className="cursor-pointer w-3.5 h-3.5 rounded border-gray-300 dark:border-slate-600 text-sky-600 focus:ring-sky-500 focus:ring-offset-0"
                          checked={isNoBun}
                          onChange={(e) => handleToggleNoBun(e.target.checked)}
                        />
                        {t('delivery.no_bun')}
                        <InformationButton infoText={t('delivery.no_bun_info')} size="3.5" />
                      </label>
                      <button
                        onClick={() => {
                          setIsDownloadDropdownOpen(false);
                          setIsBunModalOpen(true);
                        }}
                        className="text-[10px] text-sky-600 hover:underline cursor-pointer font-medium"
                      >
                        {t('common.more')}
                      </button>
                    </div>
                    <label className="flex items-center gap-1.5 text-xs text-slate-500 dark:text-slate-400 cursor-pointer select-none">
                      <input
                        type="checkbox"
                        className="cursor-pointer w-3.5 h-3.5 rounded border-gray-300 dark:border-slate-600 text-sky-600 focus:ring-sky-500 focus:ring-offset-0"
                        checked={isSplitMultitrip}
                        onChange={(e) => handleToggleSplitMultitrip(e.target.checked)}
                      />
                      {t('delivery.split_multitrip')}
                      <InformationButton infoText={t('delivery.split_multitrip_info')} size="3.5" />
                    </label>
                    <label className="flex items-center gap-1.5 text-xs text-slate-500 dark:text-slate-400 cursor-pointer select-none">
                      <input
                        type="checkbox"
                        className="cursor-pointer w-3.5 h-3.5 rounded border-gray-300 dark:border-slate-600 text-sky-600 focus:ring-sky-500 focus:ring-offset-0"
                        checked={isSplitStorageType}
                        onChange={(e) => handleToggleSplitStorageType(e.target.checked)}
                      />
                      {t('delivery.split_storage')}
                      <InformationButton infoText={t('delivery.split_storage_info')} size="3.5" />
                    </label>
                  </div>
                </div>
              )}
              <button
                onClick={() => handleDownloadTrigger('deliveryList')}
                className="w-full text-left px-3 py-2 text-sm font-medium text-slate-700 dark:text-slate-200 cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-800/30 transition-colors border-t border-gray-100 dark:border-slate-600"
              >
                Delivery List
              </button>
              <button
                onClick={() => handleDownloadTrigger('deliveryForm')}
                className="w-full text-left px-3 py-2 text-sm font-medium text-slate-700 dark:text-slate-200 cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-800/30 transition-colors border-t border-gray-100 dark:border-slate-600"
              >
                Delivery Form
              </button>
            </div>
          )}
        </div>
      ),
    },
  ];

  const getBunExcludeList = () =>
    downloadType === 'routeTransaction' && isNoBun
      ? bunSoList.filter((b) => !b.hasNonBun).map((b) => b.so)
      : [];

  const tabData = useMemo(
    () =>
      filteredVehicleRoutes.map((r) => {
        const dName = r.assigneeName || '-';
        const isManual = r.hasManual;
        const hasMT = r.trips?.some((t) => t.isMiddleHub);
        const textClass = r.hasInvalidSo ? 'text-red-600 dark:text-red-400 font-bold' : '';

        return {
          id: r.vehicleId,
          label: (
            <Tooltip tooltipContent={isEmpty(dName) ? t('common.no_driver') : dName}>
              <span
                className={`block w-full h-full rounded px-2 py-0.5 border-2 transition-all relative ${isManual ? 'bg-[#E6EEFF] border-[#b3cbfe] dark:bg-blue-900/40 dark:border-blue-900' : 'bg-transparent border-transparent'} ${textClass}`}
              >
                {r.basePlat || r.vehicleName}{' '}
                {hasMT && (
                  <span className="text-violet-600 dark:text-violet-400 font-bold mr-1">[MT]</span>
                )}
                {r.isRedelivery && (
                  <span className="text-red-600 dark:text-red-400 font-bold">[R]</span>
                )}
              </span>
            </Tooltip>
          ),
        };
      }),
    [filteredVehicleRoutes, t]
  );

  if (!isClient) return null;
  return (
    <>
      <PageTemplate
        title={t('common.delivery')}
        subtitle={
          <>
            {t('delivery.subtitle')}{' '}
            <span className="font-semibold text-sky-600">{t('delivery.subtitle_highlight')}</span>
          </>
        }
        headerItems={headerItems}
        activeTabId={activeVehicleId}
        onTabClick={setActiveVehicleId}
        tabs={tabData}
        isLoading={isLoading}
        isEmpty={!isLoading && (isEmpty(filteredVehicleRoutes) || !activeRoute)}
        footer={{ text: t('common.click_for_detail') }}
        bodyProps={{ className: 'min-h-[400px]' }}
      >
        {!isLoading && activeRoute && (
          <CustomTable
            activeRoute={activeRoute}
            searchQuery={searchQuery}
            setSearchQuery={setSearchQuery}
            isDetailView={isDetailView}
            t={t}
            sortConfig={sortConfig}
            setSortConfig={setSortConfig}
            onRowClick={handleRowClick}
          />
        )}
      </PageTemplate>

      <PartialRoutingModal
        isOpen={isRoutingModalOpen}
        title={getModalTitle()}
        onClose={() => setIsRoutingModalOpen(false)}
        onPartial={(e) => {
          if (e) e.preventDefault();
          runPartialDownload(downloadType, getStoragePrefix(storageFilter), getBunExcludeList());
          setIsRoutingModalOpen(false);
        }}
        onFull={(e) => {
          if (e) e.preventDefault();
          runFullDownload(downloadType, getStoragePrefix(storageFilter), getBunExcludeList());
          setIsRoutingModalOpen(false);
        }}
        translate={t}
      />

      <BunListModal
        isOpen={isBunModalOpen}
        onClose={() => setIsBunModalOpen(false)}
        bunSoList={bunSoList}
        onDownload={handleDownloadBunSpecific}
        t={t}
      />

      <TaskModal
        isOpen={isTaskModalOpen}
        onClose={() => setIsTaskModalOpen(false)}
        taskId={selectedTaskId}
      />
    </>
  );
}
