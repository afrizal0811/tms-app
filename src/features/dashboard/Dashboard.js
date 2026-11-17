'use client';

import CustomDatePicker from '@/components/CustomDatePicker';
import StorageTypeFilter from '@/components/dropdown/StorageTypeFilter';
import VehicleTypeFilter from '@/components/dropdown/VehicleTypeFilter';
import PageTemplate from '@/components/page/PageTemplate';
import { useLanguage } from '@/context/LanguageContext';
import DetailTab from '@/features/dashboard/tab/DetailTab';
import RoutingActualTab from '@/features/dashboard/tab/RoutingActualTab';
import { getPendingDetails, getResults, getTasks } from '@/lib/api/mileapp';
import { applyPendingDetailsToTasks } from '@/lib/deliveryStatus';
import { getCachedHubs, getLocalStorage } from '@/lib/localStorageHandler';
import { toastError } from '@/lib/toast';
import {
  formatDateUniversal,
  getBaseVehicleType,
  isEmpty,
  normalizeEmail,
  toApiDateString,
  tomorrowDate,
} from '@/lib/utils';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { calculateDashboard } from './helper/help';

export default function Dashboard({ driverData }) {
  const { t, isIndonesian } = useLanguage();
  const [selectedDate, setSelectedDate] = useState(new Date());
  const [loading, setLoading] = useState(true);
  const [storageFilter, setStorageFilter] = useState(['DRY', 'FROZEN']);
  const [typeFilter, setTypeFilter] = useState('');
  const [masterVehicleTypes, setMasterVehicleTypes] = useState([]);
  const [rawData, setRawData] = useState({ tasks: [], results: [] });
  const [activeTab, setActiveTab] = useState('Detail');
  const [hasPendingGR, setHasPendingGR] = useState(false);
  const [dismissedDots, setDismissedDots] = useState({
    Detail: false,
    RoutingVsActual: false,
  });

  const fetchStartTimeRef = useRef(null);

  const { storedLocation: hubId } = getLocalStorage();

  useEffect(() => {
    try {
      const hubs = getCachedHubs();
      const activeHub = hubs?.activeHub;
      if (activeHub) {
        setHasPendingGR(activeHub.hasPendingGR || false);
      }
    } catch (error) {
      console.error('Error Hubs:', error);
    }
  }, [hubId, driverData]);

  const handleDateChange = (date) => {
    if (!date) return;
    setSelectedDate(date);
  };

  const getPingDot = (tabId) => {
    const dismissed = dismissedDots[tabId];
    if (!loading && dismissed) return null;
    return (
      <span className="inline-flex items-center ml-2" aria-hidden>
        {loading ? (
          <span className="relative flex h-3 w-3">
            <span className="absolute inline-flex h-full w-full rounded-full bg-amber-300 opacity-75 animate-ping" />
            <span className="relative inline-flex rounded-full h-3 w-3 bg-amber-400" />
          </span>
        ) : (
          <span className="inline-flex rounded-full h-2.5 w-2.5 bg-sky-500" />
        )}
      </span>
    );
  };

  const handleTabClick = (tabId) => {
    setActiveTab(tabId);
    if (!loading) {
      setDismissedDots((prev) => ({ ...prev, [tabId]: true }));
    }
  };

  useEffect(() => {
    if (loading) {
      setDismissedDots((prev) => ({ ...prev, Detail: false, RoutingVsActual: false }));
    }
  }, [loading]);

  const fetchData = useCallback(async () => {
    if (isEmpty(driverData) || selectedDate.getDay() === 0) {
      setRawData({ tasks: [], results: [] });
      return;
    }

    try {
      if (typeof window === 'undefined') return;

      const localStart = new Date(selectedDate);
      localStart.setHours(0, 0, 0, 0);
      const localEnd = new Date(localStart);
      localEnd.setHours(23, 59, 59, 999);

      const timeFrom = toApiDateString(localStart);
      const timeTo = toApiDateString(localEnd);

      setLoading(true);
      fetchStartTimeRef.current = Date.now();

      const dateStr = formatDateUniversal(localStart, 'YYYY-MM-DD');

      const [tasksData, resultsData, pendingDetailsData] = await Promise.all([
        getTasks({
          status: 'DONE,ONGOING,UNASSIGNED',
          timeFrom,
          timeTo,
        }),
        getResults({
          deliveryDate: dateStr,
        }),
        getPendingDetails(dateStr, dateStr).catch(() => []),
      ]);

      const tasksArray = Array.isArray(tasksData) ? tasksData : tasksData?.data || [];
      const pdList = Array.isArray(pendingDetailsData)
        ? pendingDetailsData
        : Array.isArray(pendingDetailsData?.data)
          ? pendingDetailsData.data
          : [];
      const effectiveTasks = applyPendingDetailsToTasks(tasksArray, pdList);

      setRawData({ tasks: effectiveTasks, results: resultsData });
    } catch (err) {
      toastError(t('common.toast.error', { err: err.message }), err);
    } finally {
      setLoading(false);
    }
  }, [selectedDate, t, driverData]);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  useEffect(() => {
    const handlePendingUpdated = () => {
      fetchData();
    };
    const handleStorage = (e) => {
      if (e.key === 'tms_pending_details_last_update') {
        fetchData();
      }
    };
    window.addEventListener('tms-pending-details-updated', handlePendingUpdated);
    window.addEventListener('storage', handleStorage);
    return () => {
      window.removeEventListener('tms-pending-details-updated', handlePendingUpdated);
      window.removeEventListener('storage', handleStorage);
    };
  }, [fetchData]);

  const driverMap = useMemo(() => {
    const map = new Map();
    driverData?.forEach((driver) => {
      if (driver.email && driver.name) {
        map.set(normalizeEmail(driver.email), driver.name);
      }
    });
    return map;
  }, [driverData]);

  const driverTypeLookup = useMemo(() => {
    if (!driverData) return { byEmail: new Map(), byPlat: new Map() };
    const byEmail = new Map();
    const byPlat = new Map();
    driverData.forEach((dr) => {
      if (dr.email) byEmail.set(normalizeEmail(dr.email), dr);
      if (dr.plat) byPlat.set((dr.plat || '').replace(/\s+/g, '').toLowerCase(), dr);
    });
    return { byEmail, byPlat };
  }, [driverData]);

  const applyFilters = useCallback(
    (tasks) => {
      if (isEmpty(tasks)) return [];
      let filtered = tasks;
      if (storageFilter.length === 0) return [];
      if (storageFilter.length === 1) {
        filtered = filtered.filter((t) =>
          storageFilter.includes((t.typeStorage || '').toUpperCase())
        );
      }
      if (typeFilter && typeFilter !== 'all') {
        filtered = filtered.filter((t) => {
          const rawEmail = t.assignee || t.assignedTo?.email || t.doneBy;
          const email =
            typeof rawEmail === 'string' && rawEmail.includes(',')
              ? rawEmail.split(',')[0].trim()
              : rawEmail;
          const normalizedEmail = normalizeEmail(email);
          const plat = t.basePlat || t.vehicleName || t.plat;
          const platNorm = (plat || '').replace(/\s+/g, '').toLowerCase();

          const d =
            (normalizedEmail && driverTypeLookup.byEmail.get(normalizedEmail)) ||
            (platNorm && driverTypeLookup.byPlat.get(platNorm));

          if (!d) return false;
          return getBaseVehicleType(d.type, masterVehicleTypes) === typeFilter;
        });
      }
      return filtered;
    },
    [storageFilter, typeFilter, driverTypeLookup, masterVehicleTypes]
  );

  const filteredDailyTasks = useMemo(
    () => applyFilters(rawData.tasks),
    [applyFilters, rawData.tasks]
  );

  const summaryData = useMemo(() => {
    return calculateDashboard(filteredDailyTasks, driverMap, isIndonesian, hasPendingGR);
  }, [filteredDailyTasks, driverMap, isIndonesian, hasPendingGR]);

  let isCardEmpty = false;
  let emptyMessage = t('common.no_data');

  if (isEmpty(driverData) && !loading) {
    isCardEmpty = true;
    emptyMessage = t('common.no_driver');
  } else if (activeTab === 'Detail') {
    isCardEmpty =
      !loading &&
      (!summaryData ||
        isEmpty(summaryData.totalTasks) ||
        (isEmpty(summaryData.done) && isEmpty(summaryData.ongoing)));
  } else if (activeTab === 'RoutingVsActual') {
    const noOngoingAndDone = !filteredDailyTasks?.some(
      (t) => t?.status === 'ONGOING' || t?.status === 'DONE'
    );
    isCardEmpty = !loading && noOngoingAndDone;
  }

  const headerItems = [
    {
      label: t('common.storage_type'),
      component: (
        <StorageTypeFilter
          disabled={loading}
          onApply={setStorageFilter}
          selectedTypes={storageFilter}
        />
      ),
      hideLabel: false,
    },
    {
      label: t('common.vehicle_type'),
      hideLabel: false,
      component: (
        <VehicleTypeFilter
          disabled={loading}
          onApply={setTypeFilter}
          onMasterTypesLoad={setMasterVehicleTypes}
          selectedType={typeFilter}
        />
      ),
    },
    {
      label: t('common.delivery_date'),
      component: (
        <CustomDatePicker
          selected={selectedDate}
          onChange={handleDateChange}
          isLoading={loading}
          dateFormat="dd MMMM yyyy"
          maxDate={tomorrowDate()}
          disableSunday={true}
        />
      ),
      hideLabel: false,
    },
  ];

  const cardTabs = [
    { id: 'Detail', label: t('common.detail'), extraContent: getPingDot('Detail') },
    {
      id: 'RoutingVsActual',
      label: t('common.routing_vs_actual'),
      extraContent: getPingDot('RoutingVsActual'),
    },
  ];
  return (
    <PageTemplate
      title="Dashboard"
      subtitle={
        <>
          {t('dashboard.subtitle')}{' '}
          <span className="font-semibold text-sky-600">{t('dashboard.subtitle_highlight')}</span>
        </>
      }
      headerItems={headerItems}
      activeTabId={activeTab}
      onTabClick={handleTabClick}
      tabs={cardTabs}
      isLoading={loading}
      isEmpty={isCardEmpty}
      emptyMessage={emptyMessage}
      footer={
        activeTab === 'RoutingVsActual' && {
          text: t('common.click_for_detail'),
        }
      }
      bodyProps={{
        timerStartTime: fetchStartTimeRef.current,
      }}
    >
      <div className="flex-1 flex flex-col p-3 overflow-hidden dark:bg-slate-800">
        {activeTab === 'Detail' && (
          <DetailTab loading={loading} summaryData={summaryData} driverData={driverData} />
        )}

        {activeTab === 'RoutingVsActual' && (
          <RoutingActualTab
            loading={loading}
            tasks={filteredDailyTasks}
            results={rawData.results}
            drivers={driverData}
            selectedDate={selectedDate}
            hasPendingGR={hasPendingGR}
          />
        )}
      </div>
    </PageTemplate>
  );
}
