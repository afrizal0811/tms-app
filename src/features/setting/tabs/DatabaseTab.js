'use client';

import Button from '@/components/button/Button';
import SearchBar from '@/components/SearchBar';
import {
  getDailyFleet,
  getKpiData,
  getLocationHistories,
  getResults,
  getTasks,
} from '@/lib/api/mileapp';
import { getLocalStorage } from '@/lib/localStorageHandler';
import { toastError, toastSuccess } from '@/lib/toast';
import { formatDateUniversal, parseCustomerString } from '@/lib/utils';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import Card from '../components/Card';
import { downloadDateRangeZip, downloadSingleDateExcel } from '../helper/databaseHelper';
import { KpiCards, LocationTable, RoutingTable, TaskTable, VehicleDetailsTable } from '../table';

export default function DatabaseTab({ hubId, hubName = '', dateRange = [], translate }) {
  const datesList = useMemo(() => {
    if (!dateRange || !dateRange[0]) return [];
    const list = [];
    const curr = new Date(dateRange[0]);
    curr.setHours(0, 0, 0, 0);
    const end = new Date(dateRange[1] || dateRange[0]);
    end.setHours(0, 0, 0, 0);
    while (curr <= end) {
      list.push(formatDateUniversal(curr, 'YYYY-MM-DD'));
      curr.setDate(curr.getDate() + 1);
    }
    return list;
  }, [dateRange]);

  const [currentDateIndex, setCurrentDateIndex] = useState(0);
  const [activeSubTab, setActiveSubTab] = useState('routing');
  const [searchQuery, setSearchQuery] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  const [currentData, setCurrentData] = useState({
    routing: [],
    tasks: [],
    locations: [],
    kpi: null,
    dailyFleet: [],
  });

  const cacheRef = useRef({});
  const selectedDate = datesList[currentDateIndex] || datesList[0] || '';
  const isSunday = useMemo(() => {
    if (!selectedDate) return false;
    return new Date(`${selectedDate}T00:00:00`).getDay() === 0;
  }, [selectedDate]);
  const emptyText = isSunday ? translate('common.holiday_sunday') : undefined;

  const hubAcronym = useMemo(() => {
    const { storedLocationAcronym, storedLocationName } = getLocalStorage();
    return storedLocationAcronym || hubName || storedLocationName || '';
  }, [hubName]);

  const fetchDateData = useCallback(
    async (dateStr) => {
      if (!dateStr) return { routing: [], tasks: [], locations: [], kpi: null, dailyFleet: [] };
      if (cacheRef.current[dateStr]) {
        return cacheRef.current[dateStr];
      }

      const [rRes, tRes, lRes, kRes, dfRes] = await Promise.all([
        getResults({ dateFrom: dateStr, dateTo: dateStr, hubId }).catch(() => []),
        getTasks({ deliveryDate: dateStr, hubId, isShowAll: true }).catch(() => []),
        getLocationHistories({ dateFrom: dateStr, dateTo: dateStr, hubId }).catch(() => []),
        getKpiData({ date: dateStr, hubId }).catch(() => null),
        getDailyFleet({ date: dateStr, hubId }).catch(() => []),
      ]);

      const routing = Array.isArray(rRes) ? rRes : [];
      const rawTasks = Array.isArray(tRes) ? tRes : [];
      const tasks = rawTasks
        .map((t) => {
          const cust = parseCustomerString(t.customerOrder || t.customer_order);
          return {
            ...t,
            customerName: cust.name || cust.fullCustomerName || t.customerName || '',
            truncateInvoice: cust.truncateInvoice || cust.invoiceNumber || t.orderId || '',
            statusDelivery: t.statusDelivery || t.status_delivery || t.status || '',
            driverName: t.driverName || t.driver_name || '',
          };
        })
        .sort((a, b) => (a.driverName || '').localeCompare(b.driverName || ''));

      const rawLocations = Array.isArray(lRes) ? lRes : [];
      const locations = rawLocations
        .map((l) => ({
          ...l,
          driverName: l.driver_name || l.driverName || '',
        }))
        .sort((a, b) => a.driverName.localeCompare(b.driverName));
      const kpi = Array.isArray(kRes) ? kRes[0] || null : kRes || null;
      const dailyFleet = Array.isArray(dfRes) ? dfRes : [];

      const result = { routing, tasks, locations, kpi, dailyFleet };
      cacheRef.current[dateStr] = result;
      return result;
    },
    [hubId]
  );

  const loadActiveDate = useCallback(
    async (dateStr) => {
      if (!dateStr) return;
      setIsLoading(true);
      try {
        const data = await fetchDateData(dateStr);
        setCurrentData(data);
      } catch (err) {
        toastError(err.message, err);
      } finally {
        setIsLoading(false);
      }
    },
    [fetchDateData]
  );

  useEffect(() => {
    if (selectedDate) {
      setSearchQuery('');
      loadActiveDate(selectedDate);
    }
  }, [selectedDate, loadActiveDate]);

  const hasNoData = useMemo(() => {
    return (
      currentData.routing.length === 0 &&
      currentData.tasks.length === 0 &&
      currentData.locations.length === 0 &&
      currentData.dailyFleet.length === 0 &&
      !currentData.kpi
    );
  }, [currentData]);

  const handleExportZip = async () => {
    if (datesList.length === 0 || hasNoData) return;
    setIsExporting(true);
    try {
      await downloadDateRangeZip({
        datesList,
        fetchDateData,
        hubAcronym,
        translate,
      });
      toastSuccess(translate('common.toast.success'));
    } catch (err) {
      toastError(err.message, err);
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportCurrent = () => {
    if (!selectedDate || hasNoData) return;
    try {
      downloadSingleDateExcel({
        dateStr: selectedDate,
        data: currentData,
        hubAcronym,
        translate,
      });
      toastSuccess(translate('common.toast.success'));
    } catch (err) {
      toastError(err.message, err);
    }
  };

  const subTabs = [
    { key: 'routing', label: translate('common.routing') },
    { key: 'task', label: translate('common.task') },
    {
      key: 'vehicle_details',
      label: 'Vehicle Details',
    },
    {
      key: 'location',
      label: translate('common.driver_location'),
    },
    {
      key: 'kpi',
      label: 'KPI',
    },
  ];

  const renderActiveTable = () => {
    switch (activeSubTab) {
      case 'routing':
        return (
          <RoutingTable
            data={currentData.routing}
            searchQuery={searchQuery}
            isLoading={isLoading}
            emptyText={emptyText}
            translate={translate}
          />
        );
      case 'task':
        return (
          <TaskTable
            data={currentData.tasks}
            searchQuery={searchQuery}
            isLoading={isLoading}
            emptyText={emptyText}
            translate={translate}
          />
        );
      case 'vehicle_details':
        return (
          <VehicleDetailsTable
            data={currentData.dailyFleet}
            searchQuery={searchQuery}
            isLoading={isLoading}
            emptyText={emptyText}
            translate={translate}
          />
        );
      case 'location':
        return (
          <LocationTable
            data={currentData.locations}
            searchQuery={searchQuery}
            isLoading={isLoading}
            emptyText={emptyText}
            translate={translate}
          />
        );
      case 'kpi':
        return <KpiCards kpi={currentData.kpi} emptyText={emptyText} translate={translate} />;
      default:
        return null;
    }
  };

  return (
    <div className="w-full flex flex-col h-full gap-4">
      <Card className="flex flex-col min-h-[620px] max-h-[620px] p-0!">
        <div className="flex flex-col md:flex-row justify-between items-stretch md:items-center gap-3 p-4 border-b border-gray-100 dark:border-slate-800">
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center gap-2 overflow-x-auto p-1 bg-slate-100 dark:bg-slate-800/80 rounded-lg border border-gray-200 dark:border-slate-700 max-w-full">
              {subTabs.map((st) => (
                <button
                  key={st.key}
                  type="button"
                  onClick={() => setActiveSubTab(st.key)}
                  className={`inline-flex items-center justify-center h-8 px-3 rounded-md text-xs font-semibold whitespace-nowrap transition-colors cursor-pointer shrink-0 leading-none ${
                    activeSubTab === st.key
                      ? 'bg-white dark:bg-slate-900 text-sky-600 dark:text-sky-400 shadow-sm'
                      : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
                  }`}
                >
                  {st.label}
                </button>
              ))}
            </div>
            {datesList.length > 1 && (
              <div className="flex items-center gap-2 h-10 px-2 bg-slate-100 dark:bg-slate-800/80 rounded-lg border border-gray-200 dark:border-slate-700">
                <button
                  type="button"
                  onClick={() => setCurrentDateIndex((prev) => Math.max(0, prev - 1))}
                  disabled={currentDateIndex === 0}
                  className="inline-flex items-center justify-center h-7 px-2 text-xs font-bold rounded border border-gray-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors cursor-pointer"
                >
                  &lt;
                </button>

                <span className="text-xs font-semibold text-slate-700 dark:text-slate-200 px-1 whitespace-nowrap">
                  {formatDateUniversal(selectedDate, 'DD/MM/YY')}
                </span>

                <button
                  type="button"
                  onClick={() =>
                    setCurrentDateIndex((prev) => Math.min(datesList.length - 1, prev + 1))
                  }
                  disabled={currentDateIndex >= datesList.length - 1}
                  className="inline-flex items-center justify-center h-7 px-2 text-xs font-bold rounded border border-gray-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors cursor-pointer"
                >
                  &gt;
                </button>
              </div>
            )}
          </div>
          <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
            {activeSubTab !== 'kpi' && (
              <div className="w-full ">
                <SearchBar
                  value={searchQuery}
                  onChange={setSearchQuery}
                  placeholder={translate('common.search')}
                  size="sm"
                />
              </div>
            )}
            <Button
              type={datesList.length > 1 ? 'split' : 'button'}
              onClick={datesList.length > 1 ? handleExportZip : handleExportCurrent}
              isLoading={isExporting}
              disabled={isExporting || hasNoData}
              width="w-full sm:w-auto"
              size="sm"
              text={translate('common.download')}
              options={[
                {
                  id: 'current',
                  label: `${translate('common.download')} (${formatDateUniversal(selectedDate, 'DD/MM/YY')})`,
                  onClick: handleExportCurrent,
                  disabled: isExporting || hasNoData,
                },
              ]}
            />
          </div>
        </div>

        <div className="flex-1 overflow-hidden flex flex-col min-h-0">{renderActiveTable()}</div>
      </Card>
    </div>
  );
}
