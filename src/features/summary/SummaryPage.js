'use client';

import Button from '@/components/button/Button';
import CustomDatePicker from '@/components/CustomDatePicker';
import ConfirmModal from '@/components/modal/ConfirmModal';
import PageTemplate from '@/components/page/PageTemplate';
import ToggleButton from '@/components/button/ToggleButton';
import { useLanguage } from '@/context/LanguageContext';
import { getHubs, getPendingDetails, getReasons } from '@/lib/api/mileapp';
import {
  applyPendingDetailsToTasks,
  applyPendingDetailsToTaskSummaryMetrics,
  DELIVERY_STATUS_KEYS,
  normalizeDeliveryStatus,
} from '@/lib/deliveryStatus';
import useSummaryData from '@/lib/hooks/useSummaryData';
import { generateSummaryWorkbook } from '@/lib/reportGenerators/summary/summaryReport';
import { calculatePendingReasonData } from '@/lib/reportGenerators/summary/sheets/pendingReasonSheet';
import { calculateTruckDetailData } from '@/lib/reportGenerators/summary/sheets/truckDetailSheet';
import { getTargetRoutingDate } from '@/lib/reportGenerators/summary/sheets/routingTimeSheet';
import { toastError, toastSuccess, toastWarning } from '@/lib/toast';
import { formatDateUniversal, isEmpty } from '@/lib/utils';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx-js-style';
import DistanceSummaryTab from './tabs/DistanceSummaryTab';
import PendingReasonsTab from './tabs/PendingReasonsTab';
import RoutingTimeTab from './tabs/RoutingTimeTab';
import ServiceLevelTab from './tabs/ServiceLevelTab';
import TaskSummaryTab from './tabs/TaskSummaryTab';
import TimeDriverTab from './tabs/TimeDriverTab';
import TruckDetailTab from './tabs/TruckDetailTab';
import TruckUsageTab from './tabs/TruckUsageTab';
import DiagramTab from './tabs/DiagramTab';

const colorLegend = {
  truckDetail: [
    {
      name: 'orange',
      colors: 'bg-orange-400 text-white dark:bg-orange-600',
      text: 'summary.tabs.truck_detail.orange',
    },
    {
      name: 'blue',
      colors: 'bg-[#4F76C7] text-white dark:bg-[#325296]',
      text: 'summary.tabs.truck_detail.blue',
    },
    {
      name: 'magenta',
      colors: 'bg-[#C85D86] text-white dark:bg-[#964263]',
      text: 'summary.tabs.truck_detail.magenta',
    },
    {
      name: 'indigo',
      colors: 'bg-[#5C5FB2] text-white dark:bg-[#45488c]',
      text: 'summary.tabs.truck_detail.indigo',
    },
  ],
  pendingReasons: [
    {
      name: 'orange',
      colors: 'bg-amber-100 dark:bg-amber-400/30 text-white',
      text: 'summary.tabs.pending_reasons.legend_orange',
    },
  ],
};

export default function SummaryPage() {
  const { t, localeCode, isIndonesian } = useLanguage();
  const {
    selectedLocation,
    dateRange,
    setDateRange,
    driverData,
    rawData,
    isLoading,
    reportPreview,
    pendingEndpoints,
    taskSummaryMetrics,
    isCalculatingMetrics,
    historyProgress,
    masterTruckData,
    fetchData,
    dismissedDots,
    setDismissedDots,
    activeHubLocation,
  } = useSummaryData();
  const [activeTab, setActiveTab] = useState('Diagram');
  const [showWarningModal, setShowWarningModal] = useState(false);
  const [pendingDateRange, setPendingDateRange] = useState([null, null]);
  const [tempDateRange, setTempDateRange] = useState(dateRange || [null, null]);
  const [hasPendingGR, setHasPendingGR] = useState(false);
  const [reasons, setReasons] = useState([]);
  const [pendingDetails, setPendingDetails] = useState([]);
  const [pendingDetailsRange, setPendingDetailsRange] = useState(null);
  const [isDownload, setIsDownload] = useState(false);
  const [emptyMessage, setEmptyMessage] = useState('');
  const [showConvertedSuccess, setShowConvertedSuccess] = useState(false);
  const [isAddPendingModalOpen, setIsAddPendingModalOpen] = useState(false);

  const driverChecker = useCallback(
    (data) => {
      if (isEmpty(data)) {
        setEmptyMessage(t('common.no_driver'));
        return;
      }

      setEmptyMessage('');
    },
    [t]
  );

  useEffect(() => {
    driverChecker(driverData);
    getReasons()
      .then(setReasons)
      .catch(() => {});
  }, [driverData, driverChecker, t]);

  const warnedRangeKeyRef = useRef(null);

  useEffect(() => {
    let isCancelled = false;
    driverChecker(driverData);
    if (dateRange && dateRange[0] && dateRange[1]) {
      const start = formatDateUniversal(new Date(dateRange[0]));
      const end = formatDateUniversal(new Date(dateRange[1]));
      setPendingDetailsRange(null);
      getPendingDetails(start, end)
        .then((res) => {
          if (!isCancelled) {
            setPendingDetails(res || []);
            setPendingDetailsRange(`${start}_${end}`);
          }
        })
        .catch(() => {
          if (!isCancelled) {
            setPendingDetails([]);
            setPendingDetailsRange(`${start}_${end}`);
          }
        });
    }
    return () => {
      isCancelled = true;
    };
  }, [dateRange, driverChecker, driverData, t]);

  useEffect(() => {
    if (!dateRange || !dateRange[0] || !dateRange[1]) return;

    const start = formatDateUniversal(new Date(dateRange[0]));
    const end = formatDateUniversal(new Date(dateRange[1]));
    const currentRangeKey = `${start}_${end}`;

    const isReportReady =
      !isLoading &&
      reportPreview &&
      reportPreview.startDateStr === start &&
      reportPreview.endDateStr === end;

    const isPendingDetailsReady = pendingDetailsRange === currentRangeKey;

    if (!isReportReady || !isPendingDetailsReady) return;

    if (warnedRangeKeyRef.current === currentRangeKey) return;

    warnedRangeKeyRef.current = currentRangeKey;

    const pendingData = reportPreview?.pendingReasonsData || [];
    const uncompletedCount = pendingData.filter((item) => {
      const pd = (pendingDetails || []).find((d) => d.taskId === item._id) || {};
      const status = pd.new_status
        ? normalizeDeliveryStatus(pd.new_status)
        : normalizeDeliveryStatus(item.status);
      if (status === DELIVERY_STATUS_KEYS.SUCCESS) return false;
      const hasAnyData = Boolean(
        (pd.groupReason && !isEmpty(pd.groupReason) && pd.groupReason !== '-') ||
          (pd.detailReason && !isEmpty(pd.detailReason) && pd.detailReason !== '-') ||
          (pd.internalExternal && !isEmpty(pd.internalExternal) && pd.internalExternal !== '-') ||
          (pd.pic && !isEmpty(pd.pic) && pd.pic !== '-')
      );
      return !hasAnyData;
    }).length;

    if (uncompletedCount > 0) {
      const template =
        t('summary.tabs.pending_reasons.toast_uncompleted', { count: uncompletedCount }) ||
        'Terdapat {count} data pending yang perlu dilengkapi!';
      const message = template.replace('{count}', uncompletedCount);
      toastWarning(message);
    }
  }, [isLoading, reportPreview, pendingDetails, pendingDetailsRange, dateRange, t]);

  const handleUpdatePendingDetail = (updatedDetail) => {
    setPendingDetails((prev) => {
      if (updatedDetail.deleted) {
        return prev.filter((p) => p.taskId !== updatedDetail.taskId);
      }
      const exists = prev.find((p) => p.taskId === updatedDetail.taskId);
      if (exists) {
        return prev.map((p) => (p.taskId === updatedDetail.taskId ? updatedDetail : p));
      }
      return [...prev, updatedDetail];
    });
  };

  useEffect(() => {
    const fetchHubSettings = async () => {
      if (!selectedLocation) return;
      driverChecker(driverData);
      try {
        const hubs = await getHubs();
        if (hubs?.activeHub) {
          setHasPendingGR(hubs.activeHub.hasPendingGR || false);
        }
      } catch (error) {
        console.error('Error Hubs:', error);
      }
    };
    fetchHubSettings();
  }, [selectedLocation, driverData, driverChecker, t]);

  const handleTempDateChange = (update) => {
    setTempDateRange(update);
  };

  const handleApplyDate = () => {
    const [start, end] = tempDateRange;
    if (!start || !end) return;

    const diffTime = Math.abs(end - start);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

    if (diffDays > 14) {
      setPendingDateRange(tempDateRange);
      setShowWarningModal(true);
    } else {
      setDateRange(tempDateRange);
    }
  };

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  const effectiveTasks = useMemo(() => {
    return applyPendingDetailsToTasks(rawData.tasks, pendingDetails);
  }, [rawData.tasks, pendingDetails]);

  useEffect(() => {
    const handlePendingUpdated = (e) => {
      const saved = e?.detail;
      if (saved && saved.taskId) {
        handleUpdatePendingDetail(saved);
      } else if (dateRange && dateRange[0] && dateRange[1]) {
        const start = formatDateUniversal(new Date(dateRange[0]));
        const end = formatDateUniversal(new Date(dateRange[1]));
        getPendingDetails(start, end)
          .then((res) => setPendingDetails(res || []))
          .catch(() => {});
      }
    };
    const handleStorage = (e) => {
      if (
        e.key === 'tms_pending_details_last_update' &&
        dateRange &&
        dateRange[0] &&
        dateRange[1]
      ) {
        const start = formatDateUniversal(new Date(dateRange[0]));
        const end = formatDateUniversal(new Date(dateRange[1]));
        getPendingDetails(start, end)
          .then((res) => setPendingDetails(res || []))
          .catch(() => {});
      }
    };
    window.addEventListener('tms-pending-details-updated', handlePendingUpdated);
    window.addEventListener('storage', handleStorage);
    return () => {
      window.removeEventListener('tms-pending-details-updated', handlePendingUpdated);
      window.removeEventListener('storage', handleStorage);
    };
  }, [dateRange]);

  const effectiveTaskSummaryMetrics = useMemo(() => {
    return applyPendingDetailsToTaskSummaryMetrics(
      taskSummaryMetrics,
      effectiveTasks,
      pendingDetails,
      hasPendingGR,
      driverData
    );
  }, [taskSummaryMetrics, effectiveTasks, pendingDetails, hasPendingGR, driverData]);

  const effectiveTruckDetailData = useMemo(() => {
    if (!reportPreview?.truckDetailData) return null;
    if (!pendingDetails || pendingDetails.length === 0) {
      return reportPreview.truckDetailData;
    }
    const [start, end] = dateRange || [];
    if (!start || !end) return reportPreview.truckDetailData;
    const startStr = formatDateUniversal(new Date(start));
    const endStr = formatDateUniversal(new Date(end));
    const activeHubCoordsStr = activeHubLocation
      ? `${activeHubLocation.lat},${activeHubLocation.lng}`
      : null;
    const truckDetailRaw = calculateTruckDetailData(
      driverData,
      rawData.results,
      effectiveTasks,
      startStr,
      endStr,
      localeCode,
      activeHubCoordsStr
    );
    return {
      ...truckDetailRaw,
      driverMap: Object.fromEntries(truckDetailRaw.driverMap),
    };
  }, [
    reportPreview?.truckDetailData,
    pendingDetails,
    dateRange,
    effectiveTasks,
    rawData.results,
    driverData,
    localeCode,
    activeHubLocation,
  ]);

  const effectivePendingReasonsData = useMemo(() => {
    if (!reportPreview?.pendingReasonsData) return [];
    if (!pendingDetails || pendingDetails.length === 0) {
      return reportPreview.pendingReasonsData;
    }

    const hasAnyStatusChange = (pendingDetails || []).some((pd) => {
      if (!pd.new_status) return false;
      const task = (rawData.tasks || []).find(
        (t) => String(t._id || t.id || t.taskId) === String(pd.taskId)
      );
      if (!task) return true;
      const rawOrig = task.originalStatus || task.statusDelivery || task.status;
      const orig = Array.isArray(rawOrig) ? rawOrig[0] : rawOrig;
      return normalizeDeliveryStatus(orig) !== normalizeDeliveryStatus(pd.new_status);
    });

    if (!hasAnyStatusChange) {
      return reportPreview.pendingReasonsData;
    }

    const startStr = dateRange && dateRange[0] ? formatDateUniversal(new Date(dateRange[0])) : '';
    const endStr = dateRange && dateRange[1] ? formatDateUniversal(new Date(dateRange[1])) : '';

    return calculatePendingReasonData(driverData, effectiveTasks, startStr, endStr);
  }, [
    reportPreview?.pendingReasonsData,
    pendingDetails,
    rawData.tasks,
    dateRange,
    driverData,
    effectiveTasks,
  ]);

  const handleDownloadExcel = async () => {
    try {
      if (!dateRange || !dateRange[0] || !dateRange[1]) return;
      driverChecker(driverData);
      setIsDownload(true);
      const startDate = new Date(dateRange[0]);
      const endDate = new Date(dateRange[1]);
      const effectiveTasks = applyPendingDetailsToTasks(rawData.tasks, pendingDetails);
      const { wb, excelFileName } = await generateSummaryWorkbook(
        driverData,
        effectiveTasks,
        rawData.results,
        rawData.locations,
        formatDateUniversal(startDate),
        formatDateUniversal(endDate),
        effectiveTaskSummaryMetrics,
        masterTruckData,
        t,
        localeCode,
        hasPendingGR,
        pendingDetails
      );
      XLSX.writeFile(wb, excelFileName);
      toastSuccess(t('common.toast.success'));
    } catch (e) {
      toastError(t('common.toast.error', { err: e.message }), e);
    } finally {
      setIsDownload(false);
    }
  };

  const getPingDot = (tabId) => {
    const dismissed = dismissedDots[tabId];
    if (!isLoading && dismissed) return null;
    return (
      <span className="inline-flex items-center ml-2" aria-hidden>
        {isLoading ? (
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
    if (!isLoading) setDismissedDots((prev) => ({ ...prev, [tabId]: true }));
  };

  const isTabEmpty = () => {
    if (isLoading) return false;
    if (isEmpty(driverData)) return true;
    const parseDateRange = (dateRange) => {
      const [start, end] = dateRange || [];
      return {
        startStr: start ? formatDateUniversal(new Date(start), 'YYYY-MM-DD') : null,
        endStr: end ? formatDateUniversal(new Date(end), 'YYYY-MM-DD') : null,
      };
    };
    const isInDateRange = (dateKey, startStr, endStr) => dateKey >= startStr && dateKey <= endStr;

    switch (activeTab) {
      case 'Routing Time': {
        const { startStr, endStr } = parseDateRange(dateRange);
        return !rawData.tasks?.some((t) => {
          const isApi = !t.createdFrom || t.createdFrom === 'API';
          if (!isApi) return false;
          if (!startStr || !endStr) return true;
          const assignedDate = getTargetRoutingDate(t.startTime, t.createdTime);
          return assignedDate && isInDateRange(assignedDate, startStr, endStr);
        });
      }
      case 'Task Summary': {
        const metricKeys = Object.keys(taskSummaryMetrics);
        if (isEmpty(metricKeys)) return true;
        const { startStr, endStr } = parseDateRange(dateRange);
        if (!startStr || !endStr) return false;
        return !metricKeys.some((dateKey) => isInDateRange(dateKey, startStr, endStr));
      }
      case 'Pending Reasons':
        return isEmpty(effectivePendingReasonsData || reportPreview?.pendingReasonsData || []);
      case 'Time Driver': {
        const { dataMatrix, driverEmails } = reportPreview?.timeDriverData || {};
        return !(
          dataMatrix &&
          Object.values(dataMatrix).some((d) => Object.keys(d).length > 0) &&
          driverEmails?.length > 0
        );
      }
      case 'Truck Detail': {
        const dataMatrix = (effectiveTruckDetailData || reportPreview?.truckDetailData)?.dataMatrix;
        if (!dataMatrix || isEmpty(Object.keys(dataMatrix))) return true;
        const { startStr, endStr } = parseDateRange(dateRange);
        if (!startStr || !endStr) return false;
        return !Object.entries(dataMatrix).some(
          ([dateKey, val]) =>
            isInDateRange(dateKey, startStr, endStr) && Object.keys(val).length > 0
        );
      }
      case 'Truck Usage': {
        const dateMap = reportPreview?.truckUsageData?.dateMap;
        return !(
          dateMap &&
          Object.values(dateMap).some(
            (d) => (d.DryTotal || 0) > 0 || (d.FrozenTotal || 0) > 0 || (d.OTV || 0) > 0
          )
        );
      }
      case 'Distance Summary': {
        const data = reportPreview?.distanceSummaryData;
        return !data || isEmpty(data) || !data.some((row) => (row.totalKm || 0) > 0);
      }
      case 'Service Level':
        return isEmpty(rawData.tasks);
      case 'Diagram':
        return isEmpty(rawData.tasks);
      default:
        return false;
    }
  };

  const renderContent = () => {
    const renderTab = (Component, props) => (
      <div className="w-full h-full flex flex-col">
        <Component {...props} />
      </div>
    );

    const startStr = dateRange && dateRange[0] ? formatDateUniversal(new Date(dateRange[0])) : '';
    const endStr = dateRange && dateRange[1] ? formatDateUniversal(new Date(dateRange[1])) : '';

    switch (activeTab) {
      case 'Routing Time':
        return renderTab(RoutingTimeTab, {
          tasks: effectiveTasks,
          startDateStr: startStr,
          endDateStr: endStr,
          translate: t,
          localeCode: localeCode,
        });
      case 'Task Summary':
        return renderTab(TaskSummaryTab, {
          metrics: effectiveTaskSummaryMetrics,
          isLoading: isCalculatingMetrics,
          progress: historyProgress,
          startDateStr: startStr,
          endDateStr: endStr,
          isHasData: Object.entries(effectiveTaskSummaryMetrics).length > 0,
          translate: t,
          masterTruckData: masterTruckData.masterData,
        });
      case 'Pending Reasons':
        return renderTab(PendingReasonsTab, {
          data: effectivePendingReasonsData || reportPreview?.pendingReasonsData || [],
          pendingDetails: pendingDetails,
          reasons: reasons,
          onUpdatePendingDetail: handleUpdatePendingDetail,
          hasPendingGR: hasPendingGR,
          translate: t,
          driverData: driverData,
          showConvertedSuccess: showConvertedSuccess,
          allTasks: rawData.tasks,
          dateRange: dateRange,
          isAddPendingModalOpen: isAddPendingModalOpen,
          onCloseAddSuccessModal: () => setIsAddPendingModalOpen(false),
          onOpenAddSuccessModal: () => setIsAddPendingModalOpen(true),
        });
      case 'Time Driver':
        return renderTab(TimeDriverTab, {
          data: reportPreview.timeDriverData,
          translate: t,
          localeCode: localeCode,
          activeHubLocation: activeHubLocation,
        });
      case 'Truck Detail':
        return renderTab(TruckDetailTab, {
          data: effectiveTruckDetailData || reportPreview.truckDetailData,
          translate: t,
          localeCode: localeCode,
          isIndonesian: isIndonesian,
          colorLegend: colorLegend.truckDetail,
        });
      case 'Truck Usage':
        return renderTab(TruckUsageTab, {
          data: reportPreview.truckUsageData,
          translate: t,
          hubId: selectedLocation,
          onRefresh: fetchData,
          driverData: driverData,
          localeCode: localeCode,
        });
      case 'Distance Summary':
        return renderTab(DistanceSummaryTab, {
          data: reportPreview.distanceSummaryData,
          monthTotals: reportPreview.monthTotals,
          translate: t,
          localeCode: localeCode,
        });
      case 'Service Level':
        return renderTab(ServiceLevelTab, {
          tasks: effectiveTasks,
          driverData,
          startDateStr: startStr,
          endDateStr: endStr,
          translate: t,
          localeCode: localeCode,
        });
      case 'Diagram':
        return renderTab(DiagramTab, {
          tasks: effectiveTasks,
          driverData,
          startDateStr: startStr,
          endDateStr: endStr,
          dateRange,
        });
    }
  };

  const getMaxDate = () => {
    const today = new Date();
    const day = today.getDay();
    const max = new Date(today);

    if (day === 6) {
      max.setDate(max.getDate() + 2);
    } else {
      max.setDate(max.getDate() + 1);
    }
    return max;
  };

  const maxDateConfig = getMaxDate();

  const headerItems = [
    {
      label: t('summary.label'),
      component: (
        <CustomDatePicker
          selectsRange={true}
          startDate={tempDateRange ? tempDateRange[0] : null}
          endDate={tempDateRange ? tempDateRange[1] : null}
          onChange={handleTempDateChange}
          isLoading={isLoading}
          showApplyButton={true}
          onApply={handleApplyDate}
          applyText={t('common.apply')}
          maxDate={maxDateConfig}
          useCustomRangeFormat={true}
        />
      ),
      hideLabel: false,
    },
    {
      label: 'Action',
      component: (
        <Button
          disabled={isLoading || isDownload || isEmpty(rawData.tasks)}
          isLoading={isDownload}
          onClick={handleDownloadExcel}
          text={t('common.download')}
        />
      ),
      hideLabel: true,
    },
  ];

  const tabConfig = [
    {
      id: 'Diagram',
      label: 'Diagram',
    },
    { id: 'Routing Time', label: t('summary.tabs.routing_time.title') },
    { id: 'Task Summary', label: t('summary.tabs.task_summary.title') },
    { id: 'Pending Reasons', label: t('summary.tabs.pending_reasons.title') },
    { id: 'Time Driver', label: t('summary.tabs.time_driver.title') },
    { id: 'Truck Detail', label: t('summary.tabs.truck_detail.title') },
    { id: 'Truck Usage', label: t('summary.tabs.truck_usage.title') },
    { id: 'Distance Summary', label: t('summary.tabs.dist_summary.title') },
    {
      id: 'Service Level',
      label: t('summary.tabs.service_level.title'),
    },
  ];

  const longLoading = pendingEndpoints.length > 0 && (
    <div className="bg-orange-50 border border-orange-200 text-orange-700 px-4 py-3 rounded-md text-sm animate-pulse shadow-sm">
      <p>{t('summary.long_message')}</p>
      <p className="font-semibold text-center">{pendingEndpoints.join(', ')}</p>
    </div>
  );

  const getFooterConfig = () => {
    switch (activeTab) {
      case 'Routing Time':
        return {
          text: t('summary.tabs.routing_time.footer_text'),
        };
      case 'Truck Detail':
        return {
          text: t('common.click_for_detail_param', { parameter: t('summary.row') }),
          colorData: colorLegend.truckDetail,
        };
      case 'Time Driver':
        return {
          text: t('common.click_for_detail_param', { parameter: t('summary.row') }),
        };
      case 'Service Level':
        return {
          text: t('common.click_for_detail_param', { parameter: t('summary.row') }),
        };
      case 'Diagram':
        return {
          text: t('common.click_for_detail'),
        };
      case 'Pending Reasons':
        return {
          colorData: colorLegend.pendingReasons,
          text: t('common.click_for_detail_param', { parameter: t('summary.underline') }),
          data: (
            <div className="flex items-center gap-3 shrink-0">
              <div className="flex items-center gap-2 shrink-0">
                <p className="text-slate-500 dark:text-slate-400 text-xs whitespace-nowrap">
                  {t('common.all')} {t('common.task')}
                </p>
                <ToggleButton
                  size="sm"
                  value={showConvertedSuccess}
                  onChange={(val) => setShowConvertedSuccess(val)}
                  options={[
                    {
                      label: t('common.button.btn_no'),
                      value: false,
                    },
                    {
                      label: t('common.button.btn_yes'),
                      value: true,
                    },
                  ]}
                />
              </div>
              <Button
                text={t('common.button.btn_add')}
                size="sm"
                onClick={() => setIsAddPendingModalOpen(true)}
                disabled={isLoading || isEmpty(rawData.tasks)}
              />
            </div>
          ),
        };
      default:
        return {
          text: t('common.click_for_detail_param', { parameter: t('summary.underline') }),
        };
    }
  };

  return (
    <>
      <PageTemplate
        title={t('common.summary')}
        subtitle={
          <>
            {t('summary.subtitle_1')}{' '}
            <span className="font-semibold text-sky-600">{t('summary.subtitle_highlight')} </span>
          </>
        }
        headerItems={headerItems}
        activeTabId={activeTab}
        onTabClick={handleTabClick}
        tabs={tabConfig.map((tab) => ({
          id: tab.id,
          label: tab.label,
          extraContent: getPingDot(tab.id),
        }))}
        isLoading={isLoading}
        isEmpty={isTabEmpty()}
        emptyMessage={emptyMessage}
        footer={getFooterConfig()}
        bodyProps={{
          longLoadingContent: longLoading,
        }}
      >
        {!isLoading && renderContent()}
      </PageTemplate>
      <ConfirmModal
        isOpen={showWarningModal}
        title={t('common.modal.data_load_title')}
        message={t('common.modal.data_load_message', { days: 14 })}
        onConfirm={() => {
          setDateRange(pendingDateRange);
          setShowWarningModal(false);
        }}
        onCancel={() => {
          setShowWarningModal(false);
          setTempDateRange(dateRange);
        }}
      />
    </>
  );
}
