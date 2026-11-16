'use client';

import InformationButton from '@/components/button/InformationButton';
import ToggleButton from '@/components/button/ToggleButton';
import CustomDatePicker from '@/components/CustomDatePicker';
import Dropdown from '@/components/dropdown/Dropdown';
import StorageTypeFilter from '@/components/dropdown/StorageTypeFilter';
import HighlightText from '@/components/HighlightText';
import ConfirmModal from '@/components/modal/ConfirmModal';
import TaskModal from '@/components/modal/TaskModal';
import PageTemplate from '@/components/page/PageTemplate';
import SearchBar from '@/components/SearchBar';
import TableData from '@/components/table/TableData';
import Tooltip from '@/components/Tooltip';
import { getHubs, getPendingDetails, getTasks } from '@/lib/api/mileapp';
import { applyPendingDetailsToTasks, deliveryTooltip } from '@/lib/deliveryStatus';
import { useDeliveryStatus } from '@/lib/hooks/useDeliveryStatus';
import { useSuperadmin } from '@/lib/hooks/useSuperadmin';
import { toastError } from '@/lib/toast';
import {
  formatDateUniversal,
  getStorageType,
  parseCustomerString,
  toApiDateString,
} from '@/lib/utils';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

export default function TaskPage() {
  const { isSuperadmin } = useSuperadmin();
  const {
    t,
    formatStatus: formatDeliveryStatus,
    statusOptions: deliveryStatusOptions,
    getStatusLabel: getDeliveryStatusLabel,
    isStatus,
  } = useDeliveryStatus();

  const [dateRange, setDateRange] = useState([new Date(), new Date()]);
  const [startDate, endDate] = dateRange;
  const [searchQuery, setSearchQuery] = useState('');
  const [tempDateRange, setTempDateRange] = useState([new Date(), new Date()]);
  const [tempStart, tempEnd] = tempDateRange;
  const [showWarningModal, setShowWarningModal] = useState(false);
  const [statusDeliveryFilter, setStatusDeliveryFilter] = useState('ALL');
  const [statusTaskFilter, setStatusTaskFilter] = useState('ALL');
  const [storageFilter, setStorageFilter] = useState(['DRY', 'FROZEN']);
  const [isTaskModalOpen, setIsTaskModalOpen] = useState(false);
  const [selectedTaskId, setSelectedTaskId] = useState(null);
  const [isAllHub, setIsAllHub] = useState(false);
  const [hubsData, setHubsData] = useState([]);
  const [tasks, setTasks] = useState([]);
  const [loading, setLoading] = useState(true);
  const [sortConfig, setSortConfig] = useState({ key: 'startTime', direction: 'asc' });

  const cacheRef = useRef({ ONE: null, ALL: null, dateKey: '' });

  const statusTaskOptions = useMemo(
    () => [
      { label: t('common.all'), value: 'ALL' },
      { label: t('common.status.done'), value: 'DONE' },
      { label: t('common.status.unassigned'), value: 'UNASSIGNED' },
      { label: t('common.status.ongoing'), value: 'ONGOING' },
      { label: t('common.status.manual_assign'), value: 'MANUAL_ASSIGN' },
      { label: t('common.status.diff_day'), value: 'DIFF_DAY' },
    ],
    [t]
  );

  useEffect(() => {
    const fetchHubs = async () => {
      try {
        const res = await getHubs();
        setHubsData(res?.allHub || []);
      } catch (err) {
        toastError(t('common.toast.error', { err: err.message }), err);
      }
    };
    fetchHubs();
  }, [t]);

  const fetchTasksData = useCallback(async () => {
    if (!startDate || !endDate) return;

    const localStart = new Date(startDate);
    localStart.setHours(0, 0, 0, 0);
    const localEnd = new Date(endDate);
    localEnd.setHours(23, 59, 59, 999);

    const currentKey = `${localStart.getTime()}-${localEnd.getTime()}`;
    const mode = isAllHub ? 'ALL' : 'ONE';

    if (cacheRef.current.dateKey !== currentKey) {
      cacheRef.current = { ONE: null, ALL: null, dateKey: currentKey };
    } else if (cacheRef.current[mode]) {
      setTasks(cacheRef.current[mode]);
      return;
    }

    setLoading(true);
    try {
      const isShowAll = isAllHub;
      const startDateStr = formatDateUniversal(localStart, 'YYYY-MM-DD');
      const endDateStr = formatDateUniversal(localEnd, 'YYYY-MM-DD');

      const [res, pendingDetailsData] = await Promise.all([
        getTasks({
          isShowAll,
          status: 'DONE,UNASSIGNED,ONGOING',
          dateFrom: startDateStr,
          dateTo: endDateStr,
          timeFrom: toApiDateString(localStart),
          timeTo: toApiDateString(localEnd),
        }),
        getPendingDetails(startDateStr, endDateStr).catch(() => []),
      ]);

      const dataArray = Array.isArray(res) ? res : res?.data || [];
      const pdList = Array.isArray(pendingDetailsData)
        ? pendingDetailsData
        : Array.isArray(pendingDetailsData?.data)
          ? pendingDetailsData.data
          : [];
      const effectiveTasks = applyPendingDetailsToTasks(dataArray, pdList);

      cacheRef.current[mode] = effectiveTasks;
      setTasks(effectiveTasks);
    } catch (err) {
      toastError(t('common.toast.error', { err: err.message }), err);
    } finally {
      setLoading(false);
    }
  }, [startDate, endDate, isAllHub, t]);

  useEffect(() => {
    fetchTasksData();
  }, [fetchTasksData]);

  useEffect(() => {
    const handlePendingUpdated = () => {
      cacheRef.current = { ONE: null, ALL: null, dateKey: '' };
      fetchTasksData();
    };
    const handleStorage = (e) => {
      if (e.key === 'tms_pending_details_last_update') {
        cacheRef.current = { ONE: null, ALL: null, dateKey: '' };
        fetchTasksData();
      }
    };
    window.addEventListener('tms-pending-details-updated', handlePendingUpdated);
    window.addEventListener('storage', handleStorage);
    return () => {
      window.removeEventListener('tms-pending-details-updated', handlePendingUpdated);
      window.removeEventListener('storage', handleStorage);
    };
  }, [fetchTasksData]);

  const hubMap = useMemo(() => new Map(hubsData?.map((hub) => [hub._id, hub.name])), [hubsData]);

  const processedData = useMemo(() => {
    let result = tasks.map((task) => {
      const custInfo = parseCustomerString(task.customerOrder);
      const hubName = hubMap.get(task.hubId) || '-';

      return {
        ...task,
        _custInfo: custInfo,
        _hubName: hubName,
        _custName: custInfo.name || '-',
        _custId: custInfo.id || '-',
        _custLocation: custInfo.location || '-',
        _invoiceNumber: custInfo.invoiceNumber || '-',
        _truncateInvoice: custInfo.truncateInvoice || '-',
        _isTruncated: custInfo.isTruncated,
        _statusDel: task.statusDelivery ? formatDeliveryStatus(task.statusDelivery) : '-',
        _isStatusModified: Boolean(
          task.isModified ||
            (task.pendingDetail?.new_status &&
              task.originalStatus &&
              !isStatus(task.originalStatus, task.pendingDetail.new_status))
        ),
      };
    });

    if (statusDeliveryFilter !== 'ALL') {
      result = result.filter((task) => isStatus(task.statusDelivery, statusDeliveryFilter));
    }

    if (statusTaskFilter !== 'ALL') {
      result = result.filter((task) => {
        if (statusTaskFilter === 'DONE') return !!task.statusDelivery;
        if (statusTaskFilter === 'UNASSIGNED') return !task.assignee;
        if (statusTaskFilter === 'ONGOING') return task.status?.toLowerCase() === 'ongoing';
        if (statusTaskFilter === 'MANUAL_ASSIGN') {
          return (
            task.assignee &&
            (!task.routingResultId || !task.routePlannedOrder || !task.eta || !task.etd)
          );
        }
        if (statusTaskFilter === 'DIFF_DAY') {
          const startFormat = task.startTime ? formatDateUniversal(task.startTime) : null;
          const doneFormat = task.doneTime ? formatDateUniversal(task.doneTime) : null;
          return startFormat && doneFormat && startFormat !== doneFormat;
        }
        return true;
      });
    }

    if (storageFilter.length === 0) {
      result = [];
    } else if (storageFilter.length === 1) {
      result = result.filter((task) => {
        const storage = (
          getStorageType(task.typeStorage) !== '-'
            ? getStorageType(task.typeStorage)
            : getStorageType(task.driverName || '')
        ).toUpperCase();
        return storageFilter.includes(storage);
      });
    }

    const q = searchQuery.toLowerCase().trim();
    if (q) {
      result = result.filter((task) => {
        const strToSearch =
          `${task._hubName} ${task._custName} ${task._custId} ${task._invoiceNumber} ${task.driverName || ''}`.toLowerCase();
        return strToSearch.includes(q);
      });
    }

    return result;
  }, [
    tasks,
    hubMap,
    searchQuery,
    statusDeliveryFilter,
    statusTaskFilter,
    storageFilter,
    formatDeliveryStatus,
    isStatus,
  ]);

  const handleApplyDate = () => {
    if (!tempStart || !tempEnd) return;
    const diffDays = Math.ceil(Math.abs(tempEnd - tempStart) / (1000 * 60 * 60 * 24));

    if (diffDays > 14) {
      setShowWarningModal(true);
    } else {
      setDateRange([tempStart, tempEnd]);
    }
  };

  const columns = useMemo(() => {
    const cw = isAllHub
      ? {
          hub: 'w-[10%]',
          flow: 'w-[8%]',
          name: 'w-[12%]',
          id: 'w-[8%]',
          loc: 'w-[8%]',
          inv: 'w-[12%]',
          stat: 'w-[8%]',
          start: 'w-[8%]',
          assign: 'w-[8%]',
          done: 'w-[8%]',
          assignee: 'w-[10%]',
        }
      : {
          flow: 'w-[8%]',
          name: 'w-[15%]',
          id: 'w-[8%]',
          loc: 'w-[10%]',
          inv: 'w-[15%]',
          stat: 'w-[10%]',
          start: 'w-[8%]',
          assign: 'w-[8%]',
          done: 'w-[8%]',
          assignee: 'w-[10%]',
        };

    return [
      ...(isAllHub
        ? [
            {
              key: '_hubName',
              label: t('common.branch'),
              width: cw.hub,
              sortable: true,
              render: (row) => <HighlightText text={row._hubName} highlight={searchQuery} />,
            },
          ]
        : []),
      {
        key: 'flow',
        label: t('common.flow'),
        width: cw.flow,
        sortable: true,
        render: (row) => row.flow || '-',
      },
      {
        key: '_custName',
        label: t('common.customer_name'),
        width: cw.name,
        sortable: true,
        render: (row) => <HighlightText text={row._custName} highlight={searchQuery} />,
      },
      {
        key: '_custId',
        label: t('common.customer_id'),
        width: cw.id,
        sortable: true,
        render: (row) => <HighlightText text={row._custId} highlight={searchQuery} />,
      },
      {
        key: '_custLocation',
        label: t('common.location_id'),
        width: cw.loc,
        sortable: true,
      },
      {
        key: '_invoiceNumber',
        label: t('common.invoice_number'),
        width: cw.inv,
        sortable: true,
        render: (row) =>
          row._isTruncated ? (
            <Tooltip tooltipContent={row._invoiceNumber}>
              <span className="cursor-help block">
                <HighlightText text={row._truncateInvoice || '-'} highlight={searchQuery} />
              </span>
            </Tooltip>
          ) : (
            <HighlightText text={row._invoiceNumber || '-'} highlight={searchQuery} />
          ),
      },
      {
        key: '_statusDel',
        label: t('common.status.delivery_status'),
        width: cw.stat,
        sortable: true,
        render: (row) => {
          const isModified = Boolean(row._isStatusModified || row.isModified);
          if (isModified) {
            const tooltipContent = deliveryTooltip({
              originalStatus: row.originalStatus || row.statusDelivery,
              pendingDetail: row.pendingDetail,
              formatStatus: formatDeliveryStatus,
              t,
            });

            return (
              <div className="inline-flex items-center gap-1.5 flex-wrap">
                <span className="font-bold text-red-600 dark:text-red-400">{row._statusDel}</span>
                <span onClick={(e) => e.stopPropagation()}>
                  <InformationButton infoText={tooltipContent} size="4" />
                </span>
              </div>
            );
          }

          return row._statusDel;
        },
      },
      {
        key: 'assignedTime',
        label: t('common.assigned_time'),
        width: cw.assign,
        sortable: true,
        render: (row) =>
          row.startTime ? formatDateUniversal(row.assignedTime, 'DD/MM/YYYY HH:mm') : '-',
      },
      {
        key: 'startTime',
        label: t('common.start_time'),
        width: cw.start,
        sortable: true,
        render: (row) =>
          row.startTime ? formatDateUniversal(row.startTime, 'DD/MM/YYYY HH:mm') : '-',
      },
      {
        key: 'doneTime',
        label: t('common.done_time'),
        width: cw.done,
        sortable: true,
        render: (row) =>
          row.startTime ? formatDateUniversal(row.doneTime, 'DD/MM/YYYY HH:mm') : '-',
      },
      {
        key: 'driverName',
        label: t('common.driver'),
        width: cw.assignee,
        sortable: true,
        render: (row) => <HighlightText text={row.driverName} highlight={searchQuery} />,
      },
    ];
  }, [isAllHub, searchQuery, t, formatDeliveryStatus]);

  const searchPlaceholder = `${t('common.customer_name')}, ${t('common.customer_id')}, ${t('common.invoice_number')} ${t('common.driver')} ${isAllHub ? `, ${t('common.branch')}` : ''}`;

  const tableCustomSort = useCallback((items, config) => {
    if (!config) return items;
    return [...items].sort((a, b) => {
      const valA = a[config.key] ?? '';
      const valB = b[config.key] ?? '';
      let cmp =
        typeof valA === 'string' && typeof valB === 'string'
          ? valA.localeCompare(valB)
          : valA < valB
            ? -1
            : valA > valB
              ? 1
              : 0;

      if (config.direction === 'desc') cmp = -cmp;

      if (cmp === 0 && config.key === '_hubName') {
        const timeA = a.startTime ?? '';
        const timeB = b.startTime ?? '';
        return timeA.localeCompare(timeB);
      }
      return cmp;
    });
  }, []);

  const headerItems = [
    {
      label: t('common.search'),
      component: (
        <SearchBar
          disabled={loading}
          onChange={setSearchQuery}
          placeholder={t('common.search')}
          tooltip={searchPlaceholder}
          value={searchQuery}
          width="w-full"
        />
      ),
    },
    {
      label: t('common.range_delivery'),
      component: (
        <CustomDatePicker
          disabled={loading}
          endDate={tempEnd}
          onChange={(update) => setTempDateRange(update)}
          selectsRange={true}
          startDate={tempStart}
          useCustomRangeFormat={true}
          showApplyButton={true}
          onApply={handleApplyDate}
          className="w-full"
        />
      ),
    },
    {
      label: t('common.status.delivery_status'),
      component: (
        <Dropdown
          className="w-full"
          disabled={loading}
          onChange={setStatusDeliveryFilter}
          options={deliveryStatusOptions}
          value={statusDeliveryFilter}
          getLabel={getDeliveryStatusLabel}
        />
      ),
    },
    {
      label: t('common.status.task_status'),
      component: (
        <Dropdown
          className="w-full"
          disabled={loading}
          onChange={setStatusTaskFilter}
          options={statusTaskOptions}
          value={statusTaskFilter}
          getLabel={(val) => statusTaskOptions.find((o) => o.value === val)?.label || val}
        />
      ),
    },
    {
      label: t('common.storage_type'),
      component: (
        <StorageTypeFilter
          className="w-full"
          disabled={loading}
          onApply={setStorageFilter}
          selectedTypes={storageFilter}
        />
      ),
    },
    ...(isSuperadmin
      ? [
          {
            label: t('common.branch'),
            component: (
              <ToggleButton
                className="w-full"
                disabled={loading}
                onChange={(val) => {
                  const isAll = val === 'ALL';
                  setIsAllHub(isAll);
                  setSortConfig({ key: isAll ? '_hubName' : 'startTime', direction: 'asc' });
                }}
                options={[
                  { label: t('task_detail.modal.one'), value: 'ONE' },
                  { label: t('common.all'), value: 'ALL' },
                ]}
                value={isAllHub ? 'ALL' : 'ONE'}
              />
            ),
          },
        ]
      : []),
  ];

  return (
    <>
      <PageTemplate
        title={t('common.task_details')}
        subtitle={
          <>
            <span className="font-semibold text-sky-600">
              {t('task_detail.subtitle_highlight')}{' '}
            </span>
            {t('task_detail.subtitle')}
          </>
        }
        headerItems={headerItems}
        isLoading={false}
        isEmpty={false}
        footer={{
          text: t('common.click_for_detail'),
        }}
      >
        <TableData
          columns={columns}
          data={processedData}
          customSort={tableCustomSort}
          externalSortConfig={sortConfig}
          onExternalSort={setSortConfig}
          isLoading={loading}
          paginate={true}
          onRowClick={(row) => {
            setSelectedTaskId(row._id);
            setIsTaskModalOpen(true);
          }}
        />
      </PageTemplate>
      <ConfirmModal
        isOpen={showWarningModal}
        title={t('common.modal.data_load_title')}
        message={t('common.modal.data_load_message', { days: 14 })}
        onConfirm={() => {
          setShowWarningModal(false);
          setDateRange([tempStart, tempEnd || tempStart]);
        }}
        onCancel={() => setShowWarningModal(false)}
      />

      <TaskModal
        isOpen={isTaskModalOpen}
        onClose={() => setIsTaskModalOpen(false)}
        taskId={selectedTaskId}
        isAllHub={isAllHub}
      />
    </>
  );
}
