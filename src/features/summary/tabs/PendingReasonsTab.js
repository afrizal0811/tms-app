import ConfirmModal from '@/components/modal/ConfirmModal';
import TaskModal from '@/components/modal/TaskModal';
import Tooltip from '@/components/Tooltip';
import { deletePendingDetail, postPendingDetail } from '@/lib/api/mileapp';
import { DELIVERY_STATUS_KEYS, deliveryTooltip } from '@/lib/deliveryStatus';
import { useDeliveryStatus } from '@/lib/hooks/useDeliveryStatus';
import { getLocalStorage } from '@/lib/localStorageHandler';
import { toastError, toastSuccess } from '@/lib/toast';
import { getBasePlate, isEmpty, parseCustomerString } from '@/lib/utils';
import { Fragment, useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import AddPendingModal from './modals/AddPendingModal';
import PendingReasonModal from './modals/PendingReasonModal';

const TH_CLASS =
  'border border-gray-300 dark:border-slate-700 px-2 py-2 bg-sky-600 dark:bg-sky-900 text-white text-xs font-bold whitespace-nowrap';
const BASE_TD_CLASS =
  'px-2 py-2 text-xs text-slate-700 dark:text-slate-200 whitespace-nowrap text-center border-r border-r-gray-200 dark:border-r-slate-700 border-l border-l-gray-200 dark:border-l-slate-700';
const RED_CELL_STYLE = 'text-red-700 dark:text-red-400 bg-red-100 dark:bg-[#4a1c1c] font-bold';
const YELLOW_CELL_STYLE = 'bg-red-500 text-white dark:bg-red-600 font-bold';

function emitPendingUpdate(detail) {
  if (typeof window === 'undefined') return;
  window.dispatchEvent(new CustomEvent('tms-pending-details-updated', { detail }));
  try {
    localStorage.setItem('tms_pending_details_last_update', Date.now().toString());
  } catch {}
}

function getUserName() {
  const { storedUser } = getLocalStorage();
  return storedUser?.name || 'User';
}

function hasValidText(val) {
  return Boolean(val && !isEmpty(val) && val !== '-');
}

const ReasonCell = ({ text, className }) => {
  const [isTruncated, setIsTruncated] = useState(false);
  const textRef = useRef(null);

  const checkOverflow = () => {
    if (textRef.current) {
      setIsTruncated(textRef.current.scrollWidth > textRef.current.clientWidth);
    }
  };

  return (
    <Tooltip tooltipContent={isTruncated ? text : ''}>
      <td
        className={`${className} text-left max-w-[200px] p-0 ${isTruncated ? 'cursor-help' : 'cursor-default'}`}
      >
        <div ref={textRef} className="truncate w-full" onMouseEnter={checkOverflow}>
          {text}
        </div>
      </td>
    </Tooltip>
  );
};

const ActionCell = ({ text, className, onClick, isAllEmpty = false, disabled = false }) => {
  const [isTruncated, setIsTruncated] = useState(false);
  const [isHovered, setIsHovered] = useState(false);
  const textRef = useRef(null);

  const checkOverflow = () => {
    if (textRef.current) {
      setIsTruncated(textRef.current.scrollWidth > textRef.current.clientWidth);
    }
  };

  const hasData = hasValidText(text);

  return (
    <td
      className={`${className} max-w-[150px] ${disabled ? '!cursor-default hover:bg-transparent' : ''}`}
      onClick={disabled ? undefined : onClick}
      onMouseEnter={() => !disabled && setIsHovered(true)}
      onMouseLeave={() => !disabled && setIsHovered(false)}
    >
      {hasData ? (
        <Tooltip tooltipContent={isTruncated ? text : ''}>
          <div
            ref={textRef}
            className={`truncate w-full ${disabled ? 'cursor-default' : 'cursor-pointer'}`}
            onMouseEnter={checkOverflow}
          >
            {text}
          </div>
        </Tooltip>
      ) : (
        <div className="w-full h-full flex items-center justify-center transition-colors">
          {!disabled && isAllEmpty && isHovered ? (
            <span className="relative text-slate-600/30 dark:text-slate-400/30 font-bold">+</span>
          ) : (
            <span className="text-slate-400 dark:text-slate-500">-</span>
          )}
        </div>
      )}
    </td>
  );
};

const SOCell = ({
  text,
  className,
  isError,
  onClick,
  isConvertedFromSuccess = false,
  isAmberRow = false,
}) => {
  if (!text) return <td className={className}></td>;

  const textStyle = `inline-block border-b-2 border-dotted pb-0.5 border-red-700 dark:border-red-300 ${isError ? 'text-[#FF0000] dark:text-red-400 font-bold' : ''}`;

  const hoverClass =
    isAmberRow || isConvertedFromSuccess
      ? 'hover:bg-amber-200/80 dark:hover:bg-amber-900/60'
      : 'hover:bg-blue-50 dark:hover:bg-slate-700';

  return (
    <td
      className={`${className} ${hoverClass} transition-colors relative group cursor-pointer`}
      onClick={onClick}
    >
      <span className={textStyle}>{text}</span>
    </td>
  );
};

const EmptyAlertCell = ({ value, className, alertStyle }) => {
  const isValEmpty = isEmpty(value);
  return (
    <td className={isValEmpty ? `${className} ${alertStyle}` : className}>
      {isValEmpty ? '-' : value}
    </td>
  );
};

const StatusCell = ({
  item,
  pd,
  originalStatus,
  currentStatus,
  isModified,
  isWrongGR,
  hasPendingGR,
  formatStatus,
  normalizeStatus,
  isPendingGr,
  translate,
  onUpdatePendingDetail,
  className,
  isConvertedFromSuccess = false,
  isAmberRow = false,
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [isUpdating, setIsUpdating] = useState(false);
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [pendingTargetKey, setPendingTargetKey] = useState(null);
  const [menuPos, setMenuPos] = useState({ top: 0, left: 0, placeTop: false });
  const cellRef = useRef(null);
  const menuRef = useRef(null);

  const statusLabel = currentStatus ? formatStatus(currentStatus) : '-';
  const tooltipContent = isModified
    ? deliveryTooltip({
        originalStatus,
        pendingDetail: pd,
        formatStatus,
        t: translate,
      })
    : '';

  const updateMenuPosition = () => {
    if (!cellRef.current) return;
    const rect = cellRef.current.getBoundingClientRect();
    const spaceBelow = window.innerHeight - rect.bottom;
    const dropdownHeight = 180;
    const placeTop = spaceBelow < dropdownHeight && rect.top > dropdownHeight;

    setMenuPos({
      top: placeTop ? rect.top - 4 : rect.bottom + 4,
      left: Math.max(10, Math.min(rect.left, window.innerWidth - 180)),
      width: Math.max(140, rect.width),
      placeTop,
    });
  };

  const handleCellClick = (e) => {
    e.stopPropagation();
    if (isUpdating) return;
    if (!isOpen) {
      updateMenuPosition();
    }
    setIsOpen((prev) => !prev);
  };

  useEffect(() => {
    if (!isOpen) return;

    updateMenuPosition();

    const handleClickOutside = (e) => {
      if (cellRef.current?.contains(e.target) || menuRef.current?.contains(e.target)) return;
      setIsOpen(false);
    };

    const handleScroll = (e) => {
      if (menuRef.current?.contains(e.target)) return;
      setIsOpen(false);
    };

    document.addEventListener('mousedown', handleClickOutside);
    window.addEventListener('scroll', handleScroll, true);
    window.addEventListener('resize', handleScroll, true);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      window.removeEventListener('scroll', handleScroll, true);
      window.removeEventListener('resize', handleScroll, true);
    };
  }, [isOpen]);

  const statusList = [
    { key: DELIVERY_STATUS_KEYS.SUCCESS, label: formatStatus(DELIVERY_STATUS_KEYS.SUCCESS) },
    { key: DELIVERY_STATUS_KEYS.PARTIAL, label: formatStatus(DELIVERY_STATUS_KEYS.PARTIAL) },
    { key: DELIVERY_STATUS_KEYS.PENDING, label: formatStatus(DELIVERY_STATUS_KEYS.PENDING) },
    { key: DELIVERY_STATUS_KEYS.CANCEL, label: formatStatus(DELIVERY_STATUS_KEYS.CANCEL) },
  ];
  if (hasPendingGR || isPendingGr(currentStatus) || isPendingGr(originalStatus)) {
    statusList.push({
      key: DELIVERY_STATUS_KEYS.PENDING_GR,
      label: formatStatus(DELIVERY_STATUS_KEYS.PENDING_GR),
    });
  }

  const executeStatusChange = async (targetKey) => {
    setIsUpdating(true);
    try {
      const isOriginal = normalizeStatus(targetKey) === normalizeStatus(originalStatus);
      const newStatusVal = isOriginal ? null : targetKey;
      const isNowSuccess = targetKey === DELIVERY_STATUS_KEYS.SUCCESS;

      const [d, m, y] = (item.dateStr || '').split('-');
      const dbDate = y && y.length === 4 ? `${y}-${m}-${d}` : formatDateUniversal(item.dateStr);
      const userName = getUserName();

      const hasReasons = Boolean(
        pd.groupReason || pd.detailReason || pd.internalExternal || pd.pic
      );

      const isOriginalSuccess = normalizeStatus(originalStatus) === DELIVERY_STATUS_KEYS.SUCCESS;

      let resultData;
      if (newStatusVal === null && (!hasReasons || isOriginalSuccess)) {
        await deletePendingDetail(item._id);
        resultData = { taskId: item._id, deleted: true };
      } else {
        const payload = {
          taskId: item._id,
          date: dbDate,
          new_status: newStatusVal,
          internalExternal: isNowSuccess ? null : (pd.internalExternal ?? null),
          detailReason: isNowSuccess ? null : (pd.detailReason ?? null),
          groupReason: isNowSuccess ? null : (pd.groupReason ?? null),
          pic: isNowSuccess ? null : (pd.pic ?? null),
          userName,
          created_by: pd.created_by || userName,
          updated_by: userName,
        };

        const res = await postPendingDetail(payload);
        resultData = res.data || res;
      }

      onUpdatePendingDetail(resultData);
      emitPendingUpdate(resultData);
      toastSuccess(translate('common.toast.success'));
    } catch (err) {
      toastError(translate('common.toast.error', { err: err.message }), err);
    } finally {
      setIsOpen(false);
      setIsConfirmOpen(false);
      setPendingTargetKey(null);
      setIsUpdating(false);
    }
  };

  const handleSelectStatus = (targetKey) => {
    if (normalizeStatus(currentStatus) === normalizeStatus(targetKey)) {
      setIsOpen(false);
      return;
    }

    const isOriginal = normalizeStatus(targetKey) === normalizeStatus(originalStatus);
    if (isOriginal && isModified) {
      setPendingTargetKey(targetKey);
      setIsOpen(false);
      setIsConfirmOpen(true);
      return;
    }

    executeStatusChange(targetKey);
  };

  const handleCancelConfirm = () => {
    if (isUpdating) return;
    setIsConfirmOpen(false);
    setPendingTargetKey(null);
  };

  const handleConfirmRevert = () => {
    if (isUpdating) return;
    executeStatusChange(pendingTargetKey || originalStatus);
  };

  const textStyle = `inline-block border-b-2 border-dotted pb-0.5 border-red-700 dark:border-red-300 ${
    isModified
      ? 'font-bold text-red-600 dark:text-red-400'
      : isWrongGR
        ? 'font-bold text-[#FF0000] dark:text-red-400'
        : ''
  }`;
  const dropdownStyle = {
    position: 'fixed',
    top: `${menuPos.top}px`,
    left: `${menuPos.left}px`,
    transform: menuPos.placeTop ? 'translateY(-100%)' : 'none',
    minWidth: `${menuPos.width || 140}px`,
    zIndex: 9999,
  };

  const hoverClass =
    isAmberRow || isConvertedFromSuccess
      ? 'hover:bg-amber-200/80 dark:hover:bg-amber-900/60'
      : 'hover:bg-blue-50 dark:hover:bg-slate-700';

  return (
    <>
      <td
        ref={cellRef}
        className={`${className} ${hoverClass} transition-colors cursor-pointer group select-none relative`}
        onClick={handleCellClick}
      >
        <Tooltip tooltipContent={!isOpen && isModified ? tooltipContent : ''}>
          <div className="flex items-center justify-between gap-1 w-full">
            <span className={textStyle}>{statusLabel}</span>
            {isUpdating ? (
              <div className="w-3 h-3 border-2 border-sky-500 border-t-transparent rounded-full animate-spin shrink-0" />
            ) : (
              <svg
                xmlns="http://www.w3.org/2000/svg"
                className={`h-3 w-3 shrink-0 text-slate-400 group-hover:text-sky-600 dark:group-hover:text-sky-400 transition-transform ${
                  isOpen ? 'rotate-180' : ''
                }`}
                viewBox="0 0 20 20"
                fill="currentColor"
              >
                <path
                  fillRule="evenodd"
                  d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z"
                  clipRule="evenodd"
                />
              </svg>
            )}
          </div>
        </Tooltip>
      </td>

      {isOpen &&
        typeof document !== 'undefined' &&
        createPortal(
          <div
            ref={menuRef}
            style={dropdownStyle}
            className="bg-white dark:bg-slate-800 border border-gray-200 dark:border-slate-700 rounded-lg shadow-xl py-1 text-xs z-[9999] animate-in fade-in zoom-in-95 duration-100"
            onClick={(e) => e.stopPropagation()}
          >
            {statusList.map((opt) => {
              const isSelected = normalizeStatus(currentStatus) === normalizeStatus(opt.key);

              return (
                <button
                  key={opt.key}
                  type="button"
                  disabled={isUpdating}
                  onClick={() => handleSelectStatus(opt.key)}
                  className={`w-full text-left px-3 py-2 flex items-center justify-between gap-2 hover:bg-sky-50 dark:hover:bg-slate-700/80 transition-colors ${
                    isSelected
                      ? 'font-bold text-sky-600 dark:text-sky-400 bg-sky-50/50 dark:bg-slate-700/40'
                      : 'cursor-pointer text-slate-700 dark:text-slate-200'
                  }`}
                >
                  <span>{opt.label}</span>
                </button>
              );
            })}
          </div>,
          document.body
        )}

      {isConfirmOpen &&
        typeof document !== 'undefined' &&
        createPortal(
          <div onClick={(e) => e.stopPropagation()}>
            <ConfirmModal
              isOpen={isConfirmOpen}
              onCancel={handleCancelConfirm}
              onConfirm={handleConfirmRevert}
              loading={isUpdating}
              title={
                translate('summary.tabs.pending_reasons.confirm_revert_title') ||
                'Konfirmasi Pengembalian Status'
              }
              message={translate('summary.tabs.pending_reasons.confirm_revert_message', {
                status: formatStatus(originalStatus),
              })}
            />
          </div>,
          document.body
        )}
    </>
  );
};

export default function PendingReasonsTab({
  data,
  pendingDetails,
  reasons,
  hasPendingGR,
  translate,
  onUpdatePendingDetail,
  showConvertedSuccess = false,
  allTasks = [],
  dateRange = [null, null],
  driverData = [],
  isAddPendingModalOpen = false,
  onCloseAddSuccessModal = () => {},
  onOpenAddSuccessModal = () => {},
}) {
  const { formatStatus, isPendingGr, normalizeStatus } = useDeliveryStatus();
  const [modalData, setModalData] = useState(null);
  const [isTaskModalOpen, setIsTaskModalOpen] = useState(false);
  const [selectedTaskId, setSelectedTaskId] = useState(null);
  const [selectedSuccessTasksMap, setSelectedSuccessTasksMap] = useState(() => new Map());
  const [successModalDate, setSuccessModalDate] = useState(() =>
    dateRange && dateRange[0] ? new Date(dateRange[0]) : new Date()
  );
  const [prevDateRange, setPrevDateRange] = useState(dateRange);

  if (dateRange !== prevDateRange) {
    setPrevDateRange(dateRange);
    setSelectedSuccessTasksMap(new Map());
    if (dateRange && dateRange[0]) {
      setSuccessModalDate(new Date(dateRange[0]));
    }
  }

  const pendingDetailsMap = useMemo(() => {
    const map = new Map();
    if (pendingDetails) {
      for (const d of pendingDetails) {
        if (d?.taskId) map.set(String(d.taskId), d);
      }
    }
    return map;
  }, [pendingDetails]);

  const allTasksMap = useMemo(() => {
    const map = new Map();
    if (allTasks) {
      for (const t of allTasks) {
        const id = t?._id || t?.id || t?.taskId;
        if (id) map.set(String(id), t);
      }
    }
    return map;
  }, [allTasks]);

  const handleRowClick = (taskId) => {
    if (!taskId || taskId === '-') return;
    setSelectedTaskId(taskId);
    setIsTaskModalOpen(true);
  };

  const tableHeaders = useMemo(
    () => [
      translate('common.flow'),
      translate('common.delivery_date'),
      translate('common.license_number'),
      translate('common.driver'),
      translate('common.customer_name'),
      translate('common.invoice_number'),
      translate('common.status.delivery_status'),
      translate('common.reason'),
      translate('summary.tabs.pending_reasons.group_reason'),
      translate('summary.tabs.pending_reasons.detail'),
      translate('common.category'),
      'PIC',
      translate('common.open_time'),
      translate('common.close_time'),
      translate('common.eta'),
      translate('common.etd'),
      translate('common.actual_arrival'),
      translate('common.actual_departure'),
      translate('common.plan_visit'),
      translate('common.actual_visit'),
      translate('common.plan_seq'),
      translate('common.actual_seq'),
      translate('common.storage_type'),
    ],
    [translate]
  );

  const filteredData = useMemo(() => {
    return (data || []).filter((item) => {
      const taskId = String(item._id || item.id || item.taskId);
      const pd = pendingDetailsMap.get(taskId) || item.pendingDetail || {};
      const effectiveDelivery = pd.new_status || item.statusDelivery || item.status;
      const rawTask = allTasksMap.get(taskId);
      const rawOrig =
        item.originalStatus ||
        (rawTask
          ? (Array.isArray(rawTask.statusDelivery)
              ? rawTask.statusDelivery[0]
              : rawTask.statusDelivery
            )
              ?.split(',')[0]
              .trim() || rawTask.status
          : item.status);
      const isConvertedSuccess =
        normalizeStatus(effectiveDelivery) === DELIVERY_STATUS_KEYS.SUCCESS &&
        normalizeStatus(rawOrig) !== DELIVERY_STATUS_KEYS.SUCCESS;

      return !(!showConvertedSuccess && isConvertedSuccess);
    });
  }, [data, pendingDetailsMap, allTasksMap, normalizeStatus, showConvertedSuccess]);

  const hasConversionModalData = Array.isArray(modalData)
    ? modalData.some((t) => t.isSuccessConversion)
    : Boolean(modalData?.isSuccessConversion);

  return (
    <Fragment>
      <div className="rounded-b-xl overflow-auto m-0 relative">
        <AddPendingModal
          isOpen={isAddPendingModalOpen}
          onClose={() => {
            setSelectedSuccessTasksMap(new Map());
            onCloseAddSuccessModal();
          }}
          allTasks={allTasks}
          dateRange={dateRange}
          driverData={driverData}
          pendingDetails={pendingDetails}
          translate={translate}
          selectedTasksMap={selectedSuccessTasksMap}
          setSelectedTasksMap={setSelectedSuccessTasksMap}
          selectedDate={successModalDate}
          setSelectedDate={setSuccessModalDate}
          onProceed={(selectedTasks) => {
            onCloseAddSuccessModal();
            setModalData(selectedTasks);
          }}
        />

        <PendingReasonModal
          isOpen={Boolean(modalData)}
          onClose={() => {
            setModalData(null);
            setSelectedSuccessTasksMap(new Map());
          }}
          data={modalData}
          reasons={reasons || []}
          onSuccess={(res) => {
            setSelectedSuccessTasksMap(new Map());
            onUpdatePendingDetail(res);
          }}
          hasPendingGR={hasPendingGR}
          onBack={
            hasConversionModalData
              ? () => {
                  setModalData(null);
                  onOpenAddSuccessModal();
                }
              : undefined
          }
        />

        <TaskModal
          isOpen={isTaskModalOpen}
          onClose={() => setIsTaskModalOpen(false)}
          taskId={selectedTaskId}
        />

        <table className="border-collapse w-full text-sm">
          <thead className="sticky top-0 z-20">
            <tr>
              {tableHeaders.map((header, idx) => (
                <th key={idx} className={TH_CLASS}>
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="bg-white dark:bg-slate-800">
            {filteredData.map((item, idx) => {
              const {
                name: customerName,
                invoiceNumber,
                truncateInvoice,
                isTruncated,
              } = parseCustomerString(item.customerOrder || item.customerName);

              const isLastInDate = filteredData[idx + 1]?.dateStr !== item.dateStr;
              const borderBottomClass = isLastInDate
                ? 'border-b-[4px] border-b-slate-400 dark:border-b-slate-600'
                : 'border-b border-b-gray-200 dark:border-b-slate-700';
              const tdClass = `${BASE_TD_CLASS} ${borderBottomClass}`;

              const taskId = String(item._id || item.id || item.taskId);
              const pd = pendingDetailsMap.get(taskId) || item.pendingDetail || {};
              const rawTask = allTasksMap.get(taskId);
              const origStatus =
                item.originalStatus ||
                (rawTask
                  ? (Array.isArray(rawTask.statusDelivery)
                      ? rawTask.statusDelivery[0]
                      : rawTask.statusDelivery
                    )
                      ?.split(',')[0]
                      .trim() || rawTask.status
                  : item.status);
              const currentStatus = pd.new_status || item.status;

              const isConvertedSuccess =
                normalizeStatus(pd.new_status) === DELIVERY_STATUS_KEYS.SUCCESS;
              const isConvertedFromSuccess = Boolean(
                pd.new_status &&
                  normalizeStatus(origStatus) === DELIVERY_STATUS_KEYS.SUCCESS &&
                  normalizeStatus(pd.new_status) !== DELIVERY_STATUS_KEYS.SUCCESS
              );
              const isConvertedToSuccess = Boolean(
                pd.new_status &&
                  normalizeStatus(origStatus) !== DELIVERY_STATUS_KEYS.SUCCESS &&
                  normalizeStatus(pd.new_status) === DELIVERY_STATUS_KEYS.SUCCESS
              );
              const isAmberRow = isConvertedFromSuccess || isConvertedToSuccess;
              const isStatusModified = Boolean(
                pd.new_status && normalizeStatus(pd.new_status) !== normalizeStatus(origStatus)
              );
              const isWrongGR = !hasPendingGR && isPendingGr(currentStatus);

              const hasAnyData = Boolean(
                hasValidText(pd.groupReason) ||
                  hasValidText(pd.detailReason) ||
                  hasValidText(pd.internalExternal) ||
                  hasValidText(pd.pic)
              );
              const isAllEmpty = !hasAnyData;

              const actionCellHover = isAmberRow
                ? 'hover:bg-amber-200/80 dark:hover:bg-amber-900/60'
                : 'hover:bg-sky-50 dark:hover:bg-slate-700/80';
              const actionCellBase = `${BASE_TD_CLASS} ${borderBottomClass} cursor-pointer transition-colors`;
              const actionCellClass = isConvertedSuccess
                ? `${BASE_TD_CLASS} ${borderBottomClass}`
                : isAllEmpty
                  ? `${actionCellBase} bg-red-100 hover:bg-red-200 dark:bg-red-900/40 dark:hover:bg-red-900/60`
                  : `${actionCellBase} ${actionCellHover}`;

              const openModal = () =>
                setModalData({
                  ...item,
                  originalStatus: origStatus,
                  pendingDetail: pd,
                  customer: customerName,
                });

              const reasonValue = isAmberRow
                ? pd.detailReason || item.reason || '-'
                : item.reason || '-';

              return (
                <tr
                  key={item._id || idx}
                  className={
                    isAmberRow
                      ? 'bg-amber-100 dark:bg-amber-400/30 hover:bg-amber-100/90 dark:hover:bg-amber-900/40 transition-colors'
                      : 'hover:bg-gray-50 dark:hover:bg-slate-700 transition-colors'
                  }
                >
                  <td className={tdClass}>{item.flow || '-'}</td>
                  <td className={tdClass}>{item.dateStr}</td>
                  <td className={tdClass}>{getBasePlate(item.licensePlate)}</td>
                  <td className={`${tdClass} text-left`}>{item.driverName}</td>
                  <SOCell
                    text={customerName || '-'}
                    className={`${tdClass} text-left`}
                    isAmberRow={isAmberRow}
                    onClick={() => handleRowClick(item._id)}
                  />
                  <Tooltip tooltipContent={isTruncated ? invoiceNumber : ''}>
                    <td className={`${tdClass} text-left ${isTruncated ? 'cursor-help' : ''}`}>
                      {truncateInvoice || '-'}
                    </td>
                  </Tooltip>
                  <StatusCell
                    item={item}
                    pd={pd}
                    originalStatus={origStatus}
                    currentStatus={currentStatus}
                    isModified={isStatusModified}
                    isWrongGR={isWrongGR}
                    hasPendingGR={hasPendingGR}
                    formatStatus={formatStatus}
                    normalizeStatus={normalizeStatus}
                    isPendingGr={isPendingGr}
                    translate={translate}
                    onUpdatePendingDetail={onUpdatePendingDetail}
                    className={`${tdClass} text-left`}
                    isAmberRow={isAmberRow}
                  />
                  <ReasonCell text={reasonValue} className={`${tdClass} text-left`} />
                  <ActionCell
                    text={isConvertedSuccess ? '-' : pd.groupReason}
                    className={actionCellClass}
                    isAllEmpty={isAllEmpty}
                    disabled={isConvertedSuccess}
                    onClick={openModal}
                  />
                  <ActionCell
                    text={isConvertedSuccess ? '-' : pd.detailReason}
                    className={actionCellClass}
                    isAllEmpty={isAllEmpty}
                    disabled={isConvertedSuccess}
                    onClick={openModal}
                  />
                  <ActionCell
                    text={isConvertedSuccess ? '-' : pd.internalExternal}
                    className={actionCellClass}
                    isAllEmpty={isAllEmpty}
                    disabled={isConvertedSuccess}
                    onClick={openModal}
                  />
                  <ActionCell
                    text={isConvertedSuccess ? '-' : pd.pic}
                    className={actionCellClass}
                    isAllEmpty={isAllEmpty}
                    disabled={isConvertedSuccess}
                    onClick={openModal}
                  />
                  <td className={tdClass}>{item.openStr || '-'}</td>
                  <td className={tdClass}>{item.closeStr || '-'}</td>
                  <EmptyAlertCell
                    value={item.etaStr}
                    className={tdClass}
                    alertStyle={RED_CELL_STYLE}
                  />
                  <EmptyAlertCell
                    value={item.etdStr}
                    className={tdClass}
                    alertStyle={RED_CELL_STYLE}
                  />
                  <td className={tdClass}>{item.arrStr || '-'}</td>
                  <td className={tdClass}>{item.depStr || '-'}</td>
                  <td className={tdClass}>{item.visitTime || '-'}</td>
                  <td
                    className={`${tdClass} ${item.actualVisitMins === 0 ? YELLOW_CELL_STYLE : ''}`}
                  >
                    {item.actualVisitMins}
                  </td>
                  <EmptyAlertCell
                    value={item.routePlannedOrder}
                    className={tdClass}
                    alertStyle={RED_CELL_STYLE}
                  />
                  <td className={tdClass}>{item.realSequence || '-'}</td>
                  <td className={tdClass}>{item.temp}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </Fragment>
  );
}
