'use client';

import Accordion from '@/components/Accordion';
import Spinner from '@/components/Spinner';
import TableData from '@/components/table/TableData';
import { useLanguage } from '@/context/LanguageContext';
import { getResult, getTask } from '@/lib/api/mileapp';
import { useDeliveryStatus } from '@/lib/hooks/useDeliveryStatus';
import { useSuperadmin } from '@/lib/hooks/useSuperadmin';
import { toastError } from '@/lib/toast';
import {
  formatDateUniversal,
  formatMinutesToHHMM,
  isEmpty,
  parseCustomerString,
  ProperCaseText,
} from '@/lib/utils';
import { useEffect, useState } from 'react';
import { getPendingDetailByTaskId } from '@/lib/api/mileapp/pendingDetails';
import {
  DELIVERY_STATUS_KEYS,
  deliveryTooltip,
  normalizeDeliveryStatus,
} from '@/lib/deliveryStatus';
import InformationButton from '../button/InformationButton';
import CopyButton from '../button/CopyButton';
import JsonTree from '../JsonTree';
import MapView from '../Map';
import Tooltip from '../Tooltip';
import Modal from './Modal';

const parseCoord = (coordStr) => {
  if (!coordStr) return null;
  const [lat, lng] = coordStr.split(',').map(Number);
  if (isNaN(lat) || isNaN(lng)) return null;
  return [lat, lng];
};

const renderCoordinate = (val) => {
  if (!val) return '-';
  const [lat, lng] = val.split(',').map((coord) => Number(coord).toFixed(7));
  return `${lat}, ${lng}`;
};

const getSubtitle = (taskData, pendingDetail, formatStatus, translate) => {
  if (!taskData) return '';
  const status = (taskData.status || '-').toUpperCase();
  const originalDelivery = taskData.statusDelivery || taskData.status;
  const isModified = Boolean(
    pendingDetail?.new_status &&
      originalDelivery &&
      normalizeDeliveryStatus(pendingDetail.new_status) !==
        normalizeDeliveryStatus(originalDelivery)
  );

  const effectiveDelivery = isModified ? pendingDetail.new_status : taskData.statusDelivery;
  if (!effectiveDelivery) return status;

  const formattedDelivery = formatStatus ? formatStatus(effectiveDelivery) : effectiveDelivery;

  if (isModified) {
    const tooltipContent = deliveryTooltip({
      originalStatus: originalDelivery,
      pendingDetail,
      formatStatus,
      t: translate,
    });

    return (
      <span className="inline-flex items-center gap-1.5 flex-wrap">
        <span>{status} |</span>
        <span className="font-bold text-red-600 dark:text-red-400">
          {formattedDelivery.toUpperCase()}
        </span>
        <InformationButton infoText={tooltipContent} size="4" />
      </span>
    );
  }

  const subtitleText = `${status} | ${formattedDelivery}`;
  return subtitleText.toUpperCase();
};

const getActionStyle = (action) => {
  const act = (action || '').toLowerCase();
  switch (act) {
    case 'create':
      return 'bg-green-100 text-green-700 dark:bg-green-900/40 dark:text-green-400';
    case 'patch':
      return 'bg-amber-100 text-amber-700 dark:bg-amber-900/40 dark:text-amber-400';
    case 'update':
    case 'put':
      return 'bg-blue-100 text-blue-700 dark:bg-blue-900/40 dark:text-blue-400';
    case 're-assign':
    case 'reassign':
      return 'bg-indigo-100 text-indigo-700 dark:bg-indigo-900/40 dark:text-indigo-400';
    case 'assign':
      return 'bg-purple-100 text-purple-700 dark:bg-purple-900/40 dark:text-purple-400';
    case 'unassign':
      return 'bg-pink-100 text-pink-700 dark:bg-pink-900/40 dark:text-pink-400';
    case 'complete':
      return 'bg-teal-100 text-teal-700 dark:bg-teal-900/40 dark:text-teal-400';
    default:
      return 'bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300';
  }
};

const formatVal = (v) => {
  if (v === null) return 'null';
  if (v === undefined) return 'undefined';
  if (typeof v === 'object') return JSON.stringify(v);
  const str = String(v);
  if (/^\d{4}-\d{2}-\d{2}T/.test(str)) {
    return formatDateUniversal(str, 'DD-MM-YYYY HH:mm');
  }
  return str;
};

const extractPutDifferences = (historiesList) => {
  if (!historiesList || historiesList.length === 0) return [];

  const sorted = [...historiesList].sort(
    (a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime()
  );

  const isObjOrArray = (val) => val !== null && typeof val === 'object';
  const isNullish = (val) => val === undefined || val === null;

  let currentState = {};
  const differences = [];

  sorted.forEach((history) => {
    const act = (history.action || '').toLowerCase();
    const isPut = act === 'put' || act === 'update';
    const cData = history.changeData || {};

    if (isPut) {
      const diffsForThisPut = [];

      Object.keys(cData).forEach((key) => {
        const newVal = cData[key];
        const oldVal = currentState[key];
        if (isNullish(oldVal) && isNullish(newVal)) return;
        if (oldVal === newVal) return;
        if (isObjOrArray(newVal) && isObjOrArray(oldVal)) {
          if (JSON.stringify(newVal) === JSON.stringify(oldVal)) return;
        }
        if (formatVal(oldVal) === formatVal(newVal)) return;
        diffsForThisPut.push({ key, oldVal, newVal });
      });

      differences.push({ id: history._id || history.createdAt, diffs: diffsForThisPut });
    }

    currentState = { ...currentState, ...cData };
  });

  return differences;
};

const computeReassignSet = (rawHistories) => {
  const reassignSet = new Set();
  let hasPreviousAssign = false;

  const chronological = rawHistories
    .map((item, originalIndex) => ({ item, originalIndex }))
    .sort((a, b) => {
      const timeA = new Date(a.item.createdAt).getTime();
      const timeB = new Date(b.item.createdAt).getTime();
      if (timeA !== timeB) return timeA - timeB;
      return a.originalIndex - b.originalIndex;
    })
    .map(({ item }) => item);

  chronological.forEach((h) => {
    const act = (h.action || '').toLowerCase();
    if (act === 'assign' || act === 're-assign' || act === 'reassign') {
      if (!hasPreviousAssign) {
        hasPreviousAssign = true;
      } else {
        reassignSet.add(h);
      }
    }
  });

  return reassignSet;
};

const getProductColumns = (translate) => [
  {
    key: 'caption',
    sortable: true,
    label: translate('common.invoice_number'),
    render: (row) => row.caption || '-',
  },
  {
    key: 'title',
    sortable: true,
    label: translate('common.items'),
    render: (row) => (
      <div className="max-w-xs truncate" title={row.title}>
        {row.title || '-'}
      </div>
    ),
  },
  {
    key: 'qtyProcessed',
    sortable: true,
    label: translate('common.quantity'),
    render: (row) => row.qtyProcessed ?? '-',
  },
  {
    key: 'content',
    sortable: true,
    label: 'UOM',
    render: (row) => row.content || '-',
  },
  {
    key: 'volume',
    sortable: true,
    label: translate('common.volume'),
    render: (row) => (row.volume ? row.volume.toFixed(2) : '-'),
  },
  {
    key: 'weight',
    sortable: true,
    label: translate('common.weight'),
    render: (row) => (row.weight ? row.weight.toFixed(2) : '-'),
  },
];

const Field = ({
  label,
  value,
  tooltip,
  isCopy,
  copyValue,
  isTruncated,
  onValueClick,
  needEmpty = false,
  valueClassName = '',
}) => (
  <div className="mb-3">
    <div className="text-xs text-gray-500 dark:text-slate-400">{label}</div>
    <Tooltip tooltipContent={tooltip}>
      <div
        className={`text-sm font-medium ${valueClassName || 'text-slate-800 dark:text-slate-200'} ${onValueClick ? ' cursor-pointer underline decoration-dotted decoration-2 underline-offset-5' : ''} ${isCopy && 'flex items-center gap-1'} ${isTruncated && 'truncate'} ${tooltip && 'cursor-help'}`}
        onClick={onValueClick}
      >
        {!needEmpty && isEmpty(value) ? '-' : String(value)}{' '}
        {!isEmpty(value) && isCopy && <CopyButton text={copyValue || value} />}
      </div>
    </Tooltip>
  </div>
);

function TaskOverview({ taskData, translate }) {
  const custInfo = parseCustomerString(taskData.customerOrder);
  const assigneeEmail = taskData.assignee;
  const maxVehicle = taskData?.maksimumVehicleType || '-';
  const assigneeName = taskData.driverName || assigneeEmail || '-';
  const licenseNumber = taskData.basePlat || taskData.plat || '-';

  const isAutomation =
    taskData.createdFrom?.toLowerCase() === 'automation' ||
    taskData.createdBy?.toLowerCase() === 'system';

  return (
    <>
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 p-4 border border-gray-200 dark:border-slate-700 rounded-lg bg-gray-50 dark:bg-slate-900/50">
        <Field
          label={translate('common.customer_name')}
          value={custInfo.fullCustomerName}
          isCopy={true}
        />
        <Field
          label={translate('common.address')}
          value={taskData.address}
          isTruncated={true}
          tooltip={taskData.address}
        />
        <Field
          label={translate('common.invoice_number')}
          value={custInfo.truncateInvoice || custInfo.invoiceNumber}
          tooltip={custInfo.isTruncated ? custInfo.invoiceNumber : null}
          isCopy={true}
          copyValue={custInfo.invoiceNumber}
        />
        <Field label={translate('common.storage_type')} value={taskData.typeStorage} />
        <Field label={translate('common.vehicle_type')} value={maxVehicle} />
        <Field
          label={translate('common.updated_by')}
          value={taskData.updatedByName || taskData.updatedBy}
          tooltip={taskData.updatedBy}
          isCopy={true}
          copyValue={taskData.updatedBy}
        />
        <Field
          label={translate('common.updated_time')}
          value={formatDateUniversal(taskData.updatedTime, 'DD/MM/YYYY HH:mm')}
        />
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div className="border border-gray-200 dark:border-slate-700 rounded-lg p-4 bg-white dark:bg-slate-800">
          <h3 className="font-bold text-sky-600 dark:text-sky-400 border-b border-gray-200 dark:border-slate-700 pb-2 mb-4">
            {translate('task_detail.modal.creation')}
          </h3>
          <div className="grid grid-cols-2 gap-4">
            <Field
              label={translate('common.created_by')}
              value={taskData.createdByName || taskData.createdBy}
              tooltip={isAutomation ? null : taskData.createdBy}
              isCopy={!isAutomation}
              copyValue={isAutomation ? null : taskData.createdBy}
            />
            <Field
              label={translate('common.created_time')}
              value={formatDateUniversal(taskData.createdTime, 'DD/MM/YYYY HH:mm')}
            />
            <Field label={translate('common.created_from')} value={taskData.createdFrom} />
            <Field
              label={translate('common.start_time')}
              value={formatDateUniversal(taskData.startTime, 'DD/MM/YYYY HH:mm')}
            />
          </div>
        </div>

        <div className="border border-gray-200 dark:border-slate-700 rounded-lg p-4 bg-white dark:bg-slate-800">
          <h3 className="font-bold text-sky-600 dark:text-sky-400 border-b border-gray-200 dark:border-slate-700 pb-2 mb-4">
            {translate('task_detail.modal.assigment')}
          </h3>
          <div className="grid grid-cols-2 gap-4">
            <Field
              label={translate('common.driver')}
              value={assigneeName}
              tooltip={assigneeName === assigneeEmail ? null : assigneeEmail}
              isCopy={true}
              copyValue={assigneeEmail}
            />
            <Field label={translate('common.license_number')} value={licenseNumber} />
            <Field
              label={translate('common.assigned_time')}
              value={formatDateUniversal(taskData.assignedTime, 'DD/MM/YYYY HH:mm')}
            />
            <Field
              label={translate('common.done_time')}
              value={formatDateUniversal(taskData.doneTime, 'DD/MM/YYYY HH:mm')}
            />
          </div>
        </div>
      </div>
    </>
  );
}

function DataTab({ taskData, pendingDetail, goToMapTab, translate }) {
  let arrivalSource, departureSource;
  const isGrOrPickup =
    taskData.flow?.toUpperCase().includes('GR') || taskData.flow?.toUpperCase().includes('PICKUP');

  if (isGrOrPickup) {
    arrivalSource = taskData.page1DoneTime;
    departureSource = taskData.doneTime;
  } else {
    arrivalSource = taskData.klikJikaSudahSampai || taskData.klikJikaAndaSudahSampai;
    departureSource = taskData.page3DoneTime;
  }

  const arrivalDate = formatDateUniversal(arrivalSource, 'DD/MM/YYYY HH:mm');
  const departureDate = formatDateUniversal(departureSource, 'DD/MM/YYYY HH:mm');

  let actualVisitMins = 0;
  if (arrivalSource && departureSource) {
    const tArr = new Date(arrivalSource);
    const tDep = new Date(departureSource);
    tArr.setSeconds(0, 0);
    tDep.setSeconds(0, 0);
    const diff = tDep.getTime() - tArr.getTime();
    actualVisitMins = diff > 0 ? Math.floor(diff / (1000 * 60)) : 0;
  } else {
    actualVisitMins = '-';
  }

  const actualVisitTooltip =
    actualVisitMins > 60 ? `${actualVisitMins} ${translate('common.minute')}` : '';
  const travelDurationTooltip =
    taskData?.travelDuration > 60
      ? `${taskData?.travelDuration} ${translate('common.minute')}`
      : '';
  const actualVisitHour = formatMinutesToHHMM(actualVisitMins, true);

  const newCoordVal =
    taskData?.newCoordinate || taskData?.new_coordinate || taskData?.klikLokasiClient || null;

  const activePendingDetail = pendingDetail || taskData?.pendingDetail || null;
  const originalDelivery = taskData.originalStatus || taskData.statusDelivery || taskData.status;
  const originalNorm = normalizeDeliveryStatus(originalDelivery);
  const newNorm = activePendingDetail?.new_status
    ? normalizeDeliveryStatus(activePendingDetail.new_status)
    : null;

  const isConvertedFromSuccess = Boolean(
    newNorm &&
      originalNorm === DELIVERY_STATUS_KEYS.SUCCESS &&
      newNorm !== DELIVERY_STATUS_KEYS.SUCCESS
  );

  const isConvertedToSuccess = Boolean(
    newNorm &&
      originalNorm !== DELIVERY_STATUS_KEYS.SUCCESS &&
      newNorm === DELIVERY_STATUS_KEYS.SUCCESS
  );

  let showReason = false;
  let reasonValue = '';
  let isReasonRed = false;

  if (isConvertedToSuccess) {
    showReason = false;
  } else if (isConvertedFromSuccess) {
    showReason = true;
    reasonValue = activePendingDetail.detailReason || activePendingDetail.groupReason || '-';
    isReasonRed = true;
  } else if (taskData.reason || activePendingDetail?.detailReason) {
    showReason = true;
    reasonValue = taskData.reason || activePendingDetail?.detailReason || '-';
    isReasonRed = false;
  }

  return (
    <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
      <Field
        label={translate('common.task_id')}
        value={taskData._id}
        isCopy={true}
        isTruncated={true}
      />
      <Field label={translate('common.open_time')} value={taskData.openTime} />
      <Field label={translate('common.close_time')} value={taskData.closeTime} />
      <Field label={translate('common.volume')} value={taskData.volumeCbm} needEmpty={true} />
      <Field label={translate('common.weight')} value={taskData.weightKg} needEmpty={true} />
      <Field label={translate('common.actual_arrival')} value={arrivalDate} />
      <Field label={translate('common.actual_departure')} value={departureDate} />
      <Field
        label={translate('common.actual_visit')}
        tooltip={actualVisitTooltip}
        value={actualVisitHour}
        needEmpty={true}
      />
      <Field
        label={translate('common.actual_seq')}
        value={taskData.doneOrder > 0 ? taskData.doneOrder : '-'}
      />
      <Field
        label={translate('task_detail.modal.expected_coord')}
        value={renderCoordinate(taskData.longlat)}
        isCopy={true}
        onValueClick={goToMapTab}
      />
      <Field
        label={translate('task_detail.modal.done_coord')}
        value={renderCoordinate(taskData.doneCoordinate)}
        isCopy={true}
        onValueClick={goToMapTab}
      />
      {newCoordVal && (
        <Field
          label={translate('task_detail.modal.new_coord')}
          value={renderCoordinate(newCoordVal)}
          isCopy={true}
          onValueClick={goToMapTab}
        />
      )}
      {showReason && (
        <Field
          label={translate('task_detail.modal.reason')}
          value={reasonValue}
          valueClassName={isReasonRed ? 'text-red-600 dark:text-red-400 font-bold' : ''}
          isTruncated={true}
          tooltip={reasonValue && String(reasonValue).length > 40 ? String(reasonValue) : ''}
        />
      )}
      <Field label={translate('common.actual_travel_distance')} value={taskData.travelDistance} />
      <Field
        label={translate('common.actual_travel_duration')}
        value={taskData?.travelDurationHour}
        tooltip={travelDurationTooltip}
      />
    </div>
  );
}

function RoutingTab({ taskData, resultData, translate }) {
  const rName = resultData?.name || '-';

  let visitMins = Number(taskData?.visitTime ?? taskData?.others?.visitTime);
  if (isNaN(visitMins) && taskData?.eta && taskData?.etd) {
    const [etaH, etaM] = taskData.eta.split(':').map(Number);
    const [etdH, etdM] = taskData.etd.split(':').map(Number);
    if (!isNaN(etaH) && !isNaN(etdH)) {
      let diff = etdH * 60 + etdM - (etaH * 60 + etaM);
      if (diff < 0) diff += 24 * 60;
      visitMins = diff;
    }
  }
  const planVisitMins = !isNaN(visitMins) ? visitMins : 0;
  const planVisitHour = formatMinutesToHHMM(planVisitMins);
  const planVisitTooltip = planVisitMins > 0 ? `${planVisitMins} ${translate('common.minute')}` : '';

  const waitingMins = Number(taskData?.waitingTime ?? taskData?.others?.waitingTime ?? 0) || 0;
  const waitingHour = formatMinutesToHHMM(waitingMins);
  const waitingTooltip = waitingMins > 0 ? `${waitingMins} ${translate('common.minute')}` : '';

  let plannedTravelMins =
    taskData?.plannedTravelDuration != null ? Number(taskData.plannedTravelDuration) : null;
  let plannedTravelHour =
    taskData?.plannedTravelDurationHour ||
    (plannedTravelMins != null ? formatMinutesToHHMM(plannedTravelMins) : null);

  if (!plannedTravelHour && taskData?.travelDuration != null) {
    plannedTravelMins = Math.round(Number(taskData.travelDuration));
    plannedTravelHour = formatMinutesToHHMM(plannedTravelMins);
  }

  const plannedTravelTooltip =
    plannedTravelMins != null && plannedTravelMins > 0
      ? `${plannedTravelMins} ${translate('common.minute')}`
      : '';

  return (
    <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
      <Field
        label={translate('common.routing_id')}
        value={taskData?.routingResultId}
        isCopy={true}
        isTruncated={true}
      />
      <Field
        label={translate('common.routing_name')}
        value={rName}
        isCopy={true}
        isTruncated={true}
      />
      <Field label={translate('common.eta')} value={taskData?.eta} />
      <Field label={translate('common.etd')} value={taskData?.etd} />
      <Field label={translate('common.plan_visit')} tooltip={planVisitTooltip} value={planVisitHour} />
      <Field label={translate('common.plan_seq')} value={taskData?.routePlannedOrder} />
      <Field label={translate('common.plan_travel_distance')} value={taskData?.distance} />
      <Field
        label={translate('common.plan_travel_duration')}
        tooltip={plannedTravelTooltip}
        value={plannedTravelHour || '-'}
      />
      <Field
        label={translate('task_detail.modal.waiting_time')}
        tooltip={waitingTooltip}
        value={waitingHour}
      />
    </div>
  );
}

function ProductTab({ products, translate }) {
  const uniqueProducts = new Set(products.map((p) => p.title)).size;
  const totalItems = products.reduce((acc, p) => acc + (Number(p.qtyProcessed) || 0), 0);
  const columns = getProductColumns(translate);

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 px-2">
        <Field label={translate('task_detail.modal.total_product')} value={uniqueProducts} />
        <Field label={translate('task_detail.modal.total_item')} value={totalItems} />
      </div>
      <div className="overflow-hidden rounded-lg border border-gray-200 dark:border-slate-700 flex flex-col max-h-[60vh]">
        <TableData columns={columns} data={products} emptyMessage={translate('common.no_data')} />
      </div>
    </div>
  );
}

function HistoryTab({ histories, putDiffs, reassignSet, translate }) {
  if (!histories || histories.length === 0) {
    return (
      <div className="text-center text-sm text-gray-500 p-4">{translate('common.no_data')}</div>
    );
  }

  return (
    <div className="py-2 max-h-[60vh] overflow-y-auto pr-2">
      <div className="border-l-2 border-sky-300 dark:border-sky-700 ml-4 space-y-6">
        {histories.map((h, i) => {
          const act = (h.action || '').toLowerCase();
          let displayAction = act;
          let badgeText = h.action || '-';

          if (act === 'assign' || act === 're-assign' || act === 'reassign') {
            const isReassign = reassignSet.has(h);
            displayAction = isReassign ? 're-assign' : 'assign';
            badgeText = isReassign ? 'Re-Assign' : 'Assign';
          } else if (act === 'put') {
            displayAction = 'update';
            badgeText = 'Update';
          }

          const diffMatch =
            h.action &&
            (act === 'put' || act === 'update') &&
            putDiffs.find((d) => d.id === (h._id || h.createdAt));

          return (
            <div key={i} className="relative pl-6">
              <div className="absolute -left-[9px] top-1.5 h-4 w-4 rounded-full bg-sky-500 ring-4 ring-white dark:ring-slate-800" />
              <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between mb-1">
                <span
                  className={`text-[11px] px-2 py-0.5 rounded-md font-bold capitalize w-fit tracking-wide ${getActionStyle(displayAction)}`}
                >
                  {badgeText}
                </span>
                <span className="text-xs text-gray-500 dark:text-slate-400 mt-1 sm:mt-0">
                  {formatDateUniversal(h.createdAt, 'DD/MM/YYYY HH:mm')}
                </span>
              </div>
              <div className="text-xs font-medium text-sky-600 dark:text-sky-400 mb-1">
                {h.changedBy || '-'}
              </div>
              <div className="text-sm text-slate-600 dark:text-slate-300 wrap-break-words">
                {h.notes || '-'}
                {diffMatch && diffMatch.diffs.length > 0 && (
                  <details className="mt-1.5 group">
                    <summary className="text-[11px] text-sky-600 dark:text-sky-400 cursor-pointer select-none font-medium hover:underline outline-none w-fit">
                      {translate('task_detail.modal.view_changes')} ({diffMatch.diffs.length})
                    </summary>
                    <div className="mt-1.5 pl-2.5 border-l-2 border-slate-200 dark:border-slate-700 space-y-1.5 max-h-[150px] overflow-y-auto">
                      {diffMatch.diffs.map((d, idx) => (
                        <div
                          key={idx}
                          className="text-[10px] font-mono bg-slate-100 dark:bg-slate-900/60 p-1.5 rounded"
                        >
                          <div className="font-bold text-slate-700 dark:text-slate-300 mb-0.5">
                            {ProperCaseText(d.key)}
                          </div>
                          <div
                            className="text-red-500 dark:text-red-400 truncate"
                            title={formatVal(d.oldVal)}
                          >
                            - {formatVal(d.oldVal)}
                          </div>
                          <div
                            className="text-emerald-600 dark:text-emerald-400 truncate"
                            title={formatVal(d.newVal)}
                          >
                            + {formatVal(d.newVal)}
                          </div>
                        </div>
                      ))}
                    </div>
                  </details>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function MapTab({ taskData, translate, isIndonesian }) {
  const expectedCoord = parseCoord(taskData.longlat);
  const doneCoord = parseCoord(taskData.doneCoordinate);
  const newCoord = parseCoord(
    taskData.newCoordinate || taskData.new_coordinate || taskData.klikLokasiClient
  );
  const mapBounds = [expectedCoord, doneCoord, newCoord].filter(Boolean);

  return (
    <div className="w-full h-[50vh] relative z-0 border border-gray-200 dark:border-slate-700 rounded-lg overflow-hidden">
      <MapView
        bounds={mapBounds}
        onMapReady={(map) => {
          map.dragging.disable();
          setTimeout(() => {
            map.setMinZoom(map.getZoom());
          }, 500);
        }}
      >
        {(rl, L, icons) => (
          <>
            {expectedCoord && (
              <rl.Marker
                position={expectedCoord}
                icon={icons.circle('E', 'bg-blue-500', 'text-xs')}
              >
                <rl.Tooltip direction="top" offset={[0, -10]}>
                  {translate('task_detail.modal.expected_coord')}
                </rl.Tooltip>
              </rl.Marker>
            )}
            {doneCoord && (
              <rl.Marker
                position={doneCoord}
                icon={icons.circle(isIndonesian ? 'S' : 'D', 'bg-green-500', 'text-xs')}
              >
                <rl.Tooltip direction="top" offset={[0, -10]}>
                  {translate('task_detail.modal.done_coord')}
                </rl.Tooltip>
              </rl.Marker>
            )}
            {newCoord && (
              <rl.Marker
                position={newCoord}
                icon={icons.circle(isIndonesian ? 'B' : 'N', 'bg-orange-500', 'text-xs')}
              >
                <rl.Tooltip direction="top" offset={[0, -10]}>
                  {translate('task_detail.modal.new_coord')}
                </rl.Tooltip>
              </rl.Marker>
            )}
          </>
        )}
      </MapView>
    </div>
  );
}

function JsonTab({ data, expandAll = false }) {
  return (
    <div className="relative bg-slate-950 border border-slate-800 rounded-lg p-4 max-h-[60vh] overflow-y-auto">
      <div className="absolute top-2 right-2 z-10 bg-slate-900/80 rounded-md">
        <CopyButton
          text={JSON.stringify(data, null, 2)}
          className="text-white hover:text-sky-300"
        />
      </div>
      <div className="text-[13px] font-mono leading-relaxed">
        <JsonTree data={data} label={null} isLast={true} defaultOpen={true} expandAll={expandAll} />
      </div>
    </div>
  );
}

export default function TaskModal({ isOpen, onClose, taskId, isAllHub = false }) {
  const [loading, setLoading] = useState(false);
  const [taskData, setTaskData] = useState(null);
  const [pendingDetail, setPendingDetail] = useState(null);
  const [activeTab, setActiveTab] = useState('Data');
  const [resultData, setResultData] = useState(null);

  const { t: translate, isIndonesian } = useLanguage();
  const { isSuperadmin } = useSuperadmin();
  const { formatStatus } = useDeliveryStatus();

  useEffect(() => {
    if (!isOpen || !taskId || taskId === '-') {
      setTaskData(null);
      setPendingDetail(null);
      setResultData(null);
      setActiveTab('Data');
      return;
    }

    const loadData = async () => {
      setLoading(true);
      try {
        const [taskRes, pendingRes] = await Promise.allSettled([
          getTask(taskId, isAllHub),
          getPendingDetailByTaskId(taskId),
        ]);

        const task =
          taskRes.status === 'fulfilled' && taskRes.value
            ? taskRes.value?.task || taskRes.value
            : null;
        setTaskData(task);

        const pd =
          pendingRes.status === 'fulfilled' && pendingRes.value?.taskId ? pendingRes.value : null;
        setPendingDetail(pd);

        const [resultRes] = await Promise.allSettled([
          task?.routingResultId ? getResult(task.routingResultId) : Promise.resolve(null),
        ]);

        setResultData(resultRes.status === 'fulfilled' && resultRes.value ? resultRes.value : null);
      } catch (err) {
        toastError(translate('common.toast.error', { err: err.message }), err);
      } finally {
        setLoading(false);
      }
    };

    loadData();
  }, [isOpen, taskId, translate, isAllHub]);

  if (loading) {
    return (
      <Modal
        isOpen={isOpen}
        onClose={onClose}
        title={translate('common.task_details')}
        maxWidth="max-w-5xl lg:max-w-6xl"
        bodyClassName="p-6 overflow-y-auto"
      >
        <div className="flex justify-center items-center p-10 min-h-[300px]">
          <Spinner />
        </div>
      </Modal>
    );
  }

  if (!taskData) {
    return (
      <Modal
        isOpen={isOpen}
        onClose={onClose}
        title={translate('common.task_details')}
        subtitle=""
        maxWidth="max-w-5xl lg:max-w-6xl"
        bodyClassName="p-6 overflow-y-auto"
      >
        <div className="p-4 text-center text-gray-500">{translate('common.no_data')}</div>
      </Modal>
    );
  }

  const hasMap = !!taskData.doneCoordinate;
  const goToMapTab = hasMap ? () => setActiveTab(translate('task_detail.modal.map')) : undefined;

  const rawHistories = taskData.histories || [];
  const reassignSet = computeReassignSet(rawHistories);
  const histories = [...rawHistories].sort(
    (a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
  );
  const putDiffs = extractPutDifferences(rawHistories);
  const products = taskData.listProduct || [];

  const tabs = [
    'Data',
    translate('common.routing'),
    translate('task_detail.modal.list_product'),
    translate('task_detail.modal.history'),
    ...(hasMap ? [translate('task_detail.modal.map')] : []),
    ...(isSuperadmin && taskData ? [`JSON ${translate('common.task')}`] : []),
    ...(isSuperadmin && resultData ? [`JSON ${translate('common.routing')}`] : []),
  ];

  return (
    <Modal
      isOpen={isOpen}
      onClose={onClose}
      title={translate('common.task_details')}
      subtitle={getSubtitle(taskData, pendingDetail, formatStatus, translate)}
      maxWidth="max-w-5xl lg:max-w-6xl"
      bodyClassName="p-6 overflow-y-auto"
    >
      <div className="space-y-6">
        <TaskOverview taskData={taskData} translate={translate} />

        <Accordion title={translate('common.more')} defaultOpen={false}>
          <div className="flex overflow-x-auto space-x-6 border-b border-gray-200 dark:border-slate-700 mb-4 px-2">
            {tabs.map((tab) => (
              <button
                key={tab}
                onClick={() => setActiveTab(tab)}
                className={`pb-2 whitespace-nowrap text-sm font-medium transition-colors outline-none ${
                  activeTab === tab
                    ? 'border-b-2 border-sky-500 text-sky-600 dark:text-sky-400'
                    : 'text-gray-500 hover:text-gray-700 dark:text-slate-400 dark:hover:text-slate-200 cursor-pointer'
                }`}
              >
                {tab}
              </button>
            ))}
          </div>

          <div className="p-2 min-h-[200px]">
            {activeTab === 'Data' && (
              <DataTab
                taskData={taskData}
                pendingDetail={pendingDetail}
                goToMapTab={goToMapTab}
                translate={translate}
              />
            )}
            {activeTab === translate('common.routing') && (
              <RoutingTab taskData={taskData} resultData={resultData} translate={translate} />
            )}
            {activeTab === translate('task_detail.modal.list_product') && (
              <ProductTab products={products} translate={translate} />
            )}
            {activeTab === translate('task_detail.modal.history') && (
              <HistoryTab
                histories={histories}
                putDiffs={putDiffs}
                reassignSet={reassignSet}
                translate={translate}
              />
            )}
            {activeTab === translate('task_detail.modal.map') && hasMap && (
              <MapTab taskData={taskData} translate={translate} isIndonesian={isIndonesian} />
            )}
            {activeTab === `JSON ${translate('common.task')}` && <JsonTab data={taskData} />}
            {activeTab === `JSON ${translate('common.routing')}` && (
              <JsonTab data={resultData} expandAll={true} />
            )}
          </div>
        </Accordion>
      </div>
    </Modal>
  );
}
