'use client';

import HighlightText from '@/components/HighlightText';
import TableData from '@/components/table/TableData';
import { useDeliveryStatus } from '@/lib/hooks/useDeliveryStatus';
import { getBasePlate } from '@/lib/utils';
import { useMemo } from 'react';
import UserAuditCell from './UserAuditCell';

export default function TaskTable({ data = [], searchQuery = '', isLoading, emptyText, translate }) {
  const { formatStatus, getStatusClass } = useDeliveryStatus();

  const columns = useMemo(
    () => [
      {
        label: '#',
        key: 'no',
        align: 'center',
        width: 'w-12',
        render: (_, i) => (
          <div className="flex justify-center text-center">
            <span className="text-xs text-slate-500 dark:text-slate-400">{i + 1}</span>
          </div>
        ),
      },
      {
        label: translate('common.flow'),
        key: 'flow',
        align: 'center',
        sortable: true,
        width: 'w-16',
        render: (row) => (
          <div className="flex justify-center text-center">
            <span className="text-xs font-semibold text-slate-700 dark:text-slate-300">
              <HighlightText text={row.flow || '-'} highlight={searchQuery} />
            </span>
          </div>
        ),
      },
      {
        label: translate('common.customer'),
        key: 'customerName',
        sortable: true,
        width: 'w-[22%]',
        render: (row) => (
          <div className="flex flex-col text-left">
            <span
              className="text-xs font-semibold text-slate-800 dark:text-slate-200 truncate max-w-xs"
              title={row.customerName || '-'}
            >
              <HighlightText text={row.customerName || '-'} highlight={searchQuery} />
            </span>
            <span className="text-[11px] text-slate-500 dark:text-slate-400">
              <HighlightText
                text={row.truncateInvoice || row.orderId || '-'}
                highlight={searchQuery}
              />
            </span>
          </div>
        ),
      },
      {
        label: translate('common.address'),
        key: 'address',
        width: 'w-[18%]',
        render: (row) => (
          <span
            className="text-xs text-slate-700 dark:text-slate-300 line-clamp-2"
            title={row.address}
          >
            <HighlightText text={row.address || '-'} highlight={searchQuery} />
          </span>
        ),
      },
      {
        label: translate('common.status.delivery_status'),
        key: 'statusDelivery',
        align: 'center',
        sortable: true,
        width: 'w-[12%]',
        render: (row) => {
          const st = row.statusDelivery || '-';
          const badgeClass = getStatusClass(st);
          return (
            <div className="flex justify-center text-center">
              <span
                className={`inline-block px-2 py-0.5 rounded text-[10px] font-semibold tracking-wide uppercase ${badgeClass}`}
              >
                <HighlightText text={formatStatus(st)} highlight={searchQuery} />
              </span>
            </div>
          );
        },
      },
      {
        label: translate('common.driver'),
        key: 'driverName',
        sortable: true,
        width: 'w-[14%]',
        render: (row) => (
          <div className="flex flex-col text-left">
            <span className="text-xs font-medium text-slate-800 dark:text-slate-200">
              <HighlightText text={row.driverName || '-'} highlight={searchQuery} />
            </span>
            <span className="text-[11px] text-slate-500 dark:text-slate-400">
              <HighlightText text={getBasePlate(row.plat) || '-'} highlight={searchQuery} />
            </span>
          </div>
        ),
      },
      {
        label: translate('common.created_by') || 'Dibuat Oleh',
        key: 'createdBy',
        width: 'w-[9%]',
        render: (row) => {
          const tu = row.task_user || row.taskUser;
          return (
            <UserAuditCell
              user={tu?.created_by || row.createdBy}
              timestamp={tu?.created_at || row.createdAt}
              highlight={searchQuery}
            />
          );
        },
      },
      {
        label: translate('common.updated_by') || 'Diperbarui Oleh',
        key: 'updatedBy',
        width: 'w-[9%]',
        render: (row) => {
          const tu = row.task_user || row.taskUser;
          return (
            <UserAuditCell
              user={tu?.updated_by || row.updatedBy}
              timestamp={tu?.updated_at || row.updatedAt}
              highlight={searchQuery}
            />
          );
        },
      },
      {
        label: translate('common.sync_by'),
        key: 'syncBy',
        width: 'w-[9%]',
        render: (row) => {
          const su = row.sync_user || row.syncUser;
          return (
            <UserAuditCell
              user={su?.updated_by || su?.created_by}
              timestamp={su?.updated_at || su?.created_at}
              highlight={searchQuery}
            />
          );
        },
      },
    ],
    [formatStatus, getStatusClass, translate, searchQuery]
  );

  const filteredData = useMemo(() => {
    if (!searchQuery.trim()) return data;
    const q = searchQuery.toLowerCase().trim();
    return data.filter(
      (t) =>
        (t.flow || '').toLowerCase().includes(q) ||
        (t.customerOrder || t.customer_order || '').toLowerCase().includes(q) ||
        (t.customerName || '').toLowerCase().includes(q) ||
        (t.truncateInvoice || t.orderId || '').toLowerCase().includes(q) ||
        (t.address || '').toLowerCase().includes(q) ||
        (t.statusDelivery || '').toLowerCase().includes(q) ||
        (t.driverName || '').toLowerCase().includes(q) ||
        (t.plat || '').toLowerCase().includes(q)
    );
  }, [data, searchQuery]);

  return (
    <TableData
      columns={columns}
      data={filteredData}
      isLoading={isLoading}
      paginate={true}
      emptyText={emptyText}
    />
  );
}
