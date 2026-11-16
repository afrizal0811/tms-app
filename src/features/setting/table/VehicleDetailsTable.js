'use client';

import HighlightText from '@/components/HighlightText';
import TableData from '@/components/table/TableData';
import { useMemo } from 'react';
import UserAuditCell from './UserAuditCell';

export default function VehicleDetailsTable({
  data = [],
  searchQuery = '',
  isLoading,
  emptyText,
  translate,
}) {
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
        label: translate('common.license_number'),
        key: 'plat',
        sortable: true,
        render: (row) => (
          <span className="text-xs font-semibold text-slate-800 dark:text-slate-200 whitespace-nowrap">
            <HighlightText text={row.plat || '-'} highlight={searchQuery} />
          </span>
        ),
      },
      {
        label: translate('common.driver'),
        key: 'driver_name',
        sortable: true,
        render: (row) => (
          <span className="text-xs font-medium text-slate-800 dark:text-slate-200">
            <HighlightText text={row.driver_name || '-'} highlight={searchQuery} />
          </span>
        ),
      },
      {
        label: translate('common.vehicle_type'),
        key: 'vehicle_type',
        align: 'center',
        sortable: true,
        render: (row) => (
          <span className="text-xs text-slate-700 dark:text-slate-300">
            <HighlightText text={row.vehicle_type || '-'} highlight={searchQuery} />
          </span>
        ),
      },
      {
        label: 'Storage',
        key: 'storage_type',
        align: 'center',
        sortable: true,
        render: (row) => (
          <span className="text-xs text-slate-700 dark:text-slate-300">
            <HighlightText text={row.storage_type || '-'} highlight={searchQuery} />
          </span>
        ),
      },
      {
        label: 'Weight (Kg)',
        key: 'act_weight',
        align: 'center',
        sortable: true,
        render: (row) => (
          <span className="text-xs font-medium text-slate-700 dark:text-slate-300 tabular-nums whitespace-nowrap">
            {row.act_weight ?? 0} / {row.max_weight ?? 0}
          </span>
        ),
      },
      {
        label: 'Volume (Cbm)',
        key: 'act_volume',
        align: 'center',
        sortable: true,
        render: (row) => (
          <span className="text-xs font-medium text-slate-700 dark:text-slate-300 tabular-nums whitespace-nowrap">
            {row.act_volume ?? 0} / {row.max_volume ?? 0}
          </span>
        ),
      },
      {
        label: `${translate('common.distance')}`,
        key: 'total_distance',
        align: 'center',
        sortable: true,
        render: (row) => (
          <span className="text-xs text-slate-700 dark:text-slate-300 tabular-nums">
            {row.total_distance ?? 0}
          </span>
        ),
      },
      {
        label: translate('common.task'),
        key: 'task_data',
        align: 'center',
        render: (row) => {
          const td = row.task_data || {};
          const parts = [
            td.success ? `Su:${td.success}` : null,
            td.pending ? `Pe:${td.pending}` : null,
            td.cancel ? `Ca:${td.cancel}` : null,
            td.partial ? `Par:${td.partial}` : null,
            td.pending_gr ? `GR:${td.pending_gr}` : null,
          ].filter(Boolean);
          return (
            <div className="flex flex-col items-center text-xs">
              <span className="font-semibold text-slate-800 dark:text-slate-200">
                {td.total ?? 0}
              </span>
              {parts.length > 0 && (
                <span className="text-[10px] text-slate-500 dark:text-slate-400">
                  {parts.join(' | ')}
                </span>
              )}
            </div>
          );
        },
      },
      {
        label: translate('common.sync_by'),
        key: 'syncBy',
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
    [translate, searchQuery]
  );

  const filteredData = useMemo(() => {
    if (!searchQuery.trim()) return data;
    const q = searchQuery.toLowerCase().trim();
    return data.filter(
      (df) =>
        (df.plat || '').toLowerCase().includes(q) ||
        (df.driver_name || '').toLowerCase().includes(q) ||
        (df.vehicle_type || '').toLowerCase().includes(q) ||
        (df.storage_type || '').toLowerCase().includes(q)
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
