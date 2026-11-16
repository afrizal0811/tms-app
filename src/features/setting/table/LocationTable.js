'use client';

import HighlightText from '@/components/HighlightText';
import TableData from '@/components/table/TableData';
import { formatDateUniversal, formatDurationHHMMSS } from '@/lib/utils';
import { useMemo } from 'react';
import UserAuditCell from './UserAuditCell';

export default function LocationTable({
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
        label: translate('common.driver'),
        key: 'driverName',
        sortable: true,
        width: 'w-[24%]',
        render: (row) => (
          <div className="flex flex-col text-left">
            <span className="text-xs font-semibold text-slate-800 dark:text-slate-200">
              <HighlightText
                text={row.driver_name || row.driverName || '-'}
                highlight={searchQuery}
              />
            </span>
            <span className="text-[11px] text-slate-500 dark:text-slate-400">
              <HighlightText text={row.email || '-'} highlight={searchQuery} />
            </span>
          </div>
        ),
      },
      {
        label: translate('common.start_time'),
        key: 'start_time',
        sortable: true,
        align: 'center',
        width: 'w-[18%]',
        render: (row) => {
          const val = row.start_time || row.startTime;
          return (
            <div className="flex justify-center text-center">
              <span className="text-xs text-slate-700 dark:text-slate-300">
                {val ? formatDateUniversal(val, 'DD-MM-YYYY HH:mm:ss') : '-'}
              </span>
            </div>
          );
        },
      },
      {
        label: translate('common.finish_time'),
        key: 'finish_time',
        sortable: true,
        align: 'center',
        width: 'w-[18%]',
        render: (row) => {
          const finish = row.finish || {};
          const val = finish.finishTime || row.finish_time || row.finishTime;
          return (
            <div className="flex justify-center text-center">
              <span className="text-xs text-slate-700 dark:text-slate-300">
                {val ? formatDateUniversal(val, 'DD-MM-YYYY HH:mm:ss') : '-'}
              </span>
            </div>
          );
        },
      },
      {
        label: translate('common.duration'),
        key: 'duration',
        align: 'center',
        width: 'w-[12%]',
        render: (row) => {
          const finish = row.finish || {};
          const dur = finish.totalDuration ?? row.total_duration ?? row.totalDuration;
          return (
            <div className="flex justify-center text-center">
              <span className="text-xs font-medium text-slate-700 dark:text-slate-300">
                {dur != null && dur !== '' ? formatDurationHHMMSS(dur) : '-'}
              </span>
            </div>
          );
        },
      },
      {
        label: translate('common.distance'),
        key: 'distance',
        align: 'center',
        width: 'w-[12%]',
        render: (row) => {
          const finish = row.finish || {};
          const dist = finish.totalDistance ?? row.total_distance ?? row.totalDistance;
          return (
            <div className="flex justify-center text-center">
              <span className="text-xs font-medium text-slate-700 dark:text-slate-300">
                {dist != null && dist !== '' ? `${Number(dist).toFixed(2)}` : '-'}
              </span>
            </div>
          );
        },
      },
      {
        label: translate('common.sync_by'),
        key: 'syncBy',
        width: 'w-[16%]',
        render: (row) => {
          const su = row.sync_user || row.syncUser;
          return (
            <UserAuditCell
              user={su?.updated_by || su?.created_by || row.updatedBy}
              timestamp={su?.updated_at || su?.created_at || row.updatedAt}
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
      (l) =>
        (l.driver_name || l.driverName || '').toLowerCase().includes(q) ||
        (l.email || '').toLowerCase().includes(q)
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
