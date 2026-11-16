'use client';

import HighlightText from '@/components/HighlightText';
import TableData from '@/components/table/TableData';
import { useMemo } from 'react';
import UserAuditCell from './UserAuditCell';

export default function RoutingTable({ data = [], searchQuery = '', isLoading, emptyText, translate }) {
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
        label: translate('common.routing_name'),
        key: 'name',
        sortable: true,
        width: 'w-[36%]',
        render: (row) => (
          <span className="text-xs font-semibold text-slate-800 dark:text-slate-200 break-words">
            <HighlightText text={row.name || '-'} highlight={searchQuery} />
          </span>
        ),
      },
      {
        label: translate('common.dropped_task'),
        key: 'dropped',
        align: 'center',
        sortable: true,
        width: 'w-[12%]',
        render: (row) => (
          <div className="flex justify-center text-center">
            <span className="text-xs font-medium text-slate-600 dark:text-slate-400">
              {typeof row.dropped === 'object' && row.dropped !== null
                ? (row.dropped.total ?? 0)
                : (row.dropped ?? 0)}
            </span>
          </div>
        ),
      },
      {
        label: translate('common.histories'),
        key: 'histories',
        align: 'center',
        sortable: true,
        width: 'w-[14%]',
        render: (row) => {
          const totalHistories =
            typeof row.histories === 'object' && row.histories !== null
              ? (row.histories.total ?? 0)
              : typeof row.histories === 'string'
                ? (() => {
                    try {
                      return JSON.parse(row.histories)?.total ?? 0;
                    } catch {
                      return 0;
                    }
                  })()
                : 0;
          return (
            <div className="flex justify-center text-center">
              <span className="text-xs font-semibold text-slate-700 dark:text-slate-300">
                {totalHistories}
              </span>
            </div>
          );
        },
      },
      {
        label: translate('common.created_by'),
        key: 'createdBy',
        width: 'w-[19%]',
        render: (row) => {
          const ru = row.route_user || row.routeUser;
          return (
            <UserAuditCell
              user={ru?.created_by || row.createdBy}
              timestamp={ru?.created_at || row.createdAt}
              highlight={searchQuery}
            />
          );
        },
      },
      {
        label: translate('common.sync_by'),
        key: 'syncBy',
        width: 'w-[19%]',
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
      (r) =>
        (r.name || '').toLowerCase().includes(q) ||
        (r.createdBy || '').toLowerCase().includes(q) ||
        (r.updatedBy || '').toLowerCase().includes(q)
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
