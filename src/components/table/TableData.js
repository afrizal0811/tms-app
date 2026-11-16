'use client';

import { useLanguage } from '@/context/LanguageContext';
import { useMemo, useState } from 'react';
import Dropdown from '../dropdown/Dropdown';
import Spinner from '../Spinner';
import Tooltip from '../Tooltip';
import Td from './Td';
import Th from './Th';

const compareValues = (aVal, bVal) =>
  typeof aVal === 'string' && typeof bVal === 'string'
    ? aVal.localeCompare(bVal)
    : aVal < bVal
      ? -1
      : aVal > bVal
        ? 1
        : 0;

export default function TableData({
  columns = [],
  subHeaders = false,
  data = [],
  onRowClick,
  isLoading = false,
  paginate = false,
  customSort,
  externalSortConfig,
  onExternalSort,
  rowClassName,
  rowTooltip,
  tableLayout = 'table-auto',
  emptyText = null,
}) {
  const [internalSortConfig, setInternalSortConfig] = useState(null);
  const [page, setPage] = useState(1);
  const [limit, setLimit] = useState(25);
  const [prevData, setPrevData] = useState(data);

  const { t } = useLanguage();
  const sortConfig = externalSortConfig !== undefined ? externalSortConfig : internalSortConfig;

  if (data !== prevData) {
    setPrevData(data);
    setPage(1);
  }

  const handleSort = (key) => {
    let newConfig;
    if (sortConfig?.key === key) {
      if (sortConfig.direction === 'asc') {
        newConfig = { key, direction: 'desc' };
      } else {
        newConfig = null;
      }
    } else {
      newConfig = { key, direction: 'asc' };
    }

    if (onExternalSort) {
      onExternalSort(newConfig || { key: 'no', direction: 'asc' });
    } else {
      setInternalSortConfig(newConfig);
    }
  };

  const sortedData = useMemo(() => {
    let sortableItems = [...data];
    if (sortConfig !== null) {
      if (customSort) {
        sortableItems = customSort(sortableItems, sortConfig);
      } else {
        sortableItems.sort((a, b) => {
          const cmp = compareValues(a[sortConfig.key] ?? '', b[sortConfig.key] ?? '');
          return sortConfig.direction === 'asc' ? cmp : -cmp;
        });
      }
    }
    return sortableItems;
  }, [data, sortConfig, customSort]);

  const paginatedData = useMemo(() => {
    if (!paginate) return sortedData;

    const start = (page - 1) * Number(limit);
    return sortedData.slice(start, start + Number(limit));
  }, [sortedData, page, limit, paginate]);

  const totalPages = Math.ceil(sortedData.length / Number(limit));

  return (
    <div className="flex flex-col flex-1 h-full min-h-0 bg-white dark:bg-slate-900 ">
      <div className="flex-1 overflow-auto flex flex-col">
        {isLoading ? (
          <div className="flex flex-1 min-h-[300px] justify-center items-center">
            <Spinner />
          </div>
        ) : sortedData.length === 0 ? (
          <div className="p-4 text-center text-xs text-gray-500 flex flex-1 min-h-[300px] justify-center items-center">
            {emptyText || t('common.no_data')}
          </div>
        ) : (
          <table className={`w-full ${tableLayout}`} style={{ minWidth: '100%' }}>
            <thead className="bg-gray-100 dark:bg-slate-800 sticky top-0 z-10">
              <tr>
                {columns.map((col, index) => {
                  const hasSub = subHeaders && col.subColumns?.length > 0;
                  return (
                    <Th
                      key={index}
                      widthClass={col.width}
                      className={
                        col.sortable
                          ? 'cursor-pointer hover:bg-gray-200 dark:hover:bg-slate-700 select-none'
                          : ''
                      }
                      rowSpan={subHeaders && !hasSub ? 2 : 1}
                      colSpan={hasSub ? col.subColumns.length : 1}
                    >
                      <div
                        className={`flex items-center gap-1 text-center ${col.align === 'center' || hasSub ? 'justify-center' : ''}`}
                        onClick={() => col.sortable && handleSort(col.key)}
                      >
                        {col.label}
                        {col.sortable && (
                          <span className="text-[10px] text-gray-400">
                            {sortConfig?.key === col.key
                              ? sortConfig.direction === 'asc'
                                ? '▲'
                                : '▼'
                              : '↕'}
                          </span>
                        )}
                      </div>
                    </Th>
                  );
                })}
              </tr>
              {subHeaders && (
                <tr>
                  {columns.map((col) => {
                    if (col.subColumns?.length > 0) {
                      return col.subColumns.map((subCol, subIndex) => (
                        <Th
                          key={`${col.key}-sub-${subIndex}`}
                          widthClass={subCol.width}
                          className={
                            subCol.sortable
                              ? 'cursor-pointer hover:bg-gray-200 dark:hover:bg-slate-700 select-none'
                              : ''
                          }
                        >
                          <div
                            className={`flex items-center gap-1 text-center ${subCol.align === 'center' ? 'justify-center' : ''}`}
                            onClick={() => subCol.sortable && handleSort(subCol.key)}
                          >
                            {subCol.label}
                            {subCol.sortable && (
                              <span className="text-[10px] text-gray-400">
                                {sortConfig?.key === subCol.key
                                  ? sortConfig.direction === 'asc'
                                    ? '▲'
                                    : '▼'
                                  : '↕'}
                              </span>
                            )}
                          </div>
                        </Th>
                      ));
                    }
                    return null;
                  })}
                </tr>
              )}
            </thead>
            <tbody className="divide-y divide-gray-100 dark:divide-slate-800">
              {paginatedData.map((row, rowIndex) => {
                const tooltipMsg = rowTooltip ? rowTooltip(row) : null;
                const absoluteIndex = paginate ? (page - 1) * Number(limit) + rowIndex : rowIndex;

                const dataColumns = subHeaders
                  ? columns.reduce(
                      (acc, col) =>
                        col.subColumns?.length > 0 ? [...acc, ...col.subColumns] : [...acc, col],
                      []
                    )
                  : columns;

                const trContent = (
                  <tr
                    key={row._id || rowIndex}
                    onClick={() => onRowClick && onRowClick(row)}
                    className={`${onRowClick ? 'cursor-pointer' : ''} ${
                      rowClassName?.(row) || 'hover:bg-gray-50 dark:hover:bg-slate-700'
                    }`}
                  >
                    {dataColumns.map((col, colIndex) => (
                      <Td key={colIndex} className="text-[13px]!">
                        {col.render ? col.render(row, absoluteIndex) : row[col.key]}
                      </Td>
                    ))}
                  </tr>
                );

                return tooltipMsg ? (
                  <Tooltip key={`tooltip-${row._id || rowIndex}`} tooltipContent={tooltipMsg}>
                    {trContent}
                  </Tooltip>
                ) : (
                  trContent
                );
              })}
            </tbody>
          </table>
        )}
      </div>
      {paginate && (
        <div className="p-4 flex justify-between items-center bg-gray-50 dark:bg-slate-900 border-t border-gray-200 dark:border-slate-700 shrink-0">
          <div className="flex items-center gap-2 text-sm text-gray-600 dark:text-gray-300">
            <Dropdown
              value={limit}
              onChange={(val) => {
                setLimit(val);
                setPage(1);
              }}
              options={[
                { label: '25', value: 25 },
                { label: '50', value: 50 },
                { label: '100', value: 100 },
              ]}
              placement="top"
              className="w-[90px] min-w-[90px]!"
            />
            <span className="ml-4">Total data: {sortedData.length}</span>
          </div>
          {totalPages > 1 && (
            <div className="flex items-center gap-2">
              <button
                onClick={() => setPage((p) => Math.max(1, p - 1))}
                disabled={page === 1}
                className={`px-3 py-1 bg-white dark:bg-slate-800 border border-gray-300 dark:border-slate-600 rounded text-sm disabled:opacity-50 ${page !== 1 ? 'cursor-pointer' : ''}`}
              >
                &lt;
              </button>
              <span className="text-sm text-gray-600 dark:text-gray-300 px-2">
                {page} / {totalPages}
              </span>
              <button
                onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
                disabled={page === totalPages}
                className={`px-3 py-1 bg-white dark:bg-slate-800 border border-gray-300 dark:border-slate-600 rounded text-sm disabled:opacity-50 ${page !== totalPages ? 'cursor-pointer' : ''}`}
              >
                &gt;
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
