'use client';

import HighlightText from '@/components/HighlightText';
import TableData from '@/components/table/TableData';
import Tooltip from '@/components/Tooltip';
import { isEmpty } from '@/lib/utils';
import { useMemo } from 'react';

const getRowClassName = (v) => {
  if (v.isIncomplete) {
    return 'bg-red-50 dark:bg-red-500/10 hover:bg-red-100/80 dark:hover:bg-red-500/15 transition-colors cursor-help';
  }
  if (v.isDuplicateDriver) {
    return 'bg-yellow-50 dark:bg-yellow-500/10 hover:bg-yellow-100/80 dark:hover:bg-yellow-500/15 transition-colors cursor-help';
  }
  return 'hover:bg-gray-50 dark:hover:bg-slate-700/10 transition-colors';
};

export default function TemplateTab({ paginatedData, searchQuery, t }) {
  const dataWithNo = useMemo(() => {
    return paginatedData.map((item, index) => ({
      ...item,
      no: index + 1,
    }));
  }, [paginatedData]);

  const getRowTooltip = (row) => {
    const tooltips = [];
    if (row.isIncomplete) tooltips.push(t('vehicle.tabs.incomplete_data'));
    if (row.isDuplicateDriver) tooltips.push(t('vehicle.tabs.duplicate_driver'));
    const tooltip = tooltips.join(', ');
    return tooltip.charAt(0).toUpperCase() + tooltip.slice(1).toLowerCase();
  };

  const columns = [
    {
      key: 'no',
      width: 'w-[5%]',
      sortable: false,
      align: 'center',
      label: '#',
      render: (row) => <div className="text-center w-full">{row.no}</div>,
    },
    {
      key: 'plat',
      width: 'w-[10%]',
      sortable: false,
      align: 'left',
      label: t('common.name'),
      render: (row) => (
        <div className="text-left w-full">
          <HighlightText text={row.plat} highlight={searchQuery} />
        </div>
      ),
    },
    {
      key: 'email',
      width: 'w-[10%]',
      sortable: false,
      label: t('vehicle.tabs.template.assignee'),
      render: (row) => (
        <div className={`w-full ${row.email ? 'text-left' : 'text-center'}`}>
          <HighlightText text={row.email || '-'} highlight={searchQuery} />
        </div>
      ),
    },
    {
      key: 'startTime',
      width: 'w-[10%]',
      sortable: false,
      align: 'center',
      label: t('common.start_time'),
      render: (row) => (
        <div className="text-center w-full">{row.workingTime?.startTime || null}</div>
      ),
    },
    {
      key: 'endTime',
      width: 'w-[10%]',
      sortable: false,
      align: 'center',
      label: t('common.finish_time'),
      render: (row) => <div className="text-center w-full">{row.workingTime?.endTime || null}</div>,
    },
    {
      key: 'startBreakTime',
      width: 'w-[10%]',
      sortable: false,
      align: 'center',
      label: t('vehicle.tabs.template.break_start_time'),
      render: (row) => (
        <div className="text-center w-full">{row.breakTime?.startBreakTime || null}</div>
      ),
    },
    {
      key: 'endBreakTime',
      width: 'w-[10%]',
      sortable: false,
      align: 'center',
      label: t('vehicle.tabs.template.break_end_time'),
      render: (row) => (
        <div className="text-center w-full">{row.breakTime?.endBreakTime || null}</div>
      ),
    },
    {
      key: 'multiday',
      width: 'w-[5%]',
      sortable: false,
      align: 'center',
      label: t('vehicle.tabs.template.multiday'),
      render: (row) => <div className="text-center w-full">{row.workingTime?.multiday || 0}</div>,
    },
    {
      key: 'speed',
      width: 'w-[5%]',
      sortable: false,
      align: 'center',
      label: t('common.speed'),
      render: (row) => <div className="text-center w-full">{row.speed}</div>,
    },
    {
      key: 'costFactor',
      width: 'w-[5%]',
      sortable: false,
      align: 'center',
      label: t('vehicle.tabs.template.cost_factor'),
      render: (row) => <div className="text-center w-full">{row.costFactor}</div>,
    },
    {
      key: 'tags',
      width: 'w-[20%]',
      sortable: false,
      align: 'center',
      label: t('vehicle.tabs.template.vehicle_tags'),
      render: (row) => {
        const tags = row.parsedTags || [];
        if (isEmpty(tags)) return <div className="text-center w-full cursor-help">-</div>;
        const firstTag = tags[0];
        const remainingTags = tags.slice(1);
        const remainingCount = remainingTags.length;
        if (remainingCount === 0) return <div className=" w-full cursor-help">{firstTag}</div>;
        return (
          <div className="w-full cursor-help">
            <Tooltip tooltipContent={remainingTags.join('\n')}>
              <span>
                {firstTag}; (+{remainingCount} {t('common.others')})
              </span>
            </Tooltip>
          </div>
        );
      },
    },
    {
      key: 'oddEven',
      width: 'w-[5%]',
      sortable: false,
      label: t('vehicle.tabs.template.odd_even'),
      render: (row) => <div className="text-center w-full">{row.oddEven}</div>,
    },
    {
      key: 'weightMin',
      width: 'w-[5%]',
      sortable: false,
      label: t('vehicle.tabs.template.weight_min'),
      render: (row) => <div className="text-center w-full">{row.minWeight || 0}</div>,
    },
    {
      key: 'weightMax',
      width: 'w-[5%]',
      sortable: false,
      label: t('common.weight_max'),
      render: (row) => <div className="text-center w-full">{row.maxWeight || 0}</div>,
    },
    {
      key: 'volumeMin',
      width: 'w-[5%]',
      sortable: false,
      label: t('vehicle.tabs.template.volume_min'),
      render: (row) => <div className="text-center w-full">{row.minVolume || 0}</div>,
    },
    {
      key: 'volumeMax',
      width: 'w-[5%]',
      sortable: false,
      label: t('common.volume_max'),
      render: (row) => <div className="text-center w-full">{row.maxVolume.toFixed(2)}</div>,
    },
  ];

  return (
    <div className="overflow-hidden flex-1 h-full rounded-b-lg">
      <TableData
        columns={columns}
        data={dataWithNo}
        rowClassName={getRowClassName}
        rowTooltip={getRowTooltip}
      />
    </div>
  );
}
