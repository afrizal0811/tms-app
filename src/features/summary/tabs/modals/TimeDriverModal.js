'use client';

import Map from '@/components/Map';
import Modal from '@/components/modal/Modal';
import TableData from '@/components/table/TableData';
import Tooltip from '@/components/Tooltip';
import { formatDateUniversal, isEmpty } from '@/lib/utils';
import { useMemo } from 'react';

const parseDurationToMinutes = (str) => {
  if (!str) return 0;
  const [hours, minutes] = str.split(':').map(Number);
  return hours * 60 + (minutes || 0);
};

export default function TimeDriverModal({ isOpen, onClose, data, translate }) {
  const hasMultipleData = data && data.entries && data.entries.length > 1;

  const tableEntries = useMemo(() => {
    if (!data || !data.entries) return [];
    return data.entries.map((entry, index) => ({
      ...entry,
      no: index + 1,
    }));
  }, [data]);

  if (!data) return null;

  const { driverName, dateStr, entries, activeHubLocation } = data;

  const totalMinutes = entries.reduce(
    (acc, curr) => acc + parseDurationToMinutes(curr.durationDisplay),
    0
  );
  const totalHours = Math.floor(totalMinutes / 60);
  const remainingMinutes = totalMinutes % 60;
  const totalDurationFormatted = `${String(totalHours).padStart(2, '0')}:${String(remainingMinutes).padStart(2, '0')}`;

  const columns = [
    {
      key: 'no',
      width: 'w-[10%]',
      sortable: false,
      label: '#',
      align: 'center',
      render: (row) => (
        <div className="text-center text-gray-500 dark:text-slate-400 font-medium w-full">
          {row.no}
        </div>
      ),
    },
    {
      key: 'start',
      width: 'w-[30%]',
      sortable: false,
      align: 'center',
      label: translate('common.start_time'),
      render: (row) => {
        const hasOutStart = row.isStartOutRadius;
        const content = (
          <div
            className={`w-full text-center px-2 py-1 rounded ${hasOutStart ? 'bg-red-100 dark:bg-red-900/40 text-slate-700 dark:text-slate-300' : 'text-slate-700 dark:text-slate-300'}`}
          >
            {row.startDisplay}
          </div>
        );
        return hasOutStart ? (
          <Tooltip tooltipContent={translate('summary.tabs.time_driver.tooltip.out_start')}>
            {content}
          </Tooltip>
        ) : (
          content
        );
      },
    },
    {
      key: 'finish',
      width: 'w-[30%]',
      sortable: false,
      align: 'center',
      label: translate('common.finish_time'),
      render: (row) => {
        const hasOutFinish = row.isFinishOutRadius;
        const diffDay = row.dayDiff;
        const hasDiffDay = !isEmpty(diffDay);
        const diffDayTooltip = hasDiffDay
          ? `${hasOutFinish ? '\n- ' : ''}${translate('summary.tabs.time_driver.tooltip.diff_day', { days: diffDay })} `
          : '';
        const outFinishTooltip = `${hasDiffDay ? '- ' : ''}${translate('summary.tabs.time_driver.tooltip.out_finish')} ${diffDayTooltip}`;

        const content = (
          <div
            className={`w-full text-center px-2 py-1 rounded ${hasOutFinish ? 'bg-red-100 dark:bg-red-900/40 text-slate-700 dark:text-slate-300' : 'text-slate-700 dark:text-slate-300'}`}
          >
            {row.finishDisplay}
            {row.dayDiff > 0 && (
              <span className="text-red-600 dark:text-red-400 text-xs ml-1 font-bold">
                (+{row.dayDiff})
              </span>
            )}
          </div>
        );

        return hasOutFinish || hasDiffDay ? (
          <Tooltip tooltipContent={hasOutFinish ? outFinishTooltip : diffDayTooltip}>
            {content}
          </Tooltip>
        ) : (
          content
        );
      },
    },
    {
      key: 'duration',
      width: 'w-[30%]',
      sortable: false,
      align: 'center',
      label: translate('common.duration'),
      render: (row) => (
        <div className="text-center font-medium text-slate-700 dark:text-slate-200 w-full">
          {row.durationDisplay}
        </div>
      ),
    },
  ];

  const bounds = [];
  if (activeHubLocation) bounds.push([activeHubLocation.lat, activeHubLocation.lng]);
  entries.forEach((entry) => {
    if (entry.startLat && entry.startLon) bounds.push([entry.startLat, entry.startLon]);
    if (entry.finishLat && entry.finishLon) bounds.push([entry.finishLat, entry.finishLon]);
  });

  return (
    <Modal isOpen={isOpen} onClose={onClose} title={driverName} subtitle={dateStr}>
      <div className="flex flex-col gap-4">
        <div className="h-[450px] w-full rounded-lg overflow-hidden border border-slate-300 dark:border-slate-700 relative z-0">
          <Map bounds={bounds}>
            {({ Marker, Tooltip: LeafletTooltip, Circle }, L, icons) => (
              <>
                {activeHubLocation && (
                  <>
                    <Circle
                      center={[activeHubLocation.lat, activeHubLocation.lng]}
                      radius={500}
                      pathOptions={{
                        color: '#22c55e',
                        fillColor: '#22c55e',
                        fillOpacity: 0.2,
                        weight: 1,
                      }}
                    />
                    <Marker
                      position={[activeHubLocation.lat, activeHubLocation.lng]}
                      icon={icons.circle('HUB', 'bg-green-500', 'text-[10px]')}
                    />
                  </>
                )}
                {entries.map((entry, idx) => {
                  const sText = entries.length > 1 ? `S${idx + 1}` : 'S';
                  const fText = entries.length > 1 ? `F${idx + 1}` : 'F';
                  const hasDayDiff = entry.dayDiff > 0;
                  return (
                    <div key={idx}>
                      {entry.startLat && entry.startLon && (
                        <Marker
                          position={[entry.startLat, entry.startLon]}
                          icon={icons.circle(
                            sText,
                            entry.isStartOutRadius ? 'bg-red-500' : 'bg-sky-500'
                          )}
                        >
                          <LeafletTooltip direction="top" offset={[0, -10]}>
                            <div
                              dangerouslySetInnerHTML={{
                                __html: `<b>${translate('common.start_time')}</b><br>${formatDateUniversal(entry.startTime, 'DD/MM/YYYY HH:mm')}`,
                              }}
                            />
                          </LeafletTooltip>
                        </Marker>
                      )}
                      {entry.finishLat && entry.finishLon && (
                        <Marker
                          position={[entry.finishLat, entry.finishLon]}
                          icon={icons.circle(
                            fText,
                            entry.isFinishOutRadius ? 'bg-red-500' : 'bg-sky-500',
                            'text-xs',
                            hasDayDiff ? 'border-red-500' : 'border-white'
                          )}
                        >
                          <LeafletTooltip direction="top" offset={[0, -10]}>
                            <div
                              dangerouslySetInnerHTML={{
                                __html: `<b>${translate('common.finish_time')}</b><br>${formatDateUniversal(entry.finishTime, 'DD/MM/YYYY HH:mm')} <span class="text-red-500">${hasDayDiff ? `(+${entry.dayDiff})` : ''}</span>`,
                              }}
                            />
                          </LeafletTooltip>
                        </Marker>
                      )}
                    </div>
                  );
                })}
              </>
            )}
          </Map>
        </div>
        {hasMultipleData && (
          <>
            <div className="overflow-hidden border border-gray-200 dark:border-slate-700 rounded-lg flex flex-col">
              <TableData columns={columns} data={tableEntries} paginate={false} />
              <div className="flex bg-gray-50 dark:bg-slate-800/50 font-bold border-t-2 border-gray-200 dark:border-slate-700">
                <div className="w-[70%] px-4 py-3 text-center text-gray-600 dark:text-slate-400 uppercase text-[10px] tracking-wider">
                  Total
                </div>
                <div className="w-[30%] text-sm px-4 py-3 text-center text-slate-800 dark:text-slate-200">
                  {totalDurationFormatted}
                </div>
              </div>
            </div>
            <div className="mt-1 mb-2 text-xs text-slate-500 dark:text-slate-400 italic">
              {translate('summary.tabs.time_driver.modal.footer_note')}
            </div>
          </>
        )}
      </div>
    </Modal>
  );
}
