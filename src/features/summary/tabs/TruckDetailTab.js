import Tooltip from '@/components/Tooltip';
import {
  formatLongDate,
  formatMinutesToHHMM,
  getBasePlate,
  heatMap,
  isEmpty,
  isPastDate,
} from '@/lib/utils';
import { Fragment, useMemo, useState } from 'react';
import TruckDetailModal from './modals/TruckDetailModal';

const TITLE_COLOR = 'bg-[#fae2d5] dark:bg-[#3f2113]';
const DATE_COLOR = 'bg-[#dbe9f7] dark:bg-[#15233b]';
const HOLIDAY_COLOR = 'bg-[#f4cccc] dark:bg-[#451a1a]';
const TH_CLASS =
  'border border-gray-300 dark:border-slate-700 px-2 py-2 text-center text-xs font-bold text-slate-700 dark:text-slate-200';
const TD_CLASS =
  'border border-gray-300 dark:border-slate-700 px-2 py-1 text-center text-xs text-slate-700 dark:text-slate-300 whitespace-nowrap';
const TD_CLICKABLE = `${TD_CLASS} cursor-pointer hover:opacity-80 transition-opacity`;

const STICKY_TYPE = 'md:sticky md:left-0 md:z-20 md:border-r dark:md:border-r-slate-700';
const STICKY_PLATE = 'md:sticky md:left-[65px] md:z-20 md:border-r dark:md:border-r-slate-700';
const STICKY_DRIVER =
  'md:sticky md:left-[175px] md:z-20 md:border-r dark:md:border-r-slate-700 md:shadow-md';

function formatPercentage(data, maxData) {
  if (!maxData) return '0.0%';
  return ((data / maxData) * 100).toFixed(1) + '%';
}

function getCellBg(metrics, isHoliday, holidayColor, colorLegendMap) {
  if (metrics && metrics.outlets > 0) {
    if (metrics.hasManualError && metrics.hasDateDiffError) return colorLegendMap.indigo || '';
    if (metrics.hasManualError) return colorLegendMap.blue || '';
    if (metrics.hasDateDiffError) return colorLegendMap.magenta || '';
  }
  return isHoliday ? holidayColor : '';
}

function getMetricTooltip({
  key,
  metrics,
  hasSplitTask,
  hasManualError,
  hasManualTooltip,
  isTotalOutlet,
  isIndonesian,
  translate,
}) {
  const isWeightKey = key === 'weight';
  const isVolumeKey = key === 'volume';
  const metricUnit = isWeightKey ? 'Kg' : isVolumeKey ? 'Cbm' : '';

  if (hasSplitTask && isTotalOutlet) {
    return translate('summary.tabs.truck_detail.tooltip.split_task');
  }

  if (
    hasManualError &&
    (!isEmpty(metrics.weight) || !isEmpty(metrics.volume)) &&
    hasManualTooltip
  ) {
    if (isTotalOutlet) return null;
    const metricLabel = translate(isWeightKey ? 'common.weight' : 'common.volume');
    const routingText = isIndonesian
      ? `${metricLabel} ${translate('common.routing')} `
      : `${translate('common.routing')} ${metricLabel}`;
    const routingPctText = `${translate('common.percentage')} ${routingText}`;

    const totalVal = isWeightKey ? metrics.weight : isVolumeKey ? metrics.volume : 0;
    const realVal = isWeightKey ? metrics.realWeight : isVolumeKey ? metrics.realVolume : 0;
    const maxVal = isWeightKey ? metrics.maxWeight : isVolumeKey ? metrics.maxVolume : 0;
    const anyHasRouting = metrics.taskList?.some((t) => !isEmpty(t.roSequence));
    const realPct = formatPercentage(realVal, maxVal) || 0;

    return (
      <span>
        <div>
          {!anyHasRouting
            ? translate('summary.tabs.truck_detail.tooltip.no_routing')
            : translate('summary.tabs.truck_detail.tooltip.inaccurate_data')}
        </div>
        {(isWeightKey || isVolumeKey) && (
          <div className="mt-1">
            <div>
              - Total: {totalVal.toFixed(2)} {metricUnit}
            </div>
            <div>
              - Total {routingText}: {realVal.toFixed(2)} {metricUnit}
            </div>
            <div>
              - {routingPctText}: {realPct}
            </div>
          </div>
        )}
      </span>
    );
  }

  if (isWeightKey || isVolumeKey) {
    const val = isWeightKey ? metrics.weight : metrics.volume;
    return (
      <span>
        {(val ?? 0).toFixed(2)} {metricUnit}
      </span>
    );
  }

  return null;
}

export default function TruckDetailTab({ data, translate, localeCode, isIndonesian, colorLegend }) {
  const { driverEmails = [], driverMap = {}, dateKeys = [], dataMatrix = {} } = data || {};
  const [modalData, setModalData] = useState(null);

  const handleCellClick = (metrics, driverName, dateStr) => {
    if (metrics?.taskList?.length > 0) {
      setModalData({ driverName, dateStr, tasks: metrics.taskList });
    }
  };

  const closeModal = () => setModalData(null);

  const colorLegendMap = useMemo(() => {
    const map = {};
    if (colorLegend) {
      for (const item of colorLegend) {
        if (item.name) map[item.name] = item.colors;
      }
    }
    return map;
  }, [colorLegend]);

  const holidayStatusMap = useMemo(() => {
    const map = {};
    if (!dateKeys || !dataMatrix) return map;

    for (const d of dateKeys) {
      const dateStr = d.str;
      if (!dateStr) continue;
      const [y, m, day] = dateStr.split('-').map(Number);
      const isSunday = new Date(y, m - 1, day).getDay() === 0;

      const dayIsEmpty =
        !dataMatrix[dateStr] ||
        driverEmails.every((email) => {
          const metrics = dataMatrix[dateStr][email];
          return !metrics || (metrics.outlets || 0) === 0;
        });

      const isDynamic = !isSunday && isPastDate(dateStr) && dayIsEmpty;

      map[dateStr] = {
        isHoliday: isSunday || isDynamic,
        isDynamic,
        isSunday,
        dayIsEmpty,
      };
    }
    return map;
  }, [dateKeys, dataMatrix, driverEmails]);

  const displayData = useMemo(
    () => [
      {
        key: 'weight',
        border: true,
        getValue: (m) => {
          if (!m || m.maxWeight <= 0) return '-';
          const val = formatPercentage(m.weight, m.maxWeight);
          return (
            <span className={m.hasManualError ? 'text-red-200! dark:text-red-300!' : ''}>
              {val}
            </span>
          );
        },
        text: translate('common.weight'),
        hasManualTooltip: true,
      },
      {
        key: 'volume',
        getValue: (m) => {
          if (!m || m.maxVolume <= 0) return '-';
          const val = formatPercentage(m.volume, m.maxVolume);
          return (
            <span className={m.hasManualError ? 'text-red-200! dark:text-red-300!' : ''}>
              {val}
            </span>
          );
        },
        text: translate('common.volume'),
        hasManualTooltip: true,
      },
      {
        key: 'distance',
        getValue: (m) => {
          const val = m?.dist ? m.dist : '-';
          return (
            <span className={m?.hasManualError ? 'text-red-200! dark:text-red-300!' : ''}>
              {val}
            </span>
          );
        },
        text: translate('common.distance'),
        hasManualTooltip: true,
      },
      {
        key: 'total_outlet',
        getValue: (m) => {
          if (!m?.outlets) return 0;
          const hasSplitTask = m.taskList?.some((item) => item.isSplitTask);
          return (
            <span className="inline-flex items-center justify-center gap-1 leading-none">
              <span className="inline-block h-2.5 w-2.5 shrink-0" />
              <span>{m.outlets}</span>
              <span
                className={`inline-block h-2.5 w-2.5 shrink-0 rounded-full bg-orange-400 dark:bg-orange-600 ${
                  hasSplitTask ? '' : 'opacity-0'
                }`}
              />
            </span>
          );
        },
        text: translate('summary.tabs.truck_detail.total_outlet'),
        hasManualTooltip: true,
      },
      {
        key: 'total_delivery',
        getValue: (m) => m?.delivered ?? 0,
        text: translate('summary.tabs.truck_detail.total_delivery'),
        hasManualTooltip: false,
      },
      {
        key: 'ship_duration',
        getValue: (m) => {
          const val = m?.duration ? formatMinutesToHHMM(m.duration) : '-';
          return (
            <span className={m?.hasManualError ? 'text-red-200! dark:text-red-300!' : ''}>
              {val}
            </span>
          );
        },
        text: translate('summary.tabs.truck_detail.ship_duration'),
        hasManualTooltip: true,
      },
      {
        key: 'delivered',
        getValue: (m) => (m?.outlets > 0 ? formatPercentage(m.delivered, m.outlets) : 0),
        getStyle: (m) => {
          if (!m || m.outlets <= 0) return {};
          const hex = heatMap(formatPercentage(m.delivered, m.outlets));
          return hex ? { backgroundColor: `#${hex}`, color: '#334155' } : {};
        },
        text: translate('summary.tabs.truck_detail.delivered'),
        hasManualTooltip: false,
      },
    ],
    [translate]
  );

  return (
    <div className="w-full h-full flex flex-col relative p-0">
      <TruckDetailModal
        data={modalData}
        isOpen={Boolean(modalData)}
        localeCode={localeCode}
        onClose={closeModal}
        translate={translate}
        isIndonesian={isIndonesian}
      />
      <div className="overflow-auto flex-1">
        <table className="border-collapse border-0 text-sm whitespace-nowrap">
          <thead className="sticky top-0 z-30 bg-gray-100 dark:bg-slate-800">
            <tr>
              <th
                rowSpan="2"
                className={`${TH_CLASS} min-w-[65px] w-[65px] md:sticky md:left-0 md:z-40 ${TITLE_COLOR}`}
              >
                {translate('common.storage_type')}
              </th>
              <th
                rowSpan="2"
                className={`${TH_CLASS} min-w-[110px] w-[110px] md:sticky md:left-[65px] md:z-40 ${TITLE_COLOR}`}
              >
                {translate('common.license_number')}
              </th>
              <th
                rowSpan="2"
                className={`${TH_CLASS} min-w-[200px] md:sticky md:left-[175px] md:z-40 ${TITLE_COLOR} md:border-r-2 md:border-slate-400 dark:md:border-slate-600`}
              >
                {translate('common.driver')}
              </th>
              {dateKeys.map((d, i) => {
                const holidayInfo = holidayStatusMap[d.str] || {};
                const headerColor = holidayInfo.isHoliday ? HOLIDAY_COLOR : DATE_COLOR;
                const dateText = formatLongDate(d.str, localeCode);
                const headerContent = holidayInfo.isDynamic ? (
                  <Tooltip tooltipContent={translate('summary.tabs.task_summary.caution')}>
                    <span className="cursor-help border-b-2 border-dotted border-red-900 dark:border-red-300 pb-0.5">
                      {dateText}
                    </span>
                  </Tooltip>
                ) : (
                  <span>{dateText}</span>
                );

                return (
                  <th
                    key={i}
                    colSpan="7"
                    className={`${TH_CLASS} border-l-2 border-l-gray-400 dark:border-l-slate-600 ${headerColor}`}
                  >
                    {headerContent}
                  </th>
                );
              })}
            </tr>
            <tr>
              {dateKeys.map((d, i) => {
                const holidayInfo = holidayStatusMap[d.str] || {};
                const metricColor = holidayInfo.isHoliday ? HOLIDAY_COLOR : TITLE_COLOR;

                const hasFallback = driverEmails.some(
                  (email) => dataMatrix[d.str]?.[email]?.isDistFallback
                );

                return (
                  <Fragment key={`${d.day}-${i}-header`}>
                    {displayData.map(({ key, border, text }) => {
                      const isTotalDelivery = key === 'total_delivery';
                      const isDistance = key === 'distance';

                      let tooltipText = '';
                      let showDotted = false;

                      if (!holidayInfo.isHoliday) {
                        if (isTotalDelivery) {
                          tooltipText = translate('summary.tabs.truck_detail.tooltip.pct_info');
                          showDotted = true;
                        } else if (isDistance && hasFallback) {
                          tooltipText = translate(
                            'summary.tabs.truck_detail.tooltip.no_route_dist'
                          );
                          showDotted = true;
                        }
                      }

                      return (
                        <Tooltip tooltipContent={tooltipText} key={key}>
                          <th
                            className={`${TH_CLASS} ${metricColor} ${
                              border ? 'border-l-2 border-l-gray-400 dark:border-l-slate-600' : ''
                            }`}
                          >
                            <span
                              className={
                                showDotted ? 'cursor-help border-b-2 border-dotted pb-0.5' : ''
                              }
                            >
                              {text}
                            </span>
                          </th>
                        </Tooltip>
                      );
                    })}
                  </Fragment>
                );
              })}
            </tr>
          </thead>
          <tbody className="bg-white dark:bg-slate-800">
            {driverEmails.map((email, rowIndex) => {
              const driver = driverMap[email] || {};
              const driverMaxW = driver.maxWeight || 0;
              const driverMaxV = (driver.maxVolume != null ? Number(driver.maxVolume) : 0).toFixed(
                2
              );
              const capacityTooltip =
                driverMaxW > 0 || Number(driverMaxV) > 0 ? (
                  <span>
                    <div>
                      - {translate('common.weight_max')}: {driverMaxW} Kg
                    </div>{' '}
                    <div>
                      - {translate('common.volume_max')}: {driverMaxV} Cbm
                    </div>
                  </span>
                ) : (
                  translate('common.no_data')
                );

              return (
                <tr key={email} className="hover:bg-gray-50 dark:hover:bg-slate-800">
                  <td className={`${TD_CLASS} ${STICKY_TYPE} bg-white dark:bg-slate-800`}>
                    {driver.type}
                  </td>
                  <td className={`${TD_CLASS} ${STICKY_PLATE} bg-white dark:bg-slate-800`}>
                    <Tooltip tooltipContent={capacityTooltip}>
                      <span className="cursor-help border-b-2 border-dotted pb-0.4">
                        {getBasePlate(driver.plat)}
                      </span>
                    </Tooltip>
                  </td>
                  <td
                    className={`${TD_CLASS} ${STICKY_DRIVER} bg-white dark:bg-slate-800 text-left md:border-r-2 md:border-slate-400 dark:md:border-slate-600`}
                  >
                    {driver.name}
                  </td>

                  {dateKeys.map((d, i) => {
                    const metrics = dataMatrix[d.str]?.[email];
                    const outletData = metrics?.outlets || 0;
                    const onClick = () => handleCellClick(metrics, driver.name, d.str);

                    const holidayInfo = holidayStatusMap[d.str] || {};
                    const shouldMergeHoliday = holidayInfo.isHoliday && holidayInfo.dayIsEmpty;

                    const cellBg = getCellBg(
                      metrics,
                      holidayInfo.isHoliday,
                      HOLIDAY_COLOR,
                      colorLegendMap
                    );
                    const emptyBg = holidayInfo.isHoliday
                      ? HOLIDAY_COLOR
                      : 'bg-gray-50 dark:bg-slate-800';

                    if (shouldMergeHoliday) {
                      if (rowIndex === 0) {
                        return (
                          <td
                            key={`${d.str}-merged`}
                            rowSpan={driverEmails.length}
                            colSpan={displayData.length}
                            className={`border border-gray-300 dark:border-slate-700 px-2 py-1 text-center font-bold align-middle ${HOLIDAY_COLOR} text-red-900 dark:text-red-300 border-l-2 border-l-gray-400 dark:border-l-slate-600`}
                          >
                            {holidayInfo.isSunday
                              ? translate('common.holiday_sunday')
                              : translate('common.holiday')}
                          </td>
                        );
                      }
                      return null;
                    }

                    if (!metrics || outletData === 0) {
                      return (
                        <Fragment key={`${d.day}-${i}-empty`}>
                          {displayData.map(({ key, border }) => (
                            <td
                              key={key}
                              className={`${TD_CLASS} ${emptyBg} ${
                                border ? 'border-l-2 border-l-gray-400 dark:border-l-slate-600' : ''
                              }`}
                            />
                          ))}
                        </Fragment>
                      );
                    }

                    return (
                      <Fragment key={`${d.day}-${i}-data`}>
                        {displayData.map(
                          ({ key, border, getValue, getStyle, hasManualTooltip }) => {
                            const hasHeatmap = Boolean(getStyle && metrics.outlets > 0);
                            const heatmapStyle = hasHeatmap ? getStyle(metrics) : {};
                            const appliedBgClass = hasHeatmap ? '' : cellBg;
                            const isTotalOutlet = key === 'total_outlet';
                            const hasSplitTask = metrics.taskList?.some((item) => item.isSplitTask);
                            const hasManualError = Boolean(
                              metrics.hasManualError &&
                                (!isEmpty(metrics.weight) || !isEmpty(metrics.volume)) &&
                                hasManualTooltip
                            );

                            const tooltipData = getMetricTooltip({
                              key,
                              metrics,
                              hasSplitTask,
                              hasManualError,
                              hasManualTooltip,
                              isTotalOutlet,
                              isIndonesian,
                              translate,
                            });

                            return (
                              <Tooltip key={key} tooltipContent={tooltipData}>
                                <td
                                  onClick={onClick}
                                  className={`${TD_CLICKABLE} ${appliedBgClass} ${border ? 'border-l-2 border-l-gray-400 dark:border-l-slate-600' : ''} `}
                                  style={heatmapStyle}
                                >
                                  {getValue(metrics)}
                                </td>
                              </Tooltip>
                            );
                          }
                        )}
                      </Fragment>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
