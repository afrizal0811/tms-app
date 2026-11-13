import Tooltip from '@/components/Tooltip';
import { formatDateUniversal } from '@/lib/utils';
import { Fragment } from 'react';

export default function TruckUsageTable({
  dateMap,
  dateKeys,
  activeTypes,
  hubMasterData,
  isPercentage = false,
  translate,
  onCellClick,
}) {
  const bgHeader =
    'bg-[#d9d2e9] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-violet-900/30 dark:to-violet-900/30';
  const bgDry =
    'bg-[#fae2d5] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-orange-900/30 dark:to-orange-900/30';
  const bgDryTotal =
    'bg-[#f9cb9c] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-orange-900/50 dark:to-orange-900/50';
  const bgFrozen =
    'bg-[#dbe9f7] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-blue-900/30 dark:to-blue-900/30';
  const bgFrozenTotal =
    'bg-[#c9daf8] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-blue-900/50 dark:to-blue-900/50';
  const bgOTV =
    'bg-[#d9f2d0] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-green-900/30 dark:to-green-900/30';
  const bgHoliday =
    'bg-[#ffc7ce] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-red-900/40 dark:to-red-900/40';

  const bgAlert = 'bg-[#FF0000] dark:bg-red-700 text-white font-bold';
  const bgPctLow =
    'bg-[#f4cccc] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-red-900/40 dark:to-red-900/40';
  const bgPctMid =
    'bg-[#f1c232] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-yellow-900/50 dark:to-yellow-900/50';
  const bgPctHigh =
    'bg-[#b7e1cd] dark:bg-slate-800 dark:bg-gradient-to-t dark:from-emerald-900/50 dark:to-emerald-900/50';
  const bgPctOver = 'bg-[#ff0000] dark:bg-red-700 text-white font-bold';

  const thClass =
    'border border-l-0 border-gray-300 dark:border-slate-600 px-2 py-2 text-center min-w-[60px] text-xs font-bold text-slate-700 dark:text-slate-200';
  const tdClass =
    'border-b border-gray-200 dark:border-slate-700 px-2 py-1 text-center text-xs text-slate-700 dark:text-slate-300';
  const thickBorderClass = 'border-r-[3px] border-r-slate-400 dark:border-r-slate-500';

  const todayStr = formatDateUniversal(new Date());

  const getCellClass = (isLastCol) => {
    return isLastCol ? `${tdClass} ${thickBorderClass}` : tdClass;
  };

  const getMasterVal = (cat, type) => {
    let val = 0;
    if (cat === 'DryTotal') val = hubMasterData?.Dry?.Total;
    else if (cat === 'FrozenTotal') val = hubMasterData?.Frozen?.Total;
    else if (cat === 'Dry') val = hubMasterData?.Dry?.[type];
    else if (cat === 'Frozen') val = hubMasterData?.Frozen?.[type];
    else if (cat === 'OTV')
      val = (hubMasterData?.Dry?.Total || 0) + (hubMasterData?.Frozen?.Total || 0);

    return val > 0 ? val : null;
  };

  const formatValue = (val, masterTotal) => {
    const numVal = parseInt(val || 0);
    const numMaster = parseInt(masterTotal || 0);
    if (numVal === 0) return null;
    if (isPercentage) {
      if (numMaster === 0) return null;
      const pct = (numVal / numMaster) * 100;
      return `${Math.round(pct)}%`;
    }
    return numVal;
  };

  const getDataClass = (
    baseBgClass,
    isHoliday,
    value,
    masterTotal,
    isDetailRow = false,
    cellType = '',
    otherVal = 0
  ) => {
    const valNum = parseInt(value || 0);
    const maxNum = parseInt(masterTotal || 0);

    if (isPercentage) {
      if (isHoliday) return bgHoliday;
      if (maxNum > 0 && valNum > 0) {
        const pct = (valNum / maxNum) * 100;
        if (pct > 100) return bgPctOver;
        if (pct >= 75) return bgPctHigh;
        if (pct >= 50) return bgPctMid;
        return bgPctLow;
      }
      return baseBgClass;
    }

    let isOverLimit = false;
    if (isDetailRow) {
      if (cellType === 'TMS' && valNum > maxNum) isOverLimit = true;
      if (cellType === 'MANUAL' && valNum + parseInt(otherVal || 0) > maxNum) isOverLimit = true;
      if (cellType === 'TVU' && valNum > maxNum) isOverLimit = true;
    }

    if (isOverLimit) return bgAlert;
    if (isHoliday) return bgHoliday;

    return baseBgClass;
  };

  const getRowValues = (d, category, label2) => {
    let tmsRaw = 0;
    let manualRaw = 0;
    let manualDesc = '';
    let manualId = null;
    let tmsDetails = [];

    if (category === 'Dry' || category === 'Frozen') {
      tmsRaw = dateMap[d.str][category][label2] || 0;
      tmsDetails = dateMap[d.str][category][`${label2}_details`] || [];

      const manualObj = dateMap[d.str][`${category}Manual`]?.[label2];
      if (manualObj) {
        manualRaw = manualObj.count || 0;
        manualDesc = manualObj.desc || '';
        manualId = manualObj.id || null;
      }
    } else if (category === 'DryTotal' || category === 'FrozenTotal') {
      tmsRaw = dateMap[d.str][category] || 0;
      manualRaw = dateMap[d.str][`${category}Manual`] || 0;

      const actualCat = category.replace('Total', '');
      activeTypes.forEach((vt) => {
        const details = dateMap[d.str][actualCat][`${vt}_details`] || [];
        tmsDetails = tmsDetails.concat(details);
      });
    } else if (category === 'OTV') {
      tmsRaw = dateMap[d.str].OTV || 0;
      manualRaw = dateMap[d.str].OTVManual || 0;

      ['Dry', 'Frozen'].forEach((cat) => {
        activeTypes.forEach((vt) => {
          const details = dateMap[d.str][cat][`${vt}_details`] || [];
          tmsDetails = tmsDetails.concat(details);
        });
      });
    }

    const tmsDisp = tmsRaw === 0 ? null : tmsRaw;
    const manualDisp = manualRaw === 0 ? null : manualRaw;
    const totalSum = tmsRaw + manualRaw;
    const tvuDisp = totalSum > 0 ? totalSum : null;

    return { tmsDisp, manualDisp, tvuDisp, tmsRaw, manualRaw, manualDesc, manualId, tmsDetails };
  };

  const getVisibleTypes = (cat) =>
    activeTypes.filter((type) => (hubMasterData?.[cat]?.[type] || 0) > 0);
  const dryTypes = getVisibleTypes('Dry');
  const frozenTypes = getVisibleTypes('Frozen');
  const totalRows = dryTypes.length + frozenTypes.length + 5;

  const renderMasterCell = (masterTotal, isClickable, storage, type, bgClass, isBoldCell = false) => (
    <td
      onClick={() => {
        if (isClickable && onCellClick) {
          onCellClick({ isMaster: true, storage, type, masterTotal });
        }
      }}
      className={`${tdClass} w-[60px] min-w-[60px] ${isBoldCell ? 'font-bold ' : ''}sticky left-[250px] z-30 ${thickBorderClass} ${bgClass} ${isClickable ? 'cursor-pointer hover:brightness-90 dark:hover:hover:brightness-120 transition-colors' : ''}`}
    >
      {isClickable ? (
        <span className="inline-block border-b-2 border-dotted border-red-700 dark:border-red-400 px-1 pb-0.5">
          {masterTotal}
        </span>
      ) : (
        masterTotal
      )}
    </td>
  );

  const renderSectionRows = (cat, bgClass, types) => {
    return types.map((type, idx) => {
      const masterTotal = getMasterVal(cat, type);
      const isClickableMaster = !isPercentage && masterTotal > 0;
      const isAbsoluteFirstRow = (cat === 'Dry' || (dryTypes.length === 0 && cat === 'Frozen')) && idx === 0;
      return (
        <tr key={`${cat}-${type}`}>
          {idx === 0 && (
            <td
              rowSpan={types.length}
              className={`${tdClass} w-[100px] min-w-[100px] max-w-[100px] font-bold align-middle sticky left-0 z-30 border-r border-gray-300 dark:border-slate-600 ${bgClass}`}
            >
              {cat}
            </td>
          )}
          <td
            className={`${tdClass} w-[150px] min-w-[150px] max-w-[150px] text-left sticky left-[100px] z-30 border-r border-gray-300 dark:border-slate-600 ${bgClass}`}
          >
            {type}
          </td>
          {renderMasterCell(masterTotal, isClickableMaster, cat, type, bgClass, true)}
          {dateKeys.map((d, i) => {
            const {
              tmsDisp,
              manualDisp,
              tvuDisp,
              tmsRaw,
              manualRaw,
              manualDesc,
              manualId,
              tmsDetails,
            } = getRowValues(d, cat, type);
            const isFuture = d.str > todayStr;
            const isHoliday = d.isSunday || d.isDynamicHoliday;

            if (isHoliday) {
              if (isAbsoluteFirstRow) {
                return (
                  <td
                    key={i}
                    colSpan={3}
                    rowSpan={totalRows}
                    className={`border border-gray-300 dark:border-slate-700 px-2 py-1 text-center font-bold align-middle ${bgHoliday} text-red-900 dark:text-red-300 border-l-2 border-l-gray-400 dark:border-l-slate-600 ${thickBorderClass}`}
                  >
                    {d.isSunday ? translate('common.holiday_sunday') : translate('common.holiday')}
                  </td>
                );
              }
              return null;
            }

            const isClickableManual = !isPercentage && !isFuture && !isHoliday;
            const isClickableTMS = !isPercentage && tmsRaw > 0;
            const hasDataManual = manualDisp !== null;

            return (
              <Fragment key={i}>
                <td
                  className={`${getCellClass(false)} relative ${isClickableTMS ? 'cursor-pointer hover:brightness-90 dark:hover:hover:brightness-120  transition-colors' : ''} ${getDataClass(bgClass, isHoliday, tmsDisp, masterTotal, true, 'TMS')}`}
                  onClick={() => {
                    if (isClickableTMS && onCellClick) {
                      onCellClick({
                        isTms: true,
                        date: d.str,
                        storage: cat,
                        type,
                        tmsCount: tmsRaw,
                        tmsDetails: tmsDetails,
                        routingName: d.routingName || '',
                      });
                    }
                  }}
                >
                  {isClickableTMS ? (
                    <div className="relative z-20 w-full min-h-5 flex items-center justify-center text-slate-700 dark:text-slate-200 hover:text-slate-900 dark:hover:text-white transition-all group/tms">
                      <span className="inline-block border-b-2 border-dotted border-red-700 dark:border-red-400 px-1 pb-0.5">
                        {formatValue(tmsDisp, masterTotal)}
                      </span>
                    </div>
                  ) : (
                    <div className="relative z-20 w-full min-h-5 flex items-center justify-center">
                      {formatValue(tmsDisp, masterTotal)}
                    </div>
                  )}
                </td>

                <td
                  className={`${getCellClass(false)} relative ${isClickableManual ? 'cursor-pointer group/cell hover:brightness-90 dark:hover:hover:brightness-120' : ''} ${getDataClass(
                    bgClass,
                    isHoliday,
                    manualDisp,
                    masterTotal,
                    true,
                    'MANUAL',
                    tmsDisp
                  )}`}
                  onClick={() => {
                    if (isClickableManual && onCellClick) {
                      onCellClick({
                        isTms: false,
                        date: d.str,
                        storage: cat,
                        type,
                        tmsCount: tmsRaw,
                        manualCount: manualRaw,
                        description: manualDesc,
                        id: manualId,
                        masterTotal: masterTotal || 0,
                      });
                    }
                  }}
                >
                  {isClickableManual && !hasDataManual && (
                    <div className="absolute inset-0 flex items-center justify-center opacity-0 group-hover/cell:opacity-100 transition-opacity pointer-events-none z-10">
                      <div className="absolute inset-0 bg-white/30 dark:bg-white/10"></div>
                      <span className="relative text-slate-600/30 dark:text-slate-400/30 text-xl font-bold">
                        +
                      </span>
                    </div>
                  )}

                  {hasDataManual ? (
                    <div className="relative z-20 w-full min-h-5 flex items-center justify-center group/edit">
                      <span
                        className={
                          isClickableManual
                            ? `inline-block border-b-2 border-dotted border-red-700 dark:border-red-400 px-1 pb-0.5`
                            : ''
                        }
                      >
                        {formatValue(manualDisp, masterTotal)}
                      </span>
                    </div>
                  ) : undefined}
                </td>

                <td
                  className={`${getCellClass(true)} ${getDataClass(bgClass, isHoliday, tvuDisp, masterTotal, true, 'TVU')}`}
                >
                  {formatValue(tvuDisp, masterTotal)}
                </td>
              </Fragment>
            );
          })}
        </tr>
      );
    });
  };

  const renderSpecialRow = (label, cat, bgClass, isBold = false, tooltip, type = '') => {
    const masterTotal = getMasterVal(cat);
    const isInterbranch = type === 'Interbranch';
    const isClickableMaster = !isPercentage && masterTotal > 0 && !isInterbranch;

    let clickStorage = cat;
    let clickType = type || 'Gabungan';
    if (cat === 'DryTotal') {
      clickStorage = 'Dry';
      clickType = 'Gabungan';
    }
    if (cat === 'FrozenTotal') {
      clickStorage = 'Frozen';
      clickType = 'Gabungan';
    }
    if (cat === 'OTV') {
      clickStorage = 'OTV';
      clickType = 'Gabungan';
    }

    const labelTable = (
      <td
        colSpan="2"
        className={`${tooltip && 'cursor-help'} ${tdClass} w-[250px] min-w-[250px] max-w-[250px] text-left font-bold sticky left-0 z-30 border-r border-gray-300 dark:border-slate-600 pl-4 ${bgClass}`}
      >
        <span
          className={tooltip && `border-b-2 border-dotted border-red-700 dark:border-red-400 pb-px`}
        >
          {label}
        </span>
      </td>
    );

    return (
      <tr className={isBold ? 'font-bold' : ''}>
        <Tooltip tooltipContent={tooltip || ''}>{labelTable}</Tooltip>
        {renderMasterCell(masterTotal, isClickableMaster, clickStorage, clickType, bgClass)}
        {dateKeys.map((d, i) => {
          const {
            tmsDisp,
            manualDisp,
            tvuDisp,
            tmsRaw,
            manualRaw,
            manualDesc,
            manualId,
            tmsDetails,
          } = getRowValues(d, cat, type);
          const isFuture = d.str > todayStr;
          const isHoliday = d.isSunday || d.isDynamicHoliday;
          if (isHoliday) return null;
          const isClickableManual = !isPercentage && !isFuture && isInterbranch && !isHoliday;
          const isClickableTMS = !isPercentage && tmsRaw > 0;
          const hasDataManual = manualDisp !== null;

          return (
            <Fragment key={i}>
              <td
                className={`${getCellClass(false)} relative ${isClickableTMS ? 'cursor-pointer hover:brightness-90 dark:hover:hover:brightness-120 transition-colors' : ''} ${getDataClass(bgClass, isHoliday, tmsDisp, masterTotal, false)}`}
                onClick={() => {
                  if (isClickableTMS && onCellClick) {
                    onCellClick({
                      isTms: true,
                      date: d.str,
                      storage: clickStorage,
                      type: type || 'Gabungan',
                      tmsCount: tmsRaw,
                      tmsDetails: tmsDetails,
                      routingName: d.routingName || '',
                    });
                  }
                }}
              >
                {isClickableTMS ? (
                  <div className="relative z-20 w-full min-h-5 flex items-center justify-center text-slate-700 dark:text-slate-200 hover:text-slate-800 dark:hover:text-white transition-all group/tms">
                    <span className="inline-block border-b-2 border-dotted border-red-700 dark:border-red-400 px-1 pb-0.5">
                      {formatValue(tmsDisp, masterTotal)}
                    </span>
                  </div>
                ) : (
                  <div className="relative z-20 w-full min-h-5 flex items-center justify-center">
                    {formatValue(tmsDisp, masterTotal)}
                  </div>
                )}
              </td>

              {isInterbranch ? (
                <td
                  className={`${getCellClass(false)} relative ${isClickableManual ? 'cursor-pointer group/cell hover:brightness-90 dark:hover:hover:brightness-120' : ''} ${getDataClass(bgClass, isHoliday, manualDisp, 0, false, 'MANUAL', tmsDisp)}`}
                  onClick={() => {
                    if (isClickableManual && onCellClick) {
                      onCellClick({
                        isTms: false,
                        date: d.str,
                        storage: clickStorage,
                        type: 'Interbranch',
                        tmsCount: tmsRaw,
                        manualCount: manualRaw,
                        description: manualDesc,
                        id: manualId,
                        masterTotal: 0,
                      });
                    }
                  }}
                >
                  {isClickableManual && !hasDataManual && (
                    <div className="absolute inset-0 flex items-center justify-center opacity-0 group-hover/cell:opacity-100 transition-opacity pointer-events-none z-10">
                      <span className="relative text-slate-600/30 dark:text-slate-400/30 text-xl font-bold">
                        +
                      </span>
                    </div>
                  )}

                  {hasDataManual ? (
                    <div className="relative z-20 w-full min-h-5 flex items-center justify-center group/edit">
                      <span
                        className={
                          isClickableManual
                            ? `inline-block border-b-2 border-dotted border-red-700 dark:border-red-400 px-1 pb-0.5`
                            : ''
                        }
                      >
                        {formatValue(manualDisp, masterTotal)}
                      </span>
                    </div>
                  ) : undefined}
                </td>
              ) : (
                <td
                  className={`${getCellClass(false)} ${getDataClass(bgClass, isHoliday, manualDisp, 0, false)}`}
                >
                  {formatValue(manualDisp, masterTotal)}
                </td>
              )}

              <td
                className={`${getCellClass(true)} ${getDataClass(bgClass, isHoliday, tvuDisp, masterTotal, false)}`}
              >
                {formatValue(tvuDisp, masterTotal)}
              </td>
            </Fragment>
          );
        })}
      </tr>
    );
  };

  const renderHeader = (data, tooltip, text, isLast = false) => {
    const isHoliday = data.isSunday || data.isDynamicHoliday;
    return (
      <Tooltip tooltipContent={tooltip}>
        <th
          className={`${thClass} ${isLast ? thickBorderClass : ''} cursor-help ${isHoliday ? bgHoliday : bgHeader}`}
        >
          {text}
        </th>
      </Tooltip>
    );
  };

  return (
    <div className="w-full">
      <table className="border-separate border-spacing-0 border-0 text-sm whitespace-nowrap min-w-max">
        <thead className="sticky top-0 z-40">
          <tr>
            <th
              rowSpan="2"
              className={`${thClass} w-[100px] min-w-[100px] max-w-[100px] sticky left-0 z-50 ${bgHeader}`}
            >
              {translate('common.storage_type')}
            </th>
            <th
              rowSpan="2"
              className={`${thClass} w-[150px] min-w-[150px] max-w-[150px] sticky left-[100px] z-50 ${bgHeader}`}
            >
              {translate('common.vehicle_type')}
            </th>
            <th
              rowSpan="2"
              className={`${thClass} w-[60px] min-w-[60px] sticky left-[250px] z-50 ${thickBorderClass} ${bgHeader}`}
            >
              Total
            </th>
            {dateKeys.map((d, i) => {
              const isHoliday = d.isSunday || d.isDynamicHoliday;
              const headerColor = isHoliday ? bgHoliday : bgHeader;

              let headerContent;
              if (d.isDynamicHoliday) {
                headerContent = (
                  <Tooltip tooltipContent={translate('summary.tabs.task_summary.caution')}>
                    <span className="cursor-help border-b-2 border-dotted border-red-900 dark:border-red-300 pb-0.5">
                      {formatDateUniversal(d.str, 'DD-MM-YYYY')}
                    </span>
                  </Tooltip>
                );
              } else {
                headerContent = <span>{formatDateUniversal(d.str, 'DD-MM-YYYY')}</span>;
              }
              return (
                <th
                  key={i}
                  colSpan="3"
                  className={`${thClass} ${thickBorderClass} min-w-[90px] ${headerColor}`}
                >
                  {headerContent}
                </th>
              );
            })}
          </tr>
          <tr>
            {dateKeys.map((d, i) => (
              <Fragment key={i}>
                {renderHeader(d, translate('summary.tabs.truck_usage.tms'), 'TMS')}
                {renderHeader(d, translate('summary.tabs.truck_usage.non_tms'), 'Non TMS')}
                {renderHeader(d, translate('summary.tabs.truck_usage.tvu'), 'TVU', true)}
              </Fragment>
            ))}
          </tr>
        </thead>
        <tbody>
          {renderSectionRows('Dry', bgDry, dryTypes)}
          {renderSpecialRow(
            translate('summary.tabs.truck_usage.interbranch'),
            'Dry',
            bgDry,
            false,
            null,
            'Interbranch'
          )}
          {renderSpecialRow(
            translate('summary.tabs.truck_usage.total_used'),
            'DryTotal',
            bgDryTotal,
            true
          )}

          {renderSectionRows('Frozen', bgFrozen, frozenTypes)}
          {renderSpecialRow(
            translate('summary.tabs.truck_usage.interbranch'),
            'Frozen',
            bgFrozen,
            false,
            null,
            'Interbranch'
          )}
          {renderSpecialRow(
            translate('summary.tabs.truck_usage.total_used'),
            'FrozenTotal',
            bgFrozenTotal,
            true
          )}

          {renderSpecialRow('OTV', 'OTV', bgOTV, true, translate('summary.tabs.truck_usage.otv'))}
        </tbody>
      </table>
    </div>
  );
}
