import { formatDateUniversal } from '@/lib/utils';
import * as XLSX from 'xlsx-js-style';
import { BASE_STYLES, BORDERS, FILL_STYLES, FONT_STYLES, HEADER_STYLES } from './reportStyles';

function calculatePct(num, den) {
  const n = num || 0;
  const d = den || 0;
  if (d === 0) return 0;
  return n / d;
}

export function generateTaskSummarySheet(
  wb,
  metrics,
  startDateStr,
  endDateStr,
  masterTruckData,
  translate,
  isDailyReport = false,
  isManualMode = false
) {
  const headers = [
    translate('common.delivery_date'),
    translate('common.type'),
    'DP',
    'DT',
    '%DT',
    'MA',
    '%MA',
    'RT',
    '%RT',
    'CO',
    '%CO',
    'PR',
    '%PR',
    'MT',
    'TV',
    'VA',
    'TVU',
    '%TVU',
  ];

  const excelData = [headers];
  const merges = [];
  const sundayRows = new Set();
  const zeroDpRows = new Set();
  const current = new Date(startDateStr);
  const end = new Date(endDateStr);
  let currentRow = 1;
  const mtDry = masterTruckData?.masterData?.Dry?.Total || 0;
  const mtFrozen = masterTruckData?.masterData?.Frozen?.Total || 0;

  while (current <= end) {
    const day = current.getDate().toString().padStart(2, '0');
    const month = (current.getMonth() + 1).toString().padStart(2, '0');
    const year = current.getFullYear();
    const displayDate = `${day}-${month}-${year}`;
    const isSunday = current.getDay() === 0;

    const routingKey = formatDateUniversal(current);
    const data = metrics ? metrics[routingKey] : null;
    const d = data?.dry || {};
    const f = data?.frozen || {};

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const currentMidnight = new Date(current);
    currentMidnight.setHours(0, 0, 0, 0);
    const isPast = currentMidnight < today;

    const checkZero = (obj) =>
      (obj.dp || 0) === 0 &&
      (obj.dt_total || 0) === 0 &&
      (obj.ma_total || 0) === 0 &&
      (obj.rt || 0) === 0 &&
      (obj.co || 0) === 0 &&
      (obj.pr || 0) === 0 &&
      (obj.tv || 0) === 0 &&
      (obj.va || 0) === 0 &&
      (obj.tvu || 0) === 0;

    const isDynamicHoliday = isPast && checkZero(d) && checkZero(f) && !isSunday;

    if (isSunday || isDynamicHoliday) {
      excelData.push([
        displayDate,
        isSunday ? translate('common.holiday_sunday') : translate('common.holiday'),
        ...Array(16).fill(''),
      ]);
      excelData.push(['', '', ...Array(16).fill('')]);

      merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow + 1, c: 0 } });
      merges.push({ s: { r: currentRow, c: 1 }, e: { r: currentRow + 1, c: 17 } });

      sundayRows.add(currentRow);
      sundayRows.add(currentRow + 1);
      currentRow += 2;
    } else {
      if ((d.dp || 0) === 0 && (f.dp || 0) === 0 && isPast) {
        zeroDpRows.add(currentRow);
        zeroDpRows.add(currentRow + 1);
      }

      excelData.push([
        displayDate,
        'Dry',
        d.dp || 0,
        isManualMode ? 0 : d.dt_total || 0,
        isManualMode ? '-' : calculatePct(d.dt_total, d.dp),
        isManualMode ? 0 : d.ma_total || 0,
        isManualMode ? '-' : calculatePct(d.ma_total, d.dp),
        d.rt || 0,
        calculatePct(d.rt, d.dp),
        d.co || 0,
        calculatePct(d.co, d.dp),
        d.pr || 0,
        calculatePct(d.pr, d.dp),
        mtDry,
        d.tv || 0,
        d.va || 0,
        d.tvu || 0,
        calculatePct(d.tvu, mtDry),
      ]);

      excelData.push([
        '',
        'Frozen',
        f.dp || 0,
        isManualMode ? 0 : f.dt_total || 0,
        isManualMode ? '-' : calculatePct(f.dt_total, f.dp),
        isManualMode ? 0 : f.ma_total || 0,
        isManualMode ? '-' : calculatePct(f.ma_total, f.dp),
        f.rt || 0,
        calculatePct(f.rt, f.dp),
        f.co || 0,
        calculatePct(f.co, f.dp),
        f.pr || 0,
        calculatePct(f.pr, f.dp),
        mtFrozen,
        f.tv || 0,
        f.va || 0,
        f.tvu || 0,
        calculatePct(f.tvu, mtFrozen),
      ]);

      merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow + 1, c: 0 } });
      currentRow += 2;
    }

    current.setDate(current.getDate() + 1);
  }

  excelData.push([]);

  excelData.push([translate('common.explanation')]);
  const legendTitleRow = excelData.length - 1;

  merges.push({ s: { r: legendTitleRow, c: 0 }, e: { r: legendTitleRow, c: 2 } });

  const legendMap = [
    { label: 'DP', key: 'dp' },
    { label: 'DT', key: 'dt' },
    { label: 'MA', key: 'ma' },
    { label: 'RT', key: 'rt' },
    { label: 'CO', key: 'co' },
    { label: 'PR', key: 'pr' },
    { label: 'MT', key: 'mt' },
    { label: 'TV', key: 'tv' },
    { label: 'VA', key: 'va' },
    { label: 'TVU', key: 'tvu' },
  ];

  const legendStartRow = excelData.length;

  legendMap.forEach((item, idx) => {
    const label = `${item.label}`;
    const desc =
      item.key === 'ma'
        ? translate('common.status.manual_assign')
        : translate(`summary.tabs.task_summary.${item.key}`);
    excelData.push([label, desc]);

    const r = legendStartRow + idx;
    merges.push({ s: { r: r, c: 1 }, e: { r: r, c: 17 } });
  });

  const ws = XLSX.utils.aoa_to_sheet(excelData);
  ws['!merges'] = merges;

  const range = XLSX.utils.decode_range(ws['!ref']);
  const headerFills = {
    0: FILL_STYLES.dry,
    1: FILL_STYLES.dry,
    2: FILL_STYLES.taskPink,
    3: FILL_STYLES.taskGreen,
    4: FILL_STYLES.taskGreen,
    5: FILL_STYLES.taskRed,
    6: FILL_STYLES.taskRed,
    7: FILL_STYLES.taskCyan,
    8: FILL_STYLES.taskCyan,
    9: FILL_STYLES.taskBlue,
    10: FILL_STYLES.taskBlue,
    11: FILL_STYLES.taskGray,
    12: FILL_STYLES.taskGray,
    13: FILL_STYLES.dry,
    14: FILL_STYLES.dry,
    15: FILL_STYLES.dry,
    16: FILL_STYLES.taskViolet,
    17: FILL_STYLES.taskViolet,
  };
  const pctCols = [4, 6, 8, 10, 12, 17];

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
      if (R === legendTitleRow - 1) {
        if (ws[cellRef]) delete ws[cellRef];
        continue;
      }
      if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };
      const cell = ws[cellRef];

      if (R >= legendTitleRow) {
        if (R === legendTitleRow && C === 0) {
          cell.s = { font: { bold: true, underline: true }, alignment: { horizontal: 'left' } };
        } else if (R >= legendStartRow) {
          if (C === 0) {
            const legendIndex = R - legendStartRow;
            const legendKey = legendMap[legendIndex]?.key;
            let legendFill = null;
            switch (legendKey) {
              case 'dp':
                legendFill = FILL_STYLES.taskPink;
                break;
              case 'dt':
                legendFill = FILL_STYLES.taskGreen;
                break;
              case 'ma':
                legendFill = FILL_STYLES.taskRed;
                break;
              case 'rt':
                legendFill = FILL_STYLES.taskCyan;
                break;
              case 'co':
                legendFill = FILL_STYLES.taskBlue;
                break;
              case 'pr':
                legendFill = FILL_STYLES.taskGray;
                break;
              case 'mt':
              case 'tv':
              case 'va':
                legendFill = FILL_STYLES.dry;
                break;
              case 'tvu':
                legendFill = FILL_STYLES.taskViolet;
                break;
              default:
                legendFill = null;
            }

            cell.s = {
              font: { bold: true },
              alignment: { horizontal: 'center', vertical: 'top' },
              fill: legendFill,
              border: BORDERS.thin,
            };
          } else if (C === 1) {
            cell.s = { alignment: { horizontal: 'left', wrapText: true, vertical: 'top' } };
          }
        }
        continue;
      }
      if (R === 0) {
        cell.s = {
          ...HEADER_STYLES.main,
          fill: headerFills[C] || HEADER_STYLES.main.fill,
          border: BORDERS.thin,
        };
      } else {
        cell.s = { ...BASE_STYLES.center };
        if (sundayRows.has(R)) {
          cell.s.fill = FILL_STYLES.red;
          cell.s.font = { ...FONT_STYLES.bold, color: { rgb: '990000' } };
          if (C === 1) cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        } else {
          if (C === 0) {
            cell.s = {
              ...BASE_STYLES.center,
              alignment: { vertical: 'center', horizontal: 'center' },
            };
            if (zeroDpRows.has(R)) {
              cell.s.fill = { patternType: 'solid', fgColor: { rgb: 'FFC7CE' } };
              cell.s.font = { ...FONT_STYLES.bold, color: { rgb: '9C0006' } };
            }
          }
          if (C === 1) cell.s.font = FONT_STYLES.bold;
          if (pctCols.includes(C)) {
            cell.t = 'n';
            cell.s.numFmt = '0.00%';
            cell.s.fill = headerFills[C];
          } else if (isDailyReport && [2, 3, 5, 7, 9, 11, 14, 15].includes(C)) {
            cell.s.fill = FILL_STYLES.taskYellow;
          }
          if (isManualMode && [3, 4, 5, 6].includes(C)) {
            cell.s.fill = { patternType: 'solid', fgColor: { rgb: 'FFC7CE' } };
          }
        }
      }
    }
  }

  ws['!cols'] = [
    { wch: 12 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 12 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, translate('summary.tabs.task_summary.title'));
}
