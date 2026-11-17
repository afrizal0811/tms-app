// File: src/features/vehicleData/help.js
import { getLocalStorage } from '@/lib/localStorageHandler';
import { formatDateUniversal } from '@/lib/utils';
import * as XLSX from 'xlsx-js-style';
import { toastError, toastSuccess } from '../../lib/toast';

const COLOR_INCOMPLETE = 'FFD9D9';
const COLOR_DUPLICATE = 'FFF2CC';

const applyRowColorsAndLegend = (ws, rawData, t) => {
  const range = XLSX.utils.decode_range(ws['!ref']);

  rawData.forEach((v, idx) => {
    let bgColor = null;
    if (v.isIncomplete) bgColor = COLOR_INCOMPLETE;
    else if (v.isDuplicateDriver) bgColor = COLOR_DUPLICATE;

    if (bgColor) {
      const rowIndex = idx + 1;
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: C });
        if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };

        ws[cellRef].s = {
          ...ws[cellRef].s,
          fill: { fgColor: { rgb: bgColor } },
        };
      }
    }
  });

  const legendStartRow = range.e.r + 2;

  ws[XLSX.utils.encode_cell({ r: legendStartRow, c: 0 })] = {
    t: 's',
    v: t('common.color_exp'),
    s: { font: { bold: true } },
  };

  ws[XLSX.utils.encode_cell({ r: legendStartRow + 1, c: 0 })] = {
    t: 's',
    v: '',
    s: { fill: { fgColor: { rgb: COLOR_INCOMPLETE } } },
  };
  ws[XLSX.utils.encode_cell({ r: legendStartRow + 1, c: 1 })] = {
    t: 's',
    v: t('vehicle.tabs.incomplete_data'),
  };

  ws[XLSX.utils.encode_cell({ r: legendStartRow + 2, c: 0 })] = {
    t: 's',
    v: '',
    s: { fill: { fgColor: { rgb: COLOR_DUPLICATE } } },
  };
  ws[XLSX.utils.encode_cell({ r: legendStartRow + 2, c: 1 })] = {
    t: 's',
    v: t('vehicle.tabs.duplicate_driver'),
  };

  range.e.r = legendStartRow + 2;
  if (range.e.c < 1) range.e.c = 1;
  ws['!ref'] = XLSX.utils.encode_range(range);
};

export const handleConfirmDownload = ({
  conditionalData,
  masterData,
  setIsDownloadDropdownOpen = () => {},
  setIsDownloading,
  sheetSelection,
  t,
  templateData,
  fileNamePrefix = '',
}) => {
  setIsDownloading(true);
  try {
    const wb = XLSX.utils.book_new();
    const headerStyle = { font: { bold: true }, fill: { fgColor: { rgb: 'EFEFEF' } } };
    const tableHeader = [
      t('common.license_number'),
      t('common.type'),
      t('common.name'),
      t('common.email'),
    ];

    if (sheetSelection.master) {
      const data1 = masterData.map((v) => [v.plat, v.type || null, v.name || null, v.email]);
      const ws1 = XLSX.utils.aoa_to_sheet([tableHeader, ...data1]);
      ws1['!cols'] = [{ wch: 25 }, { wch: 25 }, { wch: 30 }, { wch: 30 }];
      ['A1', 'B1', 'C1', 'D1'].forEach((cell) => {
        if (ws1[cell]) ws1[cell].s = headerStyle;
      });

      applyRowColorsAndLegend(ws1, masterData, t);
      XLSX.utils.book_append_sheet(wb, ws1, t('vehicle.tabs.master_title'));
    }

    if (sheetSelection.conditional && conditionalData.length > 0) {
      const dataC = conditionalData.map((v) => [v.plat, v.type || null, v.name || null, v.email]);
      const wsC = XLSX.utils.aoa_to_sheet([tableHeader, ...dataC]);
      wsC['!cols'] = [{ wch: 25 }, { wch: 25 }, { wch: 30 }, { wch: 30 }];
      ['A1', 'B1', 'C1', 'D1'].forEach((cell) => {
        if (wsC[cell]) wsC[cell].s = headerStyle;
      });

      applyRowColorsAndLegend(wsC, conditionalData, t);
      XLSX.utils.book_append_sheet(wb, wsC, t('vehicle.tabs.conditional_title'));
    }

    if (sheetSelection.template) {
      const headers2 = [
        'Name*',
        'Assignee',
        'Start Time',
        'End Time',
        'Break Start',
        'Break End',
        'Multiday',
        'Speed Km/h',
        'Cost Factor',
        'Vehicle Tags',
        'Odd Even',
        'weight Min',
        'weight Max',
        'volume Min',
        'volume Max',
      ];
      const data2 = templateData.map((v) => [
        v.plat,
        v.email,
        v.workingTime.startTime || null,
        v.workingTime.endTime || null,
        v.breakTime.startBreakTime || null,
        v.breakTime.endBreakTime || null,
        v.workingTime.multiday || 0,
        v.speed,
        v.costFactor,
        v.parsedTags?.join('; ') || null,
        v.oddEven,
        v.minWeight || 0,
        v.maxWeight || 0,
        v.minVolume || 0,
        Number(v.maxVolume) || 0,
      ]);
      const ws2 = XLSX.utils.aoa_to_sheet([headers2, ...data2]);
      ws2['!cols'] = Array(headers2.length).fill({ wch: 20 });
      headers2.forEach((h, i) => {
        const cellRef = XLSX.utils.encode_cell({ c: i, r: 0 });
        if (ws2[cellRef]) ws2[cellRef].s = headerStyle;
      });

      XLSX.utils.book_append_sheet(wb, ws2, t('vehicle.tabs.template_title'));
    }

    const { storedLocationAcronym: locationName } = getLocalStorage() || '-';
    const date = formatDateUniversal(new Date(), 'DD.MM.YYYY');
    const prefix = fileNamePrefix ? `${fileNamePrefix} - ` : '';
    const fileName = `${t('common.vehicle_data')} - ${prefix}${date} - ${locationName}.xlsx`;
    XLSX.writeFile(wb, fileName);
    toastSuccess(t('common.toast.success'));
  } catch (e) {
    toastError(t('common.toast.error', { err: e.message }), e);
  } finally {
    setIsDownloading(false);
    setIsDownloadDropdownOpen(false);
  }
};
