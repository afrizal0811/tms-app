import { getLocalStorage } from '@/lib/localStorageHandler';
import { toastError, toastSuccess } from '@/lib/toast';
import { formatDateUniversal, isEmpty } from '@/lib/utils';
import * as XLSX from 'xlsx-js-style';

export const handleDownloadExcel = (processedData, setIsDownloading, selectedDate, t) => {
  if (isEmpty(processedData)) {
    toastError(t('common.toast.error', { err: t('common.no_data') }));
    return;
  }

  setIsDownloading(true);
  try {
    const wb = XLSX.utils.book_new();
    const headers = [
      'No',
      t('common.delivery_date'),
      t('common.invoice_number'),
      t('common.customer_name'),
      t('common.customer_id'),
      t('common.location_id'),
      t('common.coord_new'),
      t('common.dist_diff'),
      t('common.driver'),
      t('common.updated_time'),
    ];
    const sheetData = [headers];

    processedData.forEach((row, index) => {
      const displayCustId = row.isIncomplete ? '-' : row.customerId || '';
      const displayLocId = row.isIncomplete ? '-' : row.locationId || '';
      sheetData.push([
        index + 1,
        row.deliveryDateDisplay,
        row.soNumber,
        row.customerName,
        displayCustId,
        displayLocId,
        row.newLonglat,
        row.distanceDiff,
        row.driverName,
        row.updateTime,
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    const headerStyle = {
      font: { bold: true, color: { rgb: '000000' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      fill: { fgColor: { rgb: 'EFEFEF' } },
    };
    const redFillStyle = { fill: { patternType: 'solid', fgColor: { rgb: 'FFC7CE' } } };
    const range = XLSX.utils.decode_range(ws['!ref']);
    ws['!cols'] = [
      { wch: 5 },
      { wch: 15 },
      { wch: 30 },
      { wch: 30 },
      { wch: 25 },
      { wch: 20 },
      { wch: 25 },
      { wch: 15 },
      { wch: 20 },
      { wch: 15 },
    ];

    for (let R = range.s.r; R <= range.e.r; ++R) {
      const rowData = R > 0 ? processedData[R - 1] : null;
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };

        if (R === 0) {
          ws[cellRef].s = {
            ...headerStyle,
            alignment: {
              ...(headerStyle.alignment || {}),
              horizontal: 'center',
              wrapText: C === 6,
            },
          };
        } else {
          ws[cellRef].s = rowData && rowData.isIncomplete ? { ...redFillStyle } : {};
          const isLeftCol = C === 2 || C === 3 || C === 8;
          ws[cellRef].s.alignment = { horizontal: isLeftCol ? 'left' : 'center' };
        }
      }
    }

    XLSX.utils.book_append_sheet(wb, ws, t('coordinate.title'));
    let dateStr = '';
    if (Array.isArray(selectedDate)) {
      const [start, end] = selectedDate;
      const fStart = formatDateUniversal(start, 'DD.MM.YYYY');
      const fEnd = formatDateUniversal(end, 'DD.MM.YYYY');
      dateStr = fStart === fEnd ? fStart : `${fStart} - ${fEnd}`;
    } else {
      dateStr = formatDateUniversal(selectedDate, 'DD.MM.YYYY');
    }
    const { storedLocationAcronym: locationName } = getLocalStorage() || '-';
    const fileName = `${t('coordinate.title')} - ${dateStr} - ${locationName}.xlsx`;
    XLSX.writeFile(wb, fileName);
    toastSuccess(t('common.toast.success'));
  } catch (e) {
    toastError(t('common.toast.error', { err: e.message }), e);
  } finally {
    setIsDownloading(false);
  }
};
