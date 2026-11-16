import { getLocalStorage } from '@/lib/localStorageHandler';
import {
  checkInvalidSo,
  checkInvalidSoList,
  formatDateUniversal,
  getBasePlate,
  isEmpty,
  parseCustomerString,
  sortRows,
} from '@/lib/utils';
import JSZip from 'jszip';
import * as XLSX from 'xlsx-js-style';
import { toastError, toastSuccess } from '../../../lib/toast';
import {
  abortIfNoRoutingResults,
  buildEnrichedTripsMap,
  getLocationName,
  getUniqueFileName,
  resolveDedupedTrip,
  sortRoutingResultsByCreatedTime,
  triggerDownload,
} from './shared';

const appendDeliveryListSheet = (wb, plat, driverName, tripsData, isDetailView, t) => {
  const activeCols = [
    { key: 'no', title: 'No.' },
    { key: 'visit', title: t('delivery.visit') },
    { key: 'custId', title: t('common.customer_id') },
    { key: 'locId', title: t('common.location_id') },
    { key: 'so', title: t('common.invoice_number') },
    { key: 'openTime', title: t('common.open_time') },
    { key: 'closeTime', title: t('common.close_time') },
    { key: 'eta', title: t('delivery.est_arrival') },
    { key: 'etd', title: t('delivery.est_depart') },
  ];

  const hasManualInRoute = tripsData.some((pt) => pt.trip.isManual);

  const wsData = [
    [t('common.vehicle'), plat],
    [t('common.driver'), driverName],
    [],
    activeCols.map((col) => col.title),
  ];

  const stylingMeta = [
    { row: 0, col: 0, style: { font: { bold: true } } },
    { row: 0, col: 1, style: { font: { bold: true, sz: 12 } } },
    { row: 1, col: 0, style: { font: { bold: true } } },
  ];

  activeCols.forEach((_, i) => {
    stylingMeta.push({
      row: 3,
      col: i,
      style: {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '4F46E5' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
      },
    });
  });

  let currentRowIndex = 4;

  tripsData.forEach(({ trip, isFirstHub, isLastHub, isMisplacedMiddleHub }) => {
    const isHub = trip.isHub;
    const parsedCust = parseCustomerString(trip.visitName);
    const isBadCust = isEmpty(parsedCust?.id) || isEmpty(parsedCust?.location);

    let baseOutletName = isHub
      ? 'HUB'
      : trip.flow === 'Pickup' && trip.warehouseName
        ? trip.warehouseName
        : parsedCust?.name || trip.visitName;

    if (trip.isReDelivery && !isHub) baseOutletName = `[R] ${baseOutletName}`;

    const custId = isHub ? '' : parsedCust?.id || '-';
    const locId = isHub ? '' : parsedCust?.location || '-';
    const mapping = trip.soWarehouseMapping || [];

    const pushRow = (
      displayNo,
      displayCustId,
      displayLocId,
      displaySo,
      whInfo = null,
      isSplit = false,
      isUnsyncOverride = null,
      partnerOverride = null,
      isInvalidSo = false
    ) => {
      const openVal = isHub ? '' : trip.openTime || '-';
      const closeVal = isHub ? '' : trip.closeTime || '-';
      const etaVal = isFirstHub ? '' : trip.eta ? formatDateUniversal(trip.eta, 'HH:mm') : '-';
      const etdVal = isLastHub ? '' : trip.etd ? formatDateUniversal(trip.etd, 'HH:mm') : '-';

      const isRowUnsync = isUnsyncOverride !== null ? isUnsyncOverride : trip.isUnsync;
      const rowPartner = partnerOverride !== null ? partnerOverride : trip.partnerVehicle;

      let outletWithWh =
        isDetailView && whInfo ? `${baseOutletName}\n↳ Pickup: ${whInfo}` : baseOutletName;

      if (isRowUnsync && rowPartner) outletWithWh += `\n[Partner: ${rowPartner}]`;

      wsData.push(
        activeCols.map(
          (col) =>
            ({
              no: displayNo,
              visit: outletWithWh,
              custId: displayCustId,
              locId: displayLocId,
              so: displaySo,
              openTime: openVal,
              closeTime: closeVal,
              eta: etaVal,
              etd: etdVal,
            })[col.key]
        )
      );

      const cellStyle = {
        border: {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        },
        alignment: { vertical: 'center', wrapText: true },
        ...(isMisplacedMiddleHub
          ? { fill: { fgColor: { rgb: 'FEE2E2' } } }
          : trip.isManual
            ? { fill: { fgColor: { rgb: 'E6EEFF' } } }
            : {}),
        ...(isHub && { font: { color: { rgb: 'DC2626' }, bold: true } }),
      };

      activeCols.forEach((col, c) => {
        let colStyle = { ...cellStyle };
        let comment = null;

        if (col.key === 'visit' && isMisplacedMiddleHub) {
          comment = [
            {
              a: 'Info',
              t: 'Urutan Hub tambahan tidak sesuai karena perubahan pengurutan',
              h: true,
            },
          ];
        }

        if (col.key === 'no') {
          colStyle.alignment = { ...colStyle.alignment, horizontal: 'right' };
          if (isSplit && !isHub) {
            colStyle.font = { ...(colStyle.font || {}), color: { rgb: '16A34A' }, bold: true };
          }
        }
        if (col.key === 'visit' && trip.isReDelivery && !isHub) {
          colStyle.font = { ...(colStyle.font || {}), color: { rgb: 'DC2626' } };
        }
        if (col.key === 'eta' && isLastHub && hasManualInRoute && trip.eta) {
          comment = [{ a: 'Info', t: t('delivery.hub_eta_short'), h: true }];
        }
        if (col.key === 'so' && isInvalidSo && !isHub) {
          colStyle.font = { ...(colStyle.font || {}), color: { rgb: 'DC2626' }, bold: true };
        }

        stylingMeta.push({ row: currentRowIndex, col: c, style: colStyle, comment });
      });
      currentRowIndex++;
    };

    if (!isHub && isDetailView && mapping.length > 0) {
      mapping.forEach((item, idx) => {
        const letter = mapping.length > 1 ? String.fromCharCode(65 + idx) : '';
        const displayNo = trip.isManual ? '-' : `${trip.routePlannedOrder}${letter}`;
        const soPartner = trip.syncDetails?.[item.so] || null;
        pushRow(
          displayNo,
          custId,
          locId,
          item.so,
          trip.flow !== 'Pickup' ? item.wh : null,
          mapping.length > 1,
          !!soPartner,
          soPartner,
          checkInvalidSo(item.so, isBadCust)
        );
      });
      return;
    }

    const invoiceToValidate =
      trip.orderIdOverride || parsedCust?.invoiceNumber || trip.orderId || '';
    const isInvalidSoTotal = isHub ? false : checkInvalidSoList(invoiceToValidate, isBadCust);

    let displaySo = isHub
      ? null
      : mapping.length > 0
        ? mapping
            .map((item) =>
              item.wh && trip.flow !== 'Pickup' ? `${item.so} (${item.wh})` : item.so
            )
            .join(', ')
        : invoiceToValidate || '-';

    pushRow(
      isHub ? '' : trip.isManual ? '-' : trip.routePlannedOrder,
      custId,
      locId,
      displaySo,
      null,
      false,
      null,
      null,
      isInvalidSoTotal
    );
  });

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  stylingMeta.forEach((meta) => {
    const cellRef = XLSX.utils.encode_cell({ r: meta.row, c: meta.col });
    if (!ws[cellRef]) ws[cellRef] = { v: '' };
    ws[cellRef].s = { ...(ws[cellRef].s || {}), ...meta.style };
    if (meta.comment) ws[cellRef].c = meta.comment;
  });

  ws['!cols'] = activeCols.map((_, colIndex) => {
    let maxLen = 0;
    for (let r = 3; r < wsData.length; r++) {
      const val = wsData[r][colIndex];
      if (val != null) {
        const lines = val.toString().split('\n');
        maxLen = Math.max(maxLen, ...lines.map((l) => l.length));
      }
    }
    return { wch: Math.min(Math.max(maxLen + 2, 8), 60) };
  });

  let finalSheetName = plat;
  let counter = 1;
  while (wb.SheetNames.includes(finalSheetName)) {
    finalSheetName = `${plat}_${counter++}`;
  }
  XLSX.utils.book_append_sheet(wb, ws, finalSheetName);
};

export const handleFullDeliveryListDownload = async ({
  filteredVehicleRoutes,
  setIsDownloading,
  t,
  fileNamePrefix,
  isDetailView,
  sortConfig,
  selectedDate,
}) => {
  setIsDownloading(true);
  try {
    const wb = XLSX.utils.book_new();

    filteredVehicleRoutes.forEach((route) => {
      const plat = route.basePlat;
      const driverName = route.driverName || route.assigneeName || '-';

      const isDefaultSort = sortConfig?.key === 'no' && sortConfig?.direction === 'asc';

      const processedTrips = route.trips.map((trip, index) => ({
        trip,
        isFirstHub: index === 0 && trip.isHub,
        isLastHub: index === route.trips.length - 1 && trip.isHub,
        isMisplacedMiddleHub: trip.isMiddleHub && !isDefaultSort,
      }));

      if (sortConfig) {
        processedTrips.sort((a, b) => {
          if (a.isFirstHub) return -1;
          if (b.isFirstHub) return 1;
          if (a.isLastHub) return 1;
          if (b.isLastHub) return -1;

          if (!isDefaultSort) {
            if (a.trip.isMiddleHub && !b.trip.isMiddleHub) return 1;
            if (!a.trip.isMiddleHub && b.trip.isMiddleHub) return -1;
          }

          if (sortConfig.key === 'no') {
            const noA = a.trip.routePlannedOrder ?? (a.trip.isManual ? 9999 : 0);
            const noB = b.trip.routePlannedOrder ?? (b.trip.isManual ? 9999 : 0);
            return sortConfig.direction === 'asc' ? noA - noB : noB - noA;
          }
          if (sortConfig.key === 'so') {
            const soA = String(a.trip.orderId || '');
            const soB = String(b.trip.orderId || '');
            return sortConfig.direction === 'asc' ? soA.localeCompare(soB) : soB.localeCompare(soA);
          }
          return 0;
        });
      }

      appendDeliveryListSheet(wb, plat, driverName, processedTrips, isDetailView, t);
    });

    const { storedLocationAcronym: locationName } = getLocalStorage() || '-';
    const date = formatDateUniversal(selectedDate, 'DD.MM.YYYY');
    const fileName = fileNamePrefix
      ? `${t('delivery.delivery_list')} (${fileNamePrefix}) - ${date} - ${locationName}.xlsx`
      : `${t('delivery.delivery_list')} - ${date} - ${locationName}.xlsx`;

    XLSX.writeFile(wb, fileName);
    toastSuccess(t('common.toast.success'));
  } catch (e) {
    toastError(t('common.toast.error', { err: e.message }), e);
  } finally {
    setIsDownloading(false);
  }
};

export const handlePartialDeliveryListDownload = async ({
  routingResults,
  filteredVehicleRoutes,
  setIsDownloading,
  t,
  fileNamePrefix,
  isDetailView,
  selectedDate,
  sortConfig,
}) => {
  setIsDownloading(true);
  try {
    const dateForFilename = formatDateUniversal(selectedDate, 'DD.MM.YYYY');
    const locationName = getLocationName();

    if (abortIfNoRoutingResults(routingResults, t, setIsDownloading)) return;

    const sortedRoutingResults = sortRoutingResultsByCreatedTime(routingResults);

    const enrichedTripsMap = buildEnrichedTripsMap(filteredVehicleRoutes);

    const masterZip = new JSZip();
    let masterHasData = false;
    const globalSeenSOByVehicle = new Map();
    let routingIndex = 1;

    for (const routing of sortedRoutingResults) {
      const status = routing.status || routing.dispatchStatus || '';
      if (String(status).toLowerCase() !== 'done') continue;
      const plates = Array.from(
        new Set([
          ...Object.keys(routing.totalDistance || {}),
          ...Object.keys(routing.time || {}),
          ...Object.keys(routing.hubTimes || {}),
        ])
      );
      if (plates.length === 0) continue;

      const routes = plates
        .map((rawPlat) => {
          const plat = getBasePlate(rawPlat) || rawPlat;
          return (filteredVehicleRoutes || []).find(
            (r) => (r.basePlat || getBasePlate(r.plat)) === plat
          );
        })
        .filter(Boolean);
      if (routes.length === 0) continue;

      sortRows(routes, 'vehicleName', 'vehicleName');

      const cleanRoutingName = (routing.name || '').trim();
      const zipFolderName = `Routing ${routingIndex} (${cleanRoutingName})`;
      routingIndex++;

      const routingZip = new JSZip();
      let routingHasData = false;
      const seenFileNames = new Set();

      for (const route of routes) {
        const plat = route.basePlat;
        if (!globalSeenSOByVehicle.has(plat)) {
          globalSeenSOByVehicle.set(plat, new Set());
        }
        const vehicleSeenSOs = globalSeenSOByVehicle.get(plat);
        const processedTripsToRender = [];

        const isDefaultSort = sortConfig?.key === 'no' && sortConfig?.direction === 'asc';

        (route.trips || []).forEach((rawTrip, index) => {
          const isHub = rawTrip.isHub;
          const isFirstHub = index === 0 && isHub;
          const isLastHub = index === route.trips.length - 1 && isHub;
          const trip = resolveDedupedTrip(rawTrip, enrichedTripsMap, vehicleSeenSOs);
          if (trip === null) return;
          processedTripsToRender.push({
            trip,
            isFirstHub,
            isLastHub,
            isMisplacedMiddleHub: trip.isMiddleHub && !isDefaultSort,
          });
        });

        if (sortConfig) {
          processedTripsToRender.sort((a, b) => {
            if (a.isFirstHub) return -1;
            if (b.isFirstHub) return 1;
            if (a.isLastHub) return 1;
            if (b.isLastHub) return -1;

            if (!isDefaultSort) {
              if (a.trip.isMiddleHub && !b.trip.isMiddleHub) return 1;
              if (!a.trip.isMiddleHub && b.trip.isMiddleHub) return -1;
            }

            if (sortConfig.key === 'no') {
              const noA = a.trip.routePlannedOrder ?? (a.trip.isManual ? 9999 : 0);
              const noB = b.trip.routePlannedOrder ?? (b.trip.isManual ? 9999 : 0);
              return sortConfig.direction === 'asc' ? noA - noB : noB - noA;
            }
            if (sortConfig.key === 'so') {
              const soA = String(a.trip.orderId || '');
              const soB = String(b.trip.orderId || '');
              return sortConfig.direction === 'asc'
                ? soA.localeCompare(soB)
                : soB.localeCompare(soA);
            }
            return 0;
          });
        }

        if (
          processedTripsToRender.length <= 2 &&
          processedTripsToRender.every((pt) => pt.trip.isHub)
        ) {
          continue;
        }

        routingHasData = true;
        const nameFile = getUniqueFileName(plat, dateForFilename, '.xlsx', seenFileNames);
        const driverName = route.driverName || route.assigneeName || '-';
        const wb = XLSX.utils.book_new();

        appendDeliveryListSheet(wb, plat, driverName, processedTripsToRender, isDetailView, t);

        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        routingZip.file(nameFile, excelBuffer);
      }

      if (routingHasData) {
        masterHasData = true;
        const routingZipBlob = await routingZip.generateAsync({ type: 'blob' });
        masterZip.file(`${zipFolderName}.zip`, routingZipBlob);
      }
    }

    if (!masterHasData) {
      toastError(t('common.toast.error', { err: t('common.no_data') }));
      setIsDownloading(false);
      return;
    }

    const masterContent = await masterZip.generateAsync({ type: 'blob' });
    const finalFileName = fileNamePrefix
      ? `${t('delivery.delivery_list')} (${fileNamePrefix}) - ${dateForFilename} - ${locationName}.zip`
      : `${t('delivery.delivery_list')} - ${dateForFilename} - ${locationName}.zip`;

    triggerDownload(masterContent, finalFileName);
    toastSuccess(t('common.toast.success'));
  } catch (err) {
    toastError(t('common.toast.error', { err: err.message }), err);
  } finally {
    setIsDownloading(false);
  }
};
