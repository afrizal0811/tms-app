import { normalizeDeliveryStatus } from '@/lib/deliveryStatus';
import {
  formatDateUniversal,
  getBasePlate,
  isEmpty,
  normalizeEmail,
  parseCustomerString,
} from '@/lib/utils';
import {
  applyJitter,
  assignPlanSequences,
  getAngle,
  parseTaskCoordinates,
} from '../components/MapViewSection';

export const getGroupKey = (item) => {
  const driver = item?.driver || item?.assignee || 'N/A';
  const plat =
    getBasePlate(item?.plat || item?.vehicleName) || item?.plat || item?.vehicleName || '';
  return `${driver}_${plat}`;
};

export const getGroupTasks = (data, groupValue) =>
  data
    ? data
        .filter((d) => (d.groupKey || getGroupKey(d)) === groupValue)
        .map((t) => ({
          ...t,
          longlat: t.longlat || t.rawTask?.longlat || t.coordinate || t.location || '',
        }))
    : [];

export const calculateDashboard = (tasksArray, driverMap, isIndonesian) => {
  const keys = [
    'unassigned',
    'ongoing',
    'done',
    'flowDelivery',
    'flowReDelivery',
    'manualAssign',
    'diffDay',
    'success',
    'partial',
    'cancel',
    'pending',
    'pendingGr',
  ];

  const res = {
    totalTasks: tasksArray?.length || 0,
    totalDry: 0,
    totalFrozen: 0,
    assignedDry: 0,
    assignedFrozen: 0,
    assignedTasks: 0,
    taskId: null,
    ongoingList: [],
    manualAssignList: [],
    unassignedList: [],
    diffDayList: [],
    successList: [],
    partialList: [],
    cancelList: [],
    pendingList: [],
    pendingGrList: [],
  };

  keys.forEach((k) => {
    res[k] = 0;
    res[`${k}Dry`] = 0;
    res[`${k}Frozen`] = 0;
  });

  if (isEmpty(tasksArray)) return res;

  const inc = (key, isDry, isFrozen) => {
    res[key]++;
    if (isDry) res[`${key}Dry`]++;
    if (isFrozen) res[`${key}Frozen`]++;
  };

  for (const task of tasksArray) {
    const {
      name: customerName,
      invoiceNumber,
      truncateInvoice,
      isTruncated,
    } = parseCustomerString(task.customerOrder) || 'N/A';

    let rawAssignee = 'N/A';
    if (task.assignee && task.assignee !== '') {
      rawAssignee =
        typeof task.assignee === 'string' ? task.assignee.split(',')[0].trim() : task.assignee;
    }
    let finalAssignee =
      task.driverName || driverMap.get(normalizeEmail(rawAssignee)) || rawAssignee;
    if (finalAssignee === 'N/A') finalAssignee = '-';

    const typeStorage = (task.typeStorage || '').toUpperCase();
    const isDry = typeStorage === 'DRY';
    const isFrozen = typeStorage === 'FROZEN';

    if (isDry) res.totalDry++;
    if (isFrozen) res.totalFrozen++;

    const isAssigned = task.status !== 'UNASSIGNED';
    if (isAssigned) {
      if (isDry) res.assignedDry++;
      if (isFrozen) res.assignedFrozen++;
    }

    const baseData = {
      customer: customerName,
      flow: task.flow || 'N/A',
      soNumber: invoiceNumber || '-',
      truncateSoNumber: truncateInvoice,
      isTruncated,
      driver: finalAssignee,
      taskId: task._id || '-',
    };

    const normDelivery = normalizeDeliveryStatus(task.statusDelivery || task.status);
    const isDeliveryStatus = ['SUCCESS', 'PARTIAL', 'CANCEL', 'PENDING', 'PENDING_GR'].includes(
      normDelivery
    );

    if (task.status === 'DONE' || isDeliveryStatus) {
      inc('done', isDry, isFrozen);
      const statusDelivery = normDelivery;

      if (statusDelivery === 'SUCCESS') {
        res.successList.push(baseData);
        inc('success', isDry, isFrozen);
      } else if (statusDelivery === 'PARTIAL') {
        res.partialList.push(baseData);
        inc('partial', isDry, isFrozen);
      } else if (statusDelivery === 'CANCEL') {
        res.cancelList.push(baseData);
        inc('cancel', isDry, isFrozen);
      } else if (statusDelivery === 'PENDING') {
        res.pendingList.push(baseData);
        inc('pending', isDry, isFrozen);
      } else if (statusDelivery === 'PENDING_GR') {
        res.pendingGrList.push(baseData);
        inc('pendingGr', isDry, isFrozen);
      }
    } else if (task.status === 'ONGOING') {
      res.ongoingList.push(baseData);
      inc('ongoing', isDry, isFrozen);
    } else if (task.status === 'UNASSIGNED') {
      res.unassignedList.push(baseData);
      inc('unassigned', isDry, isFrozen);
    }

    const manualCategory = !task.routePlannedOrder || !task.eta || !task.etd;
    if (manualCategory && isAssigned) {
      res.manualAssignList.push(baseData);
      inc('manualAssign', isDry, isFrozen);
    }

    const flow = task.flow || '';
    if (flow === 'Delivery') {
      inc('flowDelivery', isDry, isFrozen);
    } else if (flow.includes('Re Delivery')) {
      inc('flowReDelivery', isDry, isFrozen);
    }

    if ((task.status === 'DONE' || isDeliveryStatus) && task.startTime && task.doneTime) {
      const startDateWIB = formatDateUniversal(task.startTime);
      const doneDateWIB = formatDateUniversal(task.doneTime);
      if (startDateWIB && doneDateWIB && startDateWIB !== doneDateWIB) {
        const startDate = new Date(task.startTime);
        const doneDate = new Date(task.doneTime);
        const diffInDays = Math.ceil((doneDate - startDate) / (1000 * 60 * 60 * 24));
        res.diffDayList.push({
          ...baseData,
          doneDateDisplay: `${doneDateWIB} (${isIndonesian ? 'H+' : 'D+'}${diffInDays})`,
        });
        inc('diffDay', isDry, isFrozen);
      }
    }
  }

  const sortList = (list) => list.sort((a, b) => a.driver.localeCompare(b.driver));
  res.unassignedList.sort((a, b) => a.flow.localeCompare(b.flow));
  [
    'ongoingList',
    'manualAssignList',
    'diffDayList',
    'successList',
    'partialList',
    'cancelList',
    'pendingList',
    'pendingGrList',
  ].forEach((k) => sortList(res[k]));

  res.assignedTasks = res.done + res.ongoing;
  return res;
};

export const renderGroupToJpeg = async (groupValue, data, L, toJpeg) => {
  const groupTasks = getGroupTasks(data, groupValue);
  const validTasks = parseTaskCoordinates(groupTasks).sort(
    (a, b) => (a.roSequence || 9999) - (b.roSequence || 9999)
  );
  const numberedTasks = assignPlanSequences(validTasks, 'roSequence');
  const jitteredTasks = applyJitter(numberedTasks);
  const allCoordinates = jitteredTasks.map((t) => [t.lat, t.lng]);

  if (allCoordinates.length === 0) return null;

  const rawBounds = L.latLngBounds(allCoordinates);
  const sw = rawBounds.getSouthWest();
  const ne = rawBounds.getNorthEast();

  const midLat = (sw.lat + ne.lat) / 2;
  const latDiff = Math.max(0.005, Math.abs(ne.lat - sw.lat));
  const lonDiff = Math.max(0.005, Math.abs(ne.lng - sw.lng));

  const latMeters = latDiff * 111320;
  const lonMeters = lonDiff * (111320 * Math.cos((midLat * Math.PI) / 180));

  const BASE_DIM = 1200;
  let exportWidth, exportHeight;
  const ratio = Math.max(0.05, latMeters / lonMeters);
  if (lonMeters >= latMeters) {
    exportWidth = BASE_DIM;
    exportHeight = Math.max(200, Math.min(BASE_DIM, Math.round(BASE_DIM * ratio)));
  } else {
    exportHeight = BASE_DIM;
    exportWidth = Math.max(200, Math.min(BASE_DIM, Math.round(BASE_DIM * (1 / ratio))));
  }

  const container = document.createElement('div');
  container.style.position = 'fixed';
  container.style.top = '0px';
  container.style.left = '0px';
  container.style.width = `${exportWidth}px`;
  container.style.height = `${exportHeight}px`;
  container.style.zIndex = '-99999';
  container.style.pointerEvents = 'none';
  container.style.opacity = '1';
  container.style.overflow = 'hidden';
  container.style.backgroundColor = '#ffffff';
  document.body.appendChild(container);

  const map = L.map(container, {
    zoomControl: false,
    attributionControl: false,
    fadeAnimation: false,
    zoomAnimation: false,
  });

  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    crossOrigin: 'anonymous',
  }).addTo(map);

  const padY = 50;
  const padX = 50;

  map.fitBounds(rawBounds, {
    padding: [padY, padX],
    maxZoom: 18,
    animate: false,
  });

  const hubTask = jitteredTasks.find((t) => t.type === 'HUB_START');
  const routeTasks = jitteredTasks.filter((t) => t.type !== 'HUB_START' && t.type !== 'HUB_END');
  const segments = [];

  if (hubTask && hubTask.lat && hubTask.lng) {
    let prevPoint = [hubTask.lat, hubTask.lng];
    routeTasks.forEach((t) => {
      const nextPoint = [t.lat, t.lng];
      segments.push({ start: prevPoint, end: nextPoint });
      prevPoint = nextPoint;
    });
    segments.push({ start: prevPoint, end: [hubTask.lat, hubTask.lng] });
  }

  segments.forEach((seg) => {
    const { start, end } = seg;
    L.polyline([start, end], {
      color: '#FF0000',
      weight: 4,
      opacity: 0.75,
      dashArray: '8, 12',
    }).addTo(map);

    const angle = getAngle(start[0], start[1], end[0], end[1]);
    const startArrowLat = start[0] + (end[0] - start[0]) * 0.1;
    const startArrowLng = start[1] + (end[1] - start[1]) * 0.1;
    const midLat = (start[0] + end[0]) / 2;
    const midLng = (start[1] + end[1]) / 2;
    const endLat = start[0] + (end[0] - start[0]) * 0.9;
    const endLng = start[1] + (end[1] - start[1]) * 0.9;

    const arrowIcon = L.divIcon({
      className: 'bg-transparent border-none',
      html: `<div style="transform: rotate(${-angle}deg); color: #FF0000; font-size: 16px; font-weight: 900; filter: drop-shadow(1px 1px 0px white); display: flex; justify-content: center; align-items: center; width: 100%; height: 100%;">➤</div>`,
      iconSize: [20, 20],
      iconAnchor: [10, 10],
    });

    [
      [startArrowLat, startArrowLng],
      [midLat, midLng],
      [endLat, endLng],
    ].forEach((pos) => {
      L.marker(pos, { icon: arrowIcon, zIndexOffset: -50, interactive: false }).addTo(map);
    });
  });

  const exportIcons = {
    circle: (content, bgClass, textClass, borderClass) =>
      L.divIcon({
        className: 'bg-transparent border-none',
        html: `<div class="${bgClass} w-7 h-7 rounded-full border-2 ${borderClass} shadow-md flex items-center justify-center text-white font-bold ${textClass} z-50 relative box-border overflow-hidden">${content}</div>`,
        iconSize: [28, 28],
        iconAnchor: [14, 14],
      }),
  };

  jitteredTasks.forEach((task) => {
    const seqNum = task.roSequence ?? '?';
    let finalContent = seqNum;
    let markerBgClass = 'bg-[#D97706]';
    let textSizeClass = 'text-xs';
    const isHub = task.type === 'HUB_START' || task.type === 'HUB_END';

    if (isHub) {
      markerBgClass = 'bg-[#000000]';
      finalContent = 'HUB';
      textSizeClass = 'text-[9px] tracking-tighter';
    } else if (task.isManualAssign) {
      markerBgClass = 'bg-[#DC2626]';
    }

    if (isEmpty(finalContent) || finalContent === '?') textSizeClass = 'text-sm';

    const icon = exportIcons.circle(finalContent, markerBgClass, textSizeClass, 'border-white');

    L.marker([task.lat, task.lng], {
      icon,
      zIndexOffset: isHub ? 200 : 100,
    }).addTo(map);
  });

  await new Promise((resolve) => setTimeout(resolve, 1000));

  let rawMapDataUrl = null;
  try {
    rawMapDataUrl = await toJpeg(container, {
      quality: 0.8,
      backgroundColor: '#ffffff',
      width: exportWidth,
      height: exportHeight,
      pixelRatio: 1,
    });
  } finally {
    map.remove();
    if (container && container.parentNode) {
      container.parentNode.removeChild(container);
    }
  }

  if (!rawMapDataUrl) return null;

  const legendTasks = numberedTasks.filter((t) => t.type !== 'HUB_START' && t.type !== 'HUB_END');

  if (legendTasks.length === 0) return rawMapDataUrl;

  const isLandscape = exportWidth >= exportHeight;

  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');

  const mapImg = await new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = rawMapDataUrl;
  });

  const ROW_HEIGHT = 56;
  const PADDING = 24;
  const HEADER_HEIGHT = 44;
  const GAP = 20;

  let totalCanvasWidth = exportWidth;
  let totalCanvasHeight = exportHeight;
  let legendX = 0;
  let legendY = 0;
  let colWidth = 0;
  let numRows = 5;

  if (isLandscape) {
    numRows = Math.min(5, Math.max(1, legendTasks.length));
    const minColWidth = 260;
    const availableWidth = exportWidth - PADDING * 2;
    const maxPossibleCols = Math.max(1, Math.floor((availableWidth + GAP) / (minColWidth + GAP)));
    const requiredCols = Math.ceil(legendTasks.length / numRows);
    const numCols = Math.min(maxPossibleCols, requiredCols);
    numRows = Math.ceil(legendTasks.length / numCols);

    const legendHeight = PADDING * 2 + HEADER_HEIGHT + numRows * ROW_HEIGHT;
    totalCanvasHeight = exportHeight + legendHeight;
    legendX = PADDING;
    legendY = exportHeight + PADDING;
    colWidth = (availableWidth - (numCols - 1) * GAP) / numCols;
  } else {
    const availableVerticalHeight = exportHeight - PADDING * 2 - HEADER_HEIGHT;
    const maxRowsPerCol = Math.max(1, Math.floor(availableVerticalHeight / ROW_HEIGHT));
    const numCols = Math.ceil(legendTasks.length / maxRowsPerCol);
    numRows = numCols > 1 ? maxRowsPerCol : Math.max(1, legendTasks.length);
    const targetColWidth = 320;
    const legendPanelWidth = numCols * targetColWidth + (numCols - 1) * GAP + PADDING * 2;
    totalCanvasWidth = exportWidth + legendPanelWidth;
    legendX = exportWidth + PADDING;
    legendY = PADDING;
    colWidth = targetColWidth;
  }

  canvas.width = totalCanvasWidth;
  canvas.height = totalCanvasHeight;

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, totalCanvasWidth, totalCanvasHeight);

  ctx.drawImage(mapImg, 0, 0, exportWidth, exportHeight);

  if (isLandscape) {
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(0, exportHeight);
    ctx.lineTo(exportWidth, exportHeight);
    ctx.stroke();
  } else {
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(exportWidth, 0);
    ctx.lineTo(exportWidth, exportHeight);
    ctx.stroke();
  }

  ctx.fillStyle = '#0f172a';
  ctx.font = 'bold 20px sans-serif';
  ctx.fillText('KETERANGAN', legendX, legendY + 20);

  const itemsStartY = legendY + HEADER_HEIGHT;

  const truncateText = (text, maxWidth) => {
    let display = text;
    while (ctx.measureText(display).width > maxWidth && display.length > 3) {
      display = display.slice(0, -2) + '…';
    }
    return display;
  };

  legendTasks.forEach((task, index) => {
    const col = Math.floor(index / numRows);
    const row = index % numRows;

    const itemX = legendX + col * (colWidth + GAP);
    const itemY = itemsStartY + row * ROW_HEIGHT;

    const seqNum = task.roSequence ?? index + 1;
    const isHub =
      task.type === 'HUB_START' ||
      task.type === 'HUB_END' ||
      task.type === 'HUB' ||
      task.customerName === 'HUB';
    let badgeBg = '#D97706';
    if (isHub) {
      badgeBg = '#000000';
    } else if (task.isManualAssign) {
      badgeBg = '#DC2626';
    }

    const radius = 13;
    const badgeCenterX = itemX + radius;
    const badgeCenterY = itemY + 16;

    ctx.beginPath();
    ctx.arc(badgeCenterX, badgeCenterY, radius, 0, Math.PI * 2);
    ctx.fillStyle = badgeBg;
    ctx.fill();
    ctx.lineWidth = 1.5;
    ctx.strokeStyle = '#ffffff';
    ctx.stroke();

    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 11px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(String(seqNum), badgeCenterX, badgeCenterY);

    const rawCust = task.originalCustomerString || task.customerName || '';
    const { name } = parseCustomerString(rawCust);
    const baseCustName = name || task.customerName || 'Customer';
    const custName = task.isManualAssign ? `[M] ${baseCustName}` : baseCustName;

    const rawAddress =
      task.rawTask?.address ||
      task.rawTask?.destination?.address ||
      task.address ||
      task.rawTask?.locationName ||
      '-';

    const textLeft = itemX + radius * 2 + 8;
    const maxTextWidth = colWidth - (radius * 2 + 12);

    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';

    ctx.fillStyle = '#1e293b';
    ctx.font = 'bold 13px sans-serif';
    ctx.fillText(truncateText(custName, maxTextWidth), textLeft, itemY + 14);

    ctx.fillStyle = '#64748b';
    ctx.font = 'normal 11px sans-serif';
    ctx.fillText(truncateText(rawAddress, maxTextWidth), textLeft, itemY + 28);
  });

  return canvas.toDataURL('image/jpeg', 0.8);
};
