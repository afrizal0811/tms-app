// File: src/lib/utils.js

// Mengubah format input tanggal (Date atau string) menjadi pola string tertentu (default: YYYY-MM-DD)
export function formatDateUniversal(dateInput, pattern = 'YYYY-MM-DD') {
  if (!dateInput) return '-';

  let extraDays = 0;
  let processedInput = dateInput;

  if (typeof dateInput === 'string') {
    processedInput = dateInput.replace(
      /(^|\s|T)(\d+):(\d{2})(?::(\d{2}))?/,
      (match, prefix, hoursStr, minutesStr, secondsStr) => {
        const hours = parseInt(hoursStr, 10);

        if (!isNaN(hours) && hours >= 24) {
          extraDays = Math.floor(hours / 24);
          const modHours = String(hours % 24).padStart(2, '0');
          const secPart = secondsStr !== undefined ? `:${secondsStr}` : '';
          return `${prefix}${modHours}:${minutesStr}${secPart}`;
        }

        return match;
      }
    );
  }

  let date;

  if (typeof processedInput === 'string') {
    const dmyMatch = processedInput.match(
      /^(\d{2})[\/\-](\d{2})[\/\-](\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/
    );

    if (dmyMatch) {
      const [, dd, mm, yyyy, HH = '00', min = '00', ss = '00'] = dmyMatch;
      date = new Date(`${yyyy}-${mm}-${dd}T${HH}:${min}:${ss}`);
    } else if (/^\d{1,2}:\d{2}(?::\d{2})?$/.test(processedInput.trim())) {
      date = new Date(`1970-01-01T${processedInput.trim()}`);
    } else {
      date = new Date(processedInput);
    }
  } else {
    date = new Date(processedInput);
  }

  if (isNaN(date.getTime())) return '-';

  const fullYear = date.getFullYear();

  const map = {
    YYYY: fullYear,
    YY: String(fullYear).slice(-2),
    MM: String(date.getMonth() + 1).padStart(2, '0'),
    DD: String(date.getDate()).padStart(2, '0'),
    HH: String(date.getHours()).padStart(2, '0'),
    mm: String(date.getMinutes()).padStart(2, '0'),
    ss: String(date.getSeconds()).padStart(2, '0'),
  };

  const result = pattern.replace(/YYYY|YY|MM|DD|HH|mm|ss/g, (matched) => map[matched]);

  return extraDays > 0 ? `${result} (+${extraDays})` : result;
}
// Menentukan tanggal otomatis (H-1) dengan penyesuaian jika hari tersebut adalah Minggu (mundur ke Sabtu)
export function calculateTargetDates(selectedDateStr) {
  if (!selectedDateStr) {
    const today = new Date();
    selectedDateStr = formatDateUniversal(today);
  }

  const selectedDate = new Date(selectedDateStr);

  selectedDate.setDate(selectedDate.getDate() - 1);

  if (selectedDate.getDay() === 0) {
    selectedDate.setDate(selectedDate.getDate() - 1);
  }

  const targetDateStr = formatDateUniversal(selectedDate);

  return {
    dateFrom: targetDateStr,
    dateTo: targetDateStr,
  };
}

// Mengonversi total menit menjadi string jam:menit dengan format text excel ('HH:mm)
export function formatMinutesToHHMM(totalMinutes, second = false) {
  if (totalMinutes == null || isNaN(totalMinutes) || totalMinutes < 0) {
    return '-';
  }

  const totalSeconds = Math.round(totalMinutes * 60);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  const formattedHours = String(hours).padStart(2, '0');
  const formattedMinutes = String(minutes).padStart(2, '0');

  if (second) {
    const formattedSeconds = String(seconds).padStart(2, '0');
    return `${formattedHours}:${formattedMinutes}:${formattedSeconds}`;
  }

  return `${formattedHours}:${formattedMinutes}`;
}

export function formatDurationHHMMSS(totalMinutes) {
  if (totalMinutes == null || totalMinutes === '') return '-';
  if (typeof totalMinutes === 'string' && totalMinutes.includes(':')) {
    const parts = totalMinutes.split(':');
    if (parts.length === 2) {
      return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}:00`;
    }
    return totalMinutes;
  }
  return formatMinutesToHHMM(Number(totalMinutes), true);
}

// Menghitung selisih waktu dalam satuan menit antara dua timestamp
export function calculateMinuteDifference(time1, time2) {
  if (!time1 || !time2) return null;

  const date1 = new Date(time1);
  const date2 = new Date(time2);

  if (isNaN(date1.getTime()) || isNaN(date2.getTime())) return null;

  const m1 = Math.floor(date1.getTime() / 60000);
  const m2 = Math.floor(date2.getTime() / 60000);

  return Math.abs(m2 - m1);
}

// Memecah string data pelanggan menjadi Nama, ID, dan Lokasi berdasarkan pola pemisah tertentu
export function parseCustomerString(fullString) {
  if (!fullString || typeof fullString !== 'string') {
    return {
      name: '',
      id: null,
      location: null,
      fullCustomerName: '',
      invoiceNumber: '',
      truncateInvoice: '',
      isTruncated: false,
    };
  }

  const parts = fullString.split(/\s+-\s+/);

  let idIndex = parts.findIndex((p) => /^C\d+/.test(p) && !p.includes(' '));
  if (idIndex === -1 && parts.length >= 2) {
    idIndex = 1;
  }

  let id = idIndex !== -1 ? parts[idIndex] : '';

  let location = null;
  let invoiceNumber = '';
  let truncateInvoice = '';
  let isTruncated = false;
  if (parts.length > 2 && idIndex !== -1 && idIndex < parts.length - 1) {
    const rawLocation = parts[parts.length - 1];
    const commaSplit = rawLocation.split(',');
    location = commaSplit[0].trim();
    if (commaSplit.length > 1) {
      invoiceNumber = commaSplit
        .slice(1)
        .map((s) => s.trim())
        .filter(Boolean)
        .join(', ');
    }
  } else if (parts.length === 2 && id.includes(',')) {
    const commaSplit = id.split(',');
    id = commaSplit[0].trim();
    invoiceNumber = commaSplit
      .slice(1)
      .map((s) => s.trim())
      .filter(Boolean)
      .join(', ');
  }
  if (invoiceNumber) {
    const orderParts = invoiceNumber.split(',').filter(Boolean);
    isTruncated = orderParts.length > 1;
    if (isTruncated) {
      truncateInvoice = `${orderParts[0].trim()} (+${orderParts.length - 1})`;
    } else if (orderParts.length === 1) {
      truncateInvoice = orderParts[0].trim();
    }
  }

  location = location !== null ? location : '';

  let name = '';
  if (idIndex > 0) {
    name = parts.slice(0, idIndex).join(' - ');
  } else if (parts[0] && parts[0] !== id) {
    name = parts[0];
  }

  const fullCustomerName = id !== '' || location !== '' ? `${name} - ${id} - ${location}` : name;

  return { name, id, location, fullCustomerName, invoiceNumber, isTruncated, truncateInvoice };
}

export function parseProductItem(item) {
  const rawTitle = String(item?.title || item?.inventory_name || '').trim();
  let inventoryId = item?.inventory_id ? String(item.inventory_id).trim() : null;
  let inventoryName = rawTitle;

  if (!inventoryId && rawTitle) {
    const match = rawTitle.match(/^(?:Barang\s+)?(\d+)\s*-\s*(.+)$/i);
    if (match) {
      inventoryId = match[1].trim();
      inventoryName = match[2].trim();
    }
  }

  const toTextOrNull = (val) => {
    if (val == null) return null;
    const str = String(val).trim();
    return str || null;
  };

  const toNumOrNull = (val) => {
    if (val == null || val === '') return null;
    const num = Number(val);
    return Number.isNaN(num) ? null : num;
  };

  return {
    invoice_number: toTextOrNull(item?.caption ?? item?.invoice_number),
    inventory_id: inventoryId || null,
    inventory_name: inventoryName || '-',
    uom: toTextOrNull(item?.content ?? item?.uom),
    qty: toNumOrNull(item?.qtyProcessed ?? item?.qty),
    weight: toNumOrNull(item?.weight),
    volume: toNumOrNull(item?.volume),
  };
}

export function formatListProductsForApi(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  return rows.map((row) => ({
    title: row.inventory_id ? `${row.inventory_id}-${row.inventory_name}` : row.inventory_name,
    caption: row.invoice_number || '',
    content: row.uom || '',
    qtyProcessed: row.qty ?? 0,
    weight: row.weight ?? 0,
    volume: row.volume ?? 0,
    inventoryId: row.inventory_id || null,
    inventoryName: row.inventory_name,
    invoiceNumber: row.invoice_number || null,
    uom: row.uom || null,
    qty: row.qty ?? 0,
  }));
}

export function getStorageType(input) {
  if (!input) return '-';
  if (Array.isArray(input) || typeof input === 'string') {
    const text = (Array.isArray(input) ? input.join(' ') : input).toUpperCase();
    if (text.includes('FROZEN') || text.includes('FRZ')) return 'Frozen';
    if (text.includes('DRY')) return 'Dry';
    return '-';
  }
  const storageStr = input?.storage || input?.typeStorage || '';
  if (storageStr.toUpperCase().includes('FROZEN') || storageStr.toUpperCase().includes('FRZ'))
    return 'Frozen';
  if (storageStr.toUpperCase().includes('DRY')) return 'Dry';

  const typeStr = input?.type || input?.mappedType || '';
  if (typeStr.toUpperCase().includes('FROZEN') || typeStr.toUpperCase().includes('FRZ'))
    return 'Frozen';
  if (typeStr.toUpperCase().includes('DRY')) return 'Dry';

  const nameStr = input?.name || input?.driverName || input?.description || '';
  if (nameStr.toUpperCase().includes("'FRZ'") || nameStr.toUpperCase().includes('FROZEN'))
    return 'Frozen';
  if (nameStr.toUpperCase().includes("'DRY'") || nameStr.toUpperCase().includes('DRY'))
    return 'Dry';
  return '-';
}

// Memformat koordinat
export function formatCoordinates(coordsString) {
  const coords = parseCoordinates(coordsString);
  if (!coords) return null;
  return `${coords.lat.toFixed(6)}, ${coords.lon.toFixed(6)}`;
}

// Mengubah koordinat menjadi object
export function parseCoordinates(coordString) {
  if (typeof coordString !== 'string' || !coordString.includes(',')) return null;
  const [latStr, lonStr] = coordString.split(',');
  const lat = parseFloat(latStr);
  const lon = parseFloat(lonStr);
  if (isNaN(lat) || isNaN(lon)) return null;
  return { lat, lon };
}

// Menghitung jarak antar koordinat
export function getDistance(coordsString1, coordsString2) {
  const c1 = parseCoordinates(coordsString1);
  const c2 = parseCoordinates(coordsString2);

  if (!c1 || !c2) return null;

  const toRad = (value) => (Number(value) * Math.PI) / 180;
  const R = 6371000;

  const dLat = toRad(c2.lat - c1.lat);
  const dLon = toRad(c2.lon - c1.lon);
  const rLat1 = toRad(c1.lat);
  const rLat2 = toRad(c2.lat);

  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.sin(dLon / 2) * Math.sin(dLon / 2) * Math.cos(rLat1) * Math.cos(rLat2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return Math.round(R * c);
}

// Menormalkan email
export const normalizeEmail = (email) => {
  if (typeof email !== 'string' || !email) {
    return null;
  }
  return email.toLowerCase().trim();
};

// Menambahkan jam ke Date
function addHours(date, hours) {
  const newDate = new Date(date);
  newDate.setTime(newDate.getTime() + hours * 60 * 60 * 1000);
  return newDate;
}

// Menghasilkan rentang waktu laporan
export function calculateStartFinishDates(selectedDateStr) {
  const selectedDate = new Date(selectedDateStr + 'T12:00:00');

  const timeToEnd = new Date(selectedDate);
  timeToEnd.setHours(16, 59, 59, 999);
  const timeTo = formatDateUniversal(timeToEnd, 'YYYY-MM-DD HH:mm:ss');

  const timeFromStart = new Date(selectedDate);
  timeFromStart.setDate(timeFromStart.getDate() - 1);
  if (timeFromStart.getDay() === 0) {
    timeFromStart.setDate(timeFromStart.getDate() - 2);
  }
  timeFromStart.setHours(-7, 0, 0, 0);
  const timeFrom = formatDateUniversal(timeFromStart, 'YYYY-MM-DD HH:mm:ss');

  return {
    timeFrom,
    timeTo,
  };
}

// Mengubah timestamp UTC ke UTC+7
export function parseAndShiftToUTC7(timestampStr) {
  if (!timestampStr) return null;
  const str =
    timestampStr instanceof Date
      ? timestampStr.toISOString()
      : typeof timestampStr === 'number'
        ? new Date(timestampStr).toISOString()
        : String(timestampStr);
  let utcTimestamp = str.replace(' ', 'T');
  if (
    !utcTimestamp.endsWith('Z') &&
    !utcTimestamp.substring(10).includes('+') &&
    !utcTimestamp.substring(10).includes('-')
  ) {
    utcTimestamp += 'Z';
  }

  const utcDate = new Date(utcTimestamp);
  if (isNaN(utcDate.getTime())) return null;

  return addHours(utcDate, 7);
}

// Memformat timestamp UTC+7
export function formatUTC7(timestampStr, pattern = 'YYYY-MM-DD') {
  const date = parseAndShiftToUTC7(timestampStr);
  if (!date) return null;

  const map = {
    YYYY: date.getUTCFullYear(),
    MM: String(date.getUTCMonth() + 1).padStart(2, '0'),
    DD: String(date.getUTCDate()).padStart(2, '0'),
    HH: String(date.getUTCHours()).padStart(2, '0'),
    mm: String(date.getUTCMinutes()).padStart(2, '0'),
    ss: String(date.getUTCSeconds()).padStart(2, '0'),
  };

  const result = pattern.replace(/YYYY|MM|DD|HH|mm|ss/g, (matched) => map[matched]);
  return result;
}

// Mengecek hari Minggu
export const isDateSunday = (dateStr) => {
  const date = new Date(dateStr.replace(/-/g, '/'));
  return date.getDay() === 0;
};

// Mengubah detik ke MM:SS
export const formatTimer = (seconds) => {
  const m = Math.floor(seconds / 60)
    .toString()
    .padStart(2, '0');
  const s = (seconds % 60).toString().padStart(2, '0');
  return `${m}:${s}`;
};

// Memformat tanggal panjang
export function formatLongDate(dateInput, language = 'id-ID', withDate = true) {
  if (!dateInput) return '-';
  let config = {
    month: 'long',
    year: 'numeric',
  };
  if (withDate) config.day = 'numeric';
  let d = dateInput instanceof Date ? dateInput : null;
  if (!d && typeof dateInput === 'string') {
    const isoMatch = dateInput.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
      d = new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
    } else {
      const parsed = new Date(dateInput);
      if (!isNaN(parsed.getTime())) d = parsed;
    }
  } else if (!d) {
    const parsed = new Date(dateInput);
    if (!isNaN(parsed.getTime())) d = parsed;
  }
  if (!d || isNaN(d.getTime())) return '-';
  return d.toLocaleDateString(language, config);
}

// Mengecek nilai kosong
export function isEmpty(value) {
  if (Array.isArray(value)) return value.length === 0;
  if (value && typeof value === 'object') {
    return Object.keys(value).length === 0;
  }
  return (
    !value ||
    value === undefined ||
    value === null ||
    value === '' ||
    value === false ||
    value === 0 ||
    value === '0' ||
    value === '-'
  );
}

// Mengubah Date ke format API
export function toApiDateString(dateInput) {
  const date = dateInput instanceof Date ? dateInput : new Date(dateInput);
  if (isNaN(date.getTime())) return null;
  return date.toISOString().slice(0, 19).replace('T', ' ');
}

// Mengubah teks ke Title Case
export const capitalizeText = (text = '') =>
  text.toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase());

// Mengubah teks ke Proper Case
export const ProperCaseText = (str) => {
  if (!str) return '';
  return str.replace(/([A-Z])/g, ' $1').replace(/^./, (s) => s.toUpperCase());
};

// Menghasilkan tanggal besok atau minggu depan
export const tomorrowDate = (isTomorrow = true) => {
  const date = new Date();
  const day = isTomorrow ? 1 : 7;
  date.setDate(date.getDate() + day);
  if (date.getDay() === 0) {
    date.setDate(date.getDate() + 1);
  }

  if (!isTomorrow && new Date().getDay() === 6) {
    date.setDate(date.getDate() + 1);
  }

  return date;
};

// Pola format SO yang valid: 2-5 huruf + 4 digit + strip + 6 digit
const SO_PATTERN = /^[A-Z]{2,5}\d{4}-\d{6}$/i;

export function standardizeSo(rawSo) {
  const clean = String(rawSo).replace(/[^a-zA-Z0-9-]/g, '');
  const matchStrict = clean.match(/^([a-zA-Z]{2,5})(\d{4})-(\d+)$/);
  if (matchStrict)
    return `${matchStrict[1].toUpperCase()}${matchStrict[2]}-${matchStrict[3].padStart(6, '0')}`;

  const matchBypass = clean.match(/^([a-zA-Z]{2,5})(\d{4})-(\d+)(.+)$/);
  if (matchBypass)
    return `${matchBypass[1].toUpperCase()}${matchBypass[2]}-${matchBypass[3].padStart(6, '0')}${matchBypass[4]}`;

  return clean;
}

export function isValidSo(so) {
  return SO_PATTERN.test(so);
}

export function isBypassSo(so) {
  return /^[A-Z]{2,5}\d{4}-\d{6}.+$/i.test(so);
}

export function checkInvalidSo(rawSo, isBadCust = false) {
  if (isBadCust) return true;
  return !isValidSo(standardizeSo(rawSo));
}

// Cek banyak SO sekaligus (dipisah koma) — invalid kalau salah satu aja invalid
export function checkInvalidSoList(rawSoString, isBadCust = false) {
  if (isBadCust) return true;
  const list = String(rawSoString || '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
  return list.some((so) => checkInvalidSo(so));
}

// Mengambil plat utama
export function getBasePlate(plat) {
  if (!plat) return '';
  const platStr = String(plat).trim();
  const parts = platStr.split(/\s+/);
  return parts.length > 3 ? parts.slice(0, 3).join(' ') : platStr;
}

export function getDeliveryDateFromRouting(isoString) {
  if (!isoString) return null;
  const str = String(isoString).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  const date = new Date(isoString);
  const wibMs = date.getTime() + 7 * 60 * 60 * 1000;
  const wibDate = new Date(wibMs);
  const routingDay = wibDate.getUTCDay();

  let offsetDays = 1;
  if (routingDay === 6) offsetDays = 2;

  const deliveryMs = wibMs + offsetDays * 24 * 60 * 60 * 1000;
  return new Date(deliveryMs).toISOString().split('T')[0];
}

// Mengubah nilai persentase menjadi warna heatmap pastel dalam format HEX
export function heatMap(pctInput) {
  if (pctInput === null || pctInput === undefined || pctInput === '') return null;
  let val = parseFloat(pctInput);
  if (isNaN(val)) return null;
  if (val > 0 && val <= 1) val = val * 100;
  const ratio = Math.max(0, Math.min(100, val)) / 100;
  let r, g, b;
  if (ratio <= 0.5) {
    const t = ratio / 0.5;
    r = 255;
    g = Math.round(255 * t);
    b = 0;
  } else {
    const t = (ratio - 0.5) / 0.5;
    r = Math.round(255 * (1 - t));
    g = Math.round(176 + (255 - 176) * (1 - t));
    b = 0;
  }
  return [r, g, b]
    .map((v) =>
      Math.round(v * 0.45 + 255 * 0.55)
        .toString(16)
        .padStart(2, '0')
    )
    .join('')
    .toUpperCase();
}

// Mengecek apakah tanggal lebih kecil dari hari ini
export function isPastDate(dateStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return new Date(y, m - 1, d) < today;
}

// Mengubah datetime API menjadi Date
export function parseApiDateString(dateStr) {
  if (!dateStr) return null;
  let isoStr = dateStr.toString().replace(' ', 'T');
  if (!isoStr.endsWith('Z') && !isoStr.includes('+')) isoStr += 'Z';
  const d = new Date(isoStr);
  return isNaN(d.getTime()) ? null : d;
}

// Menentukan grup urutan plat
function getSortGroup(plat) {
  if (!plat) return 1;
  const platUpper = plat.toUpperCase();
  if (platUpper.includes('DM')) return 3;
  if (platUpper.includes('SEWA')) return 2;
  return 1;
}

// Mengurutkan data kendaraan
export function sortRows(rows, platKey, driverKey) {
  return rows.sort((a, b) => {
    const rankA = getSortGroup(a[platKey]);
    const rankB = getSortGroup(b[platKey]);
    if (rankA !== rankB) return rankA - rankB;
    return (a[driverKey] || '').localeCompare(b[driverKey] || '');
  });
}

// Hitung manual jarak customer akhir ke hub
export function calculateReturnHubDistance(taskList, hubCoordsStr) {
  if (!hubCoordsStr || !taskList || taskList.length === 0) return 0;

  const sortedTasks = [...taskList].sort((a, b) => {
    const roA = a.roSequence ?? a.routePlannedOrder ?? 9999;
    const roB = b.roSequence ?? b.routePlannedOrder ?? 9999;
    return roA - roB;
  });

  const lastTask = sortedTasks[sortedTasks.length - 1];

  const targetCoord =
    lastTask.expectedCoordinate ||
    lastTask.expectedCoord ||
    lastTask.doneCoordinate ||
    lastTask.doneCoord ||
    null;

  if (targetCoord) {
    const rawDistance = getDistance(targetCoord, hubCoordsStr);
    if (rawDistance !== null) {
      return Math.round(rawDistance * 1.3);
    }
  }

  return 0;
}

export const getBaseVehicleType = (typeStr, knownTypes = []) => {
  if (!typeStr) return '';
  const upperStr = typeStr.toUpperCase();
  const sortedKnown = [...knownTypes].sort((a, b) => b.length - a.length);
  const match = sortedKnown.find((k) => upperStr.includes(k.toUpperCase()));
  if (match) return match.toUpperCase();

  const parts = upperStr.split('-');
  const typeParts = parts.filter((p) => !['FROZEN', 'DRY'].includes(p));
  if (typeParts.length === 0) return typeStr;
  let base = typeParts[0];
  if (typeParts.length > 1 && typeParts[1] === 'LONG') base = `${base}-LONG`;
  return base;
};

export function formatOdometer(value, locale = 'id-ID') {
  if (isEmpty(value)) {
    return '-';
  }
  return Number(value).toLocaleString(locale, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}
