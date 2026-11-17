import { getLocalStorage } from '@/lib/localStorageHandler';
import { apiFetch } from '../base';

// vehicles
export async function getVehicles(customHubId) {
  const { storedLocation } = getLocalStorage();
  const hubId = customHubId !== undefined ? customHubId : storedLocation;
  const params = new URLSearchParams();
  if (hubId && hubId !== 'all') params.append('hubId', hubId);
  return await apiFetch(
    `/api/mileapp/vehicles?${params.toString()}`,
    'Gagal mengambil data kendaraan dari database'
  );
}

export async function postVehicles(hubIds = [], customUpdateBy) {
  const { storedUser } = getLocalStorage();
  const updateBy = customUpdateBy !== undefined ? customUpdateBy : storedUser?.name;
  return await apiFetch('/api/mileapp/vehicles', 'Gagal sinkronisasi data kendaraan', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ hubIds, updateBy }),
  });
}

export async function updateVehicle(payload) {
  const { storedUser } = getLocalStorage();
  return await apiFetch('/api/mileapp/vehicles', 'Gagal mengubah data kendaraan', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      updateBy: storedUser?.name,
      ...payload,
    }),
  });
}

export async function patchVehicles(updates, customUpdateBy) {
  const { storedUser } = getLocalStorage();
  const updateBy = customUpdateBy !== undefined ? customUpdateBy : storedUser?.name;
  return await apiFetch('/api/mileapp/vehicles', 'Gagal update kendaraan', {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ updates, updateBy }),
  });
}

export async function deleteVehicle({ id, ids, vehicle_id, hubId }) {
  const params = new URLSearchParams();
  if (ids && Array.isArray(ids) && ids.length > 0) {
    params.append('ids', ids.join(','));
  } else if (id) {
    params.append('id', id);
  }
  if (vehicle_id) params.append('vehicle_id', vehicle_id);
  if (hubId) params.append('hubId', hubId);
  return await apiFetch(`/api/mileapp/vehicles?${params.toString()}`, 'Gagal menghapus kendaraan', {
    method: 'DELETE',
  });
}

export async function patchVehicleMceasy(activeHubId, storedLocationName, explicitMatches = null) {
  if (explicitMatches) {
    await patchVehicles(explicitMatches);
    return { success: true };
  }

  const { getVehiclesStatuses, patchVehicle: patchMceasyVehicle } = await import('../mceasy');
  const { getBasePlate } = await import('../../utils');
  const { getDrivers } = await import('./drivers');

  const rawLoc = (storedLocationName || '').replace(/^Hub\s+/i, '').trim();
  const targetLocation = rawLoc.toUpperCase() === 'GIIC' ? 'Cikarang' : rawLoc;
  const [mceasyData, dbVehicles, dbDrivers] = await Promise.all([
    getVehiclesStatuses({ location: targetLocation }),
    getVehicles(activeHubId),
    getDrivers(activeHubId),
  ]);

  const mismatched = [];
  const matched = [];
  const autoUpdates = [];

  const list = Array.isArray(dbVehicles) ? dbVehicles : [];
  const mcList = Array.isArray(mceasyData) ? mceasyData : [];
  const driversList = Array.isArray(dbDrivers) ? dbDrivers : [];

  const driverMap = new Map();
  driversList.forEach((d) => {
    driverMap.set(d.id, d);
    if (d.email) driverMap.set(d.email.toLowerCase().trim(), d);
  });

  list.forEach((v) => {
    const vBasePlat = getBasePlate(v.license_number)?.replace(/\s+/g, '')?.toUpperCase();
    let mcMatch = null;
    if (v.vms_id) {
      mcMatch = mcList.find((mc) => String(mc.vehicleId || mc.id) === String(v.vms_id));
    }
    if (!mcMatch) {
      mcMatch = mcList.find(
        (mc) => mc.licensePlate?.replace(/\s+/g, '')?.toUpperCase() === vBasePlat
      );
    }

    if (!mcMatch) return;

    const mcPlat = mcMatch.licensePlate?.replace(/\s+/g, '')?.toUpperCase();
    const vehicleKey = mcMatch.vehicleId || mcMatch.id;
    const isPlatMatch = mcPlat === vBasePlat;

    const tmsDriver = v.driver_id
      ? driverMap.get(v.driver_id)
      : v.driver_email
        ? driverMap.get(v.driver_email.toLowerCase().trim())
        : null;
    const tmsDriverVmsId = tmsDriver?.vms_id ? String(tmsDriver.vms_id) : null;
    const vmsDriverId = mcMatch.driverId ? String(mcMatch.driverId) : null;
    const isDriverMatch = !tmsDriverVmsId || vmsDriverId === tmsDriverVmsId;

    const maCleanName = (tmsDriver?.name || v.driver_name || '')
      .replace(/^'[^']+'\s*/, '')
      .replace(/^(dry|frozen)[\s\-_]*/i, '')
      .replace(/\s+/g, ' ')
      .trim();

    if (isPlatMatch && isDriverMatch) {
      matched.push({
        id: v.id,
        vehicle_id: v.vehicle_id,
        imei: mcMatch.imei,
        vms_id: vehicleKey,
      });
    } else if (!tmsDriverVmsId && !isDriverMatch) {
      mismatched.push({
        id: v.id,
        vehicle_id: v.vehicle_id,
        imei: mcMatch.imei,
        maName: tmsDriver?.name || v.driver_name || '-',
        maNameClean: maCleanName || '-',
        maPlat: v.license_number,
        maPlatBase: vBasePlat,
        mcName: mcMatch.driver1?.fullname || '-',
        mcPlat: mcMatch.licensePlate || '-',
        mcVehicleId: vehicleKey,
        vmsDriverId: mcMatch.driverId,
      });
    } else {
      autoUpdates.push({
        mcVehicleId: vehicleKey,
        targetPlat: !isPlatMatch ? getBasePlate(v.license_number) : undefined,
        targetDriverId: !isDriverMatch ? tmsDriverVmsId : undefined,
      });
      matched.push({
        id: v.id,
        vehicle_id: v.vehicle_id,
        imei: mcMatch.imei,
        vms_id: vehicleKey,
      });
    }
  });

  if (autoUpdates.length > 0) {
    const targetDriverIds = new Set(
      autoUpdates
        .map((u) => u.targetDriverId)
        .filter((id) => id && id !== '-')
        .map(String)
    );

    const vehiclesToClear = new Set();
    autoUpdates.forEach((u) => {
      if (u.mcVehicleId && u.targetDriverId !== undefined) {
        vehiclesToClear.add(u.mcVehicleId);
      }
    });
    mcList.forEach((mc) => {
      const vId = mc.vehicleId || mc.id;
      if (vId && mc.driverId && targetDriverIds.has(String(mc.driverId))) {
        vehiclesToClear.add(vId);
      }
    });

    for (const vehicleId of vehiclesToClear) {
      try {
        const clearPayload = new URLSearchParams({ driver1Id: '-' });
        await patchMceasyVehicle(vehicleId, clearPayload.toString());
      } catch (err) {
        console.warn(`[Auto-Sync VMS] Unassign driver vehicle ${vehicleId} error:`, err.message);
      }
    }

    for (const update of autoUpdates) {
      try {
        const payload = new URLSearchParams();
        if (update.targetPlat) {
          payload.append('licensePlate', update.targetPlat);
        }
        if (update.targetDriverId !== undefined) {
          payload.append('driver1Id', update.targetDriverId || '-');
        }
        await patchMceasyVehicle(update.mcVehicleId, payload.toString());
      } catch (err) {
        console.warn(`[Auto-Sync VMS] Update vehicle ${update.mcVehicleId} error:`, err.message);
      }
    }
  }

  return { success: mismatched.length === 0, matched, mismatched };
}

export async function configureVehicles(items, hubId) {
  const { storedUser, storedLocation } = getLocalStorage();
  const targetHubId = hubId || storedLocation;
  return await apiFetch('/api/mileapp/vehicles', 'Gagal menyimpan konfigurasi kendaraan', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      items,
      hubId: targetHubId,
      updateBy: storedUser?.name,
    }),
  });
}

export async function getVehicleStatus() {
  const { storedLocation } = getLocalStorage();
  const params = new URLSearchParams();
  params.append('status', 'true');
  if (storedLocation) params.append('hubId', storedLocation);

  return await apiFetch(
    `/api/mileapp/vehicles?${params.toString()}`,
    'Gagal mengambil status sync kendaraan'
  );
}

// vehicle type
export async function createVehicleType(name, customUpdateBy) {
  const { storedUser } = getLocalStorage();
  const updateBy = customUpdateBy !== undefined ? customUpdateBy : storedUser?.name;
  return await apiFetch('/api/mileapp/vehicles/types', 'Gagal menambah tipe kendaraan', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, updateBy }),
  });
}

export async function getVehicleTypes() {
  const { storedLocation } = getLocalStorage();
  const params = new URLSearchParams();
  if (storedLocation) params.append('hubId', storedLocation);

  return await apiFetch(
    `/api/mileapp/vehicles/types?${params.toString()}`,
    'Gagal memproses perhitungan types truck storage'
  );
}

export async function updateVehicleType(id, name, customUpdateBy) {
  const { storedUser } = getLocalStorage();
  const updateBy = customUpdateBy !== undefined ? customUpdateBy : storedUser?.name;
  return await apiFetch('/api/mileapp/vehicles/types', 'Gagal mengubah tipe kendaraan', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ id, name, updateBy }),
  });
}

export async function deleteVehicleType(id) {
  const params = new URLSearchParams();
  if (id) params.append('id', id);

  return await apiFetch(
    `/api/mileapp/vehicles/types?${params.toString()}`,
    'Gagal menghapus tipe kendaraan',
    {
      method: 'DELETE',
    }
  );
}

export async function getVehicleExternals({ startDate, endDate }) {
  const { storedLocation } = getLocalStorage();
  const params = new URLSearchParams();
  if (storedLocation) params.append('hubId', storedLocation);
  if (startDate) params.append('startDate', startDate);
  if (endDate) params.append('endDate', endDate);

  return await apiFetch(
    `/api/mileapp/vehicles/external?${params.toString()}`,
    'Gagal mengambil data kendaraan eksternal'
  );
}

export async function postVehicleExternal(payload) {
  return await apiFetch(
    '/api/mileapp/vehicles/external',
    'Gagal menyimpan data kendaraan eksternal',
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }
  );
}

export async function deleteVehicleExternal(payload) {
  return await apiFetch(
    '/api/mileapp/vehicles/external',
    'Gagal menghapus data kendaraan eksternal',
    {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }
  );
}

export async function getVehicleUsages(params = {}) {
  const q = new URLSearchParams();
  if (params.hubId) q.append('hubId', params.hubId);
  if (params.date) q.append('date', params.date);
  if (params.dateFrom) q.append('dateFrom', params.dateFrom);
  if (params.dateTo) q.append('dateTo', params.dateTo);

  return await apiFetch(
    `/api/mileapp/vehicles/usages?${q.toString()}`,
    'Gagal mengambil data Vehicle Usage dari database'
  );
}
