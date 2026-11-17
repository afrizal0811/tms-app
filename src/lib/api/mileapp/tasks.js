import { getLocalStorage } from '@/lib/localStorageHandler';
import { apiFetch } from '../base';
import { fields } from './fields';

export async function getTasks({
  status,
  deliveryDate,
  dateFrom,
  dateTo,
  timeFrom,
  timeTo,
  isNeedFields = true,
  isShowAll = false,
  hubId,
  source,
} = {}) {
  const { storedLocation } = getLocalStorage();
  const targetHubId = hubId || (!isShowAll ? storedLocation : null);
  const tasksFields = fields.tasks.join(',');
  const params = new URLSearchParams();
  params.append('timeBy', 'startTime');
  params.append('limit', 10000);
  if (targetHubId && targetHubId !== 'all') params.append('hubId', targetHubId);
  if (status) params.append('status', status);
  if (deliveryDate) params.append('deliveryDate', deliveryDate);
  if (dateFrom) params.append('dateFrom', dateFrom);
  if (dateTo) params.append('dateTo', dateTo);
  if (timeFrom) params.append('timeFrom', timeFrom);
  if (timeTo) params.append('timeTo', timeTo);
  if (isNeedFields && tasksFields) params.append('fields', tasksFields);
  if (source) params.append('source', source);

  return await apiFetch(`/api/mileapp/tasks?${params.toString()}`, 'Gagal mengambil data tasks');
}

export async function getTask(id, isAllHub) {
  const params = new URLSearchParams();
  const { storedLocation } = getLocalStorage();
  if (!isAllHub && storedLocation) params.append('hubId', storedLocation);

  if (!id) {
    throw new Error('ID task harus disertakan');
  }

  const result = await apiFetch(
    `/api/mileapp/tasks/${id}?${params.toString()}`,
    `Gagal mengambil data task dengan ID ${id}`
  );

  return result;
}

export async function getExtraTasks(id) {
  if (!id) {
    throw new Error('ID task harus disertakan');
  }

  return await apiFetch(
    `/api/mileapp/tasks/${id}?vendorOnly=1`,
    `Gagal mengambil riwayat task dengan ID ${id}`
  );
}

export async function syncTasks({
  deliveryDate,
  hubId,
  timeFrom,
  timeTo,
  force = false,
  created_by,
  updated_by,
} = {}) {
  const { storedLocation, storedUser } = getLocalStorage();
  const targetHubId = hubId || storedLocation;
  const userName = updated_by || created_by || storedUser?.name || null;

  return await apiFetch('/api/mileapp/tasks', 'Gagal sinkronisasi data tasks', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      deliveryDate,
      hubId: targetHubId,
      timeFrom,
      timeTo,
      force,
      created_by: userName,
      updated_by: userName,
    }),
  });
}

export async function autoUpdateTasks({ hubId } = {}) {
  const { storedLocation, storedUser } = getLocalStorage();
  const targetHubId = hubId || storedLocation;
  const userName = storedUser?.name || 'System';

  return await apiFetch('/api/mileapp/tasks', 'Gagal auto-update data tasks', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      autoUpdate: true,
      hubId: targetHubId,
      updated_by: userName,
    }),
  });
}
