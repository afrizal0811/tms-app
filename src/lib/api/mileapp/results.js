import { getLocalStorage } from '@/lib/localStorageHandler';
import { formatDateUniversal } from '../../utils';
import { apiFetch } from '../base';

export async function getResults({ deliveryDate, dateFrom, dateTo, deliveryDateObj } = {}) {
  const finalDeliveryDate =
    deliveryDate || (deliveryDateObj ? formatDateUniversal(deliveryDateObj, 'YYYY-MM-DD') : null);
  const params = new URLSearchParams();
  const { storedLocation } = getLocalStorage();
  if (finalDeliveryDate) params.append('deliveryDate', finalDeliveryDate);
  if (dateFrom) params.append('dateFrom', dateFrom);
  if (dateTo) params.append('dateTo', dateTo);
  if (storedLocation) params.append('hubId', storedLocation);

  return await apiFetch(
    `/api/mileapp/routing-results?${params.toString()}`,
    'Gagal mengambil data results'
  );
}

export async function getResult(id) {
  if (!id) {
    throw new Error('ID result harus disertakan');
  }

  return await apiFetch(
    `/api/mileapp/routing-results/${id}`,
    `Gagal mengambil data result dengan ID ${id}`
  );
}

export async function syncRoutingResults({ deliveryDate, hubId, force = false, created_by, updated_by } = {}) {
  const { storedUser } = getLocalStorage();
  const userName = storedUser?.name || null;
  return await apiFetch('/api/mileapp/routing-results', 'Gagal sinkronisasi data routing', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      deliveryDate,
      hubId,
      force,
      created_by: created_by || userName,
      updated_by: updated_by || userName,
    }),
  });
}

export async function updateRoutingResult(id, { deliveryDate, updatedBy } = {}) {
  if (!id) {
    throw new Error('ID result harus disertakan');
  }
  const { storedUser } = getLocalStorage();
  const userName = updatedBy || storedUser?.name || null;
  return await apiFetch(`/api/mileapp/routing-results/${id}`, 'Gagal memperbarui data routing', {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ deliveryDate, updatedBy: userName }),
  });
}
