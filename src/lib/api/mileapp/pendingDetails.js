import { apiFetch } from '../base';

export async function getPendingDetails(startDate, endDate) {
  const params = new URLSearchParams();
  if (startDate) params.append('startDate', startDate);
  if (endDate) params.append('endDate', endDate);

  return await apiFetch(
    `/api/mileapp/pendings?${params.toString()}`,
    'Gagal mengambil data detail pending'
  );
}

export async function getPendingDetailByTaskId(taskId) {
  if (!taskId) return null;
  const params = new URLSearchParams();
  params.append('taskId', taskId);

  return await apiFetch(
    `/api/mileapp/pendings?${params.toString()}`,
    'Gagal mengambil data detail pending'
  );
}

export async function postPendingDetail(data) {
  return await apiFetch('/api/mileapp/pendings', 'Gagal menyimpan detail pending', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
}

export async function deletePendingDetail(taskId) {
  const params = new URLSearchParams();
  if (taskId) params.append('taskId', taskId);

  return await apiFetch(
    `/api/mileapp/pendings?${params.toString()}`,
    'Gagal menghapus detail pending',
    {
      method: 'DELETE',
    }
  );
}
