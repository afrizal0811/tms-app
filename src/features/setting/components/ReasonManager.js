'use client';

import Accordion from '@/components/Accordion';
import Button from '@/components/button/Button';
import Dropdown from '@/components/dropdown/Dropdown';
import HighlightText from '@/components/HighlightText';
import ConfirmModal from '@/components/modal/ConfirmModal';
import SearchBar from '@/components/SearchBar';
import { deleteReason, postReason, updateReason } from '@/lib/api/mileapp';
import { toastError, toastSuccess } from '@/lib/toast';
import { useMemo, useState } from 'react';
import Card from './Card';
import CustomTable from './CustomTable';

export default function ReasonManager({ reasons, onRefresh, isReadOnly, translate }) {
  const [newReason, setNewReason] = useState('');
  const [newPic, setNewPic] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [isAddOpen, setIsAddOpen] = useState(false);
  const [deleteConfig, setDeleteConfig] = useState({ isOpen: false, id: null, name: null });

  const picOptions = useMemo(() => {
    const set = new Set((reasons || []).map((r) => r.pic?.trim()).filter(Boolean));
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [reasons]);

  const filteredReasons = useMemo(() => {
    const list = reasons || [];
    const q = searchQuery.toLowerCase().trim();
    const filtered = !q
      ? list
      : list.filter((r) => {
          const cat = (r.category || '').toLowerCase();
          const pic = (r.pic || '').toLowerCase();
          const reason = (r.reasons || '').toLowerCase();
          return cat.includes(q) || pic.includes(q) || reason.includes(q);
        });

    return filtered.map((item, idx) => ({
      ...item,
      _index: idx + 1,
    }));
  }, [reasons, searchQuery]);

  const handleAddReason = async () => {
    if (!newReason.trim() || !newPic || isReadOnly) return;
    const resolvedCat = ['Customer', 'Other'].includes(newPic) ? 'External' : 'Internal';
    try {
      await postReason(newReason, newPic, resolvedCat);
      setNewReason('');
      setNewPic('');
      setIsAddOpen(false);
      toastSuccess(translate('common.toast.success'));
      await onRefresh();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    }
  };

  const handleUpdateReason = async (id, editValues) => {
    if (!editValues.reasons?.trim() || !editValues.pic || isReadOnly) return;
    const resolvedCat = ['Customer', 'Other'].includes(editValues.pic) ? 'External' : 'Internal';
    try {
      await updateReason(id, editValues.reasons, editValues.pic, resolvedCat);
      toastSuccess(translate('common.toast.success'));
      await onRefresh();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
      throw e;
    }
  };

  const confirmDeleteReason = async () => {
    const targetId = deleteConfig.id;
    setDeleteConfig({
      isOpen: false,
      id: null,
      name: null,
      reason: '',
    });
    if (!targetId) return;
    try {
      await deleteReason(targetId);
      toastSuccess(translate('common.toast.success'));
      await onRefresh();
    } catch (e) {
      toastError(translate('common.toast.error', { err: e.message }), e);
    }
  };

  const columns = useMemo(
    () => [
      {
        header: '#',
        field: '_index',
        headerClassName: 'w-10 sm:w-12 text-center',
        align: 'center',
        cellClassName: 'text-center',
        render: (item) => (
          <span className="text-[10px] md:text-xs font-medium text-slate-500 dark:text-slate-400">
            {item._index}
          </span>
        ),
      },
      {
        header: translate('common.category'),
        field: 'category',
        headerClassName: 'w-24 sm:w-32 md:w-48',
        render: (item) => (
          <span className="text-[10px] md:text-xs font-medium text-slate-700 dark:text-slate-300">
            <HighlightText text={item.category} highlight={searchQuery} />
          </span>
        ),
      },
      {
        header: 'PIC',
        field: 'pic',
        headerClassName: 'w-24 sm:w-32 md:w-48',
        render: (item) => (
          <span className="text-[10px] md:text-xs font-medium text-slate-700 dark:text-slate-300">
            <HighlightText text={item.pic} highlight={searchQuery} />
          </span>
        ),
        renderEdit: (value, onChange, onSave, item, setValues) => (
          <Dropdown
            options={picOptions.map((opt) => ({ label: opt, value: opt }))}
            value={value}
            onChange={(selectedPic) => {
              const autoCat = ['Customer', 'Other'].includes(selectedPic) ? 'External' : 'Internal';
              if (typeof setValues === 'function') {
                setValues({ pic: selectedPic, category: autoCat });
              } else {
                onChange(selectedPic);
              }
            }}
            getLabel={(val) => val || '-'}
            size="sm"
            className="w-full min-w-0"
          />
        ),
      },
      {
        header: translate('setting.tab.general.reasons_title'),
        field: 'reasons',
        render: (item) => (
          <span className="font-semibold text-slate-700 dark:text-slate-200 text-[10px] md:text-sm">
            <HighlightText text={item.reasons} highlight={searchQuery} />
          </span>
        ),
        renderEdit: (value, onChange, onSave) => (
          <input
            type="text"
            value={value}
            onChange={(e) => onChange(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && onSave()}
            className="w-full min-w-0 px-2 py-1 text-[10px] md:text-sm bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-600 rounded outline-none"
            autoFocus
          />
        ),
      },
    ],
    [picOptions, searchQuery, translate]
  );
  const msgParts = translate('common.modal.delete_message', { text: '|||' }).split('|||');

  return (
    <Card>
      <ConfirmModal
        isOpen={deleteConfig.isOpen}
        onCancel={() => setDeleteConfig({ isOpen: false, id: null, name: null })}
        onConfirm={confirmDeleteReason}
        title={translate('common.modal.delete_title', {
          text: translate('setting.tab.general.reasons_title'),
        })}
        message={
          <span>
            {msgParts[0]}
            <strong>{deleteConfig.reason}</strong>
            {msgParts[1]}
          </span>
        }
      />

      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-4 border-b border-gray-100 dark:border-slate-800 pb-3">
        <div>
          <h2 className="text-lg font-bold text-gray-900 dark:text-slate-200">
            {translate('setting.tab.general.reasons_title')}
          </h2>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
            {translate('setting.tab.general.reasons_subtitle')}
          </p>
        </div>
        <div className="w-full sm:w-64 shrink-0">
          <SearchBar
            value={searchQuery}
            onChange={setSearchQuery}
            placeholder={translate('common.search')}
            size="md"
            width="w-full"
            tooltip={`${translate('common.category')}, PIC, ${translate('setting.tab.general.reasons_title')}`}
          />
        </div>
      </div>

      {!isReadOnly && (
        <Accordion
          isFloating
          isOpen={isAddOpen}
          onToggle={setIsAddOpen}
          title={
            <div className="flex items-center gap-2">
              <span className="flex items-center justify-center w-5 h-5 rounded-full bg-sky-100 dark:bg-sky-900/40 text-sky-600 dark:text-sky-400 font-bold text-xs">
                +
              </span>
              <span className="text-xs sm:text-sm font-semibold text-slate-700 dark:text-slate-200">
                {translate('common.button.btn_add')}{' '}
                {translate('setting.tab.general.reasons_title')}
              </span>
            </div>
          }
          defaultOpen={false}
          className="mb-4"
        >
          <div className="flex flex-col sm:flex-row gap-2 w-full">
            <input
              type="text"
              value={newReason}
              onChange={(e) => setNewReason(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && newReason.trim() && newPic) {
                  handleAddReason();
                }
              }}
              placeholder={translate('setting.tab.general.add_placeholder')}
              className="flex-1 min-w-0 w-full px-3 py-2 text-xs sm:text-sm text-slate-900 dark:text-slate-100 border border-gray-300 dark:border-slate-700 rounded-md outline-none focus:border-sky-500 dark:focus:border-slate-500 focus:ring-1 focus:ring-sky-500 dark:focus:ring-slate-500 bg-white dark:bg-slate-800"
            />
            <div className="w-full sm:w-56 shrink-0 min-w-0">
              <Dropdown
                options={picOptions.map((opt) => ({ label: opt, value: opt }))}
                value={newPic}
                onChange={setNewPic}
                getLabel={(val) => val || translate('setting.tab.general.select_pic')}
                size="md"
                className="w-full min-w-0"
              />
            </div>
            <div className="w-full sm:w-auto shrink-0">
              <Button
                onClick={handleAddReason}
                disabled={!newReason.trim() || !newPic}
                size="md"
                text={translate('common.button.btn_add')}
                width="w-full sm:w-auto px-5"
              />
            </div>
          </div>
        </Accordion>
      )}

      <CustomTable
        data={filteredReasons}
        columns={columns}
        isReadOnly={isReadOnly}
        containerHeight="h-[350px]"
        emptyMessage={translate('common.no_data')}
        translate={translate}
        onSave={handleUpdateReason}
        onDelete={(item) => setDeleteConfig({ isOpen: true, id: item.id, reason: item.reasons })}
      />
    </Card>
  );
}
