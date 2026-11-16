'use client';

import Button from '@/components/button/Button';
import InformationButton from '@/components/button/InformationButton';
import CustomDatePicker from '@/components/CustomDatePicker';
import { useLanguage } from '@/context/LanguageContext';
import { getTasks, getTrash } from '@/lib/api/mileapp';
import { getCachedHubs } from '@/lib/localStorageHandler';
import { generateTaskCountWorkbook } from '@/lib/reportGenerators/reports';
import { toastError, toastSuccess, toastWarning } from '@/lib/toast';
import { formatDateUniversal, formatLongDate, toApiDateString } from '@/lib/utils';
import { useEffect, useState } from 'react';
import * as XLSX from 'xlsx-js-style';

export default function TaskCountReport() {
  const { t, localeCode } = useLanguage();
  const [hubs, setHubs] = useState([]);
  const [selectedHubs, setSelectedHubs] = useState([]);
  const [isCustomMode, setIsCustomMode] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState(new Date());
  const [startDate, setStartDate] = useState(new Date(new Date().setHours(0, 0, 0, 0)));
  const [endDate, setEndDate] = useState(new Date(new Date().setHours(23, 59, 59, 999)));
  const [isLoading, setIsLoading] = useState(false);

  useEffect(() => {
    const loadHubs = async () => {
      try {
        let cached = getCachedHubs().allHub;
        setHubs(cached || []);
        setSelectedHubs((cached || []).map((h) => h._id));
      } catch (err) {
        toastError(t('common.toast.error', { err: err.message }), err);
      }
    };
    loadHubs();
  }, [t]);

  const isAllSelected = selectedHubs.length > 0 && selectedHubs.length === hubs.length;

  const handleSelectAll = () => {
    if (isAllSelected) setSelectedHubs([]);
    else setSelectedHubs(hubs.map((h) => h._id));
  };

  const handleToggleHub = (hubId) => {
    if (selectedHubs.includes(hubId)) {
      setSelectedHubs((prev) => prev.filter((id) => id !== hubId));
    } else {
      setSelectedHubs((prev) => [...prev, hubId]);
    }
  };

  const handleProcess = async () => {
    if (selectedHubs.length === 0) {
      toastError(t('report.toast.select_hub'));
      return;
    }

    let calcStart, calcEnd;

    if (!isCustomMode) {
      const y = selectedMonth.getFullYear();
      const m = selectedMonth.getMonth();
      calcStart = new Date(Date.UTC(y, m, 24, 1, 34, 0));
      calcEnd = new Date(Date.UTC(y, m + 1, 24, 1, 34, 0));
    } else {
      if (!startDate || !endDate || startDate > endDate) {
        toastError(t('common.invalid_date'));
        return;
      }
      calcStart = startDate;
      calcEnd = endDate;
    }

    setIsLoading(true);
    try {
      const chunks = [];
      let currentStart = new Date(calcStart);
      const finalEnd = new Date(calcEnd);

      while (currentStart < finalEnd) {
        let currentEnd = new Date(currentStart);
        currentEnd.setDate(currentStart.getDate() + 30);
        if (currentEnd > finalEnd) currentEnd = new Date(finalEnd);

        chunks.push({
          from: new Date(currentStart),
          to: new Date(currentEnd),
        });
        currentStart = new Date(currentEnd.getTime() + 1000);
      }

      let allFetchedTasks = [];
      let isHitLimit = false;

      for (const chunk of chunks) {
        const timeFrom = toApiDateString(chunk.from);
        const timeTo = toApiDateString(chunk.to);

        const response = await getTasks({
          status: 'DONE,ONGOING,UNASSIGNED',
          timeFrom,
          timeTo,
          isShowAll: true,
        });

        const chunkData = Array.isArray(response) ? response : response?.data || [];
        if (chunkData.length === 1000) isHitLimit = true;
        allFetchedTasks = [...allFetchedTasks, ...chunkData];
      }

      let filteredTrashTasks = [];
      try {
        const trashResponse = await getTrash();
        const trashData = Array.isArray(trashResponse) ? trashResponse : trashResponse?.data || [];

        trashData.forEach((item) => {
          if (item.dataType === 'tasks' && item.data) {
            try {
              const parsedData = typeof item.data === 'string' ? JSON.parse(item.data) : item.data;
              if (parsedData.createdTime && parsedData.hubId) {
                const createdDate = new Date(parsedData.createdTime);
                if (createdDate >= calcStart && createdDate <= calcEnd) {
                  if (selectedHubs.includes(parsedData.hubId)) {
                    parsedData._isFromTrash = true;
                    parsedData.startTime = parsedData.createdTime;
                    filteredTrashTasks.push(parsedData);
                  }
                }
              }
            } catch (e) {
              toastError(t('common.toast.error', { err: e.message }), e);
            }
          }
        });
      } catch (err) {
        toastError(t('common.toast.error', { err: err.message }), err);
        toastWarning(t('report.toast.failed_trash'));
      }

      const finalDataToProcess = [...allFetchedTasks, ...filteredTrashTasks];
      if (finalDataToProcess.length === 0) {
        toastWarning(t('common.no_data'));
        setIsLoading(false);
        return;
      }

      if (isHitLimit) toastWarning(t('report.toast.limit_warning'));

      const selectedHubsInfo = selectedHubs.map((id) => {
        const found = hubs.find((h) => h._id === id);
        return { id, name: found ? found.name : 'Unknown Hub' };
      });

      let fileName = '';
      if (!isCustomMode) {
        const startLabel = formatLongDate(calcStart, localeCode);
        const endLabel = formatLongDate(calcEnd, localeCode);
        fileName = `Task Counter (Periode ${startLabel} - ${endLabel}).xlsx`;
      } else {
        const fStart = formatDateUniversal(calcStart, 'DD.MM.YYYY');
        const fEnd = formatDateUniversal(calcEnd, 'DD.MM.YYYY');
        fileName = `Task Counter (${fStart} - ${fEnd}).xlsx`;
      }

      const wb = generateTaskCountWorkbook(
        finalDataToProcess,
        selectedHubsInfo,
        calcStart,
        calcEnd,
        t
      );
      XLSX.writeFile(wb, fileName);
      toastSuccess(t('common.toast.success'));
    } catch (err) {
      toastError(t('common.toast.error', { err: err.message }), err);
    } finally {
      setIsLoading(false);
    }
  };

  const yHelp = selectedMonth.getFullYear();
  const mHelp = selectedMonth.getMonth();
  const startHelp = new Date(yHelp, mHelp, 24, 8, 34, 0);
  const endHelp = new Date(yHelp, mHelp + 1, 24, 8, 34, 0);

  return (
    <div className="w-full max-w-5xl mx-auto p-4 sm:p-6 animate-in fade-in duration-300">
      <h1 className="text-3xl font-bold mb-6 text-slate-900 dark:text-slate-100">
        {t('report.task_count.title')}
      </h1>

      <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-200 dark:border-slate-700 p-6 flex flex-col gap-6 transition-colors">
        <div className="flex items-center gap-2 mb-2 pb-4 border-b border-gray-200 dark:border-slate-700">
          <label className="flex items-center gap-2 cursor-pointer text-sm font-bold select-none">
            <input
              type="checkbox"
              checked={isCustomMode}
              onChange={(e) => setIsCustomMode(e.target.checked)}
              disabled={isLoading}
              className="w-4 h-4 rounded bg-white dark:bg-slate-900 border-gray-300 dark:border-slate-600 cursor-pointer"
            />
            {t('report.task_count.custom_time')}
          </label>
        </div>

        {!isCustomMode ? (
          <div className="flex flex-col items-center w-full">
            <div className="flex flex-col max-w-sm w-full">
              <label className="text-sm font-semibold text-gray-600 dark:text-slate-300 mb-2">
                {t('report.task_count.date_range')}
              </label>
              <CustomDatePicker
                selected={selectedMonth}
                onChange={(date) => setSelectedMonth(date)}
                showMonthYearPicker
                dateFormat="MMMM yyyy"
                className="w-full"
                disabled={isLoading}
              />
            </div>
            <p className="text-xs text-gray-500 dark:text-slate-400 mt-2 italic leading-relaxed flex flex-wrap items-center justify-center gap-1.5">
              <strong className="text-slate-700 dark:text-slate-300 whitespace-nowrap">
                {formatLongDate(startHelp, localeCode)} 08:34 WIB
              </strong>
              <span>{t('common.to')}</span>
              <strong className="text-slate-700 dark:text-slate-300 whitespace-nowrap">
                {formatLongDate(endHelp, localeCode)} 08:34 WIB
              </strong>
              <InformationButton infoText={t('report.tooltip.cut_off')} />
            </p>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="flex flex-col">
              <label className="text-sm font-semibold text-gray-600 dark:text-slate-300 mb-2">
                {t('common.start_time')}
              </label>
              <CustomDatePicker
                selected={startDate}
                onChange={(date) => setStartDate(date)}
                showTimeInput
                dateFormat="dd/MM/yyyy HH:mm:ss"
                className="w-full"
                disabled={isLoading}
              />
            </div>
            <div className="flex flex-col">
              <label className="text-sm font-semibold text-gray-600 dark:text-slate-300 mb-2">
                {t('common.finish_time')}
              </label>
              <CustomDatePicker
                selected={endDate}
                onChange={(date) => setEndDate(date)}
                showTimeInput
                dateFormat="dd/MM/yyyy HH:mm:ss"
                className="w-full"
                disabled={isLoading}
              />
            </div>
          </div>
        )}

        <div className="flex flex-col border border-gray-300 dark:border-slate-600 rounded-lg overflow-hidden transition-colors mt-2">
          <div className="bg-gray-50 dark:bg-slate-700/50 p-4 border-b border-gray-300 dark:border-slate-600 flex justify-between items-center transition-colors">
            <span className="font-semibold text-sm text-slate-800 dark:text-slate-200">
              {t('report.task_count.choose_hub')}
            </span>
            <label className="flex items-center gap-2 cursor-pointer text-sm font-bold select-none">
              <input
                type="checkbox"
                checked={isAllSelected}
                onChange={handleSelectAll}
                disabled={isLoading}
                className="w-4 h-4 rounded bg-white dark:bg-slate-900 border-gray-300 dark:border-slate-600 cursor-pointer"
              />
              {t('common.all')}
            </label>
          </div>

          <div className="p-4 max-h-64 overflow-y-auto grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 bg-white dark:bg-slate-800">
            {hubs.map((hub) => (
              <label
                key={hub._id}
                className="flex items-center gap-3 cursor-pointer text-sm text-slate-700 dark:text-slate-300 hover:bg-sky-50 dark:hover:bg-slate-700/50 p-2 rounded transition-colors select-none border border-transparent hover:border-sky-100 dark:hover:border-slate-600"
              >
                <input
                  type="checkbox"
                  checked={selectedHubs.includes(hub._id)}
                  onChange={() => handleToggleHub(hub._id)}
                  disabled={isLoading}
                  className="w-4 h-4 rounded text-sky-600 focus:ring-sky-500 bg-white dark:bg-slate-900 border-gray-300 dark:border-slate-600 cursor-pointer"
                />
                <span className="truncate">{hub.name}</span>
              </label>
            ))}
          </div>
        </div>

        <div className="flex justify-end mt-2">
          <Button
            onClick={handleProcess}
            isLoading={isLoading}
            text={t('common.download')}
            width="w-full md:w-auto"
            disabled={selectedHubs.length === 0}
          />
        </div>
      </div>
    </div>
  );
}
