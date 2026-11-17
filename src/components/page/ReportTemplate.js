'use client';

import Button from '@/components/button/Button';
import InformationButton from '@/components/button/InformationButton';
import ToggleButton from '@/components/button/ToggleButton';
import CustomDatePicker from '@/components/CustomDatePicker';
import ConfirmModal from '@/components/modal/ConfirmModal';
import { useLanguage } from '@/context/LanguageContext';
import { tomorrowDate } from '@/lib/utils';
import { useEffect, useRef, useState } from 'react';

const ModeRadioInput = ({ checked, disabled }) => (
  <input
    type="radio"
    checked={checked}
    readOnly
    disabled={disabled}
    className="w-4 h-4 text-sky-600 border-gray-300 dark:border-slate-600 bg-white dark:bg-slate-800 focus:ring-sky-500 disabled:opacity-50 cursor-pointer disabled:cursor-not-allowed pointer-events-none"
  />
);

export default function ReportTemplate({
  title,
  selectedMode = {},
  onToggleMode,
  availableModes = [],
  singleDate,
  onSingleDateChange,
  startDate,
  endDate,
  onDateRangeChange,
  isLoading,
  onAction,
  actionText,
  extraContent,
  modals,
  setRoutingDate,
  ...props
}) {
  const { t } = useLanguage();
  const settingsRef = useRef(null);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [showWarningModal, setShowWarningModal] = useState(false);
  const [pendingAction, setPendingAction] = useState(null);

  const { bulk: isBulkMode, manual: isManualMode, custom: isCustomRouting } = selectedMode;

  const isRangeInvalid =
    isBulkMode &&
    (!startDate ||
      !endDate ||
      new Date(startDate) > new Date(endDate) ||
      new Date(startDate).getTime() === new Date(endDate).getTime());

  const modeItemStateClass = isLoading
    ? 'cursor-not-allowed text-slate-400 dark:text-slate-500'
    : 'cursor-pointer text-gray-600 dark:text-slate-300';

  const modeConfig = {
    bulk: { checked: isBulkMode, label: t('common.bulk') },
    manual: { checked: isManualMode, label: 'Manual' },
    custom: {
      checked: isCustomRouting,
      label: t('report.change_date'),
      info: t('report.tooltip.info_change_time'),
    },
  };

  useEffect(() => {
    function handleClickOutside(event) {
      if (settingsRef.current && !settingsRef.current.contains(event.target)) {
        setIsSettingsOpen(false);
      }
    }
    if (isSettingsOpen) document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isSettingsOpen]);

  const executeWithCheck = () => {
    if (isBulkMode) {
      const validEndDate = endDate || startDate;
      const diffTime = Math.abs(new Date(validEndDate) - new Date(startDate));
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      if (diffDays > 14) {
        setPendingAction(() => onAction);
        setShowWarningModal(true);
        return;
      }
    }
    onAction();
  };

  const handleConfirmProcess = () => {
    setShowWarningModal(false);
    if (pendingAction) {
      pendingAction();
      setPendingAction(null);
    }
  };

  const handleCancelProcess = () => {
    setShowWarningModal(false);
    setPendingAction(null);
  };

  return (
    <div className="flex flex-col items-center w-full max-w-6xl p-4">
      <h1 className="text-3xl sm:text-4xl font-bold mb-8 text-center text-slate-900 dark:text-slate-100">
        {title}
      </h1>
      <div className="flex flex-col sm:flex-row items-center justify-center gap-4 pb-2 w-full ">
        <div className={`w-full ${isBulkMode ? 'sm:w-[350px]' : 'sm:w-[280px]'}`}>
          <label className="text-lg mb-2 text-gray-500 dark:text-slate-400 font-medium text-center select-none flex items-center justify-center gap-1 w-full">
            {isBulkMode ? t('common.range_delivery') : t('common.delivery_date')}
            {!isCustomRouting && <InformationButton infoText={t('report.tooltip.info_delivery')} />}
          </label>
          {isBulkMode ? (
            <CustomDatePicker
              selectsRange={true}
              startDate={startDate}
              endDate={endDate}
              onChange={onDateRangeChange}
              disabled={isLoading}
              className="w-full block cursor-pointer"
              wrapperClassName="w-full block"
              maxDate={tomorrowDate()}
            />
          ) : (
            <CustomDatePicker
              disabled={isLoading}
              selected={singleDate}
              onChange={onSingleDateChange}
              className="w-full block cursor-pointer"
              wrapperClassName="w-full block"
              maxDate={tomorrowDate()}
            />
          )}
        </div>
        {isCustomRouting && (
          <div className="w-full sm:w-[280px]">
            <label
              htmlFor="routingDate"
              className="block text-lg mb-2 text-gray-500 dark:text-slate-400 font-medium text-center w-full"
            >
              {t('common.routing_date')}
            </label>
            <CustomDatePicker
              disabled={isLoading}
              id="routingDate"
              maxDate={tomorrowDate()}
              onChange={(date) => {
                if (date) setRoutingDate(date);
              }}
              selected={props.routingDate}
              className="w-full sm:w-[280px] block cursor-pointer"
              wrapperClassName="w-full block"
            />
          </div>
        )}
      </div>
      <div className="relative mt-2 mb-4" ref={settingsRef}>
        <button
          onClick={() => setIsSettingsOpen(!isSettingsOpen)}
          className="text-sm text-sky-600 dark:text-sky-400 font-medium hover:underline focus:outline-none cursor-pointer"
        >
          {t('common.setting')}
        </button>
        {isSettingsOpen && (
          <div className="absolute top-full mt-2 left-1/2 -translate-x-1/2 w-64 bg-white dark:bg-slate-800 border border-gray-200 dark:border-slate-700 rounded-md shadow-lg z-50 p-4 flex flex-col gap-4">
            {availableModes.map((modeId) => {
              const mode = modeConfig[modeId];
              if (!mode) return null;
              return (
                <div
                  key={modeId}
                  onClick={() => onToggleMode(modeId)}
                  className={`flex items-center justify-between gap-2 ${modeItemStateClass}`}
                >
                  <ModeRadioInput checked={mode.checked} disabled={isLoading} />
                  <span className="text-sm select-none w-full">{mode.label}</span>
                  {mode.info && <InformationButton infoText={mode.info} />}
                  {modeId === 'bulk' && props.onBulkFormatChange && (
                    <div onClick={(e) => e.stopPropagation()}>
                      <ToggleButton
                        options={[
                          { label: 'ZIP', value: 'zip' },
                          { label: 'XLS', value: 'xls' },
                        ]}
                        value={props.bulkFormat}
                        onChange={props.onBulkFormatChange}
                        disabled={isLoading}
                        size="md"
                        className="w-full"
                      />
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>

      {extraContent}

      <div className="flex flex-row flex-wrap gap-4 w-full justify-center">
        <Button
          onClick={executeWithCheck}
          disabled={isLoading || isRangeInvalid}
          isLoading={isLoading}
          text={actionText}
          width="w-full sm:w-auto min-w-[200px]"
        />
      </div>

      <ConfirmModal
        isOpen={showWarningModal}
        title={t('common.modal.data_load_title')}
        message={t('common.modal.data_load_message', { days: 14 })}
        onConfirm={handleConfirmProcess}
        onCancel={handleCancelProcess}
      />
      {modals}
    </div>
  );
}
