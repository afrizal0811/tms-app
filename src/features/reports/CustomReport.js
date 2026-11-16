'use client';

import RadioButton from '@/components/button/RadioButton';
import ReportTemplate from '@/components/page/ReportTemplate';
import { useState } from 'react';
import { handleCustomDownload } from './helper/customHelper';

export default function CustomReport({
  driverData,
  hubAcronym,
  hubName,
  isIndonesian,
  isLoading,
  setIsLoading,
  t,
}) {
  const [singleDate, setSingleDate] = useState(new Date());
  const [startDate, setStartDate] = useState(new Date());
  const [endDate, setEndDate] = useState(new Date());
  const [selectedMode, setSelectedMode] = useState({
    bulk: false,
    manual: false,
    custom: false,
  });
  const { bulk: isBulkMode } = selectedMode;
  const [reportType, setReportType] = useState('bread');
  const [bulkFormat, setBulkFormat] = useState('zip');

  const reportOptions = [
    {
      id: 'bread',
      label: t('report.custom.bread_items'),
      tooltip: t('report.tooltip.bread_items_info'),
    },
    {
      id: 'detail',
      label: t('report.custom.task_routing'),
      tooltip: t('report.tooltip.task_routing_info'),
    },
    {
      id: 'manual',
      label: t('report.custom.task_manual'),
      tooltip: t('report.tooltip.task_manual_info'),
    },
    {
      id: 'service_level',
      label: t('report.custom.service_level'),
      tooltip: t('report.tooltip.service_level_info'),
    },
    {
      id: 'trip_activity',
      label: t('report.custom.trip_activity'),
      tooltip: t('report.tooltip.trip_activity_info'),
    },
  ];

  const handleRadioToggle = (mode) => {
    if (isLoading) return;
    setSelectedMode((prev) => ({
      bulk: false,
      manual: false,
      custom: false,
      [mode]: !prev[mode],
    }));
  };

  const executeProcess = () => {
    handleCustomDownload({
      isBulkMode,
      bulkFormat,
      startDate,
      endDate,
      singleDate,
      driverData,
      hubAcronym,
      hubName,
      t,
      setIsLoading,
      reportType,
    });
  };
  const handleDateRangeChange = (dates) => {
    const [start, end] = dates;
    setStartDate(start);
    setEndDate(end);
  };

  const bulkText = isBulkMode ? t('common.bulk') : '';
  const titleMenu = isIndonesian
    ? `${t('common.custom_report')} ${bulkText}`.trim()
    : `${bulkText} ${t('common.custom_report')}`.trim();

  return (
    <ReportTemplate
      title={titleMenu}
      selectedMode={selectedMode}
      onToggleMode={handleRadioToggle}
      availableModes={['bulk']}
      singleDate={singleDate}
      onSingleDateChange={setSingleDate}
      startDate={startDate}
      endDate={endDate}
      onDateRangeChange={handleDateRangeChange}
      isLoading={isLoading}
      onAction={executeProcess}
      actionText={t('common.download')}
      bulkFormat={bulkFormat}
      onBulkFormatChange={setBulkFormat}
      extraContent={
        <div className="flex flex-col items-center mb-10 w-full">
          <span className="text-lg mb-3 text-gray-500 dark:text-slate-400 font-medium text-center select-none">
            {t('common.type')}
          </span>
          <RadioButton
            options={reportOptions}
            selected={reportType}
            onChange={setReportType}
            disabled={isLoading}
          />
        </div>
      }
    />
  );
}
