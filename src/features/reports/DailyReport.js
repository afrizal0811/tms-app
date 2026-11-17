'use client';

import ManualUploadModal from '@/components/modal/ManualUploadModal';
import ReportTemplate from '@/components/page/ReportTemplate';
import { toastError } from '@/lib/toast';
import { formatDateUniversal, isEmpty } from '@/lib/utils';
import { useState } from 'react';
import {
  handleBulkDownload,
  handleManualDownload,
  handleSingleDownload,
} from './helper/dailyHelper';

export default function DailyReport({
  driverData,
  hubName,
  isIndonesian,
  isLoading,
  setIsLoading,
  t,
}) {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedMode, setSelectedMode] = useState({
    bulk: false,
    manual: false,
    custom: false,
  });
  const { bulk: isBulkMode, manual: isManualMode, custom: isCustomRouting } = selectedMode;
  const [selectedRoutingFiles, setSelectedRoutingFiles] = useState([]);
  const [selectedDeliveryFiles, setSelectedDeliveryFiles] = useState([]);

  const initialDate = formatDateUniversal(new Date());
  const [selectedDate, setSelectedDate] = useState(initialDate);
  const [routingDate, setRoutingDate] = useState(initialDate);
  const [startDate, setStartDate] = useState(initialDate);
  const [endDate, setEndDate] = useState(initialDate);

  const selectedDateString = formatDateUniversal(selectedDate);

  const triggerSingleDownload = () =>
    handleSingleDownload({
      hubName,
      selectedDate,
      selectedDateString,
      isCustomRouting,
      routingDate,
      driverData,
      setIsLoading,
      t,
      isIndonesian,
    });

  const triggerBulkDownload = () =>
    handleBulkDownload({
      hubName,
      startDate,
      endDate,
      driverData,
      setIsLoading,
      t,
      isIndonesian,
    });

  const triggerManualDownload = () =>
    handleManualDownload({
      hubName,
      selectedDateString,
      selectedRoutingFiles,
      selectedDeliveryFiles,
      driverData,
      setIsLoading,
      setIsModalOpen,
      setSelectedRoutingFiles,
      setSelectedDeliveryFiles,
      t,
      isIndonesian,
    });

  const handleAction = () => {
    if (isEmpty(driverData)) {
      toastError(t('common.no_driver'));
      return;
    }
    if (isManualMode) {
      setIsModalOpen(true);
    } else if (isBulkMode) {
      triggerBulkDownload();
    } else {
      triggerSingleDownload();
    }
  };

  const handleRadioToggle = (mode) => {
    if (isLoading) return;
    setSelectedMode((prev) => ({
      bulk: false,
      manual: false,
      custom: false,
      [mode]: !prev[mode],
    }));
  };

  const manualText = isManualMode ? 'Manual' : '';
  const bulkText = isBulkMode ? t('common.bulk') : '';
  const prefixText = `${manualText} ${bulkText}`.trim();
  const titleMenu = isIndonesian
    ? `${t('common.daily_report')} ${prefixText}`.trim()
    : `${prefixText} ${t('common.daily_report')}`.trim();

  const handleSingleDateChange = (date) => {
    if (date) {
      setSelectedDate(date);
      setRoutingDate(date);
    }
  };

  const handleDateRangeChange = (dates) => {
    const [start, end] = dates;
    setStartDate(start);
    setEndDate(end);
  };

  return (
    <ReportTemplate
      title={titleMenu}
      selectedMode={selectedMode}
      onToggleMode={handleRadioToggle}
      availableModes={['bulk', 'custom', 'manual']}
      singleDate={selectedDate}
      onSingleDateChange={handleSingleDateChange}
      startDate={startDate}
      endDate={endDate}
      onDateRangeChange={handleDateRangeChange}
      isLoading={isLoading}
      onAction={handleAction}
      actionText={isManualMode ? t('common.upload') : t('common.download')}
      routingDate={routingDate}
      setRoutingDate={setRoutingDate}
      modals={
        <ManualUploadModal
          isOpen={isModalOpen}
          onClose={() => setIsModalOpen(false)}
          isLoading={isLoading}
          onDownload={triggerManualDownload}
          routingFiles={selectedRoutingFiles}
          setRoutingFiles={setSelectedRoutingFiles}
          deliveryFiles={selectedDeliveryFiles}
          setDeliveryFiles={setSelectedDeliveryFiles}
        />
      }
    />
  );
}
