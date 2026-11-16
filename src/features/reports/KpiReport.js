'use client';

import ManualUploadModal from '@/components/modal/ManualUploadModal';
import ReportTemplate from '@/components/page/ReportTemplate';
import { toastError } from '@/lib/toast';
import { formatDateUniversal, isEmpty } from '@/lib/utils';
import { useState } from 'react';
import { handleBulkDownload, handleManualDownload, handleSingleDownload } from './helper/kpiHelper';

export default function KpiReport({
  driverData,
  hubAcronym,
  isIndonesian,
  isLoading,
  setIsLoading,
  t,
}) {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedMode, setSelectedMode] = useState({
    bulk: false,
    manual: false,
  });
  const { bulk: isBulkMode, manual: isManualMode } = selectedMode;
  const [selectedRoutingFiles, setSelectedRoutingFiles] = useState([]);
  const [selectedDeliveryFiles, setSelectedDeliveryFiles] = useState([]);

  const initialDate = formatDateUniversal(new Date());
  const [selectedDate, setSelectedDate] = useState(initialDate);
  const [startDate, setStartDate] = useState(initialDate);
  const [endDate, setEndDate] = useState(initialDate);

  const triggerSingleDownload = () =>
    handleSingleDownload({
      hubAcronym,
      selectedDate,
      driverData,
      setIsLoading,
      t,
    });

  const triggerBulkDownload = () =>
    handleBulkDownload({
      hubAcronym,
      startDate,
      endDate,
      driverData,
      setIsLoading,
      t,
    });

  const triggerManualDownload = () =>
    handleManualDownload({
      hubAcronym,
      selectedRoutingFiles,
      selectedDeliveryFiles,
      driverData,
      setIsLoading,
      setIsModalOpen,
      setSelectedRoutingFiles,
      setSelectedDeliveryFiles,
      t,
    });

  const handleAction = () => {
    if (isManualMode) {
      if (isEmpty(driverData)) {
        toastError(t('common.no_driver'));
        return;
      }
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
      [mode]: !prev[mode],
    }));
  };

  const manualText = isManualMode ? 'Manual' : '';
  const bulkText = isBulkMode ? t('common.bulk') : '';
  const prefixText = `${manualText} ${bulkText}`.trim();
  const titleMenu = isIndonesian
    ? `${t('common.kpi_report')} ${prefixText}`.trim()
    : `${prefixText} ${t('common.kpi_report')}`.trim();

  const handleSingleDateChange = (date) => {
    if (date) {
      setSelectedDate(date);
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
      availableModes={['bulk', 'manual']}
      singleDate={selectedDate}
      onSingleDateChange={handleSingleDateChange}
      startDate={startDate}
      endDate={endDate}
      onDateRangeChange={handleDateRangeChange}
      isLoading={isLoading}
      onAction={handleAction}
      actionText={isManualMode ? t('common.upload') : t('common.download')}
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
