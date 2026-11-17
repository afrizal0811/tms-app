'use client';

import Button from '@/components/button/Button';
import StorageTypeFilter from '@/components/dropdown/StorageTypeFilter';
import VehicleTypeFilter from '@/components/dropdown/VehicleTypeFilter';
import PageTemplate from '@/components/page/PageTemplate';
import SearchBar from '@/components/SearchBar';
import { useLanguage } from '@/context/LanguageContext';
import { getOdometer } from '@/lib/api/mceasy/odometer';
import { getVehicles } from '@/lib/api/mileapp';
import { formatDateUniversal, getBaseVehicleType, isEmpty } from '@/lib/utils';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { toastError } from '@/lib/toast';
import TemplateTab from './components/TemplateTab';
import VehicleTab from './components/VehicleTab';
import { handleConfirmDownload } from './help';

const countSpaces = (str) => (str ? (str.match(/ /g) || []).length : 0);
const sortData = (a, b) => {
  const typeCmp = (a.type || '').localeCompare(b.type || '');
  if (typeCmp !== 0) return typeCmp;
  return (a.email || '').localeCompare(b.email || '');
};

const processVehicleData = (rawDriversData, odometer) => {
  const odoMap = new Map();
  if (Array.isArray(odometer)) {
    odometer.forEach((o) => {
      if (o.vehicleId != null) odoMap.set(String(o.vehicleId), o.odometer);
    });
  }

  const processedData = rawDriversData.map((v) => {
    let odoValue = '-';
    if (v.vmsVehicleId && odoMap.has(String(v.vmsVehicleId))) {
      odoValue = odoMap.get(String(v.vmsVehicleId));
    }

    const hasPrefix =
      v.storage && v.type && v.type.toLowerCase().startsWith(`${v.storage.toLowerCase()}-`);
    const rawFormattedType =
      v.storage && v.type ? (hasPrefix ? v.type : `${v.storage}-${v.type}`) : v.type || '-';
    const formattedType = rawFormattedType !== '-' ? rawFormattedType.toUpperCase() : '-';

    let parsedTags = [];
    if (v.tags) {
      try {
        parsedTags = JSON.parse(v.tags);
      } catch {
        parsedTags = [v.tags];
      }
    } else if (formattedType) {
      parsedTags = [formattedType];
    }

    const isIncomplete = !v.email || !v.type;

    return {
      ...v,
      type: formattedType,
      odometer: odoValue,
      parsedTags,
      isIncomplete,
    };
  });

  const emailToVehiclesMap = new Map();
  const noEmailVehicles = [];

  processedData.forEach((v) => {
    if (!v.email) {
      noEmailVehicles.push(v);
    } else {
      const emailKey = v.email.toLowerCase();
      if (!emailToVehiclesMap.has(emailKey)) emailToVehiclesMap.set(emailKey, []);
      emailToVehiclesMap.get(emailKey).push(v);
    }
  });

  const masterList = [];
  const conditionalList = [];

  noEmailVehicles.forEach((v) => {
    masterList.push({ ...v, isDuplicateDriver: false });
  });

  for (const vehicles of emailToVehiclesMap.values()) {
    const isDuplicate = vehicles.length > 1;

    if (!isDuplicate) {
      masterList.push({ ...vehicles[0], isDuplicateDriver: false });
    } else {
      const sorted = [...vehicles].sort((a, b) => countSpaces(a.plat) - countSpaces(b.plat));

      masterList.push({ ...sorted[0], isDuplicateDriver: true });

      for (let i = 1; i < sorted.length; i++) {
        const currentVehicle = { ...sorted[i], isDuplicateDriver: true };
        if (countSpaces(currentVehicle.plat) > 2) {
          conditionalList.push(currentVehicle);
        } else {
          masterList.push(currentVehicle);
        }
      }
    }
  }

  const masterSorted = masterList.sort(sortData);
  const conditionalSorted = conditionalList.sort(sortData);
  const templateData = [...masterSorted, ...conditionalSorted];

  return {
    templateData,
    masterData: masterSorted,
    conditionalData: conditionalSorted,
  };
};

const colorLegend = [
  {
    name: 'duplicate_driver',
    colors: 'text-white bg-yellow-100 dark:bg-yellow-400/30',
    text: 'vehicle.tabs.duplicate_driver',
  },
  {
    name: 'incomplete_data',
    colors: 'text-white bg-red-100 dark:bg-red-400/30',
    text: 'vehicle.tabs.incomplete_data',
  },
];

export default function VehicleData() {
  const { t, localeCode } = useLanguage();
  const [activeTab, setActiveTab] = useState('master');
  const [isLoading, setIsLoading] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [isDownloading, setIsDownloading] = useState(false);
  const [storageFilter, setStorageFilter] = useState(['DRY', 'FROZEN']);
  const [masterData, setMasterData] = useState([]);
  const [conditionalData, setConditionalData] = useState([]);
  const [templateData, setTemplateData] = useState([]);
  const [typeFilter, setTypeFilter] = useState('');
  const [masterVehicleTypes, setMasterVehicleTypes] = useState([]);

  useEffect(() => {
    let isMounted = true;

    async function loadData() {
      setIsLoading(true);
      try {
        const rawDriversData = await getVehicles();

        if (!rawDriversData || isEmpty(rawDriversData)) {
          throw new Error(t('common.toast.error', { err: t('common.no_driver') }));
        }
        const odometer = await getOdometer({ date: formatDateUniversal(new Date()) });

        const {
          templateData: temp,
          masterData: master,
          conditionalData: cond,
        } = processVehicleData(rawDriversData, odometer);

        if (!isMounted) return;

        setTemplateData(temp);
        setMasterData(master);
        setConditionalData(cond);
      } catch (e) {
        if (isMounted) toastError(t('common.toast.error', { err: e.message }), e);
      } finally {
        if (isMounted) setIsLoading(false);
      }
    }

    loadData();

    return () => {
      isMounted = false;
    };
  }, [t]);

  const applyStorageFilter = useCallback(
    (list) => {
      return list.filter((item) => {
        if (storageFilter.length === 2) return true;
        if (storageFilter.length === 0) return false;
        const s = (item.storage || '').toUpperCase();
        return storageFilter.includes(s);
      });
    },
    [storageFilter]
  );

  const applyTypeFilter = useCallback(
    (list) => {
      if (!typeFilter || typeFilter === 'all') return list;
      return list.filter(
        (item) => getBaseVehicleType(item.type, masterVehicleTypes) === typeFilter
      );
    },
    [typeFilter, masterVehicleTypes]
  );

  const filteredData = useMemo(() => {
    let data = [];
    if (activeTab === 'master') data = masterData;
    else if (activeTab === 'conditional') data = conditionalData;
    else if (activeTab === 'template') data = templateData;

    data = applyStorageFilter(data);
    data = applyTypeFilter(data);

    if (!searchQuery) return data;

    const lowerQuery = searchQuery.toLowerCase();
    return data.filter((v) => {
      const plat = (v.plat || '').toLowerCase();
      const name = (v.name || '').toLowerCase();
      const email = (v.email || '').toLowerCase();

      return plat.includes(lowerQuery) || name.includes(lowerQuery) || email.includes(lowerQuery);
    });
  }, [
    activeTab,
    masterData,
    conditionalData,
    templateData,
    searchQuery,
    applyStorageFilter,
    applyTypeFilter,
  ]);

  const handleExcelDownload = () => {
    const filteredMaster = applyTypeFilter(applyStorageFilter(masterData));
    const filteredConditional = applyTypeFilter(applyStorageFilter(conditionalData));
    const filteredTemplate = applyTypeFilter(applyStorageFilter(templateData));

    let filePrefix = '';
    if (storageFilter.includes('DRY') && !storageFilter.includes('FROZEN')) filePrefix = 'DRY';
    if (!storageFilter.includes('DRY') && storageFilter.includes('FROZEN')) filePrefix = 'FRZ';

    handleConfirmDownload({
      masterData: filteredMaster,
      conditionalData: filteredConditional,
      templateData: filteredTemplate,
      setIsDownloading,
      sheetSelection: { master: true, conditional: true, template: true },
      t,
      fileNamePrefix: filePrefix,
    });
  };
  const searchPlaceholder = `${t('common.license_number')}, ${t('common.driver')}`;

  const headerItems = [
    {
      label: t('common.search'),
      component: (
        <SearchBar
          disabled={isLoading}
          onChange={setSearchQuery}
          placeholder={t('common.search')}
          tooltip={searchPlaceholder}
          value={searchQuery}
        />
      ),
    },
    {
      label: t('common.storage_type'),
      hideLabel: false,
      component: (
        <StorageTypeFilter
          disabled={isLoading || isDownloading}
          onApply={setStorageFilter}
          selectedTypes={storageFilter}
        />
      ),
    },
    {
      label: t('common.vehicle_type'),
      hideLabel: false,
      component: (
        <VehicleTypeFilter
          disabled={isLoading || isDownloading}
          onApply={setTypeFilter}
          onMasterTypesLoad={setMasterVehicleTypes}
          selectedType={typeFilter}
        />
      ),
    },
    {
      label: 'Export',
      hideLabel: true,
      component: (
        <Button
          onClick={handleExcelDownload}
          isLoading={isDownloading}
          disabled={
            isLoading ||
            isDownloading ||
            (isEmpty(masterData) && isEmpty(conditionalData) && isEmpty(templateData))
          }
          text={t('common.download')}
          width="w-full"
        />
      ),
    },
  ];

  const totalItems = filteredData.length;
  const tabs = useMemo(
    () => [
      { id: 'master', label: t('vehicle.tabs.master_title') },
      ...(conditionalData.length > 0
        ? [{ id: 'conditional', label: t('vehicle.tabs.conditional_title') }]
        : []),
      { id: 'template', label: t('vehicle.tabs.template_title') },
    ],
    [t, conditionalData.length]
  );

  return (
    <PageTemplate
      title={t('common.vehicle_data')}
      subtitle={
        <>
          {t('vehicle.subtitle')}{' '}
          <span className="font-semibold text-sky-600">{t('vehicle.subtitle_highlight')}</span>
        </>
      }
      headerItems={headerItems}
      activeTabId={activeTab}
      onTabClick={setActiveTab}
      tabs={tabs}
      isLoading={isLoading}
      isEmpty={!isLoading && totalItems === 0}
      footer={{ colorData: colorLegend }}
    >
      <div className="flex-1 flex flex-col m-0 overflow-auto">
        {(activeTab === 'master' || activeTab === 'conditional') && (
          <VehicleTab
            localeCode={localeCode}
            paginatedData={filteredData}
            searchQuery={searchQuery}
            t={t}
          />
        )}
        {activeTab === 'template' && (
          <TemplateTab paginatedData={filteredData} searchQuery={searchQuery} t={t} />
        )}
      </div>
    </PageTemplate>
  );
}
