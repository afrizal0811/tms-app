'use client';

import Spinner from '@/components/Spinner';
import TabButton from '@/components/table/TabButton';
import { useLanguage } from '@/context/LanguageContext';
import {
  getDriverStatus,
  getHubs,
  getReasons,
  getRoles,
  getVehicleStatus,
  getVehicleTypes,
} from '@/lib/api/mileapp';
import { useSuperadmin } from '@/lib/hooks/useSuperadmin';
import { getLocalStorage } from '@/lib/localStorageHandler';
import { toastError } from '@/lib/toast';
import { formatDateUniversal } from '@/lib/utils';
import { useCallback, useEffect, useState, useMemo } from 'react';
import GeneralTab from './tabs/GeneralTab';
import PermissionTab from './tabs/PermissionTab';
import SyncDataTab from './tabs/SyncDataTab';
import DatabaseTab from './tabs/DatabaseTab';

export default function SettingPage() {
  const [isAuthorized, setIsAuthorized] = useState(false);
  const [isLoadingPage, setIsLoadingPage] = useState(true);
  const [activeTab, setActiveTab] = useState('general');
  const [detailConfig, setDetailConfig] = useState(null);
  const [lastUpdated, setLastUpdated] = useState({
    hubs: '-',
    roles: '-',
    drivers: '-',
    vehicles: '-',
  });
  const [updatedBy, setUpdatedBy] = useState({
    hubs: '-',
    roles: '-',
    drivers: '-',
    vehicles: '-',
  });
  const [vehicleTypes, setVehicleTypes] = useState([]);
  const [hubs, setHubs] = useState([]);
  const [reasons, setReasons] = useState([]);
  const [rolesList, setRolesList] = useState([]);

  const { t } = useLanguage();
  const { isSuperadmin, isChecking } = useSuperadmin();

  const isSecretMode = typeof window !== 'undefined' && window.SECRET_MODE_ACTIVE === true;
  const isReadOnlyGeneral = !isSuperadmin && !isSecretMode;
  const isReadOnlySync = !isSuperadmin && !isSecretMode;

  const buttonData = useMemo(
    () => [
      { tab: 'general', label: t('setting.tab.general.title') },
      ...(isSuperadmin ? [{ tab: 'permission', label: t('setting.tab.permission.title') }] : []),
      { tab: 'sync', label: t('common.synchronize') },
      ...(detailConfig
        ? [
            {
              tab: 'detail',
              label: 'Database',
              isClosable: true,
            },
          ]
        : []),
    ],
    [isSuperadmin, t, detailConfig]
  );

  const fetchAllData = useCallback(async () => {
    try {
      const [hubsDb, rolesDb, dStatus, vStatus, vTypes, reasonsDb] = await Promise.all([
        getHubs(),
        getRoles(),
        getDriverStatus(),
        getVehicleStatus(),
        getVehicleTypes(),
        getReasons(),
      ]);

      let maxDriverDate = null;
      let driverUpdateBy = '-';
      if (dStatus?._max?.updatedAt) {
        maxDriverDate = new Date(dStatus._max.updatedAt);
        driverUpdateBy = dStatus._max.updateBy || '-';
      }

      let maxVehicleDate = null;
      let vehicleUpdateBy = '-';
      if (vStatus?._max?.updatedAt) {
        maxVehicleDate = new Date(vStatus._max.updatedAt);
        vehicleUpdateBy = vStatus._max.updateBy || '-';
      }

      const maxOverallDate =
        maxDriverDate && maxVehicleDate
          ? maxDriverDate > maxVehicleDate
            ? maxDriverDate
            : maxVehicleDate
          : maxDriverDate || maxVehicleDate;

      const latestDriverSync = maxOverallDate
        ? formatDateUniversal(maxOverallDate, 'DD/MM/YYYY HH:mm:ss')
        : '-';

      const driverOrVehicleUpdateBy =
        maxVehicleDate && maxDriverDate
          ? maxVehicleDate > maxDriverDate
            ? vehicleUpdateBy
            : driverUpdateBy
          : driverUpdateBy !== '-'
            ? driverUpdateBy
            : vehicleUpdateBy;

      const filteredRoles = (rolesDb || []).filter((r) => !r.name.toLowerCase().includes('driver'));

      const hubCounts = {};
      (hubsDb?.allHub || []).forEach((h) => {
        const u = h.updateBy || h.update_by;
        if (u && u !== '-') {
          hubCounts[u] = (hubCounts[u] || 0) + 1;
        }
      });
      let majorityHubUpdateBy = '-';
      let maxHubCount = 0;
      Object.entries(hubCounts).forEach(([user, count]) => {
        if (count > maxHubCount) {
          maxHubCount = count;
          majorityHubUpdateBy = user;
        }
      });

      setHubs(hubsDb);
      setRolesList(filteredRoles);
      setLastUpdated({
        hubs: hubsDb?.activeHub?.updatedAt
          ? formatDateUniversal(new Date(hubsDb.activeHub.updatedAt), 'DD/MM/YYYY HH:mm:ss')
          : '-',
        roles: rolesDb?.[0]?.updatedAt
          ? formatDateUniversal(new Date(rolesDb[0].updatedAt), 'DD/MM/YYYY HH:mm:ss')
          : '-',
        drivers: latestDriverSync,
      });
      setUpdatedBy({
        hubs:
          majorityHubUpdateBy !== '-' ? majorityHubUpdateBy : hubsDb?.activeHub?.updateBy || '-',
        roles: rolesDb?.[0]?.updateBy || '-',
        drivers: driverOrVehicleUpdateBy,
      });

      setVehicleTypes(vTypes?.allTypes || []);
      setReasons(reasonsDb || []);
    } catch (e) {
      toastError(t('common.toast.error', { err: e.message }), e);
    }
  }, [t]);

  useEffect(() => {
    const checkAuthAndLoadData = async () => {
      try {
        const { storedUser } = getLocalStorage();
        const secret = typeof window !== 'undefined' && window.SECRET_MODE_ACTIVE === true;

        if (!storedUser && !secret) {
          toastError(t('home.toast.no_session'));
          return;
        }

        setIsAuthorized(true);
        await fetchAllData();
      } catch (e) {
        toastError(t('common.toast.error', { err: e.message }), e);
      } finally {
        setIsLoadingPage(false);
      }
    };

    checkAuthAndLoadData();
  }, [fetchAllData, t]);

  if (isLoadingPage || isChecking) {
    return (
      <div className="p-8 text-center text-slate-500 flex flex-col justify-center items-center h-full gap-4">
        <Spinner />
      </div>
    );
  }

  if (!isAuthorized) return null;

  const renderTabContent = () => {
    switch (activeTab) {
      case 'general':
        return (
          <GeneralTab
            vehicleTypes={vehicleTypes}
            hubs={hubs}
            reasons={reasons}
            onRefresh={fetchAllData}
            isReadOnly={isReadOnlyGeneral}
            translate={t}
          />
        );
      case 'permission':
        return (
          <PermissionTab
            roles={rolesList}
            onRefresh={fetchAllData}
            isReadOnly={isReadOnlyGeneral}
            translate={t}
          />
        );
      case 'sync':
        return (
          <SyncDataTab
            hubs={hubs}
            lastUpdated={lastUpdated}
            updatedBy={updatedBy}
            onRefresh={fetchAllData}
            isReadOnly={isReadOnlySync}
            translate={t}
            onOpenDetailTab={(config) => {
              setDetailConfig(config);
              setActiveTab('detail');
            }}
          />
        );
      case 'detail':
        return (
          <DatabaseTab
            hubId={detailConfig?.hubId}
            hubName={detailConfig?.hubName}
            dateRange={detailConfig?.dateRange}
            onClose={() => {
              setDetailConfig(null);
              setActiveTab('sync');
            }}
            translate={t}
          />
        );
      default:
        return null;
    }
  };

  return (
    <div className="mx-auto px-4 sm:px-6 py-2 pb-3 w-full h-full flex flex-col grow">
      <div className="mb-2">
        <h1 className="text-3xl font-bold text-gray-900 dark:text-slate-100">
          {t('common.setting')}
        </h1>
      </div>
      <div className="flex border-b border-gray-200 dark:border-slate-700 mb-6 mt-4 overflow-x-auto">
        {buttonData.map((data) => (
          <TabButton
            key={data.tab}
            isActive={activeTab === data.tab}
            onClick={() => setActiveTab(data.tab)}
          >
            <span>{data.label}</span>
            {data.isClosable && (
              <span
                role="button"
                tabIndex={0}
                onClick={(e) => {
                  e.stopPropagation();
                  setDetailConfig(null);
                  setActiveTab('sync');
                }}
                className="ml-1 px-1 py-0.5 rounded-full hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 text-xs cursor-pointer"
                title={t('common.close') || 'Tutup'}
              >
                ✕
              </span>
            )}
          </TabButton>
        ))}
      </div>
      {renderTabContent()}
    </div>
  );
}
