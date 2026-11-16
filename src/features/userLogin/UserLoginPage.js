'use client';

import LocationSelector from '@/components/dropdown/LocationDropdown';
import ConfirmModal from '@/components/modal/ConfirmModal';
import Spinner from '@/components/Spinner';
import { getRoles, getUsers } from '@/lib/api/mileapp';
import { getLocalStorage, setLocalStorage } from '@/lib/localStorageHandler';
import { toastError, toastSuccess } from '@/lib/toast';
import { capitalizeText } from '@/lib/utils';
import { useEffect, useState } from 'react';

export default function UserLoginPage({ t, allHubsList, currentHubListView, handleUserSelect }) {
  const [selectedLocation, setSelectedLocation] = useState('');
  const [selectedLocationName, setSelectedLocationName] = useState('');
  const [tempSelectedLocation, setTempSelectedLocation] = useState('');
  const [tempSelectedLocationName, setTempSelectedLocationName] = useState('');
  const [emailInput, setEmailInput] = useState('');
  const [loading, setLoading] = useState(false);
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [userToConfirm, setUserToConfirm] = useState(null);
  const [secretClicks, setSecretClicks] = useState(0);
  const [isSecretMode, setIsSecretMode] = useState(false);

  useEffect(() => {
    const { storedLocation, storedLocationName } = getLocalStorage();
    if (storedLocation && storedLocationName) {
      setSelectedLocation(storedLocation);
      setSelectedLocationName(storedLocationName);
      setTempSelectedLocation(storedLocation);
      setTempSelectedLocationName(storedLocationName);
    }
  }, []);

  useEffect(() => {
    if (secretClicks === 0 || isSecretMode) return;
    const timer = setTimeout(() => {
      setSecretClicks(0);
    }, 3000);
    return () => clearTimeout(timer);
  }, [secretClicks, isSecretMode]);

  const handleLocationChange = (id, name) => {
    setTempSelectedLocation(id);
    setTempSelectedLocationName(name);
  };

  const handleSaveLocation = () => {
    if (!tempSelectedLocation) return toastError(t('home.select_branch'));
    const selectedHubObj = allHubsList.find((h) => h._id === tempSelectedLocation);
    const { storedSession } = getLocalStorage();
    const currentData = storedSession || {};
    let userObj = currentData.user || {};

    userObj = {
      ...userObj,
      activeHubId: tempSelectedLocation,
      activeHubName: tempSelectedLocationName,
      activeHubAcronym: selectedHubObj?.acronym || '',
    };

    const newSession = { ...currentData, user: userObj };
    setLocalStorage('data', JSON.stringify(newSession));
    setSelectedLocation(tempSelectedLocation);
    setSelectedLocationName(tempSelectedLocationName);
  };

  const handleResetLocation = () => {
    const { storedSession } = getLocalStorage();
    if (storedSession) {
      const newSession = { ...storedSession };
      delete newSession.user;
      setLocalStorage('data', JSON.stringify(newSession));
    }
    setSelectedLocation('');
    setSelectedLocationName('');
    setEmailInput('');
    setTempSelectedLocation('');
    setTempSelectedLocationName('');
  };

  const handleSecretTrigger = () => {
    if (isSecretMode) return;
    const nextClicks = secretClicks + 1;
    if (nextClicks === 3) {
      setIsSecretMode(true);
      if (typeof window !== 'undefined') {
        window.SECRET_MODE_ACTIVE = true;
        window.dispatchEvent(new Event('secret_update'));
      }
      toastSuccess(t('home.toast.active'));
      setSecretClicks(0);
    } else {
      setSecretClicks(nextClicks);
    }
  };

  const handleSearchUser = async (e) => {
    e.preventDefault();
    if (!emailInput) return;
    if (!selectedLocation) {
      toastError(t('home.toast.no_session'));
      return;
    }
    setLoading(true);
    try {
      const response = await getUsers(emailInput);
      const usersArray = Array.isArray(response) ? response : response?.data || [];
      const foundUser = usersArray[0];

      if (!foundUser) {
        toastError(t('home.toast.no_email_inactive'));
        return;
      }

      const roles = await getRoles();
      const roleNameMap = new Map((roles || []).map((r) => [r._id, r.name?.toLowerCase()]));
      const userRoleName = roleNameMap.get(foundUser.roleId);

      if (userRoleName === 'driver' || userRoleName === 'driver jkt') {
        toastError(t('home.toast.driver_error'));
        return;
      }

      if (!isSecretMode && (userRoleName === 'superadmin' || userRoleName === 'owner')) {
        toastError(t('home.toast.no_email_inactive'));
        return;
      }

      setUserToConfirm(foundUser);
      setIsConfirmOpen(true);
    } catch (err) {
      toastError(t('common.toast.error', { err: err.message }), err);
    } finally {
      setLoading(false);
    }
  };

  const handleConfirmLogin = async () => {
    if (!userToConfirm) return;
    setLoading(true);
    const { storedLocation: currentStoredLocation } = getLocalStorage();
    const targetCheckHubId = currentStoredLocation || selectedLocation;

    if (!targetCheckHubId) {
      toastError(t('home.toast.no_session'));
      setIsConfirmOpen(false);
      return;
    }

    try {
      const selectedHubObj = allHubsList.find((h) => h._id === selectedLocation);
      const { storedSession } = getLocalStorage();
      const currentData = storedSession || {};

      let rolePaths = [];
      try {
        const roles = await getRoles();
        const userRole = (roles || []).find(
          (r) => String(r._id || r.id) === String(userToConfirm.roleId)
        );
        rolePaths = userRole?.paths || [];
      } catch (e) {
        toastError(t('common.toast.error', { err: e.message }), e);
      }

      const filteredUserSession = {
        _id: userToConfirm._id,
        email: userToConfirm.email,
        name: userToConfirm.name,
        hubId: userToConfirm.hubId,
        roleId: userToConfirm.roleId,
        status: userToConfirm.status,
        paths: rolePaths,
        activeHubId: selectedLocation,
        activeHubName: selectedLocationName,
        activeHubAcronym: selectedHubObj?.acronym || '',
      };

      const newSession = {
        ...currentData,
        user: filteredUserSession,
      };

      setLocalStorage('data', JSON.stringify(newSession));
      handleUserSelect(filteredUserSession);
      toastSuccess(t('home.toast.login_success'));
    } catch (err) {
      toastError(t('home.toast.login_failed', { err: err.message }), err);
    }

    setLoading(false);
    setIsConfirmOpen(false);
  };

  if (loading) {
    return (
      <div className="flex flex-col items-center">
        <Spinner />
      </div>
    );
  }

  const modalMessage = (
    <div className="flex flex-col gap-2">
      <div>
        {t('home.modal.question')}{' '}
        <span className="font-bold">{capitalizeText(userToConfirm?.name || '')}</span>?
      </div>
      <div className="text-sm text-gray-500 dark:text-slate-400">{userToConfirm?.email}</div>
    </div>
  );

  return (
    <div className="w-full max-w-md mx-auto relative">
      <ConfirmModal
        isOpen={isConfirmOpen}
        message={modalMessage}
        onCancel={() => {
          setIsConfirmOpen(false);
          setUserToConfirm(null);
        }}
        onConfirm={handleConfirmLogin}
        title={t('home.modal.title')}
        loading={loading}
      />

      <div className="bg-white dark:bg-slate-800 p-6 rounded-lg shadow-sm border border-gray-200 dark:border-slate-700 transition-colors text-left">
        <div className="flex flex-col gap-1 mb-6 text-center">
          <h1 className="text-3xl font-bold text-slate-900 dark:text-slate-100 select-none">
            {t('home.welcome')}
          </h1>
          <p className="text-sm text-gray-400 dark:text-slate-500">
            TMS Data Processing{' '}
            {isSecretMode && <span className="text-emerald-500 font-bold">•</span>}
          </p>
        </div>

        <div className="mb-5">
          <label
            className={`block text-sm font-medium text-gray-500 dark:text-slate-400 mb-1 ${!!selectedLocation ? 'opacity-35 pointer-events-none' : 'opacity-100 cursor-auto'}`}
          >
            {t('home.select_branch')}
          </label>
          <div className="flex gap-2">
            <LocationSelector
              className="flex-1 text-sm disabled:opacity-35 disabled:pointer-events-none"
              disabled={!!selectedLocation}
              hubsToShow={currentHubListView || allHubsList}
              onChange={handleLocationChange}
              value={tempSelectedLocation}
            />
            {!selectedLocation ? (
              <button
                onClick={handleSaveLocation}
                disabled={!tempSelectedLocation}
                className="px-4 py-2 bg-sky-600 text-white rounded hover:bg-sky-700 disabled:bg-gray-400 font-medium text-sm cursor-pointer transition-colors shrink-0"
              >
                {t('common.select')}
              </button>
            ) : (
              <button
                onClick={handleResetLocation}
                className="px-4 py-2 bg-red-100 hover:bg-red-200 text-red-600 dark:bg-red-950/40 dark:text-red-400 rounded font-medium text-sm cursor-pointer transition-colors shrink-0"
              >
                {t('common.button.btn_edit')}
              </button>
            )}
          </div>
        </div>

        <div
          className={`pt-5 border-t border-gray-200 dark:border-slate-700 transition-all duration-300 ${
            !selectedLocation ? 'opacity-35 pointer-events-none' : 'opacity-100'
          }`}
        >
          {loading ? (
            <div className="flex flex-col items-center py-4">
              <Spinner />
            </div>
          ) : (
            <form onSubmit={handleSearchUser} className="flex flex-col gap-4">
              <div>
                <label
                  htmlFor="email"
                  className="block text-sm font-medium text-gray-500 dark:text-slate-400 mb-1"
                >
                  {t('common.email')}
                </label>
                <input
                  autoComplete="off"
                  className="w-full px-4 py-2 border border-gray-300 dark:border-slate-600 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100 rounded-md focus:ring-sky-500 focus:border-sky-500 outline-none transition-colors text-sm disabled:bg-gray-100 dark:disabled:bg-slate-800"
                  id="email"
                  onChange={(e) => setEmailInput(e.target.value)}
                  placeholder="john.doe@mail.com"
                  required
                  type="email"
                  value={emailInput}
                  disabled={!selectedLocation || loading}
                />
              </div>
              <button
                className="w-full bg-sky-600 hover:bg-sky-700 text-white font-bold py-2 px-4 rounded transition-colors disabled:bg-gray-400 dark:disabled:bg-slate-700 dark:disabled:text-slate-400 disabled:cursor-not-allowed cursor-pointer text-sm"
                disabled={!emailInput || !selectedLocation || loading}
                type="submit"
              >
                {t('home.login')}
              </button>
            </form>
          )}
        </div>

        <div className="mt-4 text-right">
          <span
            onClick={handleSecretTrigger}
            className="text-xs text-gray-400 dark:text-slate-500 italic cursor-default select-none pointer-events-auto block"
          >
            *{t('home.note')}
          </span>
        </div>
      </div>
    </div>
  );
}
