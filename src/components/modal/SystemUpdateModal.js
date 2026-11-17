'use client';

import Button from '@/components/button/Button';
import Modal from '@/components/modal/Modal';
import { useLanguage } from '@/context/LanguageContext';
import { getLocalStorage, setLocalStorage } from '@/lib/localStorageHandler';
import { toastError } from '@/lib/toast';
import { useEffect, useState } from 'react';

const CURRENT_APP_VERSION = process.env.NEXT_PUBLIC_APP_VERSION;

export default function SystemUpdateModal() {
  const { t } = useLanguage();
  const [isOpen, setIsOpen] = useState(false);
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    const timer = setTimeout(() => {
      setMounted(true);
      const { storedUser, appVersion } = getLocalStorage();
      if (storedUser && appVersion !== CURRENT_APP_VERSION) {
        setIsOpen(true);
      }
    }, 0);

    return () => clearTimeout(timer);
  }, []);

  const handleApplyUpdate = async () => {
    const { storedSession, storedLanguage } = getLocalStorage();
    const storedTheme = typeof window !== 'undefined' ? localStorage.getItem('theme') : null;

    localStorage.clear();
    sessionStorage.clear();

    if (storedSession) {
      const userToSave = storedSession.user ? { ...storedSession.user } : null;
      if (userToSave) delete userToSave.paths;

      const newSessionData = {
        ...storedSession,
        user: userToSave,
        superadminRoleId: storedSession.superadminRoleId || null,
      };

      setLocalStorage('data', JSON.stringify(newSessionData));
    }

    if (storedLanguage) {
      setLocalStorage('language', storedLanguage);
    }

    if (storedTheme) {
      localStorage.setItem('theme', storedTheme);
    }

    if ('caches' in window) {
      try {
        const cacheNames = await caches.keys();
        await Promise.all(cacheNames.map((name) => caches.delete(name)));
      } catch (err) {
        toastError(t('common.toast.error', { err: err.message }), err);
      }
    }

    window.location.reload();
  };

  if (!mounted || !isOpen) return null;

  const footer = <Button text={t('common.button.btn_reload')} onClick={handleApplyUpdate} />;

  return (
    <Modal
      isOpen={isOpen}
      noClose={true}
      title={t('update.title')}
      maxWidth="max-w-md"
      footer={footer}
    >
      <div className="flex flex-col items-center text-center">
        <div className="flex items-center justify-center w-16 h-16 bg-sky-100 rounded-full mb-4">
          <svg
            className="w-8 h-8 text-sky-600 animate-[spin_3s_linear_infinite_reverse]"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
            xmlns="http://www.w3.org/2000/svg"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth="2"
              d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"
            />
          </svg>
        </div>
        <p className="text-sm text-slate-600 dark:text-slate-300 mb-2 leading-relaxed">
          {t('update.text')}
        </p>
      </div>
    </Modal>
  );
}
