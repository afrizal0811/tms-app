'use client';

import { useLanguage } from '@/context/LanguageContext';
import { getLocalStorage, removeLocalStorage } from '@/lib/localStorageHandler';
import { capitalizeText } from '@/lib/utils';
import { useTheme } from 'next-themes';
import Link from 'next/link';
import { useEffect, useRef, useState } from 'react';
import LanguageToggle from '../button/LanguageToggle';
import ThemeToggle from '../button/ThemeToggle';

export default function UserDropdown({ isDarkMode }) {
  const [isOpen, setIsOpen] = useState(false);
  const [mounted, setMounted] = useState(false);

  const { t } = useLanguage();
  const { setTheme } = useTheme();
  const dropdownRef = useRef(null);

  const { storedUser } = getLocalStorage();
  const userName = capitalizeText(storedUser?.name);

  useEffect(() => {
    const timer = setTimeout(() => {
      setMounted(true);
    }, 0);
    return () => clearTimeout(timer);
  }, []);

  useEffect(() => {
    function handleClickOutside(event) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) setIsOpen(false);
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleLogout = () => {
    removeLocalStorage('data');
    window.location.href = '/';
  };

  if (!userName || !mounted) return null;

  return (
    <div className="relative inline-block text-left w-full lg:w-auto" ref={dropdownRef}>
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className={`flex items-center justify-between lg:justify-start gap-2 text-sm font-medium transition-colors outline-none cursor-pointer w-full lg:w-auto ${
          isOpen
            ? 'text-sky-600 dark:text-sky-400'
            : 'text-slate-700 dark:text-slate-300 hover:text-slate-900 dark:hover:text-slate-100'
        }`}
      >
        <span>{userName}</span>
        <svg
          className={`w-4 h-4 transition-transform duration-200 ${isOpen ? 'rotate-180' : 'rotate-0'}`}
          viewBox="0 0 20 20"
          fill="currentColor"
        >
          <path
            fillRule="evenodd"
            d="M5.23 7.21a.75.75 0 011.06.02L10 11.168l3.71-3.938a.75.75 0 111.08 1.04l-4.25 4.5a.75.75 0 01-1.08 0l-4.25-4.5a.75.75 0 01.02-1.06z"
            clipRule="evenodd"
          />
        </svg>
      </button>

      {isOpen && (
        <div className="mt-2 rounded-md ring-1 ring-slate-100 dark:ring-slate-700 ring-opacity-50 dark:ring-opacity-100 focus:outline-none z-50 animate-in fade-in zoom-in-95 duration-100 relative w-full shadow-none border overflow-hidden border-gray-100 dark:border-slate-700 bg-gray-50 dark:bg-slate-800 lg:absolute lg:right-0 lg:w-64 lg:shadow-lg lg:border-none lg:bg-white lg:dark:bg-slate-800">
          <div>
            <ThemeToggle
              label={isDarkMode ? t('common.dark_mode') : t('common.light_mode')}
              isDarkMode={isDarkMode}
              onToggle={() => setTheme(isDarkMode ? 'light' : 'dark')}
              className="text-sm px-4"
            />
            <LanguageToggle showLabel={true} className="text-sm px-4" />
            <Link
              href="/setting"
              onClick={() => setIsOpen(false)}
              className="w-full text-left px-4 py-2.5 text-sm flex items-center gap-2 text-slate-700 dark:text-slate-300 hover:bg-sky-50 dark:hover:bg-slate-700 hover:text-sky-600 dark:hover:text-sky-400 transition-colors cursor-pointer"
            >
              {t('common.setting')}
            </Link>
            <button
              onClick={handleLogout}
              className="w-full text-left px-4 py-2.5 text-sm flex items-center gap-2 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-500/20 hover:text-red-700 dark:hover:text-red-300 transition-colors cursor-pointer border-t border-gray-100 dark:border-slate-700/50"
            >
              {t('navbar.logout')}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
