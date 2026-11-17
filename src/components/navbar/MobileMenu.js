'use client';

import { useLanguage } from '@/context/LanguageContext';
import { APP_NAVIGATION, HELP_LINKS } from '@/lib/navigation';
import { isEmpty } from '@/lib/utils';
import { useTheme } from 'next-themes';
import Link from 'next/link';
import { usePathname } from 'next/navigation';
import { useEffect, useRef, useState } from 'react';
import LanguageToggle from '../button/LanguageToggle';
import ThemeToggle from '../button/ThemeToggle';
import { LocationSwitcher } from '../dropdown/LocationDropdown';

const operationalGroup = APP_NAVIGATION.find((g) => g.group === 'Operational');

function MobileNavLink({ href, children, target = '', rel = '' }) {
  const pathname = usePathname();
  const isActive = pathname === href;
  const isExternal = target === '_blank';
  const baseClassName = `block w-full p-3 text-base font-medium ${
    isActive
      ? 'text-sky-600 dark:text-sky-400 bg-sky-50 dark:bg-sky-900/30'
      : 'text-slate-700 dark:text-slate-300 hover:bg-gray-100 dark:hover:bg-slate-800'
  }`;

  if (isExternal) {
    return (
      <a href={href} target={target} rel={rel} className={baseClassName}>
        {children}
      </a>
    );
  }

  return (
    <Link href={href} className={baseClassName}>
      {children}
    </Link>
  );
}

function MobileNavGroup({ label, links }) {
  const { t } = useLanguage();
  const [isOpen, setIsOpen] = useState(false);

  if (!links || links.length === 0) return null;

  return (
    <div>
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className="w-full flex justify-between items-center p-3 text-base font-medium text-slate-700 dark:text-slate-300 hover:bg-gray-100 dark:hover:bg-slate-800 transition-colors cursor-pointer"
      >
        <span>{label}</span>
        <svg
          className={`w-5 h-5 transition-transform duration-200 ${isOpen ? 'rotate-180' : 'rotate-0'}`}
          viewBox="0 0 20 20"
          fill="none"
          stroke="currentColor"
        >
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 8l4 4 4-4" />
        </svg>
      </button>
      <div
        className={`transition-all duration-200 ease-in-out origin-top overflow-hidden bg-slate-50 dark:bg-slate-800/50 pl-4 ${
          isOpen ? 'max-h-[500px] opacity-100' : 'max-h-0 opacity-0'
        }`}
      >
        {links.map((link) => (
          <MobileNavLink key={link.href} href={link.href}>
            {t(link.labelKey)}
          </MobileNavLink>
        ))}
      </div>
    </div>
  );
}

function Divider() {
  return (
    <div className="pt-2 pb-1 px-3">
      <div className="border-t border-gray-200 dark:border-slate-800"></div>
    </div>
  );
}

export default function MobileMenu({
  isLoggedIn,
  userName,
  userEmail,
  userPaths,
  handleLogout,
  reportLinks,
  hasVms = true,
}) {
  const { t, isIndonesian } = useLanguage();
  const { theme, setTheme, resolvedTheme } = useTheme();
  const [isOpen, setIsOpen] = useState(false);
  const [mounted, setMounted] = useState(false);
  const [isSecret, setIsSecret] = useState(false);
  const pathname = usePathname();
  const menuRef = useRef(null);
  const isAllPath = isEmpty(userPaths);

  useEffect(() => {
    const checkSecret = () => setIsSecret(window.SECRET_MODE_ACTIVE === true);
    checkSecret();
    window.addEventListener('secret_update', checkSecret);
    return () => window.removeEventListener('secret_update', checkSecret);
  }, []);

  useEffect(() => {
    const timer = setTimeout(() => {
      setMounted(true);
    }, 0);
    return () => clearTimeout(timer);
  }, []);

  const isDarkMode = mounted && (theme === 'dark' || resolvedTheme === 'dark');

  useEffect(() => {
    const timer = setTimeout(() => {
      setIsOpen(false);
    }, 0);
    return () => clearTimeout(timer);
  }, [pathname]);

  useEffect(() => {
    if (isOpen) {
      document.body.style.overflow = 'hidden';
    } else {
      document.body.style.overflow = '';
    }
    return () => {
      document.body.style.overflow = '';
    };
  }, [isOpen]);

  useEffect(() => {
    function handleClickOutside(event) {
      if (menuRef.current && !menuRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    }
    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen]);

  const primaryVehicle = isIndonesian ? 'Data' : t('common.vehicle');
  const secondaryVehicle = isIndonesian ? t('common.vehicle') : 'Data';

  const getMobileLabel = (item) => {
    if (item.isCoordinate) {
      return `${t('navbar.update')} ${t(item.labelKey)}`;
    }
    if (item.isVehicle) {
      return `${primaryVehicle} ${secondaryVehicle}`;
    }
    return t(item.labelKey);
  };

  return (
    <div ref={menuRef} className="lg:hidden">
      <button
        onClick={() => setIsOpen(!isOpen)}
        aria-label="Toggle menu"
        className="p-2 rounded-md text-slate-700 dark:text-slate-300 hover:bg-gray-100 dark:hover:bg-slate-800 cursor-pointer transition-colors duration-200"
      >
        <div className={`transition-transform duration-300 ${isOpen ? 'rotate-90' : 'rotate-0'}`}>
          <svg
            xmlns="http://www.w3.org/2000/svg"
            fill="none"
            viewBox="0 0 24 24"
            strokeWidth={2.5}
            stroke="currentColor"
            className="w-6 h-6"
          >
            {isOpen ? (
              <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
            ) : (
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5"
              />
            )}
          </svg>
        </div>
      </button>

      <div
        className={`absolute top-full left-0 right-0 bg-white dark:bg-slate-900 shadow-lg border-t border-gray-200 dark:border-slate-800 transition-all duration-200 ease-in-out ${
          isOpen ? 'max-h-[80vh] opacity-100 overflow-y-auto' : 'max-h-0 opacity-0 overflow-hidden'
        }`}
      >
        <div className="flex flex-col pb-4 ">
          {isLoggedIn ? (
            <>
              <div className=" px-4 py-5 flex items-center justify-between text-slate-800 dark:text-slate-200 gap-3">
                <div className="flex items-center gap-3 overflow-hidden min-w-0">
                  <div className="flex flex-col overflow-hidden min-w-0">
                    <span className="text-sm font-bold truncate tracking-wide">{userName}</span>
                    <span className="text-[11px] text-slate-400 dark:text-slate-500 truncate">
                      {userEmail || 'email@example.com'}
                    </span>
                  </div>
                </div>
                <LocationSwitcher />
              </div>
              <Divider />
              {reportLinks.length > 0 && (
                <MobileNavGroup label={t('navbar.report')} links={reportLinks} />
              )}
              {(operationalGroup?.items || []).map((item) => {
                if (item.requiresVms && !hasVms) return null;
                if (!isAllPath && !userPaths.some((p) => item.href.startsWith(p))) return null;
                return (
                  <MobileNavLink key={item.href} href={item.href}>
                    {getMobileLabel(item)}
                  </MobileNavLink>
                );
              })}
              <Divider />
            </>
          ) : null}
          <ThemeToggle
            label={isDarkMode ? t('common.dark_mode') : t('common.light_mode')}
            isDarkMode={isDarkMode}
            onToggle={() => setTheme(isDarkMode ? 'light' : 'dark')}
            className="text-md px-3"
          />
          <LanguageToggle showLabel={true} className="text-base px-3 py-2.5 mb-1 " />
          <MobileNavGroup label={t('navbar.help')} links={HELP_LINKS} />
          {!isLoggedIn && isSecret && (
            <MobileNavLink href="/setting">{t('common.setting')}</MobileNavLink>
          )}
          {isLoggedIn && (
            <>
              <MobileNavLink href="/setting">{t('common.setting')}</MobileNavLink>
              <button
                onClick={handleLogout}
                className="block w-full text-left px-3 py-3 text-base font-medium text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/30 hover:text-red-700 dark:hover:text-red-300 transition-colors cursor-pointer"
              >
                {t('navbar.logout')}
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
