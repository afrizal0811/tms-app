'use client';

import { useLanguage } from '@/context/LanguageContext';
import { getCachedHubs, getLocalStorage, removeLocalStorage } from '@/lib/localStorageHandler';
import { APP_NAVIGATION, HELP_LINKS } from '@/lib/navigation';
import { useTheme } from 'next-themes';
import Link from 'next/link';
import { usePathname } from 'next/navigation';
import React, { useEffect, useRef, useState } from 'react';
import LanguageToggle from '../button/LanguageToggle';
import ThemeToggle from '../button/ThemeToggle';
import { LocationSwitcher } from '../dropdown/LocationDropdown';
import MobileMenu from './MobileMenu';
import NavDropdown from './NavDropdown';
import UserDropdown from './UserDropdown';

function NavLink({ href, children, className }) {
  const pathname = usePathname();
  const isActive = pathname === href;

  return (
    <Link
      href={href}
      className={`text-sm font-medium transition-colors ${className} ${
        isActive
          ? 'text-sky-600 dark:text-sky-400 font-semibold'
          : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-slate-100'
      }`}
    >
      {children}
    </Link>
  );
}

const reportGroup = APP_NAVIGATION.find((g) => g.group === 'Report');
const operationalGroup = APP_NAVIGATION.find((g) => g.group === 'Operational');
const REPORT_LINKS = reportGroup?.items || [];

function DynamicNavMenu({ children, moreLabel = 'More' }) {
  const containerRef = useRef(null);
  const dropdownRef = useRef(null);
  const [visibleCount, setVisibleCount] = useState(React.Children.count(children));
  const [isOpen, setIsOpen] = useState(false);

  useEffect(() => {
    const checkOverflow = () => {
      const container = containerRef.current;
      if (!container) return;

      const items = Array.from(container.querySelectorAll('.nav-item-measure'));
      items.forEach((item) => (item.style.display = 'flex'));

      const containerWidth = container.clientWidth;
      const moreBtnWidth = 90;
      let currentWidth = 0;
      let count = 0;

      for (let i = 0; i < items.length; i++) {
        const itemWidth = items[i].offsetWidth;
        const requiredWidth =
          i === items.length - 1
            ? currentWidth + itemWidth
            : currentWidth + itemWidth + moreBtnWidth;

        if (requiredWidth > containerWidth) break;

        currentWidth += itemWidth;
        count++;
      }

      setVisibleCount(count);
    };

    const observer = new ResizeObserver(checkOverflow);
    if (containerRef.current) observer.observe(containerRef.current);

    checkOverflow();
    return () => observer.disconnect();
  }, [children]);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target)) setIsOpen(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const childrenArray = React.Children.toArray(children);
  const visibleItems = childrenArray.slice(0, visibleCount);
  const overflowItems = childrenArray.slice(visibleCount);

  return (
    <div className="flex-1 flex items-center min-w-0 relative h-full" ref={containerRef}>
      <div
        className="absolute top-0 left-0 h-0 overflow-hidden flex opacity-0 pointer-events-none whitespace-nowrap"
        aria-hidden="true"
      >
        {childrenArray.map((child, idx) => (
          <div key={idx} className="nav-item-measure px-2 lg:px-3">
            {child}
          </div>
        ))}
      </div>

      <div className="flex items-center space-x-2 lg:space-x-5 w-full">
        {visibleItems}

        {overflowItems.length > 0 && (
          <div className="relative shrink-0" ref={dropdownRef}>
            <button
              type="button"
              onClick={() => setIsOpen(!isOpen)}
              className={`flex items-center gap-1 text-sm font-medium transition-colors cursor-pointer ${
                isOpen
                  ? 'text-sky-600 dark:text-sky-400 font-semibold'
                  : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-slate-100'
              }`}
            >
              <span>{moreLabel}</span>
              <svg
                className={`w-4 h-4 transition-transform duration-200 ${
                  isOpen ? 'rotate-180' : 'rotate-0'
                }`}
                viewBox="0 0 20 20"
                fill="none"
                stroke="currentColor"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth="2"
                  d="M6 8l4 4 4-4"
                />
              </svg>
            </button>

            <div
              className={`absolute right-0 mt-4 w-48 bg-white dark:bg-slate-800 border border-gray-200 dark:border-slate-700 rounded-md shadow-lg z-50 transition-all duration-200 origin-top ${
                isOpen ? 'opacity-100 scale-100 visible' : 'opacity-0 scale-95 invisible'
              }`}
            >
              <div className="flex flex-col py-2">
                {overflowItems.map((item, idx) => (
                  <div
                    key={idx}
                    className="px-4 py-2 hover:bg-sky-50 dark:hover:bg-slate-700 w-full flex items-center [&>a]:w-full [&>div]:w-full"
                    onClick={() => setIsOpen(false)}
                  >
                    {item}
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default function Navbar() {
  const { t, isIndonesian } = useLanguage();
  const [mounted, setMounted] = useState(false);
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [isSecret, setIsSecret] = useState(false);
  const [sessionTick, setSessionTick] = useState(0);
  const { theme, setTheme, resolvedTheme } = useTheme();

  const isDarkMode = mounted && (theme === 'dark' || resolvedTheme === 'dark');
  const { storedUser } = getLocalStorage();
  const userPaths = storedUser?.paths || [];
  const userName = storedUser?.name || '';
  const userEmail = storedUser?.email || 'email@example.com';
  const hubs = getCachedHubs();
  const hasVms = hubs ? hubs.activeHub?.hasVms : false;

  const handleLogout = () => {
    removeLocalStorage('data');
    window.location.href = '/';
  };

  useEffect(() => {
    const checkSecret = () => setIsSecret(window.SECRET_MODE_ACTIVE === true);
    checkSecret();
    window.addEventListener('secret_update', checkSecret);
    return () => window.removeEventListener('secret_update', checkSecret);
  }, []);

  useEffect(() => {
    const handleSessionUpdate = () => setSessionTick((prev) => prev + 1);
    window.addEventListener('session_update', handleSessionUpdate);
    window.addEventListener('storage', handleSessionUpdate);

    return () => {
      window.removeEventListener('session_update', handleSessionUpdate);
      window.removeEventListener('storage', handleSessionUpdate);
    };
  }, []);

  useEffect(() => {
    const timer = setTimeout(() => {
      setMounted(true);
      setIsLoggedIn(!!getLocalStorage().storedUser);
    }, 0);
    return () => clearTimeout(timer);
  }, [sessionTick]);

  const vehicle = isIndonesian ? `Data ${t('common.vehicle')}` : `${t('common.vehicle')} Data`;

  if (!mounted) {
    return (
      <nav className="w-full bg-white dark:bg-slate-900 border-b border-gray-200 dark:border-slate-800 px-4 sm:px-6 py-4 sticky top-0 z-100 shadow-sm transition-colors duration-200">
        <div className="max-w-8xl mx-auto flex justify-between items-center px-4">
          <div className="flex items-center space-x-4 sm:space-x-6 w-full lg:w-auto">
            <Link href="/" className="flex flex-col leading-tight">
              <span className="hidden lg:block text-slate-900 dark:text-slate-100 font-bold text-lg sm:text-xl">
                TMS
              </span>
              <span className="block lg:hidden text-slate-900 dark:text-slate-100 font-bold text-lg sm:text-xl">
                TMS Data Processing
              </span>
            </Link>
          </div>
        </div>
      </nav>
    );
  }
  userPaths;

  const checkAccess = (href) => userPaths.length === 0 || userPaths.some((p) => href.startsWith(p));
  const filteredReportLinks = REPORT_LINKS.filter((link) => checkAccess(link.href));

  const getOperationalLabel = (item) => {
    if (item.isCoordinate) {
      return `${t('navbar.update')} ${t(item.labelKey)}`;
    }
    if (item.isVehicle) {
      return <span> {vehicle} </span>;
    }
    return t(item.labelKey);
  };

  const LoggedInComps = (
    <>
      {filteredReportLinks.length > 0 && (
        <NavDropdown
          label={t(reportGroup?.groupLabelKey || 'navbar.report')}
          links={filteredReportLinks}
        />
      )}
      {(operationalGroup?.items || []).map((item) => {
        if (item.requiresVms && !hasVms) return null;
        if (!checkAccess(item.href)) return null;
        return (
          <NavLink key={item.href} href={item.href}>
            {getOperationalLabel(item)}
          </NavLink>
        );
      })}
      <NavDropdown label={t('navbar.help')} links={HELP_LINKS} />
    </>
  );

  const userComps = (
    <div className="hidden lg:flex items-center space-x-4 sm:space-x-6 shrink-0">
      <LocationSwitcher />
      <div className="h-4 w-px bg-gray-300 dark:bg-slate-700" aria-hidden="true"></div>
      <UserDropdown isDarkMode={isDarkMode} />
    </div>
  );

  const mobileMenu = (
    <MobileMenu
      isLoggedIn={isLoggedIn}
      userName={userName}
      userEmail={userEmail}
      userPaths={userPaths}
      handleLogout={handleLogout}
      reportLinks={filteredReportLinks}
      hasVms={hasVms}
    />
  );

  return (
    <nav className="sticky top-0 z-100 w-full px-4 py-4 sm:px-6 bg-white dark:bg-slate-900 border-b border-gray-200 dark:border-slate-800 shadow-md dark:shadow-slate-700/40 transition-colors duration-200">
      <div className="max-w-8xl mx-auto flex justify-between items-center px-4">
        <div
          className={`flex items-center space-x-4 sm:space-x-6 ${isLoggedIn ? 'flex-1 min-w-0 pr-4 lg:pr-8' : 'w-full lg:w-auto'}`}
        >
          <Link href="/" className="flex flex-col leading-tight shrink-0 mr-2 lg:mr-4">
            <span className="hidden lg:block text-slate-900 dark:text-slate-100 font-bold text-lg sm:text-xl">
              TMS
            </span>
            <span className="block lg:hidden text-slate-900 dark:text-slate-100 font-bold text-lg sm:text-xl">
              TMS Data Processing
            </span>
          </Link>

          <div className="hidden lg:flex items-center flex-1 min-w-0">
            {isLoggedIn ? (
              <DynamicNavMenu moreLabel={t('common.others')}>
                {LoggedInComps.props.children}
              </DynamicNavMenu>
            ) : (
              <div className="flex items-center space-x-4 sm:space-x-6">
                {isSecret && <NavLink href="/setting">{t('common.setting')}</NavLink>}
                <NavDropdown label={t('navbar.help')} links={HELP_LINKS} />
              </div>
            )}
          </div>
        </div>

        {isLoggedIn ? (
          <>
            {userComps}
            {mobileMenu}
          </>
        ) : (
          <div className="flex items-center gap-3 shrink-0">
            <ThemeToggle
              isDarkMode={isDarkMode}
              onToggle={() => setTheme(isDarkMode ? 'light' : 'dark')}
              className="text-md px-3 hidden lg:block"
              isLargeIcon={true}
            />
            <LanguageToggle isLargeIcon={true} className="hidden lg:block" />
            {mobileMenu}
          </div>
        )}
      </div>
    </nav>
  );
}
