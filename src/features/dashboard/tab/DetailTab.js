'use client';

import Tooltip from '@/components/Tooltip';
import { useLanguage } from '@/context/LanguageContext';
import { forwardRef } from 'react';

const StatCard = forwardRef(function StatCard(
  { title, value, isLoading, className = '', valueClassName = '', tooltipContent },
  ref
) {
  const cardElement = (
    <div
      ref={ref}
      className={`bg-white dark:bg-slate-800 p-6 rounded-lg border border-gray-200 dark:border-slate-700 shadow-md dark:shadow-slate-700/40 hover:bg-gray-50 dark:hover:bg-slate-700/10 cursor-help ${className}`}
    >
      <h3 className="text-sm font-medium text-gray-500 dark:text-slate-400">{title}</h3>
      {isLoading ? (
        <div className="mt-2 h-8 w-12 bg-gray-200 animate-pulse rounded" />
      ) : (
        <p
          className={`mt-1 text-3xl font-semibold text-gray-900 dark:text-slate-200 ${valueClassName}`}
        >
          {value}
        </p>
      )}
    </div>
  );

  if (tooltipContent) {
    return <Tooltip tooltipContent={tooltipContent}>{cardElement}</Tooltip>;
  }
  return cardElement;
});
StatCard.displayName = 'StatCard';

export default function DetailTab({ loading, summaryData }) {
  const { t } = useLanguage();

  const renderTooltip = (titleKey, keyBase) => (
    <div className="space-y-1 text-xs">
      <div>{t(titleKey)}</div>
      <div>Total Dry : {summaryData?.[`${keyBase}Dry`] ?? 0}</div>
      <div>Total Frozen : {summaryData?.[`${keyBase}Frozen`] ?? 0}</div>
    </div>
  );

  return (
    <div className="space-y-10 animate-in fade-in duration-300 h-full flex flex-col flex-1 overflow-auto pb-2 dark:bg-slate-800">
      <div className="grid grid-cols-1 gap-6">
        <div className="flex flex-col gap-6 w-full">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <StatCard
              title={t('common.total_task')}
              value={summaryData?.totalTasks}
              isLoading={loading}
              className="flex flex-col items-center justify-center text-center h-full min-h-[150px]"
              valueClassName="text-5xl"
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.total_task', 'total')}
            />
            <StatCard
              title={t('dashboard.tab.detail.total_assigned')}
              value={summaryData?.assignedTasks}
              isLoading={loading}
              className="flex flex-col items-center justify-center text-center h-full min-h-[150px]"
              valueClassName="text-5xl"
              tooltipContent={renderTooltip(
                'dashboard.tab.detail.tooltip.total_assigned',
                'assigned'
              )}
            />
          </div>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            <StatCard
              title={t('common.status.manual_assign')}
              value={summaryData?.manualAssignList?.length}
              isLoading={loading}
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.manual', 'manualAssign')}
            />
            <StatCard
              title={t('common.status.diff_day')}
              value={summaryData?.diffDayList?.length}
              isLoading={loading}
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.diff_day', 'diffDay')}
            />
            <StatCard
              title={t('dashboard.tab.detail.delivery')}
              value={summaryData?.flowDelivery}
              isLoading={loading}
              tooltipContent={renderTooltip(
                'dashboard.tab.detail.tooltip.delivery',
                'flowDelivery'
              )}
            />
            <StatCard
              title={t('dashboard.tab.detail.redelivery')}
              value={summaryData?.flowReDelivery}
              isLoading={loading}
              tooltipContent={renderTooltip(
                'dashboard.tab.detail.tooltip.redelivery',
                'flowReDelivery'
              )}
            />
          </div>
          <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-5 gap-4">
            <StatCard
              title={t('common.status.unassigned')}
              value={summaryData?.unassigned}
              isLoading={loading}
              tooltipContent={renderTooltip(
                'dashboard.tab.detail.tooltip.unassigned',
                'unassigned'
              )}
            />
            <StatCard
              title={t('common.status.ongoing')}
              value={summaryData?.ongoing}
              isLoading={loading}
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.ongoing', 'ongoing')}
            />
            <StatCard
              title={t('common.status.done')}
              value={summaryData?.done}
              isLoading={loading}
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.done', 'done')}
            />
            <StatCard
              title={t('common.status.success')}
              value={summaryData?.success}
              isLoading={loading}
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.success', 'success')}
            />
            <StatCard
              title={t('common.status.partial')}
              value={summaryData?.partial}
              isLoading={loading}
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.partial', 'partial')}
            />
            <StatCard
              title={t('common.status.pending')}
              value={summaryData?.pending}
              isLoading={loading}
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.pending', 'pending')}
            />
            <StatCard
              title={t('common.status.cancel')}
              value={summaryData?.cancel}
              isLoading={loading}
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.cancel', 'cancel')}
            />
            <StatCard
              title={t('common.status.pending_gr')}
              value={summaryData?.pendingGr}
              isLoading={loading}
              tooltipContent={renderTooltip('dashboard.tab.detail.tooltip.pending_gr', 'pendingGr')}
            />
          </div>
        </div>
      </div>
    </div>
  );
}
