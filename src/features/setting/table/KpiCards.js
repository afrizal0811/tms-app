'use client';

import { formatDateUniversal } from '@/lib/utils';

export default function KpiCards({ kpi, emptyText, translate }) {
  if (!kpi) {
    return (
      <div className="p-4 text-center text-xs text-gray-500 flex flex-1 min-h-0 justify-center items-center">
        {emptyText || translate('common.no_data')}
      </div>
    );
  }

  const groups = [
    {
      title: translate('common.vehicle'),
      items: [
        { label: 'Dry', val: kpi.truck_dry ?? 0 },
        { label: 'Frozen', val: kpi.truck_frz ?? 0 },
        { label: 'Total', val: kpi.total_truck ?? 0, bold: true },
      ],
    },
    {
      title: translate('common.task'),
      items: [
        { label: 'Dry', val: kpi.task_dry ?? 0 },
        { label: 'Frozen', val: kpi.task_frz ?? 0 },
        { label: 'Total', val: (kpi.task_dry ?? 0) + (kpi.task_frz ?? 0), bold: true },
      ],
    },
    {
      title: 'Pending',
      items: [
        { label: 'Dry', val: kpi.pending_dry ?? 0 },
        { label: 'Frozen', val: kpi.pending_frz ?? 0 },
        { label: translate('common.routing'), val: kpi.pending_routing ?? 0 },
        { label: translate('common.driver'), val: kpi.pending_driver ?? 0 },
        { label: 'Sales', val: kpi.pending_sales ?? 0 },
        { label: translate('common.customer'), val: kpi.pending_customer ?? 0 },
        { label: translate('common.others'), val: kpi.pending_other ?? 0 },
      ],
    },
    {
      title: translate('common.duration'),
      items: [
        { label: `${translate('common.estimate')} (Dry)`, val: kpi.est_hour_dry ?? 0 },
        { label: `${translate('common.actual')} (Dry)`, val: kpi.act_hour_dry ?? 0 },
        { label: `${translate('common.estimate')} (Frozen)`, val: kpi.est_hour_frz ?? 0 },
        { label: `${translate('common.actual')} (Frozen)`, val: kpi.act_hour_frz ?? 0 },
      ],
    },
    {
      title: translate('common.distance'),
      items: [
        { label: `${translate('common.estimate')} (Dry)`, val: kpi.est_dist_dry ?? 0 },
        { label: `${translate('common.actual')} (Dry)`, val: kpi.act_dist_dry ?? 0 },
        { label: `${translate('common.estimate')} (Frozen)`, val: kpi.est_dist_frz ?? 0 },
        { label: `${translate('common.actual')} (Frozen)`, val: kpi.act_dist_frz ?? 0 },
      ],
    },
    {
      title: translate('common.capacity'),
      items: [
        { label: 'Dry (kg)', val: `${kpi.total_kg_dry ?? 0} / ${kpi.capacity_kg_dry ?? 0}` },
        { label: 'Dry (cbm)', val: `${kpi.total_cbm_dry ?? 0} / ${kpi.capacity_cbm_dry ?? 0}` },
        { label: 'Frozen (kg)', val: `${kpi.total_kg_frz ?? 0} / ${kpi.capacity_kg_frz ?? 0}` },
        {
          label: 'Frozen (cbm)',
          val: `${kpi.total_cbm_frz ?? 0} / ${kpi.capacity_cbm_frz ?? 0}`,
        },
      ],
    },
  ];

  const hasAudit = Boolean(kpi.created_by || kpi.updated_by);

  return (
    <div className="flex flex-col flex-1 min-h-0">
      <div className="flex-1 overflow-y-auto pr-1">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 m-2">
          {groups.map((g) => (
            <div
              key={g.title}
              className="p-4 rounded-xl border border-gray-200 dark:border-slate-800 bg-slate-50/70 dark:bg-slate-900/50 flex flex-col gap-2.5 shadow-xs overflow-hidden"
            >
              <div className="flex items-center justify-between border-b border-gray-200 dark:border-slate-800 pb-2">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-800 dark:text-slate-200 truncate">
                  {g.title}
                </span>
              </div>
              <div className="flex flex-col gap-1">
                {g.items.map((it) => (
                  <div
                    key={it.label}
                    className="flex items-baseline text-xs py-1 px-1.5 rounded hover:bg-white dark:hover:bg-slate-800/60 transition-colors w-full min-w-0"
                  >
                    <span className="text-slate-600 dark:text-slate-400 shrink-0 truncate max-w-[65%]">
                      {it.label}
                    </span>
                    <span className="flex-1 mx-2 border-b border-dotted border-gray-300 dark:border-slate-700 min-w-2"></span>
                    <span
                      className={`shrink-0 font-semibold text-right tabular-nums ${
                        it.bold
                          ? 'text-sky-600 dark:text-sky-400 font-bold'
                          : 'text-slate-800 dark:text-slate-200'
                      }`}
                    >
                      {it.val}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      </div>
      {hasAudit && (
        <div className="px-4 py-2.5 bg-gray-50 dark:bg-slate-900 border-t border-gray-200 dark:border-slate-700 shrink-0 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500 dark:text-slate-400">
          <div>
            <span className="font-semibold text-slate-600 dark:text-slate-300">
              {translate('common.created_by')}:{' '}
            </span>
            <span>
              {`${kpi.created_by || '-'} (${formatDateUniversal(kpi.created_time, 'DD/MM/YY HH:mm:ss')})`}
            </span>
          </div>
          <div>
            <span className="font-semibold text-slate-600 dark:text-slate-300">
              {translate('common.updated_by')}:{' '}
            </span>
            <span>
              {`${kpi.updated_by || '-'} (${formatDateUniversal(kpi.updated_time, 'DD/MM/YY HH:mm:ss')})`}
            </span>
          </div>
        </div>
      )}
    </div>
  );
}
