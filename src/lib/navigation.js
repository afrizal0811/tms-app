// File: src/lib/navigation.js

export const APP_NAVIGATION = [
  {
    group: 'Report',
    groupLabelKey: 'navbar.report',
    isDropdown: true,
    items: [
      {
        id: '/report/daily',
        href: '/report/daily',
        label: 'Daily Report',
        labelKey: 'navbar.daily_report',
      },
      { id: '/report/kpi', href: '/report/kpi', label: 'KPI Report', labelKey: 'navbar.kpi' },
      {
        id: '/report/custom',
        href: '/report/custom',
        label: 'Custom Report',
        labelKey: 'navbar.custom_report',
      },
      {
        id: '/report/counter',
        href: '/report/counter',
        label: 'Counter Report',
        labelKey: 'navbar.task_counter_report',
      },
    ],
  },
  {
    group: 'Operational',
    groupLabelKey: 'common.menu',
    isDropdown: false,
    items: [
      { id: '/task', href: '/task', label: 'Task', labelKey: 'navbar.task' },
      {
        id: '/tracking',
        href: '/tracking',
        label: 'Tracking',
        labelKey: 'navbar.tracking',
        requiresVms: true,
      },
      { id: '/summary', href: '/summary', label: 'Summary', labelKey: 'navbar.summary' },
      {
        id: '/coordinate',
        href: '/coordinate',
        label: 'Update Coordinate',
        labelKey: 'navbar.coordinate',
        isCoordinate: true,
      },
      { id: '/delivery', href: '/delivery', label: 'Delivery', labelKey: 'navbar.delivery' },
      {
        id: '/vehicles',
        href: '/vehicles',
        label: 'Vehicle Data',
        labelKey: 'common.vehicle',
        isVehicle: true,
      },
    ],
  },
];

export const AVAILABLE_PATHS = APP_NAVIGATION.flatMap((group) =>
  group.items.map((item) => ({
    id: item.id,
    href: item.href,
    label: item.label,
    labelKey: item.labelKey,
    group: group.group,
  }))
);

export const HELP_LINKS = [
  { id: '/help/tms', href: '/help/tms', label: 'TMS', labelKey: 'Transportation Management System (TMS)' },
  { id: '/help/vms', href: '/help/vms', label: 'VMS', labelKey: 'Vehicle Management System (VMS)' },
];
