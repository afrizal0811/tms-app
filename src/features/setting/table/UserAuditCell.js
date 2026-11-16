'use client';

import HighlightText from '@/components/HighlightText';
import { formatDateUniversal } from '@/lib/utils';

export default function UserAuditCell({ user, timestamp, highlight }) {
  const dateStr = timestamp ? formatDateUniversal(timestamp, 'DD-MM-YYYY HH:mm:ss') : '-';
  return (
    <div className="flex flex-col text-left">
      <span className="text-xs font-medium text-slate-800 dark:text-slate-200">
        <HighlightText text={user || '-'} highlight={highlight} />
      </span>
      <span className="text-[11px] text-slate-400 dark:text-slate-500">{dateStr}</span>
    </div>
  );
}
