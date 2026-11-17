'use client';

import Map from '@/components/Map';
import { isEmpty } from '@/lib/utils';
import { useEffect, useMemo, useRef } from 'react';
import { useMap } from 'react-leaflet';

function HighlightEffect({ activeCoords, highlightTrigger, markerRefs }) {
  const map = useMap();
  useEffect(() => {
    if (!activeCoords) return;
    const [lat, lng] = activeCoords.split(',').map(Number);
    if (isNaN(lat) || isNaN(lng)) return;
    map.flyTo([lat, lng], 18, { animate: true, duration: 1.5 });
  }, [activeCoords, highlightTrigger, map, markerRefs]);
  return null;
}

function InitialFitBounds({ points }) {
  const map = useMap();
  useEffect(() => {
    if (!points || isEmpty(points)) return;
    const bounds = L.latLngBounds(points);
    map.fitBounds(bounds, { padding: [50, 50] });
  }, [map, points]);
  return null;
}

export default function MapLocation({ data, activeCoords, highlightTrigger, t, isIndonesian }) {
  const markerRefs = useRef({});

  const latestOldPoint = useMemo(() => {
    if (!data || isEmpty(data)) return null;
    let latest = null;
    for (const item of data) {
      if (!item.oldLonglat || !item.date) continue;
      const [d, m, y] = item.date.split('/');
      const parsedDate = new Date(`${y}-${m}-${d}`);
      if (!latest || parsedDate > latest.parsedDate) {
        latest = {
          oldLonglat: item.oldLonglat,
          parsedDate,
          distanceDiff: item.distanceDiff,
        };
      }
    }
    return latest;
  }, [data]);

  const allPoints = useMemo(() => {
    const points = [];
    if (latestOldPoint) {
      const [lat, lng] = latestOldPoint.oldLonglat.split(',').map(Number);
      if (!isNaN(lat) && !isNaN(lng)) points.push([lat, lng]);
    }
    data.forEach((item) => {
      if (!item.newLonglat || !item.oldLonglat) return;
      const [newLat, newLng] = item.newLonglat.split(',').map(Number);
      if (!isNaN(newLat) && !isNaN(newLng)) points.push([newLat, newLng]);
    });
    return points;
  }, [data, latestOldPoint]);

  if (isEmpty(allPoints)) {
    return (
      <div className="flex items-center justify-center h-full bg-gray-100 dark:bg-slate-800 text-gray-400 dark:text-slate-500 text-sm rounded-lg">
        {t('coordinate.map.invalid_data')}
      </div>
    );
  }

  return (
    <Map center={allPoints[0]} zoom={13} className="h-full w-full rounded-lg z-0">
      {({ Marker, Polyline, Tooltip: LeafletTooltip }, L, icons) => {
        const elements = [];
        if (latestOldPoint) {
          const [lat, lng] = latestOldPoint.oldLonglat.split(',').map(Number);
          if (!isNaN(lat) && !isNaN(lng)) {
            elements.push(
              <Marker
                key="old-latest"
                position={[lat, lng]}
                icon={icons.circle(isIndonesian ? 'L' : 'O', 'bg-red-500')}
                opacity={0.9}
              >
                <LeafletTooltip direction="top" offset={[0, -10]}>
                  <div
                    dangerouslySetInnerHTML={{
                      __html: `<b>${t('common.coord_old')}</b>`,
                    }}
                  />
                </LeafletTooltip>
              </Marker>
            );
          }
        }

        data.forEach((item, index) => {
          if (!item.newLonglat || !item.oldLonglat) return;
          const [newLat, newLng] = item.newLonglat.split(',').map(Number);
          const [oldLat, oldLng] = item.oldLonglat.split(',').map(Number);
          if (isNaN(newLat) || isNaN(oldLat)) return;

          elements.push(
            <Marker
              key={`new-${index}`}
              position={[newLat, newLng]}
              icon={icons.circle(isIndonesian ? 'B' : 'N', 'bg-sky-500')}
              ref={(el) => {
                if (el) markerRefs.current[item.newLonglat] = el;
              }}
            >
              <LeafletTooltip direction="top" offset={[0, -10]}>
                <div
                  dangerouslySetInnerHTML={{
                    __html: `<b>${t('common.coord_new')}</b>`,
                  }}
                />
              </LeafletTooltip>
            </Marker>
          );

          elements.push(
            <Polyline
              key={`line-${index}`}
              positions={[
                [oldLat, oldLng],
                [newLat, newLng],
              ]}
              pathOptions={{ color: 'blue', dashArray: '5, 10', opacity: 0.5 }}
            />
          );
        });

        return (
          <>
            {elements}
            <InitialFitBounds points={allPoints} />
            <HighlightEffect
              activeCoords={activeCoords}
              highlightTrigger={highlightTrigger}
              markerRefs={markerRefs}
            />
          </>
        );
      }}
    </Map>
  );
}
