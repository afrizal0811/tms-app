'use client';

import 'leaflet/dist/leaflet.css';
import dynamic from 'next/dynamic';
import { useEffect } from 'react';

const DynamicMap = dynamic(
  async () => {
    const L = await import('leaflet');
    const rl = await import('react-leaflet');

    delete L.Icon.Default.prototype._getIconUrl;
    L.Icon.Default.mergeOptions({
      iconRetinaUrl:
        'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-icon-2x.png',
      iconUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-icon.png',
      shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-shadow.png',
    });

    const icons = {
      circle: (
        content,
        bgClass = 'bg-blue-500',
        textClass = 'text-xs',
        borderClass = 'border-white'
      ) =>
        L.divIcon({
          className: 'bg-transparent border-none',
          html: `<div class="${bgClass} w-7 h-7 rounded-full border-2 ${borderClass} shadow-md flex items-center justify-center text-white font-bold ${textClass} z-50 relative box-border overflow-hidden">${content}</div>`,
          iconSize: [28, 28],
          iconAnchor: [14, 14],
        }),
      arrow: (angle, color, isHighlight = false) => {
        const size = isHighlight ? 24 : 16;
        const fontSize = isHighlight ? '20px' : '14px';
        return L.divIcon({
          className: 'bg-transparent border-none',
          html: `<div style="transform: rotate(${-angle}deg); color: ${color}; font-size: ${fontSize}; font-weight: 900; filter: drop-shadow(1px 1px 0px white); display: flex; justify-content: center; align-items: center; width: 100%; height: 100%;">➤</div>`,
          iconSize: [size, size],
          iconAnchor: [size / 2, size / 2],
        });
      },
    };

    const MapInner = ({ children, bounds, onMapReady }) => {
      const map = rl.useMap();
      const boundsKey = bounds && bounds.length > 0 ? JSON.stringify(bounds) : '';

      useEffect(() => {
        if (boundsKey) {
          map.fitBounds(L.latLngBounds(JSON.parse(boundsKey)), { padding: [50, 50], maxZoom: 16 });
        }
      }, [boundsKey, map]);

      useEffect(() => {
        if (onMapReady) onMapReady(map);
      }, [map, onMapReady]);

      return typeof children === 'function' ? children(rl, L, icons, map) : children;
    };

    const MapWrapper = ({
      center = [-6.2, 106.8],
      zoom = 13,
      bounds,
      className = 'h-full w-full z-0',
      children,
      onMapReady,
    }) => {
      return (
        <rl.MapContainer center={center} zoom={zoom} className={className}>
          <rl.TileLayer
            attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
            url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
            crossOrigin="anonymous"
          />
          <MapInner bounds={bounds} onMapReady={onMapReady}>
            {children}
          </MapInner>
        </rl.MapContainer>
      );
    };

    return MapWrapper;
  },
  {
    ssr: false,
    loading: () => (
      <div className="w-full h-full bg-slate-100 dark:bg-slate-800 animate-pulse rounded-lg" />
    ),
  }
);

export default DynamicMap;
