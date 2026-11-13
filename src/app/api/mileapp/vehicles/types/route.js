import prisma from '@/lib/prisma';
import { getBasePlate, getStorageType, isEmpty } from '@/lib/utils';
import { NextResponse } from 'next/server';
export const dynamic = 'force-dynamic';

function handleApiError(logLabel, message, error) {
  console.error(logLabel, error);
  const detail = error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
  return NextResponse.json({ error: message, detail }, { status: 500 });
}

export async function GET(request) {
  try {
    const { searchParams } = new URL(request.url);
    const hubId = searchParams.get('hubId');

    const vehicleWhere = hubId && hubId !== 'all' ? { hubs_id: hubId } : {};
    const driverWhere = hubId && hubId !== 'all' ? { hub_id: hubId } : {};

    const [dbVehicles, driverList, allTypes] = await Promise.all([
      prisma.vehicle.findMany({ where: vehicleWhere }),
      prisma.driverList.findMany({ where: driverWhere }),
      prisma.vehicleType.findMany({ orderBy: { name: 'asc' } }),
    ]);

    const allowedTypes = new Set(allTypes.map((v) => v.name.toUpperCase()));

    let parsedDrivers = [];
    if (dbVehicles.length > 0) {
      const driverMap = new Map(driverList.map((d) => [d.id, d]));
      parsedDrivers = dbVehicles.map((v) => {
        const driver = v.driver_id ? driverMap.get(v.driver_id) : null;
        return {
          id: v.id,
          vehicle_id: v.vehicle_id,
          plat: v.license_number,
          basePlat: getBasePlate(v.license_number),
          storage: v.storage,
          type: v.type ? (v.storage ? `${v.storage}-${v.type}` : v.type) : null,
          email: driver ? driver.email : '',
          name: driver ? driver.name : '-',
        };
      });
    } else {
      const where = hubId && hubId !== 'all' ? { hubs: { some: { id: hubId } } } : {};
      const [rawDrivers, mappingsDB] = await Promise.all([
        prisma.driver.findMany({ where }),
        prisma.vehicleMapping.findMany(),
      ]);

      const mappingsObj = mappingsDB.reduce((acc, curr) => {
        acc[curr.plat] = curr.mapped_type ?? curr.mappedType;
        acc[getBasePlate(curr.plat)] = curr.mapped_type ?? curr.mappedType;
        return acc;
      }, {});

      parsedDrivers = rawDrivers.map((d) => {
        const basePlat = getBasePlate(d.plat);
        const mappedType = mappingsObj[basePlat] || mappingsObj[d.plat];
        let mappedTypeStr = d.type;
        if (mappedType) {
          mappedTypeStr = d.storage ? `${d.storage}-${mappedType}` : mappedType;
        }
        return { ...d, type: mappedTypeStr, basePlat };
      });
    }

    const groupedByEmail = {};
    parsedDrivers.forEach((d) => {
      const email = (d.email || '').toLowerCase().trim();
      if (email && email !== '-') {
        if (!groupedByEmail[email]) groupedByEmail[email] = [];
        groupedByEmail[email].push(d);
      }
    });

    const activeTypesSet = new Set();
    const masterData = { Dry: { Total: 0 }, Frozen: { Total: 0 } };

    parsedDrivers.forEach((d) => {
      const plat = d.basePlat || d.plat || '';
      const platUpper = plat.toUpperCase();

      if (!plat || isEmpty(plat.trim()) || platUpper.includes('DEMO')) {
        return;
      }

      let isConditional = false;
      const isSewa = platUpper.includes('SEWA');
      const email = (d.email || '').toLowerCase().trim();

      if (!isSewa && email && email !== '-' && groupedByEmail[email]) {
        const group = groupedByEmail[email];
        if (group.length > 1) {
          const spaceCount = plat.trim().split(' ').length - 1;
          const minSpaces = Math.min(
            ...group.map((v) => (v.plat || '').trim().split(' ').length - 1)
          );
          if (spaceCount > minSpaces && spaceCount > 2) {
            isConditional = true;
          }
        }
      }

      if (isConditional) return;

      let resolvedType = (d.type || 'Lainnya').trim().toUpperCase();
      if (!allowedTypes.has(resolvedType)) {
        resolvedType = resolvedType.replace(/^(DRY|FROZEN|FRZ)-/, '');
      }

      if (!allowedTypes.has(resolvedType)) return;

      activeTypesSet.add(resolvedType);
      
      let storageCategory = getStorageType(d.tags || d.name || d.type);
      if (d.type) {
        const t = d.type.toUpperCase();
        if (t.includes('DRY')) storageCategory = 'Dry';
        else if (t.includes('FROZEN') || t.includes('FRZ')) storageCategory = 'Frozen';
      }

      if (!masterData[storageCategory]) {
        masterData[storageCategory] = { Total: 0 };
      }

      if (masterData[storageCategory][resolvedType] === undefined) {
        Object.keys(masterData).forEach((key) => {
          masterData[key][resolvedType] = 0;
        });
      }

      const isMainVehicle = d.plat === d.basePlat;
      if (isMainVehicle && !isSewa) {
        masterData[storageCategory][resolvedType]++;
        masterData[storageCategory].Total++;
      }
    });

    const CUSTOM_SORT_ORDER = ['L300', 'CDE', 'CDE-LONG', 'CDD', 'CDD-LONG', 'FUSO', 'FUSO-LONG'];

    const activeTypes = Array.from(activeTypesSet).sort((a, b) => {
      const indexA = CUSTOM_SORT_ORDER.indexOf(a);
      const indexB = CUSTOM_SORT_ORDER.indexOf(b);

      if (indexA !== -1 && indexB !== -1) return indexA - indexB;
      if (indexA !== -1) return -1;
      if (indexB !== -1) return 1;

      return a.localeCompare(b);
    });

    return NextResponse.json({ masterData, activeTypes, allTypes });
  } catch (error) {
    return handleApiError('Gagal memproses tipe kendaraan', 'Gagal memproses tipe kendaraan', error);
  }
}

export async function POST(request) {
  try {
    const { name, updateBy } = await request.json();
    if (!name) return NextResponse.json({ error: 'Nama tipe wajib diisi' }, { status: 400 });

    const newType = await prisma.vehicleType.create({
      data: {
        name: name.toUpperCase(),
        created_by: updateBy || null,
        updated_by: updateBy || null,
      },
    });
    return NextResponse.json(
      { message: 'Tipe berhasil ditambahkan', data: newType },
      { status: 201 }
    );
  } catch (error) {
    return handleApiError(
      'Error Create Vehicle Type:',
      'Gagal menambahkan tipe kendaraan baru ke database',
      error
    );
  }
}

export async function PUT(request) {
  try {
    const { id, name, updateBy } = await request.json();
    if (!id || !name) return NextResponse.json({ error: 'Data tidak lengkap' }, { status: 400 });

    const updatedType = await prisma.vehicleType.update({
      where: { id: Number(id) },
      data: {
        name: name.toUpperCase(),
        ...(updateBy ? { updated_by: updateBy } : {}),
      },
    });
    return NextResponse.json(
      { message: 'Tipe berhasil diubah', data: updatedType },
      { status: 200 }
    );
  } catch (error) {
    return handleApiError(
      'Error Update Vehicle Type:',
      'Gagal memperbarui data tipe kendaraan di database',
      error
    );
  }
}

export async function DELETE(request) {
  try {
    const id = request.nextUrl.searchParams.get('id');
    if (!id) return NextResponse.json({ error: 'ID tidak ditemukan' }, { status: 400 });

    await prisma.vehicleType.delete({ where: { id: Number(id) } });
    return NextResponse.json({ message: 'Tipe berhasil dihapus' }, { status: 200 });
  } catch (error) {
    return handleApiError(
      'Error Delete Vehicle Type:',
      'Gagal menghapus data tipe kendaraan dari database',
      error
    );
  }
}
