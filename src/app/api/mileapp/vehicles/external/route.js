import prisma from '@/lib/prisma';
import { NextResponse } from 'next/server';

export async function GET(request) {
  const { searchParams } = new URL(request.url);
  const hubId = searchParams.get('hubId');
  const startDate = searchParams.get('startDate');
  const endDate = searchParams.get('endDate');

  try {
    const data = await prisma.vehicleExternal.findMany({
      where: {
        hub_id: hubId,
        date: { gte: startDate, lte: endDate },
      },
    });
    const formattedData = data.map((d) => ({
      ...d,
      hubId: d.hub_id,
      storageType: d.storage_type,
      vehicleType: d.vehicle_type,
      createdAt: d.created_at,
      updatedAt: d.updated_at,
    }));
    return NextResponse.json(formattedData);
  } catch (error) {
    console.error('Error Get Truck Usage:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      { error: 'Gagal mengambil data penggunaan truk dari database', detail: errorMessage },
      { status: 500 }
    );
  }
}

export async function POST(request) {
  try {
    const body = await request.json();
    const hubId = body.hubId || body.hub_id;
    const date = body.date;
    const storageType = body.storageType || body.storage_type;
    const vehicleType = body.vehicleType || body.vehicle_type;
    const count = body.count;
    const description = body.description;

    if (!hubId || !date || !storageType || !vehicleType) {
      return NextResponse.json(
        { error: 'Field hubId, date, storageType, dan vehicleType wajib diisi' },
        { status: 400 }
      );
    }

    const parsedCount = parseInt(count, 10) || 0;

    const upserted = await prisma.vehicleExternal.upsert({
      where: {
        hub_id_date_storage_type_vehicle_type: {
          hub_id: hubId,
          date,
          storage_type: storageType,
          vehicle_type: vehicleType,
        },
      },
      update: { count: parsedCount, description },
      create: {
        hub_id: hubId,
        date,
        storage_type: storageType,
        vehicle_type: vehicleType,
        count: parsedCount,
        description,
      },
    });
    const formatted = {
      ...upserted,
      hubId: upserted.hub_id,
      storageType: upserted.storage_type,
      vehicleType: upserted.vehicle_type,
      createdAt: upserted.created_at,
      updatedAt: upserted.updated_at,
    };
    return NextResponse.json({ message: 'Success', data: formatted });
  } catch (error) {
    console.error('Error Upsert Truck Usage:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      { error: 'Gagal menyimpan atau memperbarui data penggunaan truk', detail: errorMessage },
      { status: 500 }
    );
  }
}

export async function DELETE(request) {
  try {
    const body = await request.json();
    const hubId = body.hubId || body.hub_id;
    const date = body.date;
    const storageType = body.storageType || body.storage_type;
    const vehicleType = body.vehicleType || body.vehicle_type;

    if (!hubId || !date || !storageType || !vehicleType) {
      return NextResponse.json(
        { error: 'Field hubId, date, storageType, dan vehicleType wajib diisi' },
        { status: 400 }
      );
    }

    await prisma.vehicleExternal.delete({
      where: {
        hub_id_date_storage_type_vehicle_type: {
          hub_id: hubId,
          date,
          storage_type: storageType,
          vehicle_type: vehicleType,
        },
      },
    });

    return NextResponse.json({ message: 'Success' });
  } catch (error) {
    console.error('Error Delete Truck Usage:', error);
    const errorMessage =
      error instanceof Error ? error.message : 'Kesalahan sistem tidak diketahui';
    return NextResponse.json(
      { error: 'Gagal menghapus data penggunaan truk dari database', detail: errorMessage },
      { status: 500 }
    );
  }
}
