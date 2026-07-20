import json
from pathlib import Path

from sqlalchemy import select

from database import AsyncSessionLocal
from models import EquipmentDB


_SEED_PATH = Path(__file__).with_name("equipment_seed.json")
_FIELDS = (
    "equipment_code",
    "instrument_name",
    "model",
    "manufacturer",
    "manufacture_date",
    "installation_date",
    "verification_type",
    "purchase_date",
    "earliest_calibration_date",
    "latest_calibration_date",
    "next_calibration_date",
    "next_verification_date",
    "verification_cycle",
    "remarks",
)


async def seed_equipment() -> int:
    records = json.loads(_SEED_PATH.read_text(encoding="utf-8"))
    async with AsyncSessionLocal() as db:
        result = await db.execute(select(EquipmentDB.sequence))
        existing_sequences = set(result.scalars())
        inserted = 0
        for record in records:
            sequence = record["sequence"]
            if sequence in existing_sequences:
                continue
            db.add(
                EquipmentDB(
                    sequence=sequence,
                    **{field: record.get(field) for field in _FIELDS},
                )
            )
            inserted += 1
        if inserted:
            await db.commit()
        return inserted
