"""Dev-helper: bouwt een synthetisch 'vorige-week' snapshot op basis van
het meest recente snapshot, met enkele gemuteerde statussen/eigenaars/
deadlines. Hiermee kan de diff-detectie in monitor.py worden getest
zonder het Excel-bestand aan te raken.

Werking:
1. Draai eerst monitor.py op het dummy-bestand. Dit produceert
   data/snapshots/<vandaag>_snapshot.json.
2. Draai dit script: python tools/simulate_week.py
   Het schrijft data/snapshots/<7 dagen geleden>_snapshot.json met
   gemuteerde waardes.
3. Draai monitor.py opnieuw -> rapport toont diff t.o.v. de synthetische
   vorige week.

Dit script is ALLEEN voor ontwikkeling/test. Niet gebruiken in productie.
"""

from __future__ import annotations

import copy
import json
import sys
from datetime import date, timedelta
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SNAPSHOT_DIR = PROJECT_ROOT / "data" / "snapshots"


def latest_snapshot() -> Path | None:
    if not SNAPSHOT_DIR.exists():
        return None
    snaps = sorted(SNAPSHOT_DIR.glob("*_snapshot.json"))
    return snaps[-1] if snaps else None


def main() -> int:
    src = latest_snapshot()
    if src is None:
        print(
            "[FOUT] Geen snapshot gevonden in data/snapshots/. "
            "Draai eerst monitor.py.",
            file=sys.stderr,
        )
        return 1

    with src.open("r", encoding="utf-8") as f:
        snap = json.load(f)

    fake = copy.deepcopy(snap)
    fake_date = date.today() - timedelta(days=7)
    fake["snapshot_date"] = fake_date.isoformat()

    impls = list(fake["implementations"].items())
    mutations = 0

    # Muteer de eerste drie implementaties zodat diff-detectie iets te
    # vergelijken heeft. Wijzigingen zijn fictief en alleen bedoeld voor
    # diff-test, niet voor data-analyse.
    for idx, (kn, impl) in enumerate(impls):
        if idx >= 3:
            break
        original_status = impl.get("status")
        impl["status"] = "Vorige status"
        impl["overall_status"] = "Vorige status"
        if idx == 0:
            impl["owner"] = "Iemand anders"
        if idx == 1 and impl.get("go_live"):
            impl["go_live"] = (date.today() - timedelta(days=30)).isoformat()
        mutations += 1
        _ = original_status

    # Verwijder de laatste implementatie om "verdwenen" te triggeren.
    if len(impls) > 3:
        removed_kn = impls[-1][0]
        fake["implementations"].pop(removed_kn)

    target = SNAPSHOT_DIR / f"{fake_date.isoformat()}_snapshot.json"
    with target.open("w", encoding="utf-8") as f:
        json.dump(fake, f, indent=2, ensure_ascii=False, default=str)

    print(f"[INFO] Synthetisch snapshot geschreven: {target.name}")
    print(f"[INFO] Mutaties: {mutations} status-wijzigingen + 1 verdwenen item.")
    print("")
    print("Volgende stap:")
    print("  python monitor.py")
    print("Het nieuwe rapport toont diff t.o.v. de synthetische vorige week.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
