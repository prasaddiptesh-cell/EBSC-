# Easy Buys Aronium Sync Bridge

Foundation for three independent Aronium installations:
- Shop 1, Drasa Avenue: S1-TILL1 and S1-TILL2 (separate databases)
- Shop 2, V M Pillay Road: S2-TILL1

## Status
The dedicated cloud ingestion/query API and a queue upload client are now present on this development branch, but they have not yet been deployed to your Google Apps Script project or tested against real POS data. The bridge does not modify Aronium. Automatic extraction of actual sales still depends on verified schema mapping.

## Included
- Python 3.11+ foundation with standard library only
- Read-only SQLite connection and schema-only inspection
- Durable local queue database with unique idempotency keys
- Local validation and status commands
- Cloud sync disabled by default

## Setup on each till
1. Copy this folder to a local folder such as C:\EasyBuys\AroniumBridge.
2. Copy config.example.json to config.json.
3. Set the actual database path, shop ID and unique source ID.
4. Run: python bridge.py validate
5. Run: python bridge.py inspect-schema
6. Run: python bridge.py init-queue
7. Run: python bridge.py status

Aronium Lite commonly uses %LOCALAPPDATA%\Aronium\Data\pos.db. Confirm the path inside Aronium at Settings, Database, Open database location. Do not assume every installation uses the default.

## Safety
- Source database is opened read-only.
- Do not place the live POS database on a network share.
- Never write imported data back into Aronium.
- Do not enable cloud sync until the existing Easy Buys API and data model are reviewed.
- Receipt numbers alone are not unique across separate tills.
- Do not auto-adjust stock, payroll or financial balances during the pilot.
- Never commit API keys, passwords or device tokens.

## Planned next stages
1. Inspect schema reports from all three tills.
2. Map sales, sale lines, payments, refunds, products, customers and cashiers.
3. Review the existing Easy Buys cloud API and data model.
4. Implement versioned authenticated ingestion with idempotency, retries and audit logs.
5. Reconcile imported totals against Aronium reports.
6. Pilot one till in a test area before rolling out.
