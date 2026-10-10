# Deployment Plan

## Source identities
- S1-TILL1: Shop 1, Drasa Avenue, first independent database
- S1-TILL2: Shop 1, Drasa Avenue, second independent database
- S2-TILL1: Shop 2, V M Pillay Road

Never reuse a source ID for two databases.

## Foundation checks
- Confirm Aronium version and database type on all three computers.
- Confirm each database path independently and back up each database.
- Run validate and inspect-schema for each installation.
- Compare schema reports; do not assume all databases have identical schemas.
- Cloud upload remains disabled in this foundation.

## Data mapping required
Identify and test the source relationships for sales headers, sale lines, payments, VAT/tax, discounts, refunds, returns, voids, products, barcodes, customers, customer balances, cashiers, register IDs, stock movements and document status/timestamps. Confirm business meaning with test transactions and Aronium reports; do not infer from table names alone.

## Cloud integration gate
Review the existing Easy Buys API and data model before defining a versioned authenticated ingestion endpoint, source-specific idempotency, schema validation, retry/backoff, durable outbox, audit log, exception handling and daily reconciliation.

First live pilot must use a separate test area. Do not automatically deduct stock, alter payroll/staff debt or post financial adjustments until source mapping and reconciliation are approved.
