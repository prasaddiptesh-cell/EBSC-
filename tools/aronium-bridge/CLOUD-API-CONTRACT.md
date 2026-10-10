# Easy Buys POS ingestion API — design gate

This document records the existing cloud interface review. It is not a live endpoint and must not be treated as credentials or a production API specification.

## Existing integration

The current Google Apps Script cloud has a main-software request path:
- `main_pull`
- `main_push`
- `main_devices`

The request is authenticated against a private `MAIN_KEY` in Script Properties. The current push record envelope uses:
- `t`: lowercase table name (letters and underscores)
- `k`: record key
- `u`: updated timestamp
- `row`: record JSON object

The current code writes these records into Google Sheets tabs and records changes in a shared change-history sheet. The office software pulls all tables. This makes the existing interface a possible compatibility path, but not automatically a safe or scalable POS ingestion contract.

## Important gap discovered

The existing `main_push` path is not fully idempotent for retry safety. Re-sending a record with the same key and timestamp may rewrite the row and append another change-history entry. If the bridge sends a batch and the response is lost, retrying could create duplicate change-history events. Therefore, the bridge must NOT enable live uploads until the server side implements idempotent ingestion and batch acknowledgements.

## Proposed server-side changes before live sync

1. Add a dedicated `pos_ingest` request kind, rather than sending high-volume POS transactions through the generic record push path without safeguards.
2. Authenticate each registered bridge using a separate device token stored in Script Properties or a secure server-side registry. Do not embed the existing main-software key in a public repository or hard-code it into the desktop app.
3. Require a stable event ID and source ID for every record. A unique idempotency key should be enforced server-side; the same key and same payload returns the previous result, while same key with a different payload is rejected and audited.
4. Acknowledge each record as accepted, already processed, or rejected, including a reason for rejection.
5. Store source transaction IDs and source installation IDs separately from receipt numbers. Example sources: `S1-TILL1`, `S1-TILL2`, `S2-TILL1`.
6. Store POS records in separate tables/namespaces so they do not collide with staff, payroll, bakery or other app records.
7. Add pagination and date/source filters for queries. Do not let clients send arbitrary SQL.
8. Keep the initial pilot import-only. Do not automatically alter inventory balances, staff debts or accounting records.
9. Reconcile source totals by trading date, till, currency and payment type. Differences are exceptions, not automatic adjustments.

## Suggested initial record types

- `pos_sales`: source transaction/document header, status, timestamps, cashier, gross/net/tax/discount totals, currency, linked original sale for returns.
- `pos_sale_lines`: transaction key, source line ID, product mapping, quantity, unit price, discount/tax, line total.
- `pos_payments`: transaction key, payment method, amount, tender/change if available.
- `pos_products`: source product ID, barcode/SKU, name, tax code, active state.
- `pos_stock_movements`: only when the source schema and movement semantics are verified.
- `pos_customers`: only necessary fields and only where permitted; customer balances must not be inferred as employee debt.
- `pos_cashiers`: source cashier ID and a controlled mapping to Easy Buys staff IDs where approved.

Use a deterministic event key, for example `S1-TILL1:sale:<source-transaction-id>`, with separate keys for lines, payments, refunds and product updates. Do not use receipt number alone.

## Deployment gates

- Real schema reports from all three till databases reviewed.
- Field mapping verified against known test sales, returns, voids and end-of-day reports.
- API idempotency and partial-batch behaviour tested.
- Source token protection, TLS, logs and retry behaviour reviewed.
- Duplicate replay test passes.
- Sales/payment totals reconcile on a test trading day.
- Explicit approval before enabling a production source.
