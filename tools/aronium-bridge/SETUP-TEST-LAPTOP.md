# Test-laptop setup — safe schema discovery

Use a friend's Windows laptop with Aronium installed. Start with a test database, not a live Easy Buys till.

## 1. Prerequisites
- Windows 10/11.
- Python 3.11+ from python.org; select "Add python.exe to PATH".
- Aronium installed and opened once.

## 2. Configure the test source
Copy this folder locally, then copy \`config.example.json\` to \`config.json\` and set:
- \`shop_id\`: \`TEST\`
- \`source_id\`: \`TEST-TILL1\`
- \`database_path\`: the actual path shown in Aronium Settings > Database > Open database location
- \`queue_path\`: \`data\\bridge-queue.db\`
- \`schema_report_path\`: \`schema-report.json\`
- \`cloud_sync_enabled\`: \`false\`

Do not use production source IDs or cloud credentials on the friend's laptop.

## 3. Run these commands
\`\`\`bat
python --version
python bridge.py validate
python bridge.py inspect-schema
python schema_profile.py
python bridge.py init-queue
python bridge.py status
\`\`\`

The schema report contains table/column metadata. The profile report contains table counts and name-based candidate scores, not row contents.

## 4. Create synthetic transactions if possible
In a test database, create a cash sale, a card/other-payment sale, a discount sale, and a refund/void if available. Run the schema/profile commands again.

## 5. Share the reports for mapping
Share \`schema-report.json\` and \`schema-profile.json\`. Check them before sharing. Never send the actual POS database, customer records, production tokens, or passwords.

## Safety
- Cloud sync must stay disabled.
- The bridge does not extract transaction rows until a reviewed mapping is explicitly enabled.
- Never test with production tills or post test sales into the Easy Buys live system.
