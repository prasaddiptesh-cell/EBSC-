# Easy Buys cloud – one-time setup

Do this **once**, on a computer, signed in to the Easy Buys Google account. It takes about 15 minutes.
After this you never need to open Google again: new tablets and phones are added from inside the app
(Admin › Devices › Add a device), and new app sections (Invoices, Till & Cash, Payroll…) need no change here.

The app works offline without this. Nothing is lost before setup: the app keeps everything and sends it once linked.

## What the cloud does
- Keeps every linked tablet and phone in step: staff, Time Clock, timesheets, leave, payroll, invoices & cheques,
  till counts, deposits, bank reconciliation and every future section.
- Reads invoice photos (Google's free text reading) so the app can fill in supplier, number, date, total and VAT.
- Lists linked devices; a lost phone can be blocked from the app.
- Saves bakery shift reports, purchase orders, photos and daily backups in Google Drive (folder **Easy Buys Bakery Cloud**) and emails them.
- Live online dashboard (password protected).
- Hands everything to the main Easy Buys software on the office PC, with its own private key.

## A. You already set up the Bakery cloud
1. Open **https://script.google.com** › your **Easy Buys Bakery Cloud** project.
2. Before pasting, **copy the two values** at the top of `Code.gs` (`API_KEY` and `DASHBOARD_PASSWORD`) somewhere safe.
3. Select everything in `Code.gs`, delete it, and paste the whole new **Code.gs**. Put your two values back in. Click 💾 **Save**.
4. On the left click **Services ➕**, choose **Drive API**, click **Add** (this lets Google read invoice photos).
   Then in the function list choose **setup** and click **Run**. Allow the new permission (Google Sheets) when asked:
   choose the account › **Advanced › Go to Easy Buys Bakery Cloud (unsafe)** › **Allow**.
   The log shows the **main-software key**: keep it private, it is only needed later for the office PC.
5. **Deploy › Manage deployments › ✏️ Edit › Version: New version › Deploy.** The web address stays the same,
   so the bakery tablet keeps working without any change.

## Updating to app version 3.5 (once)
Version 3.5 adds "Send to support" problem reports and lets Admin › Backup & restore list and bring back the daily
backups kept in Google Drive. The cloud needs the new **Code.gs** once – same 3 steps as before:
1. **https://script.google.com** › your Easy Buys cloud project › copy `API_KEY` and `DASHBOARD_PASSWORD` somewhere safe.
2. Replace everything in `Code.gs` with the new **Code.gs**, put the two values back, 💾 **Save**.
3. **Deploy › Manage deployments › ✏️ Edit › Version: New version › Deploy.** The web address stays the same.
No new permissions are needed. Devices do **not** need to be linked again. The app keeps working with the old script
until you do this (only "Send to support" and the Drive backup list wait for it).

## Updating to app version 3.4 (once)
Version 3.4 keeps documents (licence copies, applications, your company contract PDF, signed contracts) in
Google Drive › **Easy Buys › Documents**, so any linked device can open them. The cloud needs the new **Code.gs** once:
1. **https://script.google.com** › your Easy Buys cloud project › copy `API_KEY` and `DASHBOARD_PASSWORD` somewhere safe.
2. Replace everything in `Code.gs` with the new **Code.gs**, put the two values back, 💾 **Save**.
3. **Deploy › Manage deployments › ✏️ Edit › Version: New version › Deploy.** The web address stays the same.
No new permissions are needed. Devices do **not** need to be linked again.

Security (recommended): turn on **2-Step Verification** on the Google account the cloud runs on
(myaccount.google.com › Security), keep the cloud web address private, and block a lost device at once in
**Admin › Devices**.

## Updating to app version 3.2 (once)
Version 3.2 adds emails to new staff (welcome + signed contract), Viber Channel alerts, the roster in Google Calendar
and full daily backups. The cloud needs the new **Code.gs** once – devices do **not** need to be linked again:
1. **https://script.google.com** › your Easy Buys cloud project › copy `API_KEY` and `DASHBOARD_PASSWORD` somewhere safe.
2. Replace everything in `Code.gs` with the new **Code.gs**, put the two values back, 💾 **Save**.
3. Choose **setup** › **Run** › allow the new permissions when Google asks (Calendar, sending email, connecting to Viber).
4. **Deploy › Manage deployments › ✏️ Edit › Version: New version › Deploy.** The web address stays the same.

Optional, only if you want automatic Viber alerts: in Viber create a **Channel** (e.g. "Easy Buys Alerts"), add the people
who should get alerts, open the channel's **Settings › Developer tools**, copy the token, then in the app
**Admin › Verification & Viber alerts › Automatic (Viber Channel)** › paste it › **Connect**. Without this, alerts are
shared to your normal Viber group with one tap.

## B. First time
1. **Make the key in the app.** On the shop tablet log in as Admin › **Admin › Online & email** › **Create a key** › **Save settings**.
2. Go to **https://script.google.com** › **New project** › name it *Easy Buys Bakery Cloud*.
3. Paste **Code.gs** into `Code.gs`. Click **+ › HTML**, name it `Dashboard` (exactly) and paste **Dashboard.html**.
4. At the top of `Code.gs` set `API_KEY` (the key from step 1) and `DASHBOARD_PASSWORD`. Click 💾 **Save**.
5. On the left click **Services ➕** › **Drive API** › **Add** (lets Google read invoice photos).
   Then choose **setup** › **Run** and allow the permissions (as in A4). Keep the main-software key private.
6. **Deploy › New deployment** › ⚙️ › **Web app** › Execute as **Me**, Who has access **Anyone** › **Deploy**.
   Copy the **Web app URL** (starts with `https://script.google.com/macros/s/`, ends with `/exec`).
7. On the tablet: **Admin › Online & email** › paste the URL › **Save settings** › **Test connection** (✅ Connected).

That's it. The shop tablet is the first device.

## Adding the other tablets and phones (no Google needed)
On the shop tablet (or any linked device), as Owner or Admin: **Admin › Devices › Add a device**, then either
- **QR code**: on the new device open Easy Buys › **Link to Easy Buys** › **Scan the QR code**;
- **Email**: type their email; they open it on the new device and tap the link › **Open in the Easy Buys app**;
- **WhatsApp / message**: send the message; they tap the link the same way, or paste the code into **Link to Easy Buys**.

Each invite works once and expires after 48 hours. The cloud key never leaves the shop tablet; each device gets its own access,
so a lost phone can be **blocked** in Admin › Devices without affecting anyone else. Owners can also switch on
**New devices need my approval** there.

New devices drop the sample logins (Admin 1111 etc.) as soon as your real staff arrive from the cloud.

## Good to know
- Gmail sends up to about 100 emails a day from a free account.
- The data spreadsheet **Easy Buys Data (do not edit by hand)** is the sync store. Look, but don't type in it.
- Backups hide PINs and passwords. A full backup of every device is saved each day in Google Drive › Backups.
- Emails go only to staff being onboarded (their address is in the app); a device cannot email anyone else.
