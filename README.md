# MEPF Asset Register

A spreadsheet-style equipment inventory for MEPF facilities. The register includes the equipment, location, capacity, life expectancy, replacement, zoning, and phasing fields from the supplied sheet.

## Run locally

```bash
npm install
npm run dev
```

Open the local URL printed by Next.js. Without Firebase settings, the app loads sample equipment and saves changes to this browser's local storage.

## Firebase setup

1. Create a Firebase project and a Cloud Firestore database.
2. Copy `.env.example` to `.env.local` and fill in the Firebase web app configuration values.
3. In Firestore, create a database. The app creates a separate top-level equipment collection for each equipment group (for example, `HVAC-1` and `HVAC-2`), plus `workbookPages` for custom page names and `system` for its initialization marker.
4. Configure Firestore security rules to permit the app's read, create, update, and delete operations on the equipment-group collections, and read/create/update operations on `workbookPages` and `system`. The `system` collection stores the initialization marker plus the editable Index and Definition of Terms data. The current deny-all rules block all syncing. Do not use public read/write rules for real facility data; this client-side sample does not include user authentication.
5. Restart the development server after changing environment variables.

When all six Firebase values are present, the app listens to each page's collection separately and writes equipment create, update, and delete operations to the matching equipment-group collection. The Index and life expectancy views combine records from all page collections. Custom pages are saved to `workbookPages` and sync across clients. On first launch after this update, records in the old shared `assets` collection are moved into their equipment-group collections. Without Firebase settings, the app uses local storage in the current browser. The table supports search, location and phase filters, and CSV export of visible rows.

The Index page contains the equipment-group descriptions, and Definition of Terms contains MEPF field and condition definitions; both can be added to, edited, and deleted. Those reference pages save to Firestore when configured and to this browser otherwise. The life expectancy page shows the same equipment records as the group tabs; adding, editing, or deleting an asset there updates that shared equipment record.

On an equipment-group page, use **Template** to download an Excel workbook containing that page's records in a table matching the on-screen register headings and life-expectancy subheadings. Fill in the blank rows to add equipment, or edit existing rows to update them. The Record ID column is hidden and must remain unchanged for existing rows; blank IDs add new rows. Import updates and adds rows only; it does not delete records omitted from the workbook. The equipment group is taken from the page you upload into.

## Checks

```bash
npm run lint
npm run build
```
