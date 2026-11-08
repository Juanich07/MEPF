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
4. In Firebase Authentication, enable the **Email/Password** provider and create an account for each staff member. Sign-up is intentionally not exposed in the app.
5. Give the administrator account the `admin: true` custom claim using a trusted environment with the Firebase Admin SDK. For example:

   ```js
   import { applicationDefault, initializeApp } from "firebase-admin/app";
   import { getAuth } from "firebase-admin/auth";

   initializeApp({ credential: applicationDefault() });
   const adminUser = await getAuth().getUserByEmail("admin@example.com");
   await getAuth().setCustomUserClaims(adminUser.uid, {
     ...adminUser.customClaims,
     admin: true,
   });
   ```

   Initialize `getAuth()` with the Admin SDK and server-side credentials; never put service-account credentials or Admin SDK code in the browser app. After changing a claim, have that user sign out and sign in again.
6. Publish the rules in [`firestore.rules`](./firestore.rules) in the Firebase Console (Firestore Database → Rules) or with the Firebase CLI. Do not replace them with public read/write rules. Signed-in users can read, add, and edit equipment and reference entries; only users with the admin claim can delete equipment, pages, or reference entries. The rules also restrict `adminUsers` profiles to their owner for writing and admins for listing, and restrict `adminActivity` to admin reads and signed-in append-only writes.
7. Restart the development server after changing environment variables.

When all six Firebase values are present, the app requires sign-in and listens to each page's collection separately. It writes equipment to the matching equipment-group collection. The Index and Definition of Terms entries are stored as individual Firestore documents so delete permissions can be enforced independently from edits. Existing reference arrays are copied into that document structure on first sign-in and left in place as a migration backup. On first launch after this update, records in the old shared `assets` collection are copied into their equipment-group collections; the originals are retained as a migration backup. The Index and life expectancy views combine records from all page collections. Without Firebase settings, the app uses local storage in the current browser, and destructive actions are unavailable. The table supports search, location and phase filters, and CSV export of visible rows.

The Index page contains the equipment-group descriptions, and Definition of Terms contains MEPF field and condition definitions; authenticated users can add and edit entries, while admins can also delete them. The life expectancy page shows the same equipment records as the group tabs; adding or editing an asset there updates that shared equipment record, and admins can delete it.

Admins can use the menu button in the header to open the app user list. Select a user to view their recorded activity, including actions, changed field names, and timestamps. A user profile is created or refreshed when that person signs in, so it only includes users who have opened the app since this feature was enabled. Activity history records future app add, edit, import, and delete actions; it does not retain previous or new field values. This UI-generated history helps with visibility but is not a tamper-proof security audit. Admin roles must still be assigned through Firebase Authentication custom claims using the Admin SDK; the app menu does not change roles.

On an equipment-group page, use **Template** to download an Excel workbook containing that page's records in a table matching the on-screen register headings and life-expectancy subheadings. Fill in the blank rows to add equipment, or edit existing rows to update them. The Record ID column is hidden and must remain unchanged for existing rows; blank IDs add new rows. Before saving, the import preview shows the detected spreadsheet row and values for every column; cancel if a value is missing or appears under the wrong heading. The equipment group is set from the page you upload into. Import updates and adds rows only; it does not delete records omitted from the workbook.

Custom equipment-group pages include **Customize table**. Choose, reorder, or hide the standard inventory columns, and add named custom columns. Custom values are stored with each equipment record. The page's table, CSV export, and Excel template use the configured columns. Excel templates retain the hidden Record ID for updates; when importing into a customized page, fields omitted from that page's template are left unchanged on existing records.

## Checks

```bash
npm run lint
npm run build
```
