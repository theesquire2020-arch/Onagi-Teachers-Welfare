# Coast Region Onagi Teachers Welfare

Google Apps Script + Google Sheets welfare application for teachers in Mombasa, Kwale, Kilifi, Lamu, Tana River and Taita Taveta. Currency: KES.

## Files

- `Code.gs`: backend, spreadsheet setup, sessions, access control, loans, payments and audit logging.
- `Index.html`, `Styles.html`, `JavaScript.html`: responsive SPA. Chart.js loads from CDN.
- `appsscript.json`: V8 runtime and scopes.

## Deploy to Google Apps Script

1. Create a **new Google Sheet**. Open **Extensions → Apps Script** from that Sheet (the script must be bound to it).
2. Replace the default `Code.gs` with this file. Add three HTML files named **Index**, **Styles**, and **JavaScript** (paste the contents of the matching `.html` files). Enable the manifest file in Project Settings and replace `appsscript.json`.
3. Select `setupDatabase` in the Apps Script editor and click **Run** to authorize access before deploying. The web app also checks on first visit and automatically creates missing sheets and headers in the bound (or configured) database, including the new `Expenditures` tab. If your Apps Script project is **not bound to a Sheet**, put the actual Google Sheet ID in the `SPREADSHEET_ID` Script Property first. Setup creates the ledger tabs plus five dedicated contribution matrices: `Saving`, `Hospital Bill`, `Benevolence`, `Education`, and `Charity`. Running it again preserves receipts, adds missing `Gender`/`Category`/`BeneficiaryID` columns to older ledgers, and rebuilds the matrices with footer totals.
4. In **Project Settings → Script Properties**, set payment credentials and callback properties as needed (below). Never put API keys in HTML or the spreadsheet.
5. Click **Deploy → New deployment → Web app**. Set **Execute as: Me** and **Who has access: Anyone** for public registration. Copy the `/exec` URL. Use a Google Workspace account and restrict spreadsheet sharing to trusted administrators only.
6. **Redeploy an updated version** after replacing these files; an existing `/exec` deployment continues serving its previous version until redeployed. Sign in and click the database badge in the navbar, or **Settings → Test connection**. It verifies the actual active Google Sheet and required tabs/headers; it does not merely check that a deployment URL exists. In Settings, SuperAdmin can label and save up to 20 compatible Google Sheets, then switch between saved Sheets. Registry entries are stored in each active Sheet's `Settings` tab and copied to the target on switch; records themselves are **not** migrated. Back up before switching. Also test registration, roles, loans and provider callbacks in sandbox before accepting payments.

The requested `https://script.google.com/macros/s/AKfycbyd4zRrxgvSHNmTBF3FMhMDPRE5RsUt6VVpZ3gzQBv-VOwVSj7iuggqsqFjkbK99jfHIw` is prefilled as the **Apps Script deployment URL** for reference and future updates. It is **not a Google Sheet URL** and cannot serve as `SPREADSHEET_ID`. Use a URL beginning `https://docs.google.com/spreadsheets/d/` to register a Sheet database. Add `/exec` when opening a deployed Apps Script web app; save the new deployment URL in Settings after redeploying.

## Required payment configuration

For M-Pesa Daraja STK Push set these Script Properties:

- `MPESA_CONSUMER_KEY`, `MPESA_CONSUMER_SECRET`
- `MPESA_SHORTCODE`, `MPESA_PASSKEY`
- `MPESA_BASE_URL` (`https://sandbox.safaricom.co.ke` for sandbox, use the production endpoint when approved)
- `PAYMENT_CALLBACK_SECRET` (long random secret)
- `PAYMENT_CALLBACK_URL` (your deployed `/exec` URL with `?key=<PAYMENT_CALLBACK_SECRET>` appended)

The callback must be publicly reachable by Safaricom. Reconcile test payments with actual provider reports before going live. A successful, authenticated callback with a provider receipt changes a pending transaction to paid. Failed callbacks change it to failed. Never record a phone PIN in this app.

For Airtel Money, set `AIRTEL_PAYMENT_URL` to your **approved merchant gateway's HTTPS collection endpoint** and `AIRTEL_BEARER_TOKEN` to its token. Merchant payloads and callback authentication vary between Airtel integrations; verify the JSON adapter in `airtelPush_` and callback contract with your provider before live use. The app fails closed when these settings are missing.

## Sign-in and roles

- Public and private teachers register with full name, title preference (Male → Mr., Female → Madam, Other → Teacher), National ID, TSC number, phone, email, password, school type, county, sub-county, school name, and school phone. Each gets a sequential `CRT-0001`-style member ID and signs in with the registered email/password. Existing members without a saved title receive the neutral Teacher greeting.
- Member emails and password hashes are read from the `Members` sheet at login. Passwords are SHA-256 with a unique random salt; plaintext passwords are not stored in the Google Sheet.
- The requested embedded Super Admin account is `admin@gmail.com` / `chief001`, checked before the Members sheet. **This is an intentionally specified default and is unsafe for a public production site. Rotate/change this constant before deployment.** Do not share it with members.
- Members can view their own financial records; Admin can manage members and financial data but cannot assign roles; SuperAdmin can also assign roles and edit settings. Tokens expire after eight hours. Users can sign out.

## Data & operational notes

The database retains the `Contributions` receipt ledger and has five distinct **matrix tabs**: `Saving`, `Hospital Bill`, `Benevolence`, `Education`, and `Charity`. In both the Sheet and frontend, each table has Member ID, registered member name in column two, dynamic columns, a row total, and a **TOTAL** row at the bottom. Every registered member gets a blank row in all five matrices. Saving columns are automatic Weekly/Monthly periods. In the other four funds, **Add beneficiary** creates a column headed by that person's name; select that beneficiary when contributing. Only paid contributions enter a matrix. The Admin-only Expenditure panel records category, amount, purpose, payee and an automatic reference in its own Sheet tab. Dashboard bars compare all five fund totals with different colours. Use spreadsheet access controls and backups. Manual contributions receive automatic `RCPT-CON-…` references; mobile payments receive provider receipts only after confirmation.

The Vercel application now keeps registration, login and every authenticated panel **inside the application**. Its same-origin `api/gas.js` bridge forwards requests server-side to the deployed Apps Script `/exec` endpoint and follows Google's ContentService redirect; the browser never needs cross-origin access or a new tab. A successful registration is saved by Apps Script to the active Google Sheet, then the user signs in on the application's login panel with those credentials. SuperAdmin signs in on the same panel. The bridge fails closed and shows an error if the deployment is unavailable; it does **not** silently fall back to sample data. Sample records are accessible only via the explicit **Explore sample demo** choice. If you create a new Apps Script deployment ID, update `GAS_WEBAPP_URL` in Vercel (or the fallback URL in `api/gas.js`) and redeploy the Vercel app. Connection tests show a circling indicator and confirmed status; Settings offers a per-device light/dark appearance choice.
