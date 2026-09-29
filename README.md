# METRA / Smart NAWI — OIML R-76 Test Report Platform

React + Vite frontend with a dependency-free Node.js backend (`server.cjs`).
Base: METRA-FINAL-PS35-Integrated, with the QR renderer, build config and demo tool from smart-nawi-platform merged in.

## Run

```bash
npm install
npm run server     # terminal 1 – API on http://localhost:8787
npm run dev        # terminal 2 – app on http://localhost:5173
```

Production build: `npm run build && npm run preview`.
If the API runs elsewhere, copy `.env.example` to `.env` and set `VITE_API_BASE`.

## Demo accounts (change before any real deployment)

All seed accounts share one password so the table below is accurate for every row; change every password before any real deployment.

| Role | Email | Password |
|---|---|---|
| Technician | r.sharma@nawi-lab.gov.in | METRA@123 |
| Technician | p.nair@nawi-lab.gov.in | METRA@123 |
| Examiner | n.verma@nawi-lab.gov.in | METRA@123 |
| Authority | s.iyer@nawi-lab.gov.in | METRA@123 |
| Admin | admin@nawi-lab.gov.in | METRA@123 |
| Viewer | viewer@nawi-lab.gov.in | METRA@123 |

These are created the first time `server.cjs` runs, in `server-data/users.json`. If that file already exists from an older run (old passwords `Navi@123` / `Admin@123` / `Viewer@123`), delete `server-data/` once so it regenerates with the table above.

## What is included

- Session auth (HttpOnly cookie, scrypt password hashes), server-side RBAC per collection, append-only audit validation
- Role-scoped metragation (`ROLE_NAV`) and per-role dashboards
- OIML R-76 class-aware MPE calculation, test plans, evidence, environmental and reference-equipment traceability
- Immutable issued-report snapshots, content-based integrity fingerprint
- PDF export (jsPDF) and editable DOCX export, with QR embedded
- QR via `qrcode.react` (standard encoder); payload `NAVI:R:<REPORT_ID>`
- Public verification endpoint `GET /api/public/reports/<REPORT_ID>` (verification-safe metadata only)
- Self-service profile update `PUT /api/auth/profile`
- **Run Demo Evaluation** button on the dashboard (roles that can register instruments); demo data is labelled DEMO

## Merge changes vs FINAL-PS35

1. Hand-built QR encoder removed; `qrcode.react` used on screen and rasterised for the PDF (`downloadPdf` is now async).
2. Added `vite.config.js`, `.gitignore`; pinned dependencies (React 18.3, Vite 5.4) instead of `latest`.
3. Ported the Run Demo Evaluation generator from smart-nawi.
4. README rewritten; package renamed.

## Not yet verified

The build has not been run in the environment where this was assembled (no npm access). `server.cjs` and `App.jsx` pass syntax checks only. Run `npm install && npm run build` and test: login per role, QR scan of an issued report, PDF/DOCX download, demo button.

## Production boundary

Prototype architecture. Production needs a transactional database, durable session store, HTTPS, rate limiting, object storage for evidence, backups, automated RBAC/OIML tests, and metrology-owner validation of the OIML tables.

## Local setup (Windows / VS Code)

Open this folder directly (the folder containing `package.json`, `server.cjs`, `src`, and `vite.config.js`).

```powershell
npm install
```

Terminal 1 — backend:

```powershell
node server.cjs
```

Terminal 2 — frontend:

```powershell
npm run dev
```

Open `http://localhost:5173/`. The backend listens on `http://localhost:8787`.

### Demo accounts

- Technician: `r.sharma@nawi-lab.gov.in` / `METRA@123`
- Examiner: `n.verma@nawi-lab.gov.in` / `METRA@123`
- Authority: `s.iyer@nawi-lab.gov.in` / `METRA@123`
- Admin: `admin@nawi-lab.gov.in` / `METRA@123`
- Viewer: `viewer@nawi-lab.gov.in` / `METRA@123`

The first backend run creates `server-data/users.json` and the local repository.


## Phase 1 fixes (PS 26035)
- Restored the missing central QR token helpers used by report rendering/PDF generation and verification.
- Added explicit PDF/DOCX error feedback instead of silent export failures.
- API base now follows the browser hostname for same-LAN demonstrations unless VITE_API_BASE is set.
- Vite is configured to listen on the LAN for phone QR demonstrations.
- Final report issuance remains restricted to Authority/Admin as part of RBAC; Technicians can prepare and submit tests but cannot issue the final report.


## Phase 2 PS 26035 hardening
- Strengthened class/Max/e/Min validation so incompatible configurations block test submission.
- Added verification-scale-interval alignment validation for entered test loads.
- Added numeric guards to accuracy/repeatability calculations so invalid inputs cannot silently produce a PASS/FAIL result.
- Existing rule-version IDs remain attached to calculation records for traceability.

## Phase 3 PS 26035 hardening (final)
- **MPE breakpoints were wrong for three of four accuracy classes.** The Phase 2 engine applied one 500e/2000e table (OIML R 76-1:2006 Table 6, Class III) to every class. It is now per class — Class I: 50 000e/200 000e, Class II: 5 000e/20 000e, Class III: 500e/2 000e, Class IIII: 50e/200e — configured per class in `OIML_CLASS_TABLE` and versioned per class in the Rules page. The active rule version is `R76-CFG-2.1`; `R76-CFG-2.0` is kept, marked SUPERSEDED, and its notes flag that Class I/II/IIII results issued under it need re-evaluation. Class III results are unaffected, since 500e/2000e was already correct for that class.
- Added the coarse-e n-range and Min-capacity band from R 76-1 Table 3 (Class II: e ≥ 0.1 g uses n ≥ 5000, Min 50e; Class III: e ≥ 5 g uses n ≥ 500, Min 20e) to class/e/Min validation.
- **Login did not check the password.** `POST /api/auth/login` now always verifies the submitted password against the account's stored scrypt hash (`verifyPassword`); a wrong password or an unrecognized email returns 401, and 5 failed attempts for one email lock it out for 5 minutes. Unknown-email auto-registration (previously always on, and the reason any string containing "admin" logged in as Admin) is now gated behind an explicit `DEMO_LOGIN=1` environment variable that defaults to **off**; leave it unset for anything other than a supervised demo.
- The demo account table above and the passwords actually seeded by `server.cjs` were out of sync (README said `METRA@123`; the server seeded `Navi@123`/`Admin@123`/`Viewer@123`). Both now seed `METRA@123` for every account.
- PDF and DOCX report exports now include a per-observation table for every test (load, observed indication, error, applicable MPE, pass/fail per point), not only the one-line test summary.
- `PUT /api/repository` now rejects any change to a report already marked `Issued` (edit or delete) — issued reports are server-enforced immutable; a correction must be a new report revision.
- `GET /api/public/reports/<id>` now only resolves reports with `status==='Issued'`, so an in-progress or draft report cannot be looked up through the public verification endpoint.

## Demo login
Login now requires the correct password for a known account (see Demo accounts above) — there is no more "any password" mode. Setting `DEMO_LOGIN=1` on the server additionally lets an unrecognized email/username self-register on first login (role inferred from `admin`/`authority`/`examiner`/`viewer` in the address, else Technician); do this only for a supervised jury/demo session, since it lets anyone create their own account.

## Still open (not fixed in this pass)
- **Digital signatures** are not implemented; the report carries a content-integrity hash only, as the UI states.
- **Report exports embed photos as a count, not as images.** Evidence photos are listed with type labels but not inlined into the PDF/DOCX.
- **Tilt, EMC/disturbance and durability tests** from OIML R 76 are not in the 12-test suite (accuracy, repeatability, discrimination, eccentricity, sensitivity, zero-return, creep, tare, temperature, warm-up, voltage, span).
- **No standalone architecture/calculation-methodology document** in this zip; there is an in-app Technical Documentation page only.
- Zero-return tolerance (0.25e), the discrimination test's 1e reference, and the class n-ranges are marked in code comments as needing sign-off from a metrologist against your specific OIML R 76 edition before this is presented as final to evaluators.
