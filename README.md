# bbz Führung SPA

Single-Page-App zum Management von Führungsmeetings bei bbz st.gallen ag.

**Live**: https://markusbaechler.github.io/bbz-Fuehrung/

## Stack

- Vanilla JavaScript (kein Framework)
- MSAL.js 3.x für Azure AD Authentication
- Microsoft Graph Workbook API für Excel-Zugriff
- Hosted auf GitHub Pages

## Views

| View | Status | Excel-Tabelle |
|---|---|---|
| Dashboard | ✅ stabil | (aggregiert) |
| FüSi Operativ | ✅ stabil | `tbl_FuesiOperativ` |
| EarlyBird | ✅ stabil | `tbl_EarlyBird` |
| Strategietage | ✅ stabil | `tbl_Strategietage` |
| Offsite | ✅ stabil | `tbl_Offsite` |
| Teamtage | 🚧 TODO | `tbl_Teamtage` |
| bbz Agenda | 🚧 TODO | `tbl_Agenda` |

## Lokale Entwicklung

```bash
python -m http.server 3000
```

Dann `http://localhost:3000` öffnen.

## Deployment

Push auf `main` → GitHub Pages deployed automatisch.

## Konfiguration

Siehe `config.js` für Tenant-IDs, SharePoint-Pfade und Excel-Tabellen-Namen.

Azure Redirect-URIs müssen folgende beinhalten:
- `http://localhost:3000` (Entwicklung)
- `https://markusbaechler.github.io/bbz-Fuehrung/` (Produktion)

## Dateistruktur

| Datei | Zweck |
|---|---|
| `index.html` | UI, alle Views, CSS, JS |
| `store.js` | State, Graph-Zugriff, CRUD |
| `graph.js` | Workbook-API-Wrapper |
| `auth.js` | MSAL-Wrapper |
| `config.js` | Tenant-IDs, Tabellennamen, Spalten |
| `msal-browser.min.js` | MSAL lokal (Brave blockt CDN) |
| `.nojekyll` | GitHub Pages: Jekyll deaktivieren |

## Lizenz

Intern – bbz st.gallen ag
