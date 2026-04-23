# Deployment-Anleitung (einmalig)

## 1. Azure Redirect-URI eintragen

Vor dem Deploy: Azure Portal → **App-Registrierungen** → `bbz-Fuehrung-SPA` → **Authentifizierung**

Unter Plattform **Single-Page-Anwendung** ergänzen:
```
https://markusbaechler.github.io/bbz-Fuehrung/
```

**Wichtig**: Trailing Slash. `http://localhost:3000` bleibt für lokales Testen.

## 2. GitHub Repo vorbereiten

Falls noch nicht existierend: Repo `bbz-Fuehrung` erstellen (public oder private je nach Lizenz).

```bash
# Im lokalen Projekt-Ordner C:\Benutzer\markus.baechler.BANKENBERATUNG\bbz Fuehrung\
git init
git branch -M main
git remote add origin https://github.com/markusbaechler/bbz-Fuehrung.git
```

## 3. Dateien kopieren und committen

Alle 8 Dateien aus diesem Bundle in den lokalen Projekt-Ordner kopieren:
- `index.html`
- `store.js`
- `graph.js`
- `auth.js`
- `config.js`
- `msal-browser.min.js`
- `.nojekyll`  ← wichtig, nicht vergessen (versteckte Datei)
- `README.md`
- `.gitignore`

Dann:

```bash
git add .
git commit -m "Initial deploy to GitHub Pages"
git push -u origin main
```

## 4. GitHub Pages aktivieren

GitHub.com → Repo `bbz-Fuehrung` → **Settings** → **Pages**

- Source: **Deploy from a branch**
- Branch: **main** / **/ (root)**
- Save

Nach 1-3 Minuten: Status "Your site is live at https://markusbaechler.github.io/bbz-Fuehrung/"

## 5. Testen

1. `https://markusbaechler.github.io/bbz-Fuehrung/` öffnen
2. "Mit Microsoft anmelden"
3. Alle Views durchklicken (Dashboard → FüSi → EarlyBird → Strategietage → Offsite)
4. Ein Test-Write machen (z.B. Traktandum erstellen, dann wieder soft-löschen)
5. Seite reloaden, prüfen ob Daten persistiert sind
6. In SharePoint-Excel prüfen ob der Test-Eintrag dort sichtbar ist

## 6. Bei Problemen

| Fehler | Ursache | Fix |
|---|---|---|
| `AADSTS50011: redirect URI mismatch` | Azure-URI fehlt | Schritt 1 |
| 404 für `msal-browser.min.js` | `.nojekyll` fehlt / Underscore-Problem | Datei committed? |
| Auth-Popup schliesst sich sofort | Popup-Blocker | Popups für Site erlauben |
| `Files.ReadWrite.All` missing | Scopes nicht granted | Azure-Consent prüfen |

## Updates später

Für jede weitere Änderung:

```bash
git add .
git commit -m "Beschreibung"
git push
```

GitHub Pages deployed innerhalb von 1-2 Minuten automatisch.
