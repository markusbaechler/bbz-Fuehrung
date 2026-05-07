// ==========================================================
// bbz Führung - Konfiguration
// ==========================================================
// Diese Werte sind spezifisch für eure Azure-Registration
// und die Excel-Datei auf SharePoint.
// ==========================================================

window.BBZ_CONFIG = {
  // --- Azure AD / MSAL ---
  auth: {
    clientId: '031720ad-b432-4e8b-8438-4b84322860ff',
    tenantId: '3643e7ab-d166-4e27-bd5f-c5bbfcd282d7',
    // Redirect URI wird automatisch aus window.location.origin abgeleitet
    // Trage in Azure beide URLs ein:
    //   - https://<dein-github-user>.github.io/<repo>/
    //   - http://localhost:3000  (für lokales Testen)
  },

  // --- Graph API / SharePoint File ---
  sharepoint: {
    hostname: 'bbzsg.sharepoint.com',
    sitePath: '/sites/bbz-Fuehrung',
    fileId: 'F2767B45-9F44-4B34-9534-50DA822E5CED',  // aus sourcedoc-URL-Parameter
    fileName: 'Traktandenliste FüSi.xlsx'
  },

  // --- Tabellen-Namen in der Excel (nach Umbenennung) ---
  tables: {
    fusi:         'tbl_FuesiOperativ',
    earlyBird:    'tbl_EarlyBird',
    strategietage:'tbl_Strategietage',
    offsite:      'tbl_Offsite',
    mgmtOffsite:  'tbl_MgmtOffsite',
    teamtage:     'tbl_Teamtage',
    agenda:       'tbl_Agenda'
  },

  // --- Spalten-Mapping pro Tabelle ---
  // Reihenfolge & Namen müssen exakt mit den Excel-Headern übereinstimmen
  columns: {
    fusi: ['FüSi', 'Thema', 'Aktion', 'erfasst durch', 'Input - (durch Themengeber/in auszufüllen)', 'Link (Optional)', 'Gedanken im Vorfeld oder bei Abwesenheit', 'Output', 'Status'],
    earlyBird: ['Datum', 'Botschaft', 'wer im Lead', 'Kommentare & Ergänzungen'],
    strategietage: ['Strategietag', 'Thema', 'Kategorie', 'erfasst durch', 'Input ', 'Link (Optional)', 'Gedanken im Vorfeld oder bei Abwesenheit', 'Traktandum', 'Startzeit', 'Zeitbudget', 'Output - Entscheid', 'Status'],
    offsite: ['Offsite', 'Thema', 'Lead', 'Schlüsselfrage/n', 'Ziel/e', 'Traktandum', 'Vorbereitungsauftrag', 'Tag', 'Startzeit', 'Zeitbudget', 'Bemerkungen / Links ', 'Entscheide', 'Status'],
    teamtage: ['Teamtag', 'Thema', 'Idee / Beschreibung', 'Was wollen wir damit erreichen?', 'eingereicht durch', 'Zeitbudget am Teamtag', 'Link ', 'Nächster Schritt - Auftrag', 'Verantwortlich', 'Status'],
    // Agenda: 27 Spalten total. Nur Index 0-8, 16-17, 25-26 werden von der App geschrieben.
    // Spalten 9-15 (UB/BAE/AF/PB/PW/SS/ML Lead neu) und 18-24 (Execution neu) sind Excel-Formel-Helper
    // für Slicer und werden von der App leer geschrieben - Excel füllt sie selbst per Formel.
    agenda: [
      'Dimension', 'intern / extern', 'Thema', 'Beschreibung | Was?', 'Massnahme | Wie?',
      'Periodizität', 'Monat', 'Lead alt', 'Lead ab 2026',
      'UB; Lead neu', 'BAE; Lead neu ', 'AF; Lead neu ', 'PB; Lead neu ',
      'PW; Lead neu ', 'SS; Lead neu ', 'ML; Lead neu ',
      'Execution alt', 'Execution ab 2026',
      'UB; Execution neu', ' BAE; Execution neu', 'AF; Execution neu ',
      'PB; Execution neu', ' PW; Execution neu', 'SS; Execution neu',
      ' ML; Execution neu',
      'Link', 'Bemerkung'
    ]
  },

  // --- Feature-Flags ---
  features: {
    useMockData: false,      // true = Mock-Modus (kein Graph), false = echte SharePoint-Daten
    autoSaveDelay: 500,      // ms Debounce nach letzter Eingabe bis Save
    workbookSessionDuration: 300,  // Sekunden, nach denen Session neu geöffnet wird
    softDeleteStatus: 'via App gelöscht'  // Wert in der Status-Spalte, wenn via App gelöscht wurde
  }
};
