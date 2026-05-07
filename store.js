// ==========================================================
// store.js - Zentrale Datenhaltung + Save-Queue
// ==========================================================
// Konvertiert zwischen Excel-Rows (Arrays) und App-Objekten,
// verwaltet Änderungen und speichert sie batched zurück.
// ==========================================================

const Store = (() => {
  let data = {
    fusi: [],
    eb: [],
    st: [],
    off: [],
    tt: [],
    ag: [],
    dropdowns: {
      persons:       ['af','bae','ml','mr','pb','pw','ss'],
      status:        ['Entscheid','In Arbeit','Parkplatz','Erledigt','via App gelöscht'],
      aktion:        ['01 - toDecide','02 - toInform'],
      kategorie:     ['Menschen@bbz','Prozesse','Kunden/Markt','für earlyBird'],
      strategietage: ['Projektportfolio','Mitarbeitende','SGF Nachwuchs','SGF Weiterbildung','SGF Zertifizierung','SGF Tailormade','B&S','Operations']
    }
  };

  let _nextId = 1;
  const ensureId = (item) => {
    if (!item._id) item._id = `i_${_nextId++}`;
    return item._id;
  };

  const listeners = new Set();

  function emit() {
    listeners.forEach(fn => fn(data));
  }

  function subscribe(fn) {
    listeners.add(fn);
    return () => listeners.delete(fn);
  }

  /* ---------- Mapping: Excel row <-> App object ---------- */
  // FüSi columns: FüSi, Thema, Aktion, erfasst durch, Input, Link, Gedanken, Output, Status
  const mapFusiRow = (row) => {
    const statusRaw = (row._values[8] || '').toString().trim();
    return {
      _index: row._index,
      _id: `fusi_${row._index}`,
      d: xlDate(row._values[0]),
      t: row._values[1],
      a: row._values[2],
      e: row._values[3],
      i: row._values[4],
      l: row._values[5],
      g: row._values[6],
      o: row._values[7],
      s: statusRaw
      // _done wird NICHT mehr hier gesetzt - ist jetzt user-spezifisch (localStorage)
      // Siehe UserDone-Helper in index.html
    };
  };
  const fusiToRow = (item) => [
    item.d || '',
    item.t || '',
    item.a || '',
    item.e || '',
    item.i || '',
    item.l || '',
    item.g || '',
    item.o || '',
    item.s || ''   // Status = einzige Wahrheit; _done nur zur Anzeige
  ];

  // EarlyBird columns: Datum, Botschaft, wer im Lead, Kommentare & Ergänzungen, Status
  const mapEbRow = (row) => ({
    _index: row._index,
    _id: `eb_${row._index}`,
    d: xlDate(row._values[0]),
    b: row._values[1],
    l: row._values[2],
    k: row._values[3],
    s: (row._values[4] || '').toString().trim()
  });
  const ebToRow = (item) => [
    item.d || '',
    item.b || '',
    item.l || '',
    item.k || '',
    item.s || ''
  ];

  // Strategietage columns: Strategietag, Thema, Kategorie, erfasst durch, Input, Link,
  //                        Gedanken, Traktandum, Startzeit, Zeitbudget, Output, Status
  const mapStRow = (row) => ({
    _index: row._index,
    _id: `st_${row._index}`,
    d: xlDate(row._values[0]),
    t: row._values[1],
    k: row._values[2],
    e: row._values[3],
    i: row._values[4],
    l: row._values[5],
    g: row._values[6],
    tra: row._values[7],
    stz: xlTime(row._values[8]),
    zb: xlDuration(row._values[9]),
    o: row._values[10],
    s: (row._values[11] || '').toString().trim()
  });
  const stToRow = (item) => [
    item.d || '',
    item.t || '',
    item.k || '',
    item.e || '',
    item.i || '',
    item.l || '',
    item.g || '',
    item.tra || '',
    item.stz || '',
    item.zb || '',
    item.o || '',
    item.s || ''
  ];

  // Offsite columns: Offsite, Thema, Lead, Schlüsselfrage/n, Ziel/e, Traktandum, Vorbereitungsauftrag, Tag, Startzeit, Zeitbudget, Bemerkungen / Links, Entscheide, Status
  const mapOffRow = (row) => ({
    _index: row._index,
    _id: `off_${row._index}`,
    d: row._values[0],
    t: row._values[1],
    l: row._values[2],
    sf: row._values[3],
    z: row._values[4],
    tra: row._values[5],
    vb: row._values[6],
    tag: row._values[7],
    st: xlTime(row._values[8]),
    zb: xlDuration(row._values[9]),
    bm: row._values[10],
    e: row._values[11],
    s: (row._values[12] || '').toString().trim()
  });
  const offToRow = (item) => [
    item.d || '',
    item.t || '',
    item.l || '',
    item.sf || '',
    item.z || '',
    item.tra || '',
    item.vb || '',
    item.tag === null || item.tag === undefined ? '' : item.tag,
    item.st || '',
    item.zb || '',
    item.bm || '',
    item.e || '',
    item.s || ''
  ];

  // Teamtage columns: Teamtag, Thema, Idee, Was wollen wir erreichen, eingereicht durch,
  //                   Zeitbudget, Link, Nächster Schritt, Verantwortlich, Status
  // Status ist nur für Soft-Delete relevant; im UI nicht angezeigt.
  const mapTtRow = (row) => ({
    _index: row._index,
    _id: `tt_${row._index}`,
    d: xlDate(row._values[0]),
    t: row._values[1],
    i: row._values[2],
    z: row._values[3],
    e: row._values[4],
    zb: row._values[5],
    l: row._values[6],
    ns: row._values[7],
    v: row._values[8],
    s: (row._values[9] || '').toString().trim()
  });
  const ttToRow = (item) => [
    item.d || '',
    item.t || '',
    item.i || '',
    item.z || '',
    item.e || '',
    item.zb || '',
    item.l || '',
    item.ns || '',
    item.v || '',
    item.s || ''
  ];

  // Agenda columns (27 total, App nutzt aber nur 11 Kern-Spalten + 2 Freitext + 2 Info):
  //   0: Dimension, 1: intern/extern, 2: Thema, 3: Beschreibung, 4: Massnahme,
  //   5: Periodizität, 6: Monat, 7: Lead alt, 8: Lead ab 2026,
  //   9-15: UB/BAE/AF/PB/PW/SS/ML Lead (Formel-Helper, werden ignoriert),
  //   16: Execution alt, 17: Execution ab 2026,
  //   18-24: Execution Helper (Formel-Helper, werden ignoriert),
  //   25: Link, 26: Bemerkung
  const mapAgRow = (row) => ({
    _index: row._index,
    _id: `ag_${row._index}`,
    dim:    normStr(row._values[0]),
    scope:  normStr(row._values[1]),
    t:      normStr(row._values[2]),  // Thema
    beschr: normStr(row._values[3]),
    mass:   normStr(row._values[4]),
    period: normStr(row._values[5]),
    mo:     normStr(row._values[6]),
    leadAlt: normStr(row._values[7]),
    l:      normStr(row._values[8]),   // Lead ab 2026 = unser "lead"
    execAlt: normStr(row._values[16]),
    ex:     normStr(row._values[17]),  // Execution ab 2026 = unser "exec"
    link:   normStr(row._values[25]),
    bem:    normStr(row._values[26])
  });

  // Schreibt nur die Kern-Spalten (11 Werte), Helper-Spalten bleiben leer.
  // Excel-Formeln in den Helper-Spalten rechnen sich von selbst neu.
  // WICHTIG: Länge muss exakt 27 sein, sonst rejected Graph API die Row.
  const agToRow = (item) => {
    const row = new Array(27).fill('');
    row[0]  = item.dim     || '';
    row[1]  = item.scope   || '';
    row[2]  = item.t       || '';
    row[3]  = item.beschr  || '';
    row[4]  = item.mass    || '';
    row[5]  = item.period  || '';
    row[6]  = item.mo      || '';
    row[7]  = item.leadAlt || '';
    row[8]  = item.l       || '';
    row[16] = item.execAlt || '';
    row[17] = item.ex      || '';
    row[25] = item.link    || '';
    row[26] = item.bem     || '';
    // Helper-Spalten (9-15, 18-24) bleiben leer - Excel füllt sie via Formel
    return row;
  };

  // Helper: null/undefined → '', sonst String-Trim
  function normStr(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function xlDate(v) {
    if (!v) return null;
    if (typeof v === 'number') {
      // Excel serial date
      const date = new Date((v - 25569) * 86400 * 1000);
      return date.toISOString().slice(0, 10);
    }
    if (typeof v === 'string') {
      // Try to parse common formats
      const dt = new Date(v);
      if (!isNaN(dt)) return dt.toISOString().slice(0, 10);
      return v;
    }
    return v;
  }

  // Excel-Zeit (Bruchteil eines Tages) → "HH:MM"
  function xlTime(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number') {
      // Fraction-of-day (0..1). Falls > 1, ist es evtl. ein Serial-Date mit Uhrzeit-Anteil
      const frac = v < 1 ? v : v - Math.floor(v);
      const totalMin = Math.round(frac * 24 * 60);
      const h = Math.floor(totalMin / 60) % 24;
      const m = totalMin % 60;
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }
    if (typeof v === 'string') {
      // Bereits formatiert (z.B. "08:00") oder als Zahl-String
      if (/^\d{1,2}:\d{2}/.test(v)) return v.slice(0, 5);
      const num = parseFloat(v);
      if (!isNaN(num) && num >= 0 && num < 2) return xlTime(num);
      return v;
    }
    return '';
  }

  // Excel-Dauer (Bruchteil) → "X Std Y min" oder "Y min"
  function xlDuration(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number') {
      const totalMin = Math.round(v * 24 * 60);
      if (totalMin === 0) return '';
      const h = Math.floor(totalMin / 60);
      const m = totalMin % 60;
      if (h === 0) return `${m} min`;
      if (m === 0) return `${h} Std`;
      return `${h}:${String(m).padStart(2, '0')} Std`;
    }
    if (typeof v === 'string') {
      const num = parseFloat(v);
      if (!isNaN(num) && v.indexOf(':') === -1 && num < 2) return xlDuration(num);
      return v;
    }
    return '';
  }

  /* ---------- Find row by content (robust gegen Index-Verschiebung) ---------- */
  // Sucht in der Excel die Zeile, die zu unserem Item gehört - per Datum + Thema + Person
  // So bleiben wir unabhängig von Position, wenn andere parallel arbeiten.
  async function findRowIndex(item) {
    const T = window.BBZ_CONFIG.tables.fusi;
    const fresh = await GraphService.getTableRows(T);
    // Match durch Datum (Spalte 0) + Thema (Spalte 1) + Person (Spalte 3)
    const targetDate = fusiDateToExcel(item.d);
    const targetTitle = item._originalTitle || item.t;  // Falls Titel geändert wurde, Originaltitel matchen
    const targetPerson = item.e;

    // Try exact match first
    let idx = fresh.findIndex(r => {
      const dateMatch = isSameDate(r._values[0], targetDate);
      const titleMatch = (r._values[1] || '') === targetTitle;
      const personMatch = (r._values[3] || '') === targetPerson;
      return dateMatch && titleMatch && personMatch;
    });

    // Fallback: nur Datum + Thema (ohne Person)
    if (idx === -1) {
      idx = fresh.findIndex(r =>
        isSameDate(r._values[0], targetDate) && (r._values[1] || '') === targetTitle
      );
    }
    return idx;
  }

  function isSameDate(a, b) {
    if (!a || !b) return false;
    // Excel gibt oft serielle Zahlen, sometimes Strings
    const da = typeof a === 'number' ? xlDate(a) : String(a).slice(0,10);
    const db = typeof b === 'number' ? xlDate(b) : String(b).slice(0,10);
    return da === db;
  }

  function fusiDateToExcel(d) {
    // Unsere interne Darstellung ist YYYY-MM-DD, Excel kann das
    return d;
  }

  /* ---------- Reload FüSi from Graph (keeps other tables) ---------- */
  async function reloadFusi() {
    try {
      const raw = await GraphService.getTableRows(window.BBZ_CONFIG.tables.fusi);
      const softDel = window.BBZ_CONFIG.features.softDeleteStatus;
      data.fusi = raw
        .map(mapFusiRow)
        .filter(x => (x.t || x.d) && x.s !== softDel);  // soft-deleted ausblenden
      emit();
    } catch (e) {
      console.error('reloadFusi failed', e);
    }
  }

  /* ---------- Load from Graph ---------- */
  async function loadFromGraph() {
    const raw = await GraphService.loadAll();
    const softDel = window.BBZ_CONFIG.features.softDeleteStatus;
    data.fusi = raw.fusi.map(mapFusiRow).filter(x => (x.t || x.d) && x.s !== softDel);
    data.eb = raw.eb.map(mapEbRow).filter(x => (x.b || x.d) && x.s !== softDel);
    data.st = raw.st.map(mapStRow).filter(x => (x.t || x.d) && x.s !== softDel);
    data.off = raw.off.map(mapOffRow).filter(x => (x.t || x.d) && x.s !== softDel);
    data.tt = raw.tt.map(mapTtRow).filter(x => (x.t || x.d) && x.s !== softDel);
    // Agenda: kein Soft-Delete-Filter nötig (keine Status-Spalte).
    // Nur Zeilen mit Thema behalten.
    data.ag = (raw.ag || []).map(mapAgRow).filter(x => x.t);
    // Dropdown-Daten im Hintergrund laden - blockiert UI nicht
    loadDropdowns().catch(e => console.warn('Dropdowns nicht geladen, Defaults verwendet:', e.message));
    emit();
    return data;
  }

  /**
   * Lädt Dropdown-Werte aus dem Daten-Sheet.
   * Zwei-Stufen-Suche:
   *   1. Name-Match (z.B. tbl_Mitarbeiter, tbl_Status)
   *   2. Header-Match bei nicht-zugeordneten Tabellen (z.B. "Tabelle1" mit Header "Mitarbeiter")
   */
  async function loadDropdowns() {
    let tables;
    try {
      tables = await GraphService.listAllTables();
    } catch (e) {
      console.warn('[dropdowns] Konnte Tabellen-Liste nicht holen:', e.message);
      return;
    }
    console.log('[dropdowns] Gefundene Tabellen:', tables.map(t => t.name));

    const mapping = [
      { key: 'persons',       nameHints: ['mitarbeiter','personen','persons'],              headerHints: ['mitarbeiter','person','kürzel','kuerzel'] },
      { key: 'status',        nameHints: ['status'],                                         headerHints: ['status'] },
      { key: 'aktion',        nameHints: ['aktion','action'],                                headerHints: ['aktion'] },
      { key: 'kategorie',     nameHints: ['kategorie','category','kategorien'],              headerHints: ['kategorie'] },
      { key: 'strategietage', nameHints: ['strategietage','strategie'],                      headerHints: ['strategietage','strategie'] }
    ];

    // Tabellen, die die Haupt-Datentabellen sind (nicht als Dropdown-Quelle nutzen)
    const mainTables = [
      'tbl_fuesioperativ', 'tbl_earlybird', 'tbl_offsite', 'tbl_teamtage', 'tbl_agenda'
    ];
    const norm = s => s.toLowerCase().replace(/[\s\-_]/g, '');

    // Stufe 1: Name-Match
    const matched = {};
    for (const cfg of mapping) {
      const hit = tables.find(t => {
        const n = norm(t.name).replace(/^tbl/, '');
        return cfg.nameHints.some(h => n === norm(h) || n.includes(norm(h)));
      });
      if (hit) matched[cfg.key] = hit.name;
    }

    // Stufe 2: Header-Match bei ungematchten Tabellen (auch 'Tabelle1' etc.)
    const remaining = mapping.filter(cfg => !matched[cfg.key]);
    if (remaining.length > 0) {
      const unassigned = tables.filter(t => {
        const nt = norm(t.name);
        if (mainTables.some(m => nt === norm(m))) return false;
        // Bereits zugeordnete Tabellen ausschliessen
        if (Object.values(matched).some(n => norm(n) === nt)) return false;
        return true;
      });
      // Header aller unzugeordneten Tabellen parallel laden
      const headerProbe = await Promise.all(unassigned.map(async (t) => {
        try {
          const cols = await GraphService.getTableColumns(t.name);
          return { name: t.name, header: (cols[0] || '').toLowerCase() };
        } catch { return { name: t.name, header: '' }; }
      }));
      for (const cfg of remaining) {
        const hit = headerProbe.find(p =>
          cfg.headerHints.some(h => p.header.includes(h))
        );
        if (hit) matched[cfg.key] = hit.name;
      }
    }

    // Werte aus gefundenen Tabellen laden (parallel)
    const loadResults = await Promise.all(
      mapping.map(async cfg => {
        const tblName = matched[cfg.key];
        if (!tblName) return { key: cfg.key, matched: false };
        try {
          const rows = await GraphService.getTableRows(tblName);
          const values = rows
            .map(r => (r._values[0] || '').toString().trim())
            .filter(v => v.length > 0);
          if (values.length > 0) {
            data.dropdowns[cfg.key] = values;
            return { key: cfg.key, matched: true, table: tblName, count: values.length };
          }
        } catch (e) {
          console.warn(`[dropdowns] Fehler bei ${tblName}:`, e.message);
        }
        return { key: cfg.key, matched: false };
      })
    );
    loadResults.forEach(r => {
      if (r.matched) console.log(`[dropdowns] ${r.key} ← "${r.table}" (${r.count} Einträge)`);
      else console.log(`[dropdowns] ${r.key} → Default verwendet`);
    });
    emit();
  }

  /* ---------- Mutations: write-through via find-by-content ---------- */
  // Jede Schreiboperation:
  // 1. Frische Tabelle holen
  // 2. Richtige Zeile per Inhalt finden
  // 3. An der aktuellen Position schreiben/löschen
  // 4. Lokale Daten nochmal refreshen

  async function updateFusi(id, changes) {
    const item = data.fusi.find(f => f._id === id);
    if (!item) return;

    const originalTitle = item.t;
    const originalDate = item.d;
    const originalPerson = item.e;

    Object.assign(item, changes);
    emit();

    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.fusi;
      const searchItem = { d: originalDate, t: originalTitle, e: originalPerson, _originalTitle: originalTitle };
      const rowIdx = await findRowIndex(searchItem);
      if (rowIdx === -1) {
        throw new Error(`Zeile nicht gefunden: "${originalTitle}"`);
      }
      await GraphService.updateRow(T, rowIdx, fusiToRow(item));
    });
  }

  async function addFusi(item) {
    ensureId(item);
    data.fusi.push(item);
    emit();
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.fusi;
      await GraphService.addRow(T, fusiToRow(item));
    });
  }

  async function deleteFusi(id) {
    const item = data.fusi.find(f => f._id === id);
    if (!item) return;
    const snapshot = { ...item };  // snapshot before local removal
    const idx = data.fusi.findIndex(f => f._id === id);
    data.fusi.splice(idx, 1);
    emit();

    // SOFT-DELETE: Statt DELETE schreiben wir den Status
    // (echter DELETE scheitert in Excel-Tabellen mit Datenschnitten mit 409)
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.fusi;
      const rowIdx = await findRowIndex(snapshot);
      if (rowIdx === -1) {
        console.warn('[deleteFusi] Zeile nicht gefunden, skip');
        return;
      }
      const softDeletedItem = { ...snapshot, s: window.BBZ_CONFIG.features.softDeleteStatus };
      await GraphService.updateRow(T, rowIdx, fusiToRow(softDeletedItem));
    });
  }

  function addEb(item) {
    ensureId(item);
    data.eb.push(item);
    emit();
    queueWrite(async () => {
      await GraphService.addRow(window.BBZ_CONFIG.tables.earlyBird, ebToRow(item));
    });
  }

  function updateEb(id, changes) {
    const item = data.eb.find(e => e._id === id);
    if (!item) return;
    const originalDate = item.d;
    const originalBotschaft = item.b;

    Object.assign(item, changes);
    emit();

    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.earlyBird;
      const fresh = await GraphService.getTableRows(T);
      const norm = (s) => (s || '').toString().trim().replace(/\s+/g, ' ');
      // 1. Versuch: exakter Match
      let rowIdx = fresh.findIndex(r =>
        isSameDate(r._values[0], originalDate) && (r._values[1] || '') === (originalBotschaft || '')
      );
      // 2. Versuch: normalisiert
      if (rowIdx === -1) {
        rowIdx = fresh.findIndex(r =>
          isSameDate(r._values[0], originalDate) && norm(r._values[1]) === norm(originalBotschaft)
        );
      }
      // 3. Versuch: nur über _index (Fallback für Daten ohne eindeutigen Text)
      if (rowIdx === -1 && item._index !== undefined) {
        rowIdx = fresh.findIndex(r => r._index === item._index);
      }
      if (rowIdx === -1) {
        console.warn(`[updateEb] Zeile nicht gefunden für "${originalBotschaft}" (${originalDate})`);
        return;  // silent fail statt throw
      }
      await GraphService.updateRow(T, rowIdx, ebToRow(item));
    });
  }

  function deleteEb(id) {
    const item = data.eb.find(e => e._id === id);
    if (!item) return;
    const snapshot = { ...item };
    const idx = data.eb.findIndex(e => e._id === id);
    data.eb.splice(idx, 1);
    emit();

    // SOFT-DELETE: Status = "via App gelöscht" setzen (nicht echtes DELETE wegen Slicer)
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.earlyBird;
      const fresh = await GraphService.getTableRows(T);
      const norm = (s) => (s || '').toString().trim().replace(/\s+/g, ' ');
      let rowIdx = fresh.findIndex(r =>
        isSameDate(r._values[0], snapshot.d) && (r._values[1] || '') === (snapshot.b || '')
      );
      if (rowIdx === -1) {
        rowIdx = fresh.findIndex(r =>
          isSameDate(r._values[0], snapshot.d) && norm(r._values[1]) === norm(snapshot.b)
        );
      }
      if (rowIdx === -1) {
        console.warn(`[deleteEb] Zeile nicht gefunden für "${snapshot.b}"`);
        return;
      }
      const softDeletedItem = { ...snapshot, s: window.BBZ_CONFIG.features.softDeleteStatus };
      await GraphService.updateRow(T, rowIdx, ebToRow(softDeletedItem));
    });
  }

  /* ---------- Strategietage CRUD ---------- */

  async function findStRowIndex(item) {
    const T = window.BBZ_CONFIG.tables.strategietage;
    const fresh = await GraphService.getTableRows(T);
    const norm = s => (s || '').toString().trim().replace(/\s+/g, ' ');
    const targetDate = item.d;
    const targetTitle = item._originalTitle || item.t;

    // Exact
    let idx = fresh.findIndex(r =>
      isSameDate(r._values[0], targetDate) && (r._values[1] || '') === targetTitle
    );
    // Normalisiert
    if (idx === -1) {
      idx = fresh.findIndex(r =>
        isSameDate(r._values[0], targetDate) && norm(r._values[1]) === norm(targetTitle)
      );
    }
    // Fallback per _index
    if (idx === -1 && item._index !== undefined) {
      idx = fresh.findIndex(r => r._index === item._index);
    }
    return idx;
  }

  function addSt(item) {
    ensureId(item);
    data.st.push(item);
    emit();
    queueWrite(async () => {
      await GraphService.addRow(window.BBZ_CONFIG.tables.strategietage, stToRow(item));
    });
  }

  function updateSt(id, changes) {
    const item = data.st.find(s => s._id === id);
    if (!item) return;
    const originalTitle = item.t;
    Object.assign(item, changes);
    emit();

    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.strategietage;
      const rowIdx = await findStRowIndex({ ...item, _originalTitle: originalTitle });
      if (rowIdx === -1) {
        console.warn(`[updateSt] Zeile nicht gefunden für "${originalTitle}"`);
        return;
      }
      await GraphService.updateRow(T, rowIdx, stToRow(item));
    });
  }

  function deleteSt(id) {
    const item = data.st.find(s => s._id === id);
    if (!item) return;
    const snapshot = { ...item };
    const idx = data.st.findIndex(s => s._id === id);
    data.st.splice(idx, 1);
    emit();
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.strategietage;
      const rowIdx = await findStRowIndex(snapshot);
      if (rowIdx === -1) {
        console.warn(`[deleteSt] Zeile nicht gefunden für "${snapshot.t}"`);
        return;
      }
      const softDeletedItem = { ...snapshot, s: window.BBZ_CONFIG.features.softDeleteStatus };
      await GraphService.updateRow(T, rowIdx, stToRow(softDeletedItem));
    });
  }

  /* ---------- Offsite CRUD ---------- */

  async function findOffRowIndex(item) {
    const T = window.BBZ_CONFIG.tables.offsite;
    const fresh = await GraphService.getTableRows(T);
    const norm = s => (s || '').toString().trim().replace(/\s+/g, ' ');
    const targetOffsite = item.d;
    const targetTitle = item._originalTitle || item.t;

    // Match über Offsite-Code + Thema (exakt)
    let idx = fresh.findIndex(r =>
      (r._values[0] || '') === targetOffsite && (r._values[1] || '') === targetTitle
    );
    if (idx === -1) {
      idx = fresh.findIndex(r =>
        (r._values[0] || '') === targetOffsite && norm(r._values[1]) === norm(targetTitle)
      );
    }
    if (idx === -1 && item._index !== undefined) {
      idx = fresh.findIndex(r => r._index === item._index);
    }
    return idx;
  }

  function addOff(item) {
    ensureId(item);
    data.off.push(item);
    emit();
    queueWrite(async () => {
      await GraphService.addRow(window.BBZ_CONFIG.tables.offsite, offToRow(item));
    });
  }

  function updateOff(id, changes) {
    const item = data.off.find(o => o._id === id);
    if (!item) return;
    const originalTitle = item.t;
    Object.assign(item, changes);
    emit();
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.offsite;
      const rowIdx = await findOffRowIndex({ ...item, _originalTitle: originalTitle });
      if (rowIdx === -1) {
        console.warn(`[updateOff] Zeile nicht gefunden für "${originalTitle}"`);
        return;
      }
      await GraphService.updateRow(T, rowIdx, offToRow(item));
    });
  }

  function deleteOff(id) {
    const item = data.off.find(o => o._id === id);
    if (!item) return;
    const snapshot = { ...item };
    const idx = data.off.findIndex(o => o._id === id);
    data.off.splice(idx, 1);
    emit();
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.offsite;
      const rowIdx = await findOffRowIndex(snapshot);
      if (rowIdx === -1) {
        console.warn(`[deleteOff] Zeile nicht gefunden für "${snapshot.t}"`);
        return;
      }
      const softDeletedItem = { ...snapshot, s: window.BBZ_CONFIG.features.softDeleteStatus };
      await GraphService.updateRow(T, rowIdx, offToRow(softDeletedItem));
    });
  }

  /* ---------- Teamtage CRUD ---------- */
  // Match-Logik: Datum + Thema. Nicht ganz eindeutig (Thema wie "Fitnesscheck"
  // kommt öfter vor), aber immer am selben Datum nur 1x.
  async function findTtRowIndex(item) {
    const T = window.BBZ_CONFIG.tables.teamtage;
    const fresh = await GraphService.getTableRows(T);
    const norm = s => (s || '').toString().trim().replace(/\s+/g, ' ');
    const targetDate = item.d;
    const targetTitle = item._originalTitle || item.t;

    // 1. Exakter Match: Datum + Thema
    let idx = fresh.findIndex(r =>
      isSameDate(r._values[0], targetDate) && (r._values[1] || '') === targetTitle
    );
    if (idx >= 0) return fresh[idx]._index;

    // 2. Normalisierter Match
    idx = fresh.findIndex(r =>
      isSameDate(r._values[0], targetDate) && norm(r._values[1]) === norm(targetTitle)
    );
    if (idx >= 0) return fresh[idx]._index;

    // 3. Fallback über _index
    if (item._index !== undefined) {
      const hit = fresh.find(r => r._index === item._index);
      if (hit) return hit._index;
    }
    return -1;
  }

  function addTt(item) {
    ensureId(item);
    data.tt.push(item);
    emit();
    queueWrite(async () => {
      await GraphService.addRow(window.BBZ_CONFIG.tables.teamtage, ttToRow(item));
    });
  }

  function updateTt(id, changes) {
    const item = data.tt.find(x => x._id === id);
    if (!item) return;
    const originalTitle = item.t;
    Object.assign(item, changes);
    emit();
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.teamtage;
      const rowIdx = await findTtRowIndex({ ...item, _originalTitle: originalTitle });
      if (rowIdx === -1) {
        console.warn(`[updateTt] Zeile nicht gefunden für "${originalTitle}"`);
        return;
      }
      await GraphService.updateRow(T, rowIdx, ttToRow(item));
    });
  }

  // Soft-Delete wie bei FüSi/Strategietage/Offsite
  function deleteTt(id) {
    const item = data.tt.find(x => x._id === id);
    if (!item) return;
    const snapshot = { ...item };
    const idx = data.tt.findIndex(x => x._id === id);
    data.tt.splice(idx, 1);
    emit();
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.teamtage;
      const rowIdx = await findTtRowIndex(snapshot);
      if (rowIdx === -1) {
        console.warn(`[deleteTt] Zeile nicht gefunden für "${snapshot.t}"`);
        return;
      }
      const softDeleted = { ...snapshot, s: window.BBZ_CONFIG.features.softDeleteStatus };
      await GraphService.updateRow(T, rowIdx, ttToRow(softDeleted));
    });
  }

  /* ---------- Agenda CRUD ---------- */
  // Match-Logik: Thema ist eindeutig in der Agenda (keine Duplikate nötig).
  // Zusätzlich: falls Thema doch doppelt, zweite Runde mit Lead ab 2026 + Periodizität.
  async function findAgRowIndex(item) {
    const T = window.BBZ_CONFIG.tables.agenda;
    const fresh = await GraphService.getTableRows(T);
    const norm = s => (s || '').toString().trim().replace(/\s+/g, ' ');
    const targetTitle = item._originalTitle || item.t;

    // 1. Exakter Match nach Thema
    let hits = fresh.map((r, i) => ({ r, i })).filter(x => (x.r._values[2] || '') === targetTitle);
    if (hits.length === 1) return hits[0].r._index;

    // 2. Normalisierter Match
    if (hits.length === 0) {
      hits = fresh.map((r, i) => ({ r, i })).filter(x => norm(x.r._values[2]) === norm(targetTitle));
    }
    if (hits.length === 1) return hits[0].r._index;

    // 3. Mehrere Treffer → zusätzlich nach Lead + Periodizität unterscheiden
    if (hits.length > 1) {
      const targetLead = (item._originalLead !== undefined ? item._originalLead : item.l) || '';
      const targetPeriod = item.period || '';
      const narrowed = hits.filter(x =>
        (x.r._values[8] || '') === targetLead &&
        (x.r._values[5] || '') === targetPeriod
      );
      if (narrowed.length === 1) return narrowed[0].r._index;
    }

    // 4. Fallback: _index
    if (item._index !== undefined) {
      const byIdx = fresh.find(r => r._index === item._index);
      if (byIdx) return byIdx._index;
    }
    return -1;
  }

  function addAg(item) {
    ensureId(item);
    data.ag.push(item);
    emit();
    queueWrite(async () => {
      await GraphService.addRow(window.BBZ_CONFIG.tables.agenda, agToRow(item));
    });
  }

  function updateAg(id, changes) {
    const item = data.ag.find(a => a._id === id);
    if (!item) return;
    const originalTitle = item.t;
    const originalLead = item.l;
    Object.assign(item, changes);
    emit();
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.agenda;
      const rowIdx = await findAgRowIndex({ ...item, _originalTitle: originalTitle, _originalLead: originalLead });
      if (rowIdx === -1) {
        console.warn(`[updateAg] Zeile nicht gefunden für "${originalTitle}"`);
        return;
      }
      await GraphService.updateRow(T, rowIdx, agToRow(item));
    });
  }

  // Hard-Delete für Agenda: die Tabelle hat keine Status-Spalte für Soft-Delete,
  // und die Agenda hat keine Slicer die 409-Konflikte verursachen.
  function deleteAg(id) {
    const item = data.ag.find(a => a._id === id);
    if (!item) return;
    const snapshot = { ...item };
    const idx = data.ag.findIndex(a => a._id === id);
    data.ag.splice(idx, 1);
    emit();
    queueWrite(async () => {
      const T = window.BBZ_CONFIG.tables.agenda;
      const rowIdx = await findAgRowIndex(snapshot);
      if (rowIdx === -1) {
        console.warn(`[deleteAg] Zeile nicht gefunden für "${snapshot.t}"`);
        return;
      }
      await GraphService.deleteRow(T, rowIdx);
    });
  }

  /* ---------- Write Queue (ordered, one at a time) ---------- */
  const writeQueue = [];
  let writing = false;

  function queueWrite(writeFn) {
    writeQueue.push(writeFn);
    emitSaveState('pending');
    processQueue();
  }

  async function processQueue() {
    if (writing || writeQueue.length === 0) return;
    writing = true;
    emitSaveState('saving');

    while (writeQueue.length > 0) {
      const fn = writeQueue.shift();
      try {
        await fn();
      } catch (e) {
        console.error('Write failed', e);
        notifySaveError(e);
        // Bei schwerem Fehler: Daten neu laden, um lokal/remote zu synchronisieren
        if (e.status === 409 || e.message.includes('nicht gefunden')) {
          await reloadFusi();
        }
      }
    }

    emitSaveState('saved');
    writing = false;

    // Nach allen Schreibvorgängen: Daten frisch laden für Konsistenz
    setTimeout(() => reloadFusi(), 500);
  }

  const saveListeners = new Set();
  const onSaveState = (fn) => { saveListeners.add(fn); return () => saveListeners.delete(fn); };
  function emitSaveState(s) { saveListeners.forEach(fn => fn(s)); }

  function notifySaveError(e) {
    console.error('Save error:', e);
    const msg = e.status === 401 ? 'Sitzung abgelaufen - bitte neu anmelden'
              : e.status === 423 ? 'Datei ist gesperrt (von anderem Nutzer)'
              : e.status === 404 ? 'Tabelle/Zeile nicht gefunden'
              : (e.message || 'Unbekannter Fehler');
    if (window.showToast) window.showToast('⚠ ' + msg);
  }

  return {
    data: () => data,
    subscribe,
    onSaveState,
    loadFromGraph,
    reloadFusi,
    updateFusi,
    addFusi,
    deleteFusi,
    addEb,
    updateEb,
    deleteEb,
    addSt,
    updateSt,
    deleteSt,
    addOff,
    updateOff,
    deleteOff,
    addTt,
    updateTt,
    deleteTt,
    addAg,
    updateAg,
    deleteAg,
    flushQueue: () => processQueue(),
    hasPending: () => writeQueue.length > 0 || writing
  };
})();
