// ==========================================================
// graph.js - Microsoft Graph API Adapter (Excel Workbook)
// ==========================================================

const GraphService = (() => {
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  let siteId = null;
  let driveId = null;
  let fileId = null;
  let sessionId = null;
  let sessionTimer = null;

  /* ---------- HTTP helper with retry & session recovery ---------- */
  async function graph(path, opts = {}, retryOn409 = true, retryOnSession = true) {
    const token = await AuthService.getToken();
    const headers = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(opts.headers || {})
    };
    if (sessionId) headers['workbook-session-id'] = sessionId;

    const url = path.startsWith('http') ? path : `${GRAPH}${path}`;

    let resp;
    try {
      resp = await fetch(url, { ...opts, headers });
    } catch (networkErr) {
      // Netzwerk-Fehler (offline, ERR_NETWORK_IO_SUSPENDED, etc.)
      // Warten und einmal neu versuchen
      if (retryOnSession) {
        await new Promise(r => setTimeout(r, 1500));
        return graph(path, opts, retryOn409, false);
      }
      const err = new Error('Netzwerk nicht erreichbar - bitte Verbindung prüfen');
      err.status = 0;
      throw err;
    }

    if (!resp.ok) {
      const errText = await resp.text();
      let errJson;
      try { errJson = JSON.parse(errText); } catch { errJson = { error: { message: errText } }; }

      // 409 Conflict: wait 500ms and retry once
      if (resp.status === 409 && retryOn409) {
        await new Promise(r => setTimeout(r, 500));
        return graph(path, opts, false, retryOnSession);
      }

      // Session invalid (400/404 with specific message): reopen session and retry once
      const msg = errJson?.error?.message || '';
      const sessionInvalid = (resp.status === 400 || resp.status === 404) &&
        (msg.includes('session') || msg.includes('Zielsitzung') || msg.includes('workbook cannot be opened'));
      if (sessionInvalid && retryOnSession && !path.includes('createSession')) {
        console.warn('[graph] Session ungültig, öffne neu und versuche erneut');
        sessionId = null;
        try {
          await openSession();
        } catch (e) {
          console.error('Session-Reopen fehlgeschlagen', e);
          throw e;
        }
        return graph(path, opts, retryOn409, false);  // retry once
      }

      const err = new Error(msg || `Graph ${resp.status}`);
      err.status = resp.status;
      err.details = errJson;
      throw err;
    }

    if (resp.status === 204) return null;
    return resp.json();
  }

  /* ---------- Resolve Site + Drive + File ---------- */
  async function resolveFile() {
    const cfg = window.BBZ_CONFIG.sharepoint;

    const siteResp = await graph(`/sites/${cfg.hostname}:${cfg.sitePath}`);
    siteId = siteResp.id;

    const driveResp = await graph(`/sites/${siteId}/drive`);
    driveId = driveResp.id;

    try {
      const search = await graph(`/sites/${siteId}/drive/root/search(q='${encodeURIComponent(cfg.fileName)}')`);
      const hit = search.value.find(item =>
        item.name === cfg.fileName && !item.folder
      );
      if (hit) {
        fileId = hit.id;
      } else {
        throw new Error('Datei im Root nicht gefunden');
      }
    } catch (e) {
      console.warn('Search failed, trying UUID lookup', e);
      const uuidResp = await graph(`/sites/${siteId}/drive/items/${cfg.fileId}`);
      fileId = uuidResp.id;
    }

    return { siteId, driveId, fileId };
  }

  /* ---------- Workbook Session ---------- */
  async function openSession() {
    const resp = await graph(
      `/sites/${siteId}/drive/items/${fileId}/workbook/createSession`,
      {
        method: 'POST',
        body: JSON.stringify({ persistChanges: true })
      }
    );
    sessionId = resp.id;

    const duration = (window.BBZ_CONFIG.features.workbookSessionDuration || 300) * 1000;
    clearInterval(sessionTimer);
    sessionTimer = setInterval(() => refreshSession(), duration * 0.8);

    return sessionId;
  }

  async function refreshSession() {
    if (!sessionId) return;
    try {
      await graph(
        `/sites/${siteId}/drive/items/${fileId}/workbook/refreshSession`,
        { method: 'POST' }
      );
    } catch (e) {
      console.warn('Session refresh failed, reopening', e);
      sessionId = null;
      await openSession();
    }
  }

  async function closeSession() {
    clearInterval(sessionTimer);
    if (!sessionId) return;
    try {
      await graph(
        `/sites/${siteId}/drive/items/${fileId}/workbook/closeSession`,
        { method: 'POST' }
      );
    } catch (e) { /* ignore */ }
    sessionId = null;
  }

  /* ---------- Check if Excel is currently open by others ---------- */
  async function isExcelOpenElsewhere() {
    try {
      const item = await graph(`/sites/${siteId}/drive/items/${fileId}`);
      const me = AuthService.getAccount();
      const lastModifier = item.lastModifiedBy?.user?.email || item.lastModifiedBy?.user?.displayName;
      const now = Date.now();
      const lastMod = new Date(item.lastModifiedDateTime).getTime();
      const recentlyModifiedByOther = (now - lastMod) < 60000 &&
        lastModifier && me?.username && !lastModifier.includes(me.username);
      return {
        isOpen: recentlyModifiedByOther,
        lastModifier: lastModifier,
        lastModified: item.lastModifiedDateTime
      };
    } catch (e) {
      return { isOpen: false };
    }
  }

  /* ---------- Table Operations ---------- */
  async function getTableRows(tableName) {
    const resp = await graph(
      `/sites/${siteId}/drive/items/${fileId}/workbook/tables/${encodeURIComponent(tableName)}/rows?$top=5000`
    );
    return resp.value.map((row, idx) => ({
      _index: idx,          // position in table (excluding header)
      _values: row.values[0]
    }));
  }

  async function getTableColumns(tableName) {
    const resp = await graph(
      `/sites/${siteId}/drive/items/${fileId}/workbook/tables/${encodeURIComponent(tableName)}/columns`
    );
    return resp.value.map(c => c.name);
  }

  async function addRow(tableName, values) {
    return graph(
      `/sites/${siteId}/drive/items/${fileId}/workbook/tables/${encodeURIComponent(tableName)}/rows/add`,
      {
        method: 'POST',
        body: JSON.stringify({ values: [values] })
      }
    );
  }

  async function updateRow(tableName, rowIndex, values) {
    return graph(
      `/sites/${siteId}/drive/items/${fileId}/workbook/tables/${encodeURIComponent(tableName)}/rows/itemAt(index=${rowIndex})`,
      {
        method: 'PATCH',
        body: JSON.stringify({ values: [values] })
      }
    );
  }

  async function deleteRow(tableName, rowIndex) {
    return graph(
      `/sites/${siteId}/drive/items/${fileId}/workbook/tables/${encodeURIComponent(tableName)}/rows/itemAt(index=${rowIndex})`,
      { method: 'DELETE' }
    );
  }

  /* ---------- Bulk load all sheets in parallel ---------- */
  async function loadAll() {
    const T = window.BBZ_CONFIG.tables;
    const [fusi, eb, st, off, tt, ag] = await Promise.all([
      getTableRows(T.fusi).catch(e => { console.error('FüSi load:', e); return []; }),
      getTableRows(T.earlyBird).catch(e => { console.error('EB load:', e); return []; }),
      getTableRows(T.strategietage).catch(e => { console.error('ST load:', e); return []; }),
      getTableRows(T.offsite).catch(e => { console.error('Offsite load:', e); return []; }),
      getTableRows(T.teamtage).catch(e => { console.error('TT load:', e); return []; }),
      getTableRows(T.agenda).catch(e => { console.error('Agenda load:', e); return []; })
    ]);
    return { fusi, eb, st, off, tt, ag };
  }

  async function listAllTables() {
    const resp = await graph(
      `/sites/${siteId}/drive/items/${fileId}/workbook/tables`
    );
    return resp.value.map(t => ({ name: t.name, id: t.id }));
  }

  return {
    resolveFile,
    openSession,
    refreshSession,
    closeSession,
    isExcelOpenElsewhere,
    getTableRows,
    getTableColumns,
    listAllTables,
    addRow,
    updateRow,
    deleteRow,
    loadAll,
    getIds: () => ({ siteId, driveId, fileId, sessionId })
  };
})();
