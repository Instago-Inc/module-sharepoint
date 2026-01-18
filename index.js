// modules_cache/sharepoint/1.0.0/index.js
// Upload files to SharePoint/OneDrive via Microsoft Graph

(function () {
  const b64 = require('b64@latest');
  const j = require('json@latest');
  const pathx = require('path@latest');
  const fmt = require('fmt@latest');
  const graph = require('graph@latest');
  const log = require('log@latest').create('sharepoint');
  const GRAPH = 'https://graph.microsoft.com/v1.0';

  // --- utils ---
  function joinPath(a, b) {
    a = (a || '').replace(/^\/+|\/+$/g, '');
    b = (b || '').replace(/^\/+|\/+$/g, '');
    return a && b ? (a + '/' + b) : (a || b);
  }
  function firstCompanyToken(s) {
    s = ('' + (s || '')).trim();
    if (!s) return '';
    // Prefer the first whitespace-delimited token
    const wsIdx = s.search(/\s/);
    if (wsIdx > 0) return s.slice(0, wsIdx);
    // If no whitespace (e.g., MicrosoftIrelandOperationsLimited), take leading letters sequence
    const m = /^([A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]{2,})/.exec(s);
    return m ? m[1] : s;
  }
  function yyyymmdd(d) {
    try {
      const dt = new Date(d);
      if (isNaN(dt.getTime())) throw 0;
      const y = dt.getUTCFullYear();
      const m = String(dt.getUTCMonth() + 1).padStart(2, '0');
      const da = String(dt.getUTCDate()).padStart(2, '0');
      return '' + y + m + da;
    } catch {
      const dt = new Date();
      const y = dt.getUTCFullYear();
      const m = String(dt.getUTCMonth() + 1).padStart(2, '0');
      const da = String(dt.getUTCDate()).padStart(2, '0');
      return '' + y + m + da;
    }
  }
  function fmtAmount(s) {
    const t = ('' + (s || '')).replace(',', '.').replace(/[^0-9.]/g, '');
    if (!t) return '0.00';
    const n = Number(t);
    return isFinite(n) ? (Math.round(n * 100) / 100).toFixed(2) : t;
  }

  // Read file from storage if path provided; otherwise use provided base64
  async function readBase64FromStorage(path, opts) {
    const storage = sys.storage.get('sharepoint', opts);
    const res = await storage.read({ path }).catch(() => null);
    if (!res || !res.dataBase64) throw new Error('sharepoint: cannot read ' + path);
    return res.dataBase64;
  }

  // Upload small files (<4MB recommended) using simple upload
  // opts: { siteId, driveId, drivePath, filename, path, dataBase64, accessToken, contentType }
  async function uploadSmall(opts) {
    const siteId = opts.siteId;
    const driveId = opts.driveId;
    const drivePath = (opts.drivePath || '');
    const filename = opts.filename;
    const contentType = opts.contentType || 'application/octet-stream';
    const token = opts.accessToken;
    if (!siteId || !driveId || !filename || !token) throw new Error('sharepoint.uploadSmall: missing siteId/driveId/filename/accessToken');

    const rel = pathx.joinURL(drivePath, filename);
    const url = `${GRAPH}/sites/${encodeURIComponent(siteId)}/drives/${encodeURIComponent(driveId)}/root:/${encodeURIComponent(rel)}:/content`;

    const bodyBase64 = opts.dataBase64 || (opts.path ? await readBase64FromStorage(opts.path, opts) : '');
    if (!bodyBase64) throw new Error('sharepoint.uploadSmall: missing data (path or dataBase64)');

    const res = await sys.http.fetch({
      url,
      method: 'PUT',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': contentType
      },
      bodyBase64
    });
    const txt = (res && res.text) || '';
    return j.parseSafe(txt, { raw: txt });
  }

  // Try to load tokens persisted by msauth poll workflow
  async function loadStoredTokens(opts) {
    const storage = sys.storage.get('sharepoint', opts);
    const candidates = ['oauth/ms_tokens.json', 'samples/ms_tokens.json'];
    for (const p of candidates) {
      try {
        const r = await storage.read({ path: p });
        const s = b64.decodeAscii(r && r.dataBase64 || '');
        const obj = j.parseSafe(s, null);
        if (obj && (obj.refresh_token || obj.access_token)) return obj;
      } catch {}
    }
    return null;
  }

  // Ensure we have an access token. If not provided, try refresh or client credentials, or auto-read from storage.
  async function ensureAccessToken(auth) {
    if (auth && auth.accessToken) return auth.accessToken;
    // Try stored tokens if no overrides
    if (!auth || (!auth.refreshToken && !auth.clientId && !auth.clientSecret)) {
      const stored = await loadStoredTokens(auth);
      if (stored && stored.access_token) return stored.access_token;
      if (stored && stored.refresh_token) {
        const over = Object.assign({ refreshToken: stored.refresh_token }, auth || {});
        const tok = await graph.ensureAccessToken(over);
        if (tok) return tok;
      }
    }
    return await graph.ensureAccessToken(auth || {});
  }

  function pickAiForAttachment(aiResults, att) {
    const uid = att && att.uid;
    const fn = att && att.filename;
    if (!Array.isArray(aiResults)) return null;
    return aiResults.find(x => x && (x.uid === uid || x.filename === fn)) || null;
  }

  function buildBaseNameFromAi(data, fallbackBase) {
    if (!data || typeof data !== 'object' || !Object.keys(data).length) return fallbackBase;
    const dateStr = fmt.yyyymmdd(data.invoice_date || '');
    const issuerObj = (data.issuer && typeof data.issuer === 'object') ? data.issuer : null;
    const buyerObj = (data.buyer && typeof data.buyer === 'object') ? data.buyer : null;
    const issuerRaw =
      (issuerObj && (issuerObj.name || issuerObj.company || issuerObj.fullName)) ||
      (buyerObj && (buyerObj.name || buyerObj.company || buyerObj.fullName)) ||
      data.issuer ||
      data.buyer ||
      'UNKNOWN';
    const issuerShort = firstCompanyToken(issuerRaw);
    const amount = fmt.fmtAmount(data.total_amount || '0');
    return `${dateStr}_${issuerShort}_${amount}`;
  }

  async function uploadPdfAndMaybeJson(att, aiMatch, cfg, token) {
    const name = (att && att.filename) || '';
    const isPdf = /\.pdf$/i.test(name) || /pdf/i.test(att && att.contentType || '');
    if (!isPdf) return { skipped: true };
    const baseFallback = name.replace(/\.[^.]+$/, '') || ('doc_' + Date.now());
    const data = aiMatch && aiMatch.data;
    const base = buildBaseNameFromAi(data, baseFallback);
    const out = { base, pdf: null, json: null };
    const tryUploadWithSuffix = async (ext, uploader) => {
      for (let i = 0; i < 10; i++) {
        const suffix = i === 0 ? '' : '_' + i;
        const fname = base + suffix + ext;
        try {
          const res = await uploader(fname);
          out.base = base + suffix;
          return res;
        } catch (e) {
          const msg = (e && (e.message || e)) || '';
          if (typeof msg === 'string' && /exist/i.test(msg)) {
            continue; // try next suffix
          }
          throw e;
        }
      }
      throw new Error('sharepoint: too many name collisions for ' + base + ext);
    };

    try {
      out.pdf = await tryUploadWithSuffix('.pdf', (fname) =>
        uploadSmall({ siteId: cfg.siteId, driveId: cfg.driveId, drivePath: cfg.drivePath, filename: fname, path: att.path, accessToken: token, contentType: (att.contentType || 'application/pdf'), workflow: cfg.workflow })
      );
    } catch (e) {
      console.error('SP upload PDF failed:', (e && (e.message || e)) || 'unknown');
    }
    if (data && typeof data === 'object' && Object.keys(data).length) {
      const jsonStr = JSON.stringify(data);
      const jsonB64 = b64.encodeAscii(jsonStr);
      const jsonRel = `mail/${att.uid}/${out.base}.json`;
      try {
        const storage = sys.storage.get('sharepoint', cfg);
        await storage.save({ path: jsonRel, dataBase64: jsonB64 });
      } catch {}
      try {
        // Place JSON into a subdirectory (default: 'json') under the configured drivePath
        const jsonSub = (cfg && typeof cfg.jsonSubdir === 'string') ? cfg.jsonSubdir : 'json';
        const drivePathJson = pathx.joinURL(cfg.drivePath || '', jsonSub);
        out.json = await tryUploadWithSuffix('.json', (fname) =>
          uploadSmall({ siteId: cfg.siteId, driveId: cfg.driveId, drivePath: drivePathJson, filename: fname, path: jsonRel, accessToken: token, contentType: 'application/json', workflow: cfg.workflow })
        );
      } catch (e) {
        console.error('SP upload JSON failed:', out.base + '.json', (e && (e.message || e)) || 'unknown');
      }
    }
    return out;
  }

  // High-level: save invoices from attachments + AI results
  // opts: { attachments: [], aiResults: [], sp: { siteId, driveId, drivePath }, auth: { accessToken?, refreshToken?, clientId, tenant?, scope? } }
  async function saveInvoices(opts) {
    const atts = Array.isArray(opts && opts.attachments) ? opts.attachments : [];
    const ai = Array.isArray(opts && opts.aiResults) ? opts.aiResults : [];
    const sp = (opts && opts.sp) || {};
    const token = await ensureAccessToken(opts && opts.auth || {});
    if (!token) throw new Error('sharepoint.saveInvoices: missing access token (provide auth.accessToken or refreshToken+clientId)');
    if (!sp.siteId || !sp.driveId) throw new Error('sharepoint.saveInvoices: missing sp.siteId or sp.driveId');
    const cfg = { siteId: sp.siteId, driveId: sp.driveId, drivePath: sp.drivePath || 'Invoices', jsonSubdir: (typeof sp.jsonSubdir === 'string' ? sp.jsonSubdir : 'json'), workflow: opts && opts.workflow };
    const results = [];
    for (const att of atts) {
      const aiMatch = pickAiForAttachment(ai, att);
      const r = await uploadPdfAndMaybeJson(att, aiMatch, cfg, token);
      results.push({ attachment: att.filename, result: r });
    }
    return { ok: true, data: results };
  }

  // Append a row to an Excel table with flexible input mapping.
  // Options:
  // - driveId + itemId OR driveId + path (path like "/Documents/Book.xlsx")
  // - table: table name (required)
  // - values: array (single row) OR row: object mapping
  //   - object keys can be column names (exact match) or column letters (A, B, C ... relative to table column order)
  // - auth: { accessToken? | refreshToken?, clientId, clientSecret?, tenant? }
  // Returns: { ok, data, error }
  async function excelAppendRow(opts) {
    try {
      if (!opts || typeof opts !== 'object') return { ok: false, error: 'excelAppendRow: options required' };
      let driveId = opts.driveId;
      let itemId = opts.itemId;
      let filePath = opts.path; // optional alternative to itemId
      const table = opts.table;
      const debug = !!opts.debug;
      if (!driveId) return { ok: false, error: 'excelAppendRow: missing driveId' };
      if (!table) return { ok: false, error: 'excelAppendRow: missing table' };
      const token = await ensureAccessToken(opts.auth || {});
      if (!token) return { ok: false, error: 'excelAppendRow: no access token' };

      // If link provided, resolve to { driveId, itemId, path }
      if (!itemId && !filePath && typeof opts.link === 'string' && opts.link) {
        try {
          const info = await (async function resolveByLink(link){
            const b64mod = require('b64@latest');
            const shareId = 'u!' + b64mod.b64urlFromUtf8(String(link||''));
            const url = `${GRAPH}/shares/${encodeURIComponent(shareId)}/driveItem?$select=id,name,parentReference,webUrl`;
            const r = await sys.http.fetch({ url, method: 'GET', headers: { 'Authorization': 'Bearer ' + token } });
            const jtxt = (r && r.text) || '';
            const obj = j.parseSafe(jtxt, {});
            const pref = obj && obj.parentReference || {};
            return { driveId: pref.driveId || '', itemId: obj && obj.id || '', path: (pref.path && String(pref.path)) || '' };
          })(opts.link);
          if (info && info.driveId && info.itemId) {
            driveId = driveId || info.driveId;
            itemId = itemId || info.itemId;
            if (!filePath) {
              // parentReference.path is like "/drives/{driveId}/root:/folder/file.xlsx"
              const p = info.path || '';
              const idx = p.indexOf('/root:');
              if (idx >= 0) filePath = decodeURIComponent(p.slice(idx + 6));
            }
          }
        } catch {}
      }

      // Build base path
      let base;
      if (driveId && itemId) {
        base = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`;
      } else if (driveId && filePath) {
        base = `/drives/${encodeURIComponent(driveId)}/root:${encodeURIComponent(filePath)}`;
      } else {
        return { ok: false, error: 'excelAppendRow: provide { driveId, itemId } or { driveId, path }' };
      }

      // Normalize input into a single row array
      const rowArray = await (async () => {
        if (Array.isArray(opts.values)) return opts.values;
        const row = (opts.row && typeof opts.row === 'object') ? opts.row : (typeof opts.values === 'object' ? opts.values : null);
        if (!row) return null;
        const keys = Object.keys(row);
        const isLetters = keys.length && keys.every(k => /^[A-Za-z]+$/.test(k));
        if (isLetters) {
          // Map A, B, C... to 0-based indexes relative to the table's first column
          function colToIndex(col) {
            let n = 0; const s = ('' + col).toUpperCase();
            for (let i = 0; i < s.length; i++) { n = n * 26 + (s.charCodeAt(i) - 64); }
            return n - 1; // zero-based
          }
          let max = -1;
          const pairs = keys.map(k => { const idx = colToIndex(k); if (idx > max) max = idx; return [idx, row[k]]; });
          const arr = new Array(max + 1);
          for (const [idx, val] of pairs) arr[idx] = val;
          return arr;
        }
        // Name-based mapping: fetch table columns to build order
        const urlCols = `${GRAPH}${base}/workbook/tables/${encodeURIComponent(table)}/columns?$select=name,index`;
        const resCols = await graph.json({ path: urlCols, method: 'GET', headers: { 'Authorization': 'Bearer ' + token } });
        const jcols = resCols && resCols.data || {};
        const cols = Array.isArray(jcols.value) ? jcols.value.slice().sort((a,b)=> (a.index||0)-(b.index||0)) : [];
        if (!cols.length) return null;
        const nameToIndex = {};
        for (const c of cols) { if (c && typeof c.name === 'string') nameToIndex[c.name] = c.index || 0; }
        const arr = new Array(cols.length);
        for (const k of keys) {
          const idx = nameToIndex[k];
          if (typeof idx === 'number') arr[idx] = row[k];
        }
        return arr;
      })();

      if (!rowArray || !Array.isArray(rowArray)) return { ok: false, error: 'excelAppendRow: missing values/row' };
      const body = { values: Array.isArray(rowArray[0]) ? rowArray : [ rowArray ] };
      const url = `${GRAPH}${base}/workbook/tables/${encodeURIComponent(table)}/rows/add`;
        if (debug) log.debug('excelAppendRow: POST', url);
      const res = await graph.json({ path: url, method: 'POST', headers: { 'Authorization': 'Bearer ' + token }, bodyObj: body });
      return res && res.ok ? { ok: true, data: res.data } : { ok: false, error: (res && res.error) || 'graph error' };
    } catch (e) {
      console.error('excelAppendRow:error', (e && (e.message || e)) || 'unknown');
      return { ok: false, error: (e && (e.message || String(e))) || 'unknown' };
    }
  }

  module.exports = { uploadSmall, saveInvoices, ensureAccessToken, excelAppendRow };
})();
