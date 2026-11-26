/* Devoldère → Odoo v1.1 — Moore Walcourt — outil chargé à chaque clic par le favori « Devoldère → Odoo » (Pneus Devoldère, pneusdevoldere.odoo.com). */
// Pneus Devoldère — TyreStream vers Odoo — Moore Walcourt
// Favori à lancer dans l'Odoo de Pneus Devoldère (pneusdevoldere.odoo.com). Ne charge rien d'Internet.
(() => {
  'use strict';
  // ---------------------------------------------------------------- lecture des exports TyreStream (.xls)
  // TyreStream enregistre une feuille Excel 4.0 (BIFF4) brute, sans conteneur OLE : enregistrements BOF 0x0409, LABEL 0x0204
  // (texte 8 bits, page de code Windows), RK 0x027E et NUMBER 0x0203 (nombres), BLANK 0x0201, EOF 0x000A.
  function lireXls(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (u8.length < 8 || dv.getUint16(0, true) !== 0x0409) return null;
    const dec = new TextDecoder('windows-1252');
    const rows = [];
    const set = (r, c, v) => { (rows[r] = rows[r] || [])[c] = v; };
    const nombre = x => (Number.isInteger(x) ? String(x) : String(Math.round(x * 1e6) / 1e6));
    let p = 0, fin = false;
    while (p + 4 <= u8.length) {
      const t = dv.getUint16(p, true), l = dv.getUint16(p + 2, true), d = p + 4;
      if (d + l > u8.length) return null;
      if (t === 0x0204) {
        const n = dv.getUint16(d + 6, true);
        if (8 + n > l) return null;
        set(dv.getUint16(d, true), dv.getUint16(d + 2, true), dec.decode(u8.subarray(d + 8, d + 8 + n)));
      } else if (t === 0x027E) {
        const rk = dv.getInt32(d + 6, true);
        let v;
        if (rk & 2) v = rk >> 2;
        else { const b = new DataView(new ArrayBuffer(8)); b.setUint32(4, rk & 0xFFFFFFFC, true); v = b.getFloat64(0, true); }
        if (rk & 1) v /= 100;
        set(dv.getUint16(d, true), dv.getUint16(d + 2, true), nombre(v));
      } else if (t === 0x0203) {
        set(dv.getUint16(d, true), dv.getUint16(d + 2, true), nombre(dv.getFloat64(d + 6, true)));
      } else if (t === 0x000A) { fin = true; break; }
      p = d + l;
    }
    if (!fin || !rows.length) return null;
    const nc = Math.max(...rows.filter(Boolean).map(r => r.length));
    const out = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] || [];
      const ligne = [];
      for (let c = 0; c < nc; c++) ligne.push(r[c] === undefined ? '' : r[c]);
      out.push(ligne);
    }
    return out;
  }

  // ---------------------------------------------------------------- réglages du dossier Pneus Devoldère
  const VERSION = '1.1 (26/11/2026)';
  const HOTE = 'pneusdevoldere.odoo.com';
  const SOCIETE = 'Pneus Devoldere';
  const COMPTE_CLIENTS = '400000';
  const COMPTES_CAISSE = { CASH: '570000', BANCONTACT: '580001', VISA: '580001' };
  const JOURNAUX_CAISSE = { CASH: 'ts_journal_caisse', BANCONTACT: 'ts_journal_cartes', VISA: 'ts_journal_cartes' };
  const JOURNAL_VENTES = 'account.1_sale';
  const DEPART = '2026-11-11';          // premier jour repris de TyreStream (import de test du 11/11/2026 avant)
  const PREMIERE_FACTURE = 111871;      // première facture TyreStream reprise dans Odoo
  const FENETRE_JOURS = 35;             // les requêtes sortent les 35 derniers jours, jusqu'à hier
  const LIMITE_LIGNES = 10000;          // un seul lot d'import par fichier (tout ou rien)

  const COLS = {
    clients: ['ID', 'NOM', 'SOCIETE', 'REFERENCE', 'TVA', 'RUE', 'CODE_POSTAL', 'VILLE', 'PAYS', 'EMAIL'],
    ventes: ['ID', 'TYPE_PIECE', 'JOURNAL', 'NUMERO', 'CLIENT', 'DATE_FACTURE', 'ECHEANCE', 'LIGNE_LIBELLE', 'LIGNE_COMPTE',
      'LIGNE_QUANTITE', 'LIGNE_PRIX', 'LIGNE_REMISE', 'LIGNE_TAXE', 'TS_HTVA', 'TS_TVA', 'TS_TVAC'],
    caisse: ['ID', 'TYPE_PIECE', 'JOURNAL', 'DATE_PIECE', 'REFERENCE', 'ECR_COMPTE', 'ECR_CLIENT', 'ECR_LIBELLE', 'ECR_DEBIT',
      'ECR_CREDIT', 'ZTRI1', 'ZTRI2', 'ZTRI3', 'ZTRI4'],
  };
  const ENTETE_VENTES = ['ID', 'TYPE_PIECE', 'JOURNAL', 'NUMERO', 'CLIENT', 'DATE_FACTURE', 'ECHEANCE', 'TS_HTVA', 'TS_TVA', 'TS_TVAC'];
  const NOMS_FICHIERS = { clients: '1 - Clients', ventes: '2 - Ventes', caisse: '3 - Caisse' };   // noms des fichiers enregistrés
  const NOMS_REQUETES = { clients: '1 - Clients (5 semaines)', ventes: '2 - Ventes (5 semaines)', caisse: '3 - Caisse (5 semaines)' };   // bibliothèque TyreStream, groupe « Devoldere - Odoo »
  // Colonnes lues pour les contrôles, jamais importées
  const NON_IMPORTEES = new Set(['TS_HTVA', 'TS_TVA', 'TS_TVAC', 'ZTRI1', 'ZTRI2', 'ZTRI3', 'ZTRI4']);
  const MODELE = { clients: 'res.partner', ventes: 'account.move', caisse: 'account.move' };
  // Correspondance des colonnes avec les champs Odoo : celle enregistrée dans Odoo par les imports du 11/11/2026
  // (base_import.mapping). L'outil la relit et s'arrête si elle manque ou si elle ne correspond pas à celle-ci.
  // (relue dans l'Odoo de Pneus Devoldère le 26/11/2026)
  const ATTENDU = {
    'res.partner': { ID: 'id', NOM: 'name', SOCIETE: 'is_company', REFERENCE: 'ref', TVA: 'vat', RUE: 'street', CODE_POSTAL: 'zip', VILLE: 'city',
      PAYS: 'country_id', EMAIL: 'email' },
    'account.move': { ID: 'id', TYPE_PIECE: 'move_type', JOURNAL: 'journal_id/id', NUMERO: 'name', CLIENT: 'partner_id/id', DATE_FACTURE: 'invoice_date',
      ECHEANCE: 'invoice_date_due', LIGNE_LIBELLE: 'invoice_line_ids/name', LIGNE_COMPTE: 'invoice_line_ids/account_id',
      LIGNE_QUANTITE: 'invoice_line_ids/quantity', LIGNE_PRIX: 'invoice_line_ids/price_unit', LIGNE_REMISE: 'invoice_line_ids/discount',
      LIGNE_TAXE: 'invoice_line_ids/tax_ids/id', DATE_PIECE: 'date', REFERENCE: 'ref', ECR_COMPTE: 'line_ids/account_id',
      ECR_CLIENT: 'line_ids/partner_id/id', ECR_LIBELLE: 'line_ids/name', ECR_DEBIT: 'line_ids/debit', ECR_CREDIT: 'line_ids/credit' },
  };

  class Arret extends Error {
    constructor(message, details) { super(message); this.details = details || []; }
  }

  // ---------------------------------------------------------------- petits outils
  const cents = s => {
    const t = String(s == null ? '' : s).trim();
    if (!/^-?\d+(\.\d+)?$/.test(t)) return NaN;
    return Math.round(parseFloat(t) * 100);
  };
  const euros = c => (c / 100).toLocaleString('fr-BE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const isoDate = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) && !isNaN(Date.parse(s + 'T00:00:00Z'));
  const fr = iso => (iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '');
  const ajouterJours = (iso, n) => { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const aujourdhui = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
  const plages = nums => {   // [111871, 111872, 111875] -> « 111871 à 111872, 111875 »
    const s = [...nums].sort((a, b) => a - b), out = [];
    for (let i = 0; i < s.length; i++) {
      let j = i;
      while (j + 1 < s.length && s[j + 1] === s[j] + 1) j++;
      out.push(i === j ? String(s[i]) : s[i] + ' à ' + s[j]);
      i = j;
    }
    return out.join(', ');
  };
  const paquets = (arr, n) => { const out = []; for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n)); return out; };

  // ---------------------------------------------------------------- lecture et contrôle des 3 fichiers
  function reconnaitre(rows) {
    const h = (rows[0] || []).map(x => String(x).trim());
    for (const [type, cols] of Object.entries(COLS)) {
      if (h.length === cols.length && cols.every((c, i) => h[i] === c)) return type;
    }
    // ancienne requête ventes (sans les totaux TyreStream)
    if (COLS.ventes.slice(0, 13).every((c, i) => h[i] === c) && h.length === 13) return 'ventes_ancienne';
    return null;
  }

  function lireFichiers(fichiers) {   // fichiers : [{nom, buf}]
    const out = {}, err = [];
    for (const f of fichiers) {
      const rows = lireXls(f.buf);
      if (!rows) { err.push(`« ${f.nom} » : ce n'est pas un fichier enregistré par TyreStream. Reprends le fichier tel que TyreStream l'a enregistré, sans l'ouvrir dans Excel.`); continue; }
      const type = reconnaitre(rows);
      if (type === 'ventes_ancienne') { err.push(`« ${f.nom} » vient de l'ancienne requête des ventes. Utilise la requête « ${NOMS_REQUETES.ventes} » indiquée dans la procédure.`); continue; }
      if (!type) { err.push(`« ${f.nom} » n'est pas un des 3 fichiers attendus (clients, ventes, caisse). Vérifie la requête utilisée dans TyreStream.`); continue; }
      if (out[type]) { err.push(`Deux fichiers « ${type} » ont été choisis (« ${out[type].nom} » et « ${f.nom} »). Choisis un seul fichier de chaque.`); continue; }
      const cols = rows[0].map(x => String(x).trim());
      const lignes = rows.slice(1).filter(r => r.some(v => String(v).trim() !== '')).map(r => {
        const o = {}; cols.forEach((c, i) => { o[c] = String(r[i] == null ? '' : r[i]).trim(); }); return o;
      });
      out[type] = { nom: f.nom, type, cols, lignes };
    }
    for (const t of ['clients', 'ventes', 'caisse']) if (!out[t] && !err.length) err.push(`Il manque le fichier « ${NOMS_FICHIERS[t]} ». Il faut les 3 fichiers TyreStream.`);
    if (err.length) throw new Arret('Les fichiers choisis ne conviennent pas.', err);
    return out;
  }

  function analyser(F, jour, reglages) {
    const { depart, premiere } = reglages;
    const err = [], horsF = [], horsC = new Set();
    const hier = ajouterJours(jour, -1);
    const dansFenetre = d => d >= depart && d <= hier;
    // clients
    const clients = new Map();
    for (const r of F.clients.lignes) {
      if (!/^ts_(pcs_\d+|prp_\d+_\d+)$/.test(r.ID)) { err.push(`Clients : identifiant inattendu « ${r.ID} ».`); continue; }
      if (!r.NOM) err.push(`Clients : ${r.ID} sans nom.`);
      if (clients.has(r.ID)) err.push(`Clients : ${r.ID} en double.`);
      clients.set(r.ID, r);
    }
    // ventes : une facture = ses lignes consécutives de même ID. La requête « à plat » répète l'en-tête sur chaque ligne
    // (plus rapide dans TyreStream) ; l'ancienne forme (en-tête sur la 1re ligne seulement) est aussi acceptée.
    const factures = new Map();
    let cur = null;
    for (const r of F.ventes.lignes) {
      if (r.ID && cur && r.ID === cur.id) {
        const h = cur.rows[0];
        if (ENTETE_VENTES.some(c => r[c] !== h[c])) err.push(`Ventes : facture ${cur.n} : lignes d'en-tête différentes dans le fichier.`);
        cur.rows.push(r);
      } else if (r.ID) {
        const n = parseInt(r.NUMERO, 10);
        if (!/^\d+$/.test(r.NUMERO) || r.ID !== 'ts_fac_' + r.NUMERO) { err.push(`Ventes : facture « ${r.ID} » / « ${r.NUMERO} » : numéro inattendu.`); cur = null; continue; }
        if (factures.has(n)) { err.push(`Ventes : facture ${n} en double.`); cur = null; continue; }
        if (!['out_invoice', 'out_refund'].includes(r.TYPE_PIECE)) err.push(`Ventes : facture ${n} : type « ${r.TYPE_PIECE} » inattendu.`);
        if (r.JOURNAL !== JOURNAL_VENTES) err.push(`Ventes : facture ${n} : journal « ${r.JOURNAL} » inattendu.`);
        if (!/^ts_(pcs_\d+|prp_\d+_\d+)$/.test(r.CLIENT)) err.push(`Ventes : facture ${n} : client « ${r.CLIENT} » inattendu.`);
        if (!isoDate(r.DATE_FACTURE)) err.push(`Ventes : facture ${n} : date « ${r.DATE_FACTURE} » illisible.`);
        else if (!dansFenetre(r.DATE_FACTURE)) horsF.push(n);
        if (r.ECHEANCE && !isoDate(r.ECHEANCE)) err.push(`Ventes : facture ${n} : échéance « ${r.ECHEANCE} » illisible.`);
        if (n < premiere) err.push(`Ventes : facture ${n} antérieure à la reprise (${premiere}).`);
        const ts = { htva: cents(r.TS_HTVA), tva: cents(r.TS_TVA), tvac: cents(r.TS_TVAC) };
        if ([ts.htva, ts.tva, ts.tvac].some(isNaN)) err.push(`Ventes : facture ${n} : totaux TyreStream illisibles.`);
        cur = { n, id: r.ID, type: r.TYPE_PIECE, client: r.CLIENT, date: r.DATE_FACTURE, ts, rows: [r], lignes: 0 };
        factures.set(n, cur);
      } else if (!cur) { err.push('Ventes : ligne de facture sans en-tête au début du fichier.'); continue; }
      else cur.rows.push(r);
      if (!cur) continue;
      const k = cur.n;
      cur.lignes++;
      if (r.LIGNE_COMPTE === 'COMPTE_INCONNU' || !/^\d{6}$/.test(r.LIGNE_COMPTE)) err.push(`Ventes : facture ${k} : article sans compte de vente dans TyreStream (« ${r.LIGNE_LIBELLE} »). À régler avec Queenie.`);
      if (/^TAXE_INCONNUE/.test(r.LIGNE_TAXE) || !r.LIGNE_TAXE) err.push(`Ventes : facture ${k} : régime de TVA inconnu (« ${r.LIGNE_TAXE} », « ${r.LIGNE_LIBELLE} »). À régler avec Queenie.`);
      if ([r.LIGNE_QUANTITE, r.LIGNE_PRIX, r.LIGNE_REMISE].some(v => isNaN(cents(v)))) err.push(`Ventes : facture ${k} : quantité, prix ou remise illisible (« ${r.LIGNE_LIBELLE} »).`);
    }
    // caisse : une pièce par jour et par mode de paiement (en-tête = ligne du total)
    const caisses = new Map();
    let pc = null;
    for (const r of F.caisse.lignes) {
      if (r.ID) {
        const m = /^ts_caisse_(\d{4})(\d{2})(\d{2})_(CASH|BANCONTACT|VISA)$/.exec(r.ID);
        if (!m) { err.push(`Caisse : identifiant « ${r.ID} » inattendu (mode de paiement inconnu ?). À régler avec Queenie.`); pc = null; continue; }
        const date = `${m[1]}-${m[2]}-${m[3]}`, mode = m[4];
        if (caisses.has(r.ID)) { err.push(`Caisse : ${r.ID} en double.`); pc = null; continue; }
        if (r.DATE_PIECE !== date) err.push(`Caisse : ${r.ID} : date « ${r.DATE_PIECE} » ≠ ${date}.`);
        if (!dansFenetre(date)) horsC.add(date);
        if (r.JOURNAL !== JOURNAUX_CAISSE[mode]) err.push(`Caisse : ${r.ID} : journal « ${r.JOURNAL} » inattendu.`);
        if (r.TYPE_PIECE !== 'entry') err.push(`Caisse : ${r.ID} : type « ${r.TYPE_PIECE} » inattendu.`);
        if (r.ECR_COMPTE !== COMPTES_CAISSE[mode]) err.push(`Caisse : ${r.ID} : compte « ${r.ECR_COMPTE} » inattendu pour ${mode}.`);
        pc = { id: r.ID, date, mode, rows: [r], debit: 0, credit: 0, total: cents(r.ECR_DEBIT) - cents(r.ECR_CREDIT), paiements: [] };
        caisses.set(r.ID, pc);
      } else if (!pc) { err.push('Caisse : ligne sans en-tête au début du fichier.'); continue; }
      else {
        pc.rows.push(r);
        if (r.ECR_COMPTE !== COMPTE_CLIENTS) err.push(`Caisse : ${pc.id} : compte « ${r.ECR_COMPTE} » inattendu.`);
        if (!/^ts_(pcs_\d+|prp_\d+_\d+)$/.test(r.ECR_CLIENT)) err.push(`Caisse : ${pc.id} : client « ${r.ECR_CLIENT} » inattendu.`);
        const f = /^Fact (\d+)$/.exec(r.ECR_LIBELLE);
        pc.paiements.push({ client: r.ECR_CLIENT, libelle: r.ECR_LIBELLE, facture: f ? parseInt(f[1], 10) : null, montant: cents(r.ECR_CREDIT) - cents(r.ECR_DEBIT) });
      }
      const d = cents(r.ECR_DEBIT), c = cents(r.ECR_CREDIT);
      if (isNaN(d) || isNaN(c) || d < 0 || c < 0) err.push(`Caisse : ${pc.id} : montant illisible (« ${r.ECR_LIBELLE} »).`);
      else { pc.debit += d; pc.credit += c; }
    }
    if (horsF.length) err.push(`Ventes : ${horsF.length} facture(s) datée(s) avant le ${fr(depart)} ou après hier (n° ${plages(horsF)}). Refais l'export dans TyreStream avec la requête indiquée, aujourd'hui.`);
    if (horsC.size) err.push(`Caisse : ${horsC.size} jour(s) avant le ${fr(depart)} ou après hier (${[...horsC].sort().map(fr).join(', ')}). Refais l'export dans TyreStream avec la requête indiquée, aujourd'hui.`);
    for (const p of caisses.values()) if (p.debit !== p.credit) err.push(`Caisse : ${p.id} non équilibrée (débit ${euros(p.debit)}, crédit ${euros(p.credit)}).`);
    // numéros de factures : pas de trou dans le fichier
    const nums = [...factures.keys()].sort((a, b) => a - b);
    if (nums.length) {
      const manq = [];
      for (let n = nums[0]; n <= nums[nums.length - 1]; n++) if (!factures.has(n)) manq.push(n);
      if (manq.length) err.push(`Ventes : numéros absents du fichier : ${plages(manq)}. À vérifier dans TyreStream avec Queenie.`);
    }
    if (err.length) throw new Arret('Les fichiers TyreStream contiennent des points à régler. Rien n\'a été envoyé dans Odoo.', err);
    return { clients, factures, caisses, jour };
  }

  // ---------------------------------------------------------------- accès à Odoo (session ouverte dans ce navigateur)
  function creerOdoo(env) {
    const fetchF = env.fetch, attendre = env.attendre || (ms => new Promise(r => setTimeout(r, ms)));
    let contexte = {};
    async function post(url, corps, essai = 0) {
      const resp = await fetchF(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(corps), credentials: 'same-origin' });
      if ((resp.status === 429 || resp.status === 503) && essai < 10) { await attendre(2000 * (essai + 1)); return post(url, corps, essai + 1); }
      if (!resp.ok) throw new Arret(`Odoo ne répond pas correctement (erreur ${resp.status}). Réessaie dans quelques minutes ; si ça continue, appelle Queenie.`);
      const j = await resp.json();
      if (j.error) {
        const m = (j.error.data && j.error.data.message) || j.error.message || 'erreur';
        if (/session|expired|SessionExpired/i.test(m + (j.error.data && j.error.data.name))) throw new Arret('Ta connexion à Odoo a expiré. Recharge la page Odoo, reconnecte-toi, puis relance le favori.');
        const e = new Arret('Odoo a refusé une opération : ' + m + '. Appelle Queenie avec une capture de ce message.');
        e.typeOdoo = (j.error.data && j.error.data.name) || '';
        throw e;
      }
      return j.result;
    }
    const rpc = (model, method, args, kwargs = {}) => post(`/web/dataset/call_kw/${model}/${method}`,
      { jsonrpc: '2.0', method: 'call', params: { model, method, args, kwargs: Object.assign({}, kwargs, { context: Object.assign({}, contexte, kwargs.context || {}) }) } });
    return {
      rpc,
      setContexte(c) { contexte = c; },
      session: () => post('/web/session/get_session_info', { jsonrpc: '2.0', method: 'call', params: {} }),
      async envoyerFichier(id, nom, texte) {
        const fd = new env.FormData();
        fd.append('csrf_token', env.csrf());
        fd.append('id', String(id));
        fd.append('ufile', new env.Blob([texte], { type: 'text/csv' }), nom);
        for (let essai = 0; ; essai++) {
          const resp = await fetchF('/base_import/set_file', { method: 'POST', body: fd, credentials: 'same-origin' });
          if ((resp.status === 429 || resp.status === 503) && essai < 10) { await attendre(2000 * (essai + 1)); continue; }
          if (!resp.ok) throw new Arret(`Odoo a refusé le fichier (erreur ${resp.status}). Recharge la page Odoo et relance le favori.`);
          return;
        }
      },
    };
  }

  // identifiants externes « ts_… » (importés sans module : module __import__) et références « module.nom »
  async function lireXids(O, noms, modele) {
    const out = new Map();
    for (const p of paquets([...new Set(noms)], 500)) {
      const dom = [['module', '=', '__import__'], ['name', 'in', p]];
      if (modele) dom.push(['model', '=', modele]);
      for (const x of await O.rpc('ir.model.data', 'search_read', [dom], { fields: ['name', 'model', 'res_id'] })) out.set(x.name, x);
    }
    return out;
  }
  async function xidsExternes(O, refs) {   // « account.1_sale », « ts_tva_21_ttc »… -> {ref: {model, res_id}}
    const out = new Map();
    for (const ref of new Set(refs)) {
      const i = ref.indexOf('.');
      const module = i > 0 ? ref.slice(0, i) : '__import__', name = i > 0 ? ref.slice(i + 1) : ref;
      const r = await O.rpc('ir.model.data', 'search_read', [[['module', '=', module], ['name', '=', name]]], { fields: ['model', 'res_id'] });
      if (r.length === 1) out.set(ref, r[0]);
    }
    return out;
  }
  async function tousLesXids(O, prefixe) {
    return O.rpc('ir.model.data', 'search_read', [[['module', '=', '__import__'], ['model', '=', 'account.move'], ['name', '=like', prefixe + '%']]], { fields: ['name', 'res_id'] });
  }

  // ---------------------------------------------------------------- construction des fichiers envoyés à Odoo (CSV)
  const csvCellule = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
  function construireCsv(cols, lignes) {
    return [cols.map(csvCellule).join(','), ...lignes.map(r => cols.map(c => csvCellule(r[c])).join(','))].join('\r\n') + '\r\n';
  }

  // Import par le même chemin que le bouton « Importer » d'Odoo : base_import (Tester, puis Importer)
  async function importer(O, modele, nom, cols, lignes, champs, essai) {
    const id = await O.rpc('base_import.import', 'create', [{ res_model: modele }]);
    await O.envoyerFichier(id, nom, construireCsv(cols, lignes));
    const base = {
      has_headers: true, advanced: true, keep_matches: false, limit: LIMITE_LIGNES, skip: 0, sheets: [], sheet: '', tracking_disable: true,
      encoding: 'utf-8', separator: ',', quoting: '"', date_format: '%Y-%m-%d', datetime_format: '', float_thousand_separator: '',
      float_decimal_separator: '.', import_skip_records: [], import_set_empty_fields: [], fallback_values: {}, name_create_enabled_fields: {},
    };
    const prev = await O.rpc('base_import.import', 'parse_preview', [id, base]);
    if (prev.error) throw new Arret(`Odoo n'arrive pas à lire le fichier ${nom} : ${prev.error}. Appelle Queenie.`);
    const options = Object.assign({}, base, prev.options || {}, { limit: LIMITE_LIGNES, skip: 0, date_format: '%Y-%m-%d', float_thousand_separator: '', float_decimal_separator: '.' });
    const fields = cols.map(c => champs[c] || false);
    const res = await O.rpc('base_import.import', 'execute_import', [id, fields, cols, options], { dryrun: !!essai });
    const msgs = (res && res.messages) || [];
    if (msgs.length) {
      throw new Arret(`${essai ? 'Le test' : 'L\'import'} de « ${nom} » dans Odoo affiche ${msgs.length} message(s). ${essai ? 'Rien n\'a été importé.' : ''} Appelle Queenie avec une capture.`,
        msgs.slice(0, 10).map(m => `${m.type === 'warning' ? 'Avertissement' : 'Erreur'} : ${m.message}${m.rows ? ` (lignes ${m.rows.from + 1} à ${m.rows.to + 1})` : ''}`));
    }
    if (res.nextrow) throw new Arret(`Le fichier ${nom} est trop grand pour un seul envoi (plus de ${LIMITE_LIGNES} lignes). Appelle Queenie.`);
    return res.ids || [];
  }

  // Lignes d'une pièce de caisse, pour comparer Odoo et TyreStream : compte | libellé | débit | crédit (en cents)
  const signatureTs = p => p.rows.map(r => `${r.ECR_COMPTE}|${r.ECR_LIBELLE}|${cents(r.ECR_DEBIT)}|${cents(r.ECR_CREDIT)}`).sort().join('\n');
  async function signaturesOdoo(O, moveIds) {
    const moves = moveIds.length ? await O.rpc('account.move', 'read', [moveIds, ['date', 'line_ids', 'state']]) : [];
    const lignes = [];
    for (const p of paquets(moves.flatMap(m => m.line_ids), 1000)) lignes.push(...await O.rpc('account.move.line', 'read', [p, ['move_id', 'name', 'debit', 'credit', 'account_id']]));
    const comptes = new Map();
    const accIds = [...new Set(lignes.map(l => l.account_id && l.account_id[0]))].filter(Boolean);
    if (accIds.length) for (const a of await O.rpc('account.account', 'read', [accIds, ['code']])) comptes.set(a.id, a.code);
    const parMove = new Map();
    for (const l of lignes) { const k = l.move_id[0]; if (!parMove.has(k)) parMove.set(k, []); parMove.get(k).push(`${comptes.get(l.account_id[0])}|${l.name || ''}|${Math.round(l.debit * 100)}|${Math.round(l.credit * 100)}`); }
    return new Map(moves.map(m => [m.id, { date: m.date, state: m.state, sig: (parMove.get(m.id) || []).sort().join('\n') }]));
  }

  // ---------------------------------------------------------------- clients TyreStream déjà présents dans Odoo sous une autre fiche
  // Même n° de TVA qu'une seule fiche Odoo active (société, pas un contact rattaché) qui n'est encore reliée à aucun client
  // TyreStream (fiche reprise de WinBooks, ou créée à la main) : on relie. Les autres ressemblances (même n° de TVA que
  // plusieurs fiches ou qu'une fiche déjà reliée à un autre client TyreStream, même nom) ne sont que signalées dans le rapport :
  // le client est créé, Queenie fusionne les fiches dans Odoo si besoin (la fusion d'Odoo déplace aussi l'identifiant).
  const normTva = v => String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const normNom = v => String(v || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z0-9]+/g, ' ').trim();
  async function rapprocherClients(O, nouveaux) {
    const lier = [], aVerifier = [];
    if (!nouveaux.length) return { lier, aVerifier };
    const fiches = await O.rpc('res.partner', 'search_read', [[['parent_id', '=', false]]], { fields: ['name', 'vat'] });
    const parTva = new Map(), parNom = new Map();
    const ajoute = (m, k, f) => { if (!k) return; if (!m.has(k)) m.set(k, []); if (!m.get(k).includes(f)) m.get(k).push(f); };
    for (const f of fiches) {
      const t = normTva(f.vat);
      if (t) { ajoute(parTva, t, f); if (/^\d/.test(t)) ajoute(parTva, 'BE' + t, f); }
      ajoute(parNom, normNom(f.name), f);
    }
    const cands = new Map();
    for (const c of nouveaux) cands.set(c, { tva: (normTva(c.TVA) && parTva.get(normTva(c.TVA))) || [], nom: parNom.get(normNom(c.NOM)) || [] });
    const ids = [...new Set([...cands.values()].flatMap(x => [...x.tva, ...x.nom].map(f => f.id)))];
    const reliees = new Set();   // fiches déjà reliées à un client TyreStream
    for (const p of paquets(ids, 500))
      for (const x of await O.rpc('ir.model.data', 'search_read', [[['module', '=', '__import__'], ['model', '=', 'res.partner'], ['res_id', 'in', p]]], { fields: ['name', 'res_id'] }))
        if (/^ts_/.test(x.name)) reliees.add(x.res_id);
    const fiche = f => `« ${f.name} » (ID ${f.id})`;
    for (const c of nouveaux) {
      const { tva, nom } = cands.get(c);
      const libres = tva.filter(f => !reliees.has(f.id));
      if (libres.length === 1) { lier.push({ c, p: libres[0] }); continue; }
      if (libres.length > 1) aVerifier.push({ c, motif: `même n° de TVA que ${libres.length} fiches Odoo : ${libres.map(fiche).join(', ')}` });
      else if (tva.length) aVerifier.push({ c, motif: `même n° de TVA que la fiche Odoo ${tva.map(fiche).join(', ')}, déjà reliée à un autre client TyreStream` });
      else if (nom.length) aVerifier.push({ c, motif: `même nom que la fiche Odoo ${nom.map(fiche).join(', ')}` });
    }
    return { lier, aVerifier };
  }

  // ---------------------------------------------------------------- Vérifier (n'écrit rien dans la comptabilité)
  async function verifier(env, fichiers, journal) {
    const log = journal || (() => {});
    const jour = env.aujourdhui ? env.aujourdhui() : aujourdhui();
    // 1. où sommes-nous ?
    if (env.hote !== HOTE) throw new Arret(`Ce favori ne marche que dans l'Odoo de Pneus Devoldère (${HOTE}). Ouvre d'abord Odoo de Pneus Devoldère, puis clique sur le favori.`);
    const O = creerOdoo(env);
    const s = await O.session();
    if (!s || !s.uid) throw new Arret('Tu n\'es pas connecté à Odoo. Connecte-toi, puis relance le favori.');
    const cids = env.cids().length ? env.cids() : (s.user_companies && s.user_companies.current_company ? [s.user_companies.current_company] : []);
    const soc = s.user_companies && s.user_companies.allowed_companies;
    if (cids.length !== 1) throw new Arret(`Plusieurs sociétés sont cochées dans Odoo. Ne coche que « ${SOCIETE} » (en haut à droite), puis relance le favori.`);
    const nomSoc = soc && soc[cids[0]] && soc[cids[0]].name;
    if (nomSoc !== SOCIETE) throw new Arret(`La société cochée dans Odoo est « ${nomSoc || '?'} ». Coche « ${SOCIETE} » seulement, puis relance le favori.`);
    O.setContexte({ allowed_company_ids: cids, lang: (s.user_context && s.user_context.lang) || 'fr_BE', tz: (s.user_context && s.user_context.tz) || 'Europe/Brussels' });
    log(`Odoo : ${HOTE}, société « ${SOCIETE} », utilisateur ${s.name || s.username || s.uid}.`);
    // l'outil lit les identifiants des imports (ir.model.data) : il faut le droit Administration « Droits d'accès »
    try { await O.rpc('ir.model.data', 'search_count', [[['module', '=', '__import__'], ['name', '=', 'ts_controle_droits']]]); }
    catch (e) {
      if (/AccessError/.test(e.typeOdoo || '')) throw new Arret(`Ton utilisateur Odoo (${s.name || s.username || s.uid}) n'a pas les droits nécessaires pour cet outil. Appelle Queenie.`,
        ['Il faut, dans Odoo de Pneus Devoldère, le droit Administration « Droits d\'accès » (Paramètres > Utilisateurs).']);
      throw e;
    }

    // 2. les fichiers
    const F = lireFichiers(fichiers);
    log(`Fichiers lus : ${F.clients.lignes.length} lignes clients, ${F.ventes.lignes.length} lignes de ventes, ${F.caisse.lignes.length} lignes de caisse.`);
    const A = analyser(F, jour, Object.assign({ depart: DEPART, premiere: PREMIERE_FACTURE }, env.reglages || {}));

    // 3. correspondance des colonnes enregistrée dans Odoo
    const champs = { 'res.partner': {}, 'account.move': {} };
    const maps = await O.rpc('base_import.mapping', 'search_read', [[['res_model', 'in', ['res.partner', 'account.move']]]], { fields: ['res_model', 'column_name', 'field_name'] });
    for (const m of maps) champs[m.res_model][m.column_name] = m.field_name;
    const errMap = [];
    for (const [type, cols] of Object.entries(COLS)) {
      const mod = MODELE[type];
      for (const c of cols) {
        if (NON_IMPORTEES.has(c)) continue;
        const f = champs[mod][c];
        if (!f) errMap.push(`${NOMS_FICHIERS[type]} : la colonne ${c} n'a pas de correspondance d'import enregistrée dans Odoo.`);
        else if (ATTENDU[mod][c] && ATTENDU[mod][c] !== f) errMap.push(`${NOMS_FICHIERS[type]} : la colonne ${c} est reliée à « ${f} » au lieu de « ${ATTENDU[mod][c]} ».`);
      }
    }
    if (errMap.length) throw new Arret('Les réglages d\'import d\'Odoo ont changé. Rien n\'a été envoyé. Appelle Queenie.', errMap);

    // 4. dates de verrouillage
    const champsSoc = await O.rpc('res.company', 'fields_get', [], { attributes: ['type'] });
    const verrous = ['fiscalyear_lock_date', 'sale_lock_date', 'tax_lock_date', 'hard_lock_date'].filter(f => champsSoc[f]);
    const socOdoo = (await O.rpc('res.company', 'read', [[cids[0]], verrous]))[0] || {};

    // 5. ce qui est déjà dans Odoo
    const facOdoo = await tousLesXids(O, 'ts_fac_');
    const caiOdoo = await tousLesXids(O, 'ts_caisse_');
    const numsOdoo = new Map(facOdoo.map(x => [parseInt(x.name.slice(7), 10), x.res_id]));
    const caiOdooMap = new Map(caiOdoo.map(x => [x.name, x.res_id]));
    const derniere = numsOdoo.size ? Math.max(...numsOdoo.keys()) : null;
    const dernierJour = caiOdoo.length ? caiOdoo.map(x => x.name.slice(10, 18)).sort().pop().replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3') : null;
    log(`Déjà dans Odoo : dernière facture TyreStream ${derniere || 'aucune'}, dernier jour de caisse ${dernierJour ? fr(dernierJour) : 'aucun'}.`);

    const err = [];
    const nouvellesF = [], existantesF = [], brouillonsF = [], brouillonsC = [];
    for (const f of A.factures.values()) (numsOdoo.has(f.n) ? existantesF : nouvellesF).push(f);
    nouvellesF.sort((a, b) => a.n - b.n);
    // trou entre Odoo et le fichier
    if (nouvellesF.length && derniere !== null) {
      const attendue = derniere + 1;
      if (nouvellesF[0].n > attendue) {
        const manq = []; for (let n = attendue; n < nouvellesF[0].n; n++) manq.push(n);
        err.push(`Il manque les factures ${plages(manq)} : elles ne sont ni dans Odoo, ni dans le fichier. Plus de 5 semaines sans import ? Appelle Queenie.`);
      }
      const avant = nouvellesF.filter(f => f.n < derniere).map(f => f.n);
      if (avant.length) err.push(`Les factures ${plages(avant)} sont dans le fichier mais pas dans Odoo, alors que des factures plus récentes y sont déjà. Appelle Queenie.`);
    }
    // factures déjà dans Odoo : elles doivent être identiques
    if (existantesF.length) {
      const moves = await O.rpc('account.move', 'read', [existantesF.map(f => numsOdoo.get(f.n)), ['name', 'move_type', 'invoice_date', 'amount_untaxed_signed', 'amount_tax_signed', 'amount_total_signed', 'state']]);
      const parId = new Map(moves.map(m => [m.id, m]));
      for (const f of existantesF) {
        const m = parId.get(numsOdoo.get(f.n));
        if (!m) { err.push(`Facture ${f.n} : introuvable dans Odoo alors que son identifiant existe. Appelle Queenie.`); continue; }
        const o = { htva: Math.round(m.amount_untaxed_signed * 100), tva: Math.round(m.amount_tax_signed * 100), tvac: Math.round(m.amount_total_signed * 100) };
        if (o.htva !== f.ts.htva || o.tva !== f.ts.tva || o.tvac !== f.ts.tvac || m.move_type !== f.type || m.invoice_date !== f.date)
          err.push(`Facture ${f.n} déjà dans Odoo mais différente de TyreStream : Odoo ${euros(o.tvac)} TVAC du ${fr(m.invoice_date)}, TyreStream ${euros(f.ts.tvac)} TVAC du ${fr(f.date)}. Appelle Queenie.`);
        else if (m.state === 'cancel') err.push(`Facture ${f.n} annulée dans Odoo. Appelle Queenie.`);
        else if (m.state === 'draft') { f.moveId = m.id; brouillonsF.push(f); }
      }
    }
    // caisse : jours déjà dans Odoo identiques, jours nouveaux à importer
    const nouvellesC = [], existantesC = [];
    for (const p of A.caisses.values()) (caiOdooMap.has(p.id) ? existantesC : nouvellesC).push(p);
    nouvellesC.sort((a, b) => (a.date + a.mode < b.date + b.mode ? -1 : 1));
    if (existantesC.length) {
      const sig = await signaturesOdoo(O, existantesC.map(p => caiOdooMap.get(p.id)));
      for (const p of existantesC) {
        const m = sig.get(caiOdooMap.get(p.id));
        if (!m || m.date !== p.date || m.sig !== signatureTs(p)) err.push(`Caisse ${p.mode} du ${fr(p.date)} déjà dans Odoo mais différente dans TyreStream (paiement ajouté ou modifié après l'import ?). Appelle Queenie.`);
        else if (m.state === 'cancel') err.push(`Caisse ${p.mode} du ${fr(p.date)} annulée dans Odoo. Appelle Queenie.`);
        else if (m.state === 'draft') { p.moveId = caiOdooMap.get(p.id); brouillonsC.push(p); }
      }
    }
    if (dernierJour && ajouterJours(jour, -FENETRE_JOURS) > ajouterJours(dernierJour, 1) && nouvellesC.length)
      err.push(`Le dernier jour de caisse dans Odoo est le ${fr(dernierJour)} : plus de 5 semaines sans import, des jours de caisse peuvent manquer. Appelle Queenie.`);
    // clients nécessaires
    const besoin = new Set([...nouvellesF.map(f => f.client), ...nouvellesC.flatMap(p => p.paiements.map(x => x.client))]);
    const cliOdoo = await lireXids(O, [...besoin], 'res.partner');
    const nouveauxClients = [];
    for (const c of besoin) {
      if (cliOdoo.has(c)) continue;
      if (!A.clients.has(c)) err.push(`Le client ${c} n'est ni dans Odoo, ni dans le fichier clients. Refais l'export « ${NOMS_REQUETES.clients} » le même jour que les autres.`);
      else nouveauxClients.push(A.clients.get(c));
    }
    // clients nouveaux pour TyreStream mais déjà dans Odoo (fiche reprise de WinBooks ou créée à la main) : même n° de TVA
    const { lier, aVerifier } = await rapprocherClients(O, nouveauxClients);
    for (const { c } of lier) nouveauxClients.splice(nouveauxClients.indexOf(c), 1);
    // références Odoo : journaux, taxes, comptes
    const refs = [JOURNAL_VENTES, ...new Set(nouvellesC.map(p => JOURNAUX_CAISSE[p.mode])), ...new Set(nouvellesF.flatMap(f => f.rows.map(r => r.LIGNE_TAXE)))];
    const trouv = await xidsExternes(O, refs);
    for (const r of new Set(refs)) if (!trouv.has(r)) err.push(`Référence Odoo introuvable : « ${r} » (journal ou taxe). Appelle Queenie.`);
    const codes = [...new Set([...nouvellesF.flatMap(f => f.rows.map(r => r.LIGNE_COMPTE)), ...nouvellesC.flatMap(p => p.rows.map(r => r.ECR_COMPTE))])];
    const comptes = codes.length ? await O.rpc('account.account', 'search_read', [[['code', 'in', codes]]], { fields: ['code'] }) : [];
    for (const c of codes) { const n = comptes.filter(a => a.code === c).length; if (n !== 1) err.push(`Compte ${c} : ${n ? 'plusieurs comptes' : 'introuvable'} dans Odoo. Appelle Queenie.`); }
    // dates de verrouillage
    const premiereDate = [...nouvellesF.map(f => f.date), ...nouvellesC.map(p => p.date)].sort()[0];
    if (premiereDate) for (const v of verrous) if (socOdoo[v] && socOdoo[v] >= premiereDate) err.push(`Odoo est verrouillé jusqu'au ${fr(socOdoo[v])} : les pièces du ${fr(premiereDate)} ne peuvent pas entrer. Appelle Queenie.`);
    if (err.length) throw new Arret('Le contrôle avec Odoo a trouvé des points à régler. Rien n\'a été envoyé dans Odoo.', err);

    const plan = { A, F, champs, nouvellesF, existantesF, nouvellesC, existantesC, brouillonsF, brouillonsC, nouveauxClients, lier, aVerifier, derniere, dernierJour, jour, verrous, cids,
      lang: (s.user_context && s.user_context.lang) || 'fr_BE', tz: (s.user_context && s.user_context.tz) || 'Europe/Brussels' };
    // 6. test d'import (rien n'est enregistré) : clients ; ventes et caisse aussi s'il n'y a ni client à créer, ni client à relier
    if (nouveauxClients.length) {
      await importer(O, 'res.partner', 'clients.csv', colsImport('clients'), nouveauxClients, champs['res.partner'], true);
      log(`Test des ${nouveauxClients.length} nouveaux clients : aucun message.`);
    }
    if (!nouveauxClients.length && !lier.length) {
      if (nouvellesF.length) { await importer(O, 'account.move', 'ventes.csv', colsImport('ventes'), lignesFactures(nouvellesF), champs['account.move'], true); log('Test des factures : aucun message.'); }
      if (nouvellesC.length) { await importer(O, 'account.move', 'caisse.csv', colsImport('caisse'), nouvellesC.flatMap(p => p.rows), champs['account.move'], true); log('Test de la caisse : aucun message.'); }
      plan.testes = true;
    }
    return plan;
  }
  const colsImport = type => COLS[type].filter(c => !NON_IMPORTEES.has(c));
  // lignes envoyées à Odoo pour des factures : l'en-tête sur la 1re ligne seulement (une ligne avec un ID = une nouvelle pièce)
  const lignesFactures = fs => fs.flatMap(f => f.rows.map((r, i) => (i === 0 ? r : Object.assign({}, r, Object.fromEntries(ENTETE_VENTES.map(c => [c, '']))))));

  function resumePlan(plan) {
    const nf = plan.nouvellesF, nc = plan.nouvellesC;
    const r = [];
    r.push(nf.length ? `${nf.length} nouvelle(s) facture(s) : n° ${plages(nf.map(f => f.n))} (${nf.filter(f => f.type === 'out_refund').length} note(s) de crédit), total TVAC ${euros(nf.reduce((s, f) => s + f.ts.tvac, 0))}.` : 'Aucune nouvelle facture.');
    if (nc.length) {
      const jours = [...new Set(nc.map(p => p.date))].sort();
      r.push(`${nc.length} pièce(s) de caisse, ${jours.length} jour(s) du ${fr(jours[0])} au ${fr(jours[jours.length - 1])} : espèces ${euros(nc.filter(p => p.mode === 'CASH').reduce((s, p) => s + p.total, 0))}, cartes ${euros(nc.filter(p => p.mode !== 'CASH').reduce((s, p) => s + p.total, 0))}.`);
    } else r.push('Aucun nouveau jour de caisse.');
    if (plan.lier.length) r.push(`${plan.lier.length} client(s) TyreStream déjà dans Odoo sous une autre fiche (même n° de TVA), qui sera reliée : ${plan.lier.map(({ c, p }) => `${c.NOM} → ${p.name}`).join(' ; ')}.`);
    r.push(`${plan.nouveauxClients.length} nouveau(x) client(s) à créer.`);
    if (plan.aVerifier.length) r.push(`Pour Queenie, rien à faire maintenant : ${plan.aVerifier.length} nouveau(x) client(s) ressemble(nt) à une fiche Odoo existante (voir le rapport).`);
    if (plan.brouillonsF.length || plan.brouillonsC.length) r.push(`Déjà importées mais pas encore comptabilisées (envoi interrompu ?) : ${plan.brouillonsF.length} facture(s), ${plan.brouillonsC.length} pièce(s) de caisse. Elles seront comptabilisées.`);
    const sautF = plan.existantesF.length - plan.brouillonsF.length, sautC = plan.existantesC.length - plan.brouillonsC.length;
    if (sautF || sautC) r.push(`Déjà dans Odoo et identiques (sautés) : ${sautF} facture(s), ${sautC} pièce(s) de caisse.`);
    return r;
  }

  // ---------------------------------------------------------------- Envoyer
  async function envoyer(env, fichiers, journal) {
    const log = journal || (() => {});
    const plan = await verifier(env, fichiers, log);   // on refait tous les contrôles juste avant d'écrire
    const O = creerOdoo(env);
    O.setContexte({ allowed_company_ids: plan.cids, lang: plan.lang, tz: plan.tz });
    const rapport = [];
    const R = (type, ref, statut, detail) => rapport.push({ type, ref, statut, detail: detail || '' });
    const { nouvellesF, nouvellesC, nouveauxClients, champs } = plan;
    if (!nouvellesF.length && !nouvellesC.length && !plan.brouillonsF.length && !plan.brouillonsC.length) {
      log('Rien de nouveau à envoyer : tout est déjà dans Odoo.');
      return { plan, rapport, fini: true, rien: true };
    }
    // 1. clients : relier ceux qui ont déjà une fiche Odoo (même n° de TVA), puis créer les nouveaux
    if (plan.lier.length) {
      await O.rpc('ir.model.data', 'create', [plan.lier.map(({ c, p }) => ({ module: '__import__', name: c.ID, model: 'res.partner', res_id: p.id, noupdate: false }))]);
      const x = await lireXids(O, plan.lier.map(({ c }) => c.ID), 'res.partner');
      for (const { c, p } of plan.lier) {
        if (!x.has(c.ID) || x.get(c.ID).res_id !== p.id) throw new Arret(`Client ${c.NOM} : la liaison avec la fiche Odoo « ${p.name} » n'a pas été enregistrée. Appelle Queenie.`);
        R('Client', c.ID, 'relié à une fiche existante', `${c.NOM} → fiche Odoo « ${p.name} » (ID ${p.id}), même n° de TVA ${c.TVA}`);
      }
      log(`${plan.lier.length} client(s) relié(s) à leur fiche Odoo existante.`);
    }
    if (nouveauxClients.length) {
      await importer(O, 'res.partner', 'clients.csv', colsImport('clients'), nouveauxClients, champs['res.partner'], true);
      const ids = await importer(O, 'res.partner', 'clients.csv', colsImport('clients'), nouveauxClients, champs['res.partner'], false);
      if (ids.length !== nouveauxClients.length) throw new Arret(`Clients : ${ids.length} créé(s) au lieu de ${nouveauxClients.length}. Appelle Queenie.`);
      const motifs = new Map(plan.aVerifier.map(v => [v.c, v.motif]));
      nouveauxClients.forEach(c => R('Client', c.ID, motifs.has(c) ? 'créé - à vérifier par Queenie' : 'créé', motifs.has(c) ? `${c.NOM} : ${motifs.get(c)}` : c.NOM));
      log(`${ids.length} client(s) créé(s).`);
    }
    // 2. tests des factures et de la caisse, avant d'importer la moindre pièce
    const rowsF = lignesFactures(nouvellesF), rowsC = nouvellesC.flatMap(p => p.rows);
    if (rowsF.length) await importer(O, 'account.move', 'ventes.csv', colsImport('ventes'), rowsF, champs['account.move'], true);
    if (rowsC.length) await importer(O, 'account.move', 'caisse.csv', colsImport('caisse'), rowsC, champs['account.move'], true);
    // 3. factures
    let movesF = [];
    if (nouvellesF.length) {
      const rows = rowsF;
      const ids = await importer(O, 'account.move', 'ventes.csv', colsImport('ventes'), rows, champs['account.move'], false);
      log(`${ids.length} facture(s) importée(s) en brouillon.`);
      const x = await lireXids(O, nouvellesF.map(f => f.id), 'account.move');
      const moves = await O.rpc('account.move', 'read', [nouvellesF.map(f => x.get(f.id) && x.get(f.id).res_id).filter(Boolean),
        ['name', 'move_type', 'invoice_date', 'amount_untaxed_signed', 'amount_tax_signed', 'amount_total_signed', 'state', 'invoice_line_ids', 'partner_id', 'journal_id']]);
      const pm = new Map(moves.map(m => [m.id, m]));
      const err = [];
      for (const f of nouvellesF) {
        const m = x.get(f.id) && pm.get(x.get(f.id).res_id);
        if (!m) { err.push(`Facture ${f.n} absente après l'import.`); continue; }
        const o = { htva: Math.round(m.amount_untaxed_signed * 100), tva: Math.round(m.amount_tax_signed * 100), tvac: Math.round(m.amount_total_signed * 100) };
        const pb = [];
        if (o.htva !== f.ts.htva) pb.push(`HTVA Odoo ${euros(o.htva)} / TyreStream ${euros(f.ts.htva)}`);
        if (o.tva !== f.ts.tva) pb.push(`TVA Odoo ${euros(o.tva)} / TyreStream ${euros(f.ts.tva)}`);
        if (o.tvac !== f.ts.tvac) pb.push(`TVAC Odoo ${euros(o.tvac)} / TyreStream ${euros(f.ts.tvac)}`);
        if (m.state !== 'draft') pb.push(`état ${m.state}`);
        if (m.move_type !== f.type) pb.push(`type ${m.move_type}`);
        if (m.invoice_date !== f.date) pb.push(`date ${m.invoice_date}`);
        if (String(m.name) !== String(f.n)) pb.push(`numéro Odoo « ${m.name} »`);
        if ((m.invoice_line_ids || []).length !== f.lignes) pb.push(`${(m.invoice_line_ids || []).length} lignes au lieu de ${f.lignes}`);
        if (pb.length) err.push(`Facture ${f.n} : ${pb.join(' ; ')}`);
        f.moveId = m.id;
      }
      if (err.length) throw new Arret('Les factures sont dans Odoo en brouillon, mais le contrôle avec TyreStream ne tombe pas juste. Rien n\'a été comptabilisé. Appelle Queenie.', err);
      log('Contrôle des factures : chaque facture = TyreStream au cent (HTVA, TVA, TVAC).');
      movesF = nouvellesF.map(f => f.moveId);
    }
    // 4. caisse
    let movesC = [];
    if (nouvellesC.length) {
      const rows = rowsC;
      const ids = await importer(O, 'account.move', 'caisse.csv', colsImport('caisse'), rows, champs['account.move'], false);
      log(`${ids.length} pièce(s) de caisse importée(s) en brouillon.`);
      const x = await lireXids(O, nouvellesC.map(p => p.id), 'account.move');
      const sig = await signaturesOdoo(O, nouvellesC.map(p => x.get(p.id) && x.get(p.id).res_id).filter(Boolean));
      const err = [];
      for (const p of nouvellesC) {
        const id = x.get(p.id) && x.get(p.id).res_id, m = id && sig.get(id);
        if (!m) { err.push(`Caisse ${p.mode} du ${fr(p.date)} absente après l'import.`); continue; }
        const pb = [];
        if (m.date !== p.date) pb.push(`date ${m.date}`);
        if (m.sig !== signatureTs(p)) pb.push('lignes ou montants différents de TyreStream');
        if (m.state !== 'draft') pb.push(`état ${m.state}`);
        if (pb.length) err.push(`Caisse ${p.mode} du ${fr(p.date)} : ${pb.join(' ; ')}`);
        p.moveId = id;
      }
      if (err.length) throw new Arret('La caisse est dans Odoo en brouillon, mais le contrôle avec TyreStream ne tombe pas juste. Rien n\'a été comptabilisé. Appelle Queenie.', err);
      log('Contrôle de la caisse : chaque pièce = TyreStream au cent.');
      movesC = nouvellesC.map(p => p.moveId);
    }
    // 5. comptabiliser (avec les pièces d'un envoi interrompu, importées mais pas comptabilisées)
    movesF = [...movesF, ...plan.brouillonsF.map(f => f.moveId)];
    movesC = [...movesC, ...plan.brouillonsC.map(p => p.moveId)];
    const aPoster = [...nouvellesF, ...plan.brouillonsF], aPosterC = [...nouvellesC, ...plan.brouillonsC];
    for (const p of paquets([...movesF, ...movesC], 50)) await O.rpc('account.move', 'action_post', [p]);
    const apres = await O.rpc('account.move', 'read', [[...movesF, ...movesC], ['state', 'name', 'date', 'invoice_date', 'amount_total_signed', 'move_type']]);
    const nonPost = apres.filter(m => m.state !== 'posted');
    const errPost = nonPost.map(m => `Pièce ${m.name} : état ${m.state}`);
    for (const f of aPoster) {
      const m = apres.find(x => x.id === f.moveId);
      if (m && (String(m.name) !== String(f.n) || m.invoice_date !== f.date || m.date !== f.date || Math.round(m.amount_total_signed * 100) !== f.ts.tvac))
        errPost.push(`Facture ${f.n} modifiée par Odoo à la comptabilisation (numéro « ${m.name} », date ${fr(m.date)}).`);
    }
    for (const p of aPosterC) { const m = apres.find(x => x.id === p.moveId); if (m && m.date !== p.date) errPost.push(`Caisse du ${fr(p.date)} datée du ${fr(m.date)} par Odoo.`); }
    if (errPost.length) throw new Arret('Odoo n\'a pas comptabilisé toutes les pièces comme prévu. Appelle Queenie.', errPost);
    aPoster.forEach(f => R(f.type === 'out_refund' ? 'Note de crédit' : 'Facture', f.n, 'comptabilisée', `${fr(f.date)} ; TVAC ${euros(f.ts.tvac)}`));
    aPosterC.forEach(p => R('Caisse', `${fr(p.date)} ${p.mode}`, 'comptabilisée', `${euros(p.total)} ; ${p.paiements.length} paiement(s)`));
    log(`${movesF.length + movesC.length} pièce(s) comptabilisée(s).`);
    // 5. lettrage (clients 400000) : facture + paiements « Fact n° » du même client, soldés au cent
    const lettrage = await lettrer(O, plan, log);
    lettrage.forEach(x => R('Lettrage', x.facture, x.statut, x.detail));
    plan.existantesF.filter(f => !plan.brouillonsF.includes(f)).forEach(f => R('Facture', f.n, 'déjà dans Odoo (sautée)', ''));
    plan.existantesC.filter(p => !plan.brouillonsC.includes(p)).forEach(p => R('Caisse', `${fr(p.date)} ${p.mode}`, 'déjà dans Odoo (sautée)', ''));
    return { plan, rapport, fini: true };
  }

  async function lettrer(O, plan, log) {
    const out = [];
    const nums = new Set([...plan.nouvellesF, ...plan.brouillonsF].map(f => f.n).concat([...plan.nouvellesC, ...plan.brouillonsC].flatMap(p => p.paiements.map(x => x.facture)).filter(Boolean)));
    const acc = await O.rpc('account.account', 'search_read', [[['code', '=', COMPTE_CLIENTS]]], { fields: ['id'] });
    if (acc.length !== 1) { log('Lettrage non fait : compte 400000 introuvable.'); return out; }
    const caiX = await tousLesXids(O, 'ts_caisse_');
    const movesCaisse = caiX.map(x => x.res_id);
    // paiements de caisse encore ouverts d'un envoi précédent (lettrage interrompu, facture payée en plusieurs fois) :
    // repris seulement si leur facture est une facture TyreStream de la reprise
    const ouverts = await O.rpc('account.move.line', 'search_read', [[['move_id', 'in', movesCaisse], ['account_id', '=', acc[0].id], ['parent_state', '=', 'posted'], ['reconciled', '=', false]]], { fields: ['name'] });
    const anciens = [...new Set(ouverts.map(l => (/^Fact (\d+)$/.exec(l.name || '') || [])[1]).filter(Boolean).map(Number))].filter(n => !nums.has(n));
    if (anciens.length) { const x = await lireXids(O, anciens.map(n => 'ts_fac_' + n), 'account.move'); anciens.filter(n => x.has('ts_fac_' + n)).forEach(n => nums.add(n)); }
    if (!nums.size) return out;
    const facX = await lireXids(O, [...nums].map(n => 'ts_fac_' + n), 'account.move');
    const champsL = ['move_id', 'name', 'balance', 'amount_residual', 'partner_id', 'reconciled'];
    const lignesF = facX.size ? await O.rpc('account.move.line', 'search_read', [[['move_id', 'in', [...facX.values()].map(x => x.res_id)], ['account_id', '=', acc[0].id], ['parent_state', '=', 'posted']]], { fields: champsL }) : [];
    const libelles = [...nums].map(n => 'Fact ' + n);
    const lignesP = [];
    for (const p of paquets(libelles, 300)) lignesP.push(...await O.rpc('account.move.line', 'search_read', [[['move_id', 'in', movesCaisse], ['account_id', '=', acc[0].id], ['parent_state', '=', 'posted'], ['name', 'in', p]]], { fields: champsL }));
    const moveDeFac = new Map([...facX.entries()].map(([k, v]) => [v.res_id, parseInt(k.slice(7), 10)]));
    const groupes = new Map();
    const g = n => { if (!groupes.has(n)) groupes.set(n, { fac: [], pay: [] }); return groupes.get(n); };
    lignesF.forEach(l => g(moveDeFac.get(l.move_id[0])).fac.push(l));
    lignesP.forEach(l => g(parseInt(l.name.slice(5), 10)).pay.push(l));
    for (const n of [...nums].sort((a, b) => a - b)) {
      const G = groupes.get(n);
      if (!G || !G.fac.length) { if (G && G.pay.length) out.push({ facture: n, statut: 'non lettré', detail: 'facture pas dans Odoo comme facture TyreStream (avant la reprise ?)' }); continue; }
      if (!G.pay.length) { out.push({ facture: n, statut: 'reste ouverte', detail: 'aucun paiement en caisse pour l\'instant' }); continue; }
      const tout = [...G.fac, ...G.pay];
      if (tout.every(l => l.reconciled)) continue;
      if (tout.some(l => l.reconciled || Math.round(l.amount_residual * 100) !== Math.round(l.balance * 100))) { out.push({ facture: n, statut: 'non lettré', detail: 'déjà lettrée en partie : à voir avec Queenie' }); continue; }
      const partenaires = new Set(tout.map(l => l.partner_id && l.partner_id[0]));
      if (partenaires.size !== 1) { out.push({ facture: n, statut: 'non lettré', detail: 'paiement à un autre client que la facture' }); continue; }
      const solde = tout.reduce((s, l) => s + Math.round(l.balance * 100), 0);
      if (solde !== 0) { out.push({ facture: n, statut: 'reste ouverte', detail: `reste ${euros(solde)} (payée en partie)` }); continue; }
      await O.rpc('account.move.line', 'reconcile', [tout.map(l => l.id)]);
      out.push({ facture: n, statut: 'lettrée', detail: `${G.pay.length} paiement(s)` });
    }
    const nb = out.filter(x => x.statut === 'lettrée').length;
    log(`Lettrage : ${nb} facture(s) lettrée(s) avec leurs paiements ; ${out.length - nb} restent ouvertes ou à voir.`);
    return out;
  }

  function rapportCsv(res, erreur) {
    const l = [['Type', 'Référence', 'Statut', 'Détail']];
    if (erreur) { l.push(['ARRÊT', '', erreur.message, '']); (erreur.details || []).forEach(d => l.push(['', '', '', d])); }
    for (const r of (res && res.rapport) || []) l.push([r.type, r.ref, r.statut, r.detail]);
    return '﻿' + l.map(x => x.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(';')).join('\r\n') + '\r\n';
  }

  // ---------------------------------------------------------------- fenêtre de l'outil (dans la page Odoo)
  function demarrer() {
    const ancien = document.getElementById('tsodoo-panneau');
    if (ancien) { ancien.style.display = 'block'; return; }
    const env = {
      fetch: window.fetch.bind(window), FormData: window.FormData, Blob: window.Blob, hote: location.hostname,
      csrf: () => (window.odoo && window.odoo.csrf_token) || '',
      cids: () => { const m = document.cookie.match(/(?:^|;\s*)cids=([^;]+)/); return m ? decodeURIComponent(m[1]).split(/[-,]/).map(Number).filter(Boolean) : []; },
    };
    const css = `
#tsodoo-panneau{position:fixed;top:64px;right:18px;width:470px;max-height:86vh;overflow:auto;z-index:2147483000;background:#fff;border:2px solid #1C9CD9;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,.25);font:15px/1.45 Arial,sans-serif;color:#1D1E1C}
#tsodoo-panneau *{box-sizing:border-box}
#tsodoo-panneau .tt{background:#1C9CD9;color:#fff;padding:12px 16px;font-weight:bold;font-size:17px;display:flex;justify-content:space-between;align-items:center}
#tsodoo-panneau .tt small{font-weight:normal;font-size:11px;opacity:.85;display:block}
#tsodoo-panneau .x{cursor:pointer;font-size:22px;line-height:1;background:none;border:0;color:#fff}
#tsodoo-panneau .co{padding:14px 16px}
#tsodoo-panneau .et{margin:0 0 14px}
#tsodoo-panneau .et b{display:block;margin-bottom:6px}
#tsodoo-panneau button.bt{font:bold 15px Arial,sans-serif;border:0;border-radius:8px;padding:11px 16px;cursor:pointer;width:100%}
#tsodoo-panneau button.bleu{background:#1C9CD9;color:#fff}
#tsodoo-panneau button.vert{background:#2E8B3D;color:#fff}
#tsodoo-panneau button.orange{background:#C55A11;color:#fff}
#tsodoo-panneau button.gris{background:#e6e6e6;color:#555;margin-top:6px}
#tsodoo-panneau button:disabled{opacity:.4;cursor:not-allowed}
#tsodoo-panneau .fi{font-size:13px;color:#444;margin-top:6px}
#tsodoo-panneau .bx{border-radius:8px;padding:10px 12px;margin:0 0 14px;font-size:14px}
#tsodoo-panneau .ok{background:#E8F5EA;border:1px solid #2E8B3D}
#tsodoo-panneau .ko{background:#FDECEC;border:1px solid #C62828}
#tsodoo-panneau .en{background:#EAF6FC;border:1px solid #1C9CD9}
#tsodoo-panneau .bx .ti{font-weight:bold;font-size:15px;margin-bottom:4px}
#tsodoo-panneau .bx ul{margin:4px 0 0 18px;padding:0}
#tsodoo-panneau details{font-size:12px;color:#666}
#tsodoo-panneau details div{white-space:pre-wrap}`;
    const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
    const P = document.createElement('div'); P.id = 'tsodoo-panneau';
    P.innerHTML = `
<div class="tt"><div>Pneus Devoldère — TyreStream vers Odoo<small>Outil Moore Walcourt, version ${VERSION}</small></div><button class="x" title="Fermer">×</button></div>
<div class="co">
  <div class="et"><b>1. Choisis les 3 fichiers TyreStream</b>
    <button class="bt bleu" data-a="choisir">Choisir les fichiers</button>
    <input type="file" multiple accept=".xls,.XLS" style="display:none">
    <div class="fi">Aucun fichier choisi.</div></div>
  <div class="et"><b>2. Vérifie</b><button class="bt bleu" data-a="verifier" disabled>Vérifier</button></div>
  <div class="bx en" data-z="res"><div class="ti">En attente</div>Choisis les 3 fichiers, puis clique sur « Vérifier ». Rien n'est envoyé à cette étape.</div>
  <div class="et"><b>3. Envoie dans Odoo</b><button class="bt vert" data-a="envoyer" disabled>Envoyer dans Odoo</button>
    <button class="bt gris" data-a="annuler" style="display:none">Annuler</button></div>
  <button class="bt gris" data-a="rapport" style="display:none">Télécharger le rapport</button>
  <details><summary>Détails</summary><div data-z="log"></div></details>
</div>`;
    document.body.appendChild(P);
    const $ = s => P.querySelector(s);
    const input = $('input[type=file]'), bChoisir = $('[data-a=choisir]'), bVerif = $('[data-a=verifier]'), bEnv = $('[data-a=envoyer]'),
      bAnn = $('[data-a=annuler]'), bRap = $('[data-a=rapport]'), zRes = $('[data-z=res]'), zLog = $('[data-z=log]'), zFi = $('.fi');
    let fichiers = [], plan = null, confirmer = false, occupe = false, dernierRapport = null;
    const log = m => { zLog.textContent += new Date().toLocaleTimeString('fr-BE') + '  ' + m + '\n'; };
    const esc = t => String(t).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
    const boite = (cls, titre, lignes, texte) => {
      zRes.className = 'bx ' + cls;
      zRes.innerHTML = `<div class="ti">${esc(titre)}</div>${texte ? esc(texte) : ''}${lignes && lignes.length ? '<ul>' + lignes.map(l => `<li>${esc(l)}</li>`).join('') + '</ul>' : ''}`;
    };
    const occuper = (o, msg) => {
      occupe = o;
      [bChoisir, bVerif, bEnv, bAnn].forEach(b => { b.disabled = o; });
      if (!o) { bVerif.disabled = fichiers.length === 0; bEnv.disabled = !plan; }
      if (o) boite('en', msg || 'Patiente…', [], 'Ne ferme pas cette page et ne clique nulle part dans Odoo.');
    };
    const telecharger = (texte, nom) => {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([texte], { type: 'text/csv;charset=utf-8' }));
      a.download = nom; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
    };
    const horodatage = () => { const d = new Date(); const z = n => String(n).padStart(2, '0'); return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}_${z(d.getHours())}h${z(d.getMinutes())}`; };
    const reset = () => { plan = null; confirmer = false; bEnv.textContent = 'Envoyer dans Odoo'; bEnv.className = 'bt vert'; bAnn.style.display = 'none'; bEnv.disabled = true; };
    $('.x').onclick = () => { if (!occupe) { P.remove(); st.remove(); } };
    bChoisir.onclick = () => input.click();
    input.onchange = async () => {
      reset();
      fichiers = await Promise.all([...input.files].map(async f => ({ nom: f.name, buf: await f.arrayBuffer() })));
      zFi.textContent = fichiers.length ? fichiers.map(f => '• ' + f.nom).join('\n') : 'Aucun fichier choisi.';
      zFi.style.whiteSpace = 'pre-line';
      bVerif.disabled = fichiers.length === 0;
      boite('en', fichiers.length === 3 ? 'Fichiers choisis' : `${fichiers.length} fichier(s) choisi(s)`, [], fichiers.length === 3 ? 'Clique maintenant sur « Vérifier ».' : 'Il faut les 3 fichiers TyreStream (clients, ventes, caisse).');
      input.value = '';
    };
    const erreur = (e, titre) => {
      const det = e instanceof Arret ? e.details : [String(e && e.stack || e)];
      boite('ko', titre || (e instanceof Arret ? e.message : 'Problème inattendu. Appelle Queenie avec une capture de cette fenêtre.'), det);
      log('ARRÊT : ' + (e && e.message) + (det.length ? '\n  ' + det.join('\n  ') : ''));
      dernierRapport = rapportCsv(null, e instanceof Arret ? e : new Arret(String(e && e.message || e)));
      bRap.style.display = 'block';
    };
    bVerif.onclick = async () => {
      reset(); bRap.style.display = 'none';
      occuper(true, 'Vérification en cours…');
      try {
        plan = await verifier(env, fichiers, log);
        const lignes = resumePlan(plan);
        if (!plan.nouvellesF.length && !plan.nouvellesC.length && !plan.brouillonsF.length && !plan.brouillonsC.length) { plan = null; boite('ok', 'Tout est déjà dans Odoo', lignes, 'Il n\'y a rien de nouveau à envoyer.'); }
        else boite('ok', 'Tout est prêt. Rien n\'a encore été envoyé.', lignes, plan.testes ? '' : 'Les nouveaux clients seront créés d\'abord, puis les factures et la caisse seront testées avant l\'envoi.');
      } catch (e) { plan = null; erreur(e); }
      occuper(false);
    };
    bAnn.onclick = () => { confirmer = false; bEnv.textContent = 'Envoyer dans Odoo'; bEnv.className = 'bt vert'; bAnn.style.display = 'none'; };
    bEnv.onclick = async () => {
      if (!plan) return;
      if (!confirmer) { confirmer = true; bEnv.textContent = 'Oui, envoyer maintenant'; bEnv.className = 'bt orange'; bAnn.style.display = 'block'; return; }
      bAnn.style.display = 'none';
      occuper(true, 'Envoi en cours… (quelques minutes)');
      let res = null;
      try {
        res = await envoyer(env, fichiers, log);
        const n = res.rapport;
        const nf = n.filter(r => /Facture|Note/.test(r.type) && r.statut === 'comptabilisée').length, nc = n.filter(r => r.type === 'Caisse' && r.statut === 'comptabilisée').length;
        const nl = n.filter(r => r.type === 'Lettrage' && r.statut === 'lettrée').length, no = n.filter(r => r.type === 'Lettrage' && r.statut !== 'lettrée').length;
        boite('ok', res.rien ? 'Rien de nouveau : tout était déjà dans Odoo.' : 'C\'est fait ! Tout est dans Odoo.', res.rien ? [] : [
          `${nf} facture(s) et note(s) de crédit comptabilisée(s).`, `${nc} pièce(s) de caisse comptabilisée(s).`,
          `${n.filter(r => r.type === 'Client' && /^créé/.test(r.statut)).length} nouveau(x) client(s)` + (n.some(r => r.type === 'Client' && /^relié/.test(r.statut)) ? `, ${n.filter(r => r.type === 'Client' && /^relié/.test(r.statut)).length} relié(s) à leur fiche existante.` : '.'), `${nl} facture(s) lettrée(s) avec leurs paiements, ${no} restée(s) ouverte(s) (normal si pas encore payées).`],
          'Le rapport est enregistré dans Téléchargements.');
        dernierRapport = rapportCsv(res, null);
        telecharger(dernierRapport, `Devoldere_TyreStream_Odoo_rapport_${horodatage()}.csv`);
        bRap.style.display = 'block';
      } catch (e) { erreur(e); }
      plan = null; confirmer = false; bEnv.textContent = 'Envoyer dans Odoo'; bEnv.className = 'bt vert';
      occuper(false);
    };
    bRap.onclick = () => dernierRapport && telecharger(dernierRapport, `Devoldere_TyreStream_Odoo_rapport_${horodatage()}.csv`);
  }

  demarrer();
})();
