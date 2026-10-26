/* Moore - Outils v1.0.0 — Moore Walcourt — aiguillage du favori « Moore - Outils », chargé à chaque clic.
   Dans MyMinfin (…minfin.fgov.be) : charge moore-myminfin.js ; partout ailleurs (bases Odoo, odoo.com) : charge moore-odoo.js.
   Les deux fichiers sont publiés à côté de celui-ci ; chacun se met à jour séparément. */
(async () => {
  const base = (window.__moore && window.__moore.base) || '';
  if (!/^https?:\/\//.test(base)) throw new Error('adresse du code en ligne inconnue : réinstaller le favori depuis sa page');
  const fichier = /(^|\.)minfin\.fgov\.be$/i.test(location.hostname) ? 'moore-myminfin.js' : 'moore-odoo.js';
  const r = await fetch(base + fichier + '?v=' + Date.now(), { cache: 'no-store', referrerPolicy: 'no-referrer' })
    .catch(e => { throw new Error('code en ligne inaccessible (' + e.message + ')'); });
  if (!r.ok) throw new Error('code en ligne introuvable (HTTP ' + r.status + ', ' + fichier + ')');
  const t = await r.text();
  try { (0, eval)(t); } catch (e) {
    if (!(e instanceof EvalError)) throw e;
    const s = document.createElement('script'); s.textContent = t; document.head.append(s); s.remove();
  }
})().catch(e => alert('Moore - Outils : ' + ((e && e.message) || e)));
