/* Fonctions pures de l'appli (testables sous Node) : détection de la zone de prescription, assemblage des lignes. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.OrdoCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const norm = (t) => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z]/g, '');
  function lev(a, b) {
    const m = a.length, n = b.length;
    let prev = Array.from({ length: n + 1 }, (_, j) => j);
    for (let i = 1; i <= m; i++) {
      const cur = [i];
      for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = cur;
    }
    return prev[n];
  }

  /* Résultat Tesseract.js (blocks) -> liste de mots {text, bbox} */
  function flattenWords(data) {
    const words = [];
    for (const b of (data && data.blocks) || [])
      for (const p of b.paragraphs || []) for (const l of p.lines || []) for (const w of l.words || []) words.push(w);
    return words;
  }

  /* Zone utile = sous le mot « ORDONNANCE », au-dessus de la signature, à droite de la colonne des praticiens */
  function detectZone(words, W, H) {
    const ords = words.filter((w) => { const n = norm(w.text); return n.length >= 8 && lev(n, 'ordonnance') <= 2; })
      .sort((a, b) => a.bbox.y0 - b.bbox.y0);
    if (!ords.length) return { x: 0, y: 0, w: W, h: H, foundOrdo: false, foundSig: false };
    const o = ords[0];
    const y0 = Math.min(H, o.bbox.y1 + 0.009 * H);
    const sigs = words.filter((w) => {
      const n = norm(w.text);
      const isDoc = n === 'dr' || n === 'docteur' || (n.length >= 6 && n.endsWith('teur') && lev(n, 'docteur') <= 2) || (n.length >= 5 && lev(n, 'docteur') <= 1);
      return isDoc && w.bbox.y0 > y0 + 0.035 * H && w.bbox.x0 > 0.2 * W;
    });
    const y1 = sigs.length ? Math.min(...sigs.map((s) => s.bbox.y0)) - 0.002 * H : Math.min(H, y0 + 0.4 * H);
    const x0 = Math.max(0, o.bbox.x0 - 0.25 * W);
    const x1 = 0.97 * W;
    return { x: Math.round(x0), y: Math.round(y0), w: Math.round(x1 - x0), h: Math.round(Math.max(40, y1 - y0)), foundOrdo: true, foundSig: sigs.length > 0 };
  }

  /* Conseil final = précisions de l'ordonnance + conseil habituel (dédoublonnés) */
  function joinConseil(ordo, habituel) {
    const out = [];
    for (const part of [ordo, habituel]) {
      for (const c of String(part || '').split(/\s*;\s*/)) {
        const t = c.trim();
        if (t && !out.some((x) => x.toLowerCase() === t.toLowerCase())) out.push(t);
      }
    }
    return out.join(' ; ');
  }

  /* Ligne éditée à l'écran -> données du modèle Word */
  function rowFromItem(it) {
    const q = it.cat === 'quotidien';
    return {
      nom: it.nom || '', dci: it.dci || '',
      matin: q ? it.matin || '' : '', midi: q ? it.midi || '' : '', soir: q ? it.soir || '' : '', coucher: q ? it.coucher || '' : '',
      posologie: q ? '' : it.posologie || '',
      utilite: it.utilite || '', conseil: joinConseil(it.conseilOrdo, it.conseilHabituel),
    };
  }
  function planDataFromItems(items, date) {
    const pick = (c) => items.filter((i) => i.cat === c).map(rowFromItem);
    return { date_bilan: date, quotidien: pick('quotidien'), besoin: pick('besoin'), autre: pick('autre'), arretes: '' };
  }

  /* ---------------------------------------------------------------- conciliation d'entrée (modèle Excel) */
  // Encadré « Bilan Médicamenteux Optimisé » de la feuille FCT : lignes 19 à 31,
  // B:D fusionnées = Nom/dosage/forme, E = Voie (liste « voie » du modèle), F = Posologie.
  const XL = { sheet: 'xl/worksheets/sheet1.xml', first: 19, last: 31, nom: 'B', voie: 'E', poso: 'F' };
  const VOIES = ['PO', 'IV', 'SC', 'IM', 'Cutanée', 'Inhalée', 'Oculaire', 'Nasale', 'Rectale', 'Sublinguale', 'Transderm', 'Gastrostomie'];

  /* Voie : PO (voie orale) ou IV uniquement, laissée vide si elle n'est pas certaine.
   * 1) écrite sur l'ordonnance (IV, perfusion, per os, VO…) ; 2) sinon déduite de la BDPM
   * (toutes les présentations retenues doivent partager la même voie). */
  const TXT_IV = /(?<![\wÀ-ÿ])(?:IVL?|IVSE|IVD|intra-?veineu\w*|perfusions?|en perf|seringue [ée]lectrique)(?![\wÀ-ÿ])/i;
  const TXT_PO = /(?<![\wÀ-ÿ])(?:per os|PO|VO|voie orale|par la bouche)(?![\wÀ-ÿ])/i;
  const TXT_AUTRE = /(?<![\wÀ-ÿ])(?:SC|IM|sous-?cutan\w*|intra-?musculaire)(?![\wÀ-ÿ])/i;
  function voieFromMatch(m, text) {
    const t = String(text || '');
    if (TXT_AUTRE.test(t)) return '';
    const iv = TXT_IV.test(t), po = TXT_PO.test(t);
    if (iv !== po) return iv ? 'IV' : 'PO';
    if (!m || !m.found) return '';
    const lists = (m.voies && m.voies.length ? m.voies : [m.voie || ''])
      .map((v) => String(v).toLowerCase().split(';').map((x) => x.trim()).filter(Boolean));
    if (!lists.length || lists.some((l) => !l.length)) return '';
    const PAR = ['intraveineuse', 'intramusculaire', 'sous-cutanée'];
    if (lists.every((l) => l.includes('orale') && !l.some((x) => PAR.includes(x)))) return 'PO';
    if (lists.every((l) => l.includes('intraveineuse') && !l.includes('orale') && !l.includes('sous-cutanée'))) return 'IV';
    return '';
  }

  /* Nom/dosage/forme : nom lu + dosage, et la forme BDPM quand elle est certaine */
  function nomForme(m, nom) {
    let base = String(nom || '').trim();
    if (m && !m.found && m.dosage && !base.toLowerCase().includes(String(m.dosage).toLowerCase())) base = `${base} ${m.dosage}`.trim();
    if (!m || !m.found || m.ambiguousForm || !m.forme) return base;
    const f = String(m.forme).toLowerCase();
    return base.toLowerCase().includes(f.split(' ')[0]) ? base : `${base}, ${f}`;
  }

  const xmlEsc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  function setCell(xml, ref, text) {
    const re = new RegExp(`<c r="${ref}"((?:\\s+[\\w:]+="[^"]*")*)\\s*(?:/>|>[\\s\\S]*?</c>)`);
    if (!re.test(xml)) throw new Error(`Cellule ${ref} absente du modèle`);
    return xml.replace(re, (_, attrs) => {
      const a = attrs.replace(/\s+t="[^"]*"/, '');
      if (!String(text).trim()) return `<c r="${ref}"${a}/>`;
      return `<c r="${ref}"${a} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(text)}</t></is></c>`;
    });
  }
  /* Remplit B (nom), E (voie) et F (posologie) dans le BMO. Rien d'autre n'est modifié. */
  function fillConciliation(zip, rows) {
    let xml = zip.file(XL.sheet).asText();
    const max = XL.last - XL.first + 1;
    rows.slice(0, max).forEach((r, i) => {
      const n = XL.first + i;
      xml = setCell(xml, XL.nom + n, r.nom || '');
      xml = setCell(xml, XL.voie + n, r.voie || '');
      xml = setCell(xml, XL.poso + n, r.posologie || '');
    });
    zip.file(XL.sheet, xml);
    return { written: Math.min(rows.length, max), max };
  }

  return { voieFromMatch, nomForme, fillConciliation, setCell, XL, VOIES, flattenWords, detectZone, joinConseil, rowFromItem, planDataFromItems, norm, lev };
});
