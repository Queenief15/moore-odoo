/* Moore - MyMinfin v1.7.0 — Moore Walcourt — outils MyMinfin : compte-provisions TVA et documents.
   Code du favori, chargé à chaque clic. Mise à jour : remplacer ce fichier dans le dépôt GitHub ; tous les favoris l’utilisent au clic suivant.
   Secours : dans MyMinfin, client sélectionné, F12 > Console, coller tout ce fichier, Entrée (la première fois : taper « allow pasting »). */
(async () => {
  'use strict';
  const VERSION = '1.7.0';

  async function outilCompteProvisions() {
  'use strict';
  const VERSION = 'v1.4.1';
  const HOST_ID = 'moore-mmf-compte-provisions';

  // ---------------------------------------------------------------- outils
  const clean = s => String(s ?? '').replace(/[\s  ]+/g, ' ').trim();
  const txt = el => (el ? clean(el.innerText || el.textContent) : '');
  const simple = s => clean(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();

  // « 1.697,63 € » -> 1697.63 ; « - » ou vide -> null
  const montant = s => {
    const m = String(s ?? '').replace(/[\s  ]/g, '').match(/[-−]?\d[\d.]*(?:,\d+)?/);
    return m ? parseFloat(m[0].replace('−', '-').replace(/\./g, '').replace(',', '.')) : null;
  };
  const xlNombre = n => (n == null ? '' : n.toFixed(2).replace('.', ','));
  const affiche = n => {
    if (n == null) return '—';
    const [e, d] = Math.abs(n).toFixed(2).split('.');
    return (n < 0 ? '-' : '') + e.replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + d + ' €';
  };
  // Texte que Excel ne doit pas transformer (formule, zéro de tête comme 0430.658.521, nombre trop long)
  const xlTexte = s => (/^[=+\-@]/.test(s) || /^\d{9,}$/.test(s) || /^0\d/.test(s) ? "'" + s : s);

  // « 31 juillet 2026 » / « 1er août 2026 » / « 08/09/2026 » -> « 31/07/2026 »
  const MOIS = [/^jan/, /^f.{0,2}v|^feb/, /^mar|^maa/, /^av|^apr/, /^ma[iy]|^mei/, /^juin|^jun/,
                /^juil|^jul/, /^ao|^aug/, /^sep/, /^oct|^okt/, /^nov/, /^d.{0,2}c/];
  const dateFR = s => {
    const t = simple(s);
    let m = t.match(/(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/);
    if (m) return `${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}/${m[3]}`;
    m = t.match(/(\d{1,2})(?:er)?\s+(\S+)\s+(\d{4})/);
    if (!m) return clean(s);
    const i = MOIS.findIndex(re => re.test(m[2]));
    return i < 0 ? clean(s) : `${m[1].padStart(2, '0')}/${String(i + 1).padStart(2, '0')}/${m[3]}`;
  };
  const triDate = d => { const m = String(d).match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + m[2] + m[1] : '99999999'; };

  const pause = ms => new Promise(r => setTimeout(r, ms));
  const attendre = async (test, ms) => {
    const fin = Date.now() + ms;
    for (;;) { const v = test(); if (v) return v; if (Date.now() > fin) return null; await pause(150); }
  };
  // Petit message en haut à droite pendant une lecture
  const bulle = texte => {
    const d = document.createElement('div');
    d.textContent = texte;
    d.style.cssText = 'position:fixed;top:16px;right:16px;z-index:2147483647;background:#16324f;color:#fff;' +
      'font:600 13px/1.4 system-ui,Segoe UI,Arial,sans-serif;padding:10px 14px;border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,.25)';
    document.body.appendChild(d);
    return d;
  };

  const pad = n => String(n).padStart(2, '0');
  const now = new Date();
  const consulte = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;

  // Tableau Angular Material : colonnes repérées par leur classe « cdk-column-… »
  const lireMatTable = table => {
    if (!table) return null;
    const cols = [...table.querySelectorAll('thead th')]
      .map(th => ({ cls: [...th.classList].find(c => c.startsWith('cdk-column-')), label: txt(th) }))
      .filter(c => c.cls && c.label && simple(c.label) !== 'actions');
    const lignes = [...table.querySelectorAll('tbody tr')]
      .filter(tr => !/detail-row|no-data/.test(tr.className) && !tr.querySelector('[class*="no-data"]'))
      .map(tr => Object.fromEntries(cols.map(c => [c.cls.replace('cdk-column-', ''), txt(tr.querySelector('td.' + c.cls))])))
      .filter(l => Object.values(l).some(v => v));
    return { entetes: cols.map(c => c.label), lignes };
  };

  const avert = [];

  // ------------------------------------------------- Excel aux couleurs de Moore
  // Bibliothèque Excel et styles repris tels quels du favori Moore - Odoo 3.3.6 : mêmes logo,
  // couleurs, en-tête, fiche du dossier et chiffres clés que les concordances et le grand livre.
  const XL = (()=>{const e=e=>String(e).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\ufffe\uffff]/g,"").replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g,"$1").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"),t=e=>{let t="";for(;e>0;e=Math.floor((e-1)/26))t=String.fromCharCode(65+(e-1)%26)+t;return t},o=(e,o)=>t(o)+e,n=e=>{const[t,o,n]=e.split("-").map(Number);return(Date.UTC(t,o-1,n)-Date.UTC(1899,11,30))/864e5},r=e=>"FF"+String(e).replace("#","").toUpperCase();class s{constructor(){this.fonts=['<font><sz val="10"/><color rgb="FF1D1E1C"/><name val="Arial"/><family val="2"/></font>'],this.fills=['<fill><patternFill patternType="none"/></fill>','<fill><patternFill patternType="gray125"/></fill>'],this.borders=["<border><left/><right/><top/><bottom/><diagonal/></border>"],this.numFmts=[],this.xfs=['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'],this.dxfs=[],this.cache=new Map}idx(e,t){let o=e.indexOf(t);return o<0&&(e.push(t),o=e.length-1),o}font(e={}){return this.idx(this.fonts,"<font>"+(e.b?"<b/>":"")+(e.i?"<i/>":"")+(e.u?"<u/>":"")+'<sz val="'+(e.sz||10)+'"/><color rgb="'+r(e.color||"1D1E1C")+'"/><name val="'+(e.name||"Arial")+'"/><family val="2"/></font>')}fill(e){return e?this.idx(this.fills,'<fill><patternFill patternType="solid"><fgColor rgb="'+r(e)+'"/><bgColor indexed="64"/></patternFill></fill>'):0}border(e){if(!e)return 0;const t=t=>{const o=e[t]||e.all;if(!o)return"<"+t+"/>";const[n,s]=o.split(" ");return"<"+t+' style="'+n+'"><color rgb="'+r(s||"BFBFBF")+'"/></'+t+">"};return this.idx(this.borders,"<border>"+t("left")+t("right")+t("top")+t("bottom")+"<diagonal/></border>")}numFmt(e){if(!e)return 0;const t={General:0,0:1,"0.00":2,"#,##0":3,"#,##0.00":4,"0%":9,"0.00%":10,"@":49};if(e in t)return t[e];let o=this.numFmts.indexOf(e);return o<0&&(this.numFmts.push(e),o=this.numFmts.length-1),164+o}get(e){if(null==e)return 0;if("number"==typeof e)return e;const t=JSON.stringify(e);if(this.cache.has(t))return this.cache.get(t);const o=this.font(e.font),n=this.fill(e.fill),r=this.border(e.border),s=this.numFmt(e.fmt),i=e.h||e.v||e.wrap||e.indent||e.shrink?"<alignment"+(e.h?' horizontal="'+e.h+'"':"")+(e.v?' vertical="'+e.v+'"':"")+(e.wrap?' wrapText="1"':"")+(e.indent?' indent="'+e.indent+'"':"")+(e.shrink?' shrinkToFit="1"':"")+"/>":"",a='<xf numFmtId="'+s+'" fontId="'+o+'" fillId="'+n+'" borderId="'+r+'" xfId="0"'+(s?' applyNumberFormat="1"':"")+(o?' applyFont="1"':"")+(n?' applyFill="1"':"")+(r?' applyBorder="1"':"")+(i?' applyAlignment="1">'+i+"</xf>":"/>"),l=this.idx(this.xfs,a);return this.cache.set(t,l),l}dxf(e){return this.idx(this.dxfs,"<dxf>"+(e.color||e.b?"<font>"+(e.b?"<b/>":"")+(e.color?'<color rgb="'+r(e.color)+'"/>':"")+"</font>":"")+(e.fill?'<fill><patternFill patternType="solid"><fgColor rgb="'+r(e.fill)+'"/><bgColor rgb="'+r(e.fill)+'"/></patternFill></fill>':"")+"</dxf>")}xml(){return'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+(this.numFmts.length?'<numFmts count="'+this.numFmts.length+'">'+this.numFmts.map((t,o)=>'<numFmt numFmtId="'+(164+o)+'" formatCode="'+e(t)+'"/>').join("")+"</numFmts>":"")+'<fonts count="'+this.fonts.length+'">'+this.fonts.join("")+'</fonts><fills count="'+this.fills.length+'">'+this.fills.join("")+'</fills><borders count="'+this.borders.length+'">'+this.borders.join("")+'</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="'+this.xfs.length+'">'+this.xfs.join("")+'</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="'+this.dxfs.length+'">'+this.dxfs.join("")+'</dxfs><tableStyles count="0"/></styleSheet>'}}class i{constructor(e,t){if(/[\[\]:*?\/\\]/.test(t)||t.length>31||!t.trim())throw new Error("Nom de feuille invalide : "+t);this.wb=e,this.name=t,this.rows=new Map,this.heights={},this.widths={},this.merges=[],this.validations=[],this.cfs=[],this.opts={grid:!1,zoom:100},this.hidden=new Set,this.images=[]}set(e,t,n,r){if("string"==typeof e&&(r=n,n=t,[e,t]=(e=>{const t=/^\$?([A-Z]+)\$?(\d+)$/.exec(e);if(!t)throw new Error("R\xe9f\xe9rence invalide "+e);let o=0;for(const e of t[1])o=26*o+e.charCodeAt(0)-64;return[+t[2],o]})(e)),!Number.isInteger(e)||!Number.isInteger(t)||e<1||t<1||e>1048576||t>16384)throw new Error("Cellule hors limites "+e+","+t);if("number"==typeof n&&!Number.isFinite(n))throw new Error("Montant non num\xe9rique en "+o(e,t));let s=this.rows.get(e);return s||(s=new Map,this.rows.set(e,s)),s.set(t,{v:n,s:this.wb.styles.get(r)}),this}style(e,t,o,n,r){for(let s=e;s<=o;s++)for(let e=t;e<=n;e++){const t=this.rows.get(s),o=t&&t.get(e);this.set(s,e,o?o.v:null,r)}}width(e,t){return this.widths[e]=t,this}height(e,t){return this.heights[e]=t,this}merge(e,t,n,r){return this.merges.push(o(e,t)+":"+o(n,r)),this}freeze(e,t){return this.opts.freeze=[e,t],this}outline(e,t){return(this.lv||(this.lv=new Map)).set(e,t),this}filter(e,t,n,r){return this.opts.filter=o(e,t)+":"+o(Math.max(n,e+1),r),this}list(e,t){return this.validations.push({sqref:e,values:t}),this}cf(e,t,o){return this.cfs.push({sqref:e,formula:t,dxf:this.wb.styles.dxf(o)}),this}print(e){return this.opts.print=e,this}image(e,t){return this.images.push(Object.assign({png:e,col:1,row:1,dx:0,dy:0},t)),this}xml(t){const s=this.opts,i=[];let a=1,l=1;for(const t of[...this.rows.keys()].sort((e,t)=>e-t)){const r=this.rows.get(t),s=[];for(const i of[...r.keys()].sort((e,t)=>e-t)){const{v:c,s:d}=r.get(i),p=o(t,i);a=Math.max(a,t),l=Math.max(l,i);const u=d?' s="'+d+'"':"";if(null==c||""===c)s.push('<c r="'+p+'"'+u+"/>");else if("number"==typeof c)s.push('<c r="'+p+'"'+u+"><v>"+c+"</v></c>");else if("boolean"==typeof c)s.push('<c r="'+p+'"'+u+' t="b"><v>'+(c?1:0)+"</v></c>");else if("object"==typeof c&&void 0!==c.f)s.push('<c r="'+p+'"'+u+"><f>"+e(String(c.f).replace(/^=/,""))+"</f></c>");else if("object"==typeof c&&void 0!==c.d)s.push('<c r="'+p+'"'+u+"><v>"+n(c.d)+"</v></c>");else{const t=String(c).slice(0,32767);s.push('<c r="'+p+'"'+u+' t="inlineStr"><is><t xml:space="preserve">'+e(t)+"</t></is></c>")}}const c=this.heights[t],lv=this.lv&&this.lv.get(t)||0,lv2=this.lv&&this.lv.get(t+1)||0;i.push('<row r="'+t+'"'+(c?' ht="'+c+'" customHeight="1"':"")+(lv?' outlineLevel="'+lv+'"'+(this.opts.replie?' hidden="1"':""):"")+(this.opts.replie&&lv2>lv?' collapsed="1"':"")+">"+s.join("")+"</row>")}for(const e in this.heights)this.rows.has(+e)||i.push('<row r="'+e+'" ht="'+this.heights[e]+'" customHeight="1"/>');i.sort((e,t)=>+/r="(\d+)"/.exec(e)[1]-+/r="(\d+)"/.exec(t)[1]);const c=s.print;let d='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';if(d+="<sheetPr>"+(s.tab?'<tabColor rgb="'+r(s.tab)+'"/>':"")+(this.lv?'<outlinePr summaryBelow="0"/>':"")+(c?'<pageSetUpPr fitToPage="1"/>':"")+"</sheetPr>",d+='<dimension ref="A1:'+o(a,l)+'"/>',d+='<sheetViews><sheetView workbookViewId="0"'+(0===t?' tabSelected="1"':"")+(s.grid?"":' showGridLines="0"')+(100!==s.zoom?' zoomScale="'+s.zoom+'" zoomScaleNormal="'+s.zoom+'"':"")+">",s.freeze){const[e,t]=s.freeze,n=t-1,r=e-1,i=n&&r?"bottomRight":r?"bottomLeft":"topRight";d+="<pane"+(n?' xSplit="'+n+'"':"")+(r?' ySplit="'+r+'"':"")+' topLeftCell="'+o(e,t)+'" activePane="'+i+'" state="frozen"/><selection pane="'+i+'" activeCell="'+o(e,t)+'" sqref="'+o(e,t)+'"/>'}d+='</sheetView></sheetViews><sheetFormatPr defaultRowHeight="13.2"'+(this.lv?' outlineLevelRow="'+[...this.lv.values()].reduce((e,t)=>e>t?e:t,0)+'"':"")+"/>";const p=Object.keys(this.widths).map(Number).sort((e,t)=>e-t);p.length&&(d+="<cols>"+p.map(e=>'<col min="'+e+'" max="'+e+'" width="'+this.widths[e]+'" customWidth="1"'+(this.hidden.has(e)?' hidden="1"':"")+"/>").join("")+"</cols>"),d+="<sheetData>"+i.join("")+"</sheetData>",s.filter&&(d+='<autoFilter ref="'+s.filter+'"/>'),this.merges.length&&(d+='<mergeCells count="'+this.merges.length+'">'+this.merges.map(e=>'<mergeCell ref="'+e+'"/>').join("")+"</mergeCells>");let u=1;for(const t of this.cfs)d+='<conditionalFormatting sqref="'+t.sqref+'"><cfRule type="expression" dxfId="'+t.dxf+'" priority="'+u+++'"><formula>'+e(t.formula)+"</formula></cfRule></conditionalFormatting>";this.validations.length&&(d+='<dataValidations count="'+this.validations.length+'">'+this.validations.map(t=>'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Valeur" error="Choisissez dans la liste." sqref="'+t.sqref+'"><formula1>&quot;'+e(t.values.join(","))+"&quot;</formula1></dataValidation>").join("")+"</dataValidations>"),c&&(d+='<printOptions horizontalCentered="1"/>'),d+='<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.6" header="0.3" footer="0.3"/>';const m=(e,t)=>t?"&"+e+'&"Arial"&8'+String(t).replace(/&/g,"&&"):"";return c&&(d+='<pageSetup paperSize="9" orientation="'+(c.landscape?"landscape":"portrait")+'" fitToWidth="1" fitToHeight="'+(c.onePage?1:0)+'"/><headerFooter><oddFooter>'+e(m("L",c.footer)+m("C",c.mid)+'&R&"Arial"&8Page &P / &N')+"</oddFooter></headerFooter>"),this.images.length&&(d+='<drawing r:id="rId1"/>'),d+"</worksheet>"}}const a=(()=>{const e=new Uint32Array(256);for(let t=0;t<256;t++){let o=t;for(let e=0;e<8;e++)o=1&o?3988292384^o>>>1:o>>>1;e[t]=o>>>0}return e})(),l=e=>{let t=4294967295;for(let o=0;o<e.length;o++)t=a[255&(t^e[o])]^t>>>8;return(4294967295^t)>>>0};async function c(e){if("undefined"==typeof CompressionStream)return null;try{const t=new CompressionStream("deflate-raw"),o=t.writable.getWriter();o.write(e),o.close();const n=[],r=t.readable.getReader();for(;;){const{done:e,value:t}=await r.read();if(e)break;n.push(t)}const s=new Uint8Array(n.reduce((e,t)=>e+t.length,0));let i=0;for(const e of n)s.set(e,i),i+=e.length;return s}catch(e){return null}}async function d(e){const t=new TextEncoder,o=[],n=[];let r=0;const s=new Date,i=s.getHours()<<11|s.getMinutes()<<5|s.getSeconds()>>1,a=s.getFullYear()-1980<<9|s.getMonth()+1<<5|s.getDate();for(const s of e){const e=t.encode(s.name),d="string"==typeof s.data?t.encode(s.data):s.data,p=l(d),u=await c(d),m=u&&u.length<d.length?8:0,f=m?u:d,h=new DataView(new ArrayBuffer(30));h.setUint32(0,67324752,!0),h.setUint16(4,20,!0),h.setUint16(6,2048,!0),h.setUint16(8,m,!0),h.setUint16(10,i,!0),h.setUint16(12,a,!0),h.setUint32(14,p,!0),h.setUint32(18,f.length,!0),h.setUint32(22,d.length,!0),h.setUint16(26,e.length,!0),h.setUint16(28,0,!0);const g=new DataView(new ArrayBuffer(46));g.setUint32(0,33639248,!0),g.setUint16(4,20,!0),g.setUint16(6,20,!0),g.setUint16(8,2048,!0),g.setUint16(10,m,!0),g.setUint16(12,i,!0),g.setUint16(14,a,!0),g.setUint32(16,p,!0),g.setUint32(20,f.length,!0),g.setUint32(24,d.length,!0),g.setUint16(28,e.length,!0),g.setUint32(42,r,!0),o.push(new Uint8Array(h.buffer),e,f),n.push(new Uint8Array(g.buffer),e),r+=30+e.length+f.length}const d=n.reduce((e,t)=>e+t.length,0),p=new DataView(new ArrayBuffer(22));p.setUint32(0,101010256,!0),p.setUint16(8,e.length,!0),p.setUint16(10,e.length,!0),p.setUint32(12,d,!0),p.setUint32(16,r,!0);const u=[...o,...n,new Uint8Array(p.buffer)],m=new Uint8Array(u.reduce((e,t)=>e+t.length,0));let f=0;for(const e of u)m.set(e,f),f+=e.length;return m}return{Workbook:class{constructor(e={}){this.styles=new s,this.sheets=[],this.meta=e,this.names=[],this.props=[]}sheet(e){const t=new i(this,e);return this.sheets.push(t),t}defineName(e,t,o){const n=String(e),r=String(o).toUpperCase().replace(/\$?([A-Z]+)\$?(\d+)/g,"$$$1$$$2");if(!/^[A-Za-z_][\w.]{0,254}$/.test(n)||/^_xlnm\./i.test(n)||/^([A-Za-z]{1,3}\d+|[RrCc]|[Rr]\d*[Cc]\d*)$/.test(n))throw new Error("Nom d\xe9fini invalide : "+n);if(!/^\$[A-Z]{1,3}\$\d+(:\$[A-Z]{1,3}\$\d+)?$/.test(r))throw new Error("R\xe9f\xe9rence invalide "+o);if(this.names.some(e=>e.name.toLowerCase()===n.toLowerCase()))throw new Error("Nom d\xe9fini en double : "+n);return this.names.push({name:n,sheet:String(t),ref:r}),this}prop(e,t){const o=String(e),n=null==t?"":String(t);if(!o.trim()||o.length>255)throw new Error("Nom de propri\xe9t\xe9 invalide : "+o);const r=this.props.find(e=>e[0].toLowerCase()===o.toLowerCase());return r?r[1]=n:this.props.push([o,n]),this}async build(){const t=this.sheets,o=e=>"'"+e.replace(/'/g,"''")+"'",n=this.props,r=[],s=(e,t,o,n)=>r.push({name:e,ref:t,local:o,hidden:n});t.forEach((e,t)=>{e.opts.filter&&s("_xlnm._FilterDatabase",o(e.name)+"!"+e.opts.filter.replace(/([A-Z]+)(\d+)/g,"$$$1$$$2"),t,!0);const n=e.opts.print;n&&n.area&&s("_xlnm.Print_Area",o(e.name)+"!"+n.area,t),n&&n.titles&&s("_xlnm.Print_Titles",o(e.name)+"!"+n.titles,t)});for(const e of this.names){if(!t.some(t=>t.name===e.sheet))throw new Error("Nom d\xe9fini "+e.name+" : feuille absente ("+e.sheet+")");s(e.name,o(e.sheet)+"!"+e.ref,-1)}const i=e=>e.name.toLowerCase();r.sort((e,t)=>i(e)<i(t)?-1:i(e)>i(t)?1:e.local-t.local);const a=r.map(t=>'<definedName name="'+e(t.name)+'"'+(t.local>=0?' localSheetId="'+t.local+'"':"")+(t.hidden?' hidden="1"':"")+">"+e(t.ref)+"</definedName>"),l=(new Date).toISOString().replace(/\.\d+Z$/,"Z"),c=[{name:"[Content_Types].xml",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+t.map((e,t)=>'<Override PartName="/xl/worksheets/sheet'+(t+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'+(e.images.length?'<Override PartName="/xl/drawings/drawing'+(t+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>':"")).join("")+'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'+(n.length?'<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>':"")+"</Types>"},{name:"_rels/.rels",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'+(n.length?'<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>':"")+"</Relationships>"},{name:"docProps/core.xml",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>'+e(this.meta.title||"")+"</dc:title><dc:creator>"+e(this.meta.creator||"")+'</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">'+l+'</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">'+l+"</dcterms:modified></cp:coreProperties>"},{name:"docProps/app.xml",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>'},...n.length?[{name:"docProps/custom.xml",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'+n.map(([t,o],n)=>'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="'+(n+2)+'" name="'+e(t)+'"><vt:lpwstr>'+e(o)+"</vt:lpwstr></property>").join("")+"</Properties>"}]:[],{name:"xl/_rels/workbook.xml.rels",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+t.map((e,t)=>'<Relationship Id="rId'+(t+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+(t+1)+'.xml"/>').join("")+'<Relationship Id="rId'+(t.length+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'},{name:"xl/workbook.xml",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000" activeTab="0"/></bookViews><sheets>'+t.map((t,o)=>'<sheet name="'+e(t.name)+'" sheetId="'+(o+1)+'" r:id="rId'+(o+1)+'"/>').join("")+"</sheets>"+(a.length?"<definedNames>"+a.join("")+"</definedNames>":"")+'<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'}];t.forEach((e,t)=>c.push({name:"xl/worksheets/sheet"+(t+1)+".xml",data:e.xml(t)}));let p=0;return t.forEach((t,o)=>{if(!t.images.length)return;const n=o+1,r=9525,s=[],i=[];t.images.forEach((t,o)=>{const n=++p;c.push({name:"xl/media/image"+n+".png",data:t.png}),s.push('<Relationship Id="rId'+(o+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image'+n+'.png"/>'),i.push("<xdr:oneCellAnchor><xdr:from><xdr:col>"+(t.col-1)+"</xdr:col><xdr:colOff>"+t.dx*r+"</xdr:colOff><xdr:row>"+(t.row-1)+"</xdr:row><xdr:rowOff>"+t.dy*r+'</xdr:rowOff></xdr:from><xdr:ext cx="'+t.w*r+'" cy="'+t.h*r+'"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="'+(o+2)+'" name="'+e(t.name||"Image "+n)+'"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId'+(o+1)+'"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+t.w*r+'" cy="'+t.h*r+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>')}),c.push({name:"xl/drawings/drawing"+n+".xml",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+i.join("")+"</xdr:wsDr>"}),c.push({name:"xl/drawings/_rels/drawing"+n+".xml.rels",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+s.join("")+"</Relationships>"}),c.push({name:"xl/worksheets/_rels/sheet"+n+".xml.rels",data:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing'+n+'.xml"/></Relationships>'})}),c.push({name:"xl/styles.xml",data:this.styles.xml()}),d(c)}},ref:o,absRef:(e,o)=>"$"+t(o)+"$"+e,colName:t,serial:n,crc32:l,zip:d}})();
  const MS = (()=>{const e={BLEU:"1C9CD9",NOIR:"1D1E1C",BLEU_TXT:"0E6C9E",TEINTE:"E3F2FA",TEINTE2:"F2F9FD",GRIS_FOND:"EDEFF1",FILET:"D9DCDF",GRIS:"6B7075",KPI:"F6F7F8",ROUGE:"C00000",ROSE:"FDF1F1",VERT:"1E7B4F",VERT_BG:"E8F5EE",AMBRE:"B45309",AMBRE_BG:"FFF7E6",SAISIE:"E9ECEF"},t='#,##0.00;\\-#,##0.00;""',o='#,##0.00;\\-#,##0.00;"-"',n='#,##0.00 "\u20ac";\\-#,##0.00 "\u20ac";0.00 "\u20ac"',r='0.0%;\\-0.0%;""',s="dd/mm/yyyy",i="hair "+e.FILET,a={top:"thin "+e.NOIR,bottom:"medium "+e.BLEU},l="FFFFFF",c={title:{font:{b:!0,sz:16,color:e.NOIR},v:"center"},sub:{font:{b:!0,sz:10,color:e.BLEU}},note:{font:{sz:8,color:e.GRIS},wrap:!0,v:"top"},noteL:{font:{sz:8,color:e.GRIS}},noteB:{font:{b:!0,sz:8,color:e.GRIS}},rule:{font:{sz:8,color:e.GRIS},border:{bottom:"medium "+e.BLEU}},brand:{font:{sz:8,color:e.GRIS},h:"right",v:"top",wrap:!0},secT:{font:{b:!0,sz:10.5,color:e.NOIR},v:"bottom",border:{bottom:"thin "+e.BLEU}},subsec:{font:{b:!0,sz:9,color:e.BLEU}},keyL:{font:{sz:9,color:e.GRIS},border:{bottom:i}},inp:{font:{sz:9},fill:e.SAISIE,border:{top:i,bottom:i}},inpWrap:{font:{sz:9},fill:e.SAISIE,border:{all:i},wrap:!0,v:"top"},th:{font:{b:!0,sz:9,color:l},fill:e.NOIR,h:"center",v:"center",wrap:!0},thl:{font:{b:!0,sz:9,color:l},fill:e.NOIR,h:"left",v:"center",wrap:!0,indent:1},thB:{font:{b:!0,sz:9,color:l},fill:e.BLEU,h:"left",v:"center",indent:1},thG:{font:{b:!0,sz:8,color:e.GRIS},h:"center",v:"center",wrap:!0,border:{bottom:"thin "+e.GRIS}},td:{font:{sz:9},border:{bottom:i}},tdB:{font:{b:!0,sz:9},border:{bottom:i}},tdGris:{font:{sz:9,color:e.GRIS},border:{bottom:i}},tdc:{font:{sz:9},border:{bottom:i},h:"center"},tdWrap:{font:{sz:8.5,color:e.GRIS},border:{bottom:i},wrap:!0,v:"center"},num:{font:{sz:9},fmt:o,border:{bottom:i}},numB:{font:{b:!0,sz:9},fmt:o,border:{bottom:i}},numGris:{font:{sz:9,color:e.GRIS},fmt:t,border:{bottom:i}},numInp:{font:{sz:9},fmt:t,fill:e.SAISIE,border:{top:i,bottom:i}},pct:{font:{sz:9},fmt:r,border:{bottom:i}},date:{font:{sz:9},fmt:s,border:{bottom:i},h:"left"},incl:{font:{b:!0,sz:9},fill:e.SAISIE,border:{top:i,bottom:i},h:"center"},hi:{font:{b:!0,sz:9.5},fill:e.KPI,border:{top:i,bottom:i}},hiNum:{font:{b:!0,sz:9.5},fill:e.KPI,fmt:o,border:{top:i,bottom:i}},tot:{font:{b:!0,sz:9.5},fill:e.GRIS_FOND,border:a},totNum:{font:{b:!0,sz:9.5},fill:e.GRIS_FOND,fmt:o,border:a},totEur:{font:{b:!0,sz:9.5},fill:e.GRIS_FOND,fmt:n,border:a},kpiL:{font:{b:!0,sz:7.5,color:e.GRIS},fill:e.KPI,v:"bottom",indent:1},kpiV:{font:{b:!0,sz:13,color:e.NOIR},fill:e.KPI,fmt:"#,##0.00;\\-#,##0.00;0.00",h:"left",v:"center",indent:1},kpiT:{font:{b:!0,sz:10,color:e.NOIR},fill:e.KPI,h:"left",v:"center",indent:1,shrink:!0},kpiIn:{font:{b:!0,sz:10,color:e.NOIR},fill:e.SAISIE,h:"left",v:"center",indent:1,shrink:!0,border:{bottom:i}},kpiS:{font:{b:!0,sz:11,color:e.NOIR},fill:e.KPI,h:"left",v:"center",indent:1},status:{font:{b:!0,sz:9.5},h:"center",v:"center",fill:e.GRIS_FOND,border:a},ok:{font:{b:!0,sz:9,color:e.VERT},border:{bottom:i}},ko:{font:{b:!0,sz:9,color:e.ROUGE},border:{bottom:i}},warn:{font:{b:!0,sz:9,color:e.AMBRE},border:{bottom:i}}},d={ok:{color:e.VERT,fill:e.VERT_BG,b:!0},ko:{color:e.ROUGE,fill:e.ROSE,b:!0},warn:{color:e.AMBRE,fill:e.AMBRE_BG,b:!0},okT:{color:e.VERT,b:!0},koT:{color:e.ROUGE,b:!0},warnT:{color:e.AMBRE,b:!0},off:{color:"A6A6A6"}},p=()=>Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAQoAAABIAgMAAABzQlOJAAAADFBMVEUQEA/+//8dnNlhfYvVScCxAAAACXBIWXMAAAsTAAALEwEAmpwYAAADZ0lEQVR42sXYPY6bQBQH8GeQFQUJueIQce8yEkeg8CxZi2jrbOUbxJdwT4OEOQVHSBF3OcQ222VF5r35YOZhF/sSKS7wMobfDH8eM2hB/fWngn9hPPKmx/cbB970TWCcWdNFYHRxy4PEYOccJMYQtzyLjDiQQWREgTzIjOikg8yIAnkWGmEgg9DowjiExiWMQ2gMYRxS4xzEITW6IA6pcQnikBpDEIfYOM9xWKOervj1/RW3L/Bqfn+CDbVOE/2M39NkjW6Owxp72OEX5HpzBAASa/0HNo9gWoA+1rjMcXhjTSehgQfiHmGpM3JmDHMc3ljRVh/ZgD3VnFNYI+XG2cfhjZTiyk3vNPQ92BGNtiU2Oh+HN4AGn+MZn2vKoYL0+oLjGzfTbxwQbIJMzYmH2NCDLrVR44hoQKWGarTHNV7hWpmUnDH4OGZjhxecm3Rr6t4MrSCjXhpnF8dsrLHTXB9S4K3ACDHmRttoqDHnRufimI0VbnLdsbmt9nbjloxyYVxcHN7QPVegOytTc+B1TwNS3lhxY3BxeOMHqGOqOyvxCvQlFI0xNHtnHDqQITZ+QVGu9IFjbvaLyhS8Nu8Z3QMzfsIOcm/UsLPvKs5YZjpcDsy4wgdY+84WBt1bqtvZGJ65oSu68IbSBoWrjitbY7ulseVGqR+JO4au8RFrnRl9xg39eKg7Bp2sFkabcKPRVXYjD2eslsYJuIGTR2mLgWUKZkLj9wVgyww9fBXXR2ikamH0ABk3xkL5Ot1jnV5dnW6m8YbRAiTcKK9kpKbWr034vFDRMuOkh8eNigraPLcV+GmaDAKZgSltmfGVDDN/lGAneTt/EBgbPRoZM8yD1dCRdCvTYB6jtSEyWjSSmwbN8I2ZT3H8dj4dV9w4UbncNBSWwojlcNQrZYMiGhhTPK+blWJ70yjd0kSLFa0Wa5rY4vWlN3vZTaNy61wNtj7R2Lvn1hmt2UtuGrXt3Y5IqXj+cIaJwwSyMOZ1v7HYaO9UZFiCAonfP97wb71amkX9aNbu6dX8NAXvH70zMvl7UOuMRG64OCgQoeEJDERm9LORSY12NhKpMceBgciMgNCBiIw+NDKZ0YZGIjPCOHQgIiMiYCsx+tjIJEYbG4nEiOMAkBiMgE/vN75w4+P7jaeJfd7+y/+l/gBsEz4CzhGsSQAAAABJRU5ErkJggg=="),e=>e.charCodeAt(0)),u=e=>String(e).padStart(2,"0"),m=e=>u(e.getDate())+"/"+u(e.getMonth()+1)+"/"+e.getFullYear();function f(e,t,o,n,r,s){e.set(t,o,r,s);for(let r=o+1;r<=n;r++)e.set(t,r,null,s);n>o&&e.merge(t,o,t,n)}return{K:e,S:c,DXF:d,NUM:t,NUMD:o,EUR:n,PCT:r,DATE:s,logo:p,mset:f,header:function(t,o){const n=o.lastCol,r=o.brandFrom||Math.max(1,n-1);t.image(p(),{col:1,row:1,dx:5,dy:12,w:132,h:36,name:"Logo Moore"}),t.height(1,40),f(t,1,r,n,"Moore Walcourt\n\xc9tabli le "+m(o.date||new Date),c.brand),t.set(2,1,o.title,c.title),t.height(2,24),o.subtitle&&t.set(3,1,o.subtitle,c.sub),t.height(3,16);for(let e=1;e<=n;e++)t.set(4,e,1===e&&o.note||null,c.rule);return t.height(5,8),t.opts.tab=void 0===o.tab?e.NOIR:o.tab,6},kpi:function(e,t,o){e.height(t,16),e.height(t+1,26);for(const[n,r,s,i,a]of o)f(e,t,n,r,s,c.kpiL),f(e,t+1,n,r,i,a||c.kpiV)},secTitle:function(e,t,o,n,r){for(let s=o;s<=n;s++)e.set(t,s,s===o?r:null,c.secT);e.height(t,20)},thRow:function(e,t,o,n,r){n.forEach((n,s)=>e.set(t,o+s,n,(r||1)>s?c.thl:c.th)),e.height(t,18)},statusCF:function(e,t,o,n){const r=n?[d.okT,d.koT,d.warnT]:[d.ok,d.ko,d.warn];e.cf(t,o+'="Concordant"',r[0]),e.cf(t,o+'="\xc9cart justifi\xe9"',r[0]),e.cf(t,o+'="\xc9cart \xe0 justifier"',r[1]),e.cf(t,"LEFT("+o+',11)="\xc0 compl\xe9ter"',r[2])},q:e=>"'"+e.replace(/'/g,"''")+"'",dateFr:m,dateTimeFr:e=>m(e)+" \xe0 "+u(e.getHours())+"h"+u(e.getMinutes())}})();

  // ---------------------------------------------------------------- client
  const ident = document.querySelector('#user-info-details .myminfin-details-identifier');
  const client = txt(ident && ident.querySelector(':scope > span'));
  const bceSrc = txt(ident && ident.querySelector(':scope > div')) || txt(document.querySelector('.profil-number'));
  const bceM = bceSrc.match(/([01]\d{3})[.\s]?(\d{3})[.\s]?(\d{3})/);
  const bce = bceM ? `${bceM[1]}.${bceM[2]}.${bceM[3]}` : '';
  if (!client) avert.push('Aucun client sélectionné : les données sont celles du profil connecté.');

  // ------------------------------------------------ blocs de « Mes paiements »
  // MyMinfin affiche « Mes paiements » avec les blocs fermés, et le contenu d'un bloc n'existe qu'une
  // fois le bloc ouvert (chargé à ce moment-là). Le script ouvre donc les blocs utiles, attend leur
  // contenu, lit, puis referme ceux qu'il a ouverts.
  const visible = el => !!(el && el.getClientRects().length);
  const corpsBloc = w => w.querySelector(':scope > .fragment-body');
  // Élément à cliquer : le titre du bloc (le clic remonte jusqu'à la flèche et à l'en-tête), sinon la flèche
  const ouvreur = (w, second) => {
    const e = w.querySelector(':scope > .fragment-header');
    if (!e) return null;
    const titre = e.querySelector('.collapse-toggler h2 span') || e.querySelector('.collapse-toggler h2') || e.querySelector('h2');
    const fleche = e.querySelector('.collapse-toggler');
    return second ? fleche || e : titre || fleche || e;
  };
  const estOuvert = w => {
    const t = w.querySelector(':scope > .fragment-header .collapse-toggler');
    if (t && t.classList.contains('uncollapsed')) return true;
    if (t && t.classList.contains('collapsed')) return false;
    const c = corpsBloc(w);
    return visible(c) && c.offsetHeight > 2;
  };
  const enChargement = w => [...w.querySelectorAll('mat-spinner, mat-progress-spinner, mat-progress-bar, [role="progressbar"], .spinner-border, .spinner, .loader, .loading')].some(visible);
  const BLOCS = [
    { cle: 'dettes', id: 'tax-balance', nom: 'Consulter mes dettes actuelles et remboursements',
      pret: w => w.querySelector('#taxBalance table') },
    { cle: 'provisions', id: 'provisions', nom: 'Gérer mes provisions TVA',
      pret: w => w.querySelector('#available-amount') &&
        (w.querySelector('provision-list table') || /pas encore de provisions|aucune provision|nog geen provisies/i.test(w.textContent)) },
    { cle: 'paiements', id: 'payments', nom: 'Gérer mes paiements en ligne',
      pret: w => w.querySelector('app-payments-current table') },
  ];
  const bloc = Object.fromEntries(BLOCS.map(b => [b.cle, b]));
  const DELAI_BLOC = 20000;
  const preparerBloc = async b => {
    const debut = Date.now();
    let sig = '', stable = 0;
    b.clics = 0;
    for (;;) {
      const w = document.getElementById(b.id);
      if (w) {
        b.present = true;
        if (b.ouvertAvant === undefined) b.ouvertAvant = estOuvert(w);
        if (b.classeAvant === undefined) b.classeAvant = (w.querySelector(':scope > .fragment-header .collapse-toggler') || {}).className || '';
        if (b.pret(w)) { b.etat = 'lu'; break; }
        if (!estOuvert(w)) {
          // un clic sur le titre du bloc ; s'il n'a rien changé après 1,5 s, un second sur la flèche
          const fl = w.querySelector(':scope > .fragment-header .collapse-toggler');
          const classe = fl ? fl.className : '';
          if (!b.clics || (b.clics === 1 && Date.now() - b.clicA > 1500 && classe === b.classeClic && !corpsBloc(w))) {
            const o = ouvreur(w, b.clics === 1);
            if (o) { b.classeClic = classe; o.click(); b.clics++; b.clicA = Date.now(); b.ouvertParOutil = true; b.via = b.clics === 1 ? 'titre' : 'fleche'; }
          }
          sig = ''; stable = 0;
        } else {
          // ouvert, chargé et inchangé depuis 2 secondes sans le tableau attendu : bloc vide
          const c = corpsBloc(w);
          const s = c ? c.textContent.length + ':' + c.getElementsByTagName('*').length : '-';
          stable = s === sig && !enChargement(w) ? stable + 1 : 0;
          sig = s;
          if (stable >= 13 && Date.now() - debut > 3000) { b.etat = 'vide'; break; }
        }
      }
      if (Date.now() - debut > DELAI_BLOC) { b.etat = w ? (estOuvert(w) ? 'delai' : 'ferme') : 'absent'; break; }
      await pause(150);
    }
    b.ms = Date.now() - debut;
  };
  const refermerBlocs = () => BLOCS.forEach(b => {
    if (!b.ouvertParOutil) return;
    const w = document.getElementById(b.id);
    const o = w && estOuvert(w) && ouvreur(w, b.via === 'fleche');
    if (o) o.click();
  });
  // Texte du bloc sans les messages d'information de MyMinfin (masqué : chiffres remplacés, pour le diagnostic)
  const texteBloc = (w, masquer) => {
    if (!w) return '';
    const c = (corpsBloc(w) || w).cloneNode(true);
    c.querySelectorAll('message-displayer, app-message-link, script, style').forEach(e => e.remove());
    const t = clean(c.textContent);
    return masquer ? t.replace(/\d/g, '9').slice(0, 240) : t;
  };
  const MSG_BLOC = {
    absent: b => `Bloc « ${b.nom} » absent de la page.`,
    ferme: b => `Bloc « ${b.nom} » : il ne s’est pas ouvert. L’ouvrir à la main, puis recliquer sur le favori.`,
    delai: b => `Bloc « ${b.nom} » : contenu pas chargé après ${DELAI_BLOC / 1000} s (MyMinfin lent ?). Recliquer sur le favori.`,
    vide: b => `Bloc « ${b.nom} » : rien à lire dans le bloc.`,
  };
  const msgBloc = b => (MSG_BLOC[b.etat] || MSG_BLOC.absent)(b);
  {
    const fermes = BLOCS.filter(b => { const w = document.getElementById(b.id); return w && !estOuvert(w) && !b.pret(w); });
    const bl = fermes.length ? bulle('Moore - MyMinfin : ouverture des blocs de « Mes paiements »…') : null;
    try { await Promise.all(BLOCS.map(preparerBloc)); } finally { if (bl) bl.remove(); }
  }

  // ------------------------------------------------- compte-provisions TVA
  const dispoEl = document.querySelector('#available-amount');
  const dispo = dispoEl ? montant(txt(dispoEl)) : null;
  const provTable = document.querySelector('provision-list table') || document.querySelector('#provisions table');
  const provBrut = lireMatTable(provTable);
  const provisions = (provBrut ? provBrut.lignes : []).map(l => ({
    type: l.natureType || '',
    reference: l.paymentReference || '',
    dateValeur: dateFR(l.valueDate || ''),
    montant: montant(l.availableAmount),
  }));
  if (!dispoEl) avert.push(bloc.provisions.etat === 'vide'
    ? 'Aucun montant disponible affiché dans « Gérer mes provisions TVA » (pas de compte-provisions TVA pour ce client ?).'
    : msgBloc(bloc.provisions));
  const totalProv = provisions.reduce((s, p) => s + (p.montant || 0), 0);
  if (dispo != null && provisions.length && Math.abs(totalProv - dispo) > 0.005) {
    avert.push(`Le détail des provisions (${affiche(totalProv)}) ne correspond pas au montant disponible (${affiche(dispo)}).`);
  }

  // -------------------------------------- dettes actuelles et remboursements
  // Le tableau mélange dettes et remboursements (crédits TVA à rembourser) sans les distinguer à l'écran.
  // Pour les séparer, le script utilise un instant le filtre « Montant initial » (Remboursements / Dettes)
  // de MyMinfin, puis le remet sur « Tout ». Rien d'autre n'est touché.
  // Ligne TVA : « TVA (déclaration périodique) », « Taxe sur la valeur ajoutée », « BTW … », réf. VALUEADDEDTAX…
  const RE_TVA = /\btva\b|\bbtw\b|valeur ajout|toegevoegde waarde|value ?added|mehrwertsteuer|\bmwst\b/i;
  const tb = document.querySelector('#taxBalance');
  const tbTable = tb && tb.querySelector('table');
  const dettes = [];
  let majSPF = '';
  let nature = null;
  let dettesLues = !!tbTable;   // faux : tableau non lu (les totaux ne veulent rien dire)
  let infoDettes = '';
  if (tbTable) {
    const ths = [...tbTable.querySelectorAll('thead th')];
    const labels = ths.map(th => simple(txt(th.querySelector(':scope > div')) || txt(th)));
    const col = (re, def) => { const i = labels.findIndex(l => re.test(l)); return i < 0 ? def : i; };
    const iRef = col(/^r.{1,2}f/, 0), iType = col(/^type/, 1), iEch = col(/ch.{1,2}ance/, 2),
          iInit = col(/initial/, 3), iSolde = col(/^solde/, 4);
    const lignesBrutes = () => [...tbTable.querySelectorAll('tbody tr')]
      .map(tr => [...tr.querySelectorAll(':scope > td')].map(txt)).filter(c => c.length >= 5);
    const cle = c => [c[iRef], c[iType], c[iEch], c[iInit]].join('|');

    const selects = [...tbTable.querySelectorAll('thead select')];
    const filtres = selects.some(s => s.selectedIndex > 0) || [...tbTable.querySelectorAll('thead input')].some(i => i.value);
    if (filtres) avert.push('Un filtre est actif sur le tableau des dettes : certaines lignes peuvent être masquées.');

    const selNature = selects.find(s => [...s.options].some(o => /rembours/.test(simple(o.text))) &&
                                        [...s.options].some(o => /^dettes?$/.test(simple(o.text))));
    if (selNature && !filtres && lignesBrutes().length) {
      const opt = re => [...selNature.options].findIndex(o => re.test(simple(o.text)));
      const iR = opt(/rembours/), iD = opt(/^dettes?$/), iTout = selNature.selectedIndex;
      const signature = () => lignesBrutes().map(cle).join('\n');
      const choisir = i => { selNature.selectedIndex = i; selNature.dispatchEvent(new Event('change', { bubbles: true })); };
      const depart = signature();
      let change = false;   // le filtre a-t-il réagi au moins une fois ?
      const appliquer = async (i, delai) => {
        const avant = signature();
        choisir(i);
        if (await attendre(() => signature() !== avant, delai)) { change = true; await pause(200); }
        return new Set(lignesBrutes().map(cle));
      };
      const bl = bulle('Moore - MyMinfin : lecture des dettes et remboursements…');
      try {
        const R = await appliquer(iR, 2500);
        const D = await appliquer(iD, 2500);
        await appliquer(iTout, change && signature() !== depart ? 2500 : 0);
        const toutes = lignesBrutes().map(cle);
        if (toutes.length && toutes.every(k => R.has(k) !== D.has(k))) {
          nature = new Map(toutes.map(k => [k, R.has(k) ? 'remboursement' : 'dette']));
        }
      } finally {
        if (selNature.selectedIndex !== iTout) { choisir(iTout); await pause(500); }
        bl.remove();
      }
    }

    lignesBrutes().forEach(c => {
      const solde = montant(c[iSolde]);
      const initial = montant(c[iInit]);
      dettes.push({
        reference: c[iRef], type: c[iType], echeance: dateFR(c[iEch]), initial, solde,
        reste: solde != null ? solde : initial, tva: RE_TVA.test(c[iType]) || /valueaddedtax/i.test(c[iRef]),
        nature: nature ? nature.get(cle(c)) || '' : '',
      });
    });
    const cpt = simple(txt(tb)).match(/remboursements?\s*:\s*(\d+)/);
    if (cpt && +cpt[1] !== dettes.length) avert.push(`MyMinfin annonce ${cpt[1]} dette(s)/remboursement(s), ${dettes.length} ligne(s) lue(s).`);
    const maj = simple(txt(tb)).match(/jour des donn.{1,2}es\s*:\s*(\d{1,2}(?:er)?\s+\S+\s+\d{4})/);
    if (maj) majSPF = dateFR(maj[1]);
  } else {
    // Pas de tableau : soit MyMinfin dit qu'il n'y a rien (« … : 0 », « aucune dette… »), soit le bloc n'a pas été lu
    const t = simple(texteBloc(document.getElementById(bloc.dettes.id)));
    if (bloc.dettes.etat === 'vide' && (/(dettes?|remboursements?) ?: ?0(?!\d)/.test(t) ||
        /aucun(e|s)? (dette|rembours)|pas de dettes?\b|n.avez (actuellement )?(aucune?|pas de)|geen (schulden|terugbetaling)/.test(t))) {
      dettesLues = true;
      infoDettes = 'MyMinfin n’affiche aucune dette ni aucun remboursement.';
    } else {
      avert.push(msgBloc(bloc.dettes) + ' Dettes et remboursements non lus.');
    }
  }
  const aujourdhui = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}`;
  const somme = l => l.reduce((s, d) => s + (d.reste || 0), 0);
  const lignesTVA = dettes.filter(d => d.tva);
  // Nature inconnue (filtre inutilisable) : traitée prudemment comme une dette, avec un point d'attention
  const dettesTVA = lignesTVA.filter(d => d.nature !== 'remboursement');
  const rembTVA = lignesTVA.filter(d => d.nature === 'remboursement');
  const rembAVenir = rembTVA.filter(d => triDate(d.echeance) >= aujourdhui);
  const rembPasses = rembTVA.filter(d => triDate(d.echeance) < aujourdhui);
  const natureInconnue = lignesTVA.some(d => !d.nature);
  if (natureInconnue) avert.push('Dettes et remboursements non distingués : toutes les lignes TVA sont comptées comme dettes. Vérifier avec le filtre « Montant initial » du tableau.');
  const totalTVA = somme(dettesTVA);
  const soldeVide = dettesTVA.filter(d => d.solde == null).length;
  if (soldeVide) avert.push(`Colonne « Solde » vide sur ${soldeVide} dette(s) TVA : montant initial repris (à vérifier via « Détail »).`);
  const plusAncienne = dettesTVA.map(d => d.echeance).sort((a, b) => triDate(a).localeCompare(triDate(b)))[0] || '';
  const NATURE = { dette: 'Dette', remboursement: 'Remboursement', '': '?' };
  const autres = dettes.filter(d => !d.tva);
  const autresDettes = autres.filter(d => d.nature !== 'remboursement');
  const autresRemb = autres.filter(d => d.nature === 'remboursement');
  const rembSansDate = rembTVA.filter(d => triDate(d.echeance) === '99999999');
  // Dettes échues (échéance passée, nature confirmée par le filtre) : à régulariser
  const echues = l => l.filter(d => d.nature === 'dette' && triDate(d.echeance) !== '99999999' && triDate(d.echeance) < aujourdhui);
  const tvaEchues = echues(dettesTVA), autresEchues = echues(autresDettes);
  if (tvaEchues.length || autresEchues.length) {
    const ancienne = tvaEchues.map(d => d.echeance).sort((a, b) => triDate(a).localeCompare(triDate(b)))[0];
    avert.push('Dettes échues non soldées : ' + [
      tvaEchues.length && `TVA ${affiche(somme(tvaEchues))} (${tvaEchues.length}, la plus ancienne au ${ancienne})`,
      autresEchues.length && `hors TVA ${affiche(somme(autresEchues))} (${autresEchues.length})`,
    ].filter(Boolean).join(' ; ') + ' : des intérêts de retard courent.');
  }
  // Provisions disponibles et dettes TVA ouvertes en même temps : paiement peut-être pas encore affecté à sa dette
  if (dispo > 0.005 && totalTVA > 0.005 && !natureInconnue) {
    avert.push(`Provisions disponibles (${affiche(dispo)}) alors que des dettes TVA sont ouvertes (${affiche(totalTVA)}) : ` +
      'paiements peut-être pas encore affectés aux dettes (retards d’affectation signalés par MyMinfin). À vérifier.');
    provisions.filter(pv => pv.montant > 0.005).forEach(pv => {
      const d = dettesTVA.find(x => x.initial != null && Math.abs(x.initial - pv.montant) < 0.005);
      if (d) avert.push(`${pv.type || 'Provision'} du ${pv.dateValeur} (${affiche(pv.montant)}) = montant initial de la dette TVA ` +
        `${d.echeance ? 'à l’échéance du ' + d.echeance : 'sans échéance'} (réf. ${d.reference}) : probablement le paiement de cette dette, resté sur le compte-provisions.`);
    });
  }
  // Blocs lus même sans le tableau attendu (pas de provisions, aucune dette)
  if (dispoEl && bloc.provisions.etat === 'vide') bloc.provisions.etat = 'lu';
  if (dettesLues && bloc.dettes.etat === 'vide') bloc.dettes.etat = 'lu';

  // ---------------------------------------------------- paiements en attente
  const attenteTable = document.querySelector('app-payments-current table') || document.querySelector('#payments table');
  const attente = lireMatTable(attenteTable);
  const nbAttente = attente ? attente.lignes.length : null;
  if (attenteTable && bloc.paiements.etat === 'vide') bloc.paiements.etat = 'lu';
  if (!attenteTable && bloc.paiements.present) avert.push(msgBloc(bloc.paiements) + ' Paiements en attente non lus.');

  // État des blocs pour le diagnostic (relevé avant de refermer les blocs ouverts par le script)
  const etatBlocs = BLOCS.map(b => {
    const w = document.getElementById(b.id);
    const o = { bloc: b.id, present: !!b.present, ouvertAvant: !!b.ouvertAvant, ouvertParLeFavori: !!b.ouvertParOutil,
                clics: b.clics || 0, etat: b.etat || '?', ms: b.ms, classeFleche: b.classeAvant || '' };
    if (w && b.etat !== 'lu') {
      const c = corpsBloc(w);
      o.balises = c ? [...new Set([...c.querySelectorAll('*')].map(e => e.tagName.toLowerCase()).filter(t => t.includes('-')))].slice(0, 20) : [];
      o.ids = c ? [...c.querySelectorAll('[id]')].map(e => e.id).slice(0, 20) : [];
      o.tableaux = c ? c.querySelectorAll('table').length : 0;
      o.texteChiffresMasques = texteBloc(w, true);
    }
    return o;
  });
  refermerBlocs();

  // ---------------------------------------------------------------- export
  const reconnu = !!(dispoEl || tbTable || bloc.dettes.present || bloc.provisions.present);
  const rienLu = !dispoEl && !dettesLues;
  const nl = v => (dettesLues ? xlNombre(v) : 'non lu');
  const L = [];
  L.push(`MyMinfin – compte-provisions, dettes et remboursements TVA\t(Moore - Outils ${VERSION})`);
  L.push(`Client\t${xlTexte(client || '(non sélectionné)')}`);
  L.push(`N° entreprise\t${xlTexte(bce)}`);
  L.push(`Consulté le\t${consulte}`);
  L.push(`Données SPF à jour au\t${majSPF}`);
  L.push('');
  L.push(`Montant disponible sur le compte-provisions\t${xlNombre(dispo)}`);
  L.push('');
  L.push('Détail des provisions');
  L.push(['Type', 'Référence', 'Date valeur', 'Montant disponible'].join('\t'));
  provisions.forEach(p => L.push([xlTexte(p.type), xlTexte(p.reference), p.dateValeur, xlNombre(p.montant)].join('\t')));
  if (!provisions.length) L.push('(aucune ligne)');
  L.push('');
  L.push('Dettes actuelles et remboursements');
  L.push(['Référence', 'Type', 'Nature', "Date d'échéance", 'Montant initial', 'Solde', 'Montant retenu', 'TVA ?'].join('\t'));
  dettes.forEach(d => L.push([xlTexte(d.reference), xlTexte(d.type), NATURE[d.nature], d.echeance, xlNombre(d.initial),
                              xlNombre(d.solde), xlNombre(d.reste), d.tva ? 'oui' : 'non'].join('\t')));
  if (!dettes.length) L.push(dettesLues ? (infoDettes || '(aucune ligne)') : '(tableau non lu)');
  L.push(`Total dettes TVA à payer${natureInconnue ? ' (nature non déterminée)' : ''}\t${nl(totalTVA)}`);
  L.push(`Total remboursements TVA à venir\t${nl(somme(rembAVenir))}`);
  L.push(`Total remboursements TVA – échéance passée\t${nl(somme(rembPasses))}`);
  if (autresDettes.length) L.push(`Total autres dettes (hors TVA)\t${nl(somme(autresDettes))}`);
  if (autresRemb.length) L.push(`Total autres remboursements (hors TVA)\t${nl(somme(autresRemb))}`);
  L.push('');
  if (nbAttente != null) {
    L.push(`Paiements en attente\t${nbAttente || 'aucun'}`);
    if (nbAttente) {
      L.push(attente.entetes.join('\t'));
      attente.lignes.forEach(l => L.push(Object.values(l).map(xlTexte).join('\t')));
    }
  }
  if (avert.length) { L.push(''); L.push('Points d’attention'); avert.forEach(a => L.push(a)); }
  const detail = L.join('\n');

  const entete = ['Consulté le', 'Client', 'N° entreprise', 'Disponible compte-provisions',
                  'Dettes TVA à payer', 'Nb dettes TVA', 'Plus ancienne échéance dette TVA',
                  'Remboursements TVA à venir', 'Remboursements TVA échéance passée',
                  'Paiements en attente', 'Données SPF au', 'Points d’attention'].join('\t');
  const ligne = [consulte, xlTexte(client), xlTexte(bce), xlNombre(dispo), nl(totalTVA), dettesLues ? dettesTVA.length : '',
                 plusAncienne, nl(somme(rembAVenir)), nl(somme(rembPasses)),
                 nbAttente ?? '', majSPF, avert.length].join('\t');

  // Diagnostic sans montants ni chiffres
  const diag = JSON.stringify({
    script: VERSION, page: location.pathname,
    trouve: { client: !!ident, montantDisponible: !!dispoEl, detailProvisions: !!provTable, dettes: !!tbTable, dettesLues, paiements: !!attenteTable },
    blocsLus: etatBlocs,
    blocs: [...document.querySelectorAll('.fragment-wrapper[id]')].map(e => e.id),
    titres: [...document.querySelectorAll('h2')].map(txt).slice(0, 15),
  }, null, 2);

  window.__mmf = { detail, ligne, entete, diag, donnees: { client, bce, dispo, provisions, dettes, dettesLues, totalTVA, rembAVenir: somme(rembAVenir), rembPasses: somme(rembPasses), majSPF, avert, blocs: etatBlocs } };

  // ------------------------------------------------------- Excel du détail
  const isoDe = fr => { const m = String(fr || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? `${m[3]}-${m[2]}-${m[1]}` : ''; };
  const nomFichier = n => String(n).replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-').replace(/\s+/g, ' ').trim().slice(0, 150);
  const jourISO = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
  const construireExcel = async () => {
    const { S: st, K, DXF, header, kpi, secTitle, thRow, mset, statusCF } = MS;
    const nomClient = client || 'Client MyMinfin';
    const bceNum = bce.replace(/\./g, '');
    const sep = '   ·   ';
    const outil = 'Compte-provisions MyMinfin ' + VERSION;
    const wb = new XL.Workbook({ title: 'Compte-provisions TVA — ' + nomClient, creator: 'Moore Walcourt — ' + outil });
    // Marqué pour Excel comme les fichiers du favori Moore - Odoo (propriétés MW_…)
    Object.entries({
      MW_Outil: 'COMPTE-PROVISIONS-TVA', MW_Version: VERSION.replace(/^v/, ''), MW_Societe: nomClient,
      MW_TVA: bceNum ? 'BE' + bceNum : '', MW_NumEntreprise: bceNum, MW_Source: 'MYMINFIN',
      MW_Consulte: jourISO, MW_DonneesSPF: isoDe(majSPF),
    }).forEach(([k, v]) => wb.prop(k, v));
    const F = wb.sheet('Compte-provisions');
    const X = wb.sheet('Extraction');
    const sousTitre = nomClient + (bce ? sep + 'BCE ' + bce : '') + (majSPF ? sep + 'données SPF au ' + majSPF : '');
    const pied = 'Moore Walcourt  ·  ' + nomClient;
    const cel = (r, c, v, style) => F.set(r, c, v === '' || v === undefined ? null : v, style || st.td);
    const ligneVide = r => { for (let c = 1; c <= 8; c++) F.set(r, c, null, st.td); };
    const date = v => (isoDe(v) ? { d: isoDe(v) } : v);
    const nonLu = Object.assign({}, st.tdGris, { h: 'right' });
    const ETAT_FR = { lu: 'lu', vide: 'rien à lire', delai: 'pas chargé', ferme: 'pas ouvert', absent: 'absent' };

    F.opts.zoom = 90;
    [19, 29, 18, 12, 14, 13, 17, 18].forEach((w, i) => F.width(i + 1, w));
    header(F, { title: 'Compte-provisions TVA', subtitle: sousTitre, lastCol: 8, tab: K.BLEU,
      note: 'Extraction MyMinfin du ' + MS.dateTimeFr(now) + sep + 'lecture seule' + sep + outil });

    // Fiche du dossier (mêmes noms que les concordances : complétée depuis DIA par Moore - Outils)
    kpi(F, 6, [[1, 2, 'DOSSIER', nomClient, st.kpiIn], [3, 3, 'N° D’ENTREPRISE', bce, st.kpiT],
               [4, 4, 'CONSULTÉ LE', { d: jourISO }, Object.assign({}, st.kpiT, { fmt: MS.DATE })],
               [5, 6, 'PROFESSION', '', st.kpiIn], [7, 7, 'RELATION CLIENT', '', st.kpiIn], [8, 8, 'SENIOR', '', st.kpiIn]]);
    wb.defineName('MW_Dossier', F.name, '$A$7').defineName('MW_Profession', F.name, '$E$7')
      .defineName('MW_RelationClient', F.name, '$G$7').defineName('MW_Senior', F.name, '$H$7');
    F.height(7, 22); F.height(8, 6);
    const JOUR = '$D$7';   // date de consultation : sépare échéances passées et à venir

    // 1. Compte-provisions
    let r = 13;
    secTitle(F, r, 1, 8, 'Compte-provisions TVA'); r++;
    thRow(F, r, 1, ['Référence', 'Type', 'Nature', 'Date valeur', '', '', 'Montant disponible', ''], 2); r++;
    const p0 = r;
    if (!provisions.length) { ligneVide(r); mset(F, r, 1, 6, dispoEl ? 'Aucune provision.' : 'Compte-provisions non lu (voir les points d’attention).', st.tdGris); r++; }
    provisions.forEach(p => {
      cel(r, 1, p.reference === '-' ? '' : p.reference); cel(r, 2, p.type); cel(r, 3, 'Provision', Object.assign({}, st.tdGris, { h: 'center' }));
      cel(r, 4, date(p.dateValeur), st.date); cel(r, 5, null); cel(r, 6, null); cel(r, 7, p.montant, st.num); cel(r, 8, null);
      r++;
    });
    const p1 = r - 1;
    mset(F, r, 1, 6, 'Total des provisions', st.tot); F.set(r, 7, { f: `SUM(G${p0}:G${p1})` }, st.totNum); F.set(r, 8, null, st.tot);
    const rTotProv = r; r++;
    mset(F, r, 1, 6, 'Montant disponible affiché par MyMinfin', st.td);
    F.set(r, 7, dispo != null ? dispo : 'non lu', dispo != null ? st.num : nonLu); F.set(r, 8, null, st.td);
    const rDispo = r; r++;
    mset(F, r, 1, 6, 'Écart (total des provisions − montant disponible)', st.tdB);
    F.set(r, 7, { f: `ROUND(G${rTotProv}-G${rDispo},2)` }, st.numB);
    if (dispo != null) {
      F.set(r, 8, { f: `IF(ABS(G${r})<0.005,"Concordant","Écart à justifier")` }, st.status);
      statusCF(F, 'H' + r, 'H' + r, true);
    } else { F.set(r, 7, null, st.numB); F.set(r, 8, 'Non contrôlé', st.status); }
    r += 2;

    // 2. Dettes actuelles et remboursements
    secTitle(F, r, 1, 8, 'Dettes actuelles et remboursements'); r++;
    thRow(F, r, 1, ['Référence', 'Type', 'Nature', 'Échéance', 'Montant initial', 'Solde', 'Montant retenu', 'Statut'], 2); r++;
    const d0 = r;
    if (!dettes.length) {
      ligneVide(r);
      mset(F, r, 1, 6, dettesLues ? 'Aucune dette ni aucun remboursement.' : 'Tableau non lu (voir les points d’attention).', st.tdGris); r++;
    }
    // Nature : « Dette TVA », « Remboursement TVA », « Dette », « Remboursement » (« ? » si non distinguée).
    // Les totaux se basent sur cette colonne : la corriger à la main met les totaux à jour.
    const vc = s => Object.assign({}, s, { v: 'center' });
    const [tdV, tdcV, dateV, numV] = [st.td, st.tdc, st.date, st.num].map(vc);
    const typeV = Object.assign({}, st.td, { v: 'center', wrap: true });
    const refV = Object.assign({}, st.td, { v: 'center', shrink: true });   // longues références (VALUEADDEDTAXMERGED)
    dettes.forEach(d => {
      cel(r, 1, d.reference, refV); cel(r, 2, d.type, typeV); cel(r, 3, NATURE[d.nature] + (d.tva ? ' TVA' : ''), tdcV);
      cel(r, 4, date(d.echeance), dateV); cel(r, 5, d.initial, numV); cel(r, 6, d.solde, numV);
      F.set(r, 7, { f: `IF(F${r}="",E${r},F${r})` }, numV);
      F.set(r, 8, { f: `IF(LEFT(C${r},5)="Dette",IF(D${r}="","Dette sans échéance",IF(D${r}<${JOUR},"Dette échue","Dette à échoir")),` +
        `IF(LEFT(C${r},5)="Rembo",IF(D${r}="","Remb. sans date",IF(D${r}<${JOUR},"Remb. échu","Remb. à venir")),"À vérifier"))` }, tdcV);
      r++;
    });
    const d1 = r - 1;
    if (dettes.length) {
      F.cf(`H${d0}:H${d1}`, `H${d0}="Dette échue"`, DXF.koT);
      F.cf(`H${d0}:H${d1}`, `OR(H${d0}="Remb. échu",H${d0}="À vérifier",H${d0}="Dette sans échéance",H${d0}="Remb. sans date")`, DXF.warnT);
    }
    const plage = c => `$${c}$${d0}:$${c}$${d1}`;
    const [G, C, D] = ['G', 'C', 'D'].map(plage);
    const totaux = [
      ['Dettes TVA à payer', `SUMIFS(${G},${C},"Dette TVA")+SUMIFS(${G},${C},"~? TVA")`],
      ['Remboursements TVA à venir' + (rembSansDate.length ? ' (y compris sans date)' : ''),
        `SUMIFS(${G},${C},"Remboursement TVA",${D},">="&${JOUR})+SUMIFS(${G},${C},"Remboursement TVA",${D},"")`],
      ['Remboursements TVA – échéance passée (à rapprocher de la banque)', `SUMIFS(${G},${C},"Remboursement TVA",${D},"<"&${JOUR})`],
      ['Autres dettes (hors TVA)', `SUMIFS(${G},${C},"Dette")+SUMIFS(${G},${C},"~?")`],
      ['Autres remboursements (hors TVA)', `SUMIFS(${G},${C},"Remboursement")`],
    ];
    const rTot = [];
    totaux.forEach(([lib, f], k) => {
      mset(F, r, 1, 6, lib, k ? st.hi : st.tot);
      F.set(r, 7, dettesLues ? { f } : 'non lu', Object.assign({}, k ? st.hiNum : st.totNum, dettesLues ? {} : { h: 'right' }));
      F.set(r, 8, null, k ? st.hi : st.tot);
      rTot.push(r); r++;
    });
    r++;

    // 3. Paiements en attente
    secTitle(F, r, 1, 8, 'Paiements en attente'); r++;
    if (!attente || !attente.lignes.length) {
      ligneVide(r);
      mset(F, r, 1, 8, attente ? 'Aucun paiement en attente.' : 'Bloc « Gérer mes paiements en ligne » non lu (replié ou pas encore chargé).', st.tdGris);
      r++;
    } else {
      thRow(F, r, 1, ['Référence', 'Intitulé', 'Statut', 'Expiration', '', '', 'Montant', ''], 2); r++;
      attente.lignes.forEach(l => {
        const val = re => { const k = Object.keys(l).find(x => re.test(x)); return k ? l[k] : ''; };
        cel(r, 1, val(/reference/i)); cel(r, 2, val(/title|intitul/i)); cel(r, 3, val(/status|statut/i));
        cel(r, 4, date(dateFR(val(/expiration/i))), st.date); cel(r, 5, null); cel(r, 6, null);
        cel(r, 7, montant(val(/amount|montant/i)), st.num); cel(r, 8, null);
        r++;
      });
    }
    r++;

    // 4. Points d'attention
    secTitle(F, r, 1, 8, 'Points d’attention'); r++;
    (avert.length ? avert : ['Aucun point d’attention.']).forEach(a => {
      mset(F, r, 1, 8, a, avert.length ? st.tdWrap : st.tdGris); F.height(r, avert.length ? 26 : 16); r++;
    });
    r++;

    // 5. Remarques internes (cellules grises à compléter)
    secTitle(F, r, 1, 8, 'Remarques internes'); r++;
    for (let k = r; k <= r + 2; k++) { for (let c = 1; c <= 8; c++) F.set(k, c, null, st.inpWrap); F.height(k, 20); }
    F.merge(r, 1, r + 2, 8);
    wb.defineName('MW_Remarques', F.name, `$A$${r}`);
    const rFin = r + 2;

    // Chiffres clés (lignes 9-10) : formules vers les totaux ci-dessus
    kpi(F, 9, [
      [1, 2, 'DISPONIBLE · COMPTE-PROVISIONS', dispo != null ? { f: `G${rDispo}` } : 'non lu'],
      [3, 4, 'DETTES TVA À PAYER', dettesLues ? { f: `G${rTot[0]}` } : 'non lu'],
      [5, 6, 'REMBOURSEMENTS TVA À VENIR', dettesLues ? { f: `G${rTot[1]}` } : 'non lu'],
      [7, 8, 'POSITION NETTE', dispo != null && dettesLues ? { f: 'A10+E10-C10' } : 'non calculée'],
    ]);
    mset(F, 11, 1, 8, 'Position nette = disponible + remboursements TVA à venir − dettes TVA à payer (positive : en faveur du client).', st.noteL);
    F.height(11, 14); F.height(12, 6);
    // Noms des montants clés, pour les reprendre ailleurs (fiche TVA, tableau de suivi…)
    wb.defineName('MW_CP_Consulte', F.name, JOUR).defineName('MW_CP_Disponible', F.name, `$G$${rDispo}`)
      .defineName('MW_CP_DettesTVA', F.name, `$G$${rTot[0]}`).defineName('MW_CP_RembAVenir', F.name, `$G$${rTot[1]}`)
      .defineName('MW_CP_RembEchus', F.name, `$G$${rTot[2]}`);
    F.print({ area: `$A$1:$H$${rFin}`, footer: pied, mid: 'Compte-provisions TVA — ' + MS.dateFr(now) });

    // Onglet Extraction (traçabilité, comme les concordances)
    X.width(1, 30); X.width(2, 100);
    header(X, { title: 'Informations d’extraction', subtitle: sousTitre, lastCol: 2, tab: '', note: 'Traçabilité : source, date et méthode de lecture.' });
    thRow(X, 6, 1, ['Élément', 'Valeur'], 2);
    const infos = [
      ['Société', nomClient],
      ['N° d’entreprise', bce || ''],
      ['Source', 'MyMinfin — Mes paiements : « Gérer mes provisions TVA » et « Consulter mes dettes actuelles et remboursements »'],
      ['Date de consultation', MS.dateTimeFr(now)],
      ['Données SPF à jour au', majSPF || 'non indiqué par MyMinfin'],
      ['Dettes et remboursements', !dettesLues ? 'Non lus (voir les points d’attention)'
        : !dettes.length ? (infoDettes || 'Aucune dette ni aucun remboursement affiché par MyMinfin')
        : nature ? 'Distingués avec le filtre « Montant initial » de MyMinfin (Remboursements / Dettes), remis ensuite sur « Tout »'
        : 'Non distingués (le filtre « Montant initial » n’a pas répondu) : les lignes TVA sont comptées comme dettes'],
      ['Blocs MyMinfin', etatBlocs.map(e => `${BLOCS.find(b => b.id === e.bloc).nom} : ${ETAT_FR[e.etat] || e.etat}` +
        (e.ouvertParLeFavori ? ' (ouvert puis refermé par le favori)' : '')).join(' ; ')],
      ['Lecture', 'Lecture seule : aucune donnée n’est modifiée dans MyMinfin'],
      ['Outil', outil + ' — Moore Walcourt'],
      ['Dossier DIA', 'à compléter depuis DIA (Excel, onglet Moore - Outils)'],
    ];
    infos.forEach(([k, v], i) => { X.set(7 + i, 1, k, st.keyL); X.set(7 + i, 2, v, st.tdWrap); });
    wb.defineName('MW_DossierDIA', X.name, '$B$' + (7 + infos.length - 1));
    X.print({ landscape: true, footer: pied, mid: 'Compte-provisions TVA — extraction' });
    return wb.build();
  };
  const telechargerExcel = async () => {
    const octets = await construireExcel();
    const blob = new Blob([octets], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const nom = nomFichier(`${jourISO} - Compte-provisions TVA - ${client || 'MyMinfin'}`) + '.xlsx';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = nom; a.style.display = 'none';
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 10000);
    return nom;
  };
  window.__mmf.excel = construireExcel;

  // ------------------------------------------------------- PDF (comme la fiche TVA)
  // Page A4 propre, imprimée par le navigateur (« Enregistrer au format PDF ») : marge d'impression nulle
  // et marges de 12 mm dans la page, répétées sur chaque page ; ni adresse ni date ajoutées par le navigateur.
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const mt = n => (n == null ? '' : affiche(n).replace(' €', ''));
  const statutDe = d => {
    const t = triDate(d.echeance), sans = t === '99999999';
    if (d.nature === 'dette') return sans ? 'Dette sans échéance' : t < aujourdhui ? 'Dette échue' : 'Dette à échoir';
    if (d.nature === 'remboursement') return sans ? 'Remb. sans date' : t < aujourdhui ? 'Remb. échu' : 'Remb. à venir';
    return 'À vérifier';
  };
  const classeStatut = st => (st === 'Dette échue' ? 'ko' : /Remb\. échu|À vérifier|sans/.test(st) ? 'warn' : '');
  const construirePDF = () => {
    const nomClient = client || 'Client MyMinfin';
    const K = MS.K;
    const logo = 'data:image/png;base64,' + btoa(Array.from(MS.logo(), c => String.fromCharCode(c)).join(''));
    const etabli = MS.dateTimeFr(now);
    const ecart = dispo != null ? Math.round((totalProv - dispo) * 100) / 100 : null;
    const sommeAVenir = somme(rembAVenir);
    const nonLus = [dispo == null && 'compte-provisions', !dettesLues && 'dettes et remboursements'].filter(Boolean);
    const etat = nonLus.length ? ['ko', `✖ Lecture incomplète : ${nonLus.join(' et ')} non lu${nonLus.length > 1 ? 's' : ''} (voir les points d’attention).`]
      : avert.length ? ['warn', `⚠ ${avert.length} point${avert.length > 1 ? 's' : ''} d’attention, en bas de page` + (ecart === 0 ? ' · compte-provisions concordant' : '') + '.']
      : ['ok', '✔ Lecture complète · compte-provisions concordant · aucun point d’attention.'];
    const position = dispo != null && dettesLues ? dispo + sommeAVenir - totalTVA : null;
    const kpi = (lib, val, cls) => `<div class="kpi"><span>${esc(lib)}</span><b class="${cls || ''}">${esc(val)}</b></div>`;
    const kv = (lib, val, cls) => `<div class="kv"><span>${lib}</span><b class="${cls || ''}">${esc(val)}</b></div>`;
    const ligneVide = (n, t) => `<tr><td colspan="${n}" class="gris">${esc(t)}</td></tr>`;
    // 1. Compte-provisions
    const prov = provisions.map(pv => `<tr><td>${esc(pv.type)}</td><td>${esc(pv.reference === '-' ? '' : pv.reference)}</td>` +
      `<td>${esc(pv.dateValeur)}</td><td class="r">${mt(pv.montant)}</td></tr>`).join('') ||
      ligneVide(4, dispoEl ? 'Aucune provision.' : 'Compte-provisions non lu.');
    const pastille = dispo == null ? '<span class="pill">Non contrôlé</span>'
      : ecart === 0 ? '<span class="pill ok">Concordant</span>' : '<span class="pill ko">Écart à justifier</span>';
    // 2. Synthèse
    const n = l => `<span class="gris">(${l.length})</span>`;
    const synthese = (!dettesLues ? kv('Dettes et remboursements', 'non lus', 'ko') : [
      kv(`Dettes TVA à payer ${n(dettesTVA)}`, affiche(totalTVA), totalTVA > 0.005 ? 'ko' : ''),
      plusAncienne ? kv('Plus ancienne échéance (dette TVA)', plusAncienne) : '',
      kv(`Remboursements TVA à venir ${n(rembAVenir)}${rembSansDate.length ? ` <span class="gris">dont ${rembSansDate.length} sans date</span>` : ''}`, affiche(sommeAVenir)),
      kv(`Remboursements TVA – échéance passée ${n(rembPasses)}<br><span class="gris">à rapprocher de la banque</span>`, affiche(somme(rembPasses))),
      autresDettes.length ? kv(`Autres dettes, hors TVA ${n(autresDettes)}`, affiche(somme(autresDettes))) : '',
      autresRemb.length ? kv(`Autres remboursements, hors TVA ${n(autresRemb)}`, affiche(somme(autresRemb))) : '',
    ].join('')) + kv('Paiements en attente', nbAttente == null ? 'non lus' : nbAttente ? String(nbAttente) : 'aucun') +
      kv('Données SPF à jour au', majSPF || '—');
    // 3. Dettes et remboursements
    const lignesDettes = dettes.map(d => {
      const st = statutDe(d);
      return `<tr><td class="nw">${esc(d.reference)}</td><td>${esc(d.type)}</td><td class="nw">${esc(NATURE[d.nature] + (d.tva ? ' TVA' : ''))}</td>` +
        `<td class="nw">${esc(d.echeance)}</td><td class="r">${mt(d.initial)}</td><td class="r">${mt(d.solde)}</td><td class="r">${mt(d.reste)}</td>` +
        `<td class="c nw ${classeStatut(st)}">${esc(st)}</td></tr>`;
    }).join('') || ligneVide(8, dettesLues ? (infoDettes || 'Aucune dette ni aucun remboursement.') : 'Tableau non lu (voir les points d’attention).');
    // 4. Paiements en attente
    let paiements;
    if (!attente) paiements = '<p class="gris">Bloc « Gérer mes paiements en ligne » non lu.</p>';
    else if (!attente.lignes.length) paiements = '<p class="gris">Aucun paiement en attente.</p>';
    else {
      const val = (l, re) => { const k = Object.keys(l).find(x => re.test(x)); return k ? l[k] : ''; };
      paiements = '<table class="t"><thead><tr><th>Référence</th><th>Intitulé</th><th>Statut</th><th>Expiration</th><th class="r">Montant</th></tr></thead><tbody>' +
        attente.lignes.map(l => `<tr><td>${esc(val(l, /reference/i))}</td><td>${esc(val(l, /title|intitul/i))}</td><td>${esc(val(l, /status|statut/i))}</td>` +
          `<td>${esc(dateFR(val(l, /expiration/i)))}</td><td class="r">${mt(montant(val(l, /amount|montant/i)))}</td></tr>`).join('') + '</tbody></table>';
    }
    const titre = nomFichier(`${jourISO} - Compte-provisions TVA - ${client || 'MyMinfin'}`);
    const css = `
      @page{size:A4 portrait;margin:0}
      *{box-sizing:border-box}
      html,body{margin:0;padding:0;background:#fff}
      body{font:8.5pt/1.35 Arial,Helvetica,sans-serif;color:#${K.NOIR};-webkit-print-color-adjust:exact;print-color-adjust:exact}
      table.page{width:100%;border-collapse:collapse}
      table.page>thead>tr>td,table.page>tfoot>tr>td,table.page>tbody>tr>td{padding:0 12mm}
      .m-haut{height:12mm}
      .pied{height:12mm;padding-top:2.5mm;border-top:.5pt solid #${K.FILET};display:flex;justify-content:space-between;gap:8px;font-size:7pt;color:#${K.GRIS}}
      .tete{display:flex;justify-content:space-between;align-items:flex-start}
      .logo{height:30px;display:block;margin:0 0 6px}
      h1{font-size:16pt;margin:0;line-height:1.15}
      .sous{color:#${K.BLEU};font-weight:bold;font-size:10pt;margin-top:2px}
      .marque{text-align:right;font-size:7.5pt;color:#${K.GRIS};line-height:1.4}
      .regle{border-bottom:1.5pt solid #${K.BLEU};font-size:7.5pt;color:#${K.GRIS};padding:5px 0 3px;margin-bottom:8px}
      .etat{border-left:4px solid;padding:6px 10px;margin:0 0 8px;font-weight:bold}
      .etat.ok{border-color:#${K.VERT};background:#${K.VERT_BG};color:#${K.VERT}}
      .etat.warn{border-color:#${K.AMBRE};background:#${K.AMBRE_BG};color:#${K.AMBRE}}
      .etat.ko{border-color:#${K.ROUGE};background:#${K.ROSE};color:#${K.ROUGE}}
      .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin:0 0 4px}
      .kpi{background:#${K.KPI};padding:6px 8px}
      .kpi span{display:block;font-size:6.5pt;font-weight:bold;color:#${K.GRIS};letter-spacing:.03em}
      .kpi b{display:block;font-size:12.5pt;margin-top:2px}
      .note{font-size:7pt;color:#${K.GRIS};margin:0 0 6px}
      h2{font-size:10pt;margin:10px 0 5px;padding-bottom:2px;border-bottom:1pt solid #${K.BLEU}}
      h2 .num{color:#${K.BLEU};margin-right:5px}
      .deux{display:grid;grid-template-columns:1fr 1fr;gap:14px}
      table.t{width:100%;border-collapse:collapse;font-size:8pt}
      table.t th{background:#${K.NOIR};color:#fff;text-align:left;padding:3px 5px;font-weight:bold}
      table.t td{padding:2.5px 5px;border-bottom:.5pt solid #${K.FILET};vertical-align:top}
      table.t .r{text-align:right;white-space:nowrap}
      table.t .c{text-align:center}
      table.t .nw,table.t th{white-space:nowrap}
      table.t tr.tot td{font-weight:bold;background:#${K.GRIS_FOND};border-top:.75pt solid #${K.NOIR};border-bottom:1.2pt solid #${K.BLEU}}
      tr,.kv,.kpi{break-inside:avoid}
      .kv{display:flex;justify-content:space-between;align-items:baseline;gap:8px;padding:3px 0;border-bottom:.5pt solid #${K.FILET}}
      .kv b{white-space:nowrap}
      .gris{color:#${K.GRIS};font-weight:normal}
      .ko{color:#${K.ROUGE};font-weight:bold}
      .warn{color:#${K.AMBRE};font-weight:bold}
      .ok{color:#${K.VERT};font-weight:bold}
      .pill{display:inline-block;padding:1px 6px;font-weight:bold;background:#${K.GRIS_FOND}}
      .pill.ok{background:#${K.VERT_BG}}
      .pill.ko{background:#${K.ROSE}}
      .alertes{background:#${K.AMBRE_BG};border-left:4px solid #${K.AMBRE};padding:6px 10px 6px 24px;margin:0;color:#5c3b06}
      .alertes li{margin:2px 0}
      p{margin:2px 0}`;
    let num = 0;
    const h2 = t => `<h2><span class="num">${++num}.</span>${esc(t)}</h2>`;
    return `<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><title>${esc(titre)}</title><style>${css}</style></head><body>
<table class="page"><thead><tr><td><div class="m-haut"></div></td></tr></thead>
<tfoot><tr><td><div class="pied"><span>Moore Walcourt · ${esc(nomClient)}</span><span>Compte-provisions TVA — ${esc(MS.dateFr(now))}</span><span>MyMinfin, lecture seule · ${esc(VERSION)}</span></div></td></tr></tfoot>
<tbody><tr><td>
<div class="tete"><div><img class="logo" src="${logo}" alt="MOORE"><h1>Compte-provisions TVA</h1>
<div class="sous">${esc(nomClient)}${bce ? ' · BCE ' + esc(bce) : ''}</div></div>
<div class="marque">Moore Walcourt<br>Établi le ${esc(etabli)}</div></div>
<div class="regle">Extraction MyMinfin (Mes paiements) du ${esc(etabli)}${majSPF ? ' · données SPF à jour au ' + esc(majSPF) : ''} · lecture seule</div>
<div class="etat ${etat[0]}">${esc(etat[1])}</div>
<div class="kpis">${kpi('DISPONIBLE · COMPTE-PROVISIONS', dispo != null ? affiche(dispo) : 'non lu')}${kpi('DETTES TVA À PAYER', dettesLues ? affiche(totalTVA) : 'non lu', dettesLues && totalTVA > 0.005 ? 'ko' : '')}${kpi('REMBOURSEMENTS TVA À VENIR', dettesLues ? affiche(sommeAVenir) : 'non lu')}${kpi('POSITION NETTE', position != null ? affiche(position) : 'non calculée', position == null ? '' : position < -0.005 ? 'ko' : 'ok')}</div>
<p class="note">Position nette = disponible + remboursements TVA à venir − dettes TVA à payer (positive : en faveur du client).</p>
<div class="deux"><div>${h2('Compte-provisions TVA')}
<table class="t"><thead><tr><th>Type</th><th>Référence</th><th>Date valeur</th><th class="r">Disponible</th></tr></thead><tbody>${prov}
<tr class="tot"><td colspan="3">Total des provisions</td><td class="r">${mt(totalProv)}</td></tr>
<tr><td colspan="3">Montant disponible affiché par MyMinfin</td><td class="r">${dispo != null ? mt(dispo) : 'non lu'}</td></tr>
<tr><td colspan="2">Écart</td><td>${pastille}</td><td class="r">${ecart != null ? mt(ecart) : ''}</td></tr></tbody></table></div>
<div>${h2('Synthèse')}${synthese}</div></div>
${h2('Dettes actuelles et remboursements')}
<table class="t"><thead><tr><th>Référence</th><th>Type</th><th>Nature</th><th>Échéance</th><th class="r">Montant initial</th><th class="r">Solde</th><th class="r">Montant retenu</th><th class="c">Statut</th></tr></thead>
<tbody>${lignesDettes}</tbody></table>
${h2('Paiements en attente')}${paiements}
${avert.length ? h2('Points d’attention') + '<ul class="alertes">' + avert.map(a => `<li>${esc(a)}</li>`).join('') + '</ul>' : ''}
</td></tr></tbody></table></body></html>`;
  };
  // Impression dans un cadre invisible ; le nom proposé pour le PDF est le titre de l'onglet, changé le temps de l'impression
  const imprimerPDF = async () => {
    const html = construirePDF();
    const titre = nomFichier(`${jourISO} - Compte-provisions TVA - ${client || 'MyMinfin'}`);
    document.getElementById('moore-mmf-pdf')?.remove();
    const f = document.createElement('iframe');
    f.id = 'moore-mmf-pdf';
    f.setAttribute('aria-hidden', 'true');
    f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
    document.body.appendChild(f);
    const d = f.contentDocument;
    d.open(); d.write(html); d.close();
    const img = d.images[0];
    if (img && !img.complete) await new Promise(r => { img.onload = img.onerror = r; setTimeout(r, 1500); });
    const ancien = document.title;
    let remis = false;
    const remettre = () => { if (!remis) { remis = true; document.title = ancien; } };
    document.title = titre;
    f.contentWindow.addEventListener('afterprint', remettre);
    f.contentWindow.focus();
    (window.__mmfImprimer || (w => w.print()))(f.contentWindow);
    setTimeout(remettre, 60000);
    return titre + '.pdf';
  };
  window.__mmf.pdf = construirePDF;

  // ----------------------------------------------------------------- copie
  const copier = async texte => {
    try { await navigator.clipboard.writeText(texte); return true; } catch (e) { /* repli ci-dessous */ }
    const ta = document.createElement('textarea');
    ta.value = texte; ta.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.appendChild(ta); ta.select();
    let ok = false; try { ok = document.execCommand('copy'); } catch (e) { /* ignore */ }
    ta.remove(); return ok;
  };
  try { if (typeof copy === 'function') copy(reconnu && !rienLu ? detail : diag); } catch (e) { /* hors console */ }

  // --------------------------------------------------------------- panneau
  document.getElementById(HOST_ID)?.remove();
  const host = document.createElement('div');
  host.id = HOST_ID;
  host.style.cssText = 'position:fixed;top:16px;right:16px;z-index:2147483647';
  const root = host.attachShadow({ mode: 'open' });
  const h = (tag, cls, texte) => { const e = document.createElement(tag); if (cls) e.className = cls; if (texte != null) e.textContent = texte; return e; };
  const style = h('style');
  style.textContent = `
    .card{width:370px;font:13px/1.4 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:#1d2433;background:#fff;
          border:1px solid #d5dbe5;border-radius:10px;box-shadow:0 8px 28px rgba(16,24,40,.18);overflow:hidden}
    .top{display:flex;align-items:center;justify-content:space-between;background:#16324f;color:#fff;padding:9px 12px}
    .top b{font-size:13px}.top span{opacity:.75;font-size:11px;margin-left:6px}
    .x{all:unset;cursor:pointer;font-size:18px;line-height:1;padding:0 4px;color:#fff}
    .body{padding:12px}
    .client{font-weight:600;font-size:14px}.bce{color:#5b6475;font-size:12px;margin-bottom:10px}
    .kv{display:flex;justify-content:space-between;gap:12px;padding:5px 0;border-top:1px solid #eef1f5}
    .kv span:first-child{color:#5b6475}.kv span:last-child{font-variant-numeric:tabular-nums;font-weight:600;text-align:right;white-space:nowrap}
    .warn{margin-top:10px;background:#fff7e6;border:1px solid #f3d08a;border-radius:6px;padding:8px 10px;color:#6b4a00;font-size:12px}
    .warn div+div{margin-top:4px}
    .err{background:#fdecec;border-color:#f2b8b5;color:#7a1c1c}
    .btns{display:flex;flex-direction:column;gap:6px;margin-top:12px}
    button.b{all:unset;box-sizing:border-box;cursor:pointer;text-align:center;padding:8px 10px;border-radius:6px;
             border:1px solid #16324f;color:#16324f;font-weight:600}
    button.b.p{background:#16324f;color:#fff}
    .duo{display:flex;gap:6px}.duo button.b{flex:1 1 0;min-width:0}
    button.b:hover{filter:brightness(1.08)}
    .note{margin-top:8px;color:#5b6475;font-size:11px}
    .note a.diag{color:#5b6475;text-decoration:underline;cursor:pointer}`;
  const card = h('div', 'card');
  const top = h('div', 'top');
  const t = h('div'); t.append(h('b', null, 'Moore - Outils'), h('span', null, `MyMinfin · compte-provisions ${VERSION}`));
  const x = h('button', 'x', '×'); x.title = 'Fermer'; x.onclick = () => host.remove();
  top.append(t, x);
  const body = h('div', 'body');
  body.append(h('div', 'client', client || '(aucun client sélectionné)'), h('div', 'bce', bce ? `BCE ${bce}` : ''));
  const kv = (k, v) => { const r = h('div', 'kv'); r.append(h('span', null, k), h('span', null, v)); body.append(r); };
  if (reconnu && !rienLu) {
    kv('Disponible sur le compte-provisions', dispo != null ? affiche(dispo) : 'non lu');
    if (!dettesLues) kv('Dettes et remboursements', 'non lus');
    else kv(natureInconnue ? `Montants TVA – nature non déterminée (${dettesTVA.length})` : `Dettes TVA à payer (${dettesTVA.length})`,
            dettesTVA.length ? affiche(totalTVA) : 'aucune');
    if (plusAncienne) kv('Plus ancienne échéance (dette TVA)', plusAncienne);
    if (rembAVenir.length) kv(`Remboursements TVA à venir (${rembAVenir.length}${rembSansDate.length ? `, dont ${rembSansDate.length} sans date` : ''})`, affiche(somme(rembAVenir)));
    if (rembPasses.length) kv(`Remboursements TVA – échéance passée (${rembPasses.length})`, affiche(somme(rembPasses)));
    if (autresDettes.length) kv(`Autres dettes, hors TVA (${autresDettes.length})`, affiche(somme(autresDettes)));
    if (autresRemb.length) kv(`Autres remboursements, hors TVA (${autresRemb.length})`, affiche(somme(autresRemb)));
    if (nbAttente != null) kv('Paiements en attente', nbAttente ? String(nbAttente) : 'aucun');
    kv('Données SPF à jour au', majSPF || '—');
  }
  if (avert.length || !reconnu || rienLu) {
    const w = h('div', reconnu && !rienLu ? 'warn' : 'warn err');
    (!reconnu ? ['Page non reconnue : ouvre MyMinfin > Mes paiements pour le client, puis relance le script.',
                 'Si le problème persiste, clique « Copier le diagnostic » et envoie-le à Claude (aucun montant dedans).']
      : rienLu ? avert.concat('Rien n’a pu être lu. Si le problème persiste, clique « Copier le diagnostic » et envoie-le à Claude (aucun montant dedans).')
      : avert).forEach(a => w.append(h('div', null, '• ' + a)));
    body.append(w);
  }
  const btns = h('div', 'btns');
  const bouton = (label, texte, principal) => {
    const b = h('button', principal ? 'b p' : 'b', label);
    b.onclick = async () => { const ok = await copier(texte); b.textContent = ok ? 'Copié ✓' : 'Copie impossible — voir console'; if (!ok) console.log(texte); setTimeout(() => (b.textContent = label), 1600); };
    btns.append(b);
  };
  if (reconnu && !rienLu) {
    const bx = h('button', 'b p', 'Télécharger l’Excel');
    bx.onclick = async () => {
      if (bx.dataset.occupe) return;
      bx.dataset.occupe = '1';
      bx.textContent = 'Création de l’Excel…';
      try { bx.title = await telechargerExcel(); bx.textContent = 'Excel téléchargé ✓'; }
      catch (e) { bx.textContent = 'Excel impossible : ' + e.message; console.error(e); }
      setTimeout(() => { bx.textContent = 'Télécharger l’Excel'; delete bx.dataset.occupe; }, 2500);
    };
    // « Enregistrer en PDF » : même page que la fiche TVA (impression du navigateur)
    const bp = h('button', 'b p', 'Enregistrer en PDF');
    bp.title = 'Choisir l’imprimante « Enregistrer au format PDF », puis le dossier du client';
    bp.onclick = async () => {
      if (bp.dataset.occupe) return;
      bp.dataset.occupe = '1';
      bp.textContent = 'Ouverture…';
      try { await imprimerPDF(); bp.textContent = 'Enregistrer en PDF'; delete bp.dataset.occupe; }
      catch (e) {
        bp.textContent = 'PDF impossible : ' + e.message; console.error(e);
        setTimeout(() => { bp.textContent = 'Enregistrer en PDF'; delete bp.dataset.occupe; }, 2500);
      }
    };
    const duo = h('div', 'duo');
    duo.append(bx, bp);
    btns.append(duo);
    bouton('Copier le détail', detail);
    bouton('Copier la ligne récap', ligne);
    bouton('Copier en-tête + ligne récap', entete + '\n' + ligne);
  } else {
    bouton('Copier le diagnostic', diag, true);
  }
  const note = h('div', 'note', 'Lecture seule : rien n’est modifié dans MyMinfin (blocs ouverts puis refermés, filtre du tableau remis sur « Tout »). ');
  if (reconnu && !rienLu) {
    const lien = h('a', 'diag', 'Copier le diagnostic');
    lien.href = '#';
    lien.onclick = async e => { e.preventDefault(); const ok = await copier(diag); lien.textContent = ok ? 'Diagnostic copié ✓' : 'Copie impossible — voir console'; if (!ok) console.log(diag); setTimeout(() => (lien.textContent = 'Copier le diagnostic'), 1600); };
    note.append(lien);
  }
  body.append(btns, note);
  card.append(top, body);
  root.append(style, card);
  document.body.appendChild(host);

  console.log(`%cMoore - Outils · MyMinfin compte-provisions ${VERSION}`, 'font-weight:bold');
  console.log(reconnu && !rienLu ? '✔ Détail copié dans le presse-papiers (Ctrl+V dans Excel).' : '✖ Rien de lisible — diagnostic copié.');
  if (avert.length) console.warn(avert.join('\n'));
}

  function outilDocuments() {
  'use strict';
  const VERSION = 'v1.4';
  const HOST_ID = 'moore-mmf-documents';

  const CONFIG = {
    FORMAT_NOM: '{client} - {date} - {titre}',   // fichiers séparés ; champs : {client} {bce} {date} {titre} {onglet}
    FORMAT_NOM_ZIP: '{date} - {titre}',          // fichiers dans le ZIP (le client est dans le nom du ZIP)
    PAUSE_MS: 1200,                               // pause entre deux documents (fichiers séparés)
    PAUSE_MS_ZIP: 500,                            // pause entre deux documents (ZIP)
    MODE_RAPIDE: true,                            // appels directs à MyMinfin quand c'est possible
    PARALLELE: 3,                                 // documents demandés en même temps (ZIP, mode rapide)
    PAUSE_RAPIDE_FICHIERS_MS: 300,                // pause entre deux fichiers séparés (mode rapide)
    DELAI_MAX_MS: 20000,                          // attente maximale par document
  };

  if (window.__mmfDocs && window.__mmfDocs.enCours) {
    console.warn('Un téléchargement MyMinfin est déjà en cours dans cet onglet.');
    return;
  }

  // ---------------------------------------------------------------- outils
  const clean = s => String(s ?? '').replace(/[\s  ]+/g, ' ').trim();
  const txt = el => (el ? clean(el.innerText || el.textContent) : '');
  const simple = s => clean(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const pause = ms => new Promise(r => setTimeout(r, ms));
  const attendre = async (test, ms, pas = 100) => {
    const fin = Date.now() + ms;
    for (;;) { const v = test(); if (v) return v; if (Date.now() > fin) return null; await pause(pas); }
  };

  const MOIS = [/^jan/, /^f.{0,2}v|^feb/, /^mar|^maa/, /^av|^apr/, /^ma[iy]|^mei/, /^juin|^jun/,
                /^juil|^jul/, /^ao|^aug/, /^sep/, /^oct|^okt/, /^nov/, /^d.{0,2}c/];
  const dateISO = s => {
    const t = simple(s);
    let m = t.match(/(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    m = t.match(/(\d{1,2})(?:er)?\s+(\S+)\s+(\d{4})/);
    if (!m) return '';
    const i = MOIS.findIndex(re => re.test(m[2]));
    return i < 0 ? '' : `${m[3]}-${String(i + 1).padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  };
  const isoVersFR = d => (d ? d.split('-').reverse().join('/') : '');
  const nettoyerNom = s => clean(s).replace(/[\\/:*?"<>|\u0000-\u001f\u007f]/g, '-').replace(/[. ]+$/, '').slice(0, 150).trim();

  // ---------------------------------------------------------------- client
  const ident = document.querySelector('#user-info-details .myminfin-details-identifier');
  const client = txt(ident && ident.querySelector(':scope > span'));
  const bceSrc = txt(ident && ident.querySelector(':scope > div')) || txt(document.querySelector('.profil-number'));
  const bceM = bceSrc.match(/([01]\d{3})[.\s]?(\d{3})[.\s]?(\d{3})/);
  const bce = bceM ? `${bceM[1]}.${bceM[2]}.${bceM[3]}` : '';

  // -------------------------------------------------- lecture des documents
  const ongletActif = () => txt(document.querySelector('a.mdc-tab--active .mdc-tab__text-label, a.mdc-tab--active')) || '';
  const lireBoutons = () => {
    let boutons = [...document.querySelectorAll('mat-tab-nav-panel button[aria-label="actions-menu"], mat-tab-nav-panel button.mat-mdc-menu-trigger')];
    if (!boutons.length) boutons = [...document.querySelectorAll('button[aria-label="actions-menu"]')];
    return [...new Set(boutons)].filter(b => b.closest('tr') && !b.closest('#' + HOST_ID));
  };
  const lireDocuments = () => {
    const boutons = lireBoutons();
    const onglet = ongletActif();
    const vus = new Set(), vusZip = new Set();
    const unique = (nom, ens) => {
      if (ens.has(nom.toLowerCase())) { let k = 2; while (ens.has(`${nom} (${k})`.toLowerCase())) k++; nom = `${nom} (${k})`; }
      ens.add(nom.toLowerCase());
      return nom;
    };
    const former = (format, titre, date) => nettoyerNom(format
      .replace('{client}', client || 'MyMinfin').replace('{bce}', bce)
      .replace('{date}', date || 'sans date').replace('{titre}', titre).replace('{onglet}', onglet));
    return boutons.map((btn, i) => {
      const tr = btn.closest('tr');
      const cellules = [...tr.querySelectorAll(':scope > td')].map(txt).filter(Boolean);
      const titre = txt(tr.querySelector('[class*="column-documentTitle"], [class*="column-NAME"]')) || cellules[0] || `Document ${i + 1}`;
      const dateTxt = txt(tr.querySelector('[class*="column-documentDate"], [class*="column-DATE"]')) || cellules.find(c => /\d{4}/.test(c)) || '';
      const date = dateISO(dateTxt);
      const nom = unique(former(CONFIG.FORMAT_NOM, titre, date), vus);
      const nomZip = unique(former(CONFIG.FORMAT_NOM_ZIP, titre, date), vusZip);
      return { i, btn, titre, date, nom, nomZip, statut: '', detail: '', fichier: '' };
    });
  };

  // ------------------------------------------------ capture du téléchargement
  const orig = {
    aClick: HTMLAnchorElement.prototype.click,
    dispatch: EventTarget.prototype.dispatchEvent,
    open: window.open,
    createURL: URL.createObjectURL,
    xhrOpen: XMLHttpRequest.prototype.open,
    xhrSend: XMLHttpRequest.prototype.send,
    xhrHeader: XMLHttpRequest.prototype.setRequestHeader,
    fetch: window.fetch,
  };
  const diag = { menus: new Set(), methodes: {}, reseau: new Set(), erreurs: [], ouvertures: new Set(),
                 appels: new Map(), ongletAuChargement: '', sondes: null };
  let cap = null;   // document en cours de capture
  // Mode rapide : en-têtes et listes vus dans les appels de MyMinfin (en mémoire seulement, jamais copiés)
  const rapide = { entetes: null, base: '', listes: new Map(), corpsListe: null, listeRecue: 0 };

  const EXT = {
    'application/pdf': '.pdf', 'application/zip': '.zip', 'application/x-zip-compressed': '.zip',
    'application/xml': '.xml', 'text/xml': '.xml', 'text/html': '.html', 'text/plain': '.txt',
    'text/csv': '.csv', 'image/png': '.png', 'image/jpeg': '.jpg', 'image/tiff': '.tif',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
    'application/vnd.ms-excel': '.xls', 'application/msword': '.doc',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
  };
  const extension = (nomServeur, mime) => {
    const m = String(nomServeur || '').match(/\.([a-z0-9]{2,5})$/i);
    if (m) return '.' + m[1].toLowerCase();
    return EXT[String(mime || '').split(';')[0].trim().toLowerCase()] || '.pdf';
  };
  const nomContentDisposition = cd => {
    if (!cd) return '';
    let m = cd.match(/filename\*\s*=\s*[^']*''([^;]+)/i);
    if (m) { try { return decodeURIComponent(m[1].trim()); } catch (e) { /* ignore */ } }
    m = cd.match(/filename\s*=\s*"?([^";]+)"?/i);
    return m ? m[1].trim() : '';
  };
  const memeOrigine = u => { try { return new URL(u, location.href).origin === location.origin; } catch (e) { return false; } };

  const enregistrerURL = (url, nom) => {
    const a = document.createElement('a');
    a.href = url; a.download = nom; a.style.display = 'none';
    document.body.appendChild(a); orig.aClick.call(a);
    setTimeout(() => a.remove(), 2000);
  };
  const enregistrerBlob = (blob, nom) => {
    const url = orig.createURL.call(URL, blob);
    enregistrerURL(url, nom);
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  };

  const terminer = (c, methode, fichier, blob) => {
    if (!c || c.fini) return;
    c.fini = true; clearTimeout(c.minuteur); clearTimeout(c.apercu);
    diag.methodes[methode] = (diag.methodes[methode] || 0) + 1;
    c.resoudre({ ok: true, methode, fichier, blob });
  };
  const echouer = (c, raison) => {
    if (!c || c.fini) return;
    c.fini = true; clearTimeout(c.minuteur); clearTimeout(c.apercu);
    diag.erreurs.push(raison);
    c.resoudre({ ok: false, raison });
  };
  // Le fichier est récupéré : gardé pour le ZIP, ou enregistré tout de suite sous son nouveau nom
  const livrer = (c, blob, nomServeur, methode) => {
    if (!c || c.fini) return;
    const nom = c.base + extension(nomServeur, blob.type);
    if (c.mode === 'zip') terminer(c, methode, nom, blob);
    else { enregistrerBlob(blob, nom); terminer(c, methode, nom); }
  };

  const recupererURL = (c, url, nomServeur = '') => {
    orig.fetch.call(window, url, /^https?:/i.test(url) ? { credentials: 'include' } : undefined)
      .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.blob().then(b => ({ b, cd: r.headers.get('content-disposition') })); })
      .then(({ b, cd }) => {
        if (/text\/html/i.test(b.type)) echouer(c, `MyMinfin ouvre une page (${masquerURL(url)}) au lieu d’un fichier`);
        else livrer(c, b, nomContentDisposition(cd) || nomServeur, 'adresse');
      })
      .catch(e => echouer(c, 'téléchargement refusé (' + e.message + ')'));
  };
  const prendreURL = (c, u) => {
    if (/^(blob|data):/i.test(u)) {
      const blob = c.blobs.get(u);
      if (blob) livrer(c, blob, '', 'ouverture');
      else if (c.mode === 'zip') recupererURL(c, u);
      else { const nom = c.base + extension('', ''); enregistrerURL(u, nom); terminer(c, 'ouverture', nom); }
      return true;
    }
    if (memeOrigine(u)) { recupererURL(c, new URL(u, location.href).href); return true; }
    return false;
  };
  // window.open('') suivi de win.location = url : on renvoie une fenêtre factice
  const fenetreFactice = c => {
    const aller = v => { const u = String(v || ''); if (u && u !== 'about:blank' && !c.fini) prendreURL(c, u); };
    const loc = { assign: aller, replace: aller, reload() {}, toString: () => '' };
    Object.defineProperty(loc, 'href', { get: () => '', set: aller });
    const w = { closed: false, opener: null, close() { this.closed = true; }, focus() {}, blur() {},
                document: { write() {}, writeln() {}, open() {}, close() {}, title: '' } };
    Object.defineProperty(w, 'location', { get: () => loc, set: aller });
    return w;
  };
  // Lien cliqué par MyMinfin pendant la capture : on reprend le fichier (renommé ou mis dans le ZIP)
  const interceptAncre = a => {
    const c = cap;
    if (!c || c.fini || !a.getAttribute('href')) return false;
    const href = a.href;
    if (/^(blob|data):/i.test(href)) {
      const blob = c.blobs.get(href);
      if (blob) { livrer(c, blob, a.download, 'lien'); return true; }          // clic d'origine remplacé
      if (c.mode === 'zip') { recupererURL(c, href, a.download); return true; }
      const nom = c.base + extension(a.download, '');
      a.download = nom; terminer(c, 'lien', nom); return false;               // clic d'origine conservé, renommé
    }
    if (/^https?:/i.test(href) && memeOrigine(href) && (a.hasAttribute('download') || a.target === '_blank')) {
      recupererURL(c, href, a.download); return true;
    }
    return false;
  };
  // Adresse sans données : numéros remplacés par #, paramètres sans leurs valeurs
  const masquerURL = u => {
    try {
      const x = new URL(u, location.href);
      return (x.host !== location.host ? x.host : '') +
        x.pathname.replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, '#').replace(/\d{3,}/g, '#') +
        (x.search ? '?' + [...new Set(x.searchParams.keys())].join('&') : '');
    } catch (e) { return ''; }
  };
  // Forme d'une réponse JSON, sans aucune valeur : noms de champs, types, nombre d'éléments
  const formeJSON = (v, prof = 0) => {
    if (prof > 5) return '…';
    if (Array.isArray(v)) return v.length ? [formeJSON(v[0], prof + 1), `${v.length} élément(s)`] : [];
    if (v && typeof v === 'object') {
      const cles = Object.keys(v), o = {};
      cles.slice(0, 40).forEach(k => { o[k.replace(/\d{3,}/g, '#')] = formeJSON(v[k], prof + 1); });
      if (cles.length > 40) o['…'] = `${cles.length} champs`;
      return o;
    }
    if (typeof v === 'string') return /^\d{4}-\d{2}-\d{2}/.test(v) ? 'date' : /^\d+$/.test(v) ? `chiffres(${v.length})` : `texte(${v.length})`;
    return v === null ? 'null' : typeof v;
  };
  const formeCorps = body => {
    if (body == null) return null;
    if (typeof body === 'string') { try { return formeJSON(JSON.parse(body)); } catch (e) { return `texte(${body.length})`; } }
    if (typeof FormData !== 'undefined' && body instanceof FormData) return { formData: [...body.keys()] };
    if (body instanceof URLSearchParams) return { parametres: [...body.keys()] };
    if (body instanceof Blob) return `blob(${body.type || '?'})`;
    if (body instanceof ArrayBuffer || ArrayBuffer.isView(body)) return 'binaire';
    return typeof body;
  };
  const estREST = u => { try { return /\/myminfin-rest\//.test(new URL(u, location.href).pathname); } catch (e) { return false; } };
  const noterAppel = (methode, url, details, pendantCapture) => {
    const m = masquerURL(url);
    if (!m) return;
    const cle = `${String(methode || 'GET').toUpperCase()} ${m}`;
    if (pendantCapture) diag.reseau.add(cle);
    if (!estREST(url)) return;
    const e = diag.appels.get(cle) || { appel: cle, nombre: 0 };
    e.nombre++;
    Object.assign(e, details);
    diag.appels.delete(cle); diag.appels.set(cle, e);
  };
  // Enregistreur passif, actif tant que le panneau est ouvert : il ne modifie aucun appel
  const ecouterReseau = () => {
    XMLHttpRequest.prototype.open = function (m, u, ...r) {
      this.__mmf = { m, u, entetes: [], valeurs: {}, capture: !!(cap && !cap.fini) };
      return orig.xhrOpen.call(this, m, u, ...r);
    };
    XMLHttpRequest.prototype.setRequestHeader = function (n, v) {
      if (this.__mmf) { this.__mmf.entetes.push(n); this.__mmf.valeurs[n] = v; }
      return orig.xhrHeader.call(this, n, v);
    };
    XMLHttpRequest.prototype.send = function (body) {
      const info = this.__mmf;
      if (info) {
        apprendreAppel(info, body);
        const corps = formeCorps(body);
        this.addEventListener('loadend', () => {
          if (estListeDocs(info.u) && this.status === 200) {
            try { retenirListe(this.responseType === 'json' ? this.response : JSON.parse(this.responseText)); } catch (e) { /* ignore */ }
          }
          const reponse = { statut: this.status, type: (this.getResponseHeader('content-type') || '').split(';')[0], responseType: this.responseType || 'text' };
          const cd = this.getResponseHeader('content-disposition');
          if (cd) reponse.fichier = extension(nomContentDisposition(cd), '');
          try {
            if (this.responseType === 'json') reponse.forme = formeJSON(this.response);
            else if ((!this.responseType || this.responseType === 'text') && /json/.test(reponse.type)) reponse.forme = formeJSON(JSON.parse(this.responseText));
          } catch (e) { /* réponse illisible */ }
          noterAppel(info.m, info.u, { entetes: [...new Set(info.entetes)], corps, reponse }, info.capture);
        });
      }
      return orig.xhrSend.call(this, body);
    };
    window.fetch = function (input, init) {
      const u = typeof input === 'string' ? input : input && input.url;
      const m = (init && init.method) || (input && input.method) || 'GET';
      const pendant = !!(cap && !cap.fini);
      const p = orig.fetch.apply(this, arguments);
      if (u) {
        const corps = formeCorps(init && init.body);
        let entetes = [];
        try { entetes = init && init.headers ? [...new Headers(init.headers).keys()] : []; } catch (e) { /* ignore */ }
        p.then(r => {
          const reponse = { statut: r.status, type: (r.headers.get('content-type') || '').split(';')[0] };
          if (/json/.test(reponse.type) && estREST(u)) {
            r.clone().json().then(j => noterAppel(m, u, { entetes, corps, reponse: { ...reponse, forme: formeJSON(j) } }, pendant),
                                  () => noterAppel(m, u, { entetes, corps, reponse }, pendant));
          } else noterAppel(m, u, { entetes, corps, reponse }, pendant);
        }, () => noterAppel(m, u, { entetes, corps, reponse: { erreur: 'réseau' } }, pendant));
      }
      return p;
    };
  };
  const arreterReseau = () => {
    XMLHttpRequest.prototype.open = orig.xhrOpen;
    XMLHttpRequest.prototype.setRequestHeader = orig.xhrHeader;
    XMLHttpRequest.prototype.send = orig.xhrSend;
    window.fetch = orig.fetch;
  };
  // Relit les listes de documents déjà demandées par MyMinfin (GET, lecture seule) pour en décrire la forme
  const sonder = async () => {
    const adresses = [...new Set(performance.getEntriesByType('resource')
      .filter(e => /xmlhttprequest|fetch/.test(e.initiatorType) && /\/myminfin-rest\/documents\//.test(e.name) && !/download|getextension/i.test(e.name))
      .map(e => e.name))].slice(0, 8);
    const res = [];
    for (const u of adresses) {
      try {
        const r = await orig.fetch.call(window, u, { credentials: 'include', headers: { Accept: 'application/json' } });
        const type = (r.headers.get('content-type') || '').split(';')[0];
        const sonde = { adresse: 'GET ' + masquerURL(u), statut: r.status, type };
        if (/json/.test(type)) { try { sonde.forme = formeJSON(await r.json()); } catch (e) { sonde.forme = 'illisible'; } }
        res.push(sonde);
      } catch (e) { res.push({ adresse: masquerURL(u), erreur: e.message }); }
    }
    return res;
  };

  // ---------------------------------------------------------- mode rapide
  // MyMinfin décrit chaque document (titre, dates, clé) dans la liste qu'il charge, puis demande le
  // fichier avec la clé du document. Le mode rapide refait ces appels avec les mêmes en-têtes.
  const cheminDe = u => { try { return new URL(u, location.href).pathname; } catch (e) { return ''; } };
  const estListeDocs = u => /\/myminfin-rest\/documents\/cdis-no-actions$/.test(cheminDe(u));
  const extraireDocs = j => {
    const l = (((j || {}).data || {}).documents || {}).documentDescriptionList;
    return Array.isArray(l) ? l : [];
  };
  const apprendreAppel = (info, body) => {
    if (!/\/myminfin-rest\/documents\//.test(cheminDe(info.u))) return;
    if (Object.keys(info.valeurs).length) {
      rapide.entetes = { ...info.valeurs };
      const x = new URL(info.u, location.href);
      rapide.base = x.origin + x.pathname.replace(/\/documents\/.*$/, '/documents/');
    }
    if (estListeDocs(info.u) && typeof body === 'string') rapide.corpsListe = body;
  };
  const retenirListe = j => { rapide.listes.set(ongletActif(), extraireDocs(j)); rapide.listeRecue = Date.now(); };
  const cleTitre = t => simple(t).replace(/[^a-z0-9]+/g, ' ').trim();
  const jourDe = v => String(v || '').slice(0, 10);
  // Lignes du panneau ↔ documents de MyMinfin, par titre et date (documentDate ou publishDate)
  const associer = (liste, lignes) => {
    const idx = new Map();
    const ajouter = (k, doc) => { if (!idx.has(k)) idx.set(k, []); if (!idx.get(k).includes(doc)) idx.get(k).push(doc); };
    liste.forEach(doc => {
      if (!doc || !doc.documentKey) return;
      const t = cleTitre(doc.documentTitle);
      ajouter(t + '|' + jourDe(doc.documentDate), doc);
      ajouter(t + '|' + jourDe(doc.publishDate), doc);
    });
    const pris = new Set(), res = new Map();
    lignes.forEach(d => {
      const doc = (idx.get(cleTitre(d.titre) + '|' + d.date) || []).find(c => !pris.has(c));
      if (doc) { pris.add(doc); res.set(d, doc); }
    });
    return res;
  };
  const appelDocs = (chemin, corps) => orig.fetch.call(window, rapide.base + chemin, {
    method: corps == null ? 'GET' : 'POST', credentials: 'include',
    headers: { ...rapide.entetes, 'Content-Type': 'application/json' },
    body: corps == null ? undefined : corps,
  });
  // Ouvre un menu « ⋮ » : MyMinfin fait alors un appel « documents » dont on retient les en-têtes
  const apprendreEntetes = async d => {
    if (rapide.entetes) return true;
    fermerMenus();
    await attendre(() => panneauxMenu().length === 0, 1500);
    if (!d || !d.btn.isConnected) return false;
    d.btn.scrollIntoView({ block: 'center' });
    d.btn.click();
    await attendre(() => rapide.entetes, 5000);
    await pause(150);
    fermerMenus();
    return !!rapide.entetes;
  };
  // Liste des documents de l'onglet : celle vue passer, sinon redemandée à MyMinfin
  const listePourOnglet = async lignes => {
    const onglet = ongletActif();
    const connue = rapide.listes.get(onglet);
    let meilleure = connue || [], score = connue ? associer(connue, lignes).size : -1;
    if (score >= Math.ceil(lignes.length / 2)) return meilleure;
    const corps = [];
    if (rapide.corpsListe) corps.push(rapide.corpsListe);
    let pro = null;
    try { const r = await appelDocs('isProEnabled'); if (r.ok) pro = await r.json(); } catch (e) { /* ignore */ }
    corps.push(JSON.stringify({ proProfile: pro === true }), JSON.stringify({ proProfile: pro !== true }));
    for (const c of [...new Set(corps)]) {
      if (score === lignes.length) break;
      try {
        const r = await appelDocs('cdis-no-actions', c);
        if (!r.ok) continue;
        const liste = extraireDocs(await r.json());
        const sc = associer(liste, lignes).size;
        if (sc > score) { score = sc; meilleure = liste; }
      } catch (e) { /* essai suivant */ }
    }
    rapide.listes.set(onglet, meilleure);
    return meilleure;
  };
  const telechargerDirect = async doc => {
    const r = await appelDocs('download', JSON.stringify(doc.documentKey));
    const type = (r.headers.get('content-type') || '').split(';')[0].trim();
    if (!r.ok || /json|html/i.test(type)) {
      let raison = 'HTTP ' + r.status;
      if (/json/i.test(type)) {
        try { const j = await r.json(); if (j && j.sessionExpired) raison = 'session MyMinfin expirée'; else if (j && j.error) raison = 'refus de MyMinfin'; } catch (e) { /* ignore */ }
      }
      throw new Error(raison);
    }
    const buf = await r.arrayBuffer();
    if (!buf.byteLength) throw new Error('fichier vide');
    return { blob: new Blob([buf], { type: type || doc.mimeType || 'application/pdf' }), nomServeur: nomContentDisposition(r.headers.get('content-disposition')) };
  };

  const installer = () => {
    HTMLAnchorElement.prototype.click = function () { if (cap && interceptAncre(this)) return; return orig.aClick.call(this); };
    EventTarget.prototype.dispatchEvent = function (ev) {
      if (cap && this instanceof HTMLAnchorElement && ev && ev.type === 'click' && interceptAncre(this)) return true;
      return orig.dispatch.call(this, ev);
    };
    window.open = function (url, ...reste) {
      const c = cap;
      if (c && !c.fini) {
        const u = String(url ?? '');
        if (!u || u === 'about:blank' || prendreURL(c, u)) return fenetreFactice(c);
        const cible = masquerURL(u);
        diag.ouvertures.add(cible);
        echouer(c, `MyMinfin ouvre une autre application (${cible.split('/')[0] || cible}) : document non téléchargeable par le favori`);
        return fenetreFactice(c);
      }
      return orig.open.call(window, url, ...reste);
    };
    URL.createObjectURL = function (obj) {
      const url = orig.createURL.call(URL, obj);
      const c = cap;
      if (c && !c.fini && obj instanceof Blob && obj.size > 200 && !/svg|javascript|css/i.test(obj.type)) {
        c.blobs.set(url, obj);
        clearTimeout(c.apercu);
        // MyMinfin se contente d'afficher le document (aperçu) : on l'enregistre nous-mêmes
        c.apercu = setTimeout(() => livrer(c, obj, '', 'aperçu'), 3000);
      }
      return url;
    };
  };
  const desinstaller = () => {
    HTMLAnchorElement.prototype.click = orig.aClick;
    EventTarget.prototype.dispatchEvent = orig.dispatch;
    window.open = orig.open;
    URL.createObjectURL = orig.createURL;
  };

  // ----------------------------------------------------------- menu « ⋮ »
  const panneauxMenu = () => [...document.querySelectorAll(
    '.cdk-overlay-container .mat-mdc-menu-panel, .cdk-overlay-container .mat-menu-panel, .cdk-overlay-container [role="menu"]')];
  const elementsMenu = p => [...p.querySelectorAll('[mat-menu-item], .mat-mdc-menu-item, [role="menuitem"]')]
    .filter(i => !i.disabled && i.getAttribute('aria-disabled') !== 'true');
  const RE_TELECHARGER = /t.{1,2}l.{1,2}charg|download|enregistr|downl/;
  const RE_VOIR = /ouvrir|afficher|consulter|visualis|voir|open|pdf/;
  // Actions qu'on ne clique jamais, même si leur libellé contient un mot ci-dessus
  const RE_INTERDIT = /supprim|archiv|effac|delete|signal|r.{1,2}pond|transf|rembours|partag|envoy|marquer/;
  const fermerMenus = () => {
    document.querySelectorAll('.cdk-overlay-container .cdk-overlay-backdrop').forEach(b => b.click());
    document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true }));
  };

  const traiter = async (d, mode) => {
    fermerMenus();
    await attendre(() => panneauxMenu().length === 0, 1500);
    if (!d.btn.isConnected) return { ok: false, raison: 'liste modifiée : clique « Actualiser la liste »' };
    d.btn.scrollIntoView({ block: 'center' });
    const c = { base: mode === 'zip' ? d.nomZip : d.nom, mode, fini: false, blobs: new Map() };
    c.promesse = new Promise(r => (c.resoudre = r));
    c.minuteur = setTimeout(() => echouer(c, 'aucun téléchargement détecté (vérifie quand même le dossier Téléchargements)'), CONFIG.DELAI_MAX_MS);
    cap = c;
    try {
      d.btn.click();
      const panel = await attendre(() => panneauxMenu().pop(), 8000);
      if (!panel) {
        echouer(c, 'le menu « ⋮ » ne s’est pas ouvert');
      } else {
        // À la première ouverture, MyMinfin peut encore charger les actions du document : on les attend,
        // d'abord « Télécharger », puis à défaut « Consulter / Ouvrir »
        const permis = () => elementsMenu(panel).filter(i => !RE_INTERDIT.test(simple(txt(i))));
        const telecharger = () => permis().find(i => RE_TELECHARGER.test(simple(txt(i))));
        const voir = () => permis().find(i => RE_VOIR.test(simple(txt(i))));
        let choix = await attendre(() => (!panel.isConnected ? 'fermé' : telecharger()), 2500);
        if (!choix) choix = await attendre(() => (!panel.isConnected ? 'fermé' : telecharger() || voir()), 4000);
        const libelles = elementsMenu(panel).map(txt);
        diag.menus.add(libelles.join(' | ') || '(menu vide)');
        if (!choix || choix === 'fermé') {
          echouer(c, (choix === 'fermé' ? 'le menu s’est refermé' : 'pas d’action de téléchargement dans le menu') +
                     ' (' + (libelles.join(' / ') || 'menu vide') + ')');
        } else {
          const sousMenu = choix.matches('.mat-mdc-menu-item-submenu-trigger, [aria-haspopup="menu"], [aria-haspopup="true"]');
          const nbAvant = panneauxMenu().length;
          const adresseAvant = location.href;
          c.surveillance = setInterval(() => {
            if (location.href !== adresseAvant) echouer(c, `MyMinfin a ouvert une autre page (${masquerURL(location.href)}) : document non téléchargeable par le favori`);
          }, 200);
          choix.click();
          if (sousMenu) {
            const p2 = await attendre(() => (panneauxMenu().length > nbAvant ? panneauxMenu().pop() : null), 3000);
            const it2 = p2 ? elementsMenu(p2) : [];
            diag.menus.add('↳ ' + it2.map(txt).join(' | '));
            const c2 = it2.find(i => /fran.{1,2}ais|\bfr\b/.test(simple(txt(i)))) || it2[0];
            if (c2) c2.click(); else echouer(c, 'sous-menu vide');
          }
        }
      }
      return await c.promesse;
    } finally {
      await pause(mode === 'zip' ? 100 : 400);   // laisser le navigateur lancer le téléchargement
      clearInterval(c.surveillance);
      cap = null;
      fermerMenus();
    }
  };

  // --------------------------------------------------------------- panneau
  document.getElementById(HOST_ID)?.remove();
  const host = document.createElement('div');
  host.id = HOST_ID;
  host.style.cssText = 'position:fixed;top:16px;right:16px;z-index:2147483647';
  const root = host.attachShadow({ mode: 'open' });
  const h = (tag, cls, texte) => { const e = document.createElement(tag); if (cls) e.className = cls; if (texte != null) e.textContent = texte; return e; };
  const style = h('style');
  style.textContent = `
    .card{width:430px;font:13px/1.4 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:#1d2433;background:#fff;
          border:1px solid #d5dbe5;border-radius:10px;box-shadow:0 8px 28px rgba(16,24,40,.18);overflow:hidden}
    .top{display:flex;align-items:center;justify-content:space-between;background:#16324f;color:#fff;padding:9px 12px}
    .top b{font-size:13px}.top span{opacity:.75;font-size:11px;margin-left:6px}
    .x{all:unset;cursor:pointer;font-size:18px;line-height:1;padding:0 4px;color:#fff}
    .body{padding:12px}
    .client{font-weight:600;font-size:14px}.sub{color:#5b6475;font-size:12px}
    .warn{margin-top:8px;background:#fff7e6;border:1px solid #f3d08a;border-radius:6px;padding:7px 10px;color:#6b4a00;font-size:12px}
    .warn div+div{margin-top:3px}
    .tools{display:flex;flex-wrap:wrap;align-items:center;gap:4px 12px;margin:10px 0 6px;font-size:12px}
    .tools a{color:#16324f;cursor:pointer;text-decoration:underline;white-space:nowrap}
    .tools .n{margin-left:auto;color:#5b6475;white-space:nowrap}
    .filtres{margin-top:10px;display:flex;flex-direction:column;gap:7px;font-size:12px}
    .chips,.dates{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
    .lab{color:#5b6475}
    button.chip{all:unset;cursor:pointer;padding:3px 9px;border:1px solid #c9d2df;border-radius:12px;color:#16324f;font-weight:600;font-size:12px}
    button.chip small{font-weight:400;color:#5b6475;margin-left:5px}
    button.chip[aria-pressed="true"]{background:#16324f;border-color:#16324f;color:#fff}
    button.chip[aria-pressed="true"] small{color:#d6deea}
    .dates input{font:12px system-ui,sans-serif;padding:2px 4px;border:1px solid #c9d2df;border-radius:4px;color:#1d2433}
    .dates a{color:#16324f;cursor:pointer;text-decoration:underline}
    .list{max-height:280px;overflow:auto;border:1px solid #e3e7ee;border-radius:6px}
    .row{display:grid;grid-template-columns:18px 78px 1fr 18px;gap:6px;align-items:start;padding:5px 8px;border-top:1px solid #f0f2f6;cursor:pointer}
    .row:first-child{border-top:0}.row:hover{background:#f6f8fb}
    .row input{margin:2px 0 0}.d{color:#5b6475;font-variant-numeric:tabular-nums}.t{word-break:break-word}
    .s{text-align:center}.s.ok{color:#1a7f37}.s.ko{color:#c0392b}.s.run{color:#b26a00}
    .prog{margin-top:8px;font-size:12px;color:#5b6475;min-height:16px}
    .btns{display:flex;gap:6px;margin-top:10px;flex-wrap:wrap}
    button.b{all:unset;box-sizing:border-box;cursor:pointer;text-align:center;padding:8px 10px;border-radius:6px;
             border:1px solid #16324f;color:#16324f;font-weight:600;flex:1 1 auto}
    button.b.p{background:#16324f;color:#fff}button.b.stop{background:#c0392b;border-color:#c0392b;color:#fff}
    button.b[disabled]{opacity:.45;cursor:default}
    .note{margin-top:8px;color:#5b6475;font-size:11px}`;
  const card = h('div', 'card');
  const top = h('div', 'top');
  const tt = h('div'); tt.append(h('b', null, 'Moore - Outils'), h('span', null, `MyMinfin · documents ${VERSION}`));
  const x = h('button', 'x', '×'); x.title = 'Fermer';
  top.append(tt, x);
  const body = h('div', 'body');
  const enTete = h('div');
  const avertBox = h('div', 'warn');
  // Filtres : années (boutons) et période de publication (du / au)
  const filtres = h('div', 'filtres');
  const chips = h('div', 'chips');
  const dates = h('div', 'dates');
  const inDu = h('input'), inAu = h('input');
  inDu.type = inAu.type = 'date';
  const lienEffacer = h('a', null, 'effacer');
  dates.append(h('span', 'lab', 'Publiés du'), inDu, h('span', 'lab', 'au'), inAu, lienEffacer);
  filtres.append(chips, dates);
  const tools = h('div', 'tools');
  const lienTous = h('a', null, 'Tout cocher'), lienAucun = h('a', null, 'Tout décocher'), lienMaj = h('a', null, 'Actualiser la liste');
  const lienDiag = h('a', null, 'Copier le diagnostic');
  const compteur = h('span', 'n');
  tools.append(lienTous, lienAucun, lienMaj, lienDiag, compteur);
  const liste = h('div', 'list');
  const prog = h('div', 'prog');
  const btns = h('div', 'btns');
  const bZip = h('button', 'b p', 'Télécharger en ZIP');
  const bFichiers = h('button', 'b', 'Fichiers séparés');
  const bJournal = h('button', 'b', 'Copier le journal');
  bJournal.style.display = 'none';
  btns.append(bZip, bFichiers, bJournal);
  body.append(enTete, avertBox, filtres, tools, liste, prog, btns,
    h('div', 'note', 'Mode rapide : les fichiers sont demandés directement à MyMinfin, 3 à la fois ; à défaut, menu ⋮ un document à la fois. Rien n’est supprimé. ZIP : un seul fichier à la fin, avec la liste des documents.'));
  card.append(top, body);
  root.append(style, card);
  document.body.appendChild(host);

  let docs = [];
  let arret = false;
  const etat = window.__mmfDocs = { enCours: false, docs: () => docs, diag: () => diagTexte() };

  // Filtre : seuls les documents filtrés sont affichés et téléchargés ; les cases restent modifiables
  let filtre = { annees: new Set(), du: '', au: '' };
  const anneeDe = d => (d.date ? d.date.slice(0, 4) : 'sans date');
  const listeAnnees = () => [...new Set(docs.map(anneeDe))]
    .sort((a, b) => (a === 'sans date') - (b === 'sans date') || b.localeCompare(a));
  const visible = d => filtre.annees.has(anneeDe(d)) &&
    (!filtre.du || (d.date && d.date >= filtre.du)) && (!filtre.au || (d.date && d.date <= filtre.au));
  const visibles = () => docs.filter(visible);
  const nbCoches = () => visibles().filter(d => d.coche).length;
  const majCompteur = () => {
    const nv = visibles().length;
    compteur.textContent = `${nbCoches()} / ${nv} coché(s)` + (nv < docs.length ? ` · ${docs.length} au total` : '');
    if (!etat.enCours) {
      const n = nbCoches();
      bZip.textContent = `Télécharger en ZIP (${n})`; bFichiers.textContent = `Fichiers séparés (${n})`;
      bZip.disabled = bFichiers.disabled = !n;
    }
  };
  const afficherFiltres = () => {
    const annees = listeAnnees();
    const toutes = annees.every(a => filtre.annees.has(a));
    chips.replaceChildren(h('span', 'lab', 'Années :'));
    const chip = (label, n, actif, action) => {
      const b = h('button', 'chip', label);
      b.append(h('small', null, String(n)));
      b.setAttribute('aria-pressed', actif ? 'true' : 'false');
      b.onclick = () => { if (!etat.enCours) { action(); rafraichir(); } };
      chips.append(b);
    };
    chip('Toutes', docs.length, toutes && !filtre.du && !filtre.au, () => {
      filtre = { annees: new Set(annees), du: '', au: '' }; inDu.value = inAu.value = '';
    });
    annees.forEach(a => chip(a, docs.filter(d => anneeDe(d) === a).length, !toutes && filtre.annees.has(a), () => {
      if (toutes) filtre.annees = new Set([a]);                       // 1er clic : seulement cette année
      else if (filtre.annees.has(a)) { filtre.annees.delete(a); if (!filtre.annees.size) filtre.annees = new Set(annees); }
      else filtre.annees.add(a);                                      // clics suivants : ajouter / retirer
    }));
  };
  const rafraichir = () => { afficherFiltres(); afficherListe(); };
  inDu.onchange = () => { if (!etat.enCours) { filtre.du = inDu.value; rafraichir(); } };
  inAu.onchange = () => { if (!etat.enCours) { filtre.au = inAu.value; rafraichir(); } };
  lienEffacer.onclick = () => { if (!etat.enCours) { inDu.value = inAu.value = ''; filtre.du = filtre.au = ''; rafraichir(); } };

  const SYMB = { '': '', run: '⏳', ok: '✔', ko: '✖' };
  const afficherListe = () => {
    liste.replaceChildren();
    if (!docs.length) liste.append(h('div', 'row', 'Aucun document affiché dans cet onglet.'));
    else if (!visibles().length) liste.append(h('div', 'row', 'Aucun document pour ce filtre.'));
    visibles().forEach(d => {
      const r = h('label', 'row');
      const cb = h('input'); cb.type = 'checkbox'; cb.checked = d.coche;
      cb.onchange = () => { d.coche = cb.checked; majCompteur(); };
      d.cb = cb;
      d.cellStatut = h('span', 's ' + d.statut, SYMB[d.statut]);
      d.cellStatut.title = d.detail;
      r.append(cb, h('span', 'd', isoVersFR(d.date) || '—'), h('span', 't', d.titre), d.cellStatut);
      liste.append(r);
    });
    majCompteur();
  };
  const statut = (d, s, detail) => {
    d.statut = s; d.detail = detail || '';
    if (d.cellStatut) { d.cellStatut.className = 's ' + s; d.cellStatut.textContent = SYMB[s]; d.cellStatut.title = d.detail; }
  };
  let signatureListe = '', signatureVue = '', listeVue = 0;
  const signatureDOM = () => {
    const b = lireBoutons();
    const ligne = i => (b[i] ? txt(b[i].closest('tr')).slice(0, 80) : '');
    return [ongletActif(), b.length, ligne(0), ligne(b.length - 1)].join('|');
  };
  const charger = () => {
    docs = lireDocuments().map(d => ({ ...d, coche: true }));
    const onglet = ongletActif();
    signatureListe = signatureDOM(); signatureVue = ''; listeVue = Date.now();
    diag.ongletAuChargement = onglet;
    enTete.replaceChildren(h('div', 'client', client || '(aucun client sélectionné)'),
      h('div', 'sub', [bce && `BCE ${bce}`, onglet && `Onglet « ${onglet} »`, `${docs.length} document(s)`].filter(Boolean).join(' · ')));
    const avert = [];
    if (!client) avert.push('Aucun client sélectionné : ce sont les documents du profil connecté.');
    if (!/documents/i.test(location.pathname)) avert.push('Cette page n’est pas « Mes documents ».');
    if (document.querySelector('mat-paginator, .mat-mdc-paginator')) avert.push('Liste paginée : seule la page affichée est prise en compte.');
    const repliees = [...document.querySelectorAll('tr.mat-table-item-row:not(.no-expand)')].length;
    if (repliees) avert.push(`${repliees} ligne(s) dépliable(s) : déplie-les puis « Actualiser la liste » pour inclure leurs pièces.`);
    avertBox.replaceChildren(...avert.map(a => h('div', null, '• ' + a)));
    avertBox.style.display = avert.length ? '' : 'none';
    filtre = { annees: new Set(listeAnnees()), du: '', au: '' };
    inDu.value = inAu.value = '';
    rafraichir();
  };

  const journalTexte = () => ['N°\tDate\tDocument\tFichier\tStatut']
    .concat(docs.filter(d => d.statut).map(d => [d.i + 1, isoVersFR(d.date), d.titre, d.fichier, d.statut === 'ok' ? 'téléchargé' : d.detail].join('\t')))
    .join('\n');
  const diagTexte = () => JSON.stringify({
    script: VERSION, page: location.pathname,
    ongletAuChargement: diag.ongletAuChargement, ongletActuel: ongletActif(), nbDocuments: docs.length,
    menus: [...diag.menus], methodes: diag.methodes,
    erreurs: Object.entries(diag.erreurs.reduce((o, e) => ((o[e] = (o[e] || 0) + 1), o), {})).map(([e, n]) => `${n} × ${e}`),
    ouvertures: [...diag.ouvertures],
    pendantTelechargement: [...diag.reseau].slice(0, 40),
    appels: [...diag.appels.values()].slice(-40),
    rapide: { actif: CONFIG.MODE_RAPIDE, entetes: rapide.entetes ? Object.keys(rapide.entetes) : 'pas encore vus',
              listes: [...rapide.listes].map(([o, l]) => `${o} : ${l.length} document(s)`) },
    sondes: diag.sondes || 'non faites',
    // Adresses appelées par MyMinfin depuis l'ouverture de la page (sans numéros ni valeurs)
    api: [...new Set(performance.getEntriesByType('resource')
      .filter(e => /xmlhttprequest|fetch/.test(e.initiatorType)).map(e => masquerURL(e.name)).filter(Boolean))].slice(-60),
  }, null, 2);
  const copier = async texte => {
    try { await navigator.clipboard.writeText(texte); return true; } catch (e) { /* repli */ }
    const ta = document.createElement('textarea');
    ta.value = texte; ta.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.appendChild(ta); ta.select();
    let ok = false; try { ok = document.execCommand('copy'); } catch (e) { /* ignore */ }
    ta.remove(); return ok;
  };
  const boutonCopie = (b, f) => { const lib = b.textContent; b.onclick = async () => { const ok = await copier(f()); b.textContent = ok ? 'Copié ✓' : 'Voir console'; if (!ok) console.log(f()); setTimeout(() => (b.textContent = lib), 1600); }; };
  boutonCopie(bJournal, journalTexte);
  lienDiag.onclick = async () => {
    if (lienDiag.dataset.occupe) return;
    lienDiag.dataset.occupe = '1';
    lienDiag.textContent = 'Diagnostic en cours…';
    if (!diag.sondes) diag.sondes = await sonder();
    const texte = diagTexte();
    const ok = await copier(texte);
    if (!ok) console.log(texte);
    lienDiag.textContent = ok ? 'Diagnostic copié ✓' : 'Diagnostic : voir console';
    setTimeout(() => { lienDiag.textContent = 'Copier le diagnostic'; delete lienDiag.dataset.occupe; }, 2000);
  };

  lienTous.onclick = () => { if (etat.enCours) return; visibles().forEach(d => { d.coche = true; d.cb.checked = true; }); majCompteur(); };
  lienAucun.onclick = () => { if (etat.enCours) return; visibles().forEach(d => { d.coche = false; d.cb.checked = false; }); majCompteur(); };
  lienMaj.onclick = () => { if (!etat.enCours) charger(); };
  // Changement d'onglet ou de filtre MyMinfin : la liste du panneau se met à jour seule
  // (relecture quand la liste affichée change et reste stable, ou quand MyMinfin a rechargé sa liste)
  const surveillance = setInterval(() => {
    if (!host.isConnected) { clearInterval(surveillance); arreterReseau(); return; }
    if (etat.enCours) return;
    const sig = signatureDOM();
    if (sig === signatureListe && rapide.listeRecue <= listeVue) { signatureVue = ''; return; }
    if (sig === signatureVue) charger(); else signatureVue = sig;
  }, 700);
  x.onclick = () => { arret = true; if (!etat.enCours) { clearInterval(surveillance); desinstaller(); arreterReseau(); host.remove(); } };

  // ------------------------------------------------------------------- ZIP
  const TABLE_CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  const crc32 = u8 => { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = TABLE_CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const aujourdhuiISO = () => { const n = new Date(); return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, '0')}-${String(n.getDate()).padStart(2, '0')}`; };
  const dateDOS = iso => {
    const [y, m, d] = (iso || aujourdhuiISO()).split('-').map(Number);
    return { date: (((y - 1980) & 0x7f) << 9) | (m << 5) | d, heure: 12 << 11 };
  };
  // ZIP sans compression (les PDF le sont déjà), noms en UTF-8, date du document comme date du fichier
  const creerZip = async fichiers => {
    const enc = new TextEncoder();
    const parts = [], central = [];
    let offset = 0;
    for (const f of fichiers) {
      const data = new Uint8Array(await f.blob.arrayBuffer());
      const nom = enc.encode(f.nom);
      const crc = crc32(data);
      const { date, heure } = dateDOS(f.date);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
      lh.setUint16(10, heure, true); lh.setUint16(12, date, true); lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, nom.length, true); lh.setUint16(28, 0, true);
      parts.push(lh, nom, f.blob);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
      ch.setUint16(10, 0, true); ch.setUint16(12, heure, true); ch.setUint16(14, date, true); ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, nom.length, true);
      ch.setUint16(30, 0, true); ch.setUint16(32, 0, true); ch.setUint16(34, 0, true); ch.setUint16(36, 0, true);
      ch.setUint32(38, 0, true); ch.setUint32(42, offset, true);
      central.push(ch, nom);
      offset += 30 + nom.length + data.length;
      if (offset > 0xFFFFFFFF) throw new Error('ZIP trop volumineux (plus de 4 Go) : télécharger par année');
    }
    const taille = central.reduce((t, p) => t + p.byteLength, 0);
    const fin = new DataView(new ArrayBuffer(22));
    fin.setUint32(0, 0x06054b50, true); fin.setUint16(8, fichiers.length, true); fin.setUint16(10, fichiers.length, true);
    fin.setUint32(12, taille, true); fin.setUint32(16, offset, true);
    return new Blob([...parts, ...central, fin], { type: 'application/zip' });
  };
  const nomDuZip = () => {
    const onglet = ongletActif();
    let suffixe = onglet && !/^mes documents$/i.test(simple(onglet)) ? ` (${onglet})` : '';
    if (!listeAnnees().every(a => filtre.annees.has(a))) suffixe += ' ' + [...filtre.annees].sort().join('-');
    if (filtre.du || filtre.au) suffixe += ` du ${filtre.du || 'début'} au ${filtre.au || aujourdhuiISO()}`;
    return nettoyerNom(`${client || 'MyMinfin'} - documents MyMinfin${suffixe} - ${aujourdhuiISO()}`) + '.zip';
  };
  const listeCSV = lot => {
    const q = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
    return '﻿' + [['Date', 'Document', 'Fichier dans le ZIP', 'Statut'].map(q).join(';')]
      .concat(lot.map(d => [isoVersFR(d.date), d.titre, d.statut === 'ok' ? d.fichier : '',
                            d.statut === 'ok' ? 'OK' : (d.detail || 'non traité')].map(q).join(';')))
      .join('\r\n');
  };

  // ------------------------------------------------------------ lancement
  let boutonActif = null;
  const lancer = async (mode, bouton) => {
    const aFaire = visibles().filter(d => d.coche);
    if (!aFaire.length) return;
    etat.enCours = true; arret = false; boutonActif = bouton;
    inDu.disabled = inAu.disabled = true;
    bouton.textContent = 'Arrêter'; bouton.className = 'b stop';
    (bouton === bZip ? bFichiers : bZip).style.display = 'none';
    bJournal.style.display = 'none';
    aFaire.forEach(d => { d.fichier = ''; statut(d, '', ''); });
    installer();
    const fichiers = [], annexes = [];
    let ok = 0, ko = 0, infoZip = '', rapides = 0;
    const pageDepart = location.pathname;
    // Fichier reçu en mode rapide : rangé pour le ZIP ou enregistré tout de suite
    const recevoir = (d, blob, nomServeur, suffixe = '', ordre2 = 0) => {
      const nom = (mode === 'zip' ? d.nomZip : d.nom) + suffixe + extension(nomServeur, blob.type);
      if (mode === 'zip') fichiers.push({ nom, blob, date: d.date, ordre: d.i, ordre2 });
      else enregistrerBlob(blob, nom);
      return nom;
    };
    const resultat = (d, r) => {
      if (r.ok) { ok++; d.fichier = r.fichier; statut(d, 'ok', r.fichier); }
      else { ko++; statut(d, 'ko', r.raison); }
    };
    try {
      let file = [...aFaire];
      // 1. Mode rapide
      if (CONFIG.MODE_RAPIDE && !arret) {
        prog.textContent = 'Préparation du mode rapide…';
        try {
          if (await apprendreEntetes(file[0])) {
            prog.textContent = 'Lecture de la liste MyMinfin…';
            const assoc = associer(await listePourOnglet(docs), file);
            const lot = file.filter(d => assoc.has(d));
            const secours = [];
            let fait = 0;
            const travailleur = async () => {
              while (lot.length && !arret) {
                const d = lot.shift();
                statut(d, 'run', 'mode rapide');
                try {
                  const doc = assoc.get(d);
                  const f = await telechargerDirect(doc);
                  const nom = recevoir(d, f.blob, f.nomServeur);
                  // Pièces jointes du document
                  for (const [k, enfant] of (Array.isArray(doc.children) ? doc.children : []).entries()) {
                    if (!enfant || !enfant.documentKey) continue;
                    const titreE = enfant.documentTitle || `annexe ${k + 1}`;
                    try {
                      const fe = await telechargerDirect(enfant);
                      annexes.push({ date: d.date, titre: `${d.titre} — ${titreE}`, statut: 'ok',
                                     fichier: recevoir(d, fe.blob, fe.nomServeur, ' - ' + nettoyerNom(titreE), k + 1) });
                    } catch (e) { annexes.push({ date: d.date, titre: `${d.titre} — ${titreE}`, statut: 'ko', detail: 'pièce jointe : ' + e.message }); }
                  }
                  rapides++;
                  resultat(d, { ok: true, fichier: nom });
                } catch (e) {
                  if (/session/i.test(e.message)) { arret = true; diag.erreurs.push('mode rapide : ' + e.message); }
                  else diag.erreurs.push('mode rapide : ' + e.message + ' (repris par le menu)');
                  statut(d, '', '');
                  secours.push(d);
                }
                fait++;
                prog.textContent = `Mode rapide : ${fait} / ${assoc.size}`;
                if (mode === 'fichiers') await pause(CONFIG.PAUSE_RAPIDE_FICHIERS_MS);
              }
            };
            await Promise.all(Array.from({ length: mode === 'zip' ? CONFIG.PARALLELE : 1 }, travailleur));
            if (rapides) diag.methodes.rapide = (diag.methodes.rapide || 0) + rapides;
            // Restent : documents absents de la liste, échecs du mode rapide (et non traités si arrêt)
            file = file.filter(d => !assoc.has(d)).concat(secours, lot).sort((x, y) => x.i - y.i);
          }
        } catch (e) { diag.erreurs.push('mode rapide indisponible : ' + e.message); }
      }
      // 2. Menu « ⋮ », un document à la fois
      const t1 = Date.now();
      for (const [k, d] of file.entries()) {
        if (arret) break;
        if (d.statut === 'ok') continue;
        const reste = k ? Math.ceil(((Date.now() - t1) / k) * (file.length - k) / 60000) : null;
        prog.textContent = `${k + 1} / ${file.length} — ${d.titre}` + (reste ? ` (≈ ${reste} min restante${reste > 1 ? 's' : ''})` : '');
        statut(d, 'run');
        let r;
        try {
          r = await traiter(d, mode);
          // Second essai automatique : la première ouverture d'un menu échoue parfois (actions encore en chargement)
          if (!r.ok && !arret && !/^MyMinfin (ouvre|a ouvert)|liste modifi/.test(r.raison)) {
            statut(d, 'run', '2e essai');
            await pause(800);
            r = await traiter(d, mode);
            if (r.ok) diag.methodes['réussi au 2e essai'] = (diag.methodes['réussi au 2e essai'] || 0) + 1;
            else r.raison += ' (2 essais)';
          }
        } catch (e) { r = { ok: false, raison: 'erreur : ' + e.message }; diag.erreurs.push(r.raison); }
        if (r.ok && mode === 'zip') fichiers.push({ nom: r.fichier, blob: r.blob, date: d.date, ordre: d.i, ordre2: 0 });
        resultat(d, r);
        if (location.pathname !== pageDepart) arret = true;   // MyMinfin a quitté la liste : on s'arrête
        if (k < file.length - 1 && !arret) await pause(mode === 'zip' ? CONFIG.PAUSE_MS_ZIP : CONFIG.PAUSE_MS);
      }
    } finally {
      desinstaller();
      cap = null;
      if (mode === 'zip' && fichiers.length) {
        prog.textContent = `Création du ZIP (${fichiers.length} fichier(s))…`;
        fichiers.sort((x, y) => x.ordre - y.ordre || x.ordre2 - y.ordre2);
        try {
          const csv = new Blob([listeCSV(aFaire.concat(annexes))], { type: 'text/csv' });
          const zip = await creerZip(fichiers.concat([{ nom: '_Liste des documents.csv', blob: csv, date: '' }]));
          const nom = nomDuZip();
          enregistrerBlob(zip, nom);
          infoZip = ` — ZIP « ${nom} » (${(zip.size / 1048576).toFixed(1).replace('.', ',')} Mo)`;
        } catch (e) {
          infoZip = ' — ZIP impossible : ' + e.message;
          diag.erreurs.push('zip : ' + e.message);
        }
      }
      etat.enCours = false; boutonActif = null;
      inDu.disabled = inAu.disabled = false;
      bZip.className = 'b p'; bFichiers.className = 'b';
      bZip.style.display = bFichiers.style.display = '';
      majCompteur();
      bJournal.style.display = '';
      prog.textContent = (arret ? 'Arrêté. ' : 'Terminé. ') +
        `${ok} ${mode === 'zip' ? 'document(s) mis dans le ZIP' : 'téléchargé(s)'}` +
        (rapides ? ` (dont ${rapides} en mode rapide)` : '') +
        (annexes.length ? `, ${annexes.filter(x => x.statut === 'ok').length} pièce(s) jointe(s)` : '') +
        (ko ? `, ${ko} en erreur (survole ✖ pour le détail)` : '') + infoZip + '.';
      console.log(`MyMinfin documents : ${ok} récupéré(s), ${ko} en erreur.${infoZip}`);
      if (ko) console.warn('Diagnostic (sans données personnelles) : lien « Copier le diagnostic » du panneau.');
    }
  };
  const surClic = (mode, bouton) => () => {
    if (!etat.enCours) lancer(mode, bouton);
    else if (boutonActif === bouton) { arret = true; bouton.textContent = 'Arrêt après ce document…'; }
  };
  bZip.onclick = surClic('zip', bZip);
  bFichiers.onclick = surClic('fichiers', bFichiers);

  ecouterReseau();
  charger();
  console.log(`%cMoore - Outils · MyMinfin documents ${VERSION}`, 'font-weight:bold');
  console.log(`${docs.length} document(s) trouvé(s). Choisis-les dans le panneau en haut à droite.`);
}

  // --------------------------------------------------------------- routeur
  const pause = ms => new Promise(r => setTimeout(r, ms));
  const attendre = async (test, ms) => {
    const fin = Date.now() + ms;
    for (;;) { const v = test(); if (v) return v; if (Date.now() > fin) return null; await pause(150); }
  };
  const txt = el => (el ? String(el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim() : '');

  if (!/(^|\.)minfin\.fgov\.be$/i.test(location.hostname)) {
    alert('Moore - MyMinfin : ouvrez MyMinfin (connecté, client sélectionné) avant de cliquer sur le favori.');
    return;
  }

  const PAGES = {
    paiements: {
      chemin: '/pages/private/payments',
      nom: 'Mes paiements',
      // Blocs de la page présents (ouverts ou fermés : l'outil ouvre lui-même ceux dont il a besoin)
      pret: () => document.querySelector('#tax-balance .fragment-header, #provisions .fragment-header, #available-amount'),
      outil: outilCompteProvisions,
    },
    documents: {
      chemin: '/pages/private/documents',
      nom: 'Mes documents',
      pret: () => document.querySelector('mat-tab-nav-panel table tbody tr'),
      outil: outilDocuments,
    },
  };
  const surPage = k => location.pathname.includes(PAGES[k].chemin);

  // Petit message pendant que MyMinfin finit de charger la page
  const bulle = texte => {
    const d = document.createElement('div');
    d.textContent = texte;
    d.style.cssText = 'position:fixed;top:16px;right:16px;z-index:2147483647;background:#16324f;color:#fff;' +
      'font:600 13px/1.4 system-ui,Segoe UI,Arial,sans-serif;padding:10px 14px;border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,.25)';
    document.body.appendChild(d);
    return d;
  };

  const lancer = async k => {
    if (k === 'documents' && window.__mmfDocs && window.__mmfDocs.enCours) {
      alert('Moore - MyMinfin : un téléchargement est déjà en cours dans cet onglet.');
      return;
    }
    const p = PAGES[k];
    if (!p.pret()) {
      const b = bulle(`Moore - MyMinfin : chargement de « ${p.nom} »…`);
      await attendre(p.pret, 12000);
      b.remove();
    }
    await p.outil();
  };

  // Fenêtre de choix (ailleurs dans MyMinfin)
  const choisirOutil = () => new Promise(resolve => {
    const ident = document.querySelector('#user-info-details .myminfin-details-identifier');
    const client = txt(ident && ident.querySelector(':scope > span'));
    const bce = (txt(ident && ident.querySelector(':scope > div')).match(/[01]\d{3}\.?\d{3}\.?\d{3}/) || [''])[0];
    const host = document.createElement('div');
    host.style.cssText = 'position:fixed;inset:0;z-index:2147483647';
    const root = host.attachShadow({ mode: 'open' });
    root.innerHTML = `<style>
      .bg{position:fixed;inset:0;background:rgba(15,23,42,.42);display:flex;align-items:center;justify-content:center;
          font:14px/1.45 Arial,Helvetica,sans-serif;color:#1D1E1C}
      .dlg{width:min(480px,calc(100vw - 32px));background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.28);border-top:5px solid #1C9CD9}
      header{display:flex;align-items:center;justify-content:space-between;padding:16px 22px 6px}
      h1{font-size:18px;margin:0;letter-spacing:.2px} h1 small{font-weight:400;color:#7F7F7F;font-size:12px;margin-left:6px}
      .x{border:0;background:none;font-size:22px;line-height:1;cursor:pointer;color:#595959;padding:4px 8px;border-radius:6px} .x:hover{background:#F4F5F5}
      .body{padding:2px 22px 18px} .intro{margin:0 0 4px;color:#595959;font-size:13px} .intro.warn{color:#B45309;font-weight:700}
      button.outil{display:block;width:100%;margin:10px 0 0;padding:13px 16px;text-align:left;background:#fff;color:#1D1E1C;border:1px solid #D9D9D9;
          border-left:5px solid #1C9CD9;border-radius:8px;font:700 15px/1.35 Arial,Helvetica,sans-serif;cursor:pointer}
      button.outil span{display:block;font-weight:400;font-size:12px;color:#595959;margin-top:2px}
      button.outil:hover,button.outil:focus-visible{background:#F2F9FD;border-color:#1C9CD9;outline:none}
      footer{padding:10px 22px 14px;color:#7F7F7F;font-size:11px;border-top:1px solid #F4F5F5}</style>
      <div class="bg"><div class="dlg" role="dialog" aria-modal="true" aria-labelledby="t">
        <header><h1 id="t">Moore - MyMinfin<small></small></h1><button class="x" title="Fermer" aria-label="Fermer">×</button></header>
        <div class="body"><p class="intro"></p>
          <button type="button" class="outil" data-k="paiements">Compte-provisions et dettes TVA<span>Mes paiements · lecture seule, export vers Excel</span></button>
          <button type="button" class="outil" data-k="documents">Télécharger les documents<span>Mes documents · sélection, un document à la fois, fichiers renommés</span></button>
        </div>
        <footer>Le favori ouvre la page MyMinfin choisie puis l’outil. Rien n’est modifié dans MyMinfin.</footer>
      </div></div>`;
    root.querySelector('small').textContent = 'v' + VERSION;
    const intro = root.querySelector('.intro');
    if (client) intro.textContent = `Client : ${client}${bce ? ' · BCE ' + bce : ''}`;
    else { intro.textContent = 'Aucun client sélectionné : choisissez d’abord le client (« Sélectionner un client »).'; intro.className = 'intro warn'; }
    const fin = k => { host.remove(); resolve(k); };
    root.querySelector('.x').onclick = () => fin(null);
    root.querySelector('.bg').addEventListener('click', e => { if (e.target.classList.contains('bg')) fin(null); });
    root.addEventListener('keydown', e => { if (e.key === 'Escape') fin(null); });
    root.querySelectorAll('.outil').forEach(b => (b.onclick = () => fin(b.dataset.k)));
    document.body.appendChild(host);
    root.querySelector('.outil').focus();
  });

  const ici = Object.keys(PAGES).find(surPage);
  if (ici) { await lancer(ici); return; }

  const k = await choisirOutil();
  if (!k) return;
  // Navigation interne de MyMinfin (garde la session et le client sélectionné)
  const lien = document.querySelector(`a.nav-link[href*="${PAGES[k].chemin}"]`) || document.querySelector(`a[href*="${PAGES[k].chemin}"]`);
  if (lien) {
    lien.click();
    if (await attendre(() => surPage(k), 6000)) { await lancer(k); return; }
  }
  alert(`Moore - MyMinfin : la page « ${PAGES[k].nom} » va s’ouvrir. Recliquez sur le favori quand elle est chargée.`);
  location.assign('/myminfin-web' + PAGES[k].chemin);
})().catch(e => alert('Moore - MyMinfin : ' + ((e && e.message) || e)));
