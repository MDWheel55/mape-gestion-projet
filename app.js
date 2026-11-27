(function(){
'use strict';

/* ============ état ============ */
var state = { objectives:[], projects:[], tasks:[] };
function normaliser(){
if(!Array.isArray(state.objectives)) state.objectives = [];
if(!Array.isArray(state.projects)) state.projects = [];
if(!Array.isArray(state.tasks)) state.tasks = [];
state.tasks.forEach(function(t){
  if(typeof t.prio!=='boolean') t.prio = (t.priority==='urgent'||t.priority==='important');
  if('priority' in t) delete t.priority;
  if(!Array.isArray(t.owners)) t.owners = t.owner ? [t.owner] : [];
  if(!Array.isArray(t.deps)) t.deps = [];
  if('owner' in t) delete t.owner;
});
/* Un prédécesseur dont la tâche a été supprimée reste écrit dans la colonne
   Predecesseurs : la pastille « ↳ 1 » continuait de s'afficher sans qu'aucune flèche ne
   soit traçable, et rien ne permettait de la faire disparaître. On nettoie à la lecture ;
   la correction part vers SharePoint au prochain enregistrement de la tâche. */
if(state.tasks.length){
  var connus={};
  state.tasks.forEach(function(t){ if(t.id) connus[t.id]=1; });
  state.tasks.forEach(function(t){
    var nets=t.deps.filter(function(d){ return connus[d]; });
    if(nets.length!==t.deps.length) t.deps=nets;
  });
}

}

var ui = {
  tab:'roadmap', zoom:'semaine', fObj:'',
  fOwners:[], fBuckets:[], prioOnly:false, hideDone:false, openMenu:null, showDeps:true, filtersOpen:false,
  editing:null, lotMenu:null,
  drawer:null, draft:null, confirmDel:false,
  collapsed:{}, save:'idle', readOnly:false, msg:'',
  pret:false
};

/* Lot présent d'office dans chaque projet, toujours affiché en dernier : c'est la zone
   de garage. Convention : une tâche du Backlog n'a ni date ni propriétaire. */
var BACKLOG='Backlog';
var bkVus={};   // lots Backlog déjà rencontrés : sert à n'appliquer le repli par défaut qu'une fois
var SLOTS=['s1','s2','s3','s4','s5','s6'];
var MONTHS=['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
var MON3=['jan','fév','mar','avr','mai','jun','jul','aoû','sep','oct','nov','déc'];
var PPD={jour:26,semaine:9,mois:3};

/* ============ utilitaires ============ */
function uid(p){ return p+'-'+Math.random().toString(36).slice(2,9); }
function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g,function(c){
  return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
function startOfDay(d){ return new Date(d.getFullYear(),d.getMonth(),d.getDate()); }
var TODAY = startOfDay(new Date());
function parseD(s){ if(!s) return null; var p=String(s).split('-'); if(p.length<3) return null;
  var d=new Date(+p[0],+p[1]-1,+p[2]); return isNaN(d.getTime())?null:d; }
function iso(d){ if(!d) return ''; var m=d.getMonth()+1,j=d.getDate();
  return d.getFullYear()+'-'+(m<10?'0':'')+m+'-'+(j<10?'0':'')+j; }
function addDays(d,n){ var x=new Date(d.getTime()); x.setDate(x.getDate()+n); return x; }
function diffDays(a,b){ return Math.round((startOfDay(b)-startOfDay(a))/86400000); }
function fmtShort(s){ var d=parseD(s); if(!d) return '—'; return d.getDate()+' '+MON3[d.getMonth()]; }
function fmtLong(s){ var d=parseD(s); if(!d) return '—'; return d.getDate()+' '+MONTHS[d.getMonth()]+' '+d.getFullYear(); }
function clamp(v,a,b){ return v<a?a:(v>b?b:v); }
function byId(arr,id){ for(var i=0;i<arr.length;i++) if(arr[i].id===id) return arr[i]; return null; }

/* ============ dérivations ============ */
function tasksOf(pid){ return state.tasks.filter(function(t){ return t.projectId===pid; }); }
function projectsOf(oid){ return state.projects.filter(function(p){ return p.objectiveId===oid; }); }
function span(items){
  var mn=null,mx=null;
  items.forEach(function(it){
    var a=parseD(it.start)||parseD(it.end), b=parseD(it.end)||a;
    if(a && (!mn||a<mn)) mn=a;
    if(b && (!mx||b>mx)) mx=b;
  });
  return mn?{start:mn,end:mx||mn}:null;
}
function projSpan(p){ return span(tasksOf(p.id)); }
function projProgress(p){
  // Le backlog n'est pas du travail engagé : il ne compte ni dans l'avancement du projet,
  // ni, par ricochet, dans celui de l'orientation.
  var ts=tasksOf(p.id).filter(function(t){ return t.lot!==BACKLOG; });
  if(!ts.length) return null;
  var w=0,s=0;
  ts.forEach(function(t){
    var a=parseD(t.start), b=parseD(t.end)||a;
    var d=(a&&b)?Math.max(1,diffDays(a,b)+1):1;
    w+=d; s+=d*(+t.progress||0);
  });
  return w? Math.round(s/w) : 0;
}
function objSpan(o){
  if(o.start) return {start:parseD(o.start), end:parseD(o.end)||parseD(o.start)};
  var sp=[]; projectsOf(o.id).forEach(function(p){ var s=projSpan(p); if(s) sp.push({start:iso(s.start),end:iso(s.end)}); });
  return span(sp);
}
function objProgress(o){
  var ps=projectsOf(o.id), vals=[];
  ps.forEach(function(p){ var v=projProgress(p); if(v!==null) vals.push(v); });
  if(!vals.length) return null;
  return Math.round(vals.reduce(function(a,b){return a+b;},0)/vals.length);
}
function objOf(p){ return p? byId(state.objectives,p.objectiveId) : null; }
function colorOf(o){ return 'var(--'+((o&&o.color)||'s1')+')'; }
function isLate(t){
  var e=parseD(t.end); return !!e && e<TODAY && (+t.progress||0)<100;
}
var BUCKETS=[['passe','Passé'],['semaine','Cette semaine'],['prochaine','Semaine prochaine'],['mois','Ce mois'],['sansdate','Sans date']];
/* Plus de colonne « Lot » : le lot se lit dans le regroupement et se change en glissant la
   ligne sur l'en-tête d'un lot, ou sur la ligne du projet pour l'en sortir. */
var COLS={name:'--w-name',own:'--w-own',due:'--w-due',pct:'--w-pct',note:'--w-note'};
var MINW={name:160,own:70,due:70,pct:44,note:80};
function savePrefs(){
  try{
    var cs=getComputedStyle(document.documentElement), w={};
    for(var k in COLS) w[k]=cs.getPropertyValue(COLS[k]).trim();
    localStorage.setItem('mape-prefs',JSON.stringify({
      tab:ui.tab, zoom:ui.zoom, fOwners:ui.fOwners, fBuckets:ui.fBuckets, prioOnly:ui.prioOnly,
      hideDone:ui.hideDone, fObj:ui.fObj, collapsed:ui.collapsed, cols:w, showDeps:ui.showDeps }));
  }catch(e){}
}
function loadPrefs(){
  try{
    var raw=localStorage.getItem('mape-prefs'); if(!raw) return;
    var o=JSON.parse(raw)||{};
    if(o.tab==='gantt'||o.tab==='roadmap') ui.tab=o.tab;
    if(o.zoom&&PPD[o.zoom]) ui.zoom=o.zoom;
    if(Array.isArray(o.fOwners)) ui.fOwners=o.fOwners;
    if(Array.isArray(o.fBuckets)) ui.fBuckets=o.fBuckets;
    ui.prioOnly=!!o.prioOnly; ui.hideDone=!!o.hideDone;
    if(typeof o.showDeps==='boolean') ui.showDeps=o.showDeps;
    if(typeof o.fObj==='string') ui.fObj=o.fObj;
    if(o.collapsed&&typeof o.collapsed==='object') ui.collapsed=o.collapsed;
    if(o.cols) for(var k in COLS){ if(o.cols[k]) document.documentElement.style.setProperty(COLS[k],o.cols[k]); }
  }catch(e){}
}
function weekRange(off){
  var dow=(TODAY.getDay()+6)%7, mon=addDays(TODAY,-dow+(off||0)*7);
  return {a:mon,b:addDays(mon,6)};
}
function endOfMonth(){ return new Date(TODAY.getFullYear(),TODAY.getMonth()+1,0); }
function bucketsOf(t){
  var e=parseD(t.end);
  if(!e) return ['sansdate'];
  var out=[], w0=weekRange(0), w1=weekRange(1);
  if(e<TODAY) out.push('passe');
  if(e>=TODAY && e<=w0.b) out.push('semaine');
  if(e>=w1.a && e<=w1.b) out.push('prochaine');
  if(e>=TODAY && e<=endOfMonth()) out.push('mois');
  return out;
}
function fOwnersActif(){
  // Cocher TOUS les propriétaires revient à n'en filtrer aucun : sinon l'application se
  // croit filtrée alors qu'elle affiche tout.
  return ui.fOwners.length>0 && ui.fOwners.length<ownerOptions().length;
}
function anyTaskFilter(){ return !!(ui.hideDone||ui.prioOnly||fOwnersActif()||ui.fBuckets.length); }
function taskMatches(t){
  if(ui.hideDone && (+t.progress||0)>=100) return false;
  if(fOwnersActif()){
    var ls=t.owners||[], ok=false;
    if(!ls.length) ok = ui.fOwners.indexOf('')>=0;
    else for(var j=0;j<ls.length;j++) if(ui.fOwners.indexOf(ls[j])>=0){ ok=true; break; }
    if(!ok) return false;
  }
  if(ui.prioOnly && !t.prio) return false;
  if(ui.fBuckets.length){
    var bs=bucketsOf(t), hit=false;
    for(var i=0;i<bs.length;i++) if(ui.fBuckets.indexOf(bs[i])>=0){ hit=true; break; }
    if(!hit) return false;
  }
  return true;
}
function toggleIn(arr,v){ var i=arr.indexOf(v); if(i<0) arr.push(v); else arr.splice(i,1); }
/* Toutes les clés de repli : un projet, puis ses lots (Backlog compris). */
function clesArbo(){
  var ks=[];
  state.projects.forEach(function(p){
    ks.push('p:'+p.id);
    lotsDuProjet(p.id).forEach(function(l){ ks.push('l:'+p.id+'|'+l); });
  });
  return ks;
}
/* Changer un filtre déplie tout, une seule fois. Ce n'est pas un mode : juste après,
   les chevrons et le bouton « Tout replier » redeviennent pleinement utilisables. */
function deplierTout(){
  ui.collapsed={};
  clesArbo().forEach(function(k){ bkVus[k]=1; });
}
function toutReplie(){
  var ks=clesArbo();
  return ks.length>0 && ks.every(function(k){ return ui.collapsed[k]===true; });
}
function basculerArbo(){
  var ks=clesArbo(), replier=!toutReplie();
  ui.collapsed={};
  if(replier) ks.forEach(function(k){ ui.collapsed[k]=true; });
  render();
}
function resetFilters(){ ui.fOwners=[]; ui.fBuckets=[]; ui.prioOnly=false; ui.hideDone=false; ui.fObj=''; }
function owners(){
  var set={};
  state.tasks.forEach(function(t){ (t.owners||[]).forEach(function(o){ if(o) set[o]=1; }); });
  state.projects.concat(state.objectives).forEach(function(x){ if(x.owner) set[x.owner]=1; });
  return Object.keys(set).sort(function(a,b){ return a.localeCompare(b,'fr'); });
}
function ownersLabel(t){
  var l=t.owners||[];
  if(!l.length) return '—';
  if(l.length===1) return esc(l[0]);
  return esc(l[0])+' <span style="color:var(--ink-3)">+'+(l.length-1)+'</span>';
}

/* ============ domaine temporel ============ */
function domainOf(items,padBefore,padAfter){
  var sp=span(items);
  var a,b;
  if(sp){ a=addDays(sp.start,-(padBefore||7)); b=addDays(sp.end,padAfter||14); }
  else { a=addDays(TODAY,-14); b=addDays(TODAY,105); }
  if(a>TODAY) a=addDays(TODAY,-7);
  if(b<TODAY) b=addDays(TODAY,21);
  a=new Date(a.getFullYear(),a.getMonth(),1);
  b=new Date(b.getFullYear(),b.getMonth()+1,0);
  return {start:a,end:b,days:diffDays(a,b)+1};
}
function monthsIn(dom){
  var out=[], cur=new Date(dom.start.getFullYear(),dom.start.getMonth(),1);
  while(cur<=dom.end){
    var last=new Date(cur.getFullYear(),cur.getMonth()+1,0);
    var s=cur<dom.start?dom.start:cur, e=last>dom.end?dom.end:last;
    out.push({y:cur.getFullYear(),m:cur.getMonth(),off:diffDays(dom.start,s),len:diffDays(s,e)+1});
    cur=new Date(cur.getFullYear(),cur.getMonth()+1,1);
  }
  return out;
}

/* ============ rendu : barres ============ */
var view={ppd:1,domStart:null};
function taskById(id){ return byId(state.tasks,id); }
function wouldCycle(fromId,depId){
  var seen={}, stack=[depId];
  while(stack.length){
    var id=stack.pop();
    if(id===fromId) return true;
    if(seen[id]) continue; seen[id]=1;
    var t=taskById(id); if(t) (t.deps||[]).forEach(function(d){ stack.push(d); });
  }
  return false;
}
/* L'ordre des tâches est manuel : il vit dans la colonne Rang, pas dans les dates. Changer
   une date ne doit pas faire sauter une ligne ailleurs dans la liste — on perd le fil de ce
   qu'on regardait. Les rangs sont espacés de 10 pour qu'un déplacement n'écrive qu'une seule
   ligne : la tâche déplacée prend le milieu entre ses deux nouveaux voisins. */
function cmpTaches(x,y){
  var rx=(x.rang==null)?Infinity:+x.rang, ry=(y.rang==null)?Infinity:+y.rang;
  if(rx!==ry) return rx<ry?-1:1;
  var dx=x.end||x.start||'9999', dy=y.end||y.start||'9999';
  if(dx!==dy) return dx<dy?-1:1;
  return String(x.name||'').localeCompare(String(y.name||''),'fr');
}
/* Première ouverture après la mise en place des rangs : on fige l'ordre affiché jusque-là,
   c'est-à-dire l'ordre par date. Rien ne bouge à l'écran, mais à partir de là plus rien ne
   bougera tout seul. Une seule fois par projet — ensuite au moins une tâche porte un rang. */
function figerOrdreInitial(){
  if(ui.readOnly) return;
  // Sans la colonne Rang, l'ordre ne survit pas au rechargement : inutile d'écrire des
  // dizaines de lignes qui échoueront toutes. On le dit une fois dans la console.
  var vues=(SP.colonnes? SP.colonnes('taches') : null);
  if(vues && vues.length && !vues.some(function(x){ return /^rang$/i.test(String(x)); })){
    console.warn('Liste "03-Taches" : colonne Rang absente. '+
      'L\'ordre des tâches restera celui des dates tant qu\'elle n\'est pas créée (type Nombre).');
    return;
  }
  state.projects.forEach(function(p){
    var ts=tasksOf(p.id);
    if(!ts.length || ts.some(function(t){ return t.rang!=null; })) return;
    appliquerOrdre(p.id, ordreAffichage(p.id), true);
  });
}
function minRang(liste){
  var m=Infinity;
  (liste||[]).forEach(function(t){ if(t.rang!=null && +t.rang<m) m=+t.rang; });
  return m;
}
function parLotDe(pid){
  var o={};
  tasksOf(pid).forEach(function(t){ if(t.lot) (o[t.lot]=o[t.lot]||[]).push(t); });
  return o;
}
/* L'ordre des lots se lit dans celui des tâches : un lot vient là où commence sa première
   tâche. Pas de colonne de plus à tenir, et déplacer un lot revient à renuméroter ses
   tâches. Le Backlog suit la même règle, à une nuance près : tant qu'on ne l'a pas déplacé
   — et toujours s'il est vide, puisqu'il n'a alors aucune tâche où écrire sa place — il
   passe en dernier. C'est le départage à égalité, pas une position imposée. */
function lotsOrdonnes(pid,parLot){
  parLot=parLot||parLotDe(pid);
  var noms=Object.keys(parLot);
  noms.sort(function(a,b){
    var ra=minRang(parLot[a]), rb=minRang(parLot[b]);
    if(ra!==rb) return ra<rb?-1:1;
    if(a===BACKLOG) return 1;
    if(b===BACKLOG) return -1;
    return a.localeCompare(b,'fr');
  });
  return noms;
}
/* Les tâches d'un projet dans l'ordre où elles s'affichent : celles sans lot d'abord, puis
   chaque lot. « ordreLots » permet d'imposer une suite de lots — c'est ce qui sert quand on
   déplace quelque chose, pour que le reste ne bouge pas. */
function ordreAffichage(pid,ordreLots){
  var parLot=parLotDe(pid);
  var noms=(ordreLots||lotsOrdonnes(pid,parLot)).filter(function(n){ return parLot[n]; });
  Object.keys(parLot).forEach(function(n){ if(noms.indexOf(n)<0) noms.push(n); });
  var out=tasksOf(pid).filter(function(t){ return !t.lot; }).sort(cmpTaches);
  noms.forEach(function(n){
    parLot[n].slice().sort(cmpTaches).forEach(function(t){ out.push(t); });
  });
  return out;
}
/* Un seul geste d'écriture : on renumérote le projet de 10 en 10 dans l'ordre voulu et on
   n'envoie que les lignes dont le rang a vraiment changé. Les projets font quelques tâches,
   c'est plus sûr que de calculer un rang intercalaire qui finit par manquer de place. */
function appliquerOrdre(pid,ordre,sansRendu){
  var chg=[];
  ordre.forEach(function(x,i){
    var r=(i+1)*10;
    if(x.rang!==r){ x.rang=r; chg.push(['task',x]); }
  });
  if(!sansRendu) render();
  if(chg.length) enregistrerLot(chg);
}
/* Déplace une tâche : avant ou après une autre (elle reprend alors son lot), à la fin d'un
   lot quand on lâche sur son en-tête, ou hors de tout lot quand on lâche sur la ligne du
   projet. Jamais d'un projet à l'autre — ça, ça passe par la fiche. */
function deplacerTache(id, cible, apres, lotCible){
  var t=taskById(id); if(!t || ui.readOnly) return;
  var pid=t.projectId, lots=lotsOrdonnes(pid);
  var lot;
  if(cible){
    var c=taskById(cible);
    if(!c || c.projectId!==pid || c.id===t.id) return;
    lot=c.lot||'';
  } else {
    if(lotCible==null) return;
    lot=lotCible;
  }
  t.lot=lot;
  var ordre=ordreAffichage(pid,lots).filter(function(x){ return x.id!==t.id; });
  var pos;
  if(cible){
    pos=ordre.map(function(x){ return x.id; }).indexOf(cible);
    if(pos<0) return;
    if(apres) pos++;
  } else {
    var der=-1;
    ordre.forEach(function(x,i){ if((x.lot||'')===lot) der=i; });
    pos = der>=0 ? der+1 : (lot===''?0:ordre.length);
  }
  ordre.splice(pos,0,t);
  appliquerOrdre(pid,ordre);
}
/* Déplacer un lot, c'est déplacer le bloc de ses tâches — sa place n'est écrite nulle part
   ailleurs. Un lot vide n'a donc pas de place à écrire : le déplacer ne donne rien, et il
   revient en bas. C'est le cas du Backlog tant que rien n'y est garé. */
function deplacerLot(pid, src, cible, apres){
  if(ui.readOnly || !src || src===cible) return;
  var lots=lotsOrdonnes(pid);
  var i=lots.indexOf(src); if(i<0) return;
  lots.splice(i,1);
  var j=lots.indexOf(cible);
  if(j<0){ lots.push(cible); j=lots.length-1; }   // cible vide : elle n'est pas dans la liste
  lots.splice(apres? j+1 : j, 0, src);
  appliquerOrdre(pid, ordreAffichage(pid,lots));
}
/* Les liens se tracent à la souris : on part du point au bout d'une barre et on lâche sur
   une autre barre. Le sens est toujours fin → début, comme dans la fiche. */
function ajouterLien(fromId,toId){
  if(ui.readOnly || !fromId || !toId || fromId===toId) return;
  var t=taskById(toId), f=taskById(fromId);
  if(!t || !f) return;
  if((t.deps||[]).indexOf(fromId)>=0){ marquer('saved','Ce lien existe déjà.'); return; }
  if(wouldCycle(toId,fromId)){ marquer('error','Lien impossible : il ferait une boucle.'); return; }
  t.deps=(t.deps||[]).concat([fromId]);
  ui.showDeps=true;              // tracer un lien sans le voir n'a pas de sens
  render();
  enregistrer('task',t);
}
function retirerLien(fromId,toId){
  var t=taskById(toId); if(!t || ui.readOnly) return;
  var i=(t.deps||[]).indexOf(fromId); if(i<0) return;
  t.deps=t.deps.slice(); t.deps.splice(i,1);
  render();
  enregistrer('task',t);
}
function shiftTask(t,days){
  if(t.start) t.start=iso(addDays(parseD(t.start),days));
  if(t.end) t.end=iso(addDays(parseD(t.end),days));
}
function barGeom(dom,ppd,startS,endS){
  var a=parseD(startS)||parseD(endS), b=parseD(endS)||a;
  if(!a) return null;
  if(b<a) b=a;
  var i0=clamp(diffDays(dom.start,a),0,dom.days);
  var i1=clamp(diffDays(dom.start,b)+1,0,dom.days);
  return {left:i0*ppd, width:Math.max(6,(i1-i0)*ppd), point:iso(a)===iso(b)};
}
function gridlines(dom,ppd,months,showWeekends,leftCss){
  var h='<div class="gridlines" style="left:'+leftCss+';width:'+(dom.days*ppd)+'px">';
  if(showWeekends){
    for(var i=0;i<dom.days;i++){
      var d=addDays(dom.start,i), w=d.getDay();
      if(w===0||w===6) h+='<div class="gwe" style="left:'+(i*ppd)+'px;width:'+ppd+'px"></div>';
    }
  }
  months.forEach(function(m){
    h+='<div class="gline m" style="left:'+(m.off*ppd)+'px"></div>';
    if(ppd>=9){
      for(var d=7;d<m.len;d+=7) h+='<div class="gline" style="left:'+((m.off+d)*ppd)+'px"></div>';
    }
  });
  var ti=diffDays(dom.start,TODAY);
  if(ti>=0&&ti<=dom.days) h+='<div class="todayline" style="left:'+(ti*ppd+ppd/2)+'px"></div>';
  return h+'</div>';
}
function headBands(dom,ppd,months,zoom){
  var w=dom.days*ppd;
  var h='<div class="band" style="width:'+w+'px">';
  months.forEach(function(m){
    var lab=(m.len*ppd<58? MON3[m.m] : MONTHS[m.m]);
    if(m.m===0||m.off===0) lab+=' '+m.y;
    h+='<div class="mcell" style="width:'+(m.len*ppd)+'px">'+esc(lab)+'</div>';
  });
  h+='</div>';
  h+='<div class="band days" style="width:'+w+'px">';
  if(zoom==='jour'){
    for(var i=0;i<dom.days;i++){
      var d=addDays(dom.start,i), we=(d.getDay()===0||d.getDay()===6);
      h+='<div class="dcell'+(we?' we':'')+(d.getDate()===1?' mstart':'')+'" style="width:'+ppd+'px">'+d.getDate()+'</div>';
    }
  } else if(zoom==='semaine'){
    var i2=0;
    while(i2<dom.days){
      var dd=addDays(dom.start,i2);
      var span2=Math.min(7,dom.days-i2);
      h+='<div class="dcell" style="width:'+(span2*ppd)+'px">'+dd.getDate()+'</div>';
      i2+=7;
    }
  } else {
    months.forEach(function(m){
      var q='T'+(Math.floor(m.m/3)+1);
      h+='<div class="dcell" style="width:'+(m.len*ppd)+'px">'+q+'</div>';
    });
  }
  h+='</div>';
  var ti=diffDays(dom.start,TODAY);
  if(ti>=0&&ti<=dom.days) h+='<div class="todayflag" style="left:'+(ti*ppd+ppd/2)+'px"></div>';
  return h;
}

/* ============ vue 1 : feuille de route ============ */
function viewRoadmap(){
  if(!state.objectives.length){
    return '<div class="empty"><h2>Aucune orientation pour l’instant</h2>'+
      '<p>Les orientations stratégiques sont le niveau macro : ce que l’entreprise vise. Chacune s’atteint par un ou plusieurs projets, planifiés ensuite dans l’onglet Projets et tâches.</p>'+
      '<div class="acts"><button class="btn primary" data-act="new-obj">Ajouter une orientation</button>'+
      '</div></div>';
  }
  var items=[];
  state.objectives.forEach(function(o){ var s=objSpan(o); if(s) items.push({start:iso(s.start),end:iso(s.end)}); });
  var dom=domainOf(items,10,20), ppd=PPD.mois, months=monthsIn(dom), W=dom.days*ppd;
  view.ppd=ppd; view.domStart=dom.start;
  var LW=344;

  var h='<div class="rm-intro"><h2>Orientations stratégiques</h2>'+
    '<p>Le niveau macro : les orientations stratégiques, leur horizon et les projets qui les portent. Le détail des tâches vit dans l’onglet suivant.</p></div>';

  h+='<div class="scroller compact"><div class="grid">';
  h+='<div class="tl-head"><div class="row rowh" style="border-bottom:1px solid var(--line-2)">'+
     '<div class="lft" style="width:'+LW+'px"><div class="c oname-c">Orientation et projets porteurs</div></div>'+
     '<div class="tl" style="width:'+W+'px;position:relative">'+headBands(dom,ppd,months,'mois')+'</div></div></div>';

  h+='<div style="position:relative">'+gridlines(dom,ppd,months,false,LW+'px');
  state.objectives.forEach(function(o){
    var col=colorOf(o), sp=objSpan(o), pr=objProgress(o);
    var open=ui.collapsed['o:'+o.id]!==true;
    var g=sp?barGeom(dom,ppd,iso(sp.start),iso(sp.end)):null;
    h+='<div class="row rowo" style="--bc:'+col+'">'+
       '<div class="lft" style="width:'+LW+'px">'+
         '<div class="orail"></div>'+
         '<div class="c oname-c" style="flex:1;min-width:0;flex-direction:column;align-items:flex-start;gap:0">'+
           '<div style="display:flex;align-items:center;gap:6px;width:100%">'+
             '<button class="twist" data-act="toggle" data-key="o:'+o.id+'" aria-expanded="'+open+'" aria-label="Replier">'+chevron()+'</button>'+
             '<span class="oname" data-act="edit-obj" data-id="'+o.id+'" style="cursor:pointer;flex:1;min-width:0">'+esc(o.name||'Orientation')+'</span>'+
           '</div>'+
           '<div class="ometa" style="padding-left:24px">'+
             (o.owner?esc(o.owner)+' · ':'')+
             (sp? fmtShort(iso(sp.start))+' → '+fmtShort(iso(sp.end)) : 'aucune date')+
             (pr!==null?' · '+pr+'%':'')+
           '</div>'+
         '</div>'+
       '</div>'+
       '<div class="tl" style="width:'+W+'px">'+
         (g?'<div class="bar obj" data-drag="obj" data-act="edit-obj" data-id="'+o.id+'" style="left:'+g.left+'px;width:'+g.width+'px;--pw:'+(pr||0)+'%" title="'+esc(o.name)+' — glisser pour déplacer, bords pour étirer"><span class="fill"></span><span class="h l" data-h="l"></span><span class="h r" data-h="r"></span></div>':'')+
       '</div></div>';
    if(open){
      var ps=projectsOf(o.id);
      if(!ps.length){
        h+='<div class="row rowpr"><div class="lft" style="width:'+LW+'px">'+
           '<div class="c" style="padding-left:38px;color:var(--ink-3);font-size:11.5px">Aucun projet rattaché</div></div>'+
           '<div class="tl" style="width:'+W+'px"></div></div>';
      }
      ps.forEach(function(p){
        var s=projSpan(p), gg=s?barGeom(dom,ppd,iso(s.start),iso(s.end)):null, pp=projProgress(p);
        h+='<div class="row rowpr" style="--bc:'+col+'">'+
           '<div class="lft" style="width:'+LW+'px">'+
             '<div class="c" style="padding-left:38px;flex:1;min-width:0">'+
               '<span class="prname" data-act="edit-proj" data-id="'+p.id+'" style="cursor:pointer">'+esc(p.name||'Projet')+'</span>'+
             '</div>'+
             '<div class="c pct">'+(pp===null?'':pp+'%')+'</div>'+
           '</div>'+
           '<div class="tl" style="width:'+W+'px">'+
             (gg?'<div class="bar proj thin" style="left:'+gg.left+'px;width:'+gg.width+'px" title="'+esc(p.name)+'"></div>':'')+
           '</div></div>';
      });
    }
  });
  h+='</div></div></div>';

  /* cartes */
  h+='<div class="sec-title">Les orientations en détail</div><div class="cards">';
  state.objectives.forEach(function(o){
    var col=colorOf(o), pr=objProgress(o), ps=projectsOf(o.id), sp=objSpan(o);
    h+='<div class="card" style="--bc:'+col+'" data-act="edit-obj" data-id="'+o.id+'" tabindex="0" role="button">'+
       '<h3>'+esc(o.name||'Orientation')+'</h3>'+
       (o.statement?'<p class="stmt">'+esc(o.statement)+'</p>':'')+
       '<div class="pbar"><i style="width:'+(pr||0)+'%"></i></div>'+
       '<div class="meta"><span class="mono">'+(pr===null?'—':pr+'%')+'</span>'+
       '<span>·</span><span>'+(sp?fmtShort(iso(sp.start))+' → '+fmtShort(iso(sp.end)):'sans horizon')+'</span>'+
       (o.owner?'<span>·</span><span>'+esc(o.owner)+'</span>':'')+'</div>'+
       (ps.length?'<ul class="plist">'+ps.map(function(p){
          var pp=projProgress(p);
          return '<li><span class="pd" style="background:'+col+'"></span>'+esc(p.name)+
                 '<span class="mono" style="margin-left:auto;color:var(--ink-3)">'+(pp===null?'':pp+'%')+'</span></li>';
        }).join('')+'</ul>':'<p class="hint" style="margin-top:9px">Aucun projet rattaché</p>')+
       '</div>';
  });
  h+='</div>';
  return h;
}

/* ============ vue 2 : projets et tâches ============ */
function viewGantt(){
  if(!state.projects.length){
    return '<div class="empty"><h2>Aucun projet pour l’instant</h2>'+
      '<p>Un projet regroupe des tâches. Chaque tâche porte un propriétaire et une échéance — c’est ce qui alimente le Gantt et l’avancement de l’orientation.</p>'+
      '<div class="acts"><button class="btn primary" data-act="new-proj">Ajouter un projet</button>'+
      '</div></div>';
  }
  var ppd=PPD[ui.zoom], filt=anyTaskFilter();
  var projs=state.projects.filter(function(p){
    if(ui.fObj && p.objectiveId!==ui.fObj) return false;
    if(filt) return tasksOf(p.id).some(taskMatches);
    return true;
  });
  if(!projs.length){
    return '<div class="empty"><h2>Rien ne correspond aux filtres</h2>'+
      '<p>Élargis la sélection, ou repars des filtres vides.</p>'+
      '<div class="acts"><button class="btn primary" data-act="reset">Réinitialiser les filtres</button></div></div>';
  }
  var items=[], byProj={};
  projs.forEach(function(p){
    var ts=tasksOf(p.id).filter(taskMatches);
    ts.sort(cmpTaches);
    byProj[p.id]=ts;
    ts.forEach(function(t){ items.push({start:t.start,end:t.end}); });
    var sp=projSpan(p); if(sp) items.push({start:iso(sp.start),end:iso(sp.end)});
  });
  var dom=domainOf(items,7,14), months=monthsIn(dom), W=dom.days*ppd;
  view.ppd=ppd; view.domStart=dom.start;

  var h='<div class="scroller"><div class="grid gantt">';
  h+='<div class="tl-head"><div class="row rowh" style="border-bottom:1px solid var(--line-2)">'+
     '<div class="lft">'+
       '<div class="c name">Projet et tâches<span class="rz" data-col="name" title="Glisser pour élargir"></span></div>'+
       '<div class="c own">Propriétaires<span class="rz" data-col="own"></span></div>'+
       '<div class="c due">Échéance<span class="rz" data-col="due"></span></div>'+
       '<div class="c pct">Av.<span class="rz" data-col="pct"></span></div>'+
       '<div class="c note">Note<span class="rz" data-col="note"></span></div></div>'+
     '<div class="tl" style="width:'+W+'px;position:relative">'+headBands(dom,ppd,months,ui.zoom)+'</div></div></div>';
  h+='<div class="rowsbox" style="position:relative">'+gridlines(dom,ppd,months,ui.zoom==='jour','var(--lw)');

  var y=0, pos={};

  // Rendu d'une ligne de tâche, réutilisé qu'elle soit dans un lot ou non.
  function ligneTache(t,col,W,dom,ppd,dansLot){
    var g2=barGeom(dom,ppd,t.start,t.end), lt=isLate(t), pv=+t.progress||0;
    if(g2) pos[t.id]={x0:g2.left, x1:g2.left+g2.width, y:y+16, point:g2.point};
    var r='<div class="row rowt'+((ui.editing&&ui.editing.id===t.id)?' editrow':'')+'" '+
       'data-tid="'+t.id+'" data-tlot="'+esc(t.lot||'')+'" data-tpid="'+esc(t.projectId||'')+'" '+
       'style="--bc:'+col+'">'+
       '<div class="lft">'+
         '<div class="c name"><span class="indent"></span>'+(dansLot?'<span class="indent"></span>':'')+
           (ui.readOnly?'':'<span class="ogrip" draggable="true" title="Glisser pour changer l’ordre">⠿</span>')+
           '<span class="tname'+(t.prio?' on':'')+'" data-act="edit-task" data-id="'+t.id+'">'+
           (t.prio?'<span class="bang">!</span> ':'')+esc(t.name||'Tâche')+'</span>'+
           ((t.deps||[]).length?'<span class="chip" title="'+esc((t.deps||[]).map(function(d){ var x=taskById(d); return x?x.name:'?'; }).join(' · '))+'">↳ '+(t.deps||[]).length+'</span>':'')+
           (ui.readOnly?'':'<button class="addrow addlot" data-act="new-lot" data-t="'+t.id+'" '+
             'title="Créer un lot à partir de cette tâche">+ lot</button>')+
         '</div>'+
         '<div class="c own'+(enEdition(t.id,'owners')?' editing':'')+'">'+
           (enEdition(t.id,'owners') ? menuProprietaires(t)
             : '<span class="cellv'+((t.owners||[]).length?'':' vide')+'" data-act="cell" data-f="owners" data-id="'+t.id+'" '+
               'title="'+esc((t.owners||[]).join(', ')||'Cliquer pour assigner')+'">'+ownersLabel(t)+'</span>')+
         '</div>'+
         '<div class="c due'+(lt?' late':'')+(enEdition(t.id,'end')?' editing':'')+'">'+
           (enEdition(t.id,'end')
             ? '<input class="inline" type="date" data-inline="end" value="'+esc(t.end||'')+'">'
             : '<span class="cellv'+(t.end?'':' vide')+'" data-act="cell" data-f="end" data-id="'+t.id+'" '+
               'title="'+(t.end?esc(fmtLong(t.end)):'Cliquer pour dater')+'">'+(t.end?fmtShort(t.end):'—')+'</span>')+
         '</div>'+
         '<div class="c pct'+(enEdition(t.id,'progress')?' editing':'')+'">'+
           (enEdition(t.id,'progress')
             ? '<input class="inline" type="number" min="0" max="100" step="5" data-inline="progress" value="'+pv+'">'
             : '<span class="cellv" data-act="cell" data-f="progress" data-id="'+t.id+'" title="Cliquer pour modifier l’avancement">'+pv+'%</span>')+
         '</div>'+
         '<div class="c note" title="'+esc(t.note||'')+'">'+
           (t.note ? '<span class="cellv" data-act="edit-task" data-id="'+t.id+'">'+esc(shorten(t.note,60))+'</span>'
                   : '<span class="nvide">—</span>')+
         '</div>'+
       '</div>'+
       '<div class="tl" style="width:'+W+'px">'+
         (g2? (g2.point
            ? '<div class="milestone" data-drag="task" data-act="edit-task" data-id="'+t.id+'" style="left:'+(g2.left+ppd/2)+'px" title="'+esc(t.name)+' — '+fmtLong(t.end||t.start)+'">'+
              (ui.readOnly?'':'<span class="dep-out" data-depout="'+t.id+'" title="Glisser vers une autre tâche pour créer un lien fin → début"></span>')+'</div>'
            : '<div class="bar task'+(lt?' late':'')+(pv>=100?' done':'')+'" data-drag="task" data-act="edit-task" data-id="'+t.id+'" style="left:'+g2.left+'px;width:'+g2.width+'px;--pw:'+pv+'%" '+
              'title="'+esc(t.name)+' — '+fmtLong(t.start)+' → '+fmtLong(t.end)+' — '+pv+'%">'+
              '<span class="fill"></span><span class="h l" data-h="l"></span><span class="h r" data-h="r"></span>'+
              (ui.readOnly?'':'<span class="dep-out" data-depout="'+t.id+'" title="Glisser vers une autre tâche pour créer un lien fin → début"></span>')+'</div>')
           : '<span class="blabel" style="left:4px;color:var(--ink-3)">sans date</span>')+
       '</div></div>';
    y+=32;
    return r;
  }

  projs.forEach(function(p){
    var o=objOf(p), col=colorOf(o), sp=projSpan(p), pp=projProgress(p);
    var open=ui.collapsed['p:'+p.id]!==true;
    var ts=byProj[p.id];
    var g=sp?barGeom(dom,ppd,iso(sp.start),iso(sp.end)):null;
    var lateN=ts.filter(isLate).length, prioN=ts.filter(function(t){ return t.prio; }).length;
    h+='<div class="row rowp" data-droplot="" data-tpid="'+esc(p.id)+'" style="--bc:'+col+'">'+
       '<div class="lft">'+
         '<div class="c name">'+
           '<button class="twist" data-act="toggle" data-key="p:'+p.id+'" aria-expanded="'+open+'" aria-label="Replier le projet">'+chevron()+'</button>'+
           '<span class="pname" data-act="edit-proj" data-id="'+p.id+'" style="cursor:pointer">'+esc(p.name||'Projet')+'</span>'+
           (o?'<span class="chip chip-obj"><span class="swatch" style="background:'+col+'"></span>'+esc(shorten(o.name,16))+'</span>':'')+
           (prioN?'<span class="chip" style="background:var(--danger-soft);color:var(--danger)">'+prioN+' prioritaire'+(prioN>1?'s':'')+'</span>':'')+
           (lateN?'<span class="chip" style="background:var(--danger-soft);color:var(--danger)">'+lateN+' en retard</span>':'')+
           (ui.readOnly?'':'<button class="addrow" data-act="new-task" data-p="'+p.id+'" title="Ajouter une tâche dans ce projet">+</button>'+
             '')+
         '</div>'+
         '<div class="c own">'+esc(p.owner||'')+'</div>'+
         '<div class="c due">'+(sp?fmtShort(iso(sp.end)):'—')+'</div>'+
         '<div class="c pct">'+(pp===null?'':pp+'%')+'</div>'+
         '<div class="c note"></div>'+
       '</div>'+
       '<div class="tl" style="width:'+W+'px">'+
         (g?'<div class="bar proj" data-drag="proj" data-act="edit-proj" data-id="'+p.id+'" style="left:'+g.left+'px;width:'+g.width+'px" title="'+esc(p.name)+' — glisser pour décaler tout le projet"></div>':'')+
         (!ts.length?'<span class="blabel" style="left:4px">aucune tâche</span>':'')+
       '</div></div>';
    y+=38;
    if(!open) return;

    // Les tâches sans lot restent directement sous le projet ; les lots forment des groupes.
    var sansLot=ts.filter(function(t){ return !t.lot; });
    var parLot={};
    ts.forEach(function(t){ if(t.lot){ (parLot[t.lot]=parLot[t.lot]||[]).push(t); } });
    // Le Backlog existe même vide — sauf pendant un filtre, où un lot sans tâche
    // correspondante n'a rien à faire à l'écran. Vide, il n'a aucun rang : il tombe en
    // dernier tout seul.
    if(!parLot[BACKLOG] && !filt) parLot[BACKLOG]=[];
    var noms=lotsOrdonnes(p.id,parLot);

    sansLot.forEach(function(t){ h+=ligneTache(t,col,W,dom,ppd,false); });

    noms.forEach(function(nom){
      var groupe=parLot[nom];
      var estBacklog=(nom===BACKLOG);
      var cle='l:'+p.id+'|'+nom;
      // Le Backlog arrive replié la première fois qu'on le voit ; ensuite l'utilisateur
      // décide — y compris « Tout déplier », qui ne doit pas se faire annuler ici.
      if(estBacklog && !bkVus[cle]){
        bkVus[cle]=1;
        if(ui.collapsed[cle]===undefined) ui.collapsed[cle]=true;
      }
      var ouvert=ui.collapsed[cle]!==true;
      var spl=estBacklog?null:span(groupe.map(function(t){ return {start:t.start,end:t.end}; }));
      var gl=spl?barGeom(dom,ppd,iso(spl.start),iso(spl.end)):null;
      var menuOuvert=!!(ui.lotMenu && ui.lotMenu.pid===p.id && ui.lotMenu.lot===nom);
      h+='<div class="row rowlot'+(menuOuvert?' editrow':'')+(estBacklog?' rowbacklog':'')+'" '+
         'data-droplot="'+esc(nom)+'" data-tpid="'+esc(p.id)+'" style="--bc:'+col+'">'+
         '<div class="lft">'+
           '<div class="c name'+(menuOuvert?' editing':'')+'"><span class="indent"></span>'+
             '<button class="twist" data-act="toggle" data-key="'+esc(cle)+'" aria-expanded="'+ouvert+'" aria-label="Replier le lot">'+chevron()+'</button>'+
             (ui.readOnly?'':'<span class="ogrip" draggable="true" title="Glisser pour changer l’ordre des lots">⠿</span>')+
             '<span class="lotname">'+esc(nom)+'</span>'+
             (groupe.length?'<span class="chip">'+groupe.length+'</span>':'')+
             (estBacklog?'' : (menuOuvert
               ? '<div class="menu-wrap"><button class="lotdots" data-act="lot-menu" data-p="'+p.id+'" data-l="'+esc(nom)+'">⋯</button>'+
                 '<div class="menu menu-lot">'+
                   '<div class="menu-add"><input type="text" id="lotrename" value="'+esc(nom)+'">'+
                   '<button class="btn" data-act="lot-rename" data-p="'+p.id+'" data-l="'+esc(nom)+'">Renommer</button></div>'+
                   '<label style="cursor:pointer" data-act="lot-ungroup" data-p="'+p.id+'" data-l="'+esc(nom)+'">Dégrouper — retirer le lot de ces '+groupe.length+' tâches</label>'+
                 '</div></div>'
               : '<button class="lotdots" data-act="lot-menu" data-p="'+p.id+'" data-l="'+esc(nom)+'" title="Renommer ou dégrouper">⋯</button>'))+
             (ui.readOnly?'':'<button class="addrow" data-act="new-task" data-p="'+p.id+'" data-l="'+esc(nom)+'" title="'+(estBacklog?'Garer une idée dans le backlog de ce projet':'Ajouter une tâche dans ce lot')+'">+</button>')+
           '</div>'+
           '<div class="c own"></div>'+
           '<div class="c due">'+(estBacklog?'':(spl?fmtShort(iso(spl.end)):'—'))+'</div>'+
           '<div class="c pct">'+(estBacklog||!groupe.length?'':moyenneAvancement(groupe)+'%')+'</div><div class="c note"></div>'+
         '</div>'+
         '<div class="tl" style="width:'+W+'px">'+
           (gl?'<div class="bar lotbar" data-drag="lot" data-id="'+esc(p.id+'|'+nom)+'" style="left:'+gl.left+'px;width:'+gl.width+'px" title="'+esc(nom)+' — glisser pour décaler tout le lot"></div>':'')+
         '</div></div>';
      y+=28;
      if(ouvert && !groupe.length) h+='<div class="row rowt rowvide"><div class="lft"><div class="c name">'+
        '<span class="indent"></span><span class="indent"></span>'+
        '<span class="vide">Rien de garé ici. Le « + » ajoute une ligne sans date ni propriétaire.</span>'+
        '</div><div class="c own"></div><div class="c due"></div>'+
        '<div class="c pct"></div><div class="c note"></div></div>'+
        '<div class="tl" style="width:'+W+'px"></div></div>';
      if(ouvert && !groupe.length) y+=28;
      if(ouvert) groupe.forEach(function(t){ h+=ligneTache(t,col,W,dom,ppd,true); });
    });
  });

  if(ui.showDeps){
    var paths='';
    projs.forEach(function(p){
      byProj[p.id].forEach(function(t){
        (t.deps||[]).forEach(function(did){
          var A=pos[did], B=pos[t.id];
          if(!A||!B) return;
          var pt=taskById(did);
          var bad = pt && pt.end && t.start && parseD(pt.end) > parseD(t.start);
          var sx=A.point? A.x1+ppd/2 : A.x1, sy=A.y, tx=B.point? B.x0+ppd/2-6 : B.x0-2, ty=B.y, m=9;
          var d;
          if(tx > sx+2*m) d='M'+sx+' '+sy+' H'+(sx+m)+' V'+ty+' H'+tx;
          else d='M'+sx+' '+sy+' H'+(sx+m)+' V'+((sy+ty)/2)+' H'+(tx-m)+' V'+ty+' H'+tx;
          paths+='<path class="depline'+(bad?' bad':'')+'" d="'+d+'"/>'+
                 '<polygon class="dephead'+(bad?' bad':'')+'" points="'+tx+','+ty+' '+(tx-5)+','+(ty-3.5)+' '+(tx-5)+','+(ty+3.5)+'"/>'+
                 '<path class="dephit" d="'+d+'" data-act="dep-del" data-from="'+esc(did)+'" data-to="'+esc(t.id)+'">'+
                 '<title>Lien fin → début — cliquer pour le supprimer</title></path>';
        });
      });
    });
    if(paths) h+='<svg class="deps'+(ui.readOnly?'':' cliquable')+'" style="left:var(--lw);width:'+W+'px;height:'+y+'px">'+paths+'</svg>';
  }
  h+='</div></div></div>';
  return h;
}
function moyenneAvancement(liste){
  if(!liste.length) return 0;
  var s=0; liste.forEach(function(t){ s+=(+t.progress||0); });
  return Math.round(s/liste.length);
}
function lotsDuProjet(pid){
  var vus={};
  state.tasks.forEach(function(t){ if(t.projectId===pid && t.lot) vus[t.lot]=1; });
  delete vus[BACKLOG];
  return Object.keys(vus).sort(function(a,b){ return a.localeCompare(b,'fr'); }).concat([BACKLOG]);
}
/* Un lot n'existe que par ses tâches : le créer, c'est en sortir une d'où elle est. La
   tâche garde son rang, donc le lot apparaît là où elle était. L'en-tête s'ouvre aussitôt
   sur le champ de nom — « Nouveau lot » n'est qu'un provisoire. */
function nouveauLotDepuis(id){
  var t=taskById(id); if(!t || ui.readOnly) return;
  var pris=lotsDuProjet(t.projectId), nom='Nouveau lot', n=2;
  while(pris.indexOf(nom)>=0){ nom='Nouveau lot '+n; n++; }
  t.lot=nom;
  ui.lotMenu={pid:t.projectId, lot:nom};
  ui.collapsed['l:'+t.projectId+'|'+nom]=false;
  render();
  enregistrer('task',t);
  var ch=document.getElementById('lotrename');
  if(ch){ ch.focus({preventScroll:true}); ch.select(); }
}
function definirLot(id,nom){
  var t=taskById(id); if(!t) return;
  t.lot=nom||'';
  ui.editing=null;
  render();
  enregistrer('task',t);
}
function renommerLot(pid,ancien,nouveau){
  nouveau=(nouveau||'').trim();
  if(!nouveau || nouveau===ancien){ ui.lotMenu=null; render(); return; }
  var touchees=state.tasks.filter(function(t){ return t.projectId===pid && t.lot===ancien; });
  touchees.forEach(function(t){ t.lot=nouveau; });
  ui.lotMenu=null; render();
  enregistrerLot(touchees.map(function(t){ return ['task',t]; }));
}
function degrouperLot(pid,nom){
  var touchees=state.tasks.filter(function(t){ return t.projectId===pid && t.lot===nom; });
  touchees.forEach(function(t){ t.lot=''; });
  ui.lotMenu=null; render();
  enregistrerLot(touchees.map(function(t){ return ['task',t]; }));
}
function menuProprietaires(t){
  var liste=owners(), sel=t.owners||[];
  var h='<div class="menu-wrap"><span class="cellv">'+ownersLabel(t)+'</span><div class="menu menu-own">';
  if(!liste.length) h+='<div class="none">Personne d’enregistré pour l’instant</div>';
  liste.forEach(function(o){
    h+='<label><input type="checkbox" data-act="own-chk" data-id="'+t.id+'" data-v="'+esc(o)+'"'+
       (sel.indexOf(o)>=0?' checked':'')+'> '+esc(o)+'</label>';
  });
  h+='<div class="menu-add"><input type="text" id="ownnew" placeholder="Autre personne">'+
     '<button class="btn" data-act="own-add" data-id="'+t.id+'">Ajouter</button></div>';
  return h+'</div></div>';
}
function basculerProprietaire(id,nom){
  var t=taskById(id); if(!t) return;
  t.owners=(t.owners||[]).slice();
  toggleIn(t.owners,nom);
  render();
  enregistrer('task',t);
}
function enEdition(id,f){ return !!(ui.editing && ui.editing.id===id && ui.editing.f===f); }
function ouvrirCellule(id,f){
  ui.editing={id:id,f:f};
  render();
  var inp=document.querySelector('[data-inline]');
  if(inp){ inp.focus({preventScroll:true}); if(inp.select && inp.type!=='date') inp.select(); }
}
function validerCellule(el){
  if(!ui.editing) return;
  var t=taskById(ui.editing.id), f=ui.editing.f;
  ui.editing=null;
  if(!t){ render(); return; }
  if(f==='owners'){
    t.owners = el.value.split(/[;,]/).map(function(x){ return x.trim(); }).filter(Boolean);
  } else if(f==='progress'){
    var v=parseInt(el.value,10); if(isNaN(v)) v=0;
    t.progress = Math.max(0, Math.min(100, v));
  } else if(f==='end'){
    t.end = el.value || '';
    // Une échéance antérieure au début rendrait la barre impossible à dessiner.
    if(t.start && t.end && t.end < t.start) t.start = t.end;
  }
  render();
  enregistrer('task',t);
}
function shorten(s,n){ s=String(s||''); return s.length>n? s.slice(0,n-1)+'…' : s; }
function chevron(){
  return '<svg width="11" height="11" viewBox="0 0 12 12" fill="none" aria-hidden="true">'+
    '<path d="M2.5 4.5L6 8l3.5-3.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';
}

/* ============ tiroir d’édition ============ */
function viewDrawer(){
  if(!ui.drawer) return '';
  var d=ui.draft, k=ui.drawer, isNew=!d.id;
  if(k==='data'){
    var dump=JSON.stringify({objectives:state.objectives,projects:state.projects,tasks:state.tasks},null,2);
    return '<div class="scrim" data-act="close"></div><aside class="drawer" role="dialog" aria-label="Donnees">'+
      '<header><h2>Données</h2><button class="btn ghost" data-act="close">Fermer</button></header><div class="body">'+
      '<p class="hint">Copie ce texte pour garder une sauvegarde de l’état actuel.</p>'+
      '<textarea id="datajson" rows="16" spellcheck="false" style="font-family:\'IBM Plex Mono\',monospace;font-size:11.5px">'+esc(dump)+'</textarea>'+
      '<p class="hint" id="datamsg"></p>'+
      '</div><footer><button class="btn" data-act="copy-data">Copier</button>'+
      '<span class="hint" style="margin:0 0 0 6px">Sauvegarde seulement — la source de vérité est SharePoint.</span>'+
      '<span class="spacer"></span></footer></aside>';
  }
  var title = k==='obj'? (isNew?'Nouvelle orientation stratégique':'Orientation stratégique')
            : k==='proj'? (isNew?'Nouveau projet':'Projet')
            : (isNew?'Nouvelle tâche':'Tâche');
  var h='<div class="scrim" data-act="close"></div><aside class="drawer" role="dialog" aria-label="'+title+'">'+
    '<header><h2>'+title+'</h2><button class="btn ghost" data-act="close" aria-label="Fermer">Fermer</button></header><div class="body">';

  h+='<label class="lb"><span>Nom</span><input type="text" data-f="name" value="'+esc(d.name||'')+'" placeholder="'+
     (k==='task'?'Commencer par un verbe : Paramétrer les articles…':(k==='proj'?'Nom du projet':'Nom de l’orientation'))+'"></label>';

  if(k==='obj'){
    h+='<label class="lb"><span>Énoncé</span><textarea data-f="statement" rows="3" placeholder="Ce que l’on vise, en une phrase">'+esc(d.statement||'')+'</textarea></label>';
    h+='<label class="lb"><span>Responsable</span><input type="text" data-f="owner" list="owners" value="'+esc(d.owner||'')+'"></label>';
    h+='<div class="two">'+
       '<label class="lb"><span>Début (optionnel)</span><input type="date" data-f="start" value="'+esc(d.start||'')+'"></label>'+
       '<label class="lb"><span>Fin (optionnel)</span><input type="date" data-f="end" value="'+esc(d.end||'')+'"></label></div>';
    h+='<div class="lb"><span>Couleur</span><div class="swatches">'+SLOTS.map(function(s){
        return '<button class="sw" style="background:var(--'+s+')" data-act="color" data-c="'+s+'" aria-pressed="'+((d.color||'s1')===s)+'" aria-label="Couleur '+s+'"></button>';
      }).join('')+'</div></div>';
    h+='<p class="hint">Sans dates, l’horizon se calcule tout seul à partir des tâches des projets rattachés.</p>';
  }
  if(k==='proj'){
    h+='<label class="lb"><span>Orientation rattachée</span><select data-f="objectiveId">'+
       '<option value="">— aucun —</option>'+
       state.objectives.map(function(o){ return '<option value="'+o.id+'"'+(d.objectiveId===o.id?' selected':'')+'>'+esc(o.name)+'</option>'; }).join('')+
       '</select></label>';
    h+='<label class="lb"><span>Responsable</span><input type="text" data-f="owner" list="owners" value="'+esc(d.owner||'')+'"></label>';
    h+='<p class="hint">Les dates et l’avancement du projet se déduisent de ses tâches.</p>';
  }
  if(k==='task'){
    h+='<label class="lb"><span>Projet</span><select data-f="projectId">'+
       state.projects.map(function(p){ return '<option value="'+p.id+'"'+(d.projectId===p.id?' selected':'')+'>'+esc(p.name)+'</option>'; }).join('')+
       '</select></label>';
    var lotsDispo=lotsDuProjet(d.projectId||'');
    if(d.lot && lotsDispo.indexOf(d.lot)<0) lotsDispo=[d.lot].concat(lotsDispo);
    h+='<div class="lb"><span>Lot</span>'+
       '<select data-f="lot"><option value="">— aucun —</option>'+
         lotsDispo.map(function(l){ return '<option value="'+esc(l)+'"'+(d.lot===l?' selected':'')+'>'+esc(l)+'</option>'; }).join('')+
       '</select>'+
       '<input type="text" id="lotdrawer" placeholder="ou créer un nouveau lot" style="margin-top:6px">'+
       '<p class="hint">Regroupe les tâches d’un même mandat à l’intérieur du projet. Un lot n’a ni propriétaire ni échéance : ce sont ses tâches qui les portent.</p>'+
       '</div>';
    var ow=d.owners||[];
    h+='<div class="lb"><span>Propriétaires</span>'+
       (ow.length?'<div class="ownchips">'+ow.map(function(o,i){
          return '<span class="ownchip">'+esc(o)+'<button class="x" data-act="del-owner" data-i="'+i+'" aria-label="Retirer '+esc(o)+'">×</button></span>';
        }).join('')+'</div>':'<p class="hint" style="margin:0">Personne d’assigné — la tâche n’apparaîtra dans la liste de personne.</p>')+
       '<div style="display:flex;gap:6px;margin-top:6px">'+
         '<input type="text" id="ownerinput" list="owners" placeholder="Ajouter une personne" style="flex:1">'+
         '<button class="btn" data-act="add-owner">Ajouter</button>'+
       '</div>'+
       (function(){
          var rest=owners().filter(function(o){ return ow.indexOf(o)<0; });
          return rest.length? '<div class="ownquick">'+rest.map(function(o){
            return '<button class="tog" data-act="quick-owner" data-o="'+esc(o)+'">+ '+esc(o)+'</button>'; }).join('')+'</div>' : '';
       })()+
       '</div>';
    h+='<div class="two">'+
       '<label class="lb"><span>Début</span><input type="date" data-f="start" value="'+esc(d.start||'')+'"></label>'+
       '<label class="lb"><span>Échéance</span><input type="date" data-f="end" value="'+esc(d.end||'')+'"></label></div>';
    h+='<label class="field" style="gap:8px"><input type="checkbox" data-f="prio"'+(d.prio?' checked':'')+'> Tâche prioritaire</label>';
    h+='<label class="lb"><span>Avancement — '+(+d.progress||0)+'%</span>'+
       '<input type="range" min="0" max="100" step="5" data-f="progress" value="'+(+d.progress||0)+'"></label>';
    var dp=d.deps||[];
    var cand=state.tasks.filter(function(x){ return x.id && x.id!==d.id && !(d.id&&wouldCycle(d.id,x.id)) && dp.indexOf(x.id)<0; });
    h+='<div class="lb"><span>Prédécesseurs (fin → début)</span>'+
       (dp.length?'<div class="ownchips">'+dp.map(function(id,i){
          var x=taskById(id);
          return '<span class="ownchip">'+esc(x?shorten(x.name,26):'tâche supprimée')+'<button class="x" data-act="del-dep" data-i="'+i+'" aria-label="Retirer">×</button></span>';
        }).join('')+'</div>':'<p class="hint" style="margin:0">Aucun prédécesseur : la tâche peut démarrer quand tu veux.</p>')+
       (cand.length?'<div style="display:flex;gap:6px;margin-top:6px">'+
         '<select id="depinput" style="flex:1;min-width:0">'+
           state.projects.map(function(pr){
             var list=cand.filter(function(x){ return x.projectId===pr.id; });
             return list.length? '<optgroup label="'+esc(pr.name)+'">'+list.map(function(x){
               return '<option value="'+x.id+'">'+esc(shorten(x.name,40))+'</option>'; }).join('')+'</optgroup>' : '';
           }).join('')+
         '</select><button class="btn" data-act="add-dep">Ajouter</button></div>':'')+
       '<p class="hint">La tâche devrait démarrer après la fin de ses prédécesseurs. Un lien non respecté s’affiche en rouge dans le Gantt — rien n’est replanifié automatiquement.</p>'+
       '</div>';
    h+='<label class="lb"><span>Note</span><textarea data-f="note" rows="3">'+esc(d.note||'')+'</textarea></label>';
    h+='<p class="hint">Même date de début et d’échéance : la tâche s’affiche comme un jalon.</p>';
  }
  h+='</div><footer>'+
     '<button class="btn primary" data-act="apply">Enregistrer</button>'+
     '<button class="btn ghost" data-act="close">Annuler</button>'+
     '<span class="spacer"></span>'+
     (isNew?'':'<button class="btn '+(ui.confirmDel?'danger':'ghost')+'" data-act="del">'+(ui.confirmDel?'Confirmer la suppression':'Supprimer')+'</button>')+
     '</footer></aside>';
  return h;
}

/* ============ chrome ============ */
function menuBox(id,label,opts,sel){
  var h='<div class="menu-wrap"><button class="btn" data-act="menu" data-m="'+id+'" aria-expanded="'+(ui.openMenu===id)+'">'+
    esc(label)+(sel.length?' <span class="cnt">'+sel.length+'</span>':'')+' ▾</button>';
  if(ui.openMenu===id){
    h+='<div class="menu">';
    if(!opts.length) h+='<div class="none">Rien à filtrer pour l’instant</div>';
    if(opts.length){
      var tout = sel.length>=opts.length;
      h+='<button class="menu-all" data-act="chk-all" data-g="'+id+'" data-v="'+(tout?'0':'1')+'">'+
         (tout?'Tout décocher':'Tout cocher')+'</button>';
    }
    opts.forEach(function(o){
      h+='<label><input type="checkbox" data-act="chk" data-g="'+id+'" data-v="'+esc(o[0])+'"'+
         (sel.indexOf(o[0])>=0?' checked':'')+'> '+esc(o[1])+'</label>';
    });
    h+='</div>';
  }
  return h+'</div>';
}
function ownerOptions(){
  var os=owners().map(function(o){ return [o,o]; });
  var orphan=state.tasks.some(function(t){ return !(t.owners||[]).length; });
  if(orphan) os.push(['','Sans propriétaire']);
  return os;
}
function isSmall(){ return window.innerWidth<=760; }
function viewFilters(){
  var vis=0;
  state.projects.forEach(function(p){
    if(ui.fObj && p.objectiveId!==ui.fObj) return;
    tasksOf(p.id).forEach(function(t){ if(taskMatches(t)) vis++; });
  });
  var h='<div class="filterbar">';
  h+='<span class="flabel">Échéance</span><div class="togs">'+
     BUCKETS.map(function(bk){
       return '<button class="tog" data-act="tog" data-b="'+bk[0]+'" aria-pressed="'+(ui.fBuckets.indexOf(bk[0])>=0)+'">'+bk[1]+'</button>';
     }).join('')+'</div>';
  h+=menuBox('owner','Propriétaire',ownerOptions(),ui.fOwners);
  h+='<button class="tog" data-act="togprio" aria-pressed="'+ui.prioOnly+'">! Prioritaires</button>';
  h+='<button class="tog" data-act="togdeps" aria-pressed="'+ui.showDeps+'" title="Afficher les liens fin → début">↳ Liens</button>';
  h+='<label class="field">Orientation <select data-act="fobj"><option value="">Toutes</option>'+
     state.objectives.map(function(o){ return '<option value="'+o.id+'"'+(ui.fObj===o.id?' selected':'')+'>'+esc(shorten(o.name,28))+'</option>'; }).join('')+
     '</select></label>';
  h+='<label class="field"><input type="checkbox" data-act="hidedone"'+(ui.hideDone?' checked':'')+'> Masquer les terminées</label>';
  if(anyTaskFilter()||ui.fObj) h+='<button class="btn ghost" data-act="reset">Réinitialiser</button>';
  h+='<span class="spacer"></span><span class="field mono">'+vis+' tâche'+(vis>1?'s':'')+' affichée'+(vis>1?'s':'')+'</span>';
  return h+'</div>';
}
function viewBar(){
  var nO=state.objectives.length, nP=state.projects.length, nT=state.tasks.length;
  var lateN=state.tasks.filter(isLate).length;
  var sv = ui.readOnly? {c:'', t:'Lecture seule'}
        : ui.save==='saving'? {c:'busy', t:'Enregistrement…'}
        : ui.save==='error'? {c:'err', t:'Non enregistré'}
        : ui.save==='saved'? {c:'ok', t:'Enregistré'}
        : {c:'', t:'Modifications enregistrées dans la page'};
  var h='<div class="appbar">'+
    '<div class="brand"><h1>Mapé</h1><span class="sub">Gestion de projet · orientations · tâches</span></div>'+
    '<div class="tabs" role="tablist">'+
      '<button class="tab" role="tab" aria-selected="'+(ui.tab==='roadmap')+'" data-act="tab" data-t="roadmap">Feuille de route <span class="n">'+nO+'</span></button>'+
      '<button class="tab" role="tab" aria-selected="'+(ui.tab==='gantt')+'" data-act="tab" data-t="gantt">Projets et tâches <span class="n">'+nP+'/'+nT+'</span></button>'+
    '</div>'+
    '<div class="savebox"><span class="dot '+sv.c+'"></span>'+esc(sv.t)+'</div>'+
  '</div>';

  h+='<div class="toolbar">';
  if(ui.tab==='gantt'){
    h+='<div class="seg" role="group" aria-label="Échelle">'+
       ['jour','semaine','mois'].map(function(z){
         return '<button data-act="zoom" data-z="'+z+'" aria-pressed="'+(ui.zoom===z)+'">'+z.charAt(0).toUpperCase()+z.slice(1)+'</button>';
       }).join('')+'</div>';
    h+='<button class="btn ghost" data-act="today">Aujourd’hui</button>';
    h+='<button class="btn ghost" data-act="arbo">'+(toutReplie()?'Tout déplier':'Tout replier')+'</button>';
    if(isSmall()){
      var nf=ui.fOwners.length+ui.fBuckets.length+(ui.prioOnly?1:0)+(ui.hideDone?1:0)+(ui.fObj?1:0);
      h+='<button class="btn'+(ui.filtersOpen?' primary':'')+'" data-act="togfilters">Filtres'+(nf?' <span class="cnt">'+nf+'</span>':'')+'</button>';
    }
    h+='<span class="spacer"></span>';
    if(state.objectives.length) h+='<div class="legend">'+state.objectives.map(function(o){
       return '<span class="it"><span class="swatch" style="background:'+colorOf(o)+'"></span>'+esc(shorten(o.name,18))+'</span>'; }).join('')+'</div>';
    if(!ui.readOnly){
      h+='<button class="btn ghost" data-act="reload" title="Relire les listes SharePoint">Recharger</button>';
      h+='<button class="btn ghost" data-act="data">Données</button>';
      h+='<button class="btn" data-act="new-proj">+ Projet</button>';
      h+='<button class="btn primary" data-act="new-task"'+(state.projects.length?'':' disabled')+'>+ Tâche</button>';
    }
  } else {
    h+='<span class="field">'+nO+' orientation'+(nO>1?'s':'')+' · '+nP+' projet'+(nP>1?'s':'')+' · '+nT+' tâche'+(nT>1?'s':'')+
       (lateN?' · <span class="late">'+lateN+' en retard</span>':'')+'</span>';
    h+='<span class="spacer"></span>';
    if(!ui.readOnly){
      h+='<button class="btn ghost" data-act="data">Données</button>';
      h+='<button class="btn" data-act="new-proj">+ Projet</button>';
      h+='<button class="btn primary" data-act="new-obj">+ Orientation</button>';
    }
  }
  h+='</div>';
  if(ui.tab==='gantt' && state.projects.length && (!isSmall() || ui.filtersOpen)) h+=viewFilters();
  if(ui.msg) h+='<div class="banner">'+esc(ui.msg)+
     (ui.needLogin?' <button class="btn" data-act="login" style="margin-left:8px">Se connecter</button>':'')+
     '</div>';
  return h;
}

/* ============ rendu ============ */
function contenuOnglet(){
  if(ui.tab==='roadmap') return viewRoadmap();
  return viewGantt();
}
/* Remplacer le DOM fait perdre le focus, ce qui declenche blur puis focusout sur le champ
   qui l'avait — et donc un enregistrement du texte a un index qui vient peut-etre de
   changer (une ligne deplacee, une ligne supprimee). Le texte etait deja ecrit par
   l'evenement change, alors pendant un rendu on ignore ces enregistrements tardifs. */
var enRendu=false;
function render(){
  enRendu=true;
  try{ rendreVraiment(); } finally { enRendu=false; }
}
function rendreVraiment(){
  var root=document.getElementById('root');
  /* Le rendu reconstruit tout le DOM : sans ça, le Gantt revient au début du domaine
     à chaque enregistrement. On relève la position avant, on la repose après. */
  var anc=root.querySelector('.scroller');
  var sx=anc?anc.scrollLeft:null, sy=anc?anc.scrollTop:null;
  var vol=root.querySelector('.pane');
  var vy=vol?vol.scrollTop:null, vx=vol?vol.scrollLeft:null;
  root.innerHTML = viewBar()+
    '<div class="pane'+(ui.tab==='gantt'?'':' scroll')+'">'+contenuOnglet()+'</div>'+
    viewDrawer()+
    '<datalist id="owners">'+owners().map(function(o){ return '<option value="'+esc(o)+'"></option>'; }).join('')+'</datalist>';
  var bar=root.querySelector('.appbar');
  if(bar) root.style.setProperty('--bartop', bar.offsetHeight+'px');
  if(sx!=null){
    var nou=root.querySelector('.scroller');
    if(nou){ nou.scrollLeft=sx; nou.scrollTop=sy; }
  }
  if(vy!=null){
    var nv=root.querySelector('.pane');
    if(nv){ nv.scrollTop=vy; nv.scrollLeft=vx; }
  }
  savePrefs();
  paintSave();
}
function scrollToToday(){
  requestAnimationFrame(function(){
    var sc=document.querySelector('.scroller'); if(!sc) return;
    var t=sc.querySelector('.todayline'); if(!t) return;
    var r=t.getBoundingClientRect(), rs=sc.getBoundingClientRect();
    var lf=sc.querySelector('.lft');
    var lw=lf? lf.getBoundingClientRect().width : 0;
    var margin=lw + Math.max(40,(sc.clientWidth-lw)*0.3);
    sc.scrollLeft = Math.max(0, sc.scrollLeft + (r.left-rs.left) - margin);
  });
}

/* ============ actions ============ */
function openDrawer(kind,obj){
  ui.drawer=kind; ui.confirmDel=false; ui.openMenu=null;
  ui.draft = obj? JSON.parse(JSON.stringify(obj,function(k,v){ return k.slice(0,2)==='__'?undefined:v; })) : {};
  render();
  var f=document.querySelector('.drawer input,.drawer textarea,.drawer select'); if(f) f.focus({preventScroll:true});
}
function addOwner(val){
  var inp=document.getElementById('ownerinput');
  var v=(val!==undefined? val : (inp?inp.value:'')).trim();
  collectDraft();
  ui.draft.owners=(ui.draft.owners||[]).slice();
  if(v && ui.draft.owners.indexOf(v)<0) ui.draft.owners.push(v);
  render();
  var i2=document.getElementById('ownerinput'); if(i2){ i2.value=''; i2.focus({preventScroll:true}); }
}
function closeDrawer(){ ui.drawer=null; ui.draft=null; ui.confirmDel=false; render(); }
function collectDraft(){
  var d=ui.draft;
  document.querySelectorAll('.drawer [data-f]').forEach(function(el){
    var f=el.getAttribute('data-f');
    d[f] = (el.type==='checkbox') ? el.checked : ((f==='progress') ? +el.value : el.value);
  });
}
function applyDraft(){
  collectDraft();
  var ld=document.getElementById('lotdrawer');
  if(ld && ld.value.trim()) ui.draft.lot=ld.value.trim();
  var d=ui.draft, k=ui.drawer;
  if(!d.name || !d.name.trim()){ ui.msg='Donne un nom avant d’enregistrer.'; render(); return; }
  if(k==='task'){
    if(!d.projectId){ ui.msg='Choisis un projet pour cette tâche.'; render(); return; }
    if(d.start && d.end && d.end < d.start){ var s=d.start; d.start=d.end; d.end=s; }
  }
  ui.msg='';
  var list = k==='obj'? state.objectives : (k==='proj'? state.projects : state.tasks);
  if(d.id){
    var cur=byId(list,d.id);
    if(cur) for(var key in d){ if(key.slice(0,2)!=='__') cur[key]=d[key]; }
  } else {
    d.id = uid(k);
    if(k==='obj' && !d.color) d.color = SLOTS[state.objectives.length % SLOTS.length];
    list.push(d);
  }
  var _k=k, _o=d; closeDrawer(); enregistrer(_k,_o);
}
function deleteCurrent(){
  var d=ui.draft, k=ui.drawer;
  var detaches=[], tachesProjet=[];
  if(k==='obj'){
    state.objectives=state.objectives.filter(function(x){ return x.id!==d.id; });
    state.projects.forEach(function(p){ if(p.objectiveId===d.id){ p.objectiveId=''; detaches.push(['proj',p]); } });
  } else if(k==='proj'){
    state.projects=state.projects.filter(function(x){ return x.id!==d.id; });
    // Les tâches partent avec le projet, dans SharePoint aussi : sinon elles y restent orphelines.
    tachesProjet=state.tasks.filter(function(t){ return t.projectId===d.id; });
    state.tasks=state.tasks.filter(function(t){ return t.projectId!==d.id; });
  } else {
    state.tasks=state.tasks.filter(function(x){ return x.id!==d.id; });
  }
  /* Supprimer une tâche laisse des prédécesseurs qui pointent dans le vide. On les retire
     tout de suite, sinon la pastille « ↳ » survit à la tâche qu'elle désignait. */
  var partis={}; partis[d.id]=1; tachesProjet.forEach(function(t){ partis[t.id]=1; });
  var orphelins=[];
  state.tasks.forEach(function(t){
    if(!(t.deps||[]).some(function(x){ return partis[x]; })) return;
    t.deps=t.deps.filter(function(x){ return !partis[x]; });
    orphelins.push(['task',t]);
  });
  closeDrawer(); supprimer(k,d);
  tachesProjet.forEach(function(t){ supprimer('task',t); });
  if(orphelins.length||detaches.length) enregistrerLot(orphelins.concat(detaches));
}

/* ============ persistance — listes SharePoint ============ */
var sync={etat:'idle', msg:''};
function marquer(e,m){ sync.etat=e; sync.msg=m||''; ui.save=(e==='busy'?'saving':e); paintSave(); }

async function chargerTout(){
  marquer('busy','Chargement…');
  try{
    var o = await SP.getItems('orientations');
    var p = await SP.getItems('projets');
    var t = await SP.getItems('taches');
    state.objectives = o.map(function(r){ return {
      spId:r.spId, id:r.OrientationId||('obj-'+r.spId), name:r.Title||'',
      statement:r.Enonce||'', owner:r.Responsable||'',
      start:SP.jour(r.DateDebut), end:SP.jour(r.DateFin), color:r.Couleur||'s1' }; });
    state.projects = p.map(function(r){ return {
      spId:r.spId, id:r.ProjetId||('proj-'+r.spId), name:r.Title||'',
      objectiveId:r.OrientationId||'', owner:r.Responsable||'' }; });
    state.tasks = t.map(function(r){ return {
      spId:r.spId, id:r.TacheId||('task-'+r.spId), projectId:r.ProjetId||'', name:r.Title||'',
      owners:SP.liste(r.Proprietaires), lot:r.Lot||'', start:SP.jour(r.DateDebut), end:SP.jour(r.DateFin),
      progress:+(r.Avancement||0), prio:(String(r.Prioritaire||'').toLowerCase()==='oui'),
      deps:SP.liste(r.Predecesseurs), note:r.Note||'',
      rang:(r.Rang===''||r.Rang==null)?null:+r.Rang }; });
    normaliser();
    figerOrdreInitial();
    marquer('saved','À jour');
  }catch(e){ marquer('error','Chargement impossible : '+e.message); }
}

function champsOrientation(o){ return {
  Title:o.name||'', OrientationId:o.id, Enonce:o.statement||'', Responsable:o.owner||'',
  DateDebut:SP.iso(o.start), DateFin:SP.iso(o.end), Couleur:o.color||'s1' }; }
function champsProjet(p){ return {
  Title:p.name||'', ProjetId:p.id, OrientationId:p.objectiveId||'', Responsable:p.owner||'' }; }
function champsTache(t){ return {
  Title:t.name||'', TacheId:t.id, ProjetId:t.projectId||'',
  Proprietaires:(t.owners||[]).join('; '), Lot:t.lot||'', DateDebut:SP.iso(t.start), DateFin:SP.iso(t.end),
  Avancement:+(t.progress||0), Prioritaire:t.prio?'Oui':'Non',
  Predecesseurs:(t.deps||[]).join('; '), Note:t.note||'',
  Rang:(t.rang==null?null:+t.rang) }; }

var CLES={obj:['orientations',champsOrientation,'objectives'],
          proj:['projets',champsProjet,'projects'],
          task:['taches',champsTache,'tasks']};

/* Les enregistrements d'un MÊME objet sont mis en file. Sans ça, deux champs remplis coup
   sur coup partaient tous les deux avant que le premier n'ait renvoyé son spId : SharePoint
   créait deux lignes, et au rechargement c'est la plus pauvre qui gagnait — d'où l'impression
   d'avoir tout perdu. La file garantit que la création se fait une seule fois. */
var enVol=0, avertiCol={};
var FACULTATIFS=[];
async function enregistrer(kind,obj){
  var c=CLES[kind]; if(!c) return;
  enVol++;
  marquer('saving','Enregistrement…');
  // Le .catch garde la file vivante : un échec ne doit pas condamner tous les
  // enregistrements suivants du même objet.
  obj.__file = Promise.resolve(obj.__file && obj.__file.then ? obj.__file : null).catch(function(){}).then(function(){
    return ecrireUn(c,obj);
  });
  var p=obj.__file;
  return p.then(function(){ if(--enVol<=0){ enVol=0; marquer('saved','Enregistré'); } },
                function(){ if(--enVol<0) enVol=0; });
}
async function ecrireUn(c,obj){
  // Les champs sont sérialisés AU MOMENT de l'écriture, donc l'envoi porte toujours
  // l'état le plus récent de l'objet, même si la saisie a continué entre-temps.
  var champs=c[1](obj);
  try{
    if(obj.spId){ await SP.updateItem(c[0],obj.spId,champs); }
    else { var r=await SP.addItem(c[0],champs); obj.spId=r.id; }
    /* Une colonne manquante dans la liste est une vraie erreur : on la nomme. */
    var perdus=SP.champsIgnores?SP.champsIgnores(c[0]):[];
    var graves=perdus.filter(function(x){ return FACULTATIFS.indexOf(x)<0; });
    if(graves.length){
      var figees=graves.filter(function(x){ return SP.colonneFigee && SP.colonneFigee(c[0],x); });
      var absentes=graves.filter(function(x){ return figees.indexOf(x)<0; });
      var msg='';
      if(figees.length) msg+=figees.join(', ')+' existe'+(figees.length>1?'nt':'')+
        ' mais n\'est pas modifiable par l\'application (colonne calculée ou système). ';
      if(absentes.length){
        var vues=SP.colonnes?SP.colonnes(c[0]):[];
        msg+='Colonne'+(absentes.length>1?'s':'')+' introuvable'+(absentes.length>1?'s':'')+
          ' : '+absentes.join(', ')+(vues.length?' — colonnes vues : '+vues.join(', '):'');
      }
      marquer('error',msg);
      throw new Error('colonnes manquantes');
    }
    if(perdus.length && !avertiCol[c[0]]){
      avertiCol[c[0]]=1;
      console.warn('Liste "'+c[0]+'" : colonne(s) '+perdus.join(', ')+' absente(s).');
    }
  }catch(e){
    console.error('Refus de SharePoint sur la liste "'+c[0]+'". Champs envoyés :', champs, e);
    if(sync.etat!=='error') marquer('error','Échec : '+e.message+' — détail dans la console (F12)');
    throw e;
  }
}
async function enregistrerLot(liste){
  marquer('saving','Enregistrement…');
  try{
    for(var i=0;i<liste.length;i++){
      var k=liste[i][0], o=liste[i][1], c=CLES[k];
      if(o.spId) await SP.updateItem(c[0],o.spId,c[1](o));
      else { var r=await SP.addItem(c[0],c[1](o)); o.spId=r.id; }
    }
    marquer('saved','Enregistré');
  }catch(e){ marquer('error','Échec : '+e.message); }
}
async function supprimer(kind,obj){
  var c=CLES[kind]; if(!c||!obj.spId) return;
  marquer('saving','Suppression…');
  try{ await SP.deleteItem(c[0],obj.spId); marquer('saved','Supprimé'); }
  catch(e){ marquer('error','Échec : '+e.message); }
}

function paintSave(){
  var box=document.querySelector('.savebox');
  if(!box) return;
  var cls = sync.etat==='saving'||sync.etat==='busy' ? 'busy'
          : sync.etat==='error' ? 'err'
          : sync.etat==='saved' ? 'ok' : '';
  var txt = sync.msg || 'Prêt';
  box.innerHTML='<span class="dot '+cls+'"></span>'+esc(txt)+
    (SP.utilisateur()? ' · '+esc(SP.utilisateur()) : '');
}

/* ============ événements ============ */
document.addEventListener('click',function(ev){
  if(suppressClick){ return; }
  var el=ev.target.closest('[data-act]');
  var dansMenu = ev.target.closest && ev.target.closest('.menu-wrap');
  if(ui.openMenu && !dansMenu){
    ui.openMenu=null; render();
    if(!el) return;
  }
  if(ui.editing && (ui.editing.f==='owners'||ui.editing.f==='lot') && !dansMenu){
    ui.editing=null; render();
    if(!el) return;
  }
  if(ui.lotMenu && !dansMenu){
    ui.lotMenu=null; render();
    if(!el) return;
  }
  if(!el) return;
  var a=el.getAttribute('data-act');
  if(a==='menu'){ var m=el.getAttribute('data-m'); ui.openMenu=(ui.openMenu===m?null:m); render(); return; }
  if(a==='tog'){ toggleIn(ui.fBuckets,el.getAttribute('data-b')); deplierTout(); render(); return; }
  if(a==='chk-all'){
    // Une liste vide veut dire « aucun filtre », donc tout coché et rien coché montrent la
    // même chose. On garde quand même les deux états : décocher tout remet le filtre à zéro.
    if(el.getAttribute('data-g')==='owner'){
      ui.fOwners = (el.getAttribute('data-v')==='1')
        ? ownerOptions().map(function(o){ return o[0]; })
        : [];
    }
    deplierTout(); render(); return;
  }
  if(a==='reset'){ resetFilters(); render(); return; }
  if(a==='togprio'){ ui.prioOnly=!ui.prioOnly; deplierTout(); render(); return; }
  if(a==='togdeps'){ ui.showDeps=!ui.showDeps; render(); return; }
  if(a==='togfilters'){ ui.filtersOpen=!ui.filtersOpen; render(); return; }
  if(a==='cell'){ ouvrirCellule(el.getAttribute('data-id'), el.getAttribute('data-f')); return; }
  if(a==='lot-menu'){
    var pid=el.getAttribute('data-p'), ln=el.getAttribute('data-l');
    ui.lotMenu=(ui.lotMenu&&ui.lotMenu.pid===pid&&ui.lotMenu.lot===ln)?null:{pid:pid,lot:ln};
    render(); return;
  }
  if(a==='lot-rename'){
    var inp=document.getElementById('lotrename');
    renommerLot(el.getAttribute('data-p'), el.getAttribute('data-l'), inp?inp.value:'');
    return;
  }
  if(a==='lot-ungroup'){ degrouperLot(el.getAttribute('data-p'), el.getAttribute('data-l')); return; }
  if(a==='own-add'){
    var nv=document.getElementById('ownnew');
    if(nv && nv.value.trim()) basculerProprietaire(el.getAttribute('data-id'), nv.value.trim());
    return;
  }
  if(a==='add-dep'){
    var sel=document.getElementById('depinput');
    if(sel&&sel.value){ collectDraft(); ui.draft.deps=(ui.draft.deps||[]).slice(); ui.draft.deps.push(sel.value); render(); }
    return;
  }
  if(a==='del-dep'){
    collectDraft(); ui.draft.deps=(ui.draft.deps||[]).slice();
    ui.draft.deps.splice(+el.getAttribute('data-i'),1); render(); return;
  }
  if(a==='add-owner'){ addOwner(); return; }
  if(a==='quick-owner'){ addOwner(el.getAttribute('data-o')); return; }
  if(a==='del-owner'){
    collectDraft();
    ui.draft.owners=(ui.draft.owners||[]).slice();
    ui.draft.owners.splice(+el.getAttribute('data-i'),1);
    render(); return;
  }
  if(a==='tab'){ ui.tab=el.getAttribute('data-t'); render(); if(ui.tab==='gantt') scrollToToday(); return; }
  if(a==='zoom'){ ui.zoom=el.getAttribute('data-z'); render(); scrollToToday(); return; }
  if(a==='today'){ scrollToToday(); return; }
  if(a==='arbo'){ basculerArbo(); return; }
  if(a==='toggle'){ var k=el.getAttribute('data-key'); ui.collapsed[k]=!ui.collapsed[k]; render(); return; }
  if(a==='new-obj'){ openDrawer('obj',{color:SLOTS[state.objectives.length%SLOTS.length]}); return; }
  if(a==='new-proj'){ openDrawer('proj',{objectiveId:(state.objectives[0]||{}).id||''}); return; }
  if(a==='new-task'){
    var pid=el.getAttribute('data-p')||(state.projects[0]||{}).id||'';
    var lot=el.getAttribute('data-l')||'';
    // Une tâche garée dans le Backlog n'a ni date ni propriétaire : c'est ce qui la
    // distingue d'une tâche planifiée, et ce qui la garde hors du Gantt et des filtres.
    var bk=(lot===BACKLOG);
    openDrawer('task',{projectId:pid, lot:lot, progress:0, prio:false, owners:[], deps:[],
      start:bk?'':iso(TODAY), end:bk?'':iso(addDays(TODAY,7))});
    return;
  }
  if(a==='new-lot'){ nouveauLotDepuis(el.getAttribute('data-t')); return; }
  if(a==='dep-del'){ retirerLien(el.getAttribute('data-from'), el.getAttribute('data-to')); return; }
  if(a==='edit-obj'){ openDrawer('obj',byId(state.objectives,el.getAttribute('data-id'))); return; }
  if(a==='edit-proj'){ openDrawer('proj',byId(state.projects,el.getAttribute('data-id'))); return; }
  if(a==='edit-task'){ openDrawer('task',byId(state.tasks,el.getAttribute('data-id'))); return; }
  if(a==='color'){ collectDraft(); ui.draft.color=el.getAttribute('data-c'); render(); return; }
  if(a==='apply'){ applyDraft(); return; }
  if(a==='close'){ closeDrawer(); return; }
  if(a==='del'){ if(!ui.confirmDel){ ui.confirmDel=true; render(); } else deleteCurrent(); return; }
  if(a==='data'){ ui.drawer='data'; ui.draft={}; ui.confirmDel=false; render(); return; }
  if(a==='copy-data'){
    var ta=document.getElementById('datajson');
    if(ta){ ta.focus(); ta.select();
      var done=false;
      try{ if(navigator.clipboard&&navigator.clipboard.writeText){ navigator.clipboard.writeText(ta.value); done=true; } }catch(e){}
      if(!done){ try{ done=document.execCommand('copy'); }catch(e2){} }
      var m=document.getElementById('datamsg');
      if(m) m.textContent = done? 'Copié.' : 'Copie automatique bloquée — le texte est sélectionné, fais Ctrl+C.';
    }
    return;
  }
  if(a==='reload'){ chargerTout().then(render); return; }
  if(a==='login'){ if(window.demarrerApp) window.demarrerApp(true); return; }
});
document.addEventListener('change',function(ev){
  if(ev.target && ev.target.hasAttribute && ev.target.hasAttribute('data-inline')){ validerCellule(ev.target); return; }
  // Le select de projet n'a pas de data-act : il faut le traiter avant le garde-fou ci-dessous.
  if(ev.target && ev.target.getAttribute && ev.target.getAttribute('data-f')==='projectId' && ui.drawer==='task'){
    collectDraft(); render(); return;
  }
  var el=ev.target.closest('[data-act]'); if(!el) return;
  var a=el.getAttribute('data-act');
  if(a==='fobj'){ ui.fObj=el.value; deplierTout(); render(); }
  if(a==='hidedone'){ ui.hideDone=el.checked; deplierTout(); render(); }
  if(a==='own-chk'){ basculerProprietaire(el.getAttribute('data-id'), el.getAttribute('data-v')); return; }
  if(a==='chk'){
    var g=el.getAttribute('data-g'), v=el.getAttribute('data-v');
    if(g==='owner') toggleIn(ui.fOwners, v);
    deplierTout(); render();
  }
});
document.addEventListener('input',function(ev){
  var el=ev.target;
  if(el.getAttribute && el.getAttribute('data-f')==='progress'){
    var lb=el.closest('.lb'); if(lb) lb.querySelector('span').textContent='Avancement — '+el.value+'%';
  }
});
document.addEventListener('focusout',function(ev){
  if(ev.target && ev.target.hasAttribute && ev.target.hasAttribute('data-inline')) validerCellule(ev.target);
});
document.addEventListener('keydown',function(ev){
  if(ev.target && ev.target.hasAttribute && ev.target.hasAttribute('data-inline')){
    if(ev.key==='Enter'){ ev.preventDefault(); validerCellule(ev.target); return; }
    if(ev.key==='Escape'){ ev.preventDefault(); ui.editing=null; render(); return; }
  }
  if(ev.key==='Enter' && ev.target && ev.target.id==='lotrename'){
    ev.preventDefault();
    var br=document.querySelector('[data-act="lot-rename"]');
    if(br) renommerLot(br.getAttribute('data-p'), br.getAttribute('data-l'), ev.target.value);
    return;
  }
  if(ev.key==='Escape' && ui.lotMenu){ ui.lotMenu=null; render(); return; }
  if(ev.key==='Enter' && ev.target && ev.target.id==='ownnew'){
    ev.preventDefault();
    var b=document.querySelector('[data-act="own-add"]');
    if(b && ev.target.value.trim()) basculerProprietaire(b.getAttribute('data-id'), ev.target.value.trim());
    return;
  }
  if(ev.key==='Escape' && ui.editing){ ui.editing=null; render(); return; }
  if(ev.key==='Enter' && ev.target && ev.target.id==='ownerinput'){ ev.preventDefault(); addOwner(); return; }
  if(ev.key==='Escape' && ui.drawer) closeDrawer();
  if((ev.key==='Enter'||ev.key===' ') && ev.target.closest && ev.target.closest('.card')){
    ev.preventDefault(); ev.target.closest('.card').click();
  }
});

var suppressClick=false, bdrag=null, drag=null, lnk=null;
function traceLien(x,y){
  var sv=document.getElementById('lienprov');
  if(!sv){
    sv=document.createElementNS('http://www.w3.org/2000/svg','svg');
    sv.id='lienprov';
    sv.setAttribute('class','lienprov');
    sv.innerHTML='<path></path>';
    document.body.appendChild(sv);
  }
  var m=14, d;
  if(x > lnk.x0+2*m) d='M'+lnk.x0+' '+lnk.y0+' H'+(lnk.x0+m)+' V'+y+' H'+x;
  else d='M'+lnk.x0+' '+lnk.y0+' H'+(lnk.x0+m)+' V'+((lnk.y0+y)/2)+' H'+(x-m)+' V'+y+' H'+x;
  sv.firstChild.setAttribute('d',d);
}
function effacerLien(){
  var sv=document.getElementById('lienprov');
  if(sv && sv.parentNode) sv.parentNode.removeChild(sv);
}

/* ---- glisser une ligne pour changer l'ordre ---- */
/* Seule la poignée est déplaçable : rendre la ligne entière draggable empêcherait de
   sélectionner du texte et entrerait en conflit avec le glissement des barres. */
var dragOrdre=null;
function netOrdre(){
  try{
    document.querySelectorAll('.odrag,.oavant,.oapres,.olot').forEach(function(o){
      o.classList.remove('odrag'); o.classList.remove('oavant');
      o.classList.remove('oapres'); o.classList.remove('olot'); });
  }catch(e){}
}
document.addEventListener('dragstart',function(ev){
  var g=ev.target.closest && ev.target.closest('.ogrip');
  if(!g) return;
  var r=g.closest('.row.rowt'), rl=g.closest('.row.rowlot');
  if(r) dragOrdre={id:r.getAttribute('data-tid'), pid:r.getAttribute('data-tpid')};
  else if(rl) dragOrdre={lot:rl.getAttribute('data-droplot'), pid:rl.getAttribute('data-tpid')};
  else return;
  (r||rl).classList.add('odrag');
  try{ ev.dataTransfer.setData('text/plain','ordre'); ev.dataTransfer.effectAllowed='move'; }catch(e){}
  ev.stopPropagation();
});
document.addEventListener('dragend',function(){ if(dragOrdre){ dragOrdre=null; netOrdre(); } });
function cibleOrdre(ev){
  if(!dragOrdre) return null;
  // Un lot ne se dépose que sur un autre en-tête de lot du même projet.
  if(dragOrdre.lot!=null){
    var gl=ev.target.closest && ev.target.closest('.row.rowlot[data-droplot]');
    if(gl && gl.getAttribute('data-tpid')===dragOrdre.pid &&
       gl.getAttribute('data-droplot')!==dragOrdre.lot){
      var bl=gl.getBoundingClientRect();
      return {el:gl, lot:gl.getAttribute('data-droplot'), apres:(ev.clientY > bl.top+bl.height/2)};
    }
    return null;
  }
  var r=ev.target.closest && ev.target.closest('.row.rowt[data-tid]');
  if(r && r.getAttribute('data-tpid')===dragOrdre.pid && r.getAttribute('data-tid')!==dragOrdre.id){
    var b=r.getBoundingClientRect();
    return {el:r, id:r.getAttribute('data-tid'), apres:(ev.clientY > b.top+b.height/2)};
  }
  // En-tête de lot : la tâche rejoint ce lot. Ligne de projet : elle en sort.
  var g=ev.target.closest && ev.target.closest('.row.rowlot[data-droplot], .row.rowp[data-droplot]');
  if(g && g.getAttribute('data-tpid')===dragOrdre.pid)
    return {el:g, lot:g.getAttribute('data-droplot'), apres:true};
  return null;
}
document.addEventListener('dragover',function(ev){
  var c=cibleOrdre(ev); if(!c) return;
  ev.preventDefault();
  try{ ev.dataTransfer.dropEffect='move'; }catch(e){}
  try{
    document.querySelectorAll('.oavant,.oapres,.olot').forEach(function(o){
      o.classList.remove('oavant'); o.classList.remove('oapres'); o.classList.remove('olot'); });
  }catch(e){}
  c.el.classList.add(c.lot!=null ? 'olot' : (c.apres?'oapres':'oavant'));
});
document.addEventListener('drop',function(ev){
  var c=cibleOrdre(ev); if(!c) return;
  ev.preventDefault();
  var d=dragOrdre; dragOrdre=null; netOrdre();
  if(d.lot!=null) deplacerLot(d.pid, d.lot, c.lot, c.apres);
  else if(c.lot!=null) deplacerTache(d.id, null, true, c.lot);
  else deplacerTache(d.id, c.id, c.apres, null);
});
document.addEventListener('mousedown',function(ev){
  var dout=ev.target.closest && ev.target.closest('[data-depout]');
  if(dout && ev.button===0){
    ev.preventDefault(); ev.stopPropagation();
    var rr=dout.getBoundingClientRect();
    lnk={from:dout.getAttribute('data-depout'), x0:rr.left+rr.width/2, y0:rr.top+rr.height/2};
    document.body.classList.add('linking');
    traceLien(lnk.x0,lnk.y0);
    return;
  }
  var bar=ev.target.closest && ev.target.closest('[data-drag]');
  if(bar && ev.button===0){
    ev.preventDefault();
    var hh=ev.target.closest('.h');
    bdrag={
      kind:bar.getAttribute('data-drag'), id:bar.getAttribute('data-id'),
      mode: hh? hh.getAttribute('data-h') : 'move',
      x0:ev.clientX, el:bar, moved:false,
      left0:parseFloat(bar.style.left)||0, w0:parseFloat(bar.style.width)||0
    };
    document.body.classList.add(bdrag.mode==='move'?'dragging':'dragging-x');
    return;
  }
  var hd=ev.target.closest && ev.target.closest('.rz'); if(!hd) return;
  ev.preventDefault();
  var c=hd.getAttribute('data-col');
  var cur=parseFloat(getComputedStyle(document.documentElement).getPropertyValue(COLS[c]))||120;
  drag={col:c,x0:ev.clientX,w0:cur,el:hd};
  hd.classList.add('on'); document.body.style.userSelect='none';
});
document.addEventListener('mousemove',function(ev){
  if(lnk){ traceLien(ev.clientX,ev.clientY); return; }
  if(bdrag){
    var dx=ev.clientX-bdrag.x0;
    if(Math.abs(dx)>3) bdrag.moved=true;
    var d=Math.round(dx/view.ppd), px=d*view.ppd;
    if(bdrag.mode==='move') bdrag.el.style.left=(bdrag.left0+px)+'px';
    else if(bdrag.mode==='l'){ var nw=Math.max(view.ppd,bdrag.w0-px); bdrag.el.style.left=(bdrag.left0+bdrag.w0-nw)+'px'; bdrag.el.style.width=nw+'px'; }
    else { bdrag.el.style.width=Math.max(view.ppd,bdrag.w0+px)+'px'; }
    bdrag.days=d;
    return;
  }
  if(!drag) return;
  var w=Math.max(MINW[drag.col], drag.w0+(ev.clientX-drag.x0));
  document.documentElement.style.setProperty(COLS[drag.col], w+'px');
});
document.addEventListener('mouseup',function(ev){
  if(lnk){
    var l=lnk; lnk=null;
    document.body.classList.remove('linking');
    effacerLien();
    var el2=document.elementFromPoint(ev.clientX,ev.clientY);
    var cible=el2 && el2.closest && el2.closest('[data-drag="task"]');
    if(cible) ajouterLien(l.from, cible.getAttribute('data-id'));
    return;
  }
  if(bdrag){
    var b=bdrag; bdrag=null;
    document.body.classList.remove('dragging','dragging-x');
    if(!b.moved || !b.days){
      b.el.style.left=b.left0+'px'; if(b.w0) b.el.style.width=b.w0+'px';
      return;
    }
    suppressClick=true; setTimeout(function(){ suppressClick=false; },300);
    var d=b.days;
    if(b.kind==='task'){
      var t=taskById(b.id);
      if(t){
        if(b.mode==='move') shiftTask(t,d);
        else if(b.mode==='l'){ var st=parseD(t.start)||parseD(t.end); if(st){ var ns=addDays(st,d); var en=parseD(t.end)||st; if(ns>en) ns=en; t.start=iso(ns); } }
        else { var e0=parseD(t.end)||parseD(t.start); if(e0){ var ne=addDays(e0,d); var s0=parseD(t.start)||e0; if(ne<s0) ne=s0; t.end=iso(ne); } }
      }
    } else if(b.kind==='proj'){
      tasksOf(b.id).forEach(function(t){ shiftTask(t,d); });
    } else if(b.kind==='lot'){
      var p0=b.id.split('|')[0], l0=b.id.slice(p0.length+1);
      state.tasks.forEach(function(t){ if(t.projectId===p0 && t.lot===l0) shiftTask(t,d); });
    } else if(b.kind==='obj'){
      var o=byId(state.objectives,b.id);
      if(o){
        if(b.mode==='move'){ if(o.start) o.start=iso(addDays(parseD(o.start),d)); if(o.end) o.end=iso(addDays(parseD(o.end),d)); }
        else if(b.mode==='l' && o.start){ var a2=addDays(parseD(o.start),d), e2=parseD(o.end)||a2; if(a2>e2) a2=e2; o.start=iso(a2); }
        else if(o.end){ var b2=addDays(parseD(o.end),d), s2=parseD(o.start)||b2; if(b2<s2) b2=s2; o.end=iso(b2); }
      }
    }
    render();
    if(b.kind==='task'){ var tt=taskById(b.id); if(tt) enregistrer('task',tt); }
    else if(b.kind==='proj'){ enregistrerLot(tasksOf(b.id).map(function(x){ return ['task',x]; })); }
    else if(b.kind==='lot'){
      var pr=b.id.split('|')[0], ln=b.id.slice(pr.length+1);
      enregistrerLot(state.tasks.filter(function(x){ return x.projectId===pr && x.lot===ln; })
                                .map(function(x){ return ['task',x]; }));
    }
    else if(b.kind==='obj'){ var oo=byId(state.objectives,b.id); if(oo) enregistrer('obj',oo); }
    return;
  }
  if(!drag) return;
  drag.el.classList.remove('on'); document.body.style.userSelect='';
  drag=null; savePrefs();
});

var rsz=null, wasSmall=isSmall();
window.addEventListener('resize',function(){
  clearTimeout(rsz);
  rsz=setTimeout(function(){ if(isSmall()!==wasSmall){ wasSmall=isSmall(); render(); } },200);
});
window.addEventListener('beforeunload',function(ev){
  if(enVol>0){ ev.preventDefault(); ev.returnValue=''; return ''; }
});
loadPrefs();
render();

async function demarrer(interactif){
  ui.needLogin=false;
  try{
    await SP.connexion(interactif);
  }catch(e){
    ui.needLogin=true;
    ui.msg='Connexion à Microsoft 365 impossible : '+(e && e.message ? e.message : 'interrompue');
    marquer('error','Non connecté');
    render();
    return;
  }
  ui.msg='';
  await chargerTout();
  render();
  if(ui.tab==='gantt') scrollToToday();
}
window.demarrerApp = demarrer;
demarrer();
})();
