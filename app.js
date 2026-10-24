const DATA_FILE_PREFIX='planning-';
const DATA_FILE_SUFFIX='.json';
const DAY_NAMES=['Lun','Mar','Mer','Jeu','Ven','Sam','Dim'];
const MONTH_NAMES=['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
const MONTH_SHORT=['jan','fév','mars','avr','mai','juin','juil','août','sept','oct','nov','déc'];
const CODE_INFO={
  R:['Réserve','Selon besoin'],RM:['Rotation matin','05:00–13:00'],RA:['Rotation après-midi','13:00–21:00'],RN:['Rotation nuit','21:00–05:00'],
  VA:['Vacances','Absence planifiée'],VJ:['Vacances','Absence planifiée'],HOR:['Récupération d\'heure optionnel','Absence planifiée'],
  HSR:['Récupération d\'heure supplémentaire','Absence planifiée'],PIQ:['Piquet','Service de piquet'],RF:['Repos férié','Jour férié'],
  CA:['Camion','Affectation'],GR:['Groupe','Affectation'],EN:['Entretien','Affectation'],FE:['Formation externe','Formation'],
  PO:['Portes','Affectation'],MI:['Militaire','Absence'],PI:['Reprise de piquet','Absence planifiée'],PC:['Protection civile','Absence planifiée'],EM:['Enfant malade','Absence planifiée'],AN:['Ancienneté','Absence'],
  AP:['Accident professionnel','Absence'],PR:['Personnel','Absence'],MS:['Maladie','Absence'],MC:['Maladie / certificat','Absence'],
  NS:['Naissance','Absence planifiée'],PA:['Congé parental','Absence planifiée'],DE:['Congé déménagement','Absence planifiée'],
  DL:['Délégation','Absence planifiée'],DC:['Congé décès','Absence planifiée'],JT:['Congé','Absence']
};

let planning={year:new Date().getFullYear(),employees:[],updatedAt:'',publicationId:'',schemaVersion:0,dateCount:0,travaux:null};
let lastLoadedRawText='';
let lastJournal=[];
let state={view:['week','month','year'].includes(localStorage.getItem('planningView'))?localStorage.getItem('planningView'):'week',employeeId:null,date:new Date(),teamMode:false,teamYear:new Date().getFullYear(),compareMode:false,statsMode:false};
const $=id=>document.getElementById(id);
const pad=n=>String(n).padStart(2,'0');
const isoDate=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const todayISO=()=>isoDate(new Date());
const safeCode=v=>v==null?'':String(v).trim().toUpperCase();
const classFor=code=>code?`code-${code.replace(/[^A-Z0-9_-]/g,'')}`:'code-off';

const ROTATION_CODES=['RM','RA','RN'];
const TEAM_BAND_COLORS={RM:'#E3E0FF',RA:'#D9EDFF',RN:'#FFEDCF'};
const TEAM_BAND_EMPTY='#F1F4F7';
const LEGEND_ORDER=[
  'RM','RA','RN','PIQ','R','VA','VJ','HOR','HSR','FE','PC','PI','EM','RF','PO','CA','GR','EN','MI','AN','AP','PR','MS','MC','NS','PA','DE','DL','DC','JT'
];
const OFFICIAL_CODES=new Set(LEGEND_ORDER);
const AUTHENTICATED_EMAIL_KEY='planningTridelAuthenticatedEmail';

const ABSENCE_CODES=['VA','VJ','FE','HOR','HSR','PC','PI','EM','MI','AN','AP','PR','MS','MC'];

/* ==========================================================
   Jours fériés officiels (Vaud) — source : État de Vaud,
   vd.ch/formation/jours-feries-et-vacances-scolaires (dates officielles,
   même source et même logique que les vacances scolaires ci-dessous,
   plutôt qu'un calcul par formule).
   ========================================================== */
// Calculés pour n'importe quelle année (donc aucune mise à jour à prévoir) :
// dates fixes + fêtes mobiles déduites de Pâques. Liste officielle de l'art. 123
// de la Loi sur le personnel de l'État de Vaud (le 26 décembre, jour de congé
// de l'administration, n'est volontairement pas inclus).
function easterSunday(y){
  const a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3);
  const h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451);
  const month=Math.floor((h+l-7*m+114)/31),day=((h+l-7*m+114)%31)+1;
  return new Date(y,month-1,day);
}
const _holidayCache={};
function swissHolidaysOfYear(y){
  if(_holidayCache[y])return _holidayCache[y];
  const iso=d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  const E=easterSunday(y),plus=n=>{const d=new Date(E);d.setDate(d.getDate()+n);return iso(d);};
  // Lundi du Jeûne fédéral : le lendemain du 3e dimanche de septembre.
  const sep1=new Date(y,8,1),firstSun=1+((7-sep1.getDay())%7),jeune=new Date(y,8,firstSun+14+1);
  const map={};
  for(const [name,date] of [
    ["Jour de l'an",`${y}-01-01`],['2 janvier',`${y}-01-02`],
    ['Vendredi Saint',plus(-2)],['Lundi de Pâques',plus(1)],['Ascension',plus(39)],['Lundi de Pentecôte',plus(50)],
    ['Fête nationale',`${y}-08-01`],['Jeûne fédéral',iso(jeune)],['Noël',`${y}-12-25`]
  ])map[date]=name;
  return _holidayCache[y]=map;
}

function swissHolidayName(dISO){
  const y=Number(String(dISO).slice(0,4));
  if(!y)return null;
  return swissHolidaysOfYear(y)[dISO]||null;
}

// ==========================================================
// Période de travaux (source : onglet "jours fériés" du fichier Excel,
// cellule "Travaux" suivie de 2 dates — voir normalizeTravaux()).
// ==========================================================

// APERÇU TEMPORAIRE — uniquement pour visualiser le rendu avant que le
// fichier Excel ne publie de vraies dates de travaux. Dès qu'une période
// valide est publiée (planning.travaux), elle prend systématiquement le
// dessus sur cette liste (voir travauxPeriod() ci-dessous). Si l'aperçu
// de l'année en cours est dépassé, on passe automatiquement à celui de
// l'année suivante listé ici ; si aucun n'est encore disponible, rien ne
// s'affiche. À retirer une fois les vraies dates publiées par Excel.
const TRAVAUX_APERCU=[
  {start:'2026-05-17', end:'2026-06-17'},
  {start:'2027-05-23', end:'2027-06-28'}
];

function travauxApercu(){
  const todayIso=todayISO();
  return TRAVAUX_APERCU.find(t=>t.end>=todayIso)||null;
}

function travauxPeriod(){
  const t=planning.travaux;
  if(t && t.start && t.end)return t;
  return travauxApercu();
}

function isInTravaux(dISO){
  const t=travauxPeriod();
  return !!t && dISO>=t.start && dISO<=t.end;
}

function parseISODateOnly(v){
  const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v||'').trim());
  if(!m)return null;
  const d=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]));
  if(isNaN(d.getTime()))return null;
  d.setHours(0,0,0,0);
  return d;
}

// ==========================================================
// Vacances scolaires vaudoises — source : État de Vaud,
// vd.ch/formation/jours-feries-et-vacances-scolaires (dates officielles).
// Affichage volontairement discret (voir CSS .is-schoolholiday / .week-school-dot).
// ==========================================================
// Source officielle : calendrier « Vacances scolaires vaudoises jusqu'en 2031 »
// publié par l'État de Vaud (fichier .ics, vd.ch/formation/jours-feries-et-
// vacances-scolaires). Dates de fin incluses (le .ics donne la fin exclusive).
// Au-delà de 2031, ajouter ici les nouvelles périodes dès qu'elles sont publiées.
const VD_SCHOOL_HOLIDAYS=[
  ['Vacances de Noël','2023-01-01','2023-01-08'],
  ['Relâches','2023-02-11','2023-02-19'],
  ['Vacances de Pâques','2023-04-07','2023-04-23'],
  ["Pont de l'Ascension",'2023-05-18','2023-05-19'],
  ["Vacances d'été",'2023-07-01','2023-08-20'],
  ["Vacances d'automne",'2023-10-14','2023-10-29'],
  ['Vacances de Noël','2023-12-23','2024-01-07'],
  ['Relâches','2024-02-10','2024-02-18'],
  ['Vacances de Pâques','2024-03-29','2024-04-14'],
  ["Pont de l'Ascension",'2024-05-09','2024-05-10'],
  ["Vacances d'été",'2024-06-29','2024-08-18'],
  ["Vacances d'automne",'2024-10-12','2024-10-27'],
  ['Vacances de Noël','2024-12-21','2025-01-05'],
  ['Relâches','2025-02-15','2025-02-23'],
  ['Vacances de Pâques','2025-04-12','2025-04-27'],
  ["Pont de l'Ascension",'2025-05-29','2025-05-30'],
  ["Vacances d'été",'2025-06-28','2025-08-17'],
  ["Vacances d'automne",'2025-10-11','2025-10-26'],
  ['Vacances de Noël','2025-12-20','2026-01-04'],
  ['Relâches','2026-02-14','2026-02-22'],
  ['Vacances de Pâques','2026-04-03','2026-04-19'],
  ["Pont de l'Ascension",'2026-05-14','2026-05-15'],
  ["Vacances d'été",'2026-06-27','2026-08-16'],
  ["Vacances d'automne",'2026-10-10','2026-10-25'],
  ['Vacances de Noël','2026-12-24','2027-01-10'],
  ['Relâches','2027-02-06','2027-02-14'],
  ['Vacances de Pâques','2027-03-26','2027-04-11'],
  ["Pont de l'Ascension",'2027-05-06','2027-05-07'],
  ["Vacances d'été",'2027-07-03','2027-08-22'],
  ["Vacances d'automne",'2027-10-09','2027-10-24'],
  ['Vacances de Noël','2027-12-24','2028-01-09'],
  ['Relâches','2028-02-12','2028-02-20'],
  ['Vacances de Pâques','2028-04-14','2028-04-30'],
  ["Pont de l'Ascension",'2028-05-25','2028-05-26'],
  ["Vacances d'été",'2028-07-01','2028-08-20'],
  ["Vacances d'automne",'2028-10-14','2028-10-29'],
  ['Vacances de Noël','2028-12-23','2029-01-07'],
  ['Relâches','2029-02-10','2029-02-18'],
  ['Vacances de Pâques','2029-03-30','2029-04-15'],
  ["Pont de l'Ascension",'2029-05-10','2029-05-11'],
  ["Vacances d'été",'2029-06-30','2029-08-19'],
  ["Vacances d'automne",'2029-10-13','2029-10-28'],
  ['Vacances de Noël','2029-12-22','2030-01-06'],
  ['Relâches','2030-02-16','2030-02-24'],
  ['Vacances de Pâques','2030-04-19','2030-05-05'],
  ["Pont de l'Ascension",'2030-05-30','2030-05-31'],
  ["Vacances d'été",'2030-07-06','2030-08-25'],
  ["Vacances d'automne",'2030-10-12','2030-10-27'],
  ['Vacances de Noël','2030-12-21','2031-01-05'],
  ['Relâches','2031-02-15','2031-02-23'],
  ['Vacances de Pâques','2031-04-11','2031-04-27'],
  ["Pont de l'Ascension",'2031-05-22','2031-05-23'],
  ["Vacances d'été",'2031-07-05','2031-08-24']
];

// Au-delà (ou avant) la période publiée par l'État de Vaud, les vacances sont
// ESTIMÉES d'après les règles observées sur les calendriers officiels 2023-2031 :
//  - automne : du premier samedi dès le 9 octobre, sur 16 jours
//  - Noël : du samedi qui précède le 25 déc. (le 24 si Noël tombe un ven./sam./dim.), jusqu'au 2e dimanche suivant
//  - Pâques : du vendredi saint au dimanche 14 jours après Pâques
//  - Ascension : jeudi + vendredi
//  - été : 7 semaines, jusqu'au dimanche avant la rentrée (lundi dès le 17 août)
//  - relâches : semaine du lundi compris entre le 12 et le 18 février (vrai 8 années sur 9 ; décalée d'une semaine en 2027)
// Dès que l'État de Vaud publie le calendrier suivant, l'ajouter dans VD_SCHOOL_HOLIDAYS.
const VD_SCHOOL_OFFICIAL_FIRST='2023-01-01',VD_SCHOOL_OFFICIAL_LAST='2031-08-24';
const _schoolEstCache={};
function estimateSchoolHolidays(year){
  if(_schoolEstCache[year])return _schoolEstCache[year];
  const D=(y,m,d)=>new Date(y,m-1,d,12);
  const add=(d,n)=>{ const x=new Date(d);x.setDate(x.getDate()+n);return x; };
  const iso=d=>isoDate(d);
  const nextDow=(d,dow)=>add(d,(dow-d.getDay()+7)%7);        // 0 = dimanche … 6 = samedi
  const prevDow=(d,dow)=>add(d,-((d.getDay()-dow+7)%7));
  const out=[];
  const easter=easterSunday(year);
  // Noël de l'année précédente (partie de janvier) et de cette année
  const xmas=y=>{
    const c=D(y,12,25),wd=c.getDay();               // 1..4 = lun..jeu
    if(wd>=1&&wd<=4){ const st=prevDow(c,6); return [iso(st),iso(add(st,15))]; }
    const st=D(y,12,24), end=nextDow(D(y+1,1,6),0);
    return [iso(st),iso(end)];
  };
  const x0=xmas(year-1),x1=xmas(year);
  out.push(['Vacances de Noël',x0[0],x0[1]],['Vacances de Noël',x1[0],x1[1]]);
  const mon7=nextDow(D(year,2,12),1);   // lundi entre le 12 et le 18 février
  out.push(['Relâches',iso(add(mon7,-2)),iso(add(mon7,6))]);
  out.push(['Vacances de Pâques',iso(add(easter,-2)),iso(add(easter,14))]);
  const asc=add(easter,39);
  out.push(["Pont de l'Ascension",iso(asc),iso(add(asc,1))]);
  const rentree=nextDow(D(year,8,17),1);
  out.push(["Vacances d'été",iso(add(rentree,-51)),iso(add(rentree,-1))]);
  const aut=nextDow(D(year,10,9),6);
  out.push(["Vacances d'automne",iso(aut),iso(add(aut,15))]);
  return _schoolEstCache[year]=out;
}
function schoolHolidayName(dISO){
  const found=VD_SCHOOL_HOLIDAYS.find(([,start,end])=>dISO>=start && dISO<=end);
  if(found)return found[0];
  if(dISO>=VD_SCHOOL_OFFICIAL_FIRST && dISO<=VD_SCHOOL_OFFICIAL_LAST)return null;
  const y=Number(dISO.slice(0,4));
  for(const yy of [y,y+1]){
    const e=estimateSchoolHolidays(yy).find(([,a,b])=>dISO>=a && dISO<=b);
    if(e)return e[0];
  }
  return null;
}

function codeMatchesFilter(code){return true;}

function colleaguesFor(dateISO,code,currentId){
  const c=safeCode(code);
  if(!ROTATION_CODES.includes(c))return [];
  return planning.employees
    .filter(e=>e.id!==currentId && safeCode(e.schedule[dateISO])===c)
    .map(e=>cleanEmployeeName(e.name));
}

function teamFor(dateISO,code){
  const c=safeCode(code);
  return planning.employees.filter(e=>safeCode(e.schedule[dateISO])===c);
}

function periodDates(){
  const dates=[];
  if(state.view==='week'){
    const start=mondayOf(state.date);
    for(let i=0;i<7;i++){const d=new Date(start);d.setDate(start.getDate()+i);dates.push(d)}
  }else if(state.view==='month'){
    const y=state.date.getFullYear(),m=state.date.getMonth();
    const last=new Date(y,m+1,0).getDate();
    for(let i=1;i<=last;i++)dates.push(new Date(y,m,i));
  }else if(state.view==='year'){
    const y=state.date.getFullYear();
    const d=new Date(y,0,1);
    while(d.getFullYear()===y){dates.push(new Date(d));d.setDate(d.getDate()+1)}
  }else{
    dates.push(new Date(state.date));
  }
  return dates;
}

function cleanEmployeeName(name){
  return String(name||'')
    .replace(/\s*\([^)]*\)\s*/g,' ')
    .replace(/\s+/g,' ')
    .trim();
}


function initials(name){
  return cleanEmployeeName(name)
    .split(/\s+/)
    .filter(Boolean)
    .slice(0,2)
    .map(part=>part.charAt(0))
    .join('')
    .toUpperCase();
}


function isVacantEmployee(name){
  const clean=cleanEmployeeName(name).toLocaleLowerCase('fr');
  return clean==='poste vacant' || clean==='poste-vacant' || clean==='vacant';
}

// Combine les employés de plusieurs fichiers (un par année) en une seule
// liste, en fusionnant par nom -- les identifiants (basés sur le numéro de
// ligne Excel) ne sont pas comparables d'un fichier à l'autre. Le planning
// et les commentaires de chaque année s'ajoutent : les clés de date étant
// propres à chaque année (yyyy-mm-dd), aucun risque de collision entre elles.
function mergeEmployeesByName(perYearData){
  const byName=new Map();

  for(const data of perYearData){
    for(const emp of data.employees){
      const key=emp.name.toLocaleUpperCase('fr');
      if(!byName.has(key)){
        byName.set(key,{id:emp.id,name:emp.name,category:'',schedule:{},comments:{}});
      }
      const merged=byName.get(key);
      // Fichiers triés par année croissante : la catégorie la plus récente l'emporte.
      if(emp.category)merged.category=emp.category;
      Object.assign(merged.schedule,emp.schedule);
      Object.assign(merged.comments,emp.comments);
    }
  }

  // Ordre du classeur Excel : l'ordre d'insertion du Map suit l'ordre de
  // première apparition de chaque personne en parcourant les années de
  // publication croissantes (voir boucle ci-dessus), donc aucun tri
  // alphabétique ici.
  return [...byName.values()];
}

function combinePlannings(perYearData){
  const employees=mergeEmployeesByName(perYearData.map(p=>p.data));

  const currentYear=new Date().getFullYear();
  const currentYearEntry=perYearData.find(p=>p.year===currentYear);

  return {
    year:currentYear,
    // Le chip "Mis à jour" ne doit refléter que l'année en cours : une
    // modification sur le fichier de l'année suivante ne doit pas laisser
    // croire que le planning affiché aujourd'hui vient d'être rafraîchi.
    updatedAt:currentYearEntry?currentYearEntry.data.updatedAt:'',
    publicationId:perYearData.map(p=>p.data.publicationId).filter(Boolean).join('+'),
    schemaVersion:perYearData.length?Math.max(...perYearData.map(p=>p.data.schemaVersion)):0,
    dateCount:perYearData.reduce((sum,p)=>sum+p.data.dateCount,0),
    // La période de travaux est une donnée globale (pas propre à une année) :
    // on prend la première publication qui en contient une.
    travaux:perYearData.map(p=>p.data.travaux).find(Boolean)||null,
    employees
  };
}

// Lit la période de travaux publiée (cellule "Travaux" + 2 dates dans l'onglet
// "jours fériés" de l'Excel). Absence ou dates invalides => aucune période
// n'est affichée (exigence explicite : ne rien afficher plutôt qu'une donnée
// partielle ou incohérente).
function normalizeTravaux(rawTravaux){
  if(!rawTravaux || typeof rawTravaux!=='object')return null;
  const start=String(rawTravaux.debut||rawTravaux.start||'').trim();
  const end=String(rawTravaux.fin||rawTravaux.end||'').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end))return null;
  if(end<start)return null;
  return {start,end};
}

function normalizeData(raw){
  if(!raw || typeof raw!=='object' || Array.isArray(raw)){
    throw new Error('planning.json ne contient pas un objet JSON valide.');
  }

  const schemaVersion=Number(raw.schema_version||raw.schemaVersion||1);
  const publicationId=String(raw.publication_id||raw.publicationId||'');
  const declaredEmployeeCount=Number(raw.employee_count||raw.employeeCount||0);
  const declaredDateCount=Number(raw.date_count||raw.dateCount||0);

  const rawEmployees=raw.employes||raw.employees;
  if(!Array.isArray(rawEmployees)){
    throw new Error('La liste des employés est absente ou invalide.');
  }

  const namesSeen=new Set();

  const employees=rawEmployees
    .map((e,index)=>{
      if(!e || typeof e!=='object'){
        throw new Error(`Employé invalide à la position ${index+1}.`);
      }

      const name=cleanEmployeeName(e.nom||e.name||'');
      const schedule=e.planning||e.schedule||{};
      const comments=e.commentaires||e.comments||{};
      // Catégorie / groupe de la personne (ex. « Électriciens »), facultatif.
      const category=String(e.categorie||e.catégorie||e.category||e.groupe||e.group||e.fonction||'').replace(/\s+/g,' ').trim();

      if(!name){
        throw new Error(`Un employé publié n’a pas de nom (position ${index+1}).`);
      }

      const nameKey=name.toLocaleUpperCase('fr');
      if(namesSeen.has(nameKey)){
        throw new Error(`Nom d’employé en double dans planning.json : ${name}.`);
      }
      namesSeen.add(nameKey);

      if(!schedule || typeof schedule!=='object' || Array.isArray(schedule)){
        throw new Error(`Planning invalide pour ${name}.`);
      }

      if(!comments || typeof comments!=='object' || Array.isArray(comments)){
        throw new Error(`Commentaires invalides pour ${name}.`);
      }

      const dateKeys=Object.keys(schedule);

      if(schemaVersion>=2 && declaredDateCount>0 && dateKeys.length!==declaredDateCount){
        throw new Error(`Nombre de dates incohérent pour ${name} (${dateKeys.length}/${declaredDateCount}).`);
      }

      for(const dateISO of dateKeys){
        if(!/^\d{4}-\d{2}-\d{2}$/.test(dateISO)){
          throw new Error(`Date invalide pour ${name} : ${dateISO}.`);
        }

        const value=schedule[dateISO];
        if(value!=null){
          const code=safeCode(value);
          if(code && !OFFICIAL_CODES.has(code)){
            /* code inconnu autorisé */
          }
        }
      }

      return {
        id:String(e.id||`ligne-${index+1}`),
        name,
        category,
        schedule,
        comments
      };
    })
    .filter(e=>e.name && !isVacantEmployee(e.name));
    // Pas de tri alphabétique : on garde l'ordre des lignes du classeur Excel
    // (celui dans lequel la macro VBA a publié le JSON).

  if(!employees.length){
    throw new Error('Aucun employé visible n’a été publié.');
  }

  if(schemaVersion>=2){
    if(!publicationId){
      throw new Error('Identifiant de publication manquant.');
    }
    if(declaredEmployeeCount!==employees.length){
      throw new Error(`Nombre d’employés incohérent (${employees.length}/${declaredEmployeeCount}).`);
    }
    if(declaredDateCount<=0){
      throw new Error('Nombre de dates publié invalide.');
    }
  }

  return {
    year:Number(raw.annee||raw.year)||new Date().getFullYear(),
    updatedAt:raw.mise_a_jour||raw.updatedAt||'',
    publicationId,
    schemaVersion,
    dateCount:declaredDateCount,
    travaux:normalizeTravaux(raw.travaux),
    employees
  };
}

async function loadPlanning(preserve=true){
  try{
    const years=neededPublicationYears();

    // Une année sans fichier publié (ex. la suivante, pas encore commencée)
    // est simplement ignorée -- pas une erreur, juste rien à ajouter d'elle.
    const fetched=await Promise.all(years.map(async year=>{
      try{
        const response=await fetch(`${DATA_FILE_PREFIX}${year}${DATA_FILE_SUFFIX}?v=${Date.now()}`,{
          cache:'no-store',
          headers:{'Accept':'application/json'}
        });
        if(!response.ok)return null;
        const raw=await response.json();
        return {year,raw,data:normalizeData(raw)};
      }catch(e){
        return null;
      }
    }));

    const valid=fetched.filter(Boolean).sort((a,b)=>a.year-b.year);

    // Une année suivante pas encore publiée (ex. 2027 avant que le fichier
    // existe) ne doit jamais générer d'erreur -- seule l'absence de l'année
    // EN COURS en est une : c'est elle qui doit toujours être disponible.
    const anneeEnCours=new Date().getFullYear();
    const anneeEnCoursTrouvee=valid.some(v=>v.year===anneeEnCours);

    if(!anneeEnCoursTrouvee){
      throw new Error(`Aucun planning publié trouvé pour ${anneeEnCours}.`);
    }

    const next=combinePlannings(valid.map(v=>({year:v.year,data:v.data})));
    if(!next.employees.length){
      throw new Error('Aucun employé visible n’a été publié.');
    }

    // Garde-fou : un export qui perd d'un coup une grande part des employés
    // ou des dates est probablement cassé -> on garde la version précédente.
    {
      let ref=null;
      if(planning.employees&&planning.employees.length)ref={e:planning.employees.length,d:planning.dateCount};
      else{ const c=loadCachedPlanning(); if(c&&c.employees.length)ref={e:c.employees.length,d:c.dateCount}; }
      if(ref&&ref.e>=10&&next.employees.length<ref.e*0.7){
        throw new Error(`Données suspectes : ${next.employees.length} employés au lieu de ${ref.e}.`);
      }
      if(ref&&ref.d>=100&&next.dateCount<ref.d*0.8){
        throw new Error(`Données suspectes : ${next.dateCount} dates au lieu de ${ref.d}.`);
      }
    }

    const rawText=JSON.stringify(valid.map(v=>v.raw));

    // Copie de secours pour un accès hors-ligne ultérieur (uniquement une fois validée).
    try{
      localStorage.setItem('planningCache',JSON.stringify({
        raws:valid.map(v=>({year:v.year,raw:v.raw})),
        cachedAt:new Date().toISOString()
      }));
    }catch(e){ /* stockage indisponible : tant pis, pas bloquant */ }

    const previousId=state.employeeId;

    // On compare le CONTENU reçu (tous les fichiers combinés), pas juste un
    // identifiant : si celui-ci ne changeait pas comme prévu pour une raison
    // quelconque, une comparaison par identifiant laisserait l'appli bloquée
    // indéfiniment sur une ancienne version malgré des republications réelles.
    const samePublication=
      preserve &&
      planning.employees?.length &&
      rawText===lastLoadedRawText;

    if(samePublication){
      $('sourceStatus').textContent='';
      $('sourceStatus').classList.remove('status-warning');
      return;
    }

    lastLoadedRawText=rawText;

    // On ne remplace la version affichée qu'après validation complète.
    planning=next;
    lastJournal=valid
      .flatMap(v=>(Array.isArray(v.raw.journal)?v.raw.journal:[]).map(entry=>({...entry,annee:v.year})))
      .sort((a,b)=>String(b.horodatage||'').localeCompare(String(a.horodatage||'')))
      .slice(0,30);

    initializeState(preserve ? previousId : null);
    populateEmployees();
    updateEmployeePicker();
    showUI();
    if(!restoreSavedPage())render();

    $('sourceStatus').textContent='';
    $('sourceStatus').classList.remove('status-warning');
  }catch(error){
    // Si une version valide est déjà affichée, on la conserve.
    if(planning.employees && planning.employees.length){
      console.error('Actualisation du planning refusée :',error);
      $('sourceStatus').textContent='⚠ Actualisation refusée';
      $('sourceStatus').classList.add('status-warning');
      return;
    }

    // Pas de version en mémoire (ex. premier chargement sans réseau) :
    // on retombe sur la dernière copie valide connue, si elle existe.
    const cached=loadCachedPlanning();
    if(cached){
      console.error('Chargement réseau échoué, utilisation de la copie hors-ligne :',error);
      planning=cached;
      initializeState(null);
      populateEmployees();
      updateEmployeePicker();
      showUI();
      if(!restoreSavedPage())render();
      $('sourceStatus').textContent='⚠ Hors-ligne';
      $('sourceStatus').classList.add('status-warning');
      return;
    }

    showError(error);
  }finally{
    try{ updateOfflineBadge(); }catch(e){ /* ignoré */ }
  }
}

// Indique discrètement quand l'appareil n'a pas de réseau : les données
// affichées sont alors la dernière copie reçue (date visible dans la pastille).
function updateOfflineBadge(){
  const el=$('sourceStatus'); if(!el)return;
  if(navigator.onLine===false){
    el.textContent='Hors ligne';
    el.classList.remove('status-warning');
    el.classList.add('is-offline');
  }else if(el.textContent==='Hors ligne'){
    el.textContent='';
    el.classList.remove('is-offline');
  }
}
window.addEventListener('offline',()=>{ try{ updateOfflineBadge(); }catch(e){} });
window.addEventListener('online',()=>{ try{ updateOfflineBadge(); loadPlanning(true); }catch(e){} });

function loadCachedPlanning(){
  try{
    const stored=localStorage.getItem('planningCache');
    if(!stored)return null;
    const parsed=JSON.parse(stored);
    if(!parsed || !Array.isArray(parsed.raws) || !parsed.raws.length)return null;
    return combinePlannings(parsed.raws.map(r=>({year:r.year,data:normalizeData(r.raw)})));
  }catch(e){
    return null;
  }
}


const EMPLOYEE_NAME_KEY='planningEmployeeName';

function persistSelectedEmployee(employee){
  try{
    localStorage.setItem('planningEmployee',employee.id);
    localStorage.setItem(EMPLOYEE_NAME_KEY,cleanEmployeeName(employee.name));
  }catch(e){ /* stockage indisponible : tant pis, pas bloquant */ }
}

function initializeState(preferredId){
  state.date=new Date();

  const authenticatedEmail=localStorage.getItem(AUTHENTICATED_EMAIL_KEY);
  const matchedEmployee=bestEmployeeForEmail(authenticatedEmail);

  const storedName=localStorage.getItem(EMPLOYEE_NAME_KEY);
  const byStoredName=storedName
    ? planning.employees.find(e=>cleanEmployeeName(e.name)===storedName)
    : null;

  // Un identifiant (basé sur le numéro de ligne Excel) peut redevenir
  // "valide" par coïncidence après une republication, tout en désignant
  // maintenant une AUTRE personne. On ne lui fait confiance que s'il
  // pointe vers le même nom que celui déjà mémorisé pour cette personne.
  const idIfMatchesStoredName=id=>{
    if(!id)return null;
    const candidate=planning.employees.find(e=>e.id===id);
    if(!candidate)return null;
    if(storedName && cleanEmployeeName(candidate.name)!==storedName)return null;
    return candidate;
  };

  // Si rien n'est mémorisé ni reconnu (première visite), on ne choisit
  // personne par défaut -- la personne doit explicitement choisir son nom.
  const resolved=
    matchedEmployee ||
    idIfMatchesStoredName(preferredId) ||
    idIfMatchesStoredName(localStorage.getItem('planningEmployee')) ||
    byStoredName ||
    null;

  state.employeeId=resolved?resolved.id:null;
  if(resolved)persistSelectedEmployee(resolved);
}

function employeeResultButtonHTML(e){
  return `
    <button type="button" class="employee-result ${e.id===state.employeeId?'active':''}" data-employee-id="${escapeHTML(e.id)}">
      <span class="employee-result-avatar">${escapeHTML(initials(e.name))}</span>
      <span class="employee-result-name">${escapeHTML(e.name)}</span>
    </button>
  `;
}

function populateEmployees(filter=''){
  const results=$('employeeResults');
  const query=filter.trim();

  if(query){
    // Recherche active : liste plate triée par pertinence, comme avant.
    const list=planning.employees
      .map(e=>({employee:e,score:fuzzyEmployeeScore(e.name,query)}))
      .filter(x=>x.score>=0)
      .sort((a,b)=>b.score-a.score || a.employee.name.localeCompare(b.employee.name,'fr',{sensitivity:'base'}))
      .map(x=>x.employee);

    results.innerHTML=list.length
      ? list.map(employeeResultButtonHTML).join('')
      : '<div class="employee-results-empty">Aucun employé trouvé</div>';
  }else{
    // Parcours libre : regroupé par équipe/fonction, dans l'ordre du classeur.
    const byGroup=new Map();
    for(const e of planning.employees){
      const group=typeof categoryLabel==='function' ? categoryLabel(e.category) : (e.category||'');
      if(!byGroup.has(group))byGroup.set(group,[]);
      byGroup.get(group).push(e);
    }

    const groups=[...byGroup.entries()].sort(([a],[b])=>{
      const orderDiff=(typeof categoryLabelOrderIndex==='function')
        ? categoryLabelOrderIndex(a)-categoryLabelOrderIndex(b)
        : 0;
      return orderDiff || a.localeCompare(b,'fr',{sensitivity:'base'});
    });

    results.innerHTML=groups.length
      ? groups.map(([name,people])=>{
          // Ordre du classeur Excel au sein du groupe (pas de tri alphabétique).
          const peopleHTML=people.map(employeeResultButtonHTML).join('');
          return `<div class="employee-result-group"><div class="employee-result-group-label">${escapeHTML(name)}</div>${peopleHTML}</div>`;
        }).join('')
      : '<div class="employee-results-empty">Aucun employé trouvé</div>';
  }

  results.querySelectorAll('.employee-result').forEach(button=>{
    button.addEventListener('click',()=>{
      state.employeeId=button.dataset.employeeId;
      persistSelectedEmployee(selectedEmployee());
      updateEmployeePicker();
      closeEmployeePicker();
      render();
    });
  });
}

function updateEmployeePicker(){
  const employee=selectedEmployee();
  if(!employee){
    $('employeePickerName').textContent='Choisir une personne';
    $('employeePickerAvatar').textContent='—';
    populateEmployees($('employeeSearch').value);
    return;
  }
  $('employeePickerName').textContent=cleanEmployeeName(cleanEmployeeName(employee.name));
  $('employeePickerAvatar').textContent=initials(cleanEmployeeName(employee.name))||'—';
  populateEmployees($('employeeSearch').value);
}

function openEmployeePicker(){
  $('employeePicker').classList.add('open');
  $('employeePickerButton').setAttribute('aria-expanded','true');
  $('employeeSearch').value='';
  populateEmployees('');
  $('employeeBackdrop').classList.toggle('hidden',!isPhoneLayout());
  lockPageScroll();
  // Sur téléphone, le clavier cacherait la moitié de la liste : pas de focus automatique.
  if(window.matchMedia('(hover:hover)').matches)setTimeout(()=>$('employeeSearch').focus(),0);
}

function closeEmployeePicker(){
  const wasOpen=$('employeePicker').classList.contains('open');
  $('employeePicker').classList.remove('open');
  $('employeeBackdrop').classList.add('hidden');
  if(wasOpen)unlockPageScroll();
  $('employeePickerButton').setAttribute('aria-expanded','false');
}

function showUI(){
  $('loading').classList.add('hidden');
  $('errorBox').classList.add('hidden');
  $('planningUI').classList.remove('hidden');
  $('sourceStatus').textContent='';

  if(planning.updatedAt){
    const d=new Date(planning.updatedAt);
    $('updateStatus').textContent=Number.isNaN(d.getTime())
      ? planning.updatedAt
      : `Mis à jour ${d.toLocaleDateString('fr-CH',{day:'2-digit',month:'2-digit',year:'numeric'})} à ${d.toLocaleTimeString('fr-CH',{hour:'2-digit',minute:'2-digit'})}`;
  }else{
    $('updateStatus').textContent='Date de publication inconnue';
  }
}

function showError(error){
  $('loading').classList.add('hidden');
  $('planningUI').classList.add('hidden');
  $('errorBox').classList.remove('hidden');
  const anneeEnCours=new Date().getFullYear();
  const fichierAttendu=`${DATA_FILE_PREFIX}${anneeEnCours}${DATA_FILE_SUFFIX}`;
  $('errorBox').innerHTML=`<strong>Impossible d’afficher le planning.</strong><br>${escapeHTML(error.message)}<br><br>Vérifie que <strong>${escapeHTML(fichierAttendu)}</strong> a bien été publié depuis Excel, dans le bon dépôt GitHub.`;
  $('sourceStatus').textContent='Erreur';
  $('updateStatus').textContent='';
}

function escapeHTML(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]))}

function parseLocalISO(dateISO){
  const [y,m,d]=String(dateISO).split('-').map(Number);
  return new Date(y,m-1,d);
}

function daysBetweenInclusive(a,b){
  return Math.round((b-a)/86400000)+1;
}

function formatShortDate(d){
  return d.toLocaleDateString('fr-CH',{day:'2-digit',month:'2-digit'});
}

function formatDateWithYear(d){
  return d.toLocaleDateString('fr-CH',{day:'2-digit',month:'2-digit',year:'numeric'});
}

function formatLongDate(d){
  return d.toLocaleDateString('fr-CH',{weekday:'short',day:'numeric',month:'short'});
}

function commentCategory(text){
  const t=String(text||'').toLocaleLowerCase('fr');
  if(t.includes('formation'))return 'training';
  if(t.includes('médecin du travail') || t.includes('medecin du travail') || t.includes('médecin travail') || t.includes('medecin travail'))return 'medical';
  return 'comment';
}


function normalizeEventText(value){
  return String(value||'')
    .replace(/\s+/g,' ')
    .trim();
}

function parseClockToken(token){
  const m=String(token||'').match(/\b([01]?\d|2[0-3])\s*(?:[:hH.](?:\s*([0-5]\d))?)?\b/);
  if(!m)return null;
  return {hour:Number(m[1]),minute:Number(m[2]||0)};
}

function extractAppointmentTimes(text){
  const value=normalizeEventText(text);

  // Formats gérés : 8h, 8h30, 08:30, 08.30, 8h30-9h15, 08:30 à 09:15
  const range=value.match(/\b([01]?\d|2[0-3])\s*(?:[:hH.](?:\s*([0-5]\d))?)?\s*(?:-|–|—|à|a)\s*([01]?\d|2[0-3])\s*(?:[:hH.](?:\s*([0-5]\d))?)?\b/i);
  if(range){
    return {
      start:{hour:Number(range[1]),minute:Number(range[2]||0)},
      end:{hour:Number(range[3]),minute:Number(range[4]||0)}
    };
  }

  const single=value.match(/\b(?:à\s*)?([01]?\d|2[0-3])\s*(?:[:hH.]\s*([0-5]\d))\b/i)
    || value.match(/\b(?:à\s*)?([01]?\d|2[0-3])\s*[hH]\b/i);

  if(single){
    const start={hour:Number(single[1]),minute:Number(single[2]||0)};
    const startMinutes=start.hour*60+start.minute;
    const endMinutes=startMinutes+60; // 1 h par défaut si aucune durée n'est indiquée.
    return {
      start,
      end:{hour:Math.floor((endMinutes%1440)/60),minute:endMinutes%60},
      rollsToNextDay:endMinutes>=1440
    };
  }

  return null;
}

function extractAddressFromComment(text){
  const value=normalizeEventText(text);

  // 1) Mention explicite : "lieu: ..." / "adresse: ..."
  const explicit=value.match(/\b(?:adresse|lieu)\s*[:\-]\s*(.+?)(?=\s+(?:à|a)\s+\d{1,2}\s*(?:h|:)|$)/i);
  if(explicit)return explicit[1].trim().replace(/[.,;]+$/,'');

  // 2) Adresse suisse classique.
  const swiss=value.match(/((?:rue|route|avenue|av\.?|chemin|ch\.?|place|pl\.?|boulevard|bd\.?|quai|impasse|all[eé]e)\s+[^,;|]+?(?:\s+\d+[a-zA-Z]?)?(?:,\s*|\s+)\d{4}\s+[A-Za-zÀ-ÿ'’\-\s]+)/i);
  if(swiss)return swiss[1].trim().replace(/[.,;]+$/,'');

  // 3) NPA + localité.
  const postal=value.match(/\b(\d{4}\s+[A-Za-zÀ-ÿ'’\-]+(?:\s+[A-Za-zÀ-ÿ'’\-]+){0,3})\b/);
  if(postal)return postal[1].trim().replace(/[.,;]+$/,'');

  // 4) Formulations naturelles : "à Lausanne à 14h", "chez Electrosuisse, Lausanne".
  // On exclut volontairement ce qui ressemble à une heure.
  const natural=value.match(/\b(?:à|a|sur|chez|site de|site)\s+([A-ZÀ-Ÿ][A-Za-zÀ-ÿ'’\-]+(?:\s+[A-ZÀ-Ÿ][A-Za-zÀ-ÿ'’\-]+){0,3})(?=\s+(?:à|a)\s+\d{1,2}\s*(?:h|:)|\s*[,;|]|$)/);
  if(natural){
    const candidate=natural[1].trim();
    if(!/^\d/.test(candidate))return candidate;
  }

  // 5) Cas courant sans préposition : "Electrosuisse Lausanne 14h".
  // On prend le dernier mot capitalisé juste avant l'heure comme localité probable.
  const beforeTime=value.match(/\b([A-ZÀ-Ÿ][A-Za-zÀ-ÿ'’\-]{2,})\s+(?=\d{1,2}\s*(?:h|:))/);
  if(beforeTime)return beforeTime[1].trim();

  return '';
}

function inferOrganizationFromEventText(text,location=''){
  let value=normalizeEventText(text);

  if(location){
    const escaped=location.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    value=value.replace(new RegExp(`\\b(?:à|a|sur|chez)?\\s*${escaped}\\b`,'i'),' ');
  }

  value=value
    .replace(/\b(?:formation(?:\s+externe)?|visite\s+m[eé]dicale|m[eé]decin\s+du\s+travail|rdv|rendez[- ]?vous)\b/gi,' ')
    .replace(/\b([01]?\d|2[0-3])\s*(?:[:hH.](?:\s*([0-5]\d))?)?\s*(?:-|–|—|à|a)\s*([01]?\d|2[0-3])\s*(?:[:hH.](?:\s*([0-5]\d))?)?\b/gi,' ')
    .replace(/\b(?:à\s*)?([01]?\d|2[0-3])\s*(?:[:hH.]\s*([0-5]\d)|[hH])\b/gi,' ')
    .replace(/\b(?:lieu|adresse)\s*[:\-]\s*.*/i,' ')
    .replace(/\s*[-–—,:;|]\s*/g,' ')
    .replace(/\s+/g,' ')
    .trim();

  const words=value.split(/\s+/).filter(Boolean);
  return words.slice(0,5).join(' ');
}

function foldAccents(value){
  return String(value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}

function stripEmployeeNamePrefix(text,employee){
  const original=String(text||'').trim();
  if(!original)return original;

  // Comparaison insensible aux accents ET à la casse : un nom stocké sans
  // accent ("Umuhire Theophile") doit être reconnu même si le commentaire
  // l'écrit avec ("Théophile Umuhire: ..."). Le repliage des accents ne
  // change pas la longueur du texte, donc la position trouvée dans le
  // texte "replié" reste valable pour découper le texte d'origine intact.
  const foldedOriginal=foldAccents(original).toLocaleLowerCase('fr');
  const escapeRe=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&').replace(/\s+/g,'\\s+');

  const allEmployees=Array.isArray(planning?.employees)?planning.employees:[];
  const ordered=employee?[employee,...allEmployees.filter(e=>e!==employee)]:allEmployees;

  for(const emp of ordered){
    const employeeName=cleanEmployeeName(emp?.name||'');
    if(!employeeName)continue;
    const parts=employeeName.split(/\s+/).filter(Boolean);

    // Nom complet, dans les deux ordres possibles (prénom nom / nom prénom).
    const fullCombos=[employeeName];
    if(parts.length===2)fullCombos.push(`${parts[1]} ${parts[0]}`);
    for(const candidate of fullCombos){
      const foldedCandidate=foldAccents(candidate).toLocaleLowerCase('fr');
      const match=foldedOriginal.match(new RegExp(`^${escapeRe(foldedCandidate)}\\s*[-–—,:;]+\\s*`));
      if(match)return original.slice(match[0].length).trim();
    }

    // Le prénom n'a pas besoin d'être présent : si une seule des parties du
    // nom (typiquement le nom de famille) est suivie d'un ":", ça suffit à
    // reconnaître une étiquette de nom -- ex. "Umuhire: Formation ...".
    for(const part of parts){
      if(part.length<2)continue;
      const foldedPart=foldAccents(part).toLocaleLowerCase('fr');
      const match=foldedOriginal.match(new RegExp(`^${escapeRe(foldedPart)}\\s*:\\s*`));
      if(match)return original.slice(match[0].length).trim();
    }
  }

  // Repli générique : la personne nommée peut ne PAS être dans la liste des
  // employés (ex. intervenant externe, formateur). Dans ce cas, on reconnaît
  // quand même le motif "Mot(s) avec Majuscule : ..." -- 1 à 3 mots
  // commençant par une majuscule, suivis directement de ":". On exclut
  // quelques mots capitalisés fréquents qui ne sont pas des noms propres,
  // pour éviter de couper "Attention : ..." ou "Important : ...".
  const NOT_A_NAME=new Set([
    'attention','important','urgent','rappel','note','info','information',
    'remarque','nota','bene','merci','bonjour','svp','stp','confidentiel',
    'rendez','rdv','objet','sujet','ps'
  ]);
  const nameWord="[A-ZÀ-ÞŒ][A-Za-zÀ-ÿœ'’.-]*";
  const genericMatch=original.match(new RegExp(`^((?:${nameWord})(?:\\s+(?:${nameWord})){0,2})\\s*:\\s*`));
  if(genericMatch){
    const firstWord=foldAccents(genericMatch[1].split(/\s+/)[0]).toLocaleLowerCase('fr');
    if(!NOT_A_NAME.has(firstWord))return original.slice(genericMatch[0].length).trim();
  }

  return original;
}

function removeEmployeePrefix(text,employee){
  let value=stripEmployeeNamePrefix(normalizeEventText(text),employee);

  // Cas fréquent : "Prénom Nom - Visite médicale..." alors que la liste ne contient que le nom.
  value=value.replace(/^[A-Za-zÀ-ÿ'’\-]+\s+[A-Za-zÀ-ÿ'’\-]+\s*[-–—,:;]\s*(?=visite|médecin|medecin|unisant[eé])/i,'');

  return value.trim();
}

function medicalEventFromComment(text,eventDate,employee){
  const cleaned=removeEmployeePrefix(text,employee);
  const times=extractAppointmentTimes(cleaned);
  const location=extractAddressFromComment(cleaned);

  const startDate=new Date(eventDate);
  let endDate=null;

  if(times){
    startDate.setHours(times.start.hour,times.start.minute,0,0);
    endDate=new Date(eventDate);
    if(times.rollsToNextDay)endDate.setDate(endDate.getDate()+1);
    endDate.setHours(times.end.hour,times.end.minute,0,0);
  }

  const info=[];
  if(times){
    info.push(
      startDate.toLocaleTimeString('fr-CH',{hour:'2-digit',minute:'2-digit'})
    );
  }
  if(location)info.push(location);

  return {
    type:'medical',
    date:startDate,
    endDate:endDate,
    allDay:!times,
    title:`Visite médicale Unisanté — ${formatLongDate(eventDate)}`,
    detail:info.join(' · ') || 'Visite médicale Unisanté',
    location,
    originalDetail:cleaned
  };
}


function stripDetectedTimeAndLocation(text, detectedLocation){
  let value=normalizeEventText(text);

  // Supprime les heures/horaires du titre potentiel.
  value=value
    .replace(/\b([01]?\d|2[0-3])\s*(?:[:hH.](?:\s*([0-5]\d))?)?\s*(?:-|–|—|à|a)\s*([01]?\d|2[0-3])\s*(?:[:hH.](?:\s*([0-5]\d))?)?\b/gi,' ')
    .replace(/\b(?:à\s*)?([01]?\d|2[0-3])\s*(?:[:hH.]\s*([0-5]\d))\b/gi,' ')
    .replace(/\b(?:à\s*)?([01]?\d|2[0-3])\s*[hH]\b/gi,' ');

  if(detectedLocation){
    const escaped=detectedLocation.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    value=value.replace(new RegExp(escaped,'i'),' ');
  }

  return value
    .replace(/\s*[-–—,:;|]\s*/g,' ')
    .replace(/\s+/g,' ')
    .trim();
}

function inferGenericEventTitle(rawText,fallbackTitle){
  const raw=normalizeEventText(rawText);
  const location=extractAddressFromComment(raw);
  const normalized=raw
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .toLocaleLowerCase('fr');

  if(normalized.includes('formation')){
    const organization=inferOrganizationFromEventText(raw,location);
    return organization ? `Formation — ${organization}` : 'Formation';
  }

  if(
    normalized.includes('visite medicale') ||
    normalized.includes('medecin du travail') ||
    normalized.includes('unisant')
  ){
    return 'Visite médicale Unisanté';
  }

  let cleaned=stripDetectedTimeAndLocation(raw,location)
    .replace(/^(rdv|rendez[- ]?vous)\s*[:\-–—]?\s*/i,'')
    .trim();

  if(!cleaned)return fallbackTitle;

  const words=cleaned.split(/\s+/);
  return (words.length>8 ? words.slice(0,8).join(' ') : cleaned) || fallbackTitle;
}

// « du 12 au 20 octobre », « du 30 octobre au 11 novembre », « du 28 décembre 2026 au 4 janvier 2027 »
function longDateRange(a,b){
  const mo=d=>MONTH_NAMES[d.getMonth()].toLowerCase();
  if(a.getTime()===b.getTime()){ const dn=['dimanche','lundi','mardi','mercredi','jeudi','vendredi','samedi'][a.getDay()]; return `${dn} ${a.getDate()} ${mo(a)}`; }
  if(a.getFullYear()!==b.getFullYear())return `du ${a.getDate()} ${mo(a)} ${a.getFullYear()} au ${b.getDate()} ${mo(b)} ${b.getFullYear()}`;
  if(a.getMonth()===b.getMonth())return `du ${a.getDate()} au ${b.getDate()} ${mo(b)}`;
  return `du ${a.getDate()} ${mo(a)} au ${b.getDate()} ${mo(b)}`;
}

function enrichEventForCalendar(ev){
  const sourceText=normalizeEventText(
    ev.originalDetail ||
    ev.detail ||
    ev.title ||
    ''
  );

  const times=extractAppointmentTimes(sourceText);
  const location=ev.location || extractAddressFromComment(sourceText);

  let start=new Date(ev.date);
  let end=ev.endDate ? new Date(ev.endDate) : null;
  let allDay=ev.allDay;

  // Si l'événement n'a pas déjà une heure explicite mais que le texte en contient une,
  // on transforme automatiquement l'événement "journée entière" en rendez-vous horaire.
  // Un événement explicitement marqué allDay:true (ex. la période de travaux, dont le
  // texte contient des dates au format "23.05.2027" qu'une regex d'heure pourrait
  // confondre avec "23h05") ne doit JAMAIS être réinterprété comme un rendez-vous horaire.
  const alreadyTimed=
    start.getHours()!==0 ||
    start.getMinutes()!==0 ||
    allDay===false ||
    allDay===true;

  if(!alreadyTimed && times){
    start=new Date(ev.date);
    start.setHours(times.start.hour,times.start.minute,0,0);

    end=new Date(ev.date);
    if(times.rollsToNextDay)end.setDate(end.getDate()+1);
    end.setHours(times.end.hour,times.end.minute,0,0);

    allDay=false;
  }

  if(allDay===undefined){
    allDay=!(
      start.getHours()!==0 ||
      start.getMinutes()!==0
    );
  }

  if(allDay===false && !end){
    end=new Date(start.getTime()+60*60*1000);
  }

  let title=ev.title;

  // Pour les formations et commentaires, fabrique un titre plus utile à partir du texte.
  if(ev.type==='training' || ev.type==='comment'){
    title=inferGenericEventTitle(sourceText,ev.title);
  }

  // Les visites médicales restent normalisées.
  if(ev.type==='medical'){
    title='Visite médicale Unisanté';
  }

  return {
    ...ev,
    date:start,
    endDate:end,
    allDay,
    location,
    title,
    originalDetail:sourceText || ev.originalDetail || ev.detail || ''
  };
}

function buildUpcomingEvents(employee){
  const start=new Date();
  start.setHours(0,0,0,0);

  const end=new Date(start);
  end.setDate(end.getDate()+29);

  // La fenêtre d'affichage (30 jours) ne doit borner QUE la recherche de
  // nouveaux événements, pas la durée réelle d'une période déjà entamée :
  // un piquet ou des vacances qui débordent de cette fenêtre doivent
  // afficher leur vraie date de fin, pas être tronqués au 30e jour.
  const probeLimit=new Date(start);
  probeLimit.setDate(probeLimit.getDate()+400);

  // Le "pont" ne doit relier que quelques jours de repos consécutifs
  // (un week-end, par exemple) — pas des semaines entières de code R,
  // sous peine de fusionner deux périodes totalement distinctes.
  const MAX_BRIDGE_DAYS=3;

  const events=[];
  const seen=new Set();
  const comments=employee.comments||{};

  const pushUnique=(event,key)=>{
    if(seen.has(key))return;
    seen.add(key);
    events.push(event);
  };

  // La période de travaux a sa propre ligne dans la carte "Prochains
  // événements" (voir renderTravauxSection / addTravauxToCalendar) --
  // elle n'apparaît volontairement pas dans la liste ci-dessous, aussi
  // lointaine soit-elle.

  // « Congé » = vacances (VA, VJ) + récupérations d'heures (HOR, HSR), cumulés en un seul bloc.
  const isVacationCode=code=>code==='VA' || code==='VJ' || code==='HOR' || code==='HSR' || code==='PI';

  // Jours qui peuvent relier deux blocs de vacances :
  // cellule vide / repos explicite.
  const isVacationBridgeCode=code=>{
    const c=safeCode(code);
    return c==='' || c==='R';
  };

  // Si une période (vacances, piquet…) est déjà en cours au tout premier
  // jour de la fenêtre (aujourd'hui), il faut remonter dans le temps pour
  // retrouver sa VRAIE date de début -- sinon l'événement affiché se
  // raccourcit artificiellement d'un jour à chaque nouvelle journée qui
  // passe, jusqu'à sembler avoir commencé "aujourd'hui".
  const backProbeLimit=new Date(start);
  backProbeLimit.setDate(backProbeLimit.getDate()-400);

  const findRealStart=(fromDate,matchFn,bridgeFn)=>{
    let realStart=new Date(fromDate);
    let probe=new Date(fromDate);
    probe.setDate(probe.getDate()-1);

    while(probe>=backProbeLimit){
      const probeCode=safeCode(employee.schedule[isoDate(probe)]);

      if(matchFn(probeCode)){
        realStart=new Date(probe);
        probe.setDate(probe.getDate()-1);
        continue;
      }

      if(bridgeFn(probeCode)){
        let bridgeStart=new Date(probe);
        let gapDays=0;
        let exceededLimit=false;

        while(bridgeStart>=backProbeLimit && bridgeFn(safeCode(employee.schedule[isoDate(bridgeStart)]))){
          bridgeStart.setDate(bridgeStart.getDate()-1);
          gapDays++;
          if(gapDays>MAX_BRIDGE_DAYS){ exceededLimit=true; break; }
        }

        if(!exceededLimit && bridgeStart>=backProbeLimit && matchFn(safeCode(employee.schedule[isoDate(bridgeStart)]))){
          realStart=new Date(bridgeStart);
          probe=new Date(bridgeStart);
          probe.setDate(probe.getDate()-1);
          continue;
        }
      }

      break;
    }

    return realStart;
  };

  // Jours qui peuvent relier deux blocs de congé : cellule vide, repos (R) ou repos férié (RF).
  const isLeaveBridgeCode=code=>{
    const c=safeCode(code);
    return c==='' || c==='R' || c==='RF';
  };

  // CONGÉS (vacances + récupérations)
  // Fusionne les périodes quand les jours intermédiaires sont uniquement des repos / cellules vides.
  let cursor=new Date(start);

  while(cursor<=end){
    const code=safeCode(employee.schedule[isoDate(cursor)]);

    if(isVacationCode(code)){
      const vacStart = cursor.getTime()===start.getTime()
        ? findRealStart(cursor,isVacationCode,isLeaveBridgeCode)
        : new Date(cursor);
      let vacEnd=new Date(cursor);

      let probe=new Date(cursor);
      probe.setDate(probe.getDate()+1);

      while(probe<=probeLimit){
        const probeCode=safeCode(employee.schedule[isoDate(probe)]);

        // Vacances qui se poursuivent normalement
        if(isVacationCode(probeCode)){
          vacEnd=new Date(probe);
          probe.setDate(probe.getDate()+1);
          continue;
        }

        // Trouver le prochain jour non-repos
        if(isLeaveBridgeCode(probeCode)){
          let bridgeEnd=new Date(probe);
          let gapDays=0;
          let exceededLimit=false;

          while(bridgeEnd<=probeLimit && isLeaveBridgeCode(employee.schedule[isoDate(bridgeEnd)])){
            bridgeEnd.setDate(bridgeEnd.getDate()+1);
            gapDays++;
            if(gapDays>MAX_BRIDGE_DAYS){ exceededLimit=true; break; }
          }

          // Si après quelques jours de repos seulement les vacances reprennent,
          // on considère que c'est une seule période.
          if(!exceededLimit && bridgeEnd<=probeLimit && isVacationCode(safeCode(employee.schedule[isoDate(bridgeEnd)]))){
            vacEnd=new Date(bridgeEnd);
            probe=new Date(bridgeEnd);
            probe.setDate(probe.getDate()+1);
            continue;
          }
        }

        break;
      }

      // Nombre de jours de congé réellement posés dans le bloc (week-ends et
      // jours fériés intercalés non comptés). Seuls les blocs de 5 jours et
      // plus sont annoncés dans les 30 prochains jours.
      let leaveDays=0,recupDays=0;
      for(let d=new Date(vacStart);d<=vacEnd;d.setDate(d.getDate()+1)){
        const c=safeCode(employee.schedule[isoDate(d)]);
        if(isVacationCode(c)){ leaveDays++; if(c==='HOR'||c==='HSR'||c==='PI')recupDays++; }
      }

      // Durée affichée : toute la période, week-ends et jours de repos inclus.
      const totalDays=Math.round((vacEnd-vacStart)/86400000)+1;

      // « Vacances » si le bloc ne contient que des VA / VJ ; « Congé » dès qu'il
      // mélange des récupérations (HOR / HSR).
      const vacOnly=recupDays===0;
      const range=longDateRange(vacStart,vacEnd);

      // Un ou plusieurs jours posés collés à un week-end (vendredi ou lundi) :
      // annoncés aussi, car ils rallongent le week-end.
      const dowStart=vacStart.getDay(),dowEnd=vacEnd.getDay();
      let touchesWeekend=dowEnd===5||dowStart===1;
      for(let d=new Date(vacStart);d<=vacEnd&&!touchesWeekend;d.setDate(d.getDate()+1)){ if(d.getDay()===6||d.getDay()===0)touchesWeekend=true; }
      const hasWeekdayLeave=(()=>{for(let d=new Date(vacStart);d<=vacEnd;d.setDate(d.getDate()+1)){const c=safeCode(employee.schedule[isoDate(d)]);if(isVacationCode(c)&&d.getDay()!==0&&d.getDay()!==6)return true;}return false;})();
      const longWeekend=leaveDays<5&&touchesWeekend&&hasWeekdayLeave;
      if(leaveDays>=5||longWeekend){
        pushUnique({
          type:'vacation',
          date:new Date(vacStart),
          endDate:new Date(vacEnd),
          allDay:true,
          title:vacOnly?`Vacances ${range}`:`Congé ${range}`,
          calendarTitle:vacOnly?`🏖️ Vacances ${range}`:`🌴 Congé ${range}`,
          detail:longWeekend
            ? `${leaveDays} jour${leaveDays>1?'s':''} ${vacOnly?'de congé':'de congé / récupération'} · week-end prolongé : ${totalDays+(dowEnd===5?2:0)+(dowStart===1?2:0)} jours de repos d'affilée`
            : vacOnly
            ? `${totalDays} jours de vacances, week-ends compris · Profitez-en bien !`
            : `${totalDays} jours de congé, week-ends compris${(recupDays<leaveDays)?` (dont ${recupDays} récupération${recupDays>1?'s':''})`:''}`
        },`vacation-${isoDate(vacStart)}-${isoDate(vacEnd)}`);
      }

      cursor=new Date(vacEnd);
      cursor.setDate(cursor.getDate()+1);
      continue;
    }

    cursor.setDate(cursor.getDate()+1);
  }

  // PIQUET : jours PIQ consécutifs, avec pont possible sur repos / cellules vides.
  cursor=new Date(start);

  while(cursor<=end){
    const code=safeCode(employee.schedule[isoDate(cursor)]);

    if(code==='PIQ'){
      const piqStart = cursor.getTime()===start.getTime()
        ? findRealStart(cursor,c=>c==='PIQ',isVacationBridgeCode)
        : new Date(cursor);
      let piqEnd=new Date(cursor);

      let probe=new Date(cursor);
      probe.setDate(probe.getDate()+1);

      while(probe<=probeLimit){
        const probeCode=safeCode(employee.schedule[isoDate(probe)]);

        if(probeCode==='PIQ'){
          piqEnd=new Date(probe);
          probe.setDate(probe.getDate()+1);
          continue;
        }

        if(isVacationBridgeCode(probeCode)){
          let bridgeEnd=new Date(probe);
          let gapDays=0;
          let exceededLimit=false;

          while(bridgeEnd<=probeLimit && isVacationBridgeCode(employee.schedule[isoDate(bridgeEnd)])){
            bridgeEnd.setDate(bridgeEnd.getDate()+1);
            gapDays++;
            if(gapDays>MAX_BRIDGE_DAYS){ exceededLimit=true; break; }
          }

          if(!exceededLimit && bridgeEnd<=probeLimit && safeCode(employee.schedule[isoDate(bridgeEnd)])==='PIQ'){
            piqEnd=new Date(bridgeEnd);
            probe=new Date(bridgeEnd);
            probe.setDate(probe.getDate()+1);
            continue;
          }
        }

        break;
      }

      pushUnique({
        type:'piquet',
        date:new Date(piqStart),
        endDate:new Date(piqEnd),
        title: piqStart.getTime()===piqEnd.getTime()
          ? `Piquet le ${formatShortDate(piqStart)}`
          : `Piquet du ${formatShortDate(piqStart)} au ${formatShortDate(piqEnd)}`,
        detail:'Service de piquet'
      },`piquet-${isoDate(piqStart)}-${isoDate(piqEnd)}`);

      cursor=new Date(piqEnd);
      cursor.setDate(cursor.getDate()+1);
      continue;
    }

    cursor.setDate(cursor.getDate()+1);
  }

  // FORMATION EXTERNE : les jours FE qui se suivent forment UN bloc. Un jour
  // avec commentaire ouvre un bloc ; les jours FE suivants sans commentaire
  // (ou avec le même texte) en font partie, un week-end ou un jour férié
  // intercalé ne coupe pas le bloc.
  const feKey=text=>String(text||'').toLocaleLowerCase('fr').replace(/\s+/g,' ').trim();
  const feComment=iso=>stripEmployeeNamePrefix(String(comments[iso]||'').trim(),employee);
  const feGapOk=(fromISO,toISO)=>{
    const d=parseLocalISO(fromISO);
    d.setDate(d.getDate()+1);
    while(isoDate(d)<toISO){
      const iso=isoDate(d);
      const dow=d.getDay();
      if(!(dow===0 || dow===6 || swissHolidayName(iso) || safeCode(employee.schedule[iso])==='RF'))return false;
      d.setDate(d.getDate()+1);
    }
    return true;
  };
  const feBlocks=[];
  for(const iso of Object.keys(employee.schedule).filter(k=>safeCode(employee.schedule[k])==='FE').sort()){
    const last=feBlocks[feBlocks.length-1];
    const c=feComment(iso);
    if(last && feGapOk(last.end,iso) && (!c || feKey(c)===feKey(last.comment))){
      last.end=iso;
      last.days++;
    }else{
      feBlocks.push({start:iso,end:iso,days:1,comment:c});
    }
  }
  const windowStartISO=isoDate(start),windowEndISO=isoDate(end);
  for(const blk of feBlocks){
    if(blk.end<windowStartISO || blk.start>windowEndISO)continue;
    const bStart=parseLocalISO(blk.start),bEnd=parseLocalISO(blk.end);
    if(blk.start===blk.end){
      pushUnique({
        type:'training',
        date:new Date(bStart),
        title:`Formation externe — ${formatLongDate(bStart)}`,
        detail:blk.comment || 'Formation externe'
      },`formation-${blk.start}`);
    }else{
      pushUnique({
        type:'training',
        date:new Date(bStart),
        endDate:new Date(bEnd),
        allDay:true,
        title:`Formation externe du ${formatShortDate(bStart)} au ${formatShortDate(bEnd)}`,
        detail:[blk.comment,`${blk.days} jours`].filter(Boolean).join(' · ')
      },`formation-${blk.start}-${blk.end}`);
    }
  }

  // COMMENTAIRES IMPORTANTS
  Object.entries(comments).forEach(([dateISO,rawText])=>{
    const raw=String(rawText||'').trim();
    if(!raw)return;

    // Si la date est déjà codée FE, l'événement "Formation externe"
    // affiche déjà ce commentaire comme détail : aucun doublon.
    // Les jours HOR/HSR n'affichent jamais leur commentaire (ni ici, ni
    // ailleurs dans l'appli) : ils n'apparaissent donc pas non plus dans
    // les "30 prochains jours".
    const dayCode=safeCode(employee.schedule[dateISO]);
    if(dayCode==='FE' || dayCode==='HOR' || dayCode==='HSR')return;

    const eventDate=parseLocalISO(dateISO);
    eventDate.setHours(0,0,0,0);

    if(eventDate<start || eventDate>end)return;

    const normalized=raw
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g,'')
      .toLocaleLowerCase('fr');

    const hasFormation=normalized.includes('formation');
    const hasMedical=
      normalized.includes('medecin du travail') ||
      normalized.includes('medecin travail') ||
      normalized.includes('visite medicale') ||
      normalized.includes('unisant');

    if(hasFormation){
      const times=extractAppointmentTimes(raw);
      const location=extractAddressFromComment(raw);
      const trainingDate=new Date(eventDate);
      let trainingEnd=null;

      if(times){
        trainingDate.setHours(times.start.hour,times.start.minute,0,0);
        trainingEnd=new Date(eventDate);
        if(times.rollsToNextDay)trainingEnd.setDate(trainingEnd.getDate()+1);
        trainingEnd.setHours(times.end.hour,times.end.minute,0,0);
      }

      const cleanedRaw=stripEmployeeNamePrefix(raw,employee);
      const trainingTitle=`${inferGenericEventTitle(cleanedRaw,'Formation')} — ${formatLongDate(eventDate)}`;

      pushUnique({
        type:'training',
        date:trainingDate,
        endDate:trainingEnd,
        allDay:!times,
        title:trainingTitle,
        detail:[
          times ? trainingDate.toLocaleTimeString('fr-CH',{hour:'2-digit',minute:'2-digit'}) : '',
          location
        ].filter(Boolean).join(' · ') || cleanedRaw,
        location,
        originalDetail:cleanedRaw
      },`formation-${dateISO}`);
    }

    if(hasMedical){
      pushUnique(
        medicalEventFromComment(raw,eventDate,employee),
        `medical-${dateISO}`
      );
    }

    if(!hasFormation && !hasMedical){
      const times=extractAppointmentTimes(raw);
      const location=extractAddressFromComment(raw);
      const commentDate=new Date(eventDate);
      let commentEnd=null;

      if(times){
        commentDate.setHours(times.start.hour,times.start.minute,0,0);
        commentEnd=new Date(eventDate);
        if(times.rollsToNextDay)commentEnd.setDate(commentEnd.getDate()+1);
        commentEnd.setHours(times.end.hour,times.end.minute,0,0);
      }

      pushUnique({
        type:'comment',
        date:commentDate,
        endDate:commentEnd,
        allDay:!times,
        title:`${formatLongDate(eventDate)}`,
        detail:stripEmployeeNamePrefix(raw,employee),
        location
      },`comment-${dateISO}-${raw}`);
    }
  });

  events.sort((a,b)=>a.date-b.date);
  return events;
}

let upcomingCalendarEvents=[];

function calendarDateOnly(date){
  const d=new Date(date);
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
}

function calendarLocalDateTime(date){
  const d=new Date(date);
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}T${String(d.getHours()).padStart(2,'0')}${String(d.getMinutes()).padStart(2,'0')}00`;
}

function nextCalendarDay(date){
  const d=new Date(date);
  d.setDate(d.getDate()+1);
  return d;
}

function escapeICS(value){
  return String(value||'')
    .replace(/\\/g,'\\\\')
    .replace(/\r?\n/g,'\\n')
    .replace(/,/g,'\\,')
    .replace(/;/g,'\\;');
}

function calendarEventData(ev){
  const intelligentEvent=enrichEventForCalendar(ev);
  const employee=selectedEmployee();
  const employeeName=employee ? cleanEmployeeName(employee.name) : '';

  if(intelligentEvent.allDay===false || (intelligentEvent.date && (intelligentEvent.date.getHours()!==0 || intelligentEvent.date.getMinutes()!==0))){
    const start=new Date(intelligentEvent.date);
    const end=intelligentEvent.endDate ? new Date(intelligentEvent.endDate) : new Date(start.getTime()+60*60*1000);

    return {
      title:intelligentEvent.title,
      start,
      end,
      allDay:false,
      location:intelligentEvent.location||'',
      description:[
        intelligentEvent.originalDetail || intelligentEvent.detail,
        employeeName ? `Planning Tridel — ${employeeName}` : 'Planning Tridel'
      ].filter(Boolean).join('\n')
    };
  }

  const start=new Date(intelligentEvent.date);
  start.setHours(0,0,0,0);

  const inclusiveEnd=intelligentEvent.endDate ? new Date(intelligentEvent.endDate) : new Date(start);
  inclusiveEnd.setHours(0,0,0,0);

  return {
    title:ev.calendarTitle||intelligentEvent.title,
    start,
    end:nextCalendarDay(inclusiveEnd),
    allDay:true,
    location:intelligentEvent.location||'',
    description:[
      intelligentEvent.originalDetail || intelligentEvent.detail,
      employeeName ? `Planning Tridel — ${employeeName}` : 'Planning Tridel'
    ].filter(Boolean).join('\n')
  };
}

function downloadCalendarICS(ev){
  const data=calendarEventData(ev);
  const uid=`planning-tridel-${Date.now()}-${Math.random().toString(36).slice(2)}@tridelsa.ch`;
  const stamp=new Date().toISOString().replace(/[-:]/g,'').replace(/\.\d{3}Z$/,'Z');

  const dateLines=data.allDay
    ? [
        `DTSTART;VALUE=DATE:${calendarDateOnly(data.start)}`,
        `DTEND;VALUE=DATE:${calendarDateOnly(data.end)}`
      ]
    : [
        `DTSTART;TZID=Europe/Zurich:${calendarLocalDateTime(data.start)}`,
        `DTEND;TZID=Europe/Zurich:${calendarLocalDateTime(data.end)}`
      ];

  const ics=[
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//TRIDEL//Planning//FR',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    `UID:${uid}`,
    `DTSTAMP:${stamp}`,
    ...dateLines,
    `SUMMARY:${escapeICS(data.title)}`,
    data.location ? `LOCATION:${escapeICS(data.location)}` : '',
    `DESCRIPTION:${escapeICS(data.description)}`,
    'STATUS:CONFIRMED',
    'TRANSP:TRANSPARENT',
    'END:VEVENT',
    'END:VCALENDAR'
  ].filter(Boolean).join('\r\n');

  const blob=new Blob([ics],{type:'text/calendar;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const link=document.createElement('a');
  link.href=url;
  link.download='planning-tridel.ics';
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1500);
}

// Android : Google Agenda n'importe pas les fichiers .ics, on ouvre donc
// directement la fiche « nouvel événement » de Google Agenda (appli ou web).
const isAndroid=()=>/android/i.test(navigator.userAgent);
function openGoogleCalendar(ev){
  const data=calendarEventData(ev);
  const dates=data.allDay
    ? `${calendarDateOnly(data.start)}/${calendarDateOnly(data.end)}`
    : `${calendarLocalDateTime(data.start)}/${calendarLocalDateTime(data.end)}`;
  const params=new URLSearchParams({action:'TEMPLATE',text:data.title,dates,details:data.description,ctz:'Europe/Zurich'});
  if(data.location)params.set('location',data.location);
  window.open(`https://calendar.google.com/calendar/render?${params.toString()}`,'_blank','noopener');
}

function openOutlookCalendar(ev){
  const data=calendarEventData(ev);

  const params=new URLSearchParams({
    path:'/calendar/action/compose',
    rru:'addevent',
    subject:data.title,
    startdt:data.start.toISOString(),
    enddt:data.end.toISOString(),
    body:data.description,
    allday:data.allDay?'true':'false'
  });

  if(data.location)params.set('location',data.location);

  window.open(
    `https://outlook.office.com/calendar/0/deeplink/compose?${params.toString()}`,
    '_blank',
    'noopener'
  );
}

function isMobileCalendarDevice(){
  return /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent) ||
    (navigator.maxTouchPoints>1 && window.innerWidth<=900);
}

function calendarButtonFeedback(index){
  const button=document.querySelector(`.upcoming-calendar-btn[data-calendar-index="${index}"]`);
  if(!button)return;

  const span=button.querySelector('span');
  const previous=span?.textContent||'📅';
  button.classList.add('calendar-opening');
  if(span)span.textContent='✓';

  window.setTimeout(()=>{
    button.classList.remove('calendar-opening');
    if(span)span.textContent=previous;
  },1300);
}

function addUpcomingEventToCalendar(index){
  const ev=upcomingCalendarEvents[index];
  if(!ev)return;

  calendarButtonFeedback(index);

  if(isAndroid()){
    openGoogleCalendar(ev);
  }else if(isMobileCalendarDevice()){
    downloadCalendarICS(ev);
  }else{
    openOutlookCalendar(ev);
  }
}

window.addUpcomingEventToCalendar=addUpcomingEventToCalendar;

function addTravauxToCalendar(){
  const t=travauxPeriod();
  if(!t)return;

  const startD=parseISODateOnly(t.start);
  const endD=parseISODateOnly(t.end);
  if(!startD || !endD)return;

  const ev={
    type:'travaux',
    date:startD,
    endDate:endD,
    allDay:true,
    title:'Période de travaux',
    detail:`Pas de congés durant cette période. (${formatDateWithYear(startD)} – ${formatDateWithYear(endD)})`
  };

  const button=$('travauxCalendarBtn');
  if(button){
    const span=button.querySelector('span');
    const previous=span?.textContent||'📅';
    button.classList.add('calendar-opening');
    if(span)span.textContent='✓';
    window.setTimeout(()=>{
      button.classList.remove('calendar-opening');
      if(span)span.textContent=previous;
    },1300);
  }

  if(isAndroid()){
    openGoogleCalendar(ev);
  }else if(isMobileCalendarDevice()){
    downloadCalendarICS(ev);
  }else{
    openOutlookCalendar(ev);
  }
}

window.addTravauxToCalendar=addTravauxToCalendar;

function renderUpcomingEvents(employee){
  const card=$('upcomingEventsCard');
  const list=$('upcomingEvents');
  const count=$('upcomingCount');
  const events=buildUpcomingEvents(employee);
  upcomingCalendarEvents=events;

  card.classList.remove('hidden');
  count.textContent=events.length||'';

  if(!events.length){
    list.innerHTML='<div class="upcoming-empty">Aucun événement important dans les 30 prochains jours.</div>';
    return;
  }

  const iconFor=type=>type==='vacation'?'🏖️':type==='piquet'?'📞':type==='training'?'🎓':type==='medical'?'🩺':type==='travaux'?'⚒️':type==='comment'?'📝':'•';

  list.innerHTML=events.map((ev,index)=>`
    <div class="upcoming-item upcoming-${ev.type}">
      <div class="upcoming-icon">${iconFor(ev.type)}</div>
      <div class="upcoming-copy">
        <strong>${escapeHTML(ev.title)}</strong>
        <small>${escapeHTML(ev.detail)}</small>
      </div>
      <button
        class="upcoming-calendar-btn"
        type="button"
        onclick="addUpcomingEventToCalendar(${index})"
        aria-label="Ajouter au calendrier"
        title="Ajouter au calendrier"
      ><span aria-hidden="true">📅</span></button>
    </div>
  `).join('');
}
function selectedEmployee(){
  if(!state.employeeId)return null;
  return planning.employees.find(e=>e.id===state.employeeId)||null;
}
function mondayOf(date){const d=new Date(date);const day=(d.getDay()+6)%7;d.setDate(d.getDate()-day);d.setHours(0,0,0,0);return d}
function isoWeek(date){const d=new Date(Date.UTC(date.getFullYear(),date.getMonth(),date.getDate()));const day=d.getUTCDay()||7;d.setUTCDate(d.getUTCDate()+4-day);const start=new Date(Date.UTC(d.getUTCFullYear(),0,1));return Math.ceil((((d-start)/86400000)+1)/7)}

// Plage de navigation autorisée : 1 mois avant aujourd'hui, jusqu'à la fin de l'année
// (fin de l'année suivante à partir de décembre).
function navBounds(){
  const min=new Date();min.setHours(0,0,0,0);min.setMonth(min.getMonth()-1);
  // Confidentialité : le planning de l'année suivante n'est affiché qu'à partir
  // de décembre (même s'il est publié plus tôt). Avant, la limite est le 31 décembre.
  const now=new Date();
  const lastYear=now.getMonth()===11?now.getFullYear()+1:now.getFullYear();
  const max=new Date(lastYear,11,31);
  return {min,max};
}

// Un fichier planning-<année>.json distinct par année : détermine lesquels
// couvrent la vraie plage de navigation (mêmes bornes que navBounds), pour
// ne jamais aller chercher plus de fichiers que nécessaire.
function neededPublicationYears(){
  const {min,max}=navBounds();
  const years=new Set();
  // L'année précédente (archive, optionnelle) sert aux statistiques et à
  // Comparer ; son absence est ignorée sans erreur.
  for(let y=min.getFullYear()-1;y<=max.getFullYear();y++)years.add(y);
  return [...years].sort((a,b)=>a-b);
}

function monthKey(date){ return date.getFullYear()*12+date.getMonth(); }

function shiftDate(amount){
  const d=new Date(state.date);
  if(state.view==='week')d.setDate(d.getDate()+7*amount);
  else if(state.view==='month')d.setMonth(d.getMonth()+amount,1);
  else if(state.view==='year')d.setFullYear(d.getFullYear()+amount,0,1);
  else d.setDate(d.getDate()+amount);

  const {min,max}=navBounds();
  if(state.view==='week'){
    if(mondayOf(d)<mondayOf(min) || mondayOf(d)>mondayOf(max))return;
  }else if(state.view==='month'){
    const key=monthKey(d);
    if(key<monthKey(min) || key>monthKey(max))return;
  }else if(state.view==='year'){
    if(d.getFullYear()<min.getFullYear() || d.getFullYear()>max.getFullYear())return;
  }

  state.date=d;render();
}
function goToday(){state.date=new Date();render()}


function openDayModal(dateObj,employee=selectedEmployee()){
  if(!employee)return;
  const d=new Date(dateObj);
  const dISO=isoDate(d);
  const code=safeCode(employee.schedule[dISO]);
  const info=CODE_INFO[code]||[code||'Repos',code?'Code du planning':'Aucun service'];
  const comment=(code==='HOR'||code==='HSR')?'':commentFor(employee,dISO);
  const detailText=comment||info[1];
  const mates=ROTATION_CODES.includes(code)?colleaguesFor(dISO,code,employee.id):[];
  const holidayName=swissHolidayName(dISO);
  const schoolName=schoolHolidayName(dISO);
  const dayInTravaux=isInTravaux(dISO);

  $('dayModalTitle').textContent=cleanEmployeeName(employee.name);
  $('dayModalDate').textContent=d.toLocaleDateString('fr-CH',{weekday:'long',day:'numeric',month:'long',year:'numeric'});

  const teamHTML=ROTATION_CODES.includes(code)
    ? `<div class="day-detail-team"><strong>Équipe</strong>
       ${mates.length?mates.map(name=>`<div class="day-detail-person">
          <span class="day-detail-person-avatar">${escapeHTML(initials(name))}</span>
          <span>${escapeHTML(cleanEmployeeName(name))}</span>
        </div>`).join(''):'<div class="day-detail-empty">Aucun collègue trouvé</div>'}
       </div>`
    : '';

  const holidayHTML=holidayName
    ? `<div class="day-detail-holiday">${escapeHTML(holidayName)}</div>`
    : '';

  const schoolHTML=schoolName
    ? `<div class="day-detail-schoolholiday">Vacances scolaires VD — ${escapeHTML(schoolName)}</div>`
    : '';

  const travauxHTML=dayInTravaux
    ? `<div class="day-detail-travaux">⚒️ Période de travaux</div>`
    : '';

  $('dayModalBody').innerHTML=`${holidayHTML}${schoolHTML}${travauxHTML}<div class="day-detail-service ${classFor(code)}">
    <div class="day-detail-top">
      <div class="day-detail-code">${escapeHTML(code||'—')}</div>
      <div class="day-detail-info"><strong>${escapeHTML(info[0])}</strong><small>${escapeHTML(detailText)}</small></div>
    </div>${teamHTML}</div>`;

  $('dayModalBackdrop').classList.toggle('is-team-day',false);$('dayModalBackdrop').classList.add('open');
}

// Comparer (semaine) : toucher une case ouvre le détail du jour (code, remarque, avec qui).
function openCmpDay(el){
  const emp=cmpGroupsOrdered().flatMap(g=>g.people).find(e=>presencePersonKey(e)===el.dataset.cmpKey);
  if(!emp)return;
  const [y,m,d]=el.dataset.cmpIso.split('-').map(Number);
  openDayModal(new Date(y,m-1,d),emp);
}
$('cmpGrid').addEventListener('click',e=>{ const el=e.target.closest('[data-cmp-iso]'); if(el)openCmpDay(el); });
$('cmpGrid').addEventListener('keydown',e=>{ if(e.key!=='Enter'&&e.key!==' ')return; const el=e.target.closest('[data-cmp-iso]'); if(el){ e.preventDefault(); openCmpDay(el); } });

function closeDayModal(){ $('dayModalBackdrop').classList.remove('open'); $('dayModalBackdrop').classList.remove('is-team-day'); }
$('dayModalClose').addEventListener('click',closeDayModal);
$('dayModalBackdrop').addEventListener('click',e=>{ if(e.target===$('dayModalBackdrop'))closeDayModal(); });

function formatJournalDate(horodatage){
  const d=new Date(String(horodatage||'').replace(' ','T'));
  if(isNaN(d))return String(horodatage||'—');
  return d.toLocaleDateString('fr-CH',{day:'2-digit',month:'2-digit',year:'numeric'})+' à '+
         d.toLocaleTimeString('fr-CH',{hour:'2-digit',minute:'2-digit'});
}

function journalStatusInfo(resultat){
  if(resultat==='succes')return {label:'Réussi',cls:'ok'};
  if(resultat==='avertissements')return {label:'Réussi (avertissements)',cls:'warn'};
  if(resultat==='echec')return {label:'Échec',cls:'fail'};
  return {label:resultat||'—',cls:'warn'};
}

function renderJournal(){
  const body=$('journalModalBody');

  if(!lastJournal.length){
    body.innerHTML='<p class="journal-empty">Aucune publication enregistrée pour l\'instant.</p>';
    return;
  }

  body.innerHTML=lastJournal.map(entree=>{
    const statut=journalStatusInfo(entree.resultat);
    const declencheur=entree.declencheur==='enregistrement' ? 'Automatique · enregistrement' : 'Bouton';
    const details=String(entree.details||'').trim();

    return `
      <div class="journal-entry journal-${statut.cls}">
        <div class="journal-entry-head">
          <span class="journal-dot"></span>
          <strong>${escapeHTML(formatJournalDate(entree.horodatage))}</strong>
          <span class="journal-status">${escapeHTML(statut.label)}</span>
        </div>
        <div class="journal-stats">
          ${entree.annee?`${escapeHTML(String(entree.annee))} · `:''}${escapeHTML(declencheur)} · ${Number(entree.employes)||0} employés · ${Number(entree.dates)||0} jours · ${Number(entree.masques)||0} masqué(s)
        </div>
        ${details ? `<div class="journal-details">${escapeHTML(details)}</div>` : ''}
      </div>`;
  }).join('');
}

function openJournalModal(){
  renderJournal();
  $('journalModalBackdrop').classList.add('open');
}
function closeJournalModal(){ $('journalModalBackdrop').classList.remove('open'); }
$('journalAccessBtn').addEventListener('click',openJournalModal);
$('journalModalClose').addEventListener('click',closeJournalModal);
$('journalModalBackdrop').addEventListener('click',e=>{ if(e.target===$('journalModalBackdrop'))closeJournalModal(); });

$('teamPrevBtn').addEventListener('click',()=>teamShiftYear(-1));
$('teamNextBtn').addEventListener('click',()=>teamShiftYear(1));

function attachDayClickHandlers(){
  document.querySelectorAll('[data-date]').forEach(el=>{
    if(!el.dataset.date)return;

    el.addEventListener('click',()=>{
      const [y,m,d]=el.dataset.date.split('-').map(Number);
      openDayModal(new Date(y,m-1,d));
    });

    el.addEventListener('keydown',event=>{
      if(event.key==='Enter' || event.key===' '){
        event.preventDefault();
        el.click();
      }
    });
  });
}

// ==========================================================
// Planning d'équipe annuel (qui est en rotation matin / après-midi / nuit,
// jour par jour, toute l'année) -- vue supplémentaire, indépendante de
// l'employé sélectionné, accessible via le bouton discret dans l'en-tête.
// ==========================================================
function teamCellStyle(counts){
  const bg=[
    counts[0]>0?'var(--tm-rm)':'var(--c-off-bg)',
    counts[1]>0?'var(--tm-ra)':'var(--c-off-bg)',
    counts[2]>0?'var(--tm-rn)':'var(--c-off-bg)'
  ];
  return `background:linear-gradient(90deg,${bg[0]} 0%,${bg[0]} 30%,${bg[1]} 36%,${bg[1]} 64%,${bg[2]} 70%,${bg[2]} 100%)`;
}

function teamCountsFor(dISO){
  return ROTATION_CODES.map(code=>teamFor(dISO,code).filter(e=>!isVacantEmployee(e.name)).length);
}

const TEAM_SCALE_KEY='teamScale';
state.teamScale=(()=>{ try{ return localStorage.getItem(TEAM_SCALE_KEY)==='year'?'year':'week'; }catch(e){ return 'week'; } })();
state.teamDate=new Date();state.teamDate.setHours(0,0,0,0);

const TEAM_SHIFTS=[['RM','rm','Matin'],['RA','ra','Après-midi'],['RN','rn','Nuit']];
function teamPeople(dISO,code){
  return teamFor(dISO,code).filter(e=>!isVacantEmployee(e.name))
    .sort((a,b)=>cleanEmployeeName(a.name).localeCompare(cleanEmployeeName(b.name),'fr'));
}
function teamBadges(dISO){
  const hol=swissHolidayName(dISO),sch=schoolHolidayName(dISO),trv=isInTravaux(dISO);
  return (hol?`<span class="tw-badge hol" title="${escapeHTML(hol)}">${escapeHTML(hol)}</span>`:'')
    +(sch?`<span class="tw-badge sch" title="Vacances scolaires VD — ${escapeHTML(sch)}">Vacances scol.</span>`:'')
    +(trv?'<span class="tw-badge trv">Travaux</span>':'');
}

function renderTeamWeek(){
  const monday=mondayOf(state.teamDate);
  const days=Array.from({length:7},(_,i)=>{const d=new Date(monday);d.setDate(monday.getDate()+i);return d;});
  const today=todayISO(),me=state.employeeId;
  const sunday=days[6];
  const mo=d=>d.toLocaleDateString('fr-CH',{month:'short'}).replace('.','');
  $('teamYearTitle').textContent=`Semaine ${isoWeek(monday)} · ${monday.getDate()}${monday.getMonth()!==sunday.getMonth()?' '+mo(monday):''} – ${sunday.getDate()} ${mo(sunday)} ${sunday.getFullYear()}`;
  const {min,max}=navBounds();
  $('teamPrevBtn').disabled=monday<=mondayOf(min);
  $('teamNextBtn').disabled=monday>=mondayOf(max);

  // --- Ordinateur : jours en colonnes, Matin / Après-midi / Nuit en lignes
  let g='<div class="tw-grid"><div class="tw-corner"></div>';
  g+=days.map((d,i)=>{const iso=isoDate(d);return `<div class="tw-dayh ${i>=5?'is-we':''} ${iso===today?'is-today':''} ${schoolHolidayName(iso)?'is-sch '+schCls(iso):''}"><small>${DAY_NAMES[i]}</small><b${markAttr(iso)}>${d.getDate()}</b>${teamBadges(iso)}</div>`;}).join('');
  for(const [code,k,label] of TEAM_SHIFTS){
    g+=`<div class="tw-rowh ${k}">${label}</div>`;
    g+=days.map((d,i)=>{
      const iso=isoDate(d),team=teamPeople(iso,code);
      return `<div class="tw-cell ${k} ${i>=5?'is-we':''} ${iso===today?'is-today':''}" data-team-date="${iso}" tabindex="0" role="button" aria-label="${label}, ${escapeHTML(formatDateWithYear(d))}">`
        +(team.length?team.map(e=>`<span class="tw-name ${e.id===me?'is-me':''}" title="${escapeHTML(cleanEmployeeName(e.name))}">${escapeHTML(cleanEmployeeName(e.name))}</span>`).join('')+(team.length>3?'<span class="tw-plus">Renforcé</span>':''):'<span class="tw-none">—</span>')
        +'</div>';
    }).join('');
  }
  g+='</div>';

  // --- Téléphone : bandeau des 7 jours + détail du jour choisi
  const sel=new Date(state.teamDate);sel.setHours(0,0,0,0);
  const selISO=isoDate(sel);
  let m='<div class="tw-mobile"><div class="tw-strip">';
  m+=days.map((d,i)=>{const iso=isoDate(d);return `<button type="button" class="tw-sd ${i>=5?'is-we':''} ${iso===today?'is-today':''} ${iso===selISO?'is-sel':''} ${swissHolidayName(iso)?'is-hol':''} ${schoolHolidayName(iso)?'is-sch '+schCls(iso):''}" data-team-pick="${iso}" aria-label="${escapeHTML(formatDateWithYear(d))}"><small>${DAY_NAMES[i]}</small><b${markAttr(iso)}>${d.getDate()}</b></button>`;}).join('');
  m+='</div>';
  const long=sel.toLocaleDateString('fr-CH',{weekday:'long',day:'numeric',month:'long'});
  m+=`<div class="tw-daytitle"><strong>${escapeHTML(long.charAt(0).toUpperCase()+long.slice(1))}</strong>${teamBadges(selISO)}</div>`;
  for(const [code,k,label] of TEAM_SHIFTS){
    const team=teamPeople(selISO,code);
    m+=`<section class="tw-card ${k}"><header>${label}<em>${team.length?team.length+' personne'+(team.length>1?'s':''):''}${team.length>3?' · renforcé':''}</em></header>`
      +(team.length?'<ul>'+team.map(e=>`<li class="tw-name ${k} ${e.id===me?'is-me':''}" data-open-emp="${escapeHTML(e.id)}">${escapeHTML(cleanEmployeeName(e.name))}</li>`).join('')+'</ul>':'<div class="tw-none">Personne en rotation</div>')
      +'</section>';
  }
  m+='</div>';
  $('teamViewContainer').innerHTML=g+m+periodNotesHTML(days.map(isoDate));
  attachTeamDayClickHandlers();
}

function renderTeamView(){
  if(!$('teamViewSection'))return;
  const {min,max}=navBounds();
  const isWeek=state.teamScale==='week';
  document.querySelectorAll('[data-team-scale]').forEach(b=>b.classList.toggle('active',b.dataset.teamScale===state.teamScale));
  $('teamLegend').classList.toggle('hidden',isWeek);
  $('teamLegend').innerHTML='<span class="team-legend-item"><i class="tl-dot tl-rf"></i>Renforcé (plus de 3)</span><span class="team-legend-item"><i class="tl-dot tl-inc"></i>Équipe incomplète</span><span class="team-legend-item"><i class="tl-circ tl-sch"></i>Vacances scolaires</span><span class="team-legend-item"><i class="tl-circ tl-hol"></i>Jour férié</span>';
  if(isWeek){
    const t=new Date(state.teamDate);
    if(t<min)state.teamDate=new Date(min);
    else if(t>max)state.teamDate=new Date(max);
    renderTeamWeek();
    return;
  }
  const y=Math.min(Math.max(state.teamYear,min.getFullYear()),max.getFullYear());
  state.teamYear=y;
  $('teamYearTitle').textContent=String(y);
  $('teamPrevBtn').disabled=y<=min.getFullYear();
  $('teamNextBtn').disabled=y>=max.getFullYear();

  let h='<div class="year-calendar">';
  for(let m=0;m<12;m++){
    const first=new Date(y,m,1),start=mondayOf(first);
    h+=`<div class="mini-month"><h3>${MONTH_NAMES[m]}</h3><div class="mini-head">${DAY_NAMES.map(d=>`<span>${d[0]}</span>`).join('')}</div><div class="mini-grid">`;
    for(let i=0;i<42;i++){
      const d=new Date(start);d.setDate(start.getDate()+i);
      const inMonth=d.getMonth()===m;
      if(!inMonth){
        h+='<span class="mini-cell mini-empty"></span>';
        continue;
      }
      const dISO=isoDate(d);
      const isToday=dISO===todayISO();
      const counts=teamCountsFor(dISO);
      const reinforced=counts.some(n=>n>3);
      const incomplete=counts.some(n=>n===0);
      const holidayName=swissHolidayName(dISO);
      const inTravaux=isInTravaux(dISO);
      const schoolName=schoolHolidayName(dISO);
      const markClasses=[holidayName?'is-holiday':'',inTravaux?'is-travaux':'',schoolName?'is-schoolholiday':''].filter(Boolean).join(' ');
      h+=`<span class="mini-cell team-cell ${isToday?'today':''} ${reinforced?'is-reinforced':''} ${incomplete?'is-incomplete':''} ${markClasses}" style="" data-team-date="${dISO}" tabindex="0" role="button" ${holidayName?`title="${escapeHTML(holidayName)}"`:(schoolName?`title="${escapeHTML(schoolName)}"`:'')} aria-label="Équipe du ${escapeHTML(formatDateWithYear(d))}">${d.getDate()}</span>`;
    }
    h+='</div></div>';
  }
  h+='</div>';
  $('teamViewContainer').innerHTML=h;
  attachTeamDayClickHandlers();
}

function openTeamDayModal(dateObj){
  const d=new Date(dateObj);
  const dISO=isoDate(d);

  $('dayModalTitle').textContent="Équipe du jour";
  $('dayModalDate').textContent=d.toLocaleDateString('fr-CH',{weekday:'long',day:'numeric',month:'long',year:'numeric'});

  const holidayName=swissHolidayName(dISO);
  const schoolName=schoolHolidayName(dISO);
  const dayInTravaux=isInTravaux(dISO);

  const holidayHTML=holidayName?`<div class="day-detail-holiday">${escapeHTML(holidayName)}</div>`:'';
  const schoolHTML=schoolName?`<div class="day-detail-schoolholiday">Vacances scolaires VD — ${escapeHTML(schoolName)}</div>`:'';
  const travauxHTML=dayInTravaux?`<div class="day-detail-travaux">⚒️ Période de travaux</div>`:'';

  const shiftsHTML=ROTATION_CODES.map(code=>{
    const info=CODE_INFO[code]||[code,''];
    const team=teamFor(dISO,code).filter(e=>!isVacantEmployee(e.name));
    const peopleHTML=team.length
      ? team.map(emp=>`<div class="day-detail-person">
          <span class="day-detail-person-avatar">${escapeHTML(initials(emp.name))}</span>
          <span>${escapeHTML(cleanEmployeeName(emp.name))}</span>
        </div>`).join('')
      : '<div class="day-detail-empty">Personne en rotation</div>';

    return `<div class="day-detail-service ${classFor(code)}" style="margin-bottom:10px">
      <div class="day-detail-top">
        <div class="day-detail-code">${escapeHTML(code)}</div>
        <div class="day-detail-info"><strong>${escapeHTML(info[0])}</strong><small>${escapeHTML(info[1])}</small></div>
      </div>
      <div class="day-detail-team"><strong>Équipe${team.length>3?' · effectif renforcé':''}</strong>${peopleHTML}</div>
    </div>`;
  }).join('');

  $('dayModalBody').innerHTML=`${holidayHTML}${schoolHTML}${travauxHTML}<div class="day-detail-shifts">${shiftsHTML}</div>`;
  $('dayModalBackdrop').classList.toggle('is-team-day',true);$('dayModalBackdrop').classList.add('open');
}

document.querySelectorAll('[data-team-scale]').forEach(b=>b.addEventListener('click',()=>{
  state.teamScale=b.dataset.teamScale==='year'?'year':'week';
  try{ localStorage.setItem(TEAM_SCALE_KEY,state.teamScale); }catch(e){ /* ignoré */ }
  state.teamYear=state.teamDate.getFullYear();
  renderTeamView();
}));
$('teamTodayBtn').addEventListener('click',()=>{
  const n=new Date();n.setHours(0,0,0,0);
  state.teamDate=n;state.teamYear=n.getFullYear();
  renderTeamView();
});
$('teamViewContainer').addEventListener('click',event=>{
  const b=event.target.closest('[data-team-pick]');
  if(!b)return;
  const [y,m,d]=b.dataset.teamPick.split('-').map(Number);
  state.teamDate=new Date(y,m-1,d);
  renderTeamView();
});
function attachTeamDayClickHandlers(){
  document.querySelectorAll('[data-team-date]').forEach(el=>{
    if(!el.dataset.teamDate)return;

    el.addEventListener('click',()=>{
      const [y,m,d]=el.dataset.teamDate.split('-').map(Number);
      openTeamDayModal(new Date(y,m-1,d));
    });

    el.addEventListener('keydown',event=>{
      if(event.key==='Enter' || event.key===' '){
        event.preventDefault();
        el.click();
      }
    });
  });
}

// Flèches : une semaine (ou une année). Balayage sur téléphone : un jour.
function teamShiftDay(dir){
  const {min,max}=navBounds();
  const d=new Date(state.teamDate);d.setDate(d.getDate()+dir);
  if(d<min||d>max)return;
  state.teamDate=d;
  renderTeamView();
}
function teamShiftYear(amount){
  if(state.teamScale==='week'){
    const {min,max}=navBounds();
    const d=new Date(state.teamDate);d.setDate(d.getDate()+7*amount);
    if(mondayOf(d)<mondayOf(min)||mondayOf(d)>mondayOf(max))return;
    state.teamDate=d;
    renderTeamView();
    return;
  }
  const {min,max}=navBounds();
  const next=state.teamYear+amount;
  if(next<min.getFullYear() || next>max.getFullYear())return;
  state.teamYear=next;
  renderTeamView();
}

// ==========================================================
// Navigation entre les trois pages : planning individuel, planning
// d'équipe et présences annuelles (boutons en haut à droite).
// ==========================================================
function currentPage(){
  if(state.statsMode)return 'stats';
  if(state.compareMode)return 'compare';
  if(state.teamMode)return 'team';
  return 'individual';
}

// Pages de l'application : pour en ajouter une, il suffit d'ajouter une ligne
// ici (et son affichage dans setPage / render).
const NAV_PAGES=[
  {id:'individual',label:'Individuel',desc:'Mon planning, semaine · mois · année',
   icon:'<circle cx="12" cy="8" r="4"/><path d="M4 21v-1a6 6 0 0 1 6-6h4a6 6 0 0 1 6 6v1"/>'},
  {id:'team',label:'Rotations',desc:'Qui est en rotation, jour par jour',
   icon:'<path d="M21 12a9 9 0 1 1-3-6.7"/><path d="M21 4v5h-5"/>'},
  {id:'compare',label:'Comparer',desc:'Collègues côte à côte · semaine, mois, année',
   icon:'<path d="M8 3v18M16 3v18M3 8h5M16 8h5M3 16h5M16 16h5"/>'},
  {id:'stats',label:'Statistiques',desc:'Mes rotations, piquets, congés, absences',
   icon:'<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>'}
];
// Regroupement de l'interface : « Planning » (individuel · collègues · rotations)
// et « Statistiques ».
const NAV_GROUPS=[
  {id:'plan',label:'Planning',desc:'Individuel, collègues ou rotations',pages:['individual','compare','team']},
  {id:'stats',label:'Statistiques',desc:'Congés, absences, rotations, piquets',pages:['stats']}
];
NAV_GROUPS[0].icon=NAV_PAGES[0].icon;
NAV_GROUPS[1].icon=NAV_PAGES[3].icon;
var lastInGroup={plan:'individual',stats:'stats'};
try{ const li=JSON.parse(localStorage.getItem('planningLastInGroup')||'{}');
  if(NAV_GROUPS[0].pages.includes(li.plan))lastInGroup.plan=li.plan;
  else if(NAV_GROUPS[0].pages.includes(li.team))lastInGroup.plan=li.team; }catch(e){ /* ignoré */ }
function groupOf(page){ return NAV_GROUPS.find(g=>g.pages.includes(page))||NAV_GROUPS[0]; }

const PAGE_KEY='planningPage';

// ---- Thème : Auto (suit l'appareil) / Clair / Sombre -------------------
const THEME_KEY='themeMode';
function themeMode(){ try{ const m=localStorage.getItem(THEME_KEY); return m==='light'||m==='dark'?m:'auto'; }catch(e){ return 'auto'; } }
function applyTheme(){
  const m=themeMode();
  const dark=m==='dark'||(m==='auto'&&window.matchMedia&&matchMedia('(prefers-color-scheme: dark)').matches);
  document.documentElement.setAttribute('data-theme',dark?'dark':'light');
}
// Installation sur l'écran d'accueil (Android : fenêtre native ; iPhone : mode d'emploi).
let deferredInstall=null;
const isStandalone=()=>(window.matchMedia&&matchMedia('(display-mode: standalone)').matches)||navigator.standalone===true;
const isIOS=()=>/iphone|ipad|ipod/i.test(navigator.userAgent)||(navigator.platform==='MacIntel'&&navigator.maxTouchPoints>1);
window.addEventListener('beforeinstallprompt',e=>{ e.preventDefault(); deferredInstall=e; updateTeamToggleLabel(); });
window.addEventListener('appinstalled',()=>{ deferredInstall=null; updateTeamToggleLabel(); });
function installHTML(){
  if(isStandalone())return '';
  if(!deferredInstall&&!isIOS()&&!/android/i.test(navigator.userAgent))return '';
  return `<button type="button" class="install-btn" data-install><span style="font-size:18px;line-height:1">＋</span><span>Installer l'appli<small>Ajouter le planning à l'écran d'accueil</small></span></button><div id="installHint" class="install-hint hidden"></div>`;
}
function themeSegHTML(){
  const m=themeMode();
  return `<span class="theme-label">Apparence</span><div class="theme-seg" role="group" aria-label="Apparence">${[['auto','Auto'],['light','Clair'],['dark','Sombre']].map(([v,l])=>`<button type="button" data-theme-mode="${v}" class="${m===v?'is-on':''}">${l}</button>`).join('')}</div>`;
}
try{ matchMedia('(prefers-color-scheme: dark)').addEventListener('change',applyTheme); }catch(e){ /* ignoré */ }
// À l'impression : toujours le thème clair (papier).
window.addEventListener('beforeprint',()=>document.documentElement.setAttribute('data-theme','light'));
window.addEventListener('afterprint',applyTheme);

function updateTeamToggleLabel(){
  const page=currentPage();
  const grp=groupOf(page);
  $('navMenuLabel').textContent=grp.label;
  $('navMenuPanel').innerHTML=NAV_GROUPS.map(g=>{
    const on=g.id===grp.id;
    return `<button type="button" class="navm-item ${on?'is-active':''}" role="menuitem" data-nav-page="${on?page:lastInGroup[g.id]}">
    <span class="navm-ico"><svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${g.icon}</svg></span>
    <span class="navm-txt"><b>${g.label}</b><small>${g.desc}</small></span>
    ${on?'<svg class="navm-check" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12l5 5 9-10"/></svg>':''}
  </button>`;}).join('')+themeSegHTML()+profileHTML()+shareHTML()+installHTML();
  const subtitle=$('heroSubtitle');
  if(subtitle){
    subtitle.textContent=page==='stats'?'Statistiques':page==='compare'?'Comparer des collègues':page==='team'?'Rotations':'Planning individuel';
  }
}

function openNavMenu(){
  const btn=$('navMenuBtn'),panel=$('navMenuPanel');
  const r=btn.getBoundingClientRect();
  panel.classList.remove('hidden');
  if(window.innerWidth<=700){
    // Téléphone : volet au-dessus de la barre d'onglets.
    const tb=document.getElementById('tabBar');
    const h=tb&&!tb.hidden?tb.getBoundingClientRect().height:0;
    panel.style.top='auto';panel.style.left='10px';panel.style.width='auto';
    panel.style.right='10px';panel.style.bottom=`${Math.round(h+8)}px`;
  }else{
    panel.style.bottom='auto';panel.style.left='auto';panel.style.width='';
    panel.style.top=`${Math.round(r.bottom+8)}px`;
    panel.style.right=`${Math.max(10,Math.round(window.innerWidth-r.right))}px`;
  }
  btn.setAttribute('aria-expanded','true');
  const st=document.getElementById('tabSettings');if(st)st.setAttribute('aria-expanded','true');
}
function closeNavMenu(){
  $('navMenuPanel').classList.add('hidden');
  $('navMenuBtn').setAttribute('aria-expanded','false');
  const st=document.getElementById('tabSettings');if(st)st.setAttribute('aria-expanded','false');
}
$('navMenuBtn').addEventListener('click',event=>{
  event.stopPropagation();
  if($('navMenuPanel').classList.contains('hidden'))openNavMenu(); else closeNavMenu();
});
$('navMenuPanel').addEventListener('click',async event=>{
  if(event.target.closest('[data-navm-close]')){ closeNavMenu(); return; }
  if(event.target.closest('[data-share-link]')){ shareCurrentView(); return; }
  if(event.target.closest('[data-profile-export]')){ profileExport(); return; }
  if(event.target.closest('[data-profile-import]')){ profileImport(); return; }
  if(event.target.closest('[data-install]')){
    if(deferredInstall){
      deferredInstall.prompt();
      try{ await deferredInstall.userChoice; }catch(e){ /* ignoré */ }
      deferredInstall=null; closeNavMenu(); updateTeamToggleLabel();
    }else{
      const h=$('installHint');
      if(h&&/android/i.test(navigator.userAgent)){ h.innerHTML="Ouvre le menu du navigateur (<b>⋮</b> en haut à droite), puis appuie sur <b>« Installer l'application »</b> ou <b>« Ajouter à l'écran d'accueil »</b>."; h.classList.toggle('hidden'); }
      else if(h){ h.innerHTML="Dans Safari, appuie sur <b>le menu</b> (les trois traits ou points en bas de l'écran), puis sur <b>Partager</b>, puis sur <b>« Sur l'écran d'accueil »</b>."; h.classList.toggle('hidden'); }
    }
    return;
  }
  const tm=event.target.closest('[data-theme-mode]');
  if(tm){
    try{ localStorage.setItem(THEME_KEY,tm.dataset.themeMode); }catch(e){ /* ignoré */ }
    applyTheme();
    $('navMenuPanel').querySelectorAll('[data-theme-mode]').forEach(b=>b.classList.toggle('is-on',b===tm));
    return;
  }
  const item=event.target.closest('[data-nav-page]');
  if(!item)return;
  closeNavMenu();
  setPage(item.dataset.navPage);
});
document.addEventListener('click',event=>{
  // Téléphone : la feuille Réglages ne se ferme que par le bouton Réglages ou la croix.
  if(window.innerWidth<=700)return;
  if(!$('navMenuPanel').contains(event.target))closeNavMenu();
});
document.addEventListener('keydown',event=>{ if(event.key==='Escape')closeNavMenu(); });
window.addEventListener('resize',()=>{
  // Téléphone : la barre du navigateur qui apparaît/disparaît en défilant déclenche un « resize » ; on garde la feuille ouverte et on la repositionne.
  if(window.innerWidth<=700){ if(!$('navMenuPanel').classList.contains('hidden'))openNavMenu(); return; }
  closeNavMenu();
});

// Page affichée au dernier passage : restaurée une seule fois, au premier
// affichage des données (fermeture / actualisation de la page).
let pageRestored=false;
function restoreSavedPage(){
  if(pageRestored)return false;
  pageRestored=true;
  let saved='individual';
  try{ saved=localStorage.getItem(PAGE_KEY)||'individual'; }catch(e){ saved='individual'; }
  if(saved==='presence'){ saved='compare'; cmp.scale='year'; saveCmp(); }  // ancienne page Maintenance
  const shared=readSharedHash();
  if(shared){
    applySharedState(shared);
    saved=shared.get('p');
    try{ history.replaceState(history.state,'',location.pathname+location.search); }catch(e){ /* ignoré */ }
  }else periodRestore();
  periodReady=true;
  if(saved==='individual' || !NAV_PAGES.some(x=>x.id===saved))return false;
  histMode='replace';
  setPage(saved);
  return true;
}

updateTeamToggleLabel();

// Accessibilité : rôles et états des onglets, page courante dans le menu.
function a11ySync(){
  if(typeof tabBarSync==='function')tabBarSync();
  document.querySelectorAll('.tab,.cmp-tab').forEach(b=>{
    b.setAttribute('role','tab');
    b.setAttribute('aria-selected',b.classList.contains('active')?'true':'false');
  });
  document.querySelectorAll('.navm-item').forEach(b=>{
    if(b.classList.contains('is-active'))b.setAttribute('aria-current','page');else b.removeAttribute('aria-current');
  });
  document.querySelectorAll('.pv-chip,.cmp-chip').forEach(b=>{
    if(b.tagName==='BUTTON'&&!b.hasAttribute('data-nopress'))b.setAttribute('aria-pressed',b.classList.contains('active')||b.classList.contains('is-on')?'true':'false');
  });
}
(function(){
  let q=0;
  const run=()=>{ q=0; a11ySync(); };
  new MutationObserver(()=>{ if(!q)q=requestAnimationFrame(run); })
    .observe(document.querySelector('.app'),{subtree:true,childList:true,attributes:true,attributeFilter:['class']});
  a11ySync();
})();
if('serviceWorker' in navigator&&location.protocol!=='file:'){
  window.addEventListener('load',()=>{ navigator.serviceWorker.register('sw.js').catch(()=>{}); });
}
let pageAnimReady=false;
setTimeout(()=>{ pageAnimReady=true; },1200);
// Transition douce entre onglets : la page s'élargit / se rétrécit, et le
// contenu glisse + apparaît en fondu depuis le côté de l'onglet choisi.
function pageTransitionStart(prev,page){
  if(!pageAnimReady||prev===page)return null;
  if(window.matchMedia&&matchMedia('(prefers-reduced-motion:reduce)').matches)return null;
  const app=document.querySelector('.app');
  if(!app||!app.animate)return null;
  app.classList.add('app-anim');
  clearTimeout(pageTransitionStart.t);
  pageTransitionStart.t=setTimeout(()=>{
    app.classList.remove('app-anim');
    const st=document.getElementById('heroArtStage');
    if(st&&heroFocusReady)st.style.transform=heroArtTransform(heroFocusCurrent||'full');
  },640);
  return {app,prev,page};
}
function pageTransitionRun(t){
  if(!t)return;
  const order=NAV_PAGES.map(p=>p.id);
  const dir=order.indexOf(t.page)>=order.indexOf(t.prev)?1:-1;
  const phone=innerWidth<=700;
  const dx=44*dir;
  const els=[...t.app.children].filter(el=>!el.classList.contains('hero')&&!el.classList.contains('hidden')&&el.offsetParent!==null&&el.tagName!=='SCRIPT');
  if(phone){
    // Mobile : simple et fluide — tout le contenu monte d'un bloc en fondu,
    // sans glissement latéral ni décalage entre les cartes.
    if(window.scrollY>40)window.scrollTo({top:0,behavior:'smooth'});
    els.forEach(el=>{
      el.animate([{opacity:0,transform:'translateY(18px)'},{opacity:1,transform:'none'}],
        {duration:340,easing:'cubic-bezier(.2,.8,.2,1)',fill:'backwards'});
    });
  }else{
    els.forEach((el,i)=>{
      el.animate(
        [{opacity:0,transform:`translateX(${dx}px) scale(.985)`,transformOrigin:'50% 0'},
         {opacity:1,transform:'none',transformOrigin:'50% 0'}],
        {duration:480,delay:i*50,easing:'cubic-bezier(.22,1,.36,1)',fill:'backwards'});
    });
  }
  // Pendant l'élargissement, on recale l'illustration du bandeau à chaque image.
  const t0=performance.now(),key=heroFocusCurrent;
  const stage=document.getElementById('heroArtStage');
  (function tick(){
    if(!stage||!heroFocusReady||performance.now()-t0>640)return;
    if(heroFocusKey()===key&&!stage.classList.contains('is-anim'))stage.style.transform=heroArtTransform(key);
    requestAnimationFrame(tick);
  })();
}
// Garde la même période (semaine / mois, date) en passant d'une vue à l'autre.
function syncScopes(from,to){
  if(from===to||from==='stats'||to==='stats')return;
  let ref,scale;
  if(from==='individual'){ ref=state.date; scale=state.view; }
  else if(from==='compare'){ ref=cmp.date; scale=cmp.scale; }
  else { ref=state.teamScale==='year'?new Date(state.teamYear,0,1):state.teamDate; scale=state.teamScale==='year'?'year':'week'; }
  ref=new Date(ref);
  if(from==='team'&&scale==='year'){ const t=new Date(); ref=t.getFullYear()===state.teamYear?t:ref; }
  if(to!=='individual'||true){ state.date=new Date(ref); cmp.date=new Date(ref); state.teamDate=new Date(ref); }
  if(to==='individual'){ if(scale==='week'||scale==='month'||scale==='year')state.view=scale; }
  else if(to==='compare'){ cmp.scale=scale; saveCmp(); }
  else if(to==='team'){ state.teamScale=scale==='year'?'year':'week'; state.teamYear=ref.getFullYear(); }
}

// Ouvre le planning individuel d'une personne (depuis Collègues ou Rotations).
function openPersonIndividual(id){
  const emp=planning.employees.find(e=>e.id===id); if(!emp)return;
  state.employeeId=emp.id; persistSelectedEmployee(emp);
  setPage('individual');
  window.scrollTo({top:0});
}
document.addEventListener('click',e=>{
  const t=e.target.closest('[data-open-emp]'); if(!t||e.target.closest('.pv-hide'))return;
  openPersonIndividual(t.dataset.openEmp);
});

function setPage(page){
  if(page!=='individual' && (!planning.employees || !planning.employees.length))return;
  syncScopes(currentPage(),page);
  const pageTrans=pageTransitionStart(currentPage(),page);
  state.teamMode=page==='team';
  state.compareMode=page==='compare';
  state.statsMode=page==='stats';
  const alternate=page!=='individual';

  document.body.classList.toggle('team-mode',page==='team');
  document.body.classList.toggle('cmp-year',page==='compare' && cmp.scale==='year');
  document.body.classList.toggle('compare-mode',page==='compare');
  document.body.classList.toggle('stats-mode',page==='stats');
  closeStPanel();
  document.body.classList.toggle('cmp-week',page==='compare' && cmp.scale==='week');
  closeEmployeePicker();
  closeCmpPanel();

  const controls=document.querySelector('.controls');
  if(controls)controls.classList.toggle('hidden',alternate);
  lastInGroup[groupOf(page).id]=page;
  try{ localStorage.setItem('planningLastInGroup',JSON.stringify(lastInGroup)); }catch(e){ /* ignoré */ }
  if(alternate){
    if($('upcomingEventsCard'))$('upcomingEventsCard').classList.add('hidden');
  }
  // « En poste / Ensuite » : affiché tout de suite sur Individuel, masqué ailleurs.
  renderRotationWidget();
  const mainCard=document.querySelector('main.main-card');
  if(mainCard)mainCard.classList.toggle('hidden',alternate);
  $('teamViewSection').classList.toggle('hidden',page!=='team');
  $('compareViewSection').classList.toggle('hidden',page!=='compare');
  $('statsViewSection').classList.toggle('hidden',page!=='stats');

  updateTeamToggleLabel();
  try{ localStorage.setItem(PAGE_KEY,page); }catch(e){ /* stockage indisponible */ }
  histNote(page);
  render();
  pageTransitionRun(pageTrans);
}

function toggleTeamView(){ setPage(state.teamMode?'individual':'team'); }
function openTeamView(){ setPage('team'); }
function closeTeamView(){ setPage('individual'); }


// ==========================================================
// Présences annuelles : une ligne par personne, une colonne par semaine
// ISO. Chaque case compte les jours « présents » de la semaine :
//   vert   = 4 jours ou plus
//   orange = 1 à 3 jours
//   rouge  = aucun jour présent
// Les personnes sont regroupées par équipe/fonction telle que publiée par
// Excel (intitulés de la colonne A, voir CATEGORY_LABELS). Les équipes de
// rotation (Équipe 1 à 5) et le personnel hors-rotation (Admin., Cadres) ne
// sont volontairement jamais proposés dans cette vue (voir
// PRESENCE_EXCLUDED_CATEGORIES) : la rotation a déjà sa propre vue dédiée
// (« Planning d'équipe »), et l'admin/les cadres ne suivent pas de présence
// au sens de cette page. Les personnes masquées individuellement restent
// mémorisées dans ce navigateur uniquement.
// ==========================================================
const PRESENCE_CODES=new Set(['R','PO','CA','EN']); // codes comptés comme « présent » (RM/RA/RN exclus : ce sont des codes de rotation, pas de présence)
const PRESENCE_RESERVE_CODES=new Set(['PIQ']); // réserve = piquet uniquement ; RM/RA/RN ne comptent ni présent ni réserve (→ absent)
const PRESENCE_FULL_DAYS=4;                        // seuil de la case verte
const PRESENCE_SHOW_TITLES_KEY='presenceShowTitles2'; // titres de groupes : masqués par défaut (clé renouvelée)
const OTHER_GROUP_LABEL='Sans catégorie';
const EYE_OFF_ICON='<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"/><path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/><path d="M14.12 14.12a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>';

// Intitulés bruts publiés depuis la colonne A du classeur Excel (dans
// l'ordre où ils apparaissent, de haut en bas) -> titre affiché. Si Excel
// change un intitulé ou en ajoute un nouveau, il apparaît ici tel quel (pas
// d'erreur), simplement sans traduction ni ordre particulier -- il suffit de
// compléter ces deux tables.
const CATEGORY_ORDER=['Equipe 1','Equipe 2','Equipe 3','Equipe 4','Equipe 5','Elect','Autom.','Mécanique','Intendance','Chauffeurs-grutiers','Admin.','Cadres'];
const CATEGORY_LABELS={
  'Equipe 1':'Équipe 1','Equipe 2':'Équipe 2','Equipe 3':'Équipe 3','Equipe 4':'Équipe 4','Equipe 5':'Équipe 5',
  'Elect':'Électricien','Autom.':'Automaticien','Mécanique':'Mécanicien',
  'Intendance':'Intendance','Chauffeurs-grutiers':'Chauffeurs-grutiers',
  'Admin.':'Administration','Cadres':'Cadres'
};
// Catégories jamais proposées dans la vue Présences (masquées en
// permanence, sans bouton ni case pour les révéler).
const PRESENCE_EXCLUDED_CATEGORIES=new Set(['Equipe 1','Equipe 2','Equipe 3','Equipe 4','Equipe 5','Admin.','Cadres']);

// Comparaison tolérante (casse, accents, espaces, point final) : « Mécanique »,
// « mecanique » ou « Autom » doivent tomber dans le bon groupe.
function categoryKey(raw){
  return String(raw||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[.\s]+/g,' ').trim();
}
const CATEGORY_LABELS_NORMALIZED=Object.fromEntries(Object.entries(CATEGORY_LABELS).map(([k,v])=>[categoryKey(k),v]));
function categoryLabel(rawCategory){
  return CATEGORY_LABELS[rawCategory]||CATEGORY_LABELS_NORMALIZED[categoryKey(rawCategory)]||rawCategory||OTHER_GROUP_LABEL;
}
// Même ordre que CATEGORY_ORDER, mais sur les titres affichés (ce que les
// groupes utilisent une fois résolus via categoryLabel).
const CATEGORY_LABEL_ORDER=CATEGORY_ORDER.map(categoryLabel);
function categoryLabelOrderIndex(label){
  const i=CATEGORY_LABEL_ORDER.indexOf(label);
  return i===-1?CATEGORY_LABEL_ORDER.length:i;
}

const presence={
  showTitles:(()=>{ try{ return localStorage.getItem(PRESENCE_SHOW_TITLES_KEY)==='1'; }catch(e){ return false; } })(),
  rows:[],
  weeks:[]
};
function savePresenceTitles(){
  try{ localStorage.setItem(PRESENCE_SHOW_TITLES_KEY,presence.showTitles?'1':'0'); }catch(e){ /* stockage indisponible */ }
}

function presencePersonKey(emp){ return cleanEmployeeName(emp.name).toLocaleUpperCase('fr'); }
function rowNumberOf(emp){
  const match=String(emp.id||'').match(/(\d+)/);
  return match?Number(match[1]):Infinity;
}

// Groupe (titre affiché) de la personne, d'après la catégorie publiée par
// Excel. Une personne sans catégorie publiée (classeur pas encore republié
// avec la macro mise à jour) tombe dans OTHER_GROUP_LABEL plutôt que de
// faire planter l'affichage.
function presenceGroupOf(emp){
  return categoryLabel(emp.category);
}

const PRESENCE_EXCLUDED_KEYS=new Set([...PRESENCE_EXCLUDED_CATEGORIES].map(categoryKey));
function isPresenceExcludedCategory(emp){
  return PRESENCE_EXCLUDED_KEYS.has(categoryKey(emp.category));
}

// Groupes affichés, dans l'ordre du classeur Excel (CATEGORY_ORDER) ;
// Équipe 1-5 / Admin. / Cadres n'apparaissent jamais ici (voir
// PRESENCE_EXCLUDED_CATEGORIES).
function presenceGroups(includeExcluded=false){
  const byGroup=new Map();
  for(const emp of planning.employees){
    if(!includeExcluded && isPresenceExcludedCategory(emp))continue;
    const group=presenceGroupOf(emp);
    if(!byGroup.has(group))byGroup.set(group,[]);
    byGroup.get(group).push(emp);
  }
  return [...byGroup.entries()]
    .sort(([a],[b])=>categoryLabelOrderIndex(a)-categoryLabelOrderIndex(b) || a.localeCompare(b,'fr',{sensitivity:'base'}))
    .map(([name,people])=>({name,people}));
}

// Années pour lesquelles un fichier planning-<année>.json a été chargé.
function presenceYears(){
  const years=new Set();
  for(const emp of planning.employees){
    for(const key of Object.keys(emp.schedule))years.add(Number(key.slice(0,4)));
  }
  return [...years].filter(Boolean).sort((a,b)=>a-b);
}

// Semaines ISO de l'année : une semaine appartient à l'année de son jeudi.
function isoWeeksOfYear(year){
  const weeks=[];
  const monday=mondayOf(new Date(year,0,4));
  for(;;){
    const thursday=new Date(monday);thursday.setDate(monday.getDate()+3);
    if(thursday.getFullYear()!==year)break;
    const days=[];
    for(let i=0;i<7;i++){const d=new Date(monday);d.setDate(monday.getDate()+i);days.push(d);}
    weeks.push({num:weeks.length+1,monday:new Date(monday),thursday,days});
    monday.setDate(monday.getDate()+7);
  }
  return weeks;
}

function weekPresence(emp,week){
  // Seuls les 5 premiers jours de la semaine ISO (lundi à vendredi) comptent :
  // le week-end n'entre jamais dans le calcul de présence / réserve.
  const weekdays=week.days.slice(0,5);
  let present=0,reserve=0,known=0;
  const reserveBreakdown={}; // détail par code réel (PIQ, RM, RA, RN) pour l'info-bulle
  for(const d of weekdays){
    const key=isoDate(d);
    if(!Object.prototype.hasOwnProperty.call(emp.schedule,key))continue;
    known++;
    const code=safeCode(emp.schedule[key]);
    if(PRESENCE_CODES.has(code))present++;
    else if(PRESENCE_RESERVE_CODES.has(code)){
      reserve++;
      reserveBreakdown[code]=(reserveBreakdown[code]||0)+1;
    }
  }
  let level='none';
  if(known){
    if(present>=PRESENCE_FULL_DAYS)level='full';
    else if(reserve>0)level='reserve'; // RM/RA/RN/PIQ comptent comme jour de réserve (même couleur que présent)
    else if(present>0)level='partial';
    else level='absent';
  }
  return {present,reserve,reserveBreakdown,known,level};
}

// Texte précis du détail « réserve » à partir des vrais codes du planning
// (ex. « 2 j. Piquet, 1 j. Rotation nuit ») plutôt qu'un libellé générique.
function reserveBreakdownText(breakdown){
  return Object.keys(breakdown)
    .filter(code=>breakdown[code]>0)
    .map(code=>{
      const n=breakdown[code];
      const label=(CODE_INFO[code]&&CODE_INFO[code][0])||code;
      return `${n} j. ${label}`;
    })
    .join(', ');
}

function presenceWeekRange(week){
  const sunday=week.days[6];
  return `${formatShortDate(week.monday)} – ${formatDateWithYear(sunday)}`;
}

function isPhoneLayout(){ return window.matchMedia('(max-width:700px)').matches; }

// Feuille ouverte sur téléphone : la page en arrière-plan est figée (sinon le
// doigt fait défiler la page au lieu de la liste).
let scrollLockY=null;
function lockPageScroll(){
  if(scrollLockY!==null || !isPhoneLayout())return;
  scrollLockY=window.scrollY||document.documentElement.scrollTop||0;
  const b=document.body.style;
  b.position='fixed'; b.top=`-${scrollLockY}px`; b.left='0'; b.right='0'; b.width='100%';
  document.body.classList.add('pv-sheet-open');
}
function unlockPageScroll(){
  if(scrollLockY===null)return;
  const y=scrollLockY; scrollLockY=null;
  const b=document.body.style;
  b.position=''; b.top=''; b.left=''; b.right=''; b.width='';
  document.body.classList.remove('pv-sheet-open');
  window.scrollTo({top:y,left:0,behavior:'instant'});
}

function openPresenceWeekModal(emp,week){
  const p=weekPresence(emp,week);
  const name=cleanEmployeeName(emp.name);
  $('dayModalTitle').textContent=name;
  $('dayModalDate').textContent=`Semaine ${week.num} · ${presenceWeekRange(week)}`;

  const summary=p.level==='none'
    ? 'Pas de données pour cette semaine'
    : p.level==='reserve'
      ? `${reserveBreakdownText(p.reserveBreakdown)}${p.present?` + ${p.present} présent${p.present>1?'s':''}`:''} sur ${p.known} (lun-ven)`
      : `${p.present} jour${p.present>1?'s':''} présent${p.present>1?'s':''} sur ${p.known} (lun-ven)`;

  const daysHTML=week.days.map(d=>{
    const key=isoDate(d);
    const known=Object.prototype.hasOwnProperty.call(emp.schedule,key);
    const code=known?safeCode(emp.schedule[key]):'';
    const info=CODE_INFO[code];
    const label=!known?'Pas de données':code?(info?info[0]:code):'Pas de service';
    const holiday=swissHolidayName(key);
    const dayName=d.toLocaleDateString('fr-CH',{weekday:'short',day:'numeric',month:'short'});
    return `<div class="pv-day ${classFor(code)} ${PRESENCE_CODES.has(code)?'is-present':PRESENCE_RESERVE_CODES.has(code)?'is-reserve':''} ${known?'':'is-missing'}">
      <span class="pv-day-name">${escapeHTML(dayName)}</span>
      <span class="pv-day-code">${escapeHTML(code||'—')}</span>
      <span class="pv-day-label">${escapeHTML(label)}${holiday?` · ${escapeHTML(holiday)}`:''}</span>
    </div>`;
  }).join('');

  $('dayModalBody').innerHTML=`<div class="pv-week-summary pv-${p.level}">${escapeHTML(summary)}</div>${daysHTML}`;
  $('dayModalBackdrop').classList.toggle('is-team-day',false);$('dayModalBackdrop').classList.add('open');
}

// ==========================================================
// Page « Comparer » : plusieurs collègues côte à côte (semaine / mois /
// année) + ligne de couverture par métier (présents vs minimum).
// « Présent » = codes de présence ou piquet (PIQ), comme dans Maintenance.
// ==========================================================
const CMP_KEY='compareState';
const cmp=(()=>{
  let s={};
  try{ s=JSON.parse(localStorage.getItem(CMP_KEY)||'{}')||{}; }catch(e){ s={}; }
  // Remise à zéro unique de la sélection : une ancienne version y ajoutait
  // automatiquement la personne choisie dans les autres pages.
  try{
    if(localStorage.getItem('compareSelV')!=='2'){
      s.people=[];
      localStorage.setItem('compareSelV','2');
      localStorage.setItem(CMP_KEY,JSON.stringify({scale:s.scale,people:[]}));
    }
  }catch(e){ /* stockage indisponible */ }
  return {
    scale:['week','month','year'].includes(s.scale)?s.scale:'week',
    date:new Date(),
    people:new Set(Array.isArray(s.people)?s.people.map(String):[]),
    search:''
  };
})();
function saveCmp(){
  try{ localStorage.setItem(CMP_KEY,JSON.stringify({scale:cmp.scale,people:[...cmp.people]})); }catch(e){ /* stockage indisponible */ }
}

const cmpHas=(emp,iso)=>Object.prototype.hasOwnProperty.call(emp.schedule,iso);
function cmpIsPresent(emp,iso){
  const code=safeCode(emp.schedule[iso]);
  return PRESENCE_CODES.has(code)||PRESENCE_RESERVE_CODES.has(code);
}
// Colonnes de la période affichée.
function cmpColumns(){
  const today=todayISO();
  if(cmp.scale==='year'){
    const year=cmp.date.getFullYear();
    const weeks=isoWeeksOfYear(year);
    return {
      kind:'year',title:String(year),weeks,
      cols:weeks.map(w=>({week:w,label:String(w.num),today:today>=isoDate(w.monday)&&today<=isoDate(w.days[6])}))
    };
  }
  if(cmp.scale==='month'){
    const y=cmp.date.getFullYear(),m=cmp.date.getMonth();
    const n=new Date(y,m+1,0).getDate();
    const cols=[];
    for(let i=1;i<=n;i++){
      const d=new Date(y,m,i);const iso=isoDate(d);const dow=(d.getDay()+6)%7;
      cols.push({iso,d,dow,num:i,name:DAY_NAMES[dow][0],weekend:dow>=5,holiday:!!swissHolidayName(iso),holName:swissHolidayName(iso)||'',school:schoolHolidayName(iso)||'',today:iso===today});
    }
    return {kind:'day',title:`${MONTH_NAMES[m]} ${y}`,cols};
  }
  const monday=mondayOf(cmp.date);
  const cols=[];
  for(let i=0;i<7;i++){
    const d=new Date(monday);d.setDate(monday.getDate()+i);const iso=isoDate(d);
    cols.push({iso,d,dow:i,num:d.getDate(),name:DAY_NAMES[i],weekend:i>=5,holiday:!!swissHolidayName(iso),holName:swissHolidayName(iso)||'',school:schoolHolidayName(iso)||'',today:iso===today});
  }
  const sunday=cols[6].d;
  let title=`Semaine ${isoWeek(monday)} · ${formatShortDate(monday)} – ${formatDateWithYear(sunday)}`;
  if(isPhoneLayout()){
    const mo=d=>d.toLocaleDateString('fr-CH',{month:'short'});
    const range=monday.getMonth()===sunday.getMonth()
      ? `${monday.getDate()}–${sunday.getDate()} ${mo(sunday)}`
      : `${monday.getDate()} ${mo(monday)}–${sunday.getDate()} ${mo(sunday)}`;
    title=`Sem. ${isoWeek(monday)} · ${range} ${sunday.getFullYear()}`;
  }
  return {kind:'day',title,cols};
}

function cmpShift(dir){
  const d=new Date(cmp.date);
  if(cmp.scale==='year'){
    const years=presenceYears(),next=d.getFullYear()+dir;
    if(years.length && (next<years[0]||next>years[years.length-1]))return;
    d.setFullYear(next);
  }
  else if(cmp.scale==='month'){ d.setDate(1); d.setMonth(d.getMonth()+dir); }
  else d.setDate(d.getDate()+7*dir);
  cmp.date=d;
  renderCompareView();
}

const CMP_LEVEL_LABEL={full:'Présent 4 jours ou plus',reserve:'Piquet',partial:'Présent 1 à 3 jours',absent:'Aucun jour présent',none:'Pas de données'};

// Groupes (catégories) dans l'ordre d'affichage : métiers d'abord, puis
// Équipes / Admin. / Cadres.
function cmpGroupsOrdered(){
  return presenceGroups(true).sort((a,b)=>
    (isPresenceExcludedCategory(a.people[0])?1:0)-(isPresenceExcludedCategory(b.people[0])?1:0));
}

// Contenu d'une case pour une personne et un jour.
function cmpDayInfo(emp,c){
  const known=cmpHas(emp,c.iso);
  const code=known?safeCode(emp.schedule[c.iso]):'';
  const info=CODE_INFO[code];
  const label=!known?'Pas de données':code?(info?info[0]:code):'Pas de service';
  const title=`${cleanEmployeeName(emp.name)} · ${c.d.toLocaleDateString('fr-CH',{weekday:'long',day:'numeric',month:'long'})} : ${label}`;
  return {code,info,label,title,cls:known&&code?classFor(code):'code-off'};
}
function cmpYearInfo(emp,c){
  const p=weekPresence(emp,c.week);
  const label=p.level==='none'?'pas de données':p.level==='reserve'?`${reserveBreakdownText(p.reserveBreakdown)}${p.present?` + ${p.present} présent${p.present>1?'s':''}`:''}`:`${p.present} jour${p.present>1?'s':''} présent${p.present>1?'s':''} sur ${p.known}`;
  return {level:p.level,title:`${cleanEmployeeName(emp.name)} · semaine ${c.week.num} : ${label}`};
}

// Affichage téléphone : une carte par personne (pas de tableau à faire défiler).




// Pastilles de codes reliées entre jours voisins (dégradé quand le code change).
const PILL_VARS={R:'--c-r-bg',RM:'--c-rm-bg',RA:'--c-ra-bg',RN:'--c-rn-bg',PIQ:'--c-piq-bg',VA:'--c-va-bg',VJ:'--c-va-bg',HOR:'--c-hor-bg',HSR:'--c-hor-bg',RF:'--c-rf-bg',CA:'--c-ca-bg',GR:'--c-gr-bg',EN:'--c-en-bg'};
function pillVar(code){ if(!code)return ''; return PILL_VARS[code]||'--c-misc-bg'; }
function linkAttrs(prev,code,next){
  const own=pillVar(code); if(!own)return {cls:'',style:''};
  const p=pillVar(prev),n=pillVar(next);
  let cls='',st=`--pb:var(${own});`;
  if(p){ cls+=' lk-l'; if(p!==own)st+=`--pl:color-mix(in srgb,var(${own}) 50%,var(${p}) 50%);`; }
  if(n){ cls+=' lk-r'; if(n!==own)st+=`--pr:color-mix(in srgb,var(${own}) 50%,var(${n}) 50%);`; }
  return {cls,style:st};
}

// Position d'un jour dans une plage de vacances scolaires (pour tracer une capsule continue).
function schCls(iso){
  if(!schoolHolidayName(iso))return '';
  const d=parseISODateOnly(iso); if(!d)return '';
  const p=new Date(d);p.setDate(p.getDate()-1);const n=new Date(d);n.setDate(n.getDate()+1);
  const s=!schoolHolidayName(isoDate(p)),e=!schoolHolidayName(isoDate(n));
  const dow=(d.getDay()+6)%7;
  return 'sc'+(s?' sc-s':'')+(e?' sc-e':'')+((s||dow===0)?' sc-l':'')+((e||dow===6)?' sc-r':'');
}

// Infobulle au toucher sur les ronds « vacances scolaires » / « jour férié ».
function markText(iso){
  const h=swissHolidayName(iso),sc=schoolHolidayName(iso);
  const t=[]; if(h)t.push('Jour férié : '+h); if(sc)t.push('Vacances scolaires : '+sc);
  return t.join('\n');
}
function markAttr(iso){ const t=markText(iso); return t?` data-mark="${escapeHTML(t)}"`:''; }
(function(){
  let pop=null,timer=null;
  const close=()=>{ if(pop){pop.remove();pop=null;} clearTimeout(timer); };
  document.addEventListener('click',e=>{
    const t=e.target.closest('[data-mark]');
    if(!t){ close(); return; }
    e.stopPropagation(); e.preventDefault();
    close();
    pop=document.createElement('div'); pop.className='mark-pop'; pop.setAttribute('role','status');
    pop.innerHTML=t.dataset.mark.split('\n').map(x=>`<div>${escapeHTML(x)}</div>`).join('');
    document.body.appendChild(pop);
    const r=t.getBoundingClientRect(),pw=pop.offsetWidth;
    pop.style.left=Math.max(8,Math.min(window.innerWidth-pw-8,r.left+r.width/2-pw/2))+'px';
    const below=r.top<80;
    pop.style.top=(below?r.bottom+8:r.top-pop.offsetHeight-8)+'px';
    timer=setTimeout(close,3500);
  },true);
  window.addEventListener('scroll',close,{passive:true});
})();

// Ruban « période » : vacances scolaires (bande violette) et jours fériés (points rouges)
// sur toute la période affichée. Hauteur fixe, même quand il est vide.
function periodStripHTML(isos){
  return '';
  const n=isos.length; if(!n)return '';
  let segs='',i=0,any=false;
  while(i<n){
    const nm=schoolHolidayName(isos[i])||'';
    let j=i; while(j+1<n&&(schoolHolidayName(isos[j+1])||'')===nm)j++;
    if(nm){ any=true; segs+=`<span class="ps-sch" style="grid-column:${i+1}/span ${j-i+1}" title="${escapeHTML(nm)}"><b>${escapeHTML(nm)}</b></span>`; }
    i=j+1;
  }
  isos.forEach((iso,k)=>{ const h=swissHolidayName(iso); if(h){ any=true; segs+=`<span class="ps-hol" style="grid-column:${k+1}" title="${escapeHTML(h)}"><i></i></span>`; } });
  return `<div class="ps${any?' has-ev':''}" style="--ps-n:${n}" aria-label="Vacances scolaires et jours fériés">${segs}</div>`;
}
// Légende sous le calendrier : noms des vacances et fériés de la période.
function periodNotesHTML(isos){
  const sch=[...new Set(isos.map(i=>schoolHolidayName(i)).filter(Boolean))];
  const hol=isos.map(i=>({i,n:swissHolidayName(i)})).filter(x=>x.n);
  if(!sch.length&&!hol.length)return '';
  const fmt=iso=>{const d=parseISODateOnly(iso);return `${DAY_NAMES[(d.getDay()+6)%7].slice(0,3).toLowerCase()}. ${d.getDate()}`;};
  return '<div class="cmp-notes">'+sch.map(n=>`<span class="cmp-note sch">🏫 ${escapeHTML(n)}</span>`).join('')+hol.map(x=>`<span class="cmp-note hol"><i></i>${fmt(x.i)} · ${escapeHTML(x.n)}</span>`).join('')+'</div>';
}

function cmpNotesHTML(view){
  if(view.kind!=='year')return periodNotesHTML(view.cols.map(c=>c.iso));
  if(view.kind==='year')return '';
  const sch=[...new Set(view.cols.map(c=>c.school).filter(Boolean))];
  const hol=view.cols.filter(c=>c.holiday);
  if(!sch.length&&!hol.length)return '';
  const fmt=c=>`${c.name.slice(0,3).toLowerCase()}. ${c.num}`;
  return '<div class="cmp-notes">'
    +sch.map(n=>`<span class="cmp-note sch">🏫 ${escapeHTML(n)}</span>`).join('')
    +hol.map(c=>`<span class="cmp-note hol"><i></i>${fmt(c)} · ${escapeHTML(c.holName)}</span>`).join('')
    +'</div>';
}
function cmpCardsHTML(view,groups){
  // Tout le monde dans un seul bloc, sans séparation par groupe.
  const everyone=groups.flatMap(g=>g.people).filter(e=>cmp.people.has(presencePersonKey(e)));
  let h='';
  {
    const people=everyone;
    h+='<div class="cmpm-block'+(view.kind!=='year'&&cmp.scale==='week'?' is-wk':'')+'">';
    if(view.kind!=='year'&&cmp.scale==='week')h+='<div class="cmpm-weekhead"><span></span>'+view.cols.map(c=>`<span class="${c.today?'is-today':''} ${c.holiday?'is-hol':''} ${c.school?'is-sch '+schCls(c.iso):''}">${c.name}<em${markAttr(c.iso)}>${c.num}</em></span>`).join('')+'</div>';
    for(const emp of people){
      const name=cleanEmployeeName(emp.name);
      h+=`<div class="cmpm-card"><div class="cmpm-name" data-open-emp="${escapeHTML(emp.id)}">${escapeHTML(name)}</div>`;
      if(view.kind==='year'){
        h+='<div class="cmpm-year">'+view.cols.map(c=>{
          const y=cmpYearInfo(emp,c);
          return `<i class="pv-${y.level} ${c.today?'is-today':''}" title="${escapeHTML(y.title)}"></i>`;
        }).join('')+'</div>';
      }else{
        const lead=cmp.scale==='month'?view.cols[0].dow:0;
        h+=`<div class="cmpm-days ${cmp.scale==='week'?'is-week':'is-month'}">`;
        if(cmp.scale==='month')h+=DAY_NAMES.map(n=>`<span class="cmpm-dow">${n[0]}</span>`).join('');
        for(let i=0;i<lead;i++)h+='<span></span>';
        const infos=view.cols.map(c=>cmpDayInfo(emp,c));
        h+=view.cols.map((c,ci)=>{
          const d=infos[ci];
          const lkc=linkAttrs(c.dow>0&&infos[ci-1]?infos[ci-1].code:'',d.code,c.dow<6&&infos[ci+1]?infos[ci+1].code:'');
          const tap=cmp.scale==='week'?` data-cmp-key="${escapeHTML(presencePersonKey(emp))}" data-cmp-iso="${c.iso}" role="button" tabindex="0"`:'';
          return `<div class="cmpm-day ${d.cls} ${c.weekend?'is-we':''} ${c.today?'is-today':''} ${c.holiday?'is-hol':''} ${c.school?'is-sch '+schCls(c.iso):''}" title="${escapeHTML(d.title)}"${tap}><small${markAttr(c.iso)}>${cmp.scale==='week'?c.name:c.num}</small><b class="${lkc.cls.trim()}" style="${lkc.style}">${escapeHTML(d.code)}</b></div>`;
        }).join('');
        h+='</div>';
      }
      h+='</div>';
    }
    h+='</div>';
  }
  return h;
}

// Personnes « de métier » (hors Équipes, Admin., Cadres) : sélection rapide « Tous ».
function cmpMainPeople(){ return presenceGroups(false).flatMap(g=>g.people); }
function cmpSetPeople(people,on){
  for(const e of people){
    const k=presencePersonKey(e);
    if(on)cmp.people.add(k); else cmp.people.delete(k);
  }
  saveCmp();
  renderCompareView();
}

function renderCompareView(){
  if(!$('compareViewSection'))return;
  const view=cmpColumns();
  const isYear=cmp.scale==='year';
  document.body.classList.toggle('cmp-week',cmp.scale==='week');
  document.body.classList.toggle('cmp-year',isYear);
  $('cmpTitle').textContent=view.title;
  $('cmpPrintYear').textContent=String(cmp.date.getFullYear());
  document.querySelectorAll('[data-cmp-scale]').forEach(b=>b.classList.toggle('active',b.dataset.cmpScale===cmp.scale));
  $('cmpCount').textContent=String(cmp.people.size);

  // Outils propres à l'année (anciennement page « Maintenance »).
  for(const id of ['cmpTitlesBtn','cmpExportBtn','cmpLegend'])$(id).classList.toggle('hidden',!isYear);
  const tb=$('cmpTitlesBtn');
  tb.setAttribute('aria-pressed',String(presence.showTitles));
  tb.classList.toggle('is-off',!presence.showTitles);
  const years=presenceYears(),yr=cmp.date.getFullYear();
  $('cmpPrevBtn').disabled=isYear && years.length>0 && yr<=years[0];
  $('cmpNextBtn').disabled=isYear && years.length>0 && yr>=years[years.length-1];
  const grid=$('cmpGrid');
  grid.classList.toggle('pv-wrap',isYear);
  grid.classList.toggle('cmp-wrap',!isYear);

  const groups=cmpGroupsOrdered();

  // Raccourcis : « Tous » (métiers) puis un par groupe.
  const main=cmpMainPeople();
  const allOn=main.length>0 && main.every(e=>cmp.people.has(presencePersonKey(e)));
  let chips=`<button type="button" class="pv-chip ${allOn?'is-on':''}" data-cmp-all aria-pressed="${allOn}" title="Tout le personnel de métier (hors Équipes, Admin., Cadres)">Tous</button>`;
  chips+=groups.map(g=>{
    const n=g.people.filter(e=>cmp.people.has(presencePersonKey(e))).length;
    const cls=n===g.people.length?'is-on':n>0?'is-partial':'';
    return `<button type="button" class="pv-chip ${cls}" data-cmp-grp="${escapeHTML(g.name)}" aria-pressed="${n===g.people.length}" title="Ajouter / retirer tout le groupe">${escapeHTML(g.name)}</button>`;
  }).join('');
  // Groupes enregistrés en premier, tout à gauche.
  $('cmpGroups').innerHTML=cmpSavedChipsHTML()+chips+cmpSaveChipHTML();

  const selected=groups.flatMap(g=>g.people).filter(e=>cmp.people.has(presencePersonKey(e)));
  if(!$('cmpPanel').classList.contains('hidden'))renderCmpPanelList();

  if(!selected.length){
    presence.rows=[];
    grid.innerHTML='<div class="pv-empty-state">Aucune personne sélectionnée.<br>Touchez un groupe ci-dessus, ou « Collègues » pour choisir des personnes.<br><button type="button" class="pv-link" data-cmp-all>Afficher tout le personnel de métier</button></div>';
    return;
  }

  if(isYear){ renderCmpYear(view,groups); return; }

  if(isPhoneLayout()){
    grid.innerHTML=`<div class="cmpm">${periodStripHTML(view.cols.map(c=>c.iso))}${cmpCardsHTML(view,groups)}${cmpNotesHTML(view)}</div>`;
    return;
  }

  // La table occupe toute la largeur disponible ; les colonnes de jours se
  // partagent l'espace à parts égales (largeur minimale = défilement horizontal).
  const minW=cmp.scale==='week'?96:34;
  let h=periodStripHTML(view.cols.map(c=>c.iso))+`<table class="cmp-table" style="width:100%;min-width:${190+view.cols.length*(minW+2)+2}px"><colgroup><col style="width:190px">${view.cols.map(()=>'<col>').join('')}</colgroup><thead><tr><th class="cmp-name"><div class="cmp-name-inner"></div></th>`;
  h+=view.cols.map(c=>`<th class="cmp-day ${c.weekend?'is-we':''} ${c.holiday?'is-hol':''} ${c.school?'is-sch '+schCls(c.iso):''} ${c.today?'is-today':''}" ${c.holiday?`title="${escapeHTML(c.holName)}"`:c.school?`title="${escapeHTML(c.school)}"`:''}><span>${c.name}</span><b${markAttr(c.iso)}>${c.num}</b></th>`).join('');
  h+='</tr></thead><tbody>';
  for(const group of groups){
    const people=group.people.filter(e=>cmp.people.has(presencePersonKey(e)));
    for(const emp of people){
      const name=cleanEmployeeName(emp.name);
      h+=`<tr class="cmp-person"><th class="cmp-name"><div class="cmp-name-inner"><span title="${escapeHTML(name)}" data-open-emp="${escapeHTML(emp.id)}">${escapeHTML(name)}</span></div></th>`;
      const wide=cmp.scale==='week';
      h+=view.cols.map(c=>{
        const d=cmpDayInfo(emp,c);
        const tap=wide?` data-cmp-key="${escapeHTML(presencePersonKey(emp))}" data-cmp-iso="${c.iso}" role="button" tabindex="0"`:'';
        return `<td class="cmp-cell ${d.cls} ${c.weekend?'is-we':''} ${c.today?'is-today':''}" title="${escapeHTML(d.title)}"${tap}><b>${escapeHTML(d.code)}</b>${wide&&d.code&&d.info?`<small>${escapeHTML(d.info[0])}</small>`:''}</td>`;
      }).join('');
      h+='</tr>';
    }
  }
  h+='</tbody></table>';
  grid.innerHTML=h+cmpNotesHTML(view);
}

// Vue « Année » : une case par personne et par semaine ISO (vert = présent
// 4 jours ou plus / piquet, orange = 1 à 3 jours, rouge = aucun), avec le
// nombre de personnes en vert en bas de chaque semaine.
function renderCmpYear(view,groups){
  const year=cmp.date.getFullYear();
  const weeks=view.weeks;
  presence.weeks=weeks;
  const today=new Date();
  const currentWeekIndex=weeks.findIndex(w=>today>=w.monday && today<new Date(w.monday.getFullYear(),w.monday.getMonth(),w.monday.getDate()+7));

  const monthSpans=[];
  for(const w of weeks){
    const m=w.thursday.getMonth();
    const last=monthSpans[monthSpans.length-1];
    if(last && last.month===m)last.span++;
    else monthSpans.push({month:m,span:1});
  }

  let h='<table class="pv-table"><thead><tr><th class="pv-name" rowspan="2"><div class="pv-name-inner"><span>Semaine</span></div></th>';
  h+=monthSpans.map(ms=>`<th class="pv-month" colspan="${ms.span}">${escapeHTML(MONTH_NAMES[ms.month])}</th>`).join('');
  h+='</tr><tr>';
  h+=weeks.map((w,i)=>`<th class="pv-week ${i===currentWeekIndex?'is-current':''}" data-pv-wk="${i}" title="Semaine ${w.num} · ${escapeHTML(presenceWeekRange(w))}">${w.num}</th>`).join('');
  h+='</tr></thead><tbody>';

  const rows=[];
  const fullCounts=weeks.map(()=>0);
  for(const group of groups){
    const people=group.people.filter(e=>cmp.people.has(presencePersonKey(e)));
    if(!people.length)continue;
    if(presence.showTitles){
      h+=`<tr class="pv-group"><th class="pv-name"><div class="pv-name-inner">${escapeHTML(group.name)}</div></th><td colspan="${weeks.length}"></td></tr>`;
    }
    for(const emp of people){
      const rowIndex=rows.length;
      rows.push(emp);
      const name=cleanEmployeeName(emp.name);
      h+=`<tr class="pv-person"><th class="pv-name"><div class="pv-name-inner"><span title="${escapeHTML(name)}" data-open-emp="${escapeHTML(emp.id)}">${escapeHTML(name)}</span><button type="button" class="pv-hide" data-pv-hide="${escapeHTML(presencePersonKey(emp))}" title="Retirer ${escapeHTML(name)}" aria-label="Retirer ${escapeHTML(name)}">${EYE_OFF_ICON}</button></div></th>`;
      weeks.forEach((w,i)=>{
        const p=weekPresence(emp,w);
        if(p.level==='full'||p.level==='reserve')fullCounts[i]++; // = cases vertes
        const label=p.level==='none'
          ? 'pas de données'
          : p.level==='reserve'
            ? `${reserveBreakdownText(p.reserveBreakdown)}${p.present?` + ${p.present} présent${p.present>1?'s':''}`:''}${p.known<5?` (données sur ${p.known} j)`:''}`
            : `${p.present} jour${p.present>1?'s':''} présent${p.present>1?'s':''}${p.known<5?` (données sur ${p.known} j)`:''}`;
        h+=`<td class="pv-cell pv-${p.level} ${i===currentWeekIndex?'is-current':''}" data-pv-row="${rowIndex}" data-pv-week="${i}" tabindex="-1" title="${escapeHTML(name)} · S${w.num} · ${escapeHTML(label)}"></td>`;
      });
      h+='</tr>';
    }
  }
  h+='</tbody><tfoot><tr><th class="pv-name"><div class="pv-name-inner" title="Nombre de personnes affichées en vert cette semaine (présentes au moins 4 jours, ou de piquet)"><span>Présents</span></div></th>';
  h+=fullCounts.map(n=>`<td class="pv-total">${n}</td>`).join('');
  h+='</tr></tfoot></table>';

  presence.rows=rows;
  const wrap=$('cmpGrid');
  const scrollLeft=wrap.scrollLeft,scrollTop=wrap.scrollTop;
  const firstRender=!wrap.querySelector('.pv-table') || wrap.dataset.year!==String(year);
  // Largeur de la colonne des noms : ajustée au nom le plus long affiché.
  try{
    const phone=window.matchMedia('(max-width:700px)').matches;
    const cv=document.createElement('canvas').getContext('2d');
    cv.font=`600 ${phone?10.5:11.5}px Inter, system-ui, sans-serif`;
    let wmax=0;
    for(const emp of rows){ wmax=Math.max(wmax,cv.measureText(cleanEmployeeName(emp.name)).width); }
    wmax=Math.max(wmax,cv.measureText('Semaine').width,cv.measureText('Présents').width);
    const extra=phone?14:8+4+22+4+6;
    const w=Math.round(Math.min(phone?150:220,Math.max(phone?66:110,wmax+extra)));
    wrap.style.setProperty('--pv-name',w+'px');
  }catch(e){ /* ignoré : largeur par défaut */ }
  wrap.innerHTML=h;
  wrap.dataset.year=String(year);
  if(firstRender && currentWeekIndex>=0){
    // Centre la semaine en cours à l'ouverture.
    const cell=wrap.querySelector('.pv-week.is-current');
    if(cell){
      const nameWidth=wrap.querySelector('thead .pv-name').offsetWidth;
      wrap.scrollLeft=Math.max(0,cell.offsetLeft-nameWidth-(wrap.clientWidth-nameWidth)/2);
    }
  }else{
    wrap.scrollLeft=scrollLeft;
    wrap.scrollTop=scrollTop;
  }
}

function renderCmpPanelList(){
  const list=$('cmpPanelList');
  const query=normalizeSearchText(cmp.search);
  let h='';
  let extrasTitleDone=false;
  for(const group of cmpGroupsOrdered()){
    const people=group.people.filter(e=>!query||normalizeSearchText(e.name).includes(query));
    if(!people.length)continue;
    if(isPresenceExcludedCategory(group.people[0]) && !extrasTitleDone){
      extrasTitleDone=true;
      h+='<div class="pv-panel-sep">Équipes, Admin. et Cadres</div>';
    }
    const n=people.filter(e=>cmp.people.has(presencePersonKey(e))).length;
    h+=`<div class="pv-panel-group"><label><input type="checkbox" data-cmp-cat="${escapeHTML(group.name)}" data-state="${n===0?'none':n===people.length?'all':'some'}" ${n===people.length?'checked':''}>${escapeHTML(group.name)}</label><small>${n}/${people.length}</small></div>`;
    for(const e of people){
      const key=presencePersonKey(e);
      const on=cmp.people.has(key);
      h+=`<label class="pv-panel-person ${on?'':'is-off'}"><input type="checkbox" data-cmp-person="${escapeHTML(key)}" ${on?'checked':''}>${escapeHTML(cleanEmployeeName(e.name))}</label>`;
    }
  }
  list.innerHTML=h||'<div class="pv-panel-empty">Aucune personne trouvée.</div>';
  list.querySelectorAll('input[data-state="some"]').forEach(i=>{ i.indeterminate=true; });
}
function openCmpPanel(){
  $('cmpPanel').classList.remove('hidden');
  $('cmpBackdrop').classList.toggle('hidden',!isPhoneLayout());
  lockPageScroll();
  $('cmpPickBtn').setAttribute('aria-expanded','true');
  renderCmpPanelList();
  if(window.matchMedia('(hover:hover)').matches)$('cmpSearch').focus();
}
function closeCmpPanel(){
  if(!$('cmpPanel'))return;
  const wasOpen=!$('cmpPanel').classList.contains('hidden');
  $('cmpPanel').classList.add('hidden');
  $('cmpBackdrop').classList.add('hidden');
  if(wasOpen)unlockPageScroll();
  $('cmpPickBtn').setAttribute('aria-expanded','false');
}

// Survol d'une case de l'année : le numéro de la semaine est mis en surbrillance.
(function(){
  const grid=$('cmpGrid');
  const setWeek=idx=>{
    grid.querySelectorAll('.pv-week.is-hover').forEach(el=>el.classList.remove('is-hover'));
    if(idx===null)return;
    const th=grid.querySelector(`.pv-week[data-pv-wk="${idx}"]`);
    if(th)th.classList.add('is-hover');
  };
  grid.addEventListener('mouseover',e=>{
    const cell=e.target.closest('[data-pv-week]');
    setWeek(cell?cell.dataset.pvWeek:null);
  });
  grid.addEventListener('mouseleave',()=>setWeek(null));
})();

$('cmpPrevBtn').addEventListener('click',()=>cmpShift(-1));
$('cmpNextBtn').addEventListener('click',()=>cmpShift(1));
$('cmpTodayBtn').addEventListener('click',()=>{ cmp.date=new Date(); renderCompareView(); });
$('cmpPickBtn').addEventListener('click',event=>{
  event.stopPropagation();
  if($('cmpPanel').classList.contains('hidden'))openCmpPanel(); else closeCmpPanel();
});
$('cmpPanelDone').addEventListener('click',event=>{ event.stopPropagation(); closeCmpPanel(); });
$('cmpSearch').addEventListener('input',event=>{ cmp.search=event.target.value; renderCmpPanelList(); });
$('cmpClearBtn').addEventListener('click',()=>{ cmp.people.clear(); saveCmp(); renderCompareView(); });
document.addEventListener('click',event=>{
  const f=$('cmpFilter');
  if(f && !f.contains(event.target))closeCmpPanel();
});
$('compareViewSection').addEventListener('click',event=>{
  const scaleBtn=event.target.closest('[data-cmp-scale]');
  if(scaleBtn){ cmp.scale=scaleBtn.dataset.cmpScale; saveCmp(); renderCompareView(); return; }
  if(event.target.closest('#cmpExportBtn')){ window.print(); return; }
  if(event.target.closest('#cmpTitlesBtn')){ presence.showTitles=!presence.showTitles; savePresenceTitles(); renderCompareView(); return; }
  if(event.target.closest('[data-cmp-all]')){
    const main=cmpMainPeople();
    const allOn=main.length>0 && main.every(e=>cmp.people.has(presencePersonKey(e)));
    cmpSetPeople(main,!allOn || !!event.target.closest('.pv-empty-state'));
    return;
  }
  const bulk=event.target.closest('[data-cmp-bulk]');
  if(bulk){
    const query=normalizeSearchText(cmp.search);
    // Sans recherche : seulement les métiers ; avec une recherche : toute personne trouvée.
    const people=presenceGroups(!!query).flatMap(g=>g.people).filter(e=>!query||normalizeSearchText(e.name).includes(query));
    cmpSetPeople(people,bulk.dataset.cmpBulk==='show');
    return;
  }
  const hideBtn=event.target.closest('[data-pv-hide]');
  if(hideBtn){ cmp.people.delete(hideBtn.dataset.pvHide); saveCmp(); renderCompareView(); return; }
  const yearCell=event.target.closest('[data-pv-row]');
  if(yearCell){
    const emp=presence.rows[Number(yearCell.dataset.pvRow)];
    const week=presence.weeks[Number(yearCell.dataset.pvWeek)];
    if(emp && week)openPresenceWeekModal(emp,week);
    return;
  }
  const grp=event.target.closest('[data-cmp-grp]');
  if(grp){
    const g=presenceGroups(true).find(x=>x.name===grp.dataset.cmpGrp);
    if(g){
      const allOn=g.people.every(e=>cmp.people.has(presencePersonKey(e)));
      for(const e of g.people){
        const k=presencePersonKey(e);
        if(allOn)cmp.people.delete(k); else cmp.people.add(k);
      }
      saveCmp(); renderCompareView();
    }
    return;
  }
  const rm=event.target.closest('[data-cmp-remove]');
  if(rm){ cmp.people.delete(rm.dataset.cmpRemove); saveCmp(); renderCompareView(); }
});
$('compareViewSection').addEventListener('change',event=>{
  const t=event.target;
  if(t.dataset.cmpCat!==undefined){
    const query=normalizeSearchText(cmp.search);
    const group=presenceGroups(true).find(g=>g.name===t.dataset.cmpCat);
    if(group)for(const e of group.people.filter(e=>!query||normalizeSearchText(e.name).includes(query))){
      const key=presencePersonKey(e);
      if(t.checked)cmp.people.add(key); else cmp.people.delete(key);
    }
  }else if(t.dataset.cmpPerson!==undefined){
    if(t.checked){
      // La sélection du Comparer est indépendante de la personne du planning individuel.
      cmp.people.add(t.dataset.cmpPerson);
    }else cmp.people.delete(t.dataset.cmpPerson);
  }else{
    return;
  }
  saveCmp();
  renderCompareView();
});
document.addEventListener('keydown',event=>{ if(event.key==='Escape')closeCmpPanel(); });
// ==========================================================
// Page Statistiques : fiche personnelle (aucune comparaison entre personnes).
// Compte les jours où chaque code apparaît dans le planning de la personne,
// sur l'année ou sur un mois. Les blocs Rotations / Piquet n'apparaissent que
// si la personne en a dans son année.
// ==========================================================
const ST_KEY='statsState';
const ST_VAC=['VA','VJ','PI','HSR','HOR']; // « congés » = vacances + récupérations (HSR, HOR)
const ST_REC=[];
const ST_AFF=['CA','PO','EN','GR']; // affectations : Camion, Portes, Entretien, Groupe
const ST_OTHER=['MS','MC','AP','MI','EM','PC','FE','PR','AN','NS','PA','DE','DL','DC','JT'];
const ST_COLORS={vac:'#E0566A',oth:'#6B7A90',rot:'#4545AA',piq:'#63428E',aff:'#3E8E6B'};
const st={month:'all',year:0,metric:'',search:'',mode:'month',more:false};
// Année affichée : par défaut l'année en cours ; les autres années ne sont
// proposées que si un fichier planning-<année>.json les contient.
function stYear(){
  const ys=presenceYears();
  return ys.includes(st.year)?st.year:(ys.includes(planning.year)?planning.year:(ys[ys.length-1]||planning.year));
}
const ST_MON_SHORT=['Jan','Fév','Mar','Avr','Mai','Juin','Juil','Aoû','Sep','Oct','Nov','Déc'];
try{ const m=JSON.parse(localStorage.getItem(ST_KEY)||'{}'); if(m&&(m.month==='all'||(Number.isInteger(m.month)&&m.month>=0&&m.month<12)))st.month=m.month; }catch(e){ /* ignoré */ }
function saveSt(){ try{ localStorage.setItem(ST_KEY,JSON.stringify({month:st.month})); }catch(e){ /* ignoré */ } }

function stCompute(emp){
  const year=String(stYear());
  const months=Array.from({length:12},()=>({}));
  const piqDays={weekend:0,total:0};
  for(const [iso,raw] of Object.entries(emp.schedule||{})){
    if(!iso.startsWith(year+'-'))continue;
    const code=safeCode(raw);
    if(!code)continue;
    const m=Number(iso.slice(5,7))-1;
    if(!(m>=0&&m<12))continue;
    months[m][code]=(months[m][code]||0)+1;
  }
  return months;
}
function stSum(months,codes,monthSel){
  let n=0;
  months.forEach((mo,i)=>{ if(monthSel!=='all'&&monthSel!==i)return; for(const c of codes)n+=mo[c]||0; });
  return n;
}
function stPiquetOffDays(emp,monthSel){
  const year=String(stYear());
  let n=0;
  for(const [iso,raw] of Object.entries(emp.schedule||{})){
    if(!iso.startsWith(year+'-') || safeCode(raw)!=='PIQ')continue;
    const m=Number(iso.slice(5,7))-1;
    if(monthSel!=='all'&&monthSel!==m)continue;
    const d=new Date(Number(iso.slice(0,4)),m,Number(iso.slice(8,10)));
    if(d.getDay()===0||d.getDay()===6||swissHolidayName(iso))n++;
  }
  return n;
}
const stDays=n=>`${n} jour${n>1?'s':''}`;

function stPeriodLabel(){
  return st.month==='all'?`l'année ${stYear()}`:`${MONTH_NAMES[st.month].toLowerCase()} ${stYear()}`;
}

function stBarRows(rows,total){
  const max=Math.max(1,...rows.map(r=>r.n));
  return rows.map(r=>`<div class="st-row"><span>${r.label}</span><div class="st-track"><i class="st-fill" style="width:${Math.round(r.n/max*100)}%;background:${r.color}"></i></div><em>${r.n}</em></div>`).join('');
}

function stSmoothPath(pts){
  const n=pts.length;
  if(n<2)return '';
  const dx=[],d=[],m=new Array(n);
  for(let i=0;i<n-1;i++){ dx[i]=pts[i+1][0]-pts[i][0]; d[i]=(pts[i+1][1]-pts[i][1])/dx[i]; }
  m[0]=d[0]; m[n-1]=d[n-2];
  for(let i=1;i<n-1;i++)m[i]=(d[i-1]*d[i]<=0)?0:(d[i-1]+d[i])/2;
  for(let i=0;i<n-1;i++){
    if(d[i]===0){ m[i]=0; m[i+1]=0; continue; }
    const a=m[i]/d[i],b=m[i+1]/d[i],t=a*a+b*b;
    if(t>9){ const k=3/Math.sqrt(t); m[i]=k*a*d[i]; m[i+1]=k*b*d[i]; }
  }
  let path=`M${pts[0][0].toFixed(1)},${pts[0][1].toFixed(1)}`;
  for(let i=0;i<n-1;i++){
    path+=` C${(pts[i][0]+dx[i]/3).toFixed(1)},${(pts[i][1]+m[i]*dx[i]/3).toFixed(1)} ${(pts[i+1][0]-dx[i]/3).toFixed(1)},${(pts[i+1][1]-m[i+1]*dx[i]/3).toFixed(1)} ${pts[i+1][0].toFixed(1)},${pts[i+1][1].toFixed(1)}`;
  }
  return path;
}

function stChartHTML(months,metrics){
  const metric=metrics.find(x=>x.id===st.metric)||metrics[0];
  const cum=st.mode==='cum';
  const per=months.map(mo=>metric.codes.reduce((a,c)=>a+(mo[c]||0),0));
  let acc=0;
  const vals=cum?per.map(v=>(acc+=v)):per;
  const niceMax=v=>{ if(v<=4)return 4; const step=Math.pow(10,Math.floor(Math.log10(v/4))); const r=v/4/step; const k=r<=1?1:r<=2?2:r<=2.5?2.5:r<=5?5:10; return Math.ceil(v/(k*step))*k*step/1; };
  const top=niceMax(Math.max(1,...vals));
  const W=640,H=240,L=34,R=14,T=32,B=28,cw=(W-L-R)/12,ch=H-T-B;
  const X=i=>L+(i+0.5)*cw, Y=v=>T+ch-(v/top)*ch;
  const pts=vals.map((v,i)=>[X(i),Y(v)]);
  const line=stSmoothPath(pts);
  const area=`${line} L${X(11).toFixed(1)},${(T+ch).toFixed(1)} L${X(0).toFixed(1)},${(T+ch).toFixed(1)} Z`;
  const now=new Date();
  const frac=stYear()<now.getFullYear()?12:stYear()>now.getFullYear()?0:(now.getMonth()+(now.getDate()-1)/new Date(now.getFullYear(),now.getMonth()+1,0).getDate());
  const xNow=L+frac*cw;
  const ticks=[0,1,2,3,4].map(k=>top*k/4);
  const grid=ticks.map(v=>`<line x1="${L}" x2="${W-R}" y1="${Y(v).toFixed(1)}" y2="${Y(v).toFixed(1)}" class="st-grid"/><text x="${L-6}" y="${(Y(v)+3.5).toFixed(1)}" text-anchor="end" class="st-axis">${Number.isInteger(v)?v:v.toFixed(1)}</text>`).join('');
  const labs=ST_MON_SHORT.map((l,i)=>`<text x="${X(i).toFixed(1)}" y="${H-8}" text-anchor="middle" class="st-axis ${st.month===i?'is-sel':''}">${l}</text>`).join('');
  const dots=vals.map((v,i)=>{
    const sel=st.month===i;
    return `<circle cx="${X(i).toFixed(1)}" cy="${Y(v).toFixed(1)}" r="${sel?5:3}" fill="${sel?'#fff':metric.color}" stroke="${metric.color}" stroke-width="${sel?3:0}"/>`+
      ((per[i]>0||cum&&i===11)?`<text x="${X(i).toFixed(1)}" y="${(Y(v)-9).toFixed(1)}" text-anchor="middle" class="st-val">${v}</text>`:'');
  }).join('');
  const nowLine=(xNow>L&&xNow<W-R)?`<line x1="${xNow.toFixed(1)}" x2="${xNow.toFixed(1)}" y1="${T-16}" y2="${T+ch}" class="st-now"/><text x="${xNow.toFixed(1)}" y="${T-20}" text-anchor="middle" class="st-nowlab">Aujourd'hui</text>`:'';
  const chips=metrics.map(x=>`<button type="button" class="pv-chip ${x.id===metric.id?'is-on':''}" data-st-metric="${x.id}" aria-pressed="${x.id===metric.id}">${x.label}</button>`).join('');
  const seg=`<div class="cmp-tabs" role="tablist" aria-label="Affichage"><button type="button" class="cmp-tab ${cum?'':'active'}" data-st-mode="month">Par mois</button><button type="button" class="cmp-tab ${cum?'active':''}" data-st-mode="cum">Cumul</button></div>`;
  const title=cum?`${metric.label} cumulés depuis janvier`:`${metric.label} par mois`;
  const id=`stc${metric.id}`;
  return `<div class="st-block st-curve" style="--st-c:${metric.color}">
    <div class="st-chart-head"><h3>${title}</h3>${seg}</div>
    <div class="pv-chips st-metrics">${chips}</div>
    <svg viewBox="0 0 ${W} ${H}" class="st-svg" role="img" aria-label="${title}" preserveAspectRatio="xMidYMid meet">
      <defs>
        <linearGradient id="${id}g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${metric.color}" stop-opacity=".28"/><stop offset="1" stop-color="${metric.color}" stop-opacity="0"/></linearGradient>
        <clipPath id="${id}p"><rect x="0" y="0" width="${xNow.toFixed(1)}" height="${H}"/></clipPath>
        <clipPath id="${id}f"><rect x="${xNow.toFixed(1)}" y="0" width="${(W-xNow).toFixed(1)}" height="${H}"/></clipPath>
      </defs>
      ${grid}
      <path d="${area}" fill="url(#${id}g)" clip-path="url(#${id}p)"/>
      <path d="${area}" fill="url(#${id}g)" opacity=".45" clip-path="url(#${id}f)"/>
      <path d="${line}" fill="none" stroke="${metric.color}" stroke-width="3" stroke-linecap="round" clip-path="url(#${id}p)"/>
      <path d="${line}" fill="none" stroke="${metric.color}" stroke-width="3" stroke-linecap="round" stroke-dasharray="2 7" opacity=".85" clip-path="url(#${id}f)"/>
      ${nowLine}${dots}${labs}
    </svg>
    <p class="st-note">${xNow>L&&xNow<W-R?'Trait plein : déjà passé · pointillés : à venir (planifié).':''}</p></div>`;
}

// Périodes de congé : jours consécutifs regroupés (un week-end ou un férié
// entre deux jours de congé ne coupe pas la période).
function stLeavePeriods(emp,monthSel){
  const year=String(stYear());
  const days=Object.keys(emp.schedule||{}).filter(iso=>iso.startsWith(year+'-') && ST_VAC.includes(safeCode(emp.schedule[iso]))).sort();
  const inSel=iso=>monthSel==='all'||Number(iso.slice(5,7))-1===monthSel;
  const off=iso=>{ const d=parseISODateOnly(iso); return d&&(d.getDay()===0||d.getDay()===6||!!swissHolidayName(iso)||safeCode(emp.schedule[iso])==='RF'); };
  const out=[];
  let cur=null;
  for(const iso of days){
    if(cur){
      let ok=true;
      const d=parseISODateOnly(cur.end); d.setDate(d.getDate()+1);
      while(isoDate(d)<iso){ if(!off(isoDate(d))){ ok=false; break; } d.setDate(d.getDate()+1); }
      if(ok){ cur.end=iso; cur.list.push(iso); continue; }
      out.push(cur);
    }
    cur={start:iso,end:iso,list:[iso]};
  }
  if(cur)out.push(cur);
  return out.map(x=>({start:x.start,end:x.end,n:x.list.length,list:x.list})).filter(x=>x.list.some(inSel)).map(x=>({...x,n:x.list.filter(inSel).length}));
}
const ST_MON=['janv.','févr.','mars','avr.','mai','juin','juil.','août','sept.','oct.','nov.','déc.'];
function stShortDate(iso){ const d=parseISODateOnly(iso); return `${d.getDate()} ${ST_MON[d.getMonth()]}`; }
function stRange(a,b){ return a===b?stShortDate(a):`${stShortDate(a)} → ${stShortDate(b)}`; }


// Grille « Mon année » : une colonne par semaine, 7 lignes (lun→dim), colorée par code.
function ygCat(code){
  if(!code)return 'o';
  if(code==='RM')return 'm'; if(code==='RA')return 'a'; if(code==='RN')return 'n';
  if(ST_VAC.includes(code))return 'l'; if(code==='RF')return 'f'; if(code==='R')return 'r';
  return 'x';
}
function stYearGridHTML(emp){
  const y=stYear(),today=todayISO();
  let head='<span class="yg-corner"></span>';
  for(let d=1;d<=31;d++)head+=`<span class="yg-dn">${(d===1||d%5===0)?d:''}</span>`;
  let rows='';const used=new Set();let usedHol=false;
  for(let m=0;m<12;m++){
    const n=new Date(y,m+1,0).getDate();
    let cells='';
    for(let d=1;d<=31;d++){
      if(d>n){cells+='<i class="yg-c yg-x"></i>';continue;}
      const dt=new Date(y,m,d),iso=isoDate(dt),code=safeCode((emp.schedule||{})[iso]);
      const we=dt.getDay()===0||dt.getDay()===6;
      used.add(ygCat(code));if(swissHolidayName(iso))usedHol=true;
      cells+=`<button type="button" class="yg-c yg-k-${ygCat(code)} ${we?'is-we':''} ${iso===today?'is-today':''} ${swissHolidayName(iso)?'is-hol':''}" data-iso="${iso}" aria-label="${escapeHTML(formatDateWithYear(dt))} ${escapeHTML(code||'repos')}"></button>`;
    }
    rows+=`<span class="yg-mn">${ST_MON_SHORT[m]}</span>${cells}`;
  }
  const LEG=[['m','Matin'],['a','Après-midi'],['n','Nuit'],['l','Congé'],['f','Repos férié'],['r','Réserve'],['x','Autre absence'],['o','Repos']];
  const leg=LEG.filter(([k])=>used.has(k)).map(([k,l])=>`<span><i class="yg-c yg-k-${k}"></i>${l}</span>`).join('')+(usedHol?'<span><i class="yg-dot"></i>Jour férié</span>':'');
  return `<div class="st-block yg-block"><h3>Mon année ${y}</h3>
    <div class="yg-grid">${head}${rows}</div>
    <div class="yg-info" id="ygInfo">Touchez une case pour voir le détail du jour.</div>
    <div class="yg-legend">${leg}</div></div>`;
}

function renderStatsView(){
  if(!$('statsViewSection'))return;
  const ys=presenceYears();
  $('stYearWrap').classList.toggle('hidden',ys.length<2);
  $('stYear').innerHTML=ys.map(y=>`<option value="${y}">${y}</option>`).join('');
  $('stYear').value=String(stYear());
  const sel=$('stMonth');
  const opts=[['all',`Année ${stYear()}`],...MONTH_NAMES.map((n,i)=>[String(i),n])];
  sel.innerHTML=opts.map(([v,l])=>`<option value="${v}">${l}</option>`).join('');
  sel.value=String(st.month);

  const emp=selectedEmployee();
  $('stName').textContent=emp?cleanEmployeeName(emp.name):'Choisir une personne';
  $('stAvatar').textContent=emp?(initials(cleanEmployeeName(emp.name))||'—'):'—';
  if(!$('stPanel').classList.contains('hidden'))renderStPanelList();

  const body=$('stBody');
  if(!emp){
    body.innerHTML='<div class="pv-empty-state">Choisissez une personne pour voir ses statistiques.</div>';
    return;
  }

  const months=stCompute(emp);
  const yearAll='all';
  const sub=st.month; // période affichée
  const vac=stSum(months,ST_VAC,sub), rec=stSum(months,ST_REC,sub), oth=stSum(months,ST_OTHER,sub);
  const rotYear=stSum(months,ROTATION_CODES,yearAll), piqYear=stSum(months,['PIQ'],yearAll);

  const today=todayISO();
  const periods=stLeavePeriods(emp,sub);
  const vacYear=stSum(months,ST_VAC,'all');
  const vacBreak=[['VA','VJ','PI'],['HSR','HOR']].map(g=>stSum(months,g,sub));
  const takenYear=Object.keys(emp.schedule||{}).filter(iso=>iso.startsWith(String(stYear())+'-') && ST_VAC.includes(safeCode(emp.schedule[iso])) && (sub==='all'||Number(iso.slice(5,7))-1===sub));
  const past=takenYear.filter(i=>i<=today).length, next=takenYear.length-past;
  const upcoming=stLeavePeriods(emp,'all').find(x=>x.end>=today);
  let nextTxt='Aucun congé à venir', nextSub='';
  if(upcoming){
    const first=upcoming.list.find(i=>i>=today)||upcoming.start;
    const diff=Math.round((parseISODateOnly(first)-parseISODateOnly(today))/86400000);
    nextTxt=diff===0?"Aujourd'hui":stShortDate(first);
    nextSub=`${diff===0?'En cours':diff===1?'Demain':`Dans ${diff} jours`} · ${stRange(upcoming.start,upcoming.end)} (${stDays(upcoming.list.length)})`;
  }

  let h=`<div class="st-kpis">
    <div class="st-kpi" style="--st-c:${ST_COLORS.vac}"><small>Congés</small><b>${vac}<i>j</i></b><span>${vacBreak[0]} vacances · ${vacBreak[1]} récupération${vacBreak[1]>1?'s':''}${(past&&next)?`<i class="st-sub">${past} pris · ${next} à venir</i>`:''}</span></div>
    <div class="st-kpi" style="--st-c:${ST_COLORS.oth}"><small>Absences</small><b>${oth}<i>j</i></b><span>Maladie, formation, militaire…</span></div>
    <div class="st-kpi" style="--st-c:var(--accent)"><small>Prochain congé</small><b style="font-size:${upcoming?'24px':'18px'};line-height:1.2">${nextTxt}</b><span>${nextSub||'—'}</span></div>
  </div>`;

  h+=stYearGridHTML(emp);
  const blocks=[];
  const affYear=stSum(months,ST_AFF,'all');
  if(affYear>0){
    const rows=[['CA','Camion','#D9A400'],['PO','Portes','#4F8FC0'],['EN','Entretien','#5E9E3C'],['GR','Groupe','#23647E']]
      .filter(([c])=>stSum(months,[c],'all')>0)
      .map(([c,l,col])=>({label:l,n:stSum(months,[c],sub),color:col}));
    const tot=rows.reduce((a,r)=>a+r.n,0);
    blocks.push(`<div class="st-block"><h3>Affectations</h3><div class="st-big">${tot}<small>${tot>1?'jours d\'affectation':'jour d\'affectation'}</small></div>${stBarRows(rows,tot)}</div>`);
  }
  if(rotYear>0){
    const rows=[['RM','Matin','#8F8AE8'],['RA','Après-midi','#5AA9E6'],['RN','Nuit','#E8A33D']].map(([c,l,col])=>({label:l,n:stSum(months,[c],sub),color:col}));
    const tot=rows.reduce((a,r)=>a+r.n,0);
    blocks.push(`<div class="st-block"><h3>Rotations</h3><div class="st-big">${tot}<small>${tot>1?'rotations':'rotation'}</small></div>${stBarRows(rows,tot)}</div>`);
  }
  if(piqYear>0){
    const n=stSum(months,['PIQ'],sub), off=stPiquetOffDays(emp,sub);
    blocks.push(`<div class="st-block"><h3>Piquet</h3><div class="st-big">${n}<small>${n>1?'jours de piquet':'jour de piquet'}</small></div>
      ${stBarRows([{label:'Semaine',n:n-off,color:'#9A7BC8'},{label:'Week-end / férié',n:off,color:'#63428E'}],n)}</div>`);
  }
  const used=[...ST_VAC,...ST_OTHER].map(c=>({c,n:stSum(months,[c],sub)})).filter(x=>x.n>0);
  blocks.push(`<div class="st-block"><h3>Périodes de congé</h3>${periods.length
    ?`<div class="st-periods">${periods.map(x=>`<div class="st-pline ${(x.end>=today&&x.start<=today)||(x.start>today&&x===periods.find(y=>y.end>=today))?'is-next':''}"><span>${stRange(x.start,x.end)}</span><span class="st-tag ${x.end<today?'is-past':''}">${x.end<today?'Pris':x.start<=today?'En cours':'À venir'}</span><b>${stDays(x.n)}</b></div>`).join('')}</div>`
    :`<div class="st-empty-note">Aucun congé sur ${stPeriodLabel()}.</div>`}</div>`);
  blocks.push(`<div class="st-block"><h3>Détail des congés et absences</h3>${used.length
    ?`<div class="st-list">${used.map(x=>`<div class="st-line"><span class="st-code ${classFor(x.c)}">${x.c}</span><span>${escapeHTML((CODE_INFO[x.c]||[x.c])[0])}</span><b>${stDays(x.n)}</b></div>`).join('')}</div>`
    :`<div class="st-empty-note">Aucun congé ni absence sur ${stPeriodLabel()}.</div>`}</div>`);
  const detailBlocks=blocks.splice(blocks.length-2,2);
  if(blocks.length)h+=`<div class="st-blocks">${blocks.join('')}</div>`;

  const metrics=[
    {id:'vac',label:'Congés',codes:ST_VAC,color:ST_COLORS.vac},
    {id:'oth',label:'Absences',codes:ST_OTHER,color:ST_COLORS.oth}
  ];
  if(rotYear>0)metrics.push({id:'rot',label:'Rotations',codes:ROTATION_CODES,color:ST_COLORS.rot});
  if(piqYear>0)metrics.push({id:'piq',label:'Piquet',codes:['PIQ'],color:ST_COLORS.piq});
  if(affYear>0)metrics.push({id:'aff',label:'Affectations',codes:ST_AFF,color:ST_COLORS.aff});
  h+=stChartHTML(months,metrics);
  h+=`<details class="st-more"${st.more?' open':''}><summary>Voir le détail des congés et absences</summary><div class="st-blocks">${detailBlocks.join('')}</div></details>`;
  body.innerHTML=h;
  // La liste des périodes s'ouvre sur le prochain congé.
  const pl=body.querySelector('.st-periods'),nx=pl&&pl.querySelector('.is-next');
  if(pl&&nx)pl.scrollTop=Math.max(0,nx.offsetTop-pl.offsetTop-8);
}

function renderStPanelList(){
  const q=normalizeSearchText(st.search);
  const cur=state.employeeId;
  let h='';
  for(const g of presenceGroups(true)){
    const people=g.people.filter(e=>!q||normalizeSearchText(e.name).includes(q));
    if(!people.length)continue;
    h+=`<div class="pv-panel-group">${escapeHTML(g.name)}</div>`;
    h+=people.map(e=>`<button type="button" class="pv-panel-person st-pick ${e.id===cur?'is-current':''}" data-st-person="${escapeHTML(e.id)}">${escapeHTML(cleanEmployeeName(e.name))}</button>`).join('');
  }
  $('stPanelList').innerHTML=h||'<div class="pv-panel-empty">Aucune personne trouvée.</div>';
}
function openStPanel(){
  $('stPanel').classList.remove('hidden');
  $('stBackdrop').classList.toggle('hidden',!isPhoneLayout());
  lockPageScroll();
  $('stPersonBtn').setAttribute('aria-expanded','true');
  renderStPanelList();
  if(window.matchMedia('(hover:hover)').matches)$('stSearch').focus();
}
function closeStPanel(){
  if(!$('stPanel'))return;
  const wasOpen=!$('stPanel').classList.contains('hidden');
  $('stPanel').classList.add('hidden');
  $('stBackdrop').classList.add('hidden');
  if(wasOpen)unlockPageScroll();
  $('stPersonBtn').setAttribute('aria-expanded','false');
}
$('stPersonBtn').addEventListener('click',event=>{
  event.stopPropagation();
  if($('stPanel').classList.contains('hidden'))openStPanel(); else closeStPanel();
});
$('stBackdrop').addEventListener('click',event=>{ event.stopPropagation(); closeStPanel(); });
$('stSearch').addEventListener('input',event=>{ st.search=event.target.value; renderStPanelList(); });
$('stPanel').addEventListener('click',event=>{
  event.stopPropagation();
  const b=event.target.closest('[data-st-person]');
  if(!b)return;
  const emp=planning.employees.find(e=>e.id===b.dataset.stPerson);
  if(!emp)return;
  state.employeeId=emp.id;
  persistSelectedEmployee(emp);
  if(typeof updateEmployeePicker==='function')try{ updateEmployeePicker(); }catch(e){ /* ignoré */ }
  closeStPanel();
  renderStatsView();
});
$('stYear').addEventListener('change',event=>{
  st.year=Number(event.target.value); renderStatsView();
});
$('stMonth').addEventListener('change',event=>{
  st.month=event.target.value==='all'?'all':Number(event.target.value);
  saveSt(); renderStatsView();
});
$('statsViewSection').addEventListener('toggle',event=>{ if(event.target.classList&&event.target.classList.contains('st-more'))st.more=event.target.open; },true);
$('statsViewSection').addEventListener('click',event=>{
  const m=event.target.closest('[data-st-metric]');
  if(m){ st.metric=m.dataset.stMetric; renderStatsView(); return; }
  const md=event.target.closest('[data-st-mode]');
  if(md){ st.mode=md.dataset.stMode; renderStatsView(); }
});
document.addEventListener('click',event=>{
  if(!$('stPerson').contains(event.target))closeStPanel();
});
document.addEventListener('keydown',event=>{ if(event.key==='Escape')closeStPanel(); });


// Bascule téléphone <-> ordinateur : on redessine la comparaison.
window.matchMedia('(max-width:700px)').addEventListener('change',()=>{ if(state.compareMode)renderCompareView(); });

function renderTravauxSection(){
  const el=$('travauxSection');
  const titleEl=$('travauxSectionTitle');
  if(!el || !titleEl)return;

  const t=travauxPeriod();
  // La section ne reste affichée que tant que la période n'est pas terminée --
  // volontairement indépendante de la fenêtre "30 prochains jours".
  if(!t || t.end<todayISO()){
    el.classList.add('hidden');
    return;
  }

  const startD=parseISODateOnly(t.start);
  const endD=parseISODateOnly(t.end);
  titleEl.textContent=`Travaux du ${formatDateWithYear(startD)} au ${formatDateWithYear(endD)}`;
  el.classList.remove('hidden');
}


// ==========================================================
// Bandeau : l'illustration de l'usine se cadre sur le bâtiment du métier de
// la personne choisie (zoom animé), et revient à la vue d'ensemble quand la
// page ne concerne personne en particulier (Comparer, Rotations, aucun nom).
// Coordonnées en unités de l'illustration (2079 x 513, sol à y = 497).
// ==========================================================
const HERO_ART_W=2079,HERO_ART_GROUND=497;
const HERO_FOCUS={
  full:  {x:[0,2079],  top:255},
  left:  {x:[0,572],   top:268},   // rotonde : chauffeurs-grutiers
  admin: {x:[610,1140],top:322},   // bâtiment rouge et bleu : cadres, administration, intendance
  right2:{x:[940,2079],top:255},   // les 2 bâtiments de droite : équipes
  right3:{x:[612,2079],top:255},   // les 3 bâtiments de droite : électriciens, automaticiens
  right1:{x:[1265,2079],top:262}   // le grand hall de droite : mécaniciens
};
function heroFocusKeyFor(employee){
  const k=categoryKey(employee&&employee.category);
  if(/chauffeur|grutier/.test(k))return 'left';
  if(/^cadres?$|^admin/.test(k))return 'admin';
  if(/intendance/.test(k))return 'admin';
  if(/^equipe/.test(k))return 'right2';
  if(/mecani/.test(k))return 'right1';
  if(/elect|autom/.test(k))return 'right3';
  return 'full';
}
function heroFocusKey(){
  const page=currentPage();
  if(page!=='individual' && page!=='stats')return 'full';
  const e=selectedEmployee();
  return e?heroFocusKeyFor(e):'full';
}
function heroArtTransform(key){
  const hero=document.querySelector('.hero');
  const W=hero.clientWidth,H=hero.clientHeight;
  const f=HERO_FOCUS[key]||HERO_FOCUS.full;
  const [x0,x1]=f.x;
  const full=key==='full';
  const phone=W<=700;
  const rl=full||phone?0:W*0.3;                  // on centre le bâtiment plutôt vers la droite
  let s=W/(x1-x0)*(full?1:0.96);
  const maxH=H/((HERO_ART_GROUND-f.top)*(full?1:0.85)); // on peut rogner un peu le haut du bâtiment
  s=Math.min(s,maxH,2.6);
  const cx=(x0+x1)/2;
  let tx=(rl+W)/2-cx*s;
  const artW=HERO_ART_W*s;
  if(artW>=W)tx=Math.min(0,Math.max(W-artW,tx));  // l'image couvre toujours le bandeau
  else tx=W-artW;                                  // vue d'ensemble : collée à droite
  const ty=H-HERO_ART_GROUND*s+2;
  return `translate(${tx.toFixed(1)}px,${ty.toFixed(1)}px) scale(${s.toFixed(4)})`;
}
let heroFocusCurrent=null,heroFocusReady=false;
function updateHeroFocus(force=false){
  const stage=document.getElementById('heroArtStage');
  if(!stage)return;
  const logo=document.getElementById('heroArtLogo'),src=document.querySelector('.hero .logo img');
  if(logo&&src&&!logo.getAttribute('href'))logo.setAttribute('href',src.src);
  const key=heroFocusKey();
  if(!force&&key===heroFocusCurrent)return;
  if(!heroFocusReady){
    // Première apparition : vue d'ensemble, puis zoom sur le métier.
    stage.classList.remove('is-anim');
    stage.style.transform=heroArtTransform('full');
    heroFocusCurrent='full';heroFocusReady=true;
    if(key!=='full'){
      setTimeout(()=>{ stage.classList.add('is-anim'); heroFocusCurrent=key; stage.style.transform=heroArtTransform(key); },450);
    }
    return;
  }
  stage.classList.toggle('is-anim',!force);
  heroFocusCurrent=key;
  stage.style.transform=heroArtTransform(key);
}
window.addEventListener('resize',()=>{ if(heroFocusReady)updateHeroFocus(true); });

function render(){
  updateHeroFocus();
  try{ renderRotationWidget(); }catch(e){ /* ignoré */ }
  if(state.statsMode){ renderStatsView(); return; }
  if(state.compareMode){
    document.body.classList.toggle('cmp-week',cmp.scale==='week');
    renderCompareView();
    return;
  }
  if(state.teamMode){
    renderTeamView();
    return;
  }

  document.body.classList.remove('view-week','view-month','view-year','view-team');
  document.body.classList.add(`view-${state.view}`);

  document.querySelectorAll('.tab').forEach(
    b=>b.classList.toggle('active',b.dataset.view===state.view)
  );

  renderTravauxSection();

  const e=selectedEmployee();

  // Première visite (rien n'est encore mémorisé) : on n'affiche aucun
  // planning tant que la personne n'a pas explicitement choisi son nom.
  if($('noEmployeeState'))$('noEmployeeState').classList.toggle('hidden',!!e);
  if($('employeeContent'))$('employeeContent').classList.toggle('hidden',!e);
  if($('upcomingEventsCard') && !e)$('upcomingEventsCard').classList.add('hidden');

  if(!e){
    renderRotationWidget();
    updateTodayButtonState();
    updateNavBoundaryState();
    return;
  }

  if($('employeeName')) $('employeeName').textContent=cleanEmployeeName(e.name);

  if(state.view!=='year')window.__yearScrollKey=null;
  if(state.view==='week'){
    renderWeek(e);
  }else if(state.view==='month'){
    renderMonth(e);
  }else{
    renderYear(e);
  }

  renderLegend(e);
  renderUpcomingEvents(e);
  renderRotationWidget();
  updateTodayButtonState();
  updateNavBoundaryState();
  attachDayClickHandlers();
}


function commentFor(employee,dateISO){
  if(!employee || !employee.comments)return '';
  const raw=String(employee.comments[dateISO]||'').trim();
  return stripEmployeeNamePrefix(raw,employee);
}


function compactShiftLabel(code){
  const c=safeCode(code);
  if(c==='RM')return 'Matin';
  if(c==='RA')return 'Après-midi';
  if(c==='RN')return 'Nuit';
  if(CODE_INFO[c])return CODE_INFO[c][0];
  return '';
}

function normalizeSearchText(value){
  return cleanEmployeeName(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .replace(/[^a-zA-Z0-9]/g,'')
    .toLocaleLowerCase('fr');
}

function fuzzyEmployeeScore(name,query){
  const hay=normalizeSearchText(name);
  const needle=normalizeSearchText(query);
  if(!needle)return 0;
  if(hay===needle)return 1000;
  if(hay.startsWith(needle))return 900-needle.length;
  const direct=hay.indexOf(needle);
  if(direct>=0)return 700-direct;

  let pos=0;
  for(const ch of needle){
    pos=hay.indexOf(ch,pos);
    if(pos<0)return -1;
    pos++;
  }
  return 400- Math.max(0,hay.length-needle.length);
}

function normalizeIdentityPart(value){
  return String(value||'')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .toLocaleLowerCase('fr')
    .replace(/[^a-z0-9]+/g,' ')
    .trim();
}

function identityTokens(value){
  return normalizeIdentityPart(value)
    .split(/\s+/)
    .filter(token=>token && token.length>=2);
}

function employeeEmailMatchScore(employeeName,email){
  const localPart=String(email||'').split('@')[0]||'';
  const mailTokens=identityTokens(localPart);
  const nameTokens=identityTokens(cleanEmployeeName(employeeName));

  if(!mailTokens.length || !nameTokens.length)return -1;

  let score=0;
  let exactMatches=0;

  for(const nameToken of nameTokens){
    let best=0;

    for(const mailToken of mailTokens){
      if(nameToken===mailToken){
        best=Math.max(best,1000 + nameToken.length*12);
      }else if(mailToken.startsWith(nameToken) || nameToken.startsWith(mailToken)){
        const minLen=Math.min(nameToken.length,mailToken.length);
        if(minLen>=4)best=Math.max(best,620 + minLen*8);
      }else if(mailToken.includes(nameToken) || nameToken.includes(mailToken)){
        const minLen=Math.min(nameToken.length,mailToken.length);
        if(minLen>=4)best=Math.max(best,430 + minLen*6);
      }
    }

    if(best>=1000)exactMatches++;
    score+=best;
  }

  // Plusieurs morceaux du nom retrouvés dans le mail rendent la correspondance très fiable.
  score+=exactMatches*250;

  // Bonus si le nom complet, sans séparateurs, se retrouve dans la partie locale du mail.
  const compactMail=normalizeIdentityPart(localPart).replace(/\s+/g,'');
  const compactName=normalizeIdentityPart(cleanEmployeeName(employeeName)).replace(/\s+/g,'');
  if(compactName && compactMail.includes(compactName))score+=900;

  return score>0 ? score : -1;
}

function bestEmployeeForEmail(email){
  if(!email || !planning.employees.length)return null;

  const ranked=planning.employees
    .map(employee=>({
      employee,
      score:employeeEmailMatchScore(employee.name,email)
    }))
    .filter(item=>item.score>=0)
    .sort((a,b)=>
      b.score-a.score ||
      b.employee.name.length-a.employee.name.length ||
      a.employee.name.localeCompare(b.employee.name,'fr',{sensitivity:'base'})
    );

  if(!ranked.length)return null;

  const best=ranked[0];
  const second=ranked[1];

  // Un nom complet ou un token exact donne généralement > 1'000 points.
  // En dessous, on préfère ne sélectionner personne plutôt que le mauvais collègue.
  const MIN_CONFIDENCE=950;
  const MIN_GAP=180;

  if(best.score<MIN_CONFIDENCE)return null;
  if(second && second.score>=MIN_CONFIDENCE && (best.score-second.score)<MIN_GAP)return null;

  return best.employee;
}

function selectEmployeeFromAuthenticatedEmail(email){
  const normalized=String(email||'').trim().toLowerCase();
  if(!normalized)return false;

  localStorage.setItem(AUTHENTICATED_EMAIL_KEY,normalized);

  const employee=bestEmployeeForEmail(normalized);
  if(!employee)return false;

  state.employeeId=employee.id;
  persistSelectedEmployee(employee);

  if($('employeePickerName')){
    updateEmployeePicker();
  }
  if($('planningUI') && !$('planningUI').classList.contains('hidden')){
    render();
  }
  return true;
}

window.selectEmployeeFromAuthenticatedEmail=selectEmployeeFromAuthenticatedEmail;

function codeHTML(code,dateISO,currentId){
  const c=safeCode(code);
  const info=CODE_INFO[c]||[c||'Repos',c?'Code du planning':'Aucun service'];
  const displayLabel=compactShiftLabel(c);
  const employee=planning.employees.find(e=>e.id===currentId);
  const feComment=c==='FE'?commentFor(employee,dateISO):'';
  const secondary=feComment||info[1];
  const mates=colleaguesFor(dateISO,c,currentId);
  const holidayName=swissHolidayName(dateISO);
  const schoolName=schoolHolidayName(dateISO);
  const inTravaux=isInTravaux(dateISO);

  const teamHTML=ROTATION_CODES.includes(c)
    ? `<div class="shift-team compact-team ${mates.length?'has-mates':'no-mates'}">
         ${mates.length
           ? `<span class="shift-team-inline"><strong>Avec :</strong> ${mates.map(name=>escapeHTML(cleanEmployeeName(name))).join(' · ')}</span>`
           : `<span class="shift-team-inline placeholder">&nbsp;</span>`}
       </div>`
    : `<div class="shift-team compact-team no-mates"><span class="shift-team-inline placeholder">&nbsp;</span></div>`;

  const badgeStackHTML=(holidayName||schoolName||inTravaux)
    ? `<div class="week-shift-badge-stack">
        ${holidayName?`<div class="week-shift-holiday-badge" title="${escapeHTML(holidayName)}">${escapeHTML(holidayName)}</div>`:''}
        ${schoolName?`<div class="week-shift-school-badge" title="${escapeHTML(schoolName)}">${escapeHTML(schoolName)}</div>`:''}
        ${inTravaux?`<div class="week-shift-travaux-badge" title="Période de travaux">⚒️ Travaux</div>`:''}
      </div>`
    : '';

  return `<div class="week-shift ${classFor(c)}">
    ${badgeStackHTML}
    <div class="shift-mainline"><b>${escapeHTML(c||'—')}</b>${displayLabel?`<span>${escapeHTML(displayLabel)}</span>`:''}</div>
    <div class="shift-time ${feComment?'fe-comment':''}">${escapeHTML(secondary)}</div>
    ${teamHTML}
  </div>`;
}


function currentShiftPointer(now=new Date()){
  const hour=now.getHours()+now.getMinutes()/60;
  const date=new Date(now);
  date.setHours(0,0,0,0);

  if(hour>=5 && hour<13)return {code:'RM',date};
  if(hour>=13 && hour<21)return {code:'RA',date};
  if(hour>=21)return {code:'RN',date};

  // Avant 05:00, la rotation de nuit en cours a commencé la veille.
  date.setDate(date.getDate()-1);
  return {code:'RN',date};
}

function nextShiftPointer(pointer){
  const idx=ROTATION_CODES.indexOf(pointer.code);
  const nextDate=new Date(pointer.date);
  if(pointer.code==='RN')nextDate.setDate(nextDate.getDate()+1);
  return {code:ROTATION_CODES[(idx+1)%ROTATION_CODES.length],date:nextDate};
}

function rotationBlockData(pointer){
  const dISO=isoDate(pointer.date);
  const names=teamFor(dISO,pointer.code)
    .map(emp=>cleanEmployeeName(emp.name))
    .sort((a,b)=>a.localeCompare(b,'fr',{sensitivity:'base'}));

  return {
    code:pointer.code,
    label:compactShiftLabel(pointer.code),
    time:CODE_INFO[pointer.code][1],
    names
  };
}

function renderRotationWidget(){
  const card=$('rotationCard');
  if(!card)return;
  // Visible seulement sur Individuel (jamais ailleurs, même après rafraîchissement).
  let pg=typeof currentPage==='function'?currentPage():'individual';
  // Avant la restauration de la page au chargement : on se fie à la page mémorisée,
  // pour ne pas faire clignoter le bandeau après un rafraîchissement.
  if(typeof pageRestored!=='undefined' && !pageRestored){
    try{ const sv=localStorage.getItem(PAGE_KEY); if(sv&&sv!=='individual')pg=sv; }catch(e){ /* ignoré */ }
    if(/[#&]p=stats/.test(location.hash))pg='stats';
  }
  // Visible sur Mon planning et Équipe ; masqué sur les statistiques.
  if(!planning.employees.length){ card.classList.add('hidden'); return; }

  const now=currentShiftPointer();
  const next=nextShiftPointer(now);
  const nowData=rotationBlockData(now);
  const nextData=rotationBlockData(next);

  // Le dégradé de fond reprend la couleur du code actuellement en poste
  // (RM/RA/RN), au lieu d'une teinte fixe.
  const TINT_VAR_BY_CODE={RM:'--c-rm-bg',RA:'--c-ra-bg',RN:'--c-rn-bg'};
  const tintVar=TINT_VAR_BY_CODE[nowData.code];
  const tintOf=code=>TINT_VAR_BY_CODE[code]?`var(${TINT_VAR_BY_CODE[code]})`:'var(--signal-tint)';
  card.style.setProperty('--rotation-tint',tintOf(nowData.code));
  card.style.setProperty('--rotation-tint-next',tintOf(nextData.code));
  // Même teintes sur l'en-tête : son bas se fond dans le bandeau (un seul dégradé continu).
  const heroEl=document.querySelector('.hero');
  if(heroEl){ heroEl.style.setProperty('--rotation-tint',tintOf(nowData.code)); heroEl.style.setProperty('--rotation-tint-next',tintOf(nextData.code)); }

  const namesHTML=names=>names.length
    ? escapeHTML(names.join(', '))
    : '<span class="rotation-empty">aucune personne assignée</span>';

  const item=(data,tag,extraClass)=>`
    <span class="rotation-item ${extraClass}">
      <span class="rotation-tag"><span class="rotation-dot"></span>${tag}</span>
      <span class="rotation-chip ${classFor(data.code)}">${escapeHTML(data.code)}</span>
      <span class="rotation-names">${namesHTML(data.names)}</span>
    </span>`;

  card.innerHTML=
    item(nowData,'En poste','') +
    '<span class="rotation-sep" aria-hidden="true"></span>' +
    item(nextData,'Ensuite','rotation-item-next') + travauxLineHTML();

  card.classList.remove('hidden');
}

function mondaysInRange(min,max){
  const list=[];
  let cursor=mondayOf(min);
  const limit=mondayOf(max);
  while(cursor<=limit){
    list.push(new Date(cursor));
    cursor=new Date(cursor);
    cursor.setDate(cursor.getDate()+7);
  }
  return list;
}

function monthsInRange(min,max){
  const list=[];
  let y=min.getFullYear(), m=min.getMonth();
  const endKey=monthKey(max);
  while(y*12+m<=endKey){
    list.push(new Date(y,m,1));
    m++;
    if(m>11){m=0;y++;}
  }
  return list;
}

function formatWeekOptionLabel(monday){
  const end=new Date(monday);end.setDate(end.getDate()+6);
  const sameMonth=monday.getMonth()===end.getMonth();
  const range=sameMonth
    ? `${monday.getDate()}–${end.getDate()} ${MONTH_SHORT[end.getMonth()]} ${end.getFullYear()}`
    : `${monday.getDate()} ${MONTH_SHORT[monday.getMonth()]} – ${end.getDate()} ${MONTH_SHORT[end.getMonth()]} ${end.getFullYear()}`;
  return `Semaine ${isoWeek(monday)} · ${range}`;
}

function populateWeekJump(){
  const select=$('periodJump');
  const {min,max}=navBounds();
  const mondays=mondaysInRange(min,max);
  select.innerHTML=mondays.map(mon=>
    `<option value="${isoDate(mon)}">${formatWeekOptionLabel(mon)}</option>`
  ).join('');
  select.value=isoDate(mondayOf(state.date));
  $('periodTitle').classList.add('hidden');
  $('periodJumpWrap').classList.remove('hidden');
}

function populateMonthJump(){
  const select=$('periodJump');
  const {min,max}=navBounds();
  const months=monthsInRange(min,max);
  select.innerHTML=months.map(d=>
    `<option value="${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-01">${MONTH_NAMES[d.getMonth()]} ${d.getFullYear()}</option>`
  ).join('');
  select.value=`${state.date.getFullYear()}-${String(state.date.getMonth()+1).padStart(2,'0')}-01`;
  $('periodTitle').classList.add('hidden');
  $('periodJumpWrap').classList.remove('hidden');
}

function hidePeriodJump(){
  $('periodJumpWrap').classList.add('hidden');
  $('periodTitle').classList.remove('hidden');
}

function renderWeek(e){
  const start=mondayOf(state.date);
  $('viewOverline').textContent='';
  populateWeekJump();
  // Vacances scolaires : un seul libellé pour la semaine (téléphone) au lieu d'une pastille par jour.
  const schoolCount={};
  for(let i=0;i<7;i++){ const dd=new Date(start);dd.setDate(start.getDate()+i); const nm=schoolHolidayName(isoDate(dd)); if(nm)schoolCount[nm]=(schoolCount[nm]||0)+1; }
  const schoolTop=Object.entries(schoolCount).sort((a,b)=>b[1]-a[1])[0];
  const schoolNote=schoolTop&&schoolTop[1]>=3?schoolTop[0]:'';
  const wkIsos=Array.from({length:7},(_,i)=>{const dd=new Date(start);dd.setDate(start.getDate()+i);return isoDate(dd);});
  let h=`<div class="week-board">`;
  for(let i=0;i<7;i++){
    const d=new Date(start);d.setDate(start.getDate()+i);
    const isToday=isoDate(d)===todayISO();
    const dISO=isoDate(d);
    const isHoliday=!!swissHolidayName(dISO);
    const inTravaux=isInTravaux(dISO);
    // Les vacances scolaires sont maintenant affichées directement sur la
    // case du jour (badge en haut à droite, voir codeHTML) -- plus besoin
    // du petit point ici, pour éviter un indicateur en double.
    const dots=
      (isHoliday?'<span class="week-holiday-dot" aria-hidden="true"></span>':'')+
      (inTravaux?'<span class="week-travaux-dot" aria-hidden="true"></span>':'');
    h+=`<div class="week-row ${isToday?'is-current-day':''}"><div class="week-day ${isToday?'today':''}"><strong>${DAY_NAMES[i]}</strong><span>${d.getDate()}${dots}</span></div>${codeHTML(e.schedule[dISO],dISO,e.id)}</div>`;
  }
  h+='</div>';
  $('viewContainer').innerHTML=h;
}

function renderMonth(e){
  const y=state.date.getFullYear(),m=state.date.getMonth();
  $('viewOverline').textContent='';
  populateMonthJump();
  const first=new Date(y,m,1),start=mondayOf(first);
  const moIsos=Array.from({length:new Date(y,m+1,0).getDate()},(_,i)=>isoDate(new Date(y,m,i+1)));
  let h=periodStripHTML(moIsos)+'<div class="month-calendar"><div class="month-head">'+DAY_NAMES.map(d=>`<span>${d}</span>`).join('')+'</div><div class="month-grid">';
  for(let i=0;i<42;i++){
    const d=new Date(start);d.setDate(start.getDate()+i);
    const dISO=isoDate(d);
    const c=safeCode(e.schedule[dISO]);
    const isToday=dISO===todayISO();
    const isHoliday=!!swissHolidayName(dISO);
    const inTravaux=isInTravaux(dISO);
    const schoolName=schoolHolidayName(dISO);
    const inMonth=d.getMonth()===m;
    const visible=codeMatchesFilter(c);
    // On affiche désormais le code même pour les jours du mois précédent/suivant
    // qui débordent dans la grille (ligne "outside-month" un peu atténuée) --
    // plus besoin de changer de mois pour voir ce qu'il y a juste après/avant.
    const colorClass=visible?classFor(c):'outside';
    const showLabel=visible&&c;
    const dowM=(d.getDay()+6)%7,pd=new Date(d),nd=new Date(d);pd.setDate(pd.getDate()-1);nd.setDate(nd.getDate()+1);
    const lk=showLabel?linkAttrs(dowM>0&&codeMatchesFilter(safeCode(e.schedule[isoDate(pd)]))?safeCode(e.schedule[isoDate(pd)]):'',c,dowM<6&&codeMatchesFilter(safeCode(e.schedule[isoDate(nd)]))?safeCode(e.schedule[isoDate(nd)]):''):{cls:'',style:''};
    const outsideClass=inMonth?'':'outside-month';
    const markClasses=[isHoliday?'is-holiday':'',inTravaux?'is-travaux':'',schoolName?'is-schoolholiday '+schCls(dISO):''].filter(Boolean).join(' ');
    h+=`<div class="month-cell ${colorClass} ${outsideClass} ${isToday?'today':''} ${markClasses}" data-date="${dISO}" tabindex="0" role="button" ${schoolName?`title="${escapeHTML(schoolName)}"`:''}><span class="day-num"${markAttr(dISO)}>${d.getDate()}</span>${showLabel?`<span class="month-code-label${lk.cls}" style="${lk.style}">${escapeHTML(c)}</span>`:''}</div>`;
  }
  h+='</div></div>'+periodNotesHTML(moIsos);$('viewContainer').innerHTML=h;
}

function yearLegendHTML(e,y){
  const used=new Set();let sch=false,hol=false;
  for(let m=0;m<12;m++)for(let d=1;d<=new Date(y,m+1,0).getDate();d++){const iso=isoDate(new Date(y,m,d));used.add(ygCat(safeCode((e.schedule||{})[iso])));if(schoolHolidayName(iso))sch=true;if(swissHolidayName(iso))hol=true;}
  const LEG=[['m','Matin'],['a','Après-midi'],['n','Nuit'],['l','Congé'],['f','Repos férié'],['r','Réserve'],['x','Autre absence'],['o','Repos']];
  return '<div class="team-legend ind-legend">'+LEG.filter(([k])=>used.has(k)).map(([k,l])=>`<span class="team-legend-item"><i class="tl-circ yg-k-${k}"></i>${l}</span>`).join('')+(sch?'<span class="team-legend-item"><i class="tl-dot tl-sch2"></i>Vacances scolaires</span>':'')+(hol?'<span class="team-legend-item"><i class="tl-circ tl-hol"></i>Jour férié</span>':'')+'</div>';
}
function renderYear(e){
  const y=state.date.getFullYear();
  $('viewOverline').textContent='';
  $('periodTitle').textContent=String(y);
  hidePeriodJump();
  let h='<div class="year-calendar">';
  for(let m=0;m<12;m++){
    const first=new Date(y,m,1),start=mondayOf(first);
    h+=`<div class="mini-month"><h3>${MONTH_NAMES[m]}</h3><div class="mini-head">${DAY_NAMES.map(d=>`<span>${d[0]}</span>`).join('')}</div><div class="mini-grid">`;
    for(let i=0;i<42;i++){
      const d=new Date(start);d.setDate(start.getDate()+i);
      // Dernière ligne entièrement hors du mois : on ne l'affiche pas (gain de place).
      if(i>=28 && i%7===0 && d.getMonth()!==m)break;
      const c=safeCode(e.schedule[isoDate(d)]);
      const isToday=isoDate(d)===todayISO();
      const dISO=isoDate(d);
      const visible=codeMatchesFilter(c);
      const hasComment=d.getMonth()===m && c!=='HOR' && c!=='HSR' && !!commentFor(e,dISO);
      const isHoliday=d.getMonth()===m && !!swissHolidayName(dISO);
      const inTravaux=d.getMonth()===m && isInTravaux(dISO);
      const schoolName=d.getMonth()===m ? schoolHolidayName(dISO) : null;
      const markClasses=[isHoliday?'is-holiday':'',inTravaux?'is-travaux':'',schoolName?'is-schoolholiday':''].filter(Boolean).join(' ');
      h+=`<span class="mini-cell ${d.getMonth()!==m?'mini-empty':(visible?'ind-cell yg-k-'+ygCat(c):'mini-empty')} ${isToday?'today':''} ${hasComment?'has-comment':''} ${markClasses}" data-date="${d.getMonth()===m?dISO:''}" ${d.getMonth()===m?'tabindex="0" role="button"':''} ${schoolName?`title="${escapeHTML(schoolName)}"`:''}>${d.getMonth()===m?d.getDate():''}</span>`;
    }
    h+='</div></div>';
  }
  h+='</div>'+yearLegendHTML(e,y);$('viewContainer').innerHTML=h;
  // Téléphone : à l'ouverture de l'année en cours, on arrive directement sur le mois actuel.
  if(window.__yearScrollKey!==String(y)){
    window.__yearScrollKey=String(y);
    if(y===new Date().getFullYear() && window.innerWidth<=700){
      requestAnimationFrame(()=>{ try{
        const el=$('viewContainer').querySelectorAll('.mini-month')[new Date().getMonth()];
        if(el&&new Date().getMonth()>0)window.scrollTo({top:el.getBoundingClientRect().top+window.scrollY-72,behavior:'auto'});
      }catch(err){ /* ignoré */ } });
    }
  }
}


function codesForCurrentView(employee){
  const found=new Set();

  for(const d of periodDates()){
    const code=safeCode(employee.schedule[isoDate(d)]);
    if(code && OFFICIAL_CODES.has(code))found.add(code);
  }

  return LEGEND_ORDER.filter(code=>found.has(code));
}

function legendLabelFor(code){
  const info=CODE_INFO[code];
  if(!info)return code;
  return info[0];
}

function updateLegendVisibility(codes){
  const details=document.querySelector('.legend-details');
  const summary=$('legendSummary');

  if(!details || !summary)return;

  if(codes.length===0){
    details.classList.add('hidden');
    details.removeAttribute('open');
    return;
  }

  details.classList.remove('hidden');
  summary.textContent='ⓘ Codes du planning';
}

function renderLegend(e){
  const codes=codesForCurrentView(e);
  updateLegendVisibility(codes);

  $('legend').innerHTML=codes.map(code=>`
    <span class="legend-item ${classFor(code)}">
      <span class="legend-code">${escapeHTML(code)}</span>
      <span class="legend-label">${escapeHTML(legendLabelFor(code))}</span>
    </span>
  `).join('');
}

document.querySelectorAll('.tab').forEach(b=>b.addEventListener('click',()=>{
  state.view=b.dataset.view;
  state.date=new Date();
  localStorage.setItem('planningView',state.view);
  render();
}));
$('previousBtn').addEventListener('click',()=>shiftDate(-1));
$('nextBtn').addEventListener('click',()=>shiftDate(1));
$('todayBtn').addEventListener('click',goToday);

$('periodJump').addEventListener('change',event=>{
  const [y,m,d]=event.target.value.split('-').map(Number);
  state.date=new Date(y,m-1,d);
  render();
});

function isViewingCurrentPeriod(){
  const today=new Date();
  if(state.view==='week')return isoDate(mondayOf(state.date))===isoDate(mondayOf(today));
  if(state.view==='month')return state.date.getFullYear()===today.getFullYear() && state.date.getMonth()===today.getMonth();
  if(state.view==='year')return state.date.getFullYear()===today.getFullYear();
  return true;
}

function updateTodayButtonState(){
  const onCurrent=isViewingCurrentPeriod();
  const btn=$('todayBtn');
  if(btn)btn.classList.toggle('is-current',onCurrent);
}

function updateNavBoundaryState(){
  const {min,max}=navBounds();
  let atMin=false,atMax=false;

  if(state.view==='week'){
    const cur=mondayOf(state.date);
    atMin=cur<=mondayOf(min);
    atMax=cur>=mondayOf(max);
  }else if(state.view==='month'){
    const key=monthKey(state.date);
    atMin=key<=monthKey(min);
    atMax=key>=monthKey(max);
  }else{
    atMin=state.date.getFullYear()<=min.getFullYear();
    atMax=state.date.getFullYear()>=max.getFullYear();
  }

  $('previousBtn').disabled=atMin;
  $('nextBtn').disabled=atMax;
}

// Navigation tactile : balayer la vue pour passer à la période précédente/suivante.
(function setupSwipeNavigation(){
  const el=$('viewContainer');
  if(!el)return;
  let startX=0,startY=0,tracking=false;

  el.addEventListener('touchstart',event=>{
    if(event.touches.length!==1)return;
    startX=event.touches[0].clientX;
    startY=event.touches[0].clientY;
    tracking=true;
  },{passive:true});

  el.addEventListener('touchend',event=>{
    if(!tracking)return;
    tracking=false;
    const touch=event.changedTouches[0];
    const dx=touch.clientX-startX;
    const dy=touch.clientY-startY;
    if(Math.abs(dx)>56 && Math.abs(dx)>Math.abs(dy)*1.6){
      shiftDate(dx<0?1:-1);
    }
  },{passive:true});
})();






// Balayage tactile aussi dans Comparer (semaine / mois) et Statistiques
// (mois, ou année quand « Année » est choisi).
function addSwipe(el,onSwipe,canSwipe){
  if(!el)return;
  let sx=0,sy=0,on=false;
  el.addEventListener('touchstart',e=>{
    if(e.touches.length!==1||(canSwipe&&!canSwipe(e))){on=false;return;}
    sx=e.touches[0].clientX;sy=e.touches[0].clientY;on=true;
  },{passive:true});
  el.addEventListener('touchend',e=>{
    if(!on)return;on=false;
    const t=e.changedTouches[0],dx=t.clientX-sx,dy=t.clientY-sy;
    if(Math.abs(dx)>56&&Math.abs(dx)>Math.abs(dy)*1.6)onSwipe(dx<0?1:-1);
  },{passive:true});
}
addSwipe($('teamViewContainer'),dir=>{ if(state.teamScale==='week'&&isPhoneLayout())teamShiftDay(dir); else teamShiftYear(dir); });
addSwipe($('cmpGrid'),dir=>cmpShift(dir),e=>{
  if(cmp.scale==='year')return false;
  const wrap=e.target.closest('.cmp-wrap');
  return !(wrap&&wrap.scrollWidth>wrap.clientWidth+2);
});
addSwipe($('stBody'),dir=>{
  if(st.month==='all'){
    const ys=presenceYears(),i=ys.indexOf(stYear())+dir;
    if(i>=0&&i<ys.length){ st.year=ys[i]; renderStatsView(); }
  }else{
    const m=st.month+dir;
    if(m>=0&&m<12){ st.month=m; saveSt(); renderStatsView(); }
  }
},e=>!e.target.closest('.st-more table,.st-periods'));

$('employeePickerButton').addEventListener('click',()=>{
  if($('employeePicker').classList.contains('open'))closeEmployeePicker();
  else openEmployeePicker();
});

$('employeeBackdrop').addEventListener('click',event=>{ event.stopPropagation(); closeEmployeePicker(); });

$('employeeSearch').addEventListener('input',event=>{
  const query=event.target.value;
  populateEmployees(query);

  if(!query.trim())return;

  const ranked=planning.employees
    .map(e=>({employee:e,score:fuzzyEmployeeScore(e.name,query)}))
    .filter(x=>x.score>=0)
    .sort((a,b)=>b.score-a.score);

  if(ranked.length){
    const first=ranked[0].employee;
    if(first.id!==state.employeeId){
      state.employeeId=first.id;
      persistSelectedEmployee(first);
      $('employeePickerName').textContent=first.name;
      $('employeePickerAvatar').textContent=initials(first.name)||'—';
      render();
    }
  }
});

document.addEventListener('click',event=>{
  if(!$('employeePicker').contains(event.target))closeEmployeePicker();
});

document.addEventListener('keydown',event=>{
  if(event.key==='Escape')closeEmployeePicker();

  if(event.key==='ArrowLeft' || event.key==='ArrowRight'){
    const tag=(document.activeElement && document.activeElement.tagName) || '';
    const isTyping=['INPUT','TEXTAREA','SELECT'].includes(tag);
    const pickerOpen=$('employeePicker').classList.contains('open');
    const modalOpen=$('dayModalBackdrop').classList.contains('open');
    if(!isTyping && !pickerOpen && !modalOpen && !state.teamMode && !state.compareMode && !state.statsMode){
      shiftDate(event.key==='ArrowLeft'?-1:1);
    }
  }
});


const legendDetails=document.querySelector('.legend-details');
if(legendDetails){
  legendDetails.addEventListener('toggle',()=>{
    if($('legendSummary')){
      $('legendSummary').textContent=legendDetails.open
        ? 'Masquer la légende'
        : 'Afficher la légende';
    }
  });
}



/* Vérifie périodiquement les données sans recharger toute la page */
setInterval(()=>loadPlanning(true),60000);

/* Garde "En poste / Ensuite" à jour même si personne ne touche à l'appli */
setInterval(renderRotationWidget,60000);

loadPlanning(false);

// ==========================================================
// Ergonomie : barre d'onglets mobile, raccourcis clavier, bouton Retour,
// mémoire de la période, groupes enregistrés (Comparer), lien partageable.
// ==========================================================
try{ history.replaceState({tr:1,page:'individual'},''); }catch(e){ /* ignoré */ }
var histMode='push',histPopping=false;
function histNote(page){
  try{
    if(histPopping)return;
    if(histMode==='replace'){ history.replaceState({tr:1,page},''); histMode='push'; return; }
    const st=history.state;
    if(!st||st.page!==page)history.pushState({tr:1,page},'');
  }catch(e){ /* ignoré */ }
}
window.addEventListener('popstate',e=>{
  const st=e.state;
  if(!st||!st.tr)return;
  histPopping=true;
  try{ if(st.page!==currentPage())setPage(st.page); }finally{ histPopping=false; }
});

// --- Barre d'onglets (téléphone) : Mon planning · Équipe · Réglages ---
(function(){
  const nav=document.createElement('nav');
  nav.id='tabBar';nav.className='tabbar';nav.hidden=true;
  nav.setAttribute('aria-label','Navigation');
  const ico=p=>`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${p}</svg>`;
  nav.innerHTML=NAV_GROUPS.map(g=>`<button type="button" class="tabbar-btn" data-tab-group="${g.id}">${ico(g.icon)}<span>${g.label}</span></button>`).join('')
    +`<button type="button" class="tabbar-btn" id="tabSettings" aria-haspopup="true" aria-expanded="false">${ico('<path d="M4 7h10M18 7h2M4 17h2M10 17h10"/><circle cx="16" cy="7" r="2"/><circle cx="8" cy="17" r="2"/>')}<span>Réglages</span></button>`;
  document.body.appendChild(nav);
  nav.addEventListener('click',e=>{
    const st=e.target.closest('#tabSettings');
    if(st){
      e.stopPropagation();
      if($('navMenuPanel').classList.contains('hidden'))openNavMenu(); else closeNavMenu();
      return;
    }
    const b=e.target.closest('[data-tab-group]');
    if(!b)return;
    closeNavMenu();
    const g=NAV_GROUPS.find(x=>x.id===b.dataset.tabGroup);
    const cur=currentPage();
    setPage(g.pages.includes(cur)?cur:lastInGroup[g.id]);
  });
})();
function tabBarSync(){
  const nav=document.getElementById('tabBar');
  if(!nav)return;
  nav.hidden=!(planning&&planning.employees&&planning.employees.length);
  const gid=groupOf(currentPage()).id;
  nav.querySelectorAll('[data-tab-group]').forEach(b=>{
    const on=b.dataset.tabGroup===gid;
    b.classList.toggle('is-active',on);
    if(on)b.setAttribute('aria-current','page');else b.removeAttribute('aria-current');
  });
}

// --- Sélecteurs de page : « Planning | Statistiques » et « Par poste | Par personne » ---
(function(){
  const mk=(items,active)=>`<div class="seg" role="tablist">${items.map(([id,l])=>`<button type="button" role="tab" aria-selected="${id===active}" class="seg-btn ${id===active?'is-on':''}" data-go="${id}">${l}</button>`).join('')}</div>`;
  const PLAN=[['individual','Individuel'],['compare','Collègues'],['team','Rotations']];
  const put=(sel,html)=>{ const el=document.querySelector(sel); if(el)el.innerHTML=html; };
  put('#teamViewSection .team-view-title',mk(PLAN,'team'));
  put('#compareViewSection .team-view-title',mk(PLAN,'compare'));
  put('#statsViewSection .team-view-title','');{const t=document.querySelector('#statsViewSection .team-view-title');if(t)t.style.display='none';}
  const ctl=document.querySelector('section.controls');
  if(ctl){ const w=document.createElement('div'); w.className='seg-wrap'; w.innerHTML=mk(PLAN,'individual'); ctl.insertBefore(w,ctl.firstChild); }
  document.addEventListener('click',e=>{
    const b=e.target.closest('[data-go]');
    if(b&&b.dataset.go!==currentPage())setPage(b.dataset.go);
  });
})();


// --- Clavier virtuel : décalage des feuilles du bas ---
(function(){
  const vv=window.visualViewport; if(!vv)return;
  const upd=()=>{
    const kb=Math.max(0,Math.round(window.innerHeight-vv.height-vv.offsetTop));
    const r=document.documentElement.style;
    r.setProperty('--kb',kb+'px'); r.setProperty('--vvh',Math.round(vv.height)+'px');
  };
  vv.addEventListener('resize',upd); vv.addEventListener('scroll',upd); upd();
})();

// --- Raccourcis clavier (ordinateur) ---
document.addEventListener('keydown',e=>{
  if(e.defaultPrevented||e.ctrlKey||e.metaKey||e.altKey)return;
  const t=e.target;
  if(t&&(/^(INPUT|SELECT|TEXTAREA)$/.test(t.tagName)||t.isContentEditable))return;
  if(document.querySelector('.day-modal-backdrop.open'))return;
  if(!(planning&&planning.employees&&planning.employees.length))return;
  const k=e.key;
  if(/^[1-4]$/.test(k)&&NAV_PAGES[+k-1]){ setPage(NAV_PAGES[+k-1].id); e.preventDefault(); return; }
  const ids={individual:['previousBtn','nextBtn','todayBtn'],team:['teamPrevBtn','teamNextBtn','teamTodayBtn'],compare:['cmpPrevBtn','cmpNextBtn','cmpTodayBtn']}[currentPage()];
  if(!ids)return;
  const id=k==='ArrowLeft'?ids[0]:k==='ArrowRight'?ids[1]:(k==='t'||k==='T')?ids[2]:null;
  if(!id)return;
  const b=document.getElementById(id);
  if(b&&!b.disabled&&b.offsetParent!==null){ b.click(); e.preventDefault(); }
});

// --- Lien partageable + mémoire de la période ---
const parseYMD=s=>{ const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(s||''); return m?new Date(+m[1],+m[2]-1,+m[3]):null; };
function shareHTML(){
  return `<button type="button" class="install-btn share-btn" data-share-link><span style="font-size:16px;line-height:1">🔗</span><span>Copier le lien de cette vue<small>Page, période et personnes affichées</small></span></button><div id="shareHint" class="install-hint hidden"></div>`;
}
function shareURL(){
  const q=new URLSearchParams();
  const page=currentPage();
  q.set('p',page);
  if(page==='individual'){ q.set('v',state.view); q.set('d',isoDate(state.date)); if(state.employeeId)q.set('e',state.employeeId); }
  else if(page==='compare'){ q.set('s',cmp.scale); q.set('cd',isoDate(cmp.date)); q.set('c',[...cmp.people].join('|')); }
  else if(page==='team'){ q.set('ts',state.teamScale); q.set('td',isoDate(state.teamDate)); }
  else if(page==='stats'){ if(state.employeeId)q.set('e',state.employeeId); }
  return location.origin+location.pathname+'#'+q.toString();
}
async function shareCurrentView(){
  const url=shareURL(),h=document.getElementById('shareHint');
  const say=t=>{ if(h){ h.textContent=t; h.classList.remove('hidden'); } };
  try{
    if(navigator.share&&matchMedia('(max-width:700px)').matches){ await navigator.share({title:'Planning Tridel',url}); return; }
  }catch(e){ if(e&&e.name==='AbortError')return; }
  try{ await navigator.clipboard.writeText(url); say('Lien copié ✓'); }
  catch(e){ say(url); }
}
function readSharedHash(){
  const h=location.hash.replace(/^#/,'');
  if(!h)return null;
  const q=new URLSearchParams(h);
  return NAV_PAGES.some(x=>x.id===q.get('p'))?q:null;
}
function applySharedState(q){
  const has=k=>planning.employees.some(x=>x.id===k);
  const e=q.get('e'); if(e&&has(e)){ state.employeeId=e; try{ updateEmployeePicker(); }catch(_){ /* ignoré */ } }
  const v=q.get('v'); if(['week','month','year'].includes(v))state.view=v;
  const d=parseYMD(q.get('d')); if(d)state.date=d;
  const s=q.get('s'); if(['week','month','year'].includes(s))cmp.scale=s;
  const cd=parseYMD(q.get('cd')); if(cd)cmp.date=cd;
  if(q.has('c')){
    const ks=q.get('c').split('|').filter(Boolean);
    cmp.people=new Set(ks.filter(k=>planning.employees.some(x=>presencePersonKey(x)===k)));
  }
  const ts=q.get('ts'); if(ts==='week'||ts==='year')state.teamScale=ts;
  const td=parseYMD(q.get('td')); if(td)state.teamDate=td;
}
const PERIOD_KEY='lastPeriod';
var periodReady=false;
function periodSave(){
  if(!periodReady)return;
  try{ localStorage.setItem(PERIOD_KEY,JSON.stringify({ts:Date.now(),d:isoDate(state.date),cd:isoDate(cmp.date),td:isoDate(state.teamDate)})); }catch(e){ /* ignoré */ }
}
function periodRestore(){
  try{
    const o=JSON.parse(localStorage.getItem(PERIOD_KEY)||'null');
    if(!o||Date.now()-o.ts>20*60*1000)return;   // au-delà de 20 min : on repart d'aujourd'hui
    const d=parseYMD(o.d),cd=parseYMD(o.cd),td=parseYMD(o.td);
    if(d)state.date=d; if(cd)cmp.date=cd; if(td)state.teamDate=td;
  }catch(e){ /* ignoré */ }
}
setInterval(periodSave,2500);
window.addEventListener('pagehide',periodSave);

// --- Comparer : groupes enregistrés ---
const CMP_SAVED_KEY='cmpSavedGroups';
function cmpSavedLoad(){
  try{ const a=JSON.parse(localStorage.getItem(CMP_SAVED_KEY)||'[]'); return Array.isArray(a)?a.filter(g=>g&&typeof g.name==='string'&&Array.isArray(g.people)):[]; }
  catch(e){ return []; }
}
function cmpSavedStore(a){ try{ localStorage.setItem(CMP_SAVED_KEY,JSON.stringify(a)); }catch(e){ /* ignoré */ } }
function cmpSavedChipsHTML(){
  const a=cmpSavedLoad();
  let h=a.map((g,i)=>{
    const on=g.people.length===cmp.people.size&&g.people.every(k=>cmp.people.has(k));
    return `<span class="pv-saved"><button type="button" class="pv-chip ${on?'is-on':''}" data-cmp-saved="${i}" aria-pressed="${on}" title="Afficher ce groupe">★ ${escapeHTML(g.name)}</button><button type="button" class="pv-saved-x" data-cmp-saved-del="${i}" aria-label="Supprimer le groupe ${escapeHTML(g.name)}">×</button></span>`;
  }).join('');
  return h;
}
function cmpSaveChipHTML(){
  let h='';
  if(cmp.people.size)h+=`<button type="button" class="pv-chip pv-save" data-cmp-save title="Enregistrer la sélection actuelle comme groupe">＋ Enregistrer la sélection</button>`;
  return h;
}
$('cmpGroups').addEventListener('click',e=>{
  const ap=e.target.closest('[data-cmp-saved]'),del=e.target.closest('[data-cmp-saved-del]'),sv=e.target.closest('[data-cmp-save]');
  if(!ap&&!del&&!sv)return;
  e.stopPropagation();
  const a=cmpSavedLoad();
  if(ap){
    const g=a[+ap.dataset.cmpSaved]; if(!g)return;
    cmp.people=new Set(g.people.filter(k=>planning.employees.some(x=>presencePersonKey(x)===k)));
    saveCmp(); renderCompareView(); return;
  }
  if(del){
    const g=a[+del.dataset.cmpSavedDel]; if(!g)return;
    if(!confirm(`Supprimer le groupe « ${g.name} » ?`))return;
    a.splice(+del.dataset.cmpSavedDel,1); cmpSavedStore(a); renderCompareView(); return;
  }
  const name=(prompt('Nom du groupe (ex. Mon équipe) :')||'').trim().slice(0,24);
  if(!name)return;
  const entry={name,people:[...cmp.people]};
  const i=a.findIndex(g=>g.name.toLocaleLowerCase('fr')===name.toLocaleLowerCase('fr'));
  if(i>=0)a[i]=entry;else a.push(entry);
  cmpSavedStore(a); renderCompareView();
});

// --- Code profil : transférer personne, groupes et apparence vers un autre appareil ou
// vers l'application installée (qui a sa propre mémoire, vide au départ).
const PROFILE_KEYS=['planningEmployee','planningEmployeeName','cmpSavedGroups','compareState','themeMode','planningLastInGroup'];
function profileHTML(){
  return `<div class="profile-row"><span class="profile-lab"><span style="font-size:16px;line-height:1">👤</span><span>Code profil<small>Changer d'appareil ou d'application</small></span></span><span class="profile-btns"><button type="button" data-profile-export>Copier</button><button type="button" data-profile-import>Importer</button></span></div><div id="profileHint" class="install-hint hidden"></div>`;
}
function profileSay(t){ const h=document.getElementById('profileHint'); if(h){ h.textContent=t; h.classList.remove('hidden'); } }
function profileEncode(){
  const o={};
  for(const k of PROFILE_KEYS){ try{ const v=localStorage.getItem(k); if(v!=null)o[k]=v; }catch(e){ /* ignoré */ } }
  return 'TRIDEL:'+btoa(unescape(encodeURIComponent(JSON.stringify(o)))).replace(/=+$/,'');
}
async function profileExport(){
  const code=profileEncode();
  try{ await navigator.clipboard.writeText(code); profileSay('Code copié ✓ — colle-le dans l\'autre appareil (Importer un code profil).'); }
  catch(e){ prompt('Copie ce code :',code); }
}
function profileImport(text){
  let code=typeof text==='string'?text:prompt('Colle ici ton code profil :');
  if(!code)return;
  try{
    code=code.trim().replace(/^.*?(TRIDEL:)/s,'$1');
    if(!code.startsWith('TRIDEL:'))throw new Error('format');
    const o=JSON.parse(decodeURIComponent(escape(atob(code.slice(7)))));
    for(const k of PROFILE_KEYS){ if(typeof o[k]==='string')localStorage.setItem(k,o[k]); }
    localStorage.setItem('compareSelV','2');
    location.reload();
  }catch(e){ profileSay('Code non reconnu. Copie-le à nouveau depuis l\'autre appareil.'); if(typeof text!=='string')alert('Code non reconnu.'); }
}
document.addEventListener('click',e=>{ if(e.target.closest('.no-emp-import [data-profile-import]'))profileImport(); });

// Période de travaux : une ligne discrète dans le bandeau En poste / Ensuite.
function travauxLineHTML(){
  const t=travauxPeriod();
  if(!t||t.end<todayISO())return '';
  const a=parseISODateOnly(t.start),b=parseISODateOnly(t.end);
  if(!a||!b)return '';
  const yr=new Date().getFullYear();
  const f=(d,y)=>d.toLocaleDateString('fr-CH',y?{day:'numeric',month:'short',year:'numeric'}:{day:'numeric',month:'short'});
  const showYear=a.getFullYear()!==yr||b.getFullYear()!==yr;
  const label=`Travaux · ${f(a,showYear&&a.getFullYear()!==b.getFullYear())} → ${f(b,showYear)}`;
  return `<button type="button" class="rotation-travaux" onclick="addTravauxToCalendar()" title="Ajouter la période de travaux au calendrier" aria-label="${label} — ajouter au calendrier"><span class="rt-ico" aria-hidden="true">⚒️</span><span>${label}</span><span class="rt-cal" aria-hidden="true">📅</span></button>`;
}

document.addEventListener('click',e=>{
  const c=e.target.closest('.yg-grid .yg-c[data-iso]'); if(!c)return;
  const emp=selectedEmployee(); if(!emp)return;
  const iso=c.dataset.iso,d=parseISODateOnly(iso),code=safeCode((emp.schedule||{})[iso]);
  const info=CODE_INFO[code]||[code||'Repos',code?'Code du planning':'Aucun service'];
  const hol=swissHolidayName(iso),sch=schoolHolidayName(iso);
  const box=$('ygInfo'); if(!box)return;
  document.querySelectorAll('.yg-c.is-sel').forEach(x=>x.classList.remove('is-sel')); c.classList.add('is-sel');
  box.innerHTML=`<b>${escapeHTML(formatDateWithYear(d))}</b> · <span class="st-code ${classFor(code)}">${escapeHTML(code||'—')}</span> ${escapeHTML(info[0])}${hol?` · ${escapeHTML(hol)}`:''}${sch?` · ${escapeHTML(sch)}`:''}`;
});

// Filtres (Collègues) : repliés par défaut, mémorisés ; le bouton affiche le nombre de personnes.
(function(){
  const btn=document.getElementById('cmpGroupsToggle'),box=document.getElementById('cmpFiltersBox'),groups=document.getElementById('cmpGroups'),cnt=document.getElementById('cmpCount');
  if(!btn||!box)return;
  let open=false; try{ open=localStorage.getItem('cmpGroupsOpen')==='1'; }catch(e){}
  const apply=()=>{
    box.classList.toggle('is-collapsed',!open);
    btn.setAttribute('aria-expanded',String(open));
    btn.classList.toggle('is-open',open);
  };
  const mark=()=>{
    const n=cnt?cnt.textContent.trim():'';
    let b=btn.querySelector('.cmp-gtoggle-n'); if(!b){ b=document.createElement('span'); b.className='cmp-gtoggle-n'; btn.appendChild(b); }
    b.textContent=n&&n!=='0'?n:'';
  };
  btn.addEventListener('click',()=>{ open=!open; try{ localStorage.setItem('cmpGroupsOpen',open?'1':'0'); }catch(e){} apply(); });
  if(cnt)new MutationObserver(mark).observe(cnt,{childList:true,characterData:true,subtree:true});
  apply(); mark();
})();

// En-tête + bandeau « En poste » : classe posée en JS (plus fiable que :has() sur iOS pendant les changements de page).
(function(){
  const card=document.getElementById('rotationCard'),hero=document.querySelector('.hero');
  if(!card||!hero)return;
  const sync=()=>hero.classList.toggle('hero-joined',!card.classList.contains('hidden'));
  sync();
  new MutationObserver(sync).observe(card,{attributes:true,attributeFilter:['class']});
})();
