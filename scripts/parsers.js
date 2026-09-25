// Parsers for the CBG raw workbooks. Used by scripts/build-data.js (run by GitHub Actions).
const XLSX=require('xlsx');
const MONTHS=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const NUM_KEYS=['cases','availMin','fillMin','bbtHl','bottledHl','pmPlanned','pmDone','pmPct','elec','water','fuel','co2','utilHl','plMtd','plYtd','plTgt'];
const pad=n=>String(n).padStart(2,'0');
const dstr=(y,m,d)=>`${y}-${pad(m+1)}-${pad(d)}`;

const isDate=v=>Object.prototype.toString.call(v)==='[object Date]';
function parseDate(v){
  if(v==null||v==='')return null;
  if(isDate(v)&&!isNaN(v)){const d=new Date(v.getTime()+12*3600e3);return ok(d.getFullYear(),d.getMonth(),d.getDate())}
  if(typeof v==='number'){if(v>20000&&v<80000){const d=new Date(Math.round((v-25569)*864e5));return ok(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())}return null}
  const s=String(v).trim();let m;
  if((m=s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/)))return ok(+m[1],+m[2]-1,+m[3]);
  if((m=s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/))){let y=+m[3];if(y<100)y+=2000;let a=+m[1],b=+m[2];if(b>12&&a<=12)[a,b]=[b,a];return ok(y,b-1,a)}
  return null;
  function ok(y,mo,da){if(y<2000||y>2100||mo<0||mo>11||da<1||da>31)return null;const t=new Date(y,mo,da);if(t.getMonth()!==mo)return null;return dstr(y,mo,da)}
}
function num(v){if(v==null||v===''||isDate(v))return null;if(typeof v==='number')return isFinite(v)?v:null;
  const s=String(v).replace(/[,\s%$]/g,'');if(!/^-?\d*\.?\d+(e-?\d+)?$/i.test(s))return null;const n=parseFloat(s);return isFinite(n)?n:null}
const txt=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();
function monthFromText(t){const m=String(t).toLowerCase().match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/);return m?['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[0]):null}
function yearFromText(t){const m=String(t).match(/(20\d{2})/);if(m)return +m[1];const n=String(t).match(/_(\d{2})_/);return n?2000+(+n[1]):null}
function readSheet(ws,maxR=600,maxC=160){
  if(!ws||!ws['!ref'])return [];
  const r=XLSX.utils.decode_range(ws['!ref']);r.e.r=Math.min(r.e.r,r.s.r+maxR);r.e.c=Math.min(r.e.c,maxC);r.s.r=0;r.s.c=0;
  return XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:true,range:r});
}
function detectType(wb){
  const names=wb.SheetNames;
  if(names.some(n=>/^utility analysis/i.test(n.trim())))return 'util';
  for(const n of names.slice(0,4)){const a=readSheet(wb.Sheets[n],12,30);if(a.some(r=>r&&r.some(c=>txt(c).toUpperCase()==='DATE')&&r.some(c=>/^CASES/i.test(txt(c)))&&r.some(c=>/^FILL/i.test(txt(c)))))return 'gross'}
  for(const n of names){const a=readSheet(wb.Sheets[n],130,20);if(a.some(r=>r&&r.some(c=>/process loss \(volume\)/i.test(txt(c)))))return 'process'}
  return 'generic';
}
// Gross Efficiency workbook: one sheet per month, one row per product run
function parseGross(wb){
  const out={};
  for(const n of wb.SheetNames){
    const a=readSheet(wb.Sheets[n],300,40);
    const h=a.findIndex(r=>r&&r.some(c=>txt(c).toUpperCase()==='DATE')&&r.some(c=>/^FILL/i.test(txt(c))));
    if(h<0)continue;
    const comb=i=>`${txt(a[h][i])} ${txt(a[h+1]&&a[h+1][i])}`.trim().toUpperCase();
    const W=a[h].length,C={};
    for(let i=0;i<W;i++){const c=comb(i);
      if(C.date==null&&/^DATE/.test(c))C.date=i;
      else if(C.avail==null&&/^AVAIL/.test(c))C.avail=i;
      else if(C.fill==null&&/^FILL/.test(c))C.fill=i;
      else if(C.cases==null&&/^CASES( BOTTLED)?$/.test(c))C.cases=i;
      else if(C.bbt==null&&/^BBT/.test(c))C.bbt=i;
      else if(C.bottled==null&&/^BOTTLED HLS/.test(c))C.bottled=i;}
    // some months leave the "AVAIL." heading blank: it is the TIME column just before FILL
    if(C.avail==null&&C.fill>0&&/^TIME$/.test(comb(C.fill-1)))C.avail=C.fill-1;
    if(C.date==null||C.cases==null)continue;
    const sm=monthFromText(n),sy=yearFromText(n.replace(/_206$/,'_2026'));
    for(let r=h+2;r<a.length;r++){
      const row=a[r];if(!row)continue;let ds=parseDate(row[C.date]);if(!ds)continue;
      // keep rows for the sheet's own month; fix a mistyped year (e.g. 2005 typed for 2026)
      if(sm!=null&&+ds.slice(5,7)-1!==sm)continue;
      if(sy!=null&&+ds.slice(0,4)!==sy){const fx=parseDate(`${sy}-${ds.slice(5)}`);if(!fx)continue;ds=fx}
      const rec=out[ds]||(out[ds]={cases:0,availMin:0,fillMin:0,bbtHl:0,bottledHl:0});
      const add=(k,c)=>{if(c==null)return;const v=num(row[c]);if(v!=null)rec[k]+=v};
      add('cases',C.cases);add('availMin',C.avail);add('fillMin',C.fill);add('bbtHl',C.bbt);add('bottledHl',C.bottled);
    }
  }
  for(const ds in out){const r=out[ds];for(const k in r)r[k]=+r[k].toFixed(4);if(!Object.values(r).some(v=>v))delete out[ds]}
  return {rows:out,targets:{}};
}
// Utilities Tracking workbook: "Utility Analysis <Month> <Year>" has one row per day of the month
function parseUtil(wb,fname){
  const rows={},targets={};
  const sn=wb.SheetNames.find(n=>/^utility analysis/i.test(n.trim()));
  let mo=monthFromText(sn.replace(/utility analysis/i,'')),yr=yearFromText(sn);
  if(mo==null)mo=monthFromText(fname);if(yr==null)yr=yearFromText(fname);
  if(mo==null||yr==null)throw new Error('Could not tell which month this utilities file covers.');
  const a=readSheet(wb.Sheets[sn],80,40);
  const h=a.findIndex(r=>r&&r.some(c=>/electricity total/i.test(txt(c))));
  if(h<0)throw new Error('Could not find the utility columns.');
  const C={};a[h].forEach((c,i)=>{const t=txt(c).toUpperCase();
    if(/ELECTRICITY TOTAL/.test(t))C.elec=i;else if(t==='WATER')C.water=i;else if(t==='FUEL')C.fuel=i;else if(/^CO2/.test(t))C.co2=i;else if(/^PRODUCTION/.test(t))C.prod=i});
  for(let r=h+1;r<Math.min(h+4,a.length);r++){const row=a[r]||[];for(let i=(C.prod||0);i<row.length;i++)if(txt(row[i]).toUpperCase()==='TOTAL'){C.vol=i;break}if(C.vol!=null)break}
  let seen=false;
  for(let r=h+1;r<a.length;r++){
    const row=a[r];if(!row)continue;
    if(/^total$/i.test(txt(row[0]))||row.some(c=>txt(c).toUpperCase()==='MTD'))break;
    let d=num(row[0]);
    if(d==null&&seen)d=31;            // unlabelled rows under the day list still count in the sheet totals
    if(d==null||d<1||d>31||d%1)continue;seen=true;
    // rows past the month's last day (e.g. "29"-"31" in February) still count in the sheet's totals, so fold them into the last day
    const last=new Date(yr,mo+1,0).getDate(),ds=dstr(yr,mo,Math.min(d,last));
    const v={elec:num(row[C.elec]),water:num(row[C.water]),fuel:num(row[C.fuel]),co2:num(row[C.co2]),utilHl:C.vol!=null?num(row[C.vol]):null};
    if(!Object.values(v).some(x=>x))continue;
    const rec=rows[ds]||(rows[ds]={elec:0,water:0,fuel:0,co2:0,utilHl:0});for(const k in v)rec[k]=+(rec[k]+(v[k]||0)).toFixed(4);
  }
  // monthly targets from the Daily KPIs sheet
  const kn=wb.SheetNames.find(n=>/^daily kpis$/i.test(n.trim()));
  if(kn){const k=readSheet(wb.Sheets[kn],200,160);
    let tcol=null,mcol=null;
    k.forEach(r=>{if(!r)return;r.forEach((c,i)=>{const t=txt(c).toLowerCase();if(t==='monthly target'&&mcol==null)mcol=i})});
    for(let r=0;r<k.length;r++){const row=k[r]||[];const lab=txt(row[0]).toLowerCase();
      if(lab==='area'){row.forEach((c,i)=>{if(txt(c).toLowerCase()==='target')tcol=i})}
      if(lab==='cases produced'&&mcol!=null){const v=num(row[mcol]);if(v)targets.cases=v}
      if(tcol!=null){const v=num(row[tcol]);if(v==null)continue;
        if(/^total water consumption$/.test(lab))targets.water=v;
        else if(/^co2 \(kg\.\/hl\.\)\s*kpi/.test(lab))targets.co2=v;
        else if(/^electricity \(kwh\.\/hl\.\)/.test(lab))targets.elec=v;
        else if(/^fuel \(l\/hl\.\)/.test(lab))targets.fuel=v;}
    }}
  return {rows,targets,month:dstr(yr,mo,1).slice(0,7)};
}
// Daily Process Report workbook: one sheet per report day
function parseProcess(wb,fname){
  const rows={},targets={};
  let mo=monthFromText(fname),yr=yearFromText(fname);
  for(const n of wb.SheetNames){
    const m=n.trim().match(/^([A-Za-z]+)\.?\s*(\d{1,2})(st|nd|rd|th)?\b/);if(!m)continue;
    const sm=monthFromText(m[1]);if(sm==null)continue;
    const a=readSheet(wb.Sheets[n],140,20);
    let y=yr;if(y==null){const c4=a[3]&&a[3][2];const d=parseDate(c4);if(d)y=+d.slice(0,4)}
    if(y==null)continue;
    if(mo!=null&&sm!==mo)continue;
    const h=a.findIndex(r=>r&&r.some(c=>/process loss \(volume\)/i.test(txt(c))));if(h<0)continue;
    const hr=a[h+1]||[];let cm=null,cy=null,ct=null;
    hr.forEach((c,i)=>{const t=txt(c).toLowerCase();if(t==='mtd'&&cm==null)cm=i;else if(t==='ytd'&&cy==null)cy=i;else if(t==='target'&&ct==null)ct=i});
    if(cm==null)continue;
    for(let r=h+2;r<Math.min(h+22,a.length);r++){
      const row=a[r]||[];if(txt(row[1])!=='')continue;const v=num(row[cm]);if(v==null||v<=0||v>50)continue;
      const ds=dstr(y,sm,+m[2]);if(!parseDate(ds))break;
      rows[ds]={plMtd:+v.toFixed(4)};
      const yv=cy!=null?num(row[cy]):null;if(yv!=null&&yv>0&&yv<50)rows[ds].plYtd=+yv.toFixed(4);
      const tv=ct!=null?num(row[ct]):null;if(tv!=null&&tv>0){rows[ds].plTgt=tv;targets.loss=tv}
      break;
    }
  }
  return {rows,targets};
}
function parseKnown(wb,fname){
  const t=detectType(wb);
  if(t==='gross')return {type:'Gross Efficiency',...parseGross(wb)};
  if(t==='util')return {type:'Utilities Tracking',...parseUtil(wb,fname)};
  if(t==='process')return {type:'Daily Process Report',...parseProcess(wb,fname)};
  return null;
}

/* ---- generic files (e.g. a PM tracker): column matching ---- */
const PAT={
  date:/\bdate\b|^day$/,
  pmPct:/(pm|prevent).*(%|complian)|complian/,
  pmPlanned:/(pm|prevent).*(plan|sched|due|target)|planned\s*(pm|task|work)|scheduled/,
  pmDone:/(pm|prevent).*(done|complet|actual|execut)|complet\w*/,
  availMin:/avail/,fillMin:/fill\w*\s*time|^fill/,
  plMtd:/process loss|loss/,
  elec:/elec|kwh/,water:/water/,fuel:/fuel|diesel/,co2:/co2|co₂/,
  cases:/case/,utilHl:/\bhl\b|hecto|volume|production/,
};
const PAT_ORDER=['date','pmPct','pmPlanned','pmDone','availMin','fillMin','plMtd','elec','water','fuel','co2','cases','utilHl'];
function autoMap(headers){
  const map={},used=new Set(),norm=headers.map(h=>txt(h).toLowerCase());
  for(const k of PAT_ORDER)for(let i=0;i<norm.length;i++){const h=norm[i];if(!h||used.has(i))continue;
    if(k==='cases'&&/hour|\/|per|hold|reject/.test(h))continue;if(PAT[k].test(h)){map[k]=i;used.add(i);break}}
  return map;
}
function colLetter(i){let s='';i++;while(i>0){const m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=Math.floor((i-1)/26)}return s}
function analyzeSheet(name,ws,forceHdr){
  const aoa=readSheet(ws,1500,60).filter((r,i,arr)=>true);
  if(!aoa.some(r=>r&&r.some(c=>c!=null)))return null;
  let best=0,score=-1;
  if(forceHdr!=null)best=forceHdr;else for(let i=0;i<Math.min(30,aoa.length);i++){const m=autoMap(aoa[i]||[]);const s=Object.keys(m).length+(m.date!=null?3:0);if(s>score){score=s;best=i}}
  const width=Math.max(0,...aoa.slice(best,best+50).map(r=>(r||[]).length));
  const hdr=aoa[best]||[];
  const headers=Array.from({length:width},(_,i)=>txt(hdr[i])||`Column ${colLetter(i)}`);
  const rows=aoa.slice(best+1),map=autoMap(headers);
  if(map.date==null)for(let c=0;c<width;c++){if(rows.slice(0,25).filter(r=>r&&parseDate(r[c])).length>=5){map.date=c;break}}
  return {name,hdr:best,headers,rows,map,score:Object.keys(map).length+(map.date!=null?3:0)};
}
function buildGeneric(sh){
  const out={};
  for(const row of sh.rows){if(!row)continue;const ds=parseDate(row[sh.map.date]);if(!ds)continue;const rec=out[ds]||(out[ds]={});
    for(const k of NUM_KEYS){if(sh.map[k]==null)continue;const v=num(row[sh.map[k]]);if(v!=null)rec[k]=(rec[k]||0)+v}}
  const vals=Object.values(out).map(r=>r.pmPct).filter(v=>v!=null);
  if(vals.length&&Math.max(...vals.map(Math.abs))<=1.5)Object.values(out).forEach(r=>{if(r.pmPct!=null)r.pmPct*=100});
  for(const ds in out)if(!Object.keys(out[ds]).length)delete out[ds];
  return out;
}

module.exports={XLSX,detectType,parseKnown,analyzeSheet,buildGeneric,parseDate};
