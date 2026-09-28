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
function detectType(wb,fname=''){
  const names=wb.SheetNames;
  for(const n of names.slice(0,24)){const a=readSheet(wb.Sheets[n],40,30);if(a.some(r=>r&&r.some(c=>/%\s*completion/i.test(txt(c))))&&a.some(r=>r&&r.some(c=>/pm'?s\s*planned|^equipment$/i.test(txt(c)))))return 'pm'}
  if(names.some(n=>/^mtd$/i.test(n.trim()))){const a=readSheet(wb.Sheets[names.find(n=>/^mtd$/i.test(n.trim()))],20,10);if(a.some(r=>r&&r.some(c=>/^fv hls$/i.test(txt(c)))))return 'procvol'}
  if(names.some(n=>/^summary$/i.test(n.trim()))&&(/oee/i.test(fname)||names.some(n=>/utili[sz]ation/i.test(n))))return 'oee';
  if(names.some(n=>/^summary report$/i.test(n.trim()))&&(/ftr/i.test(fname)||names.some(n=>/^ftr/i.test(n.trim()))))return 'ftr';
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
  const rows={},targets={},seenFilt=new Set();
  const row=ds=>rows[ds]||(rows[ds]={});
  let mo=monthFromText(fname),yr=yearFromText(fname);
  for(const n of wb.SheetNames){
    const m=n.trim().match(/^([A-Za-z]+)\.?\s*(\d{1,2})(st|nd|rd|th)?\b/);if(!m)continue;
    const sm=monthFromText(m[1]);if(sm==null)continue;
    const a=readSheet(wb.Sheets[n],160,20);
    let y=yr;if(y==null){const d=parseDate(a[3]&&a[3][2]);if(d)y=+d.slice(0,4)}
    if(y==null)continue;
    if(mo!=null&&sm!==mo)continue;
    const ds=dstr(y,sm,+m[2]);if(!parseDate(ds))continue;

    // 1) official process loss (volume) month-to-date / year-to-date for this report day
    const h=a.findIndex(r=>r&&r.some(c=>/process loss \(volume\)/i.test(txt(c))));
    if(h>=0){
      const hr=a[h+1]||[];let ca=null,cm=null,cy=null,ct=null;
      hr.forEach((c,i)=>{const t=txt(c).toLowerCase();if(t==='actual'&&ca==null)ca=i;else if(t==='mtd'&&cm==null)cm=i;else if(t==='ytd'&&cy==null)cy=i;else if(t==='target'&&ct==null)ct=i});
      const brands=[],r2=v=>v==null?null:+v.toFixed(3);
      if(cm!=null)for(let r=h+2;r<Math.min(h+22,a.length);r++){
        const rw=a[r]||[];
        if(txt(rw[1])!==''){   // one row per brand: Actual, MTD, YTD, Target
          const act=ca!=null?num(rw[ca]):null,mtd=num(rw[cm]),ytd=cy!=null?num(rw[cy]):null,tg=ct!=null?num(rw[ct]):null;
          const ok=v=>v!=null&&v>0&&v<50?v:null;
          if(ok(act)!=null||ok(mtd)!=null||ok(ytd)!=null)brands.push([txt(rw[1]),r2(ok(act)),r2(ok(mtd)),r2(ok(ytd)),r2(tg)]);
          continue;
        }const v=num(rw[cm]);if(v==null||v<=0||v>50)continue;
        row(ds).plMtd=+v.toFixed(4);
        const yv=cy!=null?num(rw[cy]):null;if(yv!=null&&yv>0&&yv<50)row(ds).plYtd=+yv.toFixed(4);
        const tv=ct!=null?num(rw[ct]):null;if(tv!=null&&tv>0){row(ds).plTgt=tv;targets.loss=tv}
        if(brands.length)row(ds).plBrands=brands;
        break;
      }
    }

    // 2) actual process loss (volume) for each filtration: (FV + GFE - BBT) / (FV + GFE), dated by the filtration date
    const fh=a.findIndex(r=>r&&r.some(c=>/^filt\.?\s*date$/i.test(txt(c)))&&r.some(c=>/^bbt hls$/i.test(txt(c))));
    if(fh>=0){
      const H=a[fh].map(c=>txt(c).toLowerCase());
      const cB=H.findIndex(t=>t==='brand'),cD=H.findIndex(t=>/^filt\.?\s*date$/.test(t)),cF=H.findIndex(t=>/^fv hls$/.test(t)),
            cG=H.findIndex(t=>/^gfe/.test(t)),cT=H.findIndex(t=>/^bbt hls$/.test(t));
      for(let r=fh+1;r<a.length;r++){
        const rw=a[r]||[];if(rw.some(c=>/process loss \(volume\)/i.test(txt(c))))break;
        const fd=parseDate(rw[cD]);const fv=num(rw[cF])||0,g=cG>=0?(num(rw[cG])||0):0,bbt=num(rw[cT]);
        if(!fd||!(fv+g>0)||bbt==null||bbt<=0)continue;
        if(mo!=null&&+fd.slice(5,7)-1!==mo)continue;   // earlier months' filtrations repeated at the top of a new month's report
        const key=[txt(rw[cB]).toUpperCase(),fd,fv,g,bbt].join('|');if(seenFilt.has(key))continue;seenFilt.add(key);
        const t=row(fd);t.plDayFv=+((t.plDayFv||0)+fv+g).toFixed(3);t.plDayBbt=+((t.plDayBbt||0)+bbt).toFixed(3);
      }
    }

    // 3) brewhouse extract recovery: brewing loss = 100 - extract recovery
    const bh=a.findIndex(r=>r&&r.some(c=>/mtd no\.? of brews/i.test(txt(c))));
    if(bh>=0){
      const H=a[bh].map(c=>txt(c).toLowerCase()),cn=H.findIndex(t=>/mtd no\.? of brews/.test(t));
      const after=t=>H.findIndex((x,i)=>i>cn&&x===t),cm=after('mtd'),cy=after('ytd'),ct=after('target');
      let cb=-1;H.forEach((x,i)=>{if(i<cn&&x==='brand')cb=i});   // the brand column nearest the block (another table sits to its left)
      if(cm>=0)for(let r=bh+1;r<Math.min(bh+18,a.length);r++){
        const rw=a[r]||[];if(cb>=0&&txt(rw[cb])!=='')continue;const v=num(rw[cm]);if(v==null||v<50||v>110)continue;
        row(ds).bwMtd=+v.toFixed(4);
        const nb=num(rw[cn]);if(nb!=null&&nb>=0)row(ds).brewsMtd=nb;
        const yv=cy>=0?num(rw[cy]):null;if(yv!=null&&yv>50&&yv<=110)row(ds).bwYtd=+yv.toFixed(4);
        const tv=ct>=0?num(rw[ct]):null;if(tv!=null&&tv>50&&tv<=100){row(ds).bwTgt=tv;targets.brew=+(100-tv).toFixed(3)}
        break;
      }
    }
    // 4) Daily Brewing Plan: average brews per day MTD / YTD / target
    const ph=a.findIndex(r=>r&&r.some(c=>/average brews per day/i.test(txt(c))));
    if(ph>=0){
      for(let r=ph;r<Math.min(ph+4,a.length);r++){
        const rw=a[r]||[],cm=rw.findIndex(c=>txt(c).toLowerCase()==='mtd');if(cm<0)continue;
        const cy=rw.findIndex((c,i)=>i>cm&&txt(c).toLowerCase()==='ytd'),ct=rw.findIndex((c,i)=>i>cm&&txt(c).toLowerCase()==='target');
        const nx=a[r+1]||[],m=num(nx[cm]),y=cy>=0?num(nx[cy]):null,t=ct>=0?num(nx[ct]):null;
        if(m!=null&&m>=0&&m<50)row(ds).bpdMtd=+m.toFixed(4);
        if(y!=null&&y>=0&&y<50)row(ds).bpdYtd=+y.toFixed(4);
        if(t!=null&&t>0&&t<50){row(ds).bpdTgt=t;targets.bpd=t}
        break;
      }
    }
    const eh=a.findIndex(r=>r&&r.some(c=>/^extr\.?\s*rec/i.test(txt(c))));
    if(eh>=0&&eh<40){
      const ce=a[eh].findIndex(c=>/^extr\.?\s*rec/i.test(txt(c)));let sum=0,k=0;
      for(let r=eh+1;r<Math.min(eh+30,a.length);r++){const rw=a[r]||[];if(rw.some(c=>/cooling temp|^brand$/i.test(txt(c))))break;
        const v=num(rw[ce]);if(v!=null&&v>50&&v<=110&&txt(rw[2])!==''){sum+=v;k++}}
      if(k){row(ds).bwDay=+(sum/k).toFixed(4);row(ds).bwBrews=k}
    }
  }
  return {rows,targets};
}

/* ---- monthly process loss workbooks (e.g. September_2026.xls): FV and BBT volumes for the month ---- */
function parseProcVol(wb,fname){
  const mo=monthFromText(fname),yr=yearFromText(fname);
  if(mo==null||yr==null)throw new Error('Name the file after its month, e.g. September_2026.xls');
  const sn=wb.SheetNames.find(n=>/^mtd$/i.test(n.trim()));const a=readSheet(wb.Sheets[sn],60,12);
  const h=a.findIndex(r=>r&&r.some(c=>/^fv hls$/i.test(txt(c))));const H=a[h].map(c=>txt(c).toLowerCase());
  const cF=H.findIndex(t=>t==='fv hls'),cB=H.findIndex(t=>t==='bbt hls');
  for(let r=h+1;r<a.length;r++){const rw=a[r]||[];if(rw.some(c=>/^total$/i.test(txt(c)))){
    const fv=num(rw[cF]),bbt=num(rw[cB]);if(fv==null||bbt==null)break;
    return {rows:{},targets:{},monthly:{[dstr(yr,mo,1).slice(0,7)]:{vol:{fv:+fv.toFixed(3),bbt:+bbt.toFixed(3)}}}};}}
  throw new Error('Could not find the Total row on the MTD sheet.');
}
/* ---- month-by-month KPI tables (OEE summary, FTR summary report) ---- */
function monthCols(rw){const m={};(rw||[]).forEach((c,i)=>{const t=txt(c);if(/^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i.test(t)){const k=monthFromText(t);if(m[k]==null)m[k]=i}});return m}
function pctRow(rw,cols){
  const raw={};for(const k in cols){const v=num(rw[cols[k]]);if(v!=null)raw[k]=v}
  const vals=Object.values(raw).sort((x,y)=>x-y);if(!vals.length)return {};
  const asFrac=vals[Math.floor(vals.length/2)]<=1.5,out={};
  for(const k in raw){const p=asFrac?(raw[k]<=1.5?raw[k]*100:null):raw[k];if(p!=null&&p>0&&p<=100)out[k]=+p.toFixed(3)}
  return out;
}
function kpiTable(a,label,fname,key){
  const yr=yearFromText(fname);if(yr==null)throw new Error('Put the year in the file name, e.g. ..._2026.xlsx');
  const lab=r=>(a[r]||[]).slice(0,4).map(txt).map(t=>t.toLowerCase());
  const mr=a.findIndex((r,i)=>lab(i).includes('mtd')&&lab(i+1).includes('ytd'));
  if(mr<0)throw new Error(`Could not find the MTD / YTD rows for ${label}.`);
  let hr=-1;for(let i=mr-1;i>=0&&i>=mr-12;i--){if(Object.keys(monthCols(a[i])).length>=6){hr=i;break}}
  if(hr<0)throw new Error(`Could not find the month headings for ${label}.`);
  const cols=monthCols(a[hr]),mtd=pctRow(a[mr],cols),ytd=pctRow(a[mr+1],cols);
  let target=null;
  const tc=(a[hr]||[]).findIndex(c=>/^target$/i.test(txt(c)));
  if(tc>=0){const v=num(a[mr][tc]);if(v!=null)target=v<=1.5?v*100:v}
  if(target==null){for(let r=mr+2;r<mr+5&&r<a.length;r++){if(lab(r).includes('target')){const v=(a[r]||[]).map(num).find((x,i)=>x!=null&&i>0);if(v!=null)target=v<=1.5?v*100:v}}}
  const monthly={};
  for(let m=0;m<12;m++){if(mtd[m]==null&&ytd[m]==null)continue;const id=dstr(yr,m,1).slice(0,7);
    monthly[id]={kpi:{[key+'Mtd']:mtd[m]??null,[key+'Ytd']:ytd[m]??(m===0?mtd[m]:null)}};
    if(target!=null)monthly[id].targets={[key]:+target.toFixed(3)}}
  return {rows:{},targets:target!=null?{[key]:target}:{},monthly};
}
function parseOEE(wb,fname){
  const sn=wb.SheetNames.find(n=>/^summary$/i.test(n.trim()));const a=readSheet(wb.Sheets[sn],80,30);
  const top=a.findIndex(r=>r&&r.slice(0,4).some(c=>/^oee$/i.test(txt(c))));
  return kpiTable(top>=0?a.slice(top):a,'OEE',fname,'oee');
}
function parseFTR(wb,fname){
  const sn=wb.SheetNames.find(n=>/^summary report$/i.test(n.trim()));const a=readSheet(wb.Sheets[sn],200,30);
  return kpiTable(a,'FTR',fname,'ftr');
}
const MON3=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
function parsePM(wb,fname){
  const found=[];
  for(const n of wb.SheetNames){
    const a=readSheet(wb.Sheets[n],80,40);
    const pr=a.findIndex(r=>r&&r.some(c=>/%\s*completion/i.test(txt(c))));if(pr<0)continue;
    const hr=a.findIndex(r=>r&&r.some(c=>txt(c)==='%'));
    let pc=hr>=0?a[hr].findIndex(c=>txt(c)==='%'):-1;
    let v=pc>=0?num(a[pr][pc]):null;
    if(v==null){const vals=(a[pr]||[]).map(num).filter(x=>x!=null);v=vals.length?vals[vals.length-1]:null}
    if(v==null)continue;if(v<=1.5)v*=100;if(v<0||v>100)continue;
    // month from the sheet name (e.g. "Aug PM"), else the date under the title; year from that date, else the file name
    let mo=monthFromText(n),yr=null,d=null;
    for(let r=0;r<4&&!d;r++)for(const c of (a[r]||[])){const x=parseDate(c);if(x){d=x;break}}
    if(d){yr=+d.slice(0,4);if(mo==null)mo=+d.slice(5,7)-1}
    if(yr==null)yr=yearFromText(n)||yearFromText(fname);
    if(mo==null||yr==null)continue;
    found.push({id:dstr(yr,mo,1).slice(0,7),v:+v.toFixed(3)});
  }
  if(!found.length)throw new Error('Could not find a "% COMPLETION" row with a month on any sheet.');
  found.sort((x,y)=>x.id.localeCompare(y.id));
  const monthly={},acc={};
  for(const f of found){const y=f.id.slice(0,4),s=acc[y]||(acc[y]={t:0,n:0});s.t+=f.v;s.n++;
    monthly[f.id]={kpi:{pmMtd:f.v,pmYtd:+(s.t/s.n).toFixed(3)}}}
  return {rows:{},targets:{},monthly};
}
function parseKnown(wb,fname){
  const t=detectType(wb,fname);
  if(t==='pm')return {type:'PM compliance',...parsePM(wb,fname)};
  if(t==='procvol')return {type:'Monthly process loss (volume)',...parseProcVol(wb,fname)};
  if(t==='oee')return {type:'OEE summary',...parseOEE(wb,fname)};
  if(t==='ftr')return {type:'FTR summary',...parseFTR(wb,fname)};
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
