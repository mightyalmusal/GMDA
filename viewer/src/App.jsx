import { useState, useEffect, useCallback, useMemo } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { syncIfNeeded } from './lib/sync.js';
import { loadAllData } from './lib/db.js';

const COLORS = {
  blue:"#3B82F6",teal:"#14B8A6",amber:"#F59E0B",coral:"#F97316",
  pink:"#EC4899",purple:"#8B5CF6",green:"#22C55E",red:"#EF4444",
  indigo:"#6366F1",cyan:"#06B6D4",lime:"#84CC16",gray:"#6B7280",
};
const CHART_COLORS=["#3B82F6","#14B8A6","#F59E0B","#F97316","#EC4899","#8B5CF6","#22C55E","#EF4444","#6366F1","#06B6D4"];

const DIVISIONS=["Retail","Ecomm","HR","LSA","Desktop","LTS"];
const LOBS=["Desktop","LTS","LSA","Brand Value","Housebrand","Lazada"];
const DESKTOP_SEGS=["Productivity","Gaming","High-end Gamer"];
const LTS_SEGS=["Productivity","Gaming"];
const LSA_SEGS=["Alabang","Angeles","Bacoor","Baliwag","Fairview","Manila","Marikina","Muntinlupa","Novaliches","Pasig","Quezon City","San Jose del Monte","Taytay"];
const ALL_SEGS=[...DESKTOP_SEGS,...LTS_SEGS,...LSA_SEGS,"DDS"];
const OBJECTIVES=["Inquiry","Engagement","Awareness","Sales"];
const LSA_DASH_SEGMENTS=["Alabang","Bacoor","Fairview","Oasis","Marikina","Makati","Shaw Blvd","Manila","Monumento","San Fernando","Sucat","Touchpoint","Malolos","Angeles","Baliwag"];
const DESKTOP_DASH_SEGMENTS=["Productivity","Gaming","High-end Gamer","Desktop Engagement","Expert-ease"];
const DESKTOP_INQUIRY_TARGET_KEYS=["Productivity","Gaming","High-end Gamer"];
const LTS_DASH_SEGMENTS=["Productivity","Gaming"];
const LTS_INQUIRY_BUDGET_KEYS=["Productivity","Gaming"];
const LTS_ENGAGEMENT_BUDGET_KEYS=["Productivity Engagement","Gaming Engagement"];
const LTS_ALL_BUDGET_KEYS=[...LTS_INQUIRY_BUDGET_KEYS,...LTS_ENGAGEMENT_BUDGET_KEYS];
const MONTH_NAMES=["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
const MONTH_OPTIONS=[2025,2026].flatMap(y=>MONTH_NAMES.map(m=>`${y} ${m}`));
const MONTH_IDX=Object.fromEntries(MONTH_NAMES.map((m,i)=>[m,i]));
const MONTH_RANGES={"2026 JANUARY":{since:"2026-01-01",until:"2026-01-31"},"2026 FEBRUARY":{since:"2026-02-01",until:"2026-02-28"},"2026 MARCH":{since:"2026-03-01",until:"2026-03-31"},"2026 APRIL":{since:"2026-04-01",until:"2026-04-30"},"2026 MAY":{since:"2026-05-01",until:"2026-05-31"},"2026 JUNE":{since:"2026-06-01",until:"2026-06-30"}};
const LS_KEY="easypc_v3";
const THEME_KEY="easypc_theme";
const LOCAL_ONLY=(import.meta.env.VITE_LOCAL_ONLY||"false").toLowerCase()==="true";
const AAD_TENANT_ID=import.meta.env.VITE_AAD_TENANT_ID||"973ec11f-980d-4bd7-9443-fe528f0a752b";
const AAD_CLIENT_ID=import.meta.env.VITE_AAD_CLIENT_ID||"e7c8038f-4c5a-4be8-bce1-a3d42e0e38f5";
const AAD_ALLOWED_EMAILS=(import.meta.env.VITE_ALLOWED_EMAILS||"").split(",").map(v=>v.trim().toLowerCase()).filter(Boolean);
const DEFAULT_MAPPING_OPTIONS={divisions:[...DIVISIONS],lobs:[...LOBS],segments:[...ALL_SEGS],objectives:[...OBJECTIVES]};

const fmt=v=>`₱${Number(v||0).toLocaleString("en-PH",{minimumFractionDigits:0,maximumFractionDigits:0})}`;
const fmtN=v=>Number(v||0).toLocaleString();
const fmtP=v=>`${Number(v||0).toFixed(2)}%`;
const fmtK=v=>v>=1000000?`₱${(v/1000000).toFixed(2)}M`:v>=1000?`₱${(v/1000).toFixed(1)}K`:`₱${v.toFixed(0)}`;
const fmtMoneyExact=v=>`₱${Number(v||0).toLocaleString("en-PH",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const fmtPctExact=v=>`${Number(v||0).toLocaleString("en-PH",{minimumFractionDigits:2,maximumFractionDigits:2})}%`;
const CTR_BENCHMARK_ROWS=[
  {rating:"Excellent",range:"Over 2.0%",context:"Top-tier creative; highly resonant with a warm or high-intent audience."},
  {rating:"Good",range:"1.2% - 2.0%",context:"Healthy, competitive performance for most retail and consumer industries."},
  {rating:"Average",range:"0.9% - 1.2%",context:"Acceptable baseline, but likely room for creative or targeting optimization."},
  {rating:"Needs Work",range:"Under 0.9%",context:"High risk of creative fatigue or ad-to-audience mismatch."},
];

function getCtrBenchmark(ctr){
  const value=Number(ctr)||0;
  if(value>2.0)return{tier:"Excellent",range:"Over 2.0%",tone:"up"};
  if(value>=1.2)return{tier:"Good",range:"1.2% - 2.0%",tone:"up"};
  if(value>=0.9)return{tier:"Average",range:"0.9% - 1.2%",tone:"neutral"};
  return{tier:"Needs Work",range:"Under 0.9%",tone:"down"};
}

function loadLS(){try{return JSON.parse(localStorage.getItem(LS_KEY))||null;}catch{return null;}}
function saveLS(s){try{localStorage.setItem(LS_KEY,JSON.stringify(s));}catch{}}
function loadTheme(){try{const t=localStorage.getItem(THEME_KEY);return t==="light"?"light":"dark";}catch{return "dark";}}
function saveTheme(t){try{localStorage.setItem(THEME_KEY,t);}catch{}}
function localMetaFromRows(rows){
  const safe=Array.isArray(rows)?rows:[];
  const days=safe.map(r=>String(r?.day||"")).filter(Boolean).sort();
  const uniqAccounts=[...new Set(safe.map(r=>String(r?.account_id||"")).filter(Boolean))];
  return {
    totalRows:safe.length,
    fetchedAt:new Date().toISOString(),
    since:days[0]||null,
    until:days[days.length-1]||null,
    accountsQueried:uniqAccounts.length,
    accountsSucceeded:uniqAccounts.length,
    discoveredAccounts:uniqAccounts,
    errors:[],
    businessNames:[],
    fromCache:true,
    localMode:true,
  };
}

const DEFAULT_SETTINGS={mappingOptions:DEFAULT_MAPPING_OPTIONS,defaultMonth:"2026 MARCH",budgets:{"2026 JANUARY":1590000,"2026 FEBRUARY":1795298,"2026 MARCH":1897079},targets:{"2026 JANUARY":40000,"2026 FEBRUARY":43000,"2026 MARCH":46000},desktopBudgets:{},desktopTargets:{},ltsBudgets:{},ltsTargets:{},lsaBudgets:{},lsaTargets:{}};
let API_BEARER_TOKEN="";
function setApiAuthToken(token){API_BEARER_TOKEN=String(token||"");}
function apiHeaders(){const headers={"Content-Type":"application/json"};if(API_BEARER_TOKEN)headers.Authorization=`Bearer ${API_BEARER_TOKEN}`;return headers;}


function applyIdentifiers(rows,identifiers){
  return rows.map(r=>{
    const id=identifiers.find(i=>(i.adsetId&&i.adsetId===r.adset_id)||(i.adset===r.adset_name));
    if(id) return{...r,_meta:{div:id.division,lob:id.lob,seg:id.segment,obj:id.objective}};
    return r;
  });
}
function filterMonth(rows,month){
  const [yearTxt,monthName]=String(month||"").split(" ");
  const yr=Number(yearTxt);
  const mo=MONTH_IDX[monthName];
  if(!Number.isFinite(yr)||mo==null)return rows;
  return rows.filter(r=>{const d=new Date(r.day+"T00:00:00");return d.getMonth()===mo&&d.getFullYear()===yr;});
}
function aggRows(rows){return{spend:rows.reduce((s,r)=>s+r.spend,0),inquiries:rows.reduce((s,r)=>s+r.inquiries,0),post_engagement:rows.reduce((s,r)=>s+r.post_engagement,0),reach:rows.reduce((s,r)=>s+r.reach,0),impressions:rows.reduce((s,r)=>s+r.impressions,0),clicks:rows.reduce((s,r)=>s+r.clicks,0)};}
function groupRows(rows,key){const map={};rows.forEach(r=>{const k=r._meta?.[key]||"Unknown";if(!map[k])map[k]={key:k,spend:0,inquiries:0,post_engagement:0,reach:0,impressions:0,clicks:0};map[k].spend+=r.spend;map[k].inquiries+=r.inquiries;map[k].post_engagement+=r.post_engagement;map[k].reach+=r.reach;map[k].impressions+=r.impressions;map[k].clicks+=r.clicks;});return Object.values(map).sort((a,b)=>b.spend-a.spend);}
function dailySeries(rows,metric="inquiries"){const map={};rows.forEach(r=>{map[r.day]=(map[r.day]||0)+r[metric];});return Object.entries(map).sort(([a],[b])=>a.localeCompare(b));}
function isoToUTCDate(iso){return new Date(`${iso}T00:00:00Z`);}
function utcDateToIso(d){return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;}
function addDaysUTC(d,days){const n=new Date(d.getTime());n.setUTCDate(n.getUTCDate()+days);return n;}
function weekInfoFromIsoDay(day){
  const base=isoToUTCDate(day);
  const weekday=base.getUTCDay()||7;
  const start=addDaysUTC(base,1-weekday);
  const end=addDaysUTC(start,6);
  const thursday=addDaysUTC(start,3);
  const weekYear=thursday.getUTCFullYear();
  const yearStart=new Date(Date.UTC(weekYear,0,1));
  const weekNo=Math.ceil((((thursday-yearStart)/86400000)+1)/7);
  return {weekYear,weekNo,start,end,startIso:utcDateToIso(start),endIso:utcDateToIso(end)};
}

// CSS
const css=`
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
html,body,#root{height:100%}
body{background:#0A0C10;color:#E2E8F0;font-family:'DM Sans',sans-serif}
::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-track{background:#0A0C10}::-webkit-scrollbar-thumb{background:#1E2D40;border-radius:2px}
.app{display:flex;height:100vh;overflow:hidden}
.sidebar{width:218px;min-width:218px;background:#0D1117;border-right:1px solid #151E2A;display:flex;flex-direction:column;overflow:hidden;transition:width .22s ease,min-width .22s ease}
.sidebar.collapsed{width:72px;min-width:72px}
.logo-wrap{padding:20px 17px 13px;border-bottom:1px solid #151E2A;display:flex;align-items:flex-start;justify-content:space-between;gap:8px}
.logo-mark{font-size:11px;font-weight:600;letter-spacing:.12em;color:#2D3B50;text-transform:uppercase}
.logo-name{font-size:17px;font-weight:600;color:#F1F5F9;margin-top:3px}
.sidebar.collapsed .logo-name,.sidebar.collapsed .logo-mark,.sidebar.collapsed .nav-lbl{display:none}
.collapse-btn{width:24px;height:24px;border-radius:6px;border:1px solid #182838;background:#0A111A;color:#64748B;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .12s}
.collapse-btn:hover{color:#93C5FD;border-color:#263848}
.nav-sec{padding:13px 9px 4px}
.nav-lbl{font-size:10px;font-weight:600;color:#1E2D3D;letter-spacing:.13em;text-transform:uppercase;padding:0 8px;margin-bottom:5px}
.nav-item{display:flex;align-items:center;gap:9px;padding:7px 9px;border-radius:7px;cursor:pointer;font-size:14px;color:#4A5568;margin-bottom:1px;border:none;background:none;width:100%;text-align:left;transition:all .12s}
.nav-item:hover{background:#0F1720;color:#94A3B8}
.nav-item.active{background:#111E30;color:#60A5FA;font-weight:500}
.nav-item svg{width:13px;height:13px;flex-shrink:0;opacity:.5}
.nav-item.active svg{opacity:1}
.nav-text{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sidebar.collapsed .nav-item{justify-content:center;padding:8px}
.sidebar.collapsed .nav-text{display:none}
.main{flex:1;display:flex;flex-direction:column;min-width:0;overflow:hidden}
.topbar{background:#0D1117;border-bottom:1px solid #151E2A;padding:12px 22px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.pg-title{font-size:16px;font-weight:500;color:#F1F5F9}
.pg-sub{font-size:11.5px;color:#2D3B50;margin-top:1px}
.topbar-r{display:flex;align-items:center;gap:7px}
.btn{padding:6px 12px;border-radius:7px;font-size:13px;font-weight:500;cursor:pointer;border:1px solid #1A2838;background:#0E1820;color:#4A5568;font-family:'DM Sans',sans-serif;transition:all .12s}
.btn:hover{border-color:#263848;color:#94A3B8;background:#131E2C}
.btn:disabled{opacity:.35;cursor:not-allowed}
.btn-p{background:#0E1E30;color:#60A5FA;border-color:#182840}
.btn-p:hover:not(:disabled){background:#121E30;color:#93C5FD}
.btn-d{background:#180808;color:#F87171;border-color:#2E1010}
.btn-d:hover{background:#200A0A}
.scroll-area{flex:1;overflow-y:auto;padding:18px 22px;scrollbar-width:auto;scrollbar-color:#3B82F6 #0A1018}
.scroll-area::-webkit-scrollbar{width:14px;height:14px}
.scroll-area::-webkit-scrollbar-track{background:#0A1018;border-radius:8px}
.scroll-area::-webkit-scrollbar-thumb{background:#3B82F6;border-radius:8px;border:2px solid #0A1018}
.g4{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:11px;margin-bottom:14px}
.g3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:11px;margin-bottom:14px}
.g2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:11px;margin-bottom:14px}
.kpi{background:#0D1117;border:1px solid #151E2A;border-radius:10px;padding:14px 16px}
.kpi-hdr{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:7px}
.kpi-lbl{font-size:10px;font-weight:500;color:#2D3B50;text-transform:uppercase;letter-spacing:.1em;margin-bottom:7px}
.kpi-hdr .kpi-lbl{margin-bottom:0}
.kpi-info-wrap{position:relative;display:inline-flex;align-items:center}
.kpi-info-btn{width:17px;height:17px;border-radius:999px;border:1px solid #1B2A3C;background:#0A1626;color:#93C5FD;font-size:10px;font-family:'DM Mono',monospace;line-height:1;cursor:pointer;display:inline-flex;align-items:center;justify-content:center}
.kpi-info-btn:hover{background:#102036}
.kpi-info-btn:focus-visible{outline:2px solid #60A5FA;outline-offset:1px}
.kpi-tip{position:absolute;right:0;top:calc(100% + 8px);width:min(430px,82vw);background:#0B1220;border:1px solid #203247;border-radius:10px;padding:10px 11px;box-shadow:0 12px 28px rgba(2,8,23,.45);opacity:0;pointer-events:none;transform:translateY(-4px);transition:opacity .14s ease,transform .14s ease;z-index:90}
.kpi-info-wrap:hover .kpi-tip,.kpi-info-wrap:focus-within .kpi-tip{opacity:1;pointer-events:auto;transform:translateY(0)}
.kpi-tip-title{font-size:11px;font-weight:600;color:#BFDBFE;margin-bottom:7px}
.kpi-tip-head,.kpi-tip-row{display:grid;grid-template-columns:minmax(78px,95px) minmax(88px,98px) minmax(0,1fr);gap:8px;align-items:start}
.kpi-tip-head{font-size:9.5px;letter-spacing:.08em;text-transform:uppercase;color:#64748B;padding-bottom:5px;border-bottom:1px solid #162434;margin-bottom:5px}
.kpi-tip-row{padding:5px 0;border-bottom:1px solid #111C2B}
.kpi-tip-row:last-child{border-bottom:none;padding-bottom:0}
.kpi-tip-rating{font-size:11px;color:#E2E8F0;font-weight:600}
.kpi-tip-range{font-size:11px;color:#BFDBFE;font-family:'DM Mono',monospace}
.kpi-tip-context{font-size:11px;color:#94A3B8;line-height:1.35}
.kpi-val{font-size:23px;font-weight:600;color:#F1F5F9;font-family:'DM Mono',monospace;line-height:1;letter-spacing:-.02em}
.kpi-sub{font-size:12px;color:#374151;margin-top:6px}
.kpi-badge{display:inline-flex;font-size:11px;font-weight:500;padding:2px 8px;border-radius:20px;margin-top:6px}
.b-up{background:#071610;color:#34D399;border:1px solid #0E2818}
.b-dn{background:#160707;color:#F87171;border:1px solid #281010}
.b-nu{background:#091220;color:#60A5FA;border:1px solid #121E30}
.card{background:#0D1117;border:1px solid #151E2A;border-radius:10px;padding:16px;margin-bottom:14px}
.card-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:13px}
.card-ttl{font-size:14px;font-weight:500;color:#4A5568}
.sec-ttl{font-size:11px;font-weight:600;color:#2D3B50;text-transform:uppercase;letter-spacing:.12em;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid #111A24}
.tw{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12.5px}
thead th{text-align:left;font-size:10px;font-weight:500;color:#2D3B50;text-transform:uppercase;letter-spacing:.08em;padding:6px 10px;border-bottom:1px solid #111A24;white-space:nowrap}
tbody td{padding:8px 10px;border-bottom:1px solid #0D1420;color:#4A5568;font-family:'DM Mono',monospace;font-size:12.5px}
tbody tr:hover td{background:#080E18}
tbody tr:last-child td{border-bottom:none}
.td-p{color:#94A3B8!important;font-weight:500;font-family:'DM Sans',sans-serif!important;font-size:13px!important}
.pill{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:500;white-space:nowrap;font-family:'DM Sans',sans-serif}
.pl-blue{background:#071828;color:#60A5FA;border:1px solid #0F2438}.pl-teal{background:#051A14;color:#2DD4BF;border:1px solid #0A2820}.pl-amb{background:#180E00;color:#FCD34D;border:1px solid #281800}.pl-cor{background:#180A00;color:#FB923C;border:1px solid #281200}.pl-pur{background:#0E0C20;color:#A78BFA;border:1px solid #181430}.pl-grn{background:#051608;color:#4ADE80;border:1px solid #0A2410}.pl-ind{background:#0A0C20;color:#818CF8;border:1px solid #121430}.pl-gray{background:#0C1018;color:#4A5568;border:1px solid #141C28}
.hbar{display:grid;grid-template-columns:minmax(90px,120px) minmax(0,1fr) minmax(84px,110px);align-items:center;gap:8px;margin-bottom:6px}
.hbar-lbl{font-size:12px;color:#374151;text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.hbar-track{background:#0A1018;border-radius:4px;height:21px;overflow:hidden}
.hbar-fill{height:100%;border-radius:4px;transition:width .5s cubic-bezier(.4,0,.2,1)}
.hbar-val{font-size:11px;font-family:'DM Mono',monospace;color:#94A3B8;font-weight:500;text-align:right;white-space:nowrap}
.pbar-wrap{background:#0A1018;border-radius:5px;height:8px;overflow:hidden;margin:4px 0 11px}
.pbar{height:100%;border-radius:5px;transition:width .6s cubic-bezier(.4,0,.2,1)}
.pbar-row{display:flex;justify-content:space-between;font-size:11.5px;color:#374151;margin-bottom:3px}
.mini-chart{display:flex;align-items:flex-end;gap:2px;height:48px}
.mini-bar{flex:1;border-radius:2px 2px 0 0;min-height:3px}
.line-chart{position:relative;height:120px}
.donut-wrap{display:flex;align-items:center;gap:16px}
.donut-legend{flex:1}
.donut-item{display:flex;align-items:center;gap:7px;margin-bottom:7px;font-size:12.5px}
.donut-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.donut-name{color:#4A5568;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.donut-pct{font-family:'DM Mono',monospace;font-size:12px;color:#64748B}
.tabs{display:flex;gap:2px;background:#07090D;border:1px solid #111A24;border-radius:7px;padding:2px;width:fit-content}
.tab{padding:4px 12px;border-radius:6px;font-size:13px;cursor:pointer;color:#2D3B50;border:none;background:none;font-family:'DM Sans',sans-serif;transition:all .12s}
.tab.active{background:#0D1420;color:#E2E8F0;font-weight:500}
.tab:hover:not(.active){color:#64748B}
input,select,textarea{background:#0A1018;border:1px solid #182838;border-radius:7px;padding:8px 10px;color:#E2E8F0;font-size:14px;font-family:'DM Sans',sans-serif;width:100%;outline:none;transition:border-color .15s}
input:focus,select:focus{border-color:#204060}
input::placeholder{color:#2D3B50}
select option{background:#0A1018;color:#E2E8F0}
label{display:block;font-size:12px;font-weight:500;color:#374151;margin-bottom:5px}
.fr{display:grid;gap:11px;margin-bottom:11px}
.fr2{grid-template-columns:1fr 1fr}.fr3{grid-template-columns:1fr 1fr 1fr}.fr4{grid-template-columns:1fr 1fr 1fr 1fr}
.id-row{display:grid;grid-template-columns:2fr .9fr .9fr 1fr .9fr 30px;gap:5px;align-items:start;margin-bottom:6px}
.id-hdr{display:grid;grid-template-columns:2fr .9fr .9fr 1fr .9fr 30px;gap:5px;margin-bottom:5px}
.id-h{font-size:10px;font-weight:500;color:#2D3B50;text-transform:uppercase;letter-spacing:.08em;padding:0 4px}
.rm-btn{width:30px;height:34px;border-radius:6px;display:flex;align-items:center;justify-content:center;border:1px solid #280E0E;background:#120606;color:#F87171;cursor:pointer;font-size:14px;flex-shrink:0;transition:background .12s}
.rm-btn:hover{background:#1E0808}
.acc-chip{display:flex;align-items:center;justify-content:space-between;background:#0A1824;border:1px solid #142035;border-radius:7px;padding:8px 12px;margin-bottom:6px}
.acc-id{font-family:'DM Mono',monospace;font-size:11.5px;color:#60A5FA}
.seg-card{background:#080E18;border:1px solid #111A24;border-radius:9px;padding:13px;margin-bottom:0}
.seg-card-click{cursor:pointer;transition:transform .12s ease,border-color .12s ease,box-shadow .12s ease}
.seg-card-click:hover{transform:translateY(-1px);border-color:#1A2B3E;box-shadow:0 8px 20px rgba(2,8,23,.18)}
.seg-lbl{font-size:10px;font-weight:600;color:#9FB3C8;text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px}
.seg-val{font-size:21px;font-weight:600;font-family:'DM Mono',monospace;line-height:1}
.seg-sub{font-size:11px;color:#C7D2E0;margin-top:4px;margin-bottom:8px}
.share-list{display:grid;gap:8px}
.share-row{background:#08121E;border:1px solid #122236;border-radius:8px;padding:7px 9px}
.share-row-top{display:flex;justify-content:space-between;gap:8px;font-size:11.5px}
.share-name{color:#C7D2E0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.share-num{color:#E2E8F0;font-family:'DM Mono',monospace}
.share-pct{color:#93C5FD;font-family:'DM Mono',monospace}
.share-bar{height:5px;background:#0A1018;border-radius:999px;overflow:hidden;margin-top:6px}
.share-fill{height:100%;border-radius:999px}
.daily-val{display:flex;align-items:center;gap:7px;white-space:nowrap}
.daily-delta{font-size:10.5px}
.seg-modal-backdrop{position:fixed;inset:0;background:rgba(2,6,23,.72);display:flex;align-items:center;justify-content:center;padding:14px;z-index:12000}
.seg-modal{width:min(1280px,96vw);max-height:92vh;overflow:auto;background:#0D1117;border:1px solid #1A2B3E;border-radius:12px;padding:14px 14px 10px}
.seg-modal-h{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:10px}
.seg-modal-t{font-size:14px;font-weight:600;color:#BFDBFE}
.seg-modal-sub{font-size:11.5px;color:#64748B}
.seg-modal-close{border:1px solid #223348;background:#0A1626;color:#93C5FD;border-radius:8px;padding:4px 10px;cursor:pointer}
.seg-modal-close:hover{background:#102036}
.db-cell{position:relative;min-width:140px;height:22px;display:flex;align-items:center;justify-content:flex-end;padding:0 6px;border-radius:4px;overflow:hidden}
.db-fill{position:absolute;left:0;top:0;height:100%;border-right:1px solid rgba(255,255,255,.35)}
.db-spend{background:linear-gradient(90deg,rgba(16,185,129,.35),rgba(16,185,129,.18))}
.db-inq{background:linear-gradient(90deg,rgba(59,130,246,.35),rgba(59,130,246,.18))}
.db-cpi{background:linear-gradient(90deg,rgba(236,72,153,.35),rgba(236,72,153,.18))}
.db-text{position:relative;z-index:1;font-family:'DM Mono',monospace}
.drop-zone{border:1.5px dashed #1A2838;border-radius:9px;padding:24px;text-align:center;cursor:pointer;transition:all .15s;margin-bottom:12px}
.drop-zone:hover,.drop-zone.drag{border-color:#3B82F6;background:#070F1C}
.drop-zone p{font-size:13.5px;color:#374151;margin-top:6px}
.drop-zone small{font-size:12px;color:#2D3B50;margin-top:3px;display:block}
.info-box{background:#07101C;border:1px solid #10203A;border-radius:9px;padding:11px 14px;margin-bottom:13px;font-size:13px;color:#60A5FA}
.err-box{background:#140707;border:1px solid #2E1010;border-radius:9px;padding:11px 14px;margin-bottom:13px;font-size:13px;color:#F87171}
.toast{position:fixed;bottom:20px;right:20px;padding:10px 15px;border-radius:9px;font-size:13.5px;z-index:9999;animation:tIn .15s;max-width:320px}
.tok{background:#0E1E30;border:1px solid #182840;color:#93C5FD}
.ter{background:#160707;border:1px solid #2E1010;color:#FCA5A5}
@keyframes tIn{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:translateY(0)}}
.spin{width:13px;height:13px;border:2px solid #0E1E30;border-top-color:#60A5FA;border-radius:50%;animation:rot .6s linear infinite;display:inline-block;vertical-align:middle}
@keyframes rot{to{transform:rotate(360deg)}}
.fetch-bar{position:relative;height:4px;background:#0F172A;border-bottom:1px solid #1E293B;overflow:hidden}
.fetch-bar::after{content:"";position:absolute;left:-35%;top:0;height:100%;width:35%;background:linear-gradient(90deg,#0EA5E9,#60A5FA,#0EA5E9);animation:barMove 1.1s linear infinite}
@keyframes barMove{from{left:-35%}to{left:100%}}
.empty{text-align:center;padding:32px;color:#2D3B50;font-size:13.5px}
.sb-bot{padding:10px 15px 15px;margin-top:auto;font-size:11px;color:#2D3B50;display:flex;flex-direction:column;gap:5px}
.sb-meta{display:flex;flex-direction:column;gap:5px}
.theme-toggle{margin-top:8px;display:flex;align-items:center;justify-content:center;gap:8px;border:none;background:transparent;color:#94A3B8;cursor:pointer;padding:0}
.theme-pill{position:relative;width:60px;height:30px;border-radius:999px;border:1px solid #263848;background:#0E1820;display:block;transition:all .2s}
.theme-pill .ico{position:absolute;top:50%;transform:translateY(-50%);font-size:12px;opacity:.65;user-select:none}
.theme-pill .sun{left:9px;color:#FBBF24}
.theme-pill .moon{right:9px;color:#93C5FD}
.theme-thumb{position:absolute;top:2px;left:2px;width:24px;height:24px;border-radius:50%;background:#E2E8F0;box-shadow:0 1px 3px rgba(0,0,0,.3);transition:transform .2s ease}
.theme-toggle.light .theme-thumb{transform:translateX(30px)}
.theme-toggle-label{font-size:11px;color:#64748B;min-width:42px;text-align:left}
.sidebar.collapsed .sb-bot{padding:8px 8px 12px;align-items:center}
.sidebar.collapsed .sb-meta{display:none}
.sidebar.collapsed .theme-toggle-label{display:none}
.sdot{width:5px;height:5px;border-radius:50%;display:inline-block;flex-shrink:0}
.page-switch{animation:pageIn .24s ease}
@keyframes pageIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
.guard-wrap{position:relative;min-height:240px}
.guard-content.is-locked{filter:blur(7px);pointer-events:none;user-select:none}
.guard-overlay{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;padding:18px;z-index:8}
.guard-card{width:min(440px,100%);background:rgba(13,17,23,.95);border:1px solid #1F2E44;border-radius:12px;padding:16px;box-shadow:0 12px 36px rgba(2,6,23,.42);backdrop-filter:blur(4px)}
.guard-title{font-size:14px;font-weight:600;color:#BFDBFE;margin-bottom:6px}
.guard-sub{font-size:12px;color:#94A3B8;margin-bottom:10px}
.guard-err{margin-top:8px;font-size:12px;color:#FCA5A5}
.weekly-note{font-size:12.5px;color:#64748B;margin-top:-4px;margin-bottom:10px}
.weekly-table th,.weekly-table td{padding:7px 9px;font-size:13px}
.weekly-sticky{position:sticky;left:0;background:#0D1117;z-index:2;min-width:150px}
.weekly-sec td{font-weight:600;color:#93C5FD;background:#0B1320}
.weekly-section-title{z-index:3}
.weekly-metric{font-family:'DM Sans',sans-serif!important;color:#94A3B8!important}
.weekly-wrap{overflow-x:auto;padding-bottom:6px;scrollbar-width:auto;scrollbar-color:#3B82F6 #0A1018}
.weekly-wrap::-webkit-scrollbar{height:14px}
.weekly-wrap::-webkit-scrollbar-track{background:#0A1018;border-radius:8px}
.weekly-wrap::-webkit-scrollbar-thumb{background:#3B82F6;border-radius:8px;border:2px solid #0A1018}
.wk-main{font-family:'DM Mono',monospace}
.wk-delta{font-size:11px;margin-top:2px}
.wk-up{color:#34D399}
.wk-down{color:#F87171}
.wk-flat{color:#64748B}
.org-auth-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:18px;background:radial-gradient(120% 120% at 10% 0%,#0E1B2C 0%,#0A0F18 45%,#070B12 100%)}
.org-auth-card{width:min(520px,100%);background:#0D1117;border:1px solid #1A2B3E;border-radius:14px;padding:18px 18px 16px;box-shadow:0 18px 42px rgba(2,6,23,.42)}
.org-auth-kicker{font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:#60A5FA;font-weight:600;margin-bottom:8px}
.org-auth-title{font-size:24px;color:#F1F5F9;font-weight:600;line-height:1.2}
.org-auth-sub{margin-top:7px;color:#94A3B8;font-size:13px;line-height:1.45}
.org-auth-meta{margin-top:12px;padding:9px 10px;border-radius:8px;background:#07101C;border:1px solid #10203A;color:#93C5FD;font-size:12px}
.org-auth-err{margin-top:10px;padding:8px 10px;border-radius:8px;background:#140707;border:1px solid #2E1010;color:#FCA5A5;font-size:12px}
.org-auth-actions{margin-top:13px;display:flex;gap:8px;flex-wrap:wrap}
.org-user-chip{font-size:11.5px;color:#93C5FD;background:#0A1626;border:1px solid #223348;border-radius:999px;padding:4px 9px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
body.theme-light{background:#EEF3F8;color:#0F172A}
body.theme-light .sidebar{background:#F8FAFC;border-right-color:#DBE3EF}
body.theme-light .topbar{background:#F8FAFC;border-bottom-color:#DBE3EF}
body.theme-light .logo-name{color:#0F172A}
body.theme-light .logo-mark,body.theme-light .nav-lbl,body.theme-light .pg-sub{color:#64748B}
body.theme-light .nav-item{color:#475569}
body.theme-light .nav-item:hover{background:#EDF2F7;color:#1F2937}
body.theme-light .nav-item.active{background:#E2E8F0;color:#1D4ED8}
body.theme-light .main{background:#F1F5F9}
body.theme-light .card,body.theme-light .kpi{background:#FFFFFF;border-color:#DBE3EF}
body.theme-light .card-ttl,body.theme-light .kpi-lbl,body.theme-light .kpi-sub,body.theme-light .hbar-lbl,body.theme-light label{color:#475569}
body.theme-light .kpi-val,body.theme-light .pg-title{color:#0F172A}
body.theme-light .kpi-info-btn{background:#EFF6FF;border-color:#BFDBFE;color:#1D4ED8}
body.theme-light .kpi-info-btn:hover{background:#DBEAFE}
body.theme-light .kpi-tip{background:#FFFFFF;border-color:#DBE3EF;box-shadow:0 12px 28px rgba(15,23,42,.15)}
body.theme-light .kpi-tip-title{color:#1E40AF}
body.theme-light .kpi-tip-head{color:#64748B;border-bottom-color:#E2E8F0}
body.theme-light .kpi-tip-row{border-bottom-color:#E2E8F0}
body.theme-light .kpi-tip-rating{color:#0F172A}
body.theme-light .kpi-tip-range{color:#1D4ED8}
body.theme-light .kpi-tip-context{color:#475569}
body.theme-light .btn{background:#FFFFFF;border-color:#CBD5E1;color:#334155}
body.theme-light .btn:hover{background:#F8FAFC;border-color:#94A3B8;color:#0F172A}
body.theme-light .btn-p{background:#DBEAFE;border-color:#BFDBFE;color:#1D4ED8}
body.theme-light .btn-p:hover:not(:disabled){background:#CFE2FF;color:#1E40AF}
body.theme-light input,body.theme-light select,body.theme-light textarea{background:#FFFFFF;border-color:#CBD5E1;color:#0F172A}
body.theme-light select option,body.theme-light optgroup{background:#FFFFFF;color:#0F172A}
body.theme-light input::placeholder{color:#94A3B8}
body.theme-light .tw thead th{color:#64748B;border-bottom-color:#E2E8F0}
body.theme-light .tw tbody td{color:#334155;border-bottom-color:#E2E8F0}
body.theme-light tbody tr:hover td{background:#F8FAFC}
body.theme-light .theme-pill{background:#E2E8F0;border-color:#CBD5E1}
body.theme-light .theme-toggle-label{color:#475569}
body.theme-light .guard-card{background:rgba(255,255,255,.96);border-color:#CBD5E1;box-shadow:0 10px 30px rgba(15,23,42,.12)}
body.theme-light .guard-title{color:#1E40AF}
body.theme-light .guard-sub{color:#475569}
body.theme-light .guard-err{color:#B91C1C}
body.theme-light .weekly-sticky{background:#FFFFFF}
body.theme-light .weekly-sec td{background:#EFF6FF;color:#1E40AF}
body.theme-light .weekly-metric{color:#334155!important}
body.theme-light .weekly-wrap{scrollbar-color:#60A5FA #E2E8F0}
body.theme-light .weekly-wrap::-webkit-scrollbar-track{background:#E2E8F0}
body.theme-light .weekly-wrap::-webkit-scrollbar-thumb{background:#60A5FA;border-color:#E2E8F0}
body.theme-light .scroll-area{scrollbar-color:#60A5FA #E2E8F0}
body.theme-light .scroll-area::-webkit-scrollbar-track{background:#E2E8F0}
body.theme-light .scroll-area::-webkit-scrollbar-thumb{background:#60A5FA;border-color:#E2E8F0}
body.theme-light .share-row{background:#FFFFFF;border-color:#DBE3EF}
body.theme-light .share-name{color:#334155}
body.theme-light .share-num{color:#0F172A}
body.theme-light .share-pct{color:#1D4ED8}
body.theme-light .org-auth-wrap{background:radial-gradient(120% 120% at 10% 0%,#E0ECFF 0%,#EEF3F8 45%,#E8EEF6 100%)}
body.theme-light .org-auth-card{background:#FFFFFF;border-color:#DBE3EF;box-shadow:0 16px 36px rgba(15,23,42,.12)}
body.theme-light .org-auth-title{color:#0F172A}
body.theme-light .org-auth-sub{color:#475569}
body.theme-light .org-auth-meta{background:#EFF6FF;border-color:#BFDBFE;color:#1E40AF}
body.theme-light .org-user-chip{background:#EFF6FF;border-color:#BFDBFE;color:#1D4ED8}
body.theme-light .seg-modal{background:#FFFFFF;border-color:#CBD5E1}
body.theme-light .seg-modal-t{color:#1E40AF}
body.theme-light .seg-modal-sub{color:#475569}
body.theme-light .seg-modal-close{background:#EFF6FF;border-color:#BFDBFE;color:#1D4ED8}
body.theme-light .db-fill{border-right-color:rgba(15,23,42,.2)}
@media (max-width:980px){
  .app{flex-direction:column;height:auto;min-height:100vh}
  .sidebar,.sidebar.collapsed{width:100%;min-width:0;border-right:none;border-bottom:1px solid #151E2A;max-height:none}
  .sidebar.collapsed .logo-name,.sidebar.collapsed .logo-mark,.sidebar.collapsed .nav-lbl{display:block}
  .sidebar.collapsed .nav-text{display:inline}
  .sidebar.collapsed .nav-item{justify-content:flex-start;padding:7px 9px}
  .logo-wrap{padding:12px 14px 10px}
  .nav-sec{padding:8px 10px 4px}
  .nav-lbl{padding:0 2px}
  .nav-scroll{display:flex;overflow-x:auto;gap:6px;padding-bottom:4px}
  .nav-scroll .nav-item{width:auto;min-width:max-content;margin-bottom:0}
  .sb-bot{display:none}
  .main{min-width:0}
  .topbar{padding:10px 12px;flex-wrap:wrap;gap:8px}
  .topbar-r{width:100%;justify-content:flex-start;flex-wrap:wrap}
  .scroll-area{padding:12px}
  .g4,.g3,.g2{grid-template-columns:1fr}
  .fr2,.fr3,.fr4{grid-template-columns:1fr}
  .id-hdr{display:none}
  .id-row{grid-template-columns:1fr;gap:6px}
  .seg-grid{grid-template-columns:1fr!important}
  .info-box{overflow-wrap:anywhere}
}
@media (max-width:640px){
  .pg-title{font-size:18px}
  .pg-sub{font-size:11px}
  .hbar{grid-template-columns:minmax(70px,96px) minmax(0,1fr) minmax(72px,96px)}
  .seg-modal{padding:10px}
  table{font-size:12px}
  thead th,tbody td{padding:6px 8px}
  .btn{padding:6px 10px;font-size:12px}
  .tabs{max-width:100%;overflow:auto}
  .tab{white-space:nowrap}
}
`;

// Shared UI
function KPI({label,value,sub,badge,badgeUp,badgeNeutral,color,infoTitle,infoRows}){
  const hasInfo=Array.isArray(infoRows)&&infoRows.length>0;
  return(<div className="kpi">
    <div className="kpi-hdr">
      <div className="kpi-lbl">{label}</div>
      {hasInfo&&(<div className="kpi-info-wrap">
        <button type="button" className="kpi-info-btn" aria-label={`${label} benchmark details`} title="Benchmark details">i</button>
        <div className="kpi-tip" role="tooltip">
          <div className="kpi-tip-title">{infoTitle||"Benchmark"}</div>
          <div className="kpi-tip-head"><span>Rating</span><span>CTR Range</span><span>Context</span></div>
          {infoRows.map((row,i)=><div className="kpi-tip-row" key={`${row.rating}-${i}`}><span className="kpi-tip-rating">{row.rating}</span><span className="kpi-tip-range">{row.range}</span><span className="kpi-tip-context">{row.context}</span></div>)}
        </div>
      </div>)}
    </div>
    <div className="kpi-val" style={color?{color}:{}}>{value}</div>
    {sub&&<div className="kpi-sub">{sub}</div>}
    {badge&&<div className={`kpi-badge ${badgeNeutral?"b-nu":badgeUp?"b-up":"b-dn"}`}>{badge}</div>}
  </div>);
}

function HBar({data,valueFormat}){
  const max=Math.max(...data.map(d=>d.value),1);
  return(<div>{data.map((d,i)=>(<div className="hbar" key={i}><div className="hbar-lbl" title={d.label}>{d.label}</div><div className="hbar-track"><div className="hbar-fill" style={{width:`${Math.max(4,(d.value/max)*100)}%`,background:d.color||COLORS.blue}}/></div><div className="hbar-val">{valueFormat?valueFormat(d.value):fmtN(d.value)}</div></div>))}</div>);
}

function Spark({values,color}){
  const max=Math.max(...values,1);
  return(<div className="mini-chart">{values.map((v,i)=>(<div key={i} className="mini-bar" style={{background:color||COLORS.blue,opacity:.3+.7*(v/max),height:`${Math.max(5,(v/max)*100)}%`}}/>))}</div>);
}

function LineChart({series,color,showValueLabels=false}){
  if(!series||series.length<2)return null;
  const vals=series.map(([,v])=>v),max=Math.max(...vals,1),W=400,H=100,pad=6;
  const pts=series.map(([,v],i)=>[pad+(i/(series.length-1))*(W-pad*2),H-pad-((v/max)*(H-pad*2))]);
  const d=pts.map((p,i)=>i===0?`M${p[0]},${p[1]}`:`L${p[0]},${p[1]}`).join(" ");
  const fill=[...pts.map(p=>`L${p[0]},${p[1]}`),`L${pts[pts.length-1][0]},${H}`,`L${pts[0][0]},${H}`].join(" ");
  const gid=`g${(color||"").replace(/[^a-z0-9]/gi,"")}`;
  const labelStep=Math.max(1,Math.ceil(vals.length/10));
  return(<div className="line-chart"><svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="none" style={{width:"100%",height:"100%"}}><defs><linearGradient id={gid} x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={color||COLORS.blue} stopOpacity=".2"/><stop offset="100%" stopColor={color||COLORS.blue} stopOpacity="0"/></linearGradient></defs><path d={`M${pts[0][0]},${pts[0][1]} ${fill}`} fill={`url(#${gid})`}/><path d={d} fill="none" stroke={color||COLORS.blue} strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round"/>{pts.map(([x,y],i)=><circle key={i} cx={x} cy={y} r="2.2" fill={color||COLORS.blue} opacity=".7"/>)}{showValueLabels&&pts.map(([x,y],i)=>i%labelStep===0||i===pts.length-1?<text key={`lbl-${i}`} x={x} y={Math.max(8,y-6)} fill="#94A3B8" fontSize="7.5" textAnchor="middle" style={{fontFamily:"'DM Mono',monospace"}}>{fmtN(vals[i])}</text>:null)}</svg></div>);
}

function DonutChart({data,size=100,showLegend=true}){
  const total=data.reduce((s,d)=>s+d.value,0)||1;
  let off=0;const r=36,cx=50,cy=50,circ=2*Math.PI*r;
  const slices=data.map(d=>{const pct=d.value/total,dash=pct*circ,gap=circ-dash,s={...d,dash,gap,off,pct};off+=dash;return s;});
  return(<div className="donut-wrap"><svg width={size} height={size} viewBox="0 0 100 100" style={{flexShrink:0}}><circle cx={cx} cy={cy} r={r} fill="none" stroke="#0A1018" strokeWidth="13"/>{slices.map((s,i)=>(<circle key={i} cx={cx} cy={cy} r={r} fill="none" stroke={s.color} strokeWidth="13" strokeDasharray={`${s.dash} ${s.gap}`} strokeDashoffset={-s.off+circ*.25}/>))}<text x={cx} y={cy} textAnchor="middle" dominantBaseline="central" fill="#F1F5F9" fontSize="12" fontWeight="600" fontFamily="'DM Mono',monospace">{total>=1000?`${(total/1000).toFixed(1)}K`:total}</text></svg>{showLegend&&<div className="donut-legend">{slices.map((s,i)=>(<div className="donut-item" key={i}><span className="donut-dot" style={{background:s.color}}/><span className="donut-name">{s.label}</span><span className="donut-pct">{(s.pct*100).toFixed(1)}%</span></div>))}</div>}</div>);
}

function PacingRow({label,value,max,color,vfmt}){
  const pct=max>0?Math.min((value/max)*100,100):0;
  return(<div style={{marginBottom:12}}><div className="pbar-row"><span>{label}</span><span style={{fontFamily:"'DM Mono',monospace"}}>{vfmt?vfmt(value):fmtN(value)} / {vfmt?vfmt(max):fmtN(max)}</span></div><div className="pbar-wrap"><div className="pbar" style={{width:`${pct}%`,background:color||COLORS.blue}}/></div></div>);
}

const PM={"Retail":"pl-blue","Ecomm":"pl-teal","LSA":"pl-amb","HR":"pl-pur","Desktop":"pl-blue","LTS":"pl-ind","Lazada":"pl-teal","Brand Value":"pl-pur","Housebrand":"pl-grn","Inquiry":"pl-blue","Sales":"pl-teal","Engagement":"pl-amb","Awareness":"pl-grn","Productivity":"pl-blue","Gaming":"pl-ind","High-end Gamer":"pl-pur"};

// LOB Dashboard (reused for Desktop, LTS, LSA)
function LobDash({data,allData,lob,segments,colors,settings,month,bKey,tKey,allowedDivisions,allowedLobs,allowedSegments,allowedObjectives,monthlyBudgetMap,monthlyTargetMap,budgetKeys,targetKeys}){
  const [segmentModal,setSegmentModal]=useState(null);
  const [trendModal,setTrendModal]=useState(false);
  const matchesFilters=(r)=>{
    const div=r._meta?.div||"";
    const mappedLob=r._meta?.lob||"";
    const seg=r._meta?.seg||"";
    const obj=r._meta?.obj||"";

    if(Array.isArray(allowedDivisions)&&allowedDivisions.length&&!allowedDivisions.includes(div))return false;
    if(Array.isArray(allowedLobs)&&allowedLobs.length){
      if(!allowedLobs.includes(mappedLob))return false;
    }else if(!(mappedLob===lob||div===lob)){
      return false;
    }
    if(Array.isArray(allowedSegments)&&allowedSegments.length&&!allowedSegments.includes(seg))return false;
    if(Array.isArray(allowedObjectives)&&allowedObjectives.length&&!allowedObjectives.includes(obj))return false;
    return true;
  };
  const filtered=data.filter(matchesFilters);
  const total=aggRows(filtered);
  const accountsUsed=[...new Set(filtered.map(r=>String(r.account_id||"").trim()).filter(Boolean))].length;
  const [yrTxt,moName]=String(month||"").split(" ");
  const yr=Number(yrTxt);
  const mo=MONTH_IDX[moName];
  const monthRange=(Number.isFinite(yr)&&mo!=null)
    ? `${yr}-${String(mo+1).padStart(2,"0")}-01 to ${yr}-${String(mo+1).padStart(2,"0")}-${String(new Date(yr,mo+1,0).getDate()).padStart(2,"0")}`
    : "Unknown";
  const sumMonthValues=(monthMap,keys)=>{
    const row=monthMap?.[month]||{};
    const selected=Array.isArray(keys)&&keys.length?keys:Object.keys(row);
    return selected.reduce((s,key)=>s+Number(row?.[key]||0),0);
  };
  const budget=monthlyBudgetMap?sumMonthValues(monthlyBudgetMap,budgetKeys):Number(settings.budgets?.[month]||0);
  const target=monthlyTargetMap?sumMonthValues(monthlyTargetMap,targetKeys):Number(settings.targets?.[month]||0);
  const overall=aggRows(data||[]);
  const cpi=total.inquiries>0?total.spend/total.inquiries:0;
  const cpm=total.impressions>0?(total.spend/total.impressions)*1000:0;
  const ctr=total.impressions>0?(total.clicks/total.impressions)*100:0;
  const ctrBenchmark=getCtrBenchmark(ctr);
  const spPct=budget>0?(total.spend/budget)*100:0;
  const inPct=target>0?(total.inquiries/target)*100:0;
  const reachShare=overall.reach>0?(total.reach/overall.reach)*100:0;
  const impressionShare=overall.impressions>0?(total.impressions/overall.impressions)*100:0;
  const engagementShare=overall.post_engagement>0?(total.post_engagement/overall.post_engagement)*100:0;

  const bySeg=segments.map((seg,i)=>{
    const rows=filtered.filter(r=>r._meta?.seg===seg);
    const a=aggRows(rows);
    const series=dailySeries(rows).slice(-14);
    const dayMap={};
    rows.forEach(r=>{
      const k=r.day;
      if(!k)return;
      if(!dayMap[k])dayMap[k]={day:k,spend:0,inquiries:0};
      dayMap[k].spend+=Number(r.spend||0);
      dayMap[k].inquiries+=Number(r.inquiries||0);
    });
    const dailyDetails=Object.values(dayMap).sort((a,b)=>a.day.localeCompare(b.day)).map(d=>({...d,cpi:d.inquiries>0?d.spend/d.inquiries:0}));
    const segBudget=Number(settings[bKey]?.[month]?.[seg]||0);
    const segTarget=Number(settings[tKey]?.[month]?.[seg]||0);
    return{seg,a,series,color:colors[i%colors.length],segBudget,segTarget,dailyDetails};
  });

  const dailyAll=dailySeries(filtered);
  const dailyRows=dailyAll.map(([day,inquiries])=>{
    const rowsForDay=filtered.filter(r=>r.day===day);
    const spend=rowsForDay.reduce((s,r)=>s+Number(r.spend||0),0);
    const cpi=inquiries>0?spend/inquiries:0;
    return {day,inquiries,spend,cpi};
  });
  const donutData=bySeg.filter(s=>s.a.inquiries>0).map(s=>({label:s.seg,value:s.a.inquiries,color:s.color}));
  const donutSpend=bySeg.filter(s=>s.a.spend>0).map(s=>({label:s.seg,value:Math.round(s.a.spend),color:s.color}));
  const inquiryShareRows=bySeg.map(s=>{
    const pct=total.inquiries>0?(s.a.inquiries/total.inquiries)*100:0;
    return{seg:s.seg,color:s.color,inquiries:s.a.inquiries,pct};
  }).sort((a,b)=>b.inquiries-a.inquiries);
  const spendShareRows=bySeg.map(s=>{
    const pct=total.spend>0?(s.a.spend/total.spend)*100:0;
    return{seg:s.seg,color:s.color,spend:s.a.spend,pct};
  }).sort((a,b)=>b.spend-a.spend);
  const sourceParts=[
    `Month: ${month}`,
    `Date range: ${monthRange}`,
    `Rows used: ${fmtN(filtered.length)}`,
    `Accounts used: ${fmtN(accountsUsed)}`,
    `Division: ${(allowedDivisions||[]).join(", ")||"Any"}`,
    `LOB: ${(allowedLobs||[]).join(", ")||lob||"Any"}`,
    `Segment: ${(allowedSegments||[]).join(", ")||"Any"}`,
    `Objective: ${(allowedObjectives||[]).join(", ")||"Any"}`,
  ];

  const weeklyMonitoring=(lob==="Desktop")?(()=>{
    const allRows=(Array.isArray(allData)&&allData.length?allData:data).filter(matchesFilters);
    const validRows=allRows.filter(r=>r.day);
    if(!validRows.length)return null;

    const minDay=validRows.reduce((m,r)=>!m||String(r.day)<m?String(r.day):m,null);
    const maxDay=validRows.reduce((m,r)=>!m||String(r.day)>m?String(r.day):m,null);
    const minW=weekInfoFromIsoDay(minDay);
    const maxW=weekInfoFromIsoDay(maxDay);

    const weekMap={};
    validRows.forEach(r=>{
      const wi=weekInfoFromIsoDay(String(r.day));
      const k=`${wi.weekYear}-W${wi.weekNo}`;
      if(!weekMap[k]){
        weekMap[k]={key:k,weekYear:wi.weekYear,weekNo:wi.weekNo,startIso:wi.startIso,endIso:wi.endIso,total:{inquiries:0,reach:0,engagement:0,spent:0},segments:{"Productivity":{inquiries:0,reach:0,engagement:0,spent:0},"Gaming":{inquiries:0,reach:0,engagement:0,spent:0},"High-end Gamer":{inquiries:0,reach:0,engagement:0,spent:0}}};
      }
      const wk=weekMap[k];
      const inquiries=Number(r.inquiries||0),reach=Number(r.reach||0),eng=Number(r.post_engagement||0),spent=Number(r.spend||0);
      wk.total.inquiries+=inquiries;wk.total.reach+=reach;wk.total.engagement+=eng;wk.total.spent+=spent;
      const seg=r._meta?.seg;
      if(seg&&wk.segments[seg]){wk.segments[seg].inquiries+=inquiries;wk.segments[seg].reach+=reach;wk.segments[seg].engagement+=eng;wk.segments[seg].spent+=spent;}
    });

    const weeks=[];
    for(let d=minW.start;d<=maxW.end;d=addDaysUTC(d,7)){
      const wi=weekInfoFromIsoDay(utcDateToIso(d));
      const k=`${wi.weekYear}-W${wi.weekNo}`;
      const blank={key:k,weekYear:wi.weekYear,weekNo:wi.weekNo,startIso:wi.startIso,endIso:wi.endIso,total:{inquiries:0,reach:0,engagement:0,spent:0},segments:{"Productivity":{inquiries:0,reach:0,engagement:0,spent:0},"Gaming":{inquiries:0,reach:0,engagement:0,spent:0},"High-end Gamer":{inquiries:0,reach:0,engagement:0,spent:0}}};
      weeks.push(weekMap[k]||blank);
    }
    return weeks;
  })():null;

  const renderWeeklyCell=(key,current,previous,formatter)=>{
    const c=Number(current);
    const p=Number(previous);
    const hasCurrent=Number.isFinite(c);
    const canDelta=Number.isFinite(c)&&Number.isFinite(p)&&p!==0;
    const delta=canDelta?((c-p)/Math.abs(p))*100:null;
    return(<td key={key}><div className="wk-main">{hasCurrent?formatter(c):"-"}</div><div className={`wk-delta ${delta==null?"wk-flat":delta>=0?"wk-up":"wk-down"}`}>{delta==null?"-":`${delta>=0?"+":""}${delta.toFixed(2)}%`}</div></td>);
  };
  const renderDailyValueWithDelta=(current,previous,formatter)=>{
    const c=Number(current);
    const p=Number(previous);
    const validCurrent=Number.isFinite(c);
    const canDelta=Number.isFinite(c)&&Number.isFinite(p)&&p!==0;
    const delta=canDelta?((c-p)/Math.abs(p))*100:null;
    return(<div className="daily-val"><span>{validCurrent?formatter(c):"-"}</span><span className={`daily-delta ${delta==null?"wk-flat":delta>=0?"wk-up":"wk-down"}`}>{delta==null?"-":`${delta>=0?"+":""}${delta.toFixed(2)}%`}</span></div>);
  };

  return(<div>
    <div className="info-box" style={{marginBottom:12}}>
      <div style={{fontSize:11,fontWeight:600,letterSpacing:".06em",textTransform:"uppercase",marginBottom:5}}>Source Summary</div>
      <div style={{fontSize:12,lineHeight:1.45,color:"inherit"}}>{sourceParts.join(" | ")}</div>
      <div style={{fontSize:11,lineHeight:1.4,marginTop:5,opacity:.9}}>Sync cutoff uses Asia/Manila. KPI totals are summed from Meta daily rows and may differ slightly from Ads Manager depending on attribution/reporting settings.</div>
    </div>
    <div className="g4">
      <KPI label="Total Spend" value={fmtMoneyExact(total.spend)} sub={budget>0?`${fmtPctExact(spPct)} of ${fmtMoneyExact(budget)}`:undefined} badge={budget>0?`${fmtPctExact(spPct)} used`:undefined} badgeNeutral/>
      <KPI label="Total Inquiries" value={fmtN(total.inquiries)} sub={target>0?`${fmtPctExact(inPct)} of ${fmtN(target)} target`:undefined} badge={target>0?`${fmtPctExact(inPct)} of target`:undefined} badgeUp={inPct>=80}/>
      <KPI label="Cost per Inquiry" value={fmtMoneyExact(cpi)} sub="Spend ÷ inquiries"/>
      <KPI label="CTR" value={fmtPctExact(ctr)} sub={`Link click-through rate · ${ctrBenchmark.range}`} badge={ctrBenchmark.tier} badgeUp={ctrBenchmark.tone==="up"} badgeNeutral={ctrBenchmark.tone==="neutral"} infoTitle="CTR Performance Tiers (2026 Benchmarks)" infoRows={CTR_BENCHMARK_ROWS}/>
    </div>
    <div className="g4">
      <KPI label="Total Reach" value={fmtN(total.reach)} sub={`${fmtPctExact(reachShare)} share`} color={COLORS.teal}/>
      <KPI label="Total Impression" value={fmtN(total.impressions)} sub={`${fmtPctExact(impressionShare)} share`} color={COLORS.teal}/>
      <KPI label="CPM" value={fmtMoneyExact(cpm)} sub="(Spend ÷ impressions) × 1000"/>
      <KPI label="Total Engagement" value={fmtN(total.post_engagement)} sub={`${fmtPctExact(engagementShare)} share`} color={COLORS.amber}/>
    </div>

    <div className="seg-grid" style={{display:"grid",gridTemplateColumns:`repeat(${bySeg.length},minmax(0,1fr))`,gap:11,marginBottom:14}}>
      {bySeg.map(({seg,a,series,color,segBudget,segTarget,dailyDetails})=>{
        const sCPI=a.inquiries>0?a.spend/a.inquiries:0;
        const sPct=total.inquiries>0?(a.inquiries/total.inquiries*100):0;
        return(<div className="seg-card seg-card-click" key={seg} role="button" tabIndex={0} onClick={()=>setSegmentModal({seg,color,dailyDetails})} onKeyDown={e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();setSegmentModal({seg,color,dailyDetails});}}}>
          <div className="seg-lbl">{seg}</div>
          <div className="seg-val" style={{color}}>{fmtN(a.inquiries)}</div>
          <div className="seg-sub">{fmtPctExact(sPct)} share · CPI {fmtMoneyExact(sCPI)}</div>
          {segBudget>0&&<PacingRow label="Budget" value={a.spend} max={segBudget} color={color} vfmt={fmt}/>}
          {segTarget>0&&<PacingRow label="Target" value={a.inquiries} max={segTarget} color={color}/>}
          <Spark values={series.map(([,v])=>v)} color={color}/>
          <div style={{fontSize:10,color:"rgba(226,232,240,.7)",marginTop:5}}>Spend {fmtMoneyExact(a.spend)} · Reach {fmtN(a.reach)}</div>
        </div>);
      })}
    </div>

    {segmentModal&&(<div className="seg-modal-backdrop" onClick={()=>setSegmentModal(null)}>
      <div className="seg-modal" onClick={e=>e.stopPropagation()}>
        <div className="seg-modal-h">
          <div>
            <div className="seg-modal-t">{segmentModal.seg} Daily Breakdown</div>
            <div className="seg-modal-sub">{month} · Daily spend, daily inquiries, and daily CPI</div>
          </div>
          <button className="seg-modal-close" onClick={()=>setSegmentModal(null)}>Close</button>
        </div>
        <div className="tw"><table>
          <thead><tr><th>Day</th><th>Spend</th><th>Inquiries</th><th>Cost per Inquiry</th></tr></thead>
          <tbody>{segmentModal.dailyDetails.length?segmentModal.dailyDetails.map((d,_,arr)=>{
            const maxSpend=Math.max(...arr.map(x=>Number(x.spend)||0),1);
            const maxInq=Math.max(...arr.map(x=>Number(x.inquiries)||0),1);
            const maxCpi=Math.max(...arr.map(x=>Number(x.cpi)||0),1);
            const spendW=Math.max(4,(Number(d.spend)||0)/maxSpend*100);
            const inqW=Math.max(4,(Number(d.inquiries)||0)/maxInq*100);
            const cpiW=Math.max(4,(Number(d.cpi)||0)/maxCpi*100);
            return(<tr key={d.day}>
              <td>{d.day}</td>
              <td><div className="db-cell"><div className="db-fill db-spend" style={{width:`${spendW}%`}}/><span className="db-text">{fmtMoneyExact(d.spend)}</span></div></td>
              <td><div className="db-cell"><div className="db-fill db-inq" style={{width:`${inqW}%`}}/><span className="db-text">{fmtN(d.inquiries)}</span></div></td>
              <td><div className="db-cell"><div className="db-fill db-cpi" style={{width:`${cpiW}%`}}/><span className="db-text">{d.inquiries>0?fmtMoneyExact(d.cpi):"-"}</span></div></td>
            </tr>);
          }):(<tr><td colSpan={4}>No daily rows for this segment in selected month.</td></tr>)}</tbody>
        </table></div>
      </div>
    </div>)}

    <div className="g2">
      <div className="card seg-card-click" role="button" tabIndex={0} onClick={()=>setTrendModal(true)} onKeyDown={e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();setTrendModal(true);}}}>
        <div className="card-hdr"><div className="card-ttl">Daily Inquiry Trend</div></div>
        <LineChart series={dailyAll} color={colors[0]}/>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:9.5,color:"#2D3B50",marginTop:4}}>
          <span>{dailyAll[0]?.[0]}</span><span>{dailyAll[dailyAll.length-1]?.[0]}</span>
        </div>
      </div>
      <div className="card">
        <div className="card-hdr"><div className="card-ttl">Inquiry Share by Segment</div></div>
        <div className="donut-wrap" style={{alignItems:"flex-start",gap:14}}>
          <DonutChart data={donutData.length?donutData:[{label:"No data",value:1,color:"#1A2838"}]} size={170} showLegend={false}/>
          <div className="share-list" style={{flex:1,minWidth:260}}>
            {inquiryShareRows.map(r=>(<div key={r.seg} className="share-row"><div className="share-row-top"><span className="share-name">{r.seg}</span><span className="share-num">{fmtN(r.inquiries)} · <span className="share-pct">{fmtPctExact(r.pct)}</span></span></div><div className="share-bar"><div className="share-fill" style={{width:`${Math.max(2,r.pct)}%`,background:r.color}}/></div></div>))}
          </div>
        </div>
      </div>
    </div>

    <div className="g2">
      <div className="card">
        <div className="card-hdr"><div className="card-ttl">Spend by Segment</div></div>
        <HBar data={bySeg.map(s=>({label:s.seg,value:s.a.spend,color:s.color}))} valueFormat={fmt}/>
      </div>
      <div className="card">
        <div className="card-hdr"><div className="card-ttl">Spend Distribution</div></div>
        <div className="donut-wrap" style={{alignItems:"flex-start",gap:14}}>
          <DonutChart data={donutSpend.length?donutSpend:[{label:"No data",value:1,color:"#1A2838"}]} size={170} showLegend={false}/>
          <div className="share-list" style={{flex:1,minWidth:260}}>
            {spendShareRows.map(r=>(<div key={r.seg} className="share-row"><div className="share-row-top"><span className="share-name">{r.seg}</span><span className="share-num">{fmtMoneyExact(r.spend)} · <span className="share-pct">{fmtPctExact(r.pct)}</span></span></div><div className="share-bar"><div className="share-fill" style={{width:`${Math.max(2,r.pct)}%`,background:r.color}}/></div></div>))}
          </div>
        </div>
      </div>
    </div>

    {trendModal&&(<div className="seg-modal-backdrop" onClick={()=>setTrendModal(false)}>
      <div className="seg-modal" onClick={e=>e.stopPropagation()}>
        <div className="seg-modal-h">
          <div>
            <div className="seg-modal-t">Daily Inquiry Trend — Full Month</div>
            <div className="seg-modal-sub">{month} · Includes complete daily labels and dates</div>
          </div>
          <button className="seg-modal-close" onClick={()=>setTrendModal(false)}>Close</button>
        </div>
        <div className="card" style={{padding:10,marginBottom:10}}>
          <LineChart series={dailyAll} color={colors[0]} showValueLabels/>
        </div>
        <div className="tw"><table>
          <thead><tr><th>Date</th><th>Daily Inquiries</th><th>Daily Spent</th><th>Cost per Inquiry</th></tr></thead>
          <tbody>{dailyRows.length?dailyRows.map((r,i)=>{const p=i>0?dailyRows[i-1]:null;return(<tr key={r.day}><td>{r.day}</td><td>{renderDailyValueWithDelta(r.inquiries,p?.inquiries,fmtN)}</td><td>{renderDailyValueWithDelta(r.spend,p?.spend,fmtMoneyExact)}</td><td>{r.inquiries>0?renderDailyValueWithDelta(r.cpi,p?.cpi,fmtMoneyExact):"-"}</td></tr>);}):<tr><td colSpan={4}>No daily rows in selected month.</td></tr>}</tbody>
        </table></div>
      </div>
    </div>)}

    {budget>0&&<div className="card">
      <div className="card-hdr"><div className="card-ttl">Overall Pacing</div></div>
      <PacingRow label="Budget utilization" value={total.spend} max={budget} color={spPct>90?COLORS.red:spPct>70?COLORS.amber:COLORS.blue} vfmt={fmt}/>
      {target>0&&<PacingRow label="Inquiry target" value={total.inquiries} max={target} color={inPct>90?COLORS.green:inPct>60?COLORS.teal:COLORS.purple}/>}
    </div>}

    <div className="card">
      <div className="card-hdr"><div className="card-ttl">Segment Performance Table</div></div>
      <div className="tw"><table>
        <thead><tr><th>Segment</th><th>Spend</th><th>Inquiries</th><th>CPI</th><th>Reach</th><th>Impressions</th><th>CPM</th><th>Engagement</th><th>CTR</th></tr></thead>
        <tbody>{bySeg.map(({seg,a,color})=>{
          const cpi=a.inquiries>0?a.spend/a.inquiries:0,cpm=a.impressions>0?(a.spend/a.impressions)*1000:0,ctr=a.impressions>0?(a.clicks/a.impressions)*100:0;
          return(<tr key={seg}><td><span className={`pill ${PM[seg]||"pl-gray"}`} style={{color,borderColor:color+"44"}}>{seg}</span></td><td className="td-p">{fmtMoneyExact(a.spend)}</td><td>{fmtN(a.inquiries)}</td><td>{fmtMoneyExact(cpi)}</td><td>{fmtN(a.reach)}</td><td>{fmtN(a.impressions)}</td><td>{fmtMoneyExact(cpm)}</td><td>{fmtN(a.post_engagement)}</td><td>{fmtPctExact(ctr)}</td></tr>);
        })}</tbody>
      </table></div>
    </div>

    {weeklyMonitoring&&(<div className="card">
      <div className="card-hdr"><div className="card-ttl">Weekly Monitoring</div></div>
      <div className="weekly-note">Independent from month selector. Weeks run Monday to Sunday using week ranges</div>
      <div className="tw weekly-wrap"><table className="weekly-table">
        <thead>
          <tr>
            <th className="weekly-sticky">Metric</th>
            {weeklyMonitoring.map(w=><th key={w.key}><div>Week {w.weekNo}</div><div style={{fontSize:10,fontWeight:400,color:"#64748B"}}>{w.startIso} to {w.endIso}</div></th>)}
          </tr>
        </thead>
        <tbody>
          <tr className="weekly-sec"><td className="weekly-sticky weekly-section-title">Total</td>{weeklyMonitoring.map(w=><td key={`sec-t-${w.key}`}/>)}</tr>
          <tr><td className="weekly-sticky weekly-metric">Total Inquiries</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`ti-${w.key}`,w.total.inquiries,i>0?weeklyMonitoring[i-1].total.inquiries:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Total Reach</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`tr-${w.key}`,w.total.reach,i>0?weeklyMonitoring[i-1].total.reach:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Total Engagement</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`te-${w.key}`,w.total.engagement,i>0?weeklyMonitoring[i-1].total.engagement:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Total Spent</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`ts-${w.key}`,w.total.spent,i>0?weeklyMonitoring[i-1].total.spent:NaN,fmtMoneyExact))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Cost Per Inquiry</td>{weeklyMonitoring.map((w,i)=>{const curr=w.total.inquiries>0?w.total.spent/w.total.inquiries:NaN;const prev=i>0&&weeklyMonitoring[i-1].total.inquiries>0?weeklyMonitoring[i-1].total.spent/weeklyMonitoring[i-1].total.inquiries:NaN;return renderWeeklyCell(`tci-${w.key}`,curr,prev,fmtMoneyExact);})}</tr>
          <tr><td className="weekly-sticky weekly-metric">Cost Per Engagement</td>{weeklyMonitoring.map((w,i)=>{const curr=w.total.engagement>0?w.total.spent/w.total.engagement:NaN;const prev=i>0&&weeklyMonitoring[i-1].total.engagement>0?weeklyMonitoring[i-1].total.spent/weeklyMonitoring[i-1].total.engagement:NaN;return renderWeeklyCell(`tce-${w.key}`,curr,prev,fmtMoneyExact);})}</tr>

          <tr className="weekly-sec"><td className="weekly-sticky weekly-section-title">Productivity</td>{weeklyMonitoring.map(w=><td key={`sec-p-${w.key}`}/>)}</tr>
          <tr><td className="weekly-sticky weekly-metric">Inquiry</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`pi-${w.key}`,w.segments["Productivity"].inquiries,i>0?weeklyMonitoring[i-1].segments["Productivity"].inquiries:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Reach</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`pr-${w.key}`,w.segments["Productivity"].reach,i>0?weeklyMonitoring[i-1].segments["Productivity"].reach:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Engagement</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`pe-${w.key}`,w.segments["Productivity"].engagement,i>0?weeklyMonitoring[i-1].segments["Productivity"].engagement:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Spent</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`ps-${w.key}`,w.segments["Productivity"].spent,i>0?weeklyMonitoring[i-1].segments["Productivity"].spent:NaN,fmtMoneyExact))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Cost Per Inquiry</td>{weeklyMonitoring.map((w,i)=>{const s=w.segments["Productivity"];const p=i>0?weeklyMonitoring[i-1].segments["Productivity"]:null;const curr=s.inquiries>0?s.spent/s.inquiries:NaN;const prev=p&&p.inquiries>0?p.spent/p.inquiries:NaN;return renderWeeklyCell(`pci-${w.key}`,curr,prev,fmtMoneyExact);})}</tr>

          <tr className="weekly-sec"><td className="weekly-sticky weekly-section-title">Gaming</td>{weeklyMonitoring.map(w=><td key={`sec-g-${w.key}`}/>)}</tr>
          <tr><td className="weekly-sticky weekly-metric">Inquiry</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`gi-${w.key}`,w.segments["Gaming"].inquiries,i>0?weeklyMonitoring[i-1].segments["Gaming"].inquiries:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Reach</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`gr-${w.key}`,w.segments["Gaming"].reach,i>0?weeklyMonitoring[i-1].segments["Gaming"].reach:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Engagement</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`ge-${w.key}`,w.segments["Gaming"].engagement,i>0?weeklyMonitoring[i-1].segments["Gaming"].engagement:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Spent</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`gs-${w.key}`,w.segments["Gaming"].spent,i>0?weeklyMonitoring[i-1].segments["Gaming"].spent:NaN,fmtMoneyExact))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Cost Per Inquiry</td>{weeklyMonitoring.map((w,i)=>{const s=w.segments["Gaming"];const p=i>0?weeklyMonitoring[i-1].segments["Gaming"]:null;const curr=s.inquiries>0?s.spent/s.inquiries:NaN;const prev=p&&p.inquiries>0?p.spent/p.inquiries:NaN;return renderWeeklyCell(`gci-${w.key}`,curr,prev,fmtMoneyExact);})}</tr>

          <tr className="weekly-sec"><td className="weekly-sticky weekly-section-title">High-end Gamer</td>{weeklyMonitoring.map(w=><td key={`sec-h-${w.key}`}/>)}</tr>
          <tr><td className="weekly-sticky weekly-metric">Inquiry</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`hi-${w.key}`,w.segments["High-end Gamer"].inquiries,i>0?weeklyMonitoring[i-1].segments["High-end Gamer"].inquiries:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Reach</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`hr-${w.key}`,w.segments["High-end Gamer"].reach,i>0?weeklyMonitoring[i-1].segments["High-end Gamer"].reach:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Engagement</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`he-${w.key}`,w.segments["High-end Gamer"].engagement,i>0?weeklyMonitoring[i-1].segments["High-end Gamer"].engagement:NaN,fmtN))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Spent</td>{weeklyMonitoring.map((w,i)=>renderWeeklyCell(`hs-${w.key}`,w.segments["High-end Gamer"].spent,i>0?weeklyMonitoring[i-1].segments["High-end Gamer"].spent:NaN,fmtMoneyExact))}</tr>
          <tr><td className="weekly-sticky weekly-metric">Cost Per Inquiry</td>{weeklyMonitoring.map((w,i)=>{const s=w.segments["High-end Gamer"];const p=i>0?weeklyMonitoring[i-1].segments["High-end Gamer"]:null;const curr=s.inquiries>0?s.spent/s.inquiries:NaN;const prev=p&&p.inquiries>0?p.spent/p.inquiries:NaN;return renderWeeklyCell(`hci-${w.key}`,curr,prev,fmtMoneyExact);})}</tr>
        </tbody>
      </table></div>
    </div>)}
  </div>);
}

// Overview
function Overview({data,month,settings}){
  const budget=Number(settings.budgets?.[month]||0),target=Number(settings.targets?.[month]||0);
  const total=aggRows(data);
  const cpi=total.inquiries>0?total.spend/total.inquiries:0,cpm=total.impressions>0?(total.spend/total.impressions)*1000:0,ctr=total.impressions>0?(total.clicks/total.impressions)*100:0;
  const ctrBenchmark=getCtrBenchmark(ctr);
  const spPct=budget>0?(total.spend/budget)*100:0,inPct=target>0?(total.inquiries/target)*100:0;
  const byLob=groupRows(data,"lob").filter(g=>g.inquiries>0);
  const dailyAll=dailySeries(data).slice(-14);
  const accounts=[...new Set(data.map(r=>r.account_name).filter(Boolean))];

  return(<div>
    {accounts.length>1&&<div className="info-box">Aggregated across <strong>{accounts.length}</strong> accounts: {accounts.join(", ")}</div>}
    <div className="g4">
      <KPI label="Total Spend" value={fmtK(total.spend)} sub={budget>0?`${spPct.toFixed(1)}% of ${fmt(budget)}`:undefined} badge={budget>0?`${spPct.toFixed(0)}% used`:undefined} badgeNeutral/>
      <KPI label="Total Inquiries" value={fmtN(total.inquiries)} sub={target>0?`${inPct.toFixed(1)}% of ${fmtN(target)} target`:undefined} badge={target>0?`${inPct.toFixed(0)}% of target`:undefined} badgeUp={inPct>=80}/>
      <KPI label="Cost per Inquiry" value={fmt(cpi)} sub="Overall avg. CPI"/>
      <KPI label="Avg. CTR" value={fmtP(ctr)} sub={`Link click-through rate · ${ctrBenchmark.range}`} badge={ctrBenchmark.tier} badgeUp={ctrBenchmark.tone==="up"} badgeNeutral={ctrBenchmark.tone==="neutral"} infoTitle="CTR Performance Tiers (2026 Benchmarks)" infoRows={CTR_BENCHMARK_ROWS}/>
    </div>
    <div className="g4">
      <KPI label="Total Reach" value={fmtN(total.reach)} color={COLORS.teal}/>
      <KPI label="Impressions" value={fmtN(total.impressions)} color={COLORS.teal}/>
      <KPI label="CPM" value={fmt(cpm)}/>
      <KPI label="Engagement" value={fmtN(total.post_engagement)} color={COLORS.amber}/>
    </div>
    <div className="g2">
      <div className="card">
        <div className="card-hdr"><div className="card-ttl">Daily Inquiry Trend — Last 14 Days</div></div>
        <LineChart series={dailyAll} color={COLORS.blue}/>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:9.5,color:"#2D3B50",marginTop:4}}><span>{dailyAll[0]?.[0]}</span><span>{dailyAll[dailyAll.length-1]?.[0]}</span></div>
      </div>
      <div className="card">
        <div className="card-hdr"><div className="card-ttl">Budget & Target Pacing</div></div>
        {budget>0&&<PacingRow label="Budget utilization" value={total.spend} max={budget} color={spPct>90?COLORS.red:spPct>70?COLORS.amber:COLORS.blue} vfmt={fmt}/>}
        {target>0&&<PacingRow label="Inquiry target" value={total.inquiries} max={target} color={inPct>90?COLORS.green:inPct>60?COLORS.teal:COLORS.purple}/>}
        {accounts.length>0&&<div style={{marginTop:10,paddingTop:10,borderTop:"1px solid #111A24"}}><div style={{fontSize:9,color:"#2D3B50",marginBottom:6,textTransform:"uppercase",letterSpacing:".1em",fontWeight:500}}>Active Accounts</div>{accounts.map(a=>(<div key={a} style={{fontSize:11,color:"#374151",marginBottom:3,display:"flex",alignItems:"center",gap:5}}><span style={{width:4,height:4,borderRadius:"50%",background:COLORS.teal,display:"inline-block"}}/>{a}</div>))}</div>}
      </div>
    </div>
    <div className="g2">
      <div className="card"><div className="card-hdr"><div className="card-ttl">Inquiry Share by LOB</div></div><DonutChart data={byLob.map((g,i)=>({label:g.key,value:g.inquiries,color:CHART_COLORS[i]}))} size={105}/></div>
      <div className="card"><div className="card-hdr"><div className="card-ttl">Spend Distribution by LOB</div></div><DonutChart data={byLob.map((g,i)=>({label:g.key,value:Math.round(g.spend),color:CHART_COLORS[i]}))} size={105}/></div>
    </div>
    <div className="card">
      <div className="card-hdr"><div className="card-ttl">LOB Performance Summary</div></div>
      <div className="tw"><table>
        <thead><tr><th>LOB</th><th>Spend</th><th>Inquiries</th><th>CPI</th><th>Reach</th><th>Impressions</th><th>CPM</th><th>CTR</th></tr></thead>
        <tbody>{byLob.map(g=>{const cpi=g.inquiries>0?g.spend/g.inquiries:0,cpm=g.impressions>0?(g.spend/g.impressions)*1000:0,ctr=g.impressions>0?(g.clicks/g.impressions)*100:0;return(<tr key={g.key}><td><span className={`pill ${PM[g.key]||"pl-gray"}`}>{g.key}</span></td><td className="td-p">{fmt(g.spend)}</td><td>{fmtN(g.inquiries)}</td><td>{fmt(cpi)}</td><td>{fmtN(g.reach)}</td><td>{fmtN(g.impressions)}</td><td>{fmt(cpm)}</td><td>{fmtP(ctr)}</td></tr>);})}</tbody>
      </table></div>
    </div>
  </div>);
}

// Breakdown
function Breakdown({data}){
  const [gk,setGk]=useState("lob"),[metric,setMetric]=useState("inquiries");
  const grouped=useMemo(()=>groupRows(data,gk),[data,gk]);
  const mFmt={inquiries:fmtN,spend:fmt,reach:fmtN,impressions:fmtN};
  const mLabel={inquiries:"Inquiries",spend:"Spend",reach:"Reach",impressions:"Impressions"};
  return(<div>
    <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}}>
      <div className="tabs">{[["lob","LOB"],["div","Division"],["seg","Segment"],["obj","Objective"]].map(([k,l])=>(<button key={k} className={`tab ${gk===k?"active":""}`} onClick={()=>setGk(k)}>{l}</button>))}</div>
      <div className="tabs" style={{marginLeft:"auto"}}>{Object.entries(mLabel).map(([k,l])=>(<button key={k} className={`tab ${metric===k?"active":""}`} onClick={()=>setMetric(k)}>{l}</button>))}</div>
    </div>
    <div className="g2">
      <div className="card"><div className="card-hdr"><div className="card-ttl">Inquiries by {gk==="lob"?"LOB":gk==="div"?"Division":gk==="seg"?"Segment":"Objective"}</div></div><HBar data={[...grouped].sort((a,b)=>b.inquiries-a.inquiries).map((g,i)=>({label:g.key,value:g.inquiries,color:CHART_COLORS[i%CHART_COLORS.length]}))} valueFormat={fmtN}/></div>
      <div className="card"><div className="card-hdr"><div className="card-ttl">Spend by {gk==="lob"?"LOB":gk==="div"?"Division":gk==="seg"?"Segment":"Objective"}</div></div><HBar data={grouped.map((g,i)=>({label:g.key,value:g.spend,color:CHART_COLORS[i%CHART_COLORS.length]}))} valueFormat={fmt}/></div>
    </div>
    <div className="g2">
      <div className="card"><div className="card-hdr"><div className="card-ttl">Inquiry Distribution</div></div><DonutChart data={grouped.filter(g=>g.inquiries>0).map((g,i)=>({label:g.key,value:g.inquiries,color:CHART_COLORS[i%CHART_COLORS.length]}))} size={105}/></div>
      <div className="card"><div className="card-hdr"><div className="card-ttl">Spend Distribution</div></div><DonutChart data={grouped.filter(g=>g.spend>0).map((g,i)=>({label:g.key,value:Math.round(g.spend),color:CHART_COLORS[i%CHART_COLORS.length]}))} size={105}/></div>
    </div>
    <div className="card"><div className="card-hdr"><div className="card-ttl">Full Performance Table</div></div>
      <div className="tw"><table>
        <thead><tr><th>Name</th><th>Spend</th><th>Inquiries</th><th>CPI</th><th>Reach</th><th>Impressions</th><th>CPM</th><th>Engagement</th><th>CTR</th></tr></thead>
        <tbody>{grouped.map(g=>{const cpi=g.inquiries>0?g.spend/g.inquiries:0,cpm=g.impressions>0?(g.spend/g.impressions)*1000:0,ctr=g.impressions>0?(g.clicks/g.impressions)*100:0;return(<tr key={g.key}><td><span className={`pill ${PM[g.key]||"pl-gray"}`}>{g.key}</span></td><td className="td-p">{fmt(g.spend)}</td><td>{fmtN(g.inquiries)}</td><td>{fmt(cpi)}</td><td>{fmtN(g.reach)}</td><td>{fmtN(g.impressions)}</td><td>{fmt(cpm)}</td><td>{fmtN(g.post_engagement)}</td><td>{fmtP(ctr)}</td></tr>);})}</tbody>
      </table></div>
    </div>
  </div>);
}

function Trends({rawData,identifiers}){
  const [metric,setMetric]=useState("inquiries");
  const allData=useMemo(()=>applyIdentifiers(rawData,identifiers),[rawData,identifiers]);
  const byML=useMemo(()=>{
    const map={};
    allData.forEach(r=>{
      const d=new Date(r.day+"T00:00:00"),mk=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`,lob=r._meta?.lob||"Unknown",k=`${mk}|${lob}`;
      if(!map[k])map[k]={month:mk,lob,spend:0,inquiries:0,reach:0,impressions:0};
      map[k].spend+=r.spend;map[k].inquiries+=r.inquiries;map[k].reach+=r.reach;map[k].impressions+=r.impressions;
    });
    return Object.values(map);
  },[allData]);
  const months=[...new Set(byML.map(r=>r.month))].sort(),lobs=[...new Set(byML.map(r=>r.lob))];
  const byM={};months.forEach(m=>{byM[m]={spend:0,inquiries:0,reach:0,impressions:0};byML.filter(r=>r.month===m).forEach(r=>{byM[m].spend+=r.spend;byM[m].inquiries+=r.inquiries;byM[m].reach+=r.reach;byM[m].impressions+=r.impressions;});});
  const mFmt={inquiries:fmtN,spend:fmt,reach:fmtN,impressions:fmtN},mLabel={inquiries:"Inquiries",spend:"Spend (₱)",reach:"Reach",impressions:"Impressions"};
  const ml=m=>{const[y,mo]=m.split("-");return new Date(+y,+mo-1,1).toLocaleString("en-PH",{month:"short",year:"2-digit"});};
  const monthSeries=months.map(m=>[ml(m),byM[m]?.[metric]||0]);
  return(<div>
    <div className="tabs" style={{marginBottom:14}}>{Object.entries(mLabel).map(([k,l])=>(<button key={k} className={`tab ${metric===k?"active":""}`} onClick={()=>setMetric(k)}>{l}</button>))}</div>
    <div className="card"><div className="card-hdr"><div className="card-ttl">Monthly {mLabel[metric]} Trend</div></div>
      <LineChart series={monthSeries} color={COLORS.blue}/>
      <div style={{display:"flex",justifyContent:"space-between",fontSize:9.5,color:"#2D3B50",marginTop:4}}>{monthSeries.map(([l])=><span key={l}>{l}</span>)}</div>
    </div>
    <div className="g2">
      <div className="card"><div className="card-hdr"><div className="card-ttl">Monthly {mLabel[metric]}</div></div><HBar data={months.map(m=>({label:ml(m),value:byM[m]?.[metric]||0,color:COLORS.blue}))} valueFormat={mFmt[metric]}/></div>
      <div className="card"><div className="card-hdr"><div className="card-ttl">{mLabel[metric]} by LOB — All Time</div></div><HBar data={lobs.map((lob,i)=>({label:lob,value:byML.filter(r=>r.lob===lob).reduce((s,r)=>s+r[metric],0),color:CHART_COLORS[i%CHART_COLORS.length]})).sort((a,b)=>b.value-a.value)} valueFormat={mFmt[metric]}/></div>
    </div>
    <div className="card"><div className="card-hdr"><div className="card-ttl">LOB × Month Cross-Table</div></div>
      <div className="tw"><table>
        <thead><tr><th>LOB</th>{months.map(m=><th key={m}>{ml(m)}</th>)}<th>Total</th></tr></thead>
        <tbody>{lobs.map(lob=>{const vals=months.map(m=>byML.find(r=>r.month===m&&r.lob===lob)?.[metric]||0),total=vals.reduce((s,v)=>s+v,0);return(<tr key={lob}><td className="td-p">{lob}</td>{vals.map((v,i)=><td key={i}>{mFmt[metric](v)}</td>)}<td style={{color:COLORS.blue,fontWeight:600,fontFamily:"'DM Mono',monospace"}}>{mFmt[metric](total)}</td></tr>);})}</tbody>
      </table></div>
    </div>
  </div>);
}

const NAV=[
  {id:"overview",  label:"Monthly Overview",  icon:"M3 12l2-2m0 0l7-7 7 7M5 10v10a1 1 0 001 1h3m10-11l2 2m-2-2v10a1 1 0 01-1 1h-3m-6 0a1 1 0 001-1v-4a1 1 0 011-1h2a1 1 0 011 1v4a1 1 0 001 1m-6 0h6"},
  {id:"breakdown", label:"Monthly Breakdown", icon:"M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"},
  {id:"desktop",   label:"Desktop Dashboard", icon:"M9.75 17L9 20l-1 1h8l-1-1-.75-3M3 13h18M5 17h14a2 2 0 002-2V5a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"},
  {id:"lts",       label:"LTS Dashboard",     icon:"M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4"},
  {id:"lsa",       label:"LSA Dashboard",     icon:"M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z M15 11a3 3 0 11-6 0 3 3 0 016 0z"},
  {id:"trends",    label:"Trends & Graphs",   icon:"M7 12l3-3 3 3 4-4M8 21l4-4 4 4M3 4h18M4 4h16v12a1 1 0 01-1 1H5a1 1 0 01-1-1V4z"},
];
const PT={overview:"Monthly Overview",breakdown:"Monthly Breakdown",desktop:"Desktop Dashboard",lts:"LTS Dashboard",lsa:"LSA Dashboard",trends:"Trends & Graphs"};

export default function App(){
  const saved=loadLS();
  const [page,setPage]=useState("overview");
  const [navCollapsed,setNavCollapsed]=useState(false);
  const [theme,setTheme]=useState(loadTheme);
  const [settings,setSettings]=useState(()=>({...DEFAULT_SETTINGS,...(saved||{})}));
  const [identifiers,setIdentifiers]=useState([]);
  const [month,setMonth]=useState(()=>saved?.defaultMonth||"2026 MARCH");
  const [rawData,setRawData]=useState([]);
  const [loading,setLoading]=useState(false);
  const [syncState,setSyncState]=useState(null);
  const [syncMsg,setSyncMsg]=useState("");
  const [manifestVersion,setManifestVersion]=useState(null);
  const [orgAuthBusy,setOrgAuthBusy]=useState(false);
  const [orgAuthReady,setOrgAuthReady]=useState(false);
  const [orgAuthError,setOrgAuthError]=useState("");
  const [orgAccount,setOrgAccount]=useState(null);
  const [toast,setToast]=useState(null);
  const showToast=(msg,isErr=false)=>{setToast({msg,isErr});setTimeout(()=>setToast(null),3500);};
  const allowedOrgEmails=useMemo(()=>new Set(AAD_ALLOWED_EMAILS),[]);
  const msalClient=useMemo(()=>{
    if(!AAD_TENANT_ID||!AAD_CLIENT_ID)return null;
    return new PublicClientApplication({
      auth:{clientId:AAD_CLIENT_ID,authority:`https://login.microsoftonline.com/${AAD_TENANT_ID}`,redirectUri:window.location.origin},
      cache:{cacheLocation:"sessionStorage"},
    });
  },[]);
  const isAllowedOrgEmail=useCallback((email)=>{
    const value=String(email||"").trim().toLowerCase();
    if(!value)return false;
    if(!allowedOrgEmails.size)return true;
    return allowedOrgEmails.has(value);
  },[allowedOrgEmails]);
  const completeOrgAuth=useCallback(async(account)=>{
    if(!msalClient||!account)throw new Error("No Microsoft account selected");
    const resp=await msalClient.acquireTokenSilent({account,scopes:["openid","profile","email"]});
    if(!isAllowedOrgEmail(account.username||""))throw new Error("This Microsoft account is not in the allowed list for this app.");
    setApiAuthToken(resp.idToken||"");
    setOrgAccount(account);
    setOrgAuthError("");
  },[msalClient,isAllowedOrgEmail]);
  const signInWithMicrosoft=useCallback(async()=>{
    if(!msalClient){setOrgAuthError("Microsoft login is not configured.");return;}
    setOrgAuthBusy(true);setOrgAuthError("");
    try{
      const login=await msalClient.loginPopup({scopes:["openid","profile","email"],prompt:"select_account"});
      await completeOrgAuth(login.account);
      showToast("Signed in with Microsoft organization account ✓");
    }catch(err){
      setApiAuthToken("");setOrgAccount(null);
      setOrgAuthError(err?.message||"Microsoft sign-in failed.");
    }finally{setOrgAuthBusy(false);setOrgAuthReady(true);}
  },[msalClient,completeOrgAuth]);
  const signOutMicrosoft=useCallback(async()=>{
    if(!msalClient)return;
    setOrgAuthBusy(true);
    try{await msalClient.logoutPopup({account:orgAccount||undefined,postLogoutRedirectUri:window.location.origin});}catch{}
    setApiAuthToken("");setOrgAccount(null);setOrgAuthBusy(false);
    showToast("Signed out from Microsoft account");
  },[msalClient,orgAccount]);
  useEffect(()=>{
    document.body.classList.toggle("theme-light",theme==="light");
    saveTheme(theme);
  },[theme]);
  useEffect(()=>{
    let cancelled=false;
    const boot=async()=>{
      if(LOCAL_ONLY){setOrgAuthReady(true);setOrgAccount({username:"local@localhost"});return;}
      if(!msalClient){setOrgAuthError("Microsoft login is not configured.");setOrgAuthReady(true);return;}
      setOrgAuthBusy(true);
      try{
        if(typeof msalClient.initialize==="function")await msalClient.initialize();
        await msalClient.handleRedirectPromise().catch(()=>null);
        const current=msalClient.getActiveAccount()||msalClient.getAllAccounts()[0]||null;
        if(current){await completeOrgAuth(current);}else{setApiAuthToken("");setOrgAccount(null);}
      }catch(err){
        if(!cancelled){setApiAuthToken("");setOrgAccount(null);setOrgAuthError(err?.message||"Microsoft session check failed.");}
      }finally{
        if(!cancelled){setOrgAuthBusy(false);setOrgAuthReady(true);}
      }
    };
    boot();
    return()=>{cancelled=true;};
  },[msalClient,completeOrgAuth]);
  useEffect(()=>{
    if(!orgAccount)return;
    const run=async()=>{
      setLoading(true);setSyncState("checking");setSyncMsg("Checking for updates...");
      const applyStoredData=(cached)=>{
        // meta-budget-targets.json may be {"budgetTargets":{...}} or a flat object
        const bt=cached.budgets?.budgetTargets||cached.budgets||{};
        const mappings=cached.mappings||[];
        const selections=cached.selections||{};
        if(mappings.length)setIdentifiers(mappings);
        setSettings(prev=>({
          ...prev,
          ...(bt.budgets?{budgets:bt.budgets}:{}),
          ...(bt.targets?{targets:bt.targets}:{}),
          ...(bt.desktopBudgets?{desktopBudgets:bt.desktopBudgets}:{}),
          ...(bt.desktopTargets?{desktopTargets:bt.desktopTargets}:{}),
          ...(bt.ltsBudgets?{ltsBudgets:bt.ltsBudgets}:{}),
          ...(bt.ltsTargets?{ltsTargets:bt.ltsTargets}:{}),
          ...(bt.lsaBudgets?{lsaBudgets:bt.lsaBudgets}:{}),
          ...(bt.lsaTargets?{lsaTargets:bt.lsaTargets}:{}),
          mappingOptions:selections.mappingOptions||DEFAULT_MAPPING_OPTIONS,
        }));
      };
      try{
        const cached=await loadAllData();
        if(cached.rows.length){setRawData(cached.rows);applyStoredData(cached);}
        setSyncState("syncing");setSyncMsg("Syncing data from cloud...");
        const result=await syncIfNeeded();
        setManifestVersion(result.version);
        if(result.updated){
          const fresh=await loadAllData();
          setRawData(fresh.rows);
          applyStoredData(fresh);
          showToast(`Data updated to v${result.version} ✓`);
        }
        setSyncState("done");setSyncMsg("");
      }catch(e){
        setSyncState("error");setSyncMsg(e.message||"Sync failed");
        showToast(`Failed to load data: ${e.message}`,true);
      }finally{setLoading(false);}
    };
    run();
  },[orgAccount]);
  const allMappedData=useMemo(()=>applyIdentifiers(rawData,identifiers),[rawData,identifiers]);
  const monthData=useMemo(()=>filterMonth(allMappedData,month),[allMappedData,month]);
  if(!orgAuthReady){
    if(LOCAL_ONLY)return null;
    return(<><style>{css}</style><div className="org-auth-wrap"><div className="org-auth-card"><div className="org-auth-kicker">EasyPC Marketing Hub</div><div className="org-auth-title">Checking organization sign-in...</div><div className="org-auth-sub">Please wait while we verify your Microsoft 365 session.</div></div></div></>);
  }
  if(!orgAccount){
    if(LOCAL_ONLY)return null;
    return(<>
      <style>{css}</style>
      <div className="org-auth-wrap">
        <div className="org-auth-card">
          <div className="org-auth-kicker">Organization Access Only</div>
          <div className="org-auth-title">Sign in with your SharePoint work account</div>
          <div className="org-auth-sub">Access is restricted to your Microsoft Entra tenant. Use your organization Microsoft 365 account to continue.</div>
          <div className="org-auth-meta">Tenant ID: {AAD_TENANT_ID}</div>
          {allowedOrgEmails.size>0&&<div className="org-auth-meta" style={{marginTop:8}}>Email allowlist is enabled for this deployment.</div>}
          {orgAuthError&&<div className="org-auth-err">{orgAuthError}</div>}
          <div className="org-auth-actions"><button className="btn btn-p" onClick={signInWithMicrosoft} disabled={orgAuthBusy}>{orgAuthBusy?"Signing in...":"Sign in with Microsoft"}</button></div>
        </div>
      </div>
    </>);
  }
  return(<>
    <style>{css}</style>
    <div className="app">
      <aside className={`sidebar ${navCollapsed?"collapsed":""}`}>
        <div className="logo-wrap"><div><div className="logo-mark">EasyPC</div><div className="logo-name">Marketing Hub</div></div><button className="collapse-btn" onClick={()=>setNavCollapsed(v=>!v)} title={navCollapsed?"Expand navigation":"Collapse navigation"}>{navCollapsed?">":"<"}</button></div>
        <nav className="nav-sec">
          <div className="nav-lbl">Dashboards</div>
          <div className="nav-scroll">{NAV.map(n=>(<button key={n.id} title={n.label} className={`nav-item ${page===n.id?"active":""}`} onClick={()=>setPage(n.id)}><svg fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}><path strokeLinecap="round" strokeLinejoin="round" d={n.icon}/></svg><span className="nav-text">{n.label}</span></button>))}</div>
        </nav>
        <div className="sb-bot">
          <div className="sb-meta">
            <div className="org-user-chip" title={LOCAL_ONLY?"Local mode":(orgAccount?.username||"")}>{LOCAL_ONLY?"Local Mode":(orgAccount?.username||"Signed in")}</div>
            {manifestVersion!==null&&<div style={{fontSize:11,color:"#60A5FA"}}>Data v{manifestVersion}</div>}
            {(syncState==="checking"||syncState==="syncing")&&<div style={{fontSize:11,color:"#F59E0B",display:"flex",alignItems:"center",gap:5}}><span className="spin"/>{syncMsg}</div>}
            {syncState==="error"&&<div style={{fontSize:11,color:"#F87171"}}>{syncMsg}</div>}
          </div>
          {!LOCAL_ONLY&&<button className="btn" onClick={signOutMicrosoft} disabled={orgAuthBusy} style={{marginTop:7}}>{orgAuthBusy?"Signing out...":"Sign out"}</button>}
          <button className={`theme-toggle ${theme==="light"?"light":""}`} onClick={()=>setTheme(t=>t==="dark"?"light":"dark")} title={theme==="dark"?"Switch to light mode":"Switch to dark mode"}>
            <span className="theme-pill" aria-hidden="true"><span className="ico sun">☀</span><span className="ico moon">☾</span><span className="theme-thumb"/></span>
            <span className="theme-toggle-label">{theme==="dark"?"Dark":"Light"}</span>
          </button>
        </div>
      </aside>
      <div className="main">
        <div className="topbar">
          <div><div className="pg-title">{PT[page]}</div><div className="pg-sub">EasyPC Marketing Analytics · {month}</div></div>
          {["overview","breakdown","desktop","lts","lsa"].includes(page)&&(
            <div className="topbar-r">
              <select value={month} onChange={e=>setMonth(e.target.value)} style={{width:"auto",padding:"5px 10px",fontSize:14}}>{MONTH_OPTIONS.map(m=><option key={m}>{m}</option>)}</select>
            </div>
          )}
        </div>
        {loading&&syncState==="syncing"&&<div><div className="fetch-bar"/><div style={{padding:"7px 22px",fontSize:12,color:"#60A5FA",background:"#0B1220",borderBottom:"1px solid #1F2937"}}>Syncing data from cloud...</div></div>}
        <div className="scroll-area">
          <div key={page} className="page-switch">
            {page==="overview"  &&<Overview data={monthData} month={month} settings={settings}/>}
            {page==="breakdown" &&<Breakdown data={monthData}/>}
            {page==="desktop"   &&<LobDash data={monthData} allData={allMappedData} lob="Desktop" segments={DESKTOP_DASH_SEGMENTS} colors={[COLORS.blue,COLORS.indigo,COLORS.purple,COLORS.coral,COLORS.teal]} settings={settings} month={month} bKey="desktop" tKey="desktop" allowedDivisions={["Retail"]} allowedLobs={["Desktop"]} allowedSegments={DESKTOP_DASH_SEGMENTS} allowedObjectives={["Inquiry","Engagement"]} monthlyBudgetMap={settings.desktopBudgets} monthlyTargetMap={settings.desktopTargets} budgetKeys={DESKTOP_DASH_SEGMENTS} targetKeys={DESKTOP_INQUIRY_TARGET_KEYS}/>}
            {page==="lts"       &&<LobDash data={monthData} lob="LTS" segments={LTS_DASH_SEGMENTS} colors={[COLORS.teal,COLORS.cyan]} settings={settings} month={month} bKey="lts" tKey="lts" allowedDivisions={["Retail"]} allowedLobs={["LTS"]} allowedSegments={LTS_DASH_SEGMENTS} allowedObjectives={["Inquiry","Engagement"]} monthlyBudgetMap={settings.ltsBudgets} monthlyTargetMap={settings.ltsTargets} budgetKeys={LTS_ALL_BUDGET_KEYS} targetKeys={LTS_DASH_SEGMENTS}/>}
            {page==="lsa"       &&<LobDash data={monthData} lob="LSA" segments={LSA_DASH_SEGMENTS} colors={CHART_COLORS} settings={settings} month={month} bKey="lsa" tKey="lsa" allowedDivisions={["Retail"]} allowedLobs={["LSA"]} allowedSegments={LSA_DASH_SEGMENTS} allowedObjectives={["Inquiry"]} monthlyBudgetMap={settings.lsaBudgets} monthlyTargetMap={settings.lsaTargets} budgetKeys={LSA_DASH_SEGMENTS} targetKeys={LSA_DASH_SEGMENTS}/>}
            {page==="trends"    &&<Trends rawData={rawData} identifiers={identifiers}/>}
          </div>
        </div>
      </div>
    </div>
    {toast&&<div className={`toast ${toast.isErr?"ter":"tok"}`}>{toast.msg}</div>}
  </>);
}
