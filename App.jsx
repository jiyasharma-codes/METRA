import React, { useState, useEffect, useRef, useMemo } from "react";
import { jsPDF } from "jspdf";
import { QRCodeSVG } from "qrcode.react";
import { renderToStaticMarkup } from "react-dom/server";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, WidthType } from "docx";
/*
 * Local icon primitives.
 * The prototype previously depended on Lucide. These small, text-based controls keep the
 * existing component API intact without adding an icon library to the product shell.
 */
const makeIcon = (glyph, weight = 700) => function LocalIcon({size = 15, color = "currentColor", className = "", style = {}, ...rest}) {
  return <span aria-hidden="true" className={`n-icon ${className}`} style={{fontSize:size, lineHeight:1, color, fontWeight:weight, ...style}} {...rest}>{glyph}</span>;
};
const Home = makeIcon("⌂");
const Package = makeIcon("□");
const ClipboardList = makeIcon("☷");
const Cpu = makeIcon("▦");
const FileText = makeIcon("▤");
const Database = makeIcon("▥");
const BarChart3 = makeIcon("▥");
const ScrollText = makeIcon("≡");
const Users = makeIcon("◉");
const GitBranch = makeIcon("⑂");
const Settings = makeIcon("⚙", 600);
const Search = makeIcon("⌕");
const Camera = makeIcon("▣");
const Mic = makeIcon("◌");
const Upload = makeIcon("↑");
const Printer = makeIcon("▧");
const Download = makeIcon("↓");
const ChevronRight = makeIcon("›");
const ChevronDown = makeIcon("⌄");
const X = makeIcon("×");
const Check = makeIcon("✓");
const RefreshCw = makeIcon("↻");
const Plus = makeIcon("+");
const AlertTriangle = makeIcon("!");
const CheckCircle2 = makeIcon("✓");
const XCircle = makeIcon("×");
const HelpCircle = makeIcon("?");
const Wifi = makeIcon("⌁");
const WifiOff = makeIcon("×");
const Thermometer = makeIcon("T");
const Droplets = makeIcon("H");
const Radio = makeIcon("◉");
const Activity = makeIcon("·");
const Bot = makeIcon("AI");
const Send = makeIcon("↑");
const Fingerprint = makeIcon("≋");
const ShieldCheck = makeIcon("✓");
const ShieldAlert = makeIcon("!");
const Clock = makeIcon("◷");
const ChevronLeft = makeIcon("‹");
const FlaskConical = makeIcon("△");
const History = makeIcon("↺");
const Signal = makeIcon("⌁");
const Library = makeIcon("▧");
const ExternalLink = makeIcon("↗");

/* ============================== STYLE TOKENS ============================== */
const GlobalStyle = () => (
  <style>{`
    .nawi-root{
      --bg:#f7f8f9;
      --surface:#ffffff;
      --surface-2:#f3f4f6;
      --surface-3:#e5e7eb;
      --ink:#111827;
      --ink-dim:#4b5563;
      --ink-faint:#9ca3af;
      --line:#e5e7eb;
      --line-strong:#d1d5db;
      --accent:#111827;
      --accent-soft:#f3f4f6;
      --accent-ink:#ffffff;
      --navy:#111827;
      --navy-dim:#374151;
      --seal:#059669;
      --seal-bg:#ecfdf5;
      --seal-line:#a7f3d0;
      --amber:#d97706;
      --amber-bg:#fffbeb;
      --amber-line:#fde68a;
      --rose:#dc2626;
      --rose-bg:#fef2f2;
      --rose-line:#fecaca;
      --gold:#b45309;
      font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
      color:var(--ink);
      background:var(--bg);
      letter-spacing:-.011em;
      -webkit-font-smoothing:antialiased;
      -moz-osx-font-smoothing:grayscale;
    }
    .nawi-root *{box-sizing:border-box;}
    .nawi-root button,.nawi-root input,.nawi-root select,.nawi-root textarea{font:inherit;}
    .nawi-root button:focus-visible,.nawi-root input:focus-visible,.nawi-root select:focus-visible,.nawi-root textarea:focus-visible,
    .nawi-root [tabindex]:focus-visible{outline:2px solid var(--accent);outline-offset:2px;}
    .nawi-root .n-icon{display:inline-flex;align-items:center;justify-content:center;min-width:1em;font-family:Arial,sans-serif;}
    .f-display{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;}
    .f-mono{font-family:"SFMono-Regular",Consolas,"Liberation Mono",monospace;}
    .nawi-root ::selection{background:#e5e7eb;}
    .n-scroll::-webkit-scrollbar{width:7px;height:7px;}
    .n-scroll::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:9999px;}
    .n-scroll::-webkit-scrollbar-track{background:transparent;}
    .n-panel{background:var(--surface);border:1px solid rgba(229,231,235,0.9);border-radius:28px!important;box-shadow:0 2px 40px -12px rgba(0,0,0,0.05),0 1px 3px rgba(0,0,0,0.02)!important;transition:transform .25s cubic-bezier(0.16,1,0.3,1),box-shadow .25s cubic-bezier(0.16,1,0.3,1),border-color .2s ease;}
    .n-panel-2{background:var(--surface-2);border:1px solid var(--line);border-radius:20px!important;}
    .n-hr{border-top:1px solid var(--line);}
    .n-label{font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-faint);font-weight:700;}
    .n-eyebrow{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--ink-dim);font-weight:750;}
    .n-btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;font-size:13.5px;font-weight:600;line-height:1.25;padding:9px 20px;min-height:38px;border:1px solid var(--line);border-radius:9999px!important;background:var(--surface);color:var(--ink);cursor:pointer;box-shadow:0 1px 2px rgba(0,0,0,0.04);transition:all .2s cubic-bezier(0.16,1,0.3,1);}
    .n-btn:hover{background:var(--surface-2);border-color:var(--line-strong);transform:translateY(-1.5px);box-shadow:0 4px 14px rgba(0,0,0,0.07);}
    .n-btn:active{transform:scale(0.98);}
    .n-btn:disabled{opacity:.45;cursor:not-allowed;transform:none!important;box-shadow:none!important;}
    .n-btn-primary{background:var(--accent);color:var(--accent-ink);border-color:var(--accent);box-shadow:0 2px 8px rgba(17,24,39,0.12);}
    .n-btn-primary:hover{background:#1f2937;border-color:#1f2937;box-shadow:0 6px 20px rgba(17,24,39,0.2);}
    .n-btn-seal{background:var(--seal);color:#fff;border-color:var(--seal);box-shadow:0 2px 8px rgba(5,150,105,0.15);}
    .n-btn-seal:hover{background:#047857;border-color:#047857;box-shadow:0 6px 18px rgba(5,150,105,0.25);}
    .n-btn-rose{background:var(--rose);color:#fff;border-color:var(--rose);}
    .n-btn-rose:hover{background:#b91c1c;border-color:#b91c1c;}
    .n-btn-ghost{background:transparent;border-color:transparent;box-shadow:none;color:var(--ink-dim);}
    .n-btn-ghost:hover{background:var(--surface-2);color:var(--ink);}
    .n-btn-sm{padding:6px 14px;min-height:32px;font-size:12px;}
    .n-input,.n-select,.n-textarea{width:100%;font-size:13.5px;padding:10px 14px;border:1px solid var(--line);border-radius:14px!important;background:var(--surface);color:var(--ink);font-family:inherit;transition:border-color .18s ease,box-shadow .18s ease,background .18s ease;}
    .n-input::placeholder,.n-textarea::placeholder{color:var(--ink-faint);}
    .n-input:focus,.n-select:focus,.n-textarea:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px rgba(17,24,39,0.08);}
    .n-field-label{font-size:12px;font-weight:650;color:var(--ink-dim);margin-bottom:6px;display:block;}
    .n-badge{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:600;padding:4px 10px;border:1px solid;letter-spacing:.02em;border-radius:9999px!important;}
    .n-badge-pass{background:var(--seal-bg);color:#065f46;border-color:var(--seal-line);}
    .n-badge-fail{background:var(--rose-bg);color:#991b1b;border-color:var(--rose-line);}
    .n-badge-review{background:var(--amber-bg);color:#92400e;border-color:var(--amber-line);}
    .n-badge-neutral{background:var(--surface-2);color:var(--ink-dim);border-color:var(--line);}
    .n-badge-navy{background:var(--accent);color:var(--accent-ink);border-color:var(--accent);}
    .n-table{width:100%;border-collapse:separate;border-spacing:0;font-size:13px;}
    .n-table th{text-align:left;font-size:10.5px;letter-spacing:.07em;text-transform:uppercase;color:var(--ink-faint);font-weight:700;padding:12px 16px;border-bottom:1px solid var(--line);background:var(--surface-2);white-space:nowrap;}
    .n-table th:first-child{border-top-left-radius:24px;}
    .n-table th:last-child{border-top-right-radius:24px;}
    .n-table td{padding:13px 16px;border-bottom:1px solid var(--line);vertical-align:middle;transition:background .15s ease;}
    .n-table tr:last-child td:first-child{border-bottom-left-radius:24px;}
    .n-table tr:last-child td:last-child{border-bottom-right-radius:24px;}
    .n-table tr:last-child td{border-bottom:none;}
    .n-table tr.n-row-hover:hover td{background:var(--surface-2);cursor:pointer;}
    .n-row-hover{cursor:pointer;transition:background .15s ease,border-color .15s ease;}
    div.n-row-hover:hover{background:var(--surface-2);border-color:var(--line-strong);}
    .n-tab{padding:10px 8px;font-size:13px;font-weight:600;color:var(--ink-faint);border-bottom:2px solid transparent;cursor:pointer;transition:all .18s ease;}
    .n-tab.active{color:var(--ink);border-color:var(--accent);}
    .n-sidebar-item{display:flex;align-items:center;gap:10px;padding:9px 14px;margin-bottom:2px;font-size:13px;font-weight:550;color:var(--ink-dim);cursor:pointer;border-radius:9999px;transition:all .2s cubic-bezier(0.16,1,0.3,1);}
    .n-sidebar-item:hover{background:var(--surface-2);color:var(--ink);transform:translateX(2px);}
    .n-sidebar-item.active{background:var(--accent);color:#ffffff;font-weight:600;box-shadow:0 4px 14px rgba(17,24,39,0.18);}
    .n-stamp-ring{fill:none;stroke-width:2;}
    .n-flag-valid{color:var(--seal);}
    .n-flag-warn{color:var(--amber);}
    .n-flag-bad{color:var(--rose);}
    .n-kv{display:grid;grid-template-columns:auto 1fr;gap:5px 14px;font-size:13px;}
    .n-kv dt{color:var(--ink-faint);font-weight:600;}
    .n-kv dd{color:var(--ink);font-family:"SFMono-Regular",Consolas,"Liberation Mono",monospace;}
    .nawi-sidebar{width:256px;background:#ffffff;border-right:1px solid var(--line);flex-shrink:0;display:flex;flex-direction:column;}
    .nawi-brand{padding:24px 20px 20px;border-bottom:1px solid var(--line);}
    .nawi-brand-name{color:var(--ink);font-size:20px;font-weight:800;letter-spacing:-.03em;display:flex;align-items:center;gap:8px;}
    .nawi-brand-name::before{content:"";display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--ink);}
    .nawi-brand-sub{color:var(--ink-faint);font-size:9.5px;letter-spacing:.1em;line-height:1.45;margin-top:6px;font-weight:600;}
    .nawi-nav{padding:14px 10px;flex:1;overflow-y:auto;}
    .nawi-support-label{padding:20px 14px 8px;font-size:9.5px;letter-spacing:.1em;color:var(--ink-faint);font-weight:750;text-transform:uppercase;}
    .nawi-sidebar-note{margin:14px;padding:12px 14px;border-radius:16px;background:var(--surface-2);border:1px solid var(--line);color:var(--ink-dim);font-size:11px;line-height:1.5;}
    .nawi-topbar{display:flex;align-items:center;justify-content:space-between;gap:18px;padding:12px 32px;border-bottom:1px solid var(--line);background:var(--surface);min-height:64px;}
    .nawi-main{flex:1;display:flex;flex-direction:column;min-width:0;min-height:100vh;}
    .nawi-body{flex:1;overflow:auto;padding:32px 38px 44px;}
    .nawi-page{max-width:1420px;margin:0 auto;}
    .nawi-footer{display:flex;justify-content:space-between;gap:20px;align-items:center;margin-top:56px;padding:24px 0 16px;border-top:1px solid var(--line);font-size:12px;color:var(--ink-faint);}
    .nawi-footer a{color:var(--ink-dim);text-decoration:none;font-weight:550;transition:color .15s ease;}
    .nawi-footer a:hover{color:var(--ink);text-decoration:underline;}
    .nawi-page-title{font-size:34px!important;letter-spacing:-.03em!important;line-height:1.12!important;font-weight:800!important;}
    .nawi-page-desc{max-width:760px;line-height:1.6!important;font-size:14.5px!important;color:var(--ink-dim)!important;}
    .nawi-product-note{display:grid;grid-template-columns:minmax(0,1.5fr) minmax(260px,.75fr);gap:28px;align-items:end;margin-bottom:30px;}
    .nawi-product-note .lead{font-size:15px;line-height:1.65;color:var(--ink-dim);max-width:760px;}
    @keyframes fadeUp{from{opacity:0;transform:translateY(16px);}to{opacity:1;transform:translateY(0);}}
    @keyframes orbitRotate{from{transform:rotate(0deg);}to{transform:rotate(360deg);}}
    @keyframes orbitRotateReverse{from{transform:rotate(0deg);}to{transform:rotate(-360deg);}}
    .nawi-body > div{animation:fadeUp .45s cubic-bezier(0.16,1,0.3,1) forwards;}
    @media (max-width:1100px){
      .nawi-sidebar{width:220px;}
      .nawi-body{padding:26px 24px 36px;}
      .nawi-topbar{padding-left:20px;padding-right:20px;}
    }
    @media (max-width:800px){
      .nawi-root .nawi-shell{display:block!important;min-height:100vh;}
      .nawi-sidebar{width:100%;height:auto;position:sticky;top:0;z-index:60;border-right:0;border-bottom:1px solid var(--line);}
      .nawi-brand{padding:12px 16px;display:flex;align-items:center;justify-content:space-between;gap:12px;}
      .nawi-brand-sub{margin-top:0;text-align:right;max-width:210px;}
      .nawi-nav{display:flex;overflow-x:auto;padding:6px;-webkit-overflow-scrolling:touch;}
      .nawi-nav .n-sidebar-item{white-space:nowrap;margin:0 4px;padding:8px 14px;}
      .nawi-support-label,.nawi-sidebar-note{display:none;}
      .nawi-topbar{align-items:stretch;flex-direction:column;gap:12px;padding:12px 16px;}
      .nawi-topbar-left,.nawi-topbar-right{width:100%;display:flex!important;flex-wrap:wrap;gap:8px!important;}
      .nawi-topbar-right{align-items:center;}
      .nawi-topbar-right > div[style*="position:relative"]{flex:1;min-width:180px;}
      .nawi-topbar-right .n-input{width:100%!important;}
      .nawi-body{padding:20px 16px 32px;overflow-x:hidden;}
      .nawi-page-title{font-size:26px!important;}
      .nawi-product-note{grid-template-columns:1fr;gap:14px;}
      .nawi-footer{align-items:flex-start;flex-direction:column;gap:10px;margin-top:36px;}
      .nawi-root [style*="grid-template-columns"]{grid-template-columns:1fr!important;}
      .nawi-root .n-table{display:block;overflow-x:auto;white-space:nowrap;}
      .nawi-root [style*="width:320px"]{width:min(320px,calc(100vw - 28px))!important;}
    }
    @media (max-width:520px){
      .nawi-brand-name{font-size:16px;}
      .nawi-brand-sub{font-size:8.5px;max-width:165px;}
      .nawi-topbar-right select{max-width:100%;flex:1;}
      .n-repo-status{font-size:10.5px;color:var(--ink-faint);white-space:nowrap;}
      .nawi-page-title{font-size:22px!important;}
      .n-btn{max-width:100%;}
      .nawi-footer{font-size:11px;}
    }
    @media print{
      @page{size:A4;margin:14mm 12mm;}
      html,body{background:#fff!important;}
      body{print-color-adjust:exact;-webkit-print-color-adjust:exact;}
      .no-print{display:none!important;}
      .nawi-shell,.nawi-main,.nawi-body{display:block!important;min-height:0!important;}
      .nawi-sidebar,.nawi-topbar,.nawi-footer{display:none!important;}
      .print-area{position:static!important;width:100%!important;border:0!important;padding:0!important;margin:0!important;}
      .print-area table{page-break-inside:auto;}
      .print-area tr{page-break-inside:avoid;page-break-after:auto;}
      .print-section{page-break-inside:avoid;}
      .print-evidence-image{max-width:150mm!important;max-height:75mm!important;object-fit:contain;}
      .n-panel{border:0!important;}
    }
    @media (prefers-reduced-motion:reduce){
      .nawi-root *{scroll-behavior:auto!important;transition:none!important;animation:none!important;}
    }
  `}</style>
);

/* ============================== UTILITIES ============================== */
let _seq = 1000;
let _nawiSeq = 143;
const nextNawiId = () => `NAWI-${String(_nawiSeq++).padStart(6,"0")}`;
const uid = (p) => `${p}-${(_seq++).toString(36)}${Math.random().toString(36).slice(2,5)}`;
const nowISO = () => new Date().toISOString();
const fmtDT = (iso) => new Date(iso).toLocaleString("en-IN",{ day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"});
const fmtT = (iso) => new Date(iso).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit",second:"2-digit"});
const round = (n,d=3) => Math.round(n*10**d)/10**d;
const METRA_API_BASE = (typeof import.meta !== "undefined" && import.meta.env?.VITE_API_BASE) || `http://${window.location.hostname || "localhost"}:8787`;
async function naviApi(path, options={}){
  const headers={"Content-Type":"application/json", ...(options.headers||{})};
  const response=await fetch(`${METRA_API_BASE}${path}`, {...options, headers, credentials:"include"});
  let data=null;
  try{ data=await response.json(); }catch{ data={}; }
  if(!response.ok){ throw new Error(data?.error || `METRA API request failed (${response.status})`); }
  return data;
}

function hashStr(str){
  let h1=0xdeadbeef, h2=0x41c6ce57;
  for(let i=0;i<str.length;i++){
    const ch=str.charCodeAt(i);
    h1=Math.imul(h1^ch,2654435761); h2=Math.imul(h2^ch,1597334677);
  }
  h1=Math.imul(h1^(h1>>>16),2246822507)^Math.imul(h2^(h2>>>13),3266489909);
  h2=Math.imul(h2^(h2>>>16),2246822507)^Math.imul(h1^(h1>>>13),3266489909);
  const out = 4294967296*(2097151&h2)+(h1>>>0);
  return out.toString(16).toUpperCase().padStart(14,"0");
}

// Content-based report fingerprint. Unlike a timestamp-only hash, this changes when the
// issued report's instrument data, linked tests, observations, or rule version changes.
// Prototype integrity fingerprint only; production should use a cryptographic signature.
function reportContentFingerprint(instrument, tests, ruleVersionId){
  const payload = {
    instrumentId: instrument.id, manufacturer: instrument.manufacturer, model: instrument.model,
    serialNumber: instrument.serialNumber, accuracyClass: instrument.accuracyClass,
    maxCapacity: instrument.maxCapacity, minCapacity: instrument.minCapacity, e: instrument.e,
    ruleVersionId,
    tests: tests.map(t=>({id:t.id, testKey:t.testKey, status:t.status, result:t.result,
      observations:t.observations||[], remarks:t.remarks||""})).sort((a,b)=>a.id.localeCompare(b.id))
  };
  return hashStr(JSON.stringify(payload));
}

/* ============================== OIML R-76 RULE ENGINE ==============================
   Centralized, class-aware, verification-stage-aware MPE engine (OIML R-76-1:2006 §3.6/3.7).
   All class/stage parameters live in OIML_CLASS_TABLE / MPE_STEPS_INITIAL so a future OIML
   revision can be incorporated by editing this table rather than rewriting components that
   consume it. Nothing outside this block should hardcode an MPE threshold or class limit. */

// Structural limits per accuracy class: permitted range of verification scale intervals (n),
// and the minimum capacity expressed as a multiple of e. Source: OIML R76-1:2006 Table (Min/n).
// NOTE: these figures are the commonly published R76 reference values; a metrology reviewer
// should confirm them against the exact edition/amendment in force before certifying results.
const OIML_CLASS_TABLE = {
  I:    { label:"Class I — Special accuracy",   nMin:50000, nMax:1000000, minCapacityInE:100,
          mpeSteps:[{upTo:50000,mpeE:0.5},{upTo:200000,mpeE:1.0},{upTo:null,mpeE:1.5}] },
  II:   { label:"Class II — High accuracy",     nMin:100,   nMax:100000,  minCapacityInE:20,
          coarse:{ eFromG:0.1, nMin:5000, minCapacityInE:50 },
          mpeSteps:[{upTo:5000,mpeE:0.5},{upTo:20000,mpeE:1.0},{upTo:null,mpeE:1.5}] },
  III:  { label:"Class III — Medium accuracy",  nMin:100,   nMax:10000,   minCapacityInE:20,
          coarse:{ eFromG:5, nMin:500, minCapacityInE:20 },
          mpeSteps:[{upTo:500,mpeE:0.5},{upTo:2000,mpeE:1.0},{upTo:null,mpeE:1.5}] },
  IIII: { label:"Class IIII — Ordinary accuracy", nMin:100, nMax:1000,    minCapacityInE:10,
          mpeSteps:[{upTo:50,mpeE:0.5},{upTo:200,mpeE:1.0},{upTo:null,mpeE:1.5}] },
};
// e expressed in grams (the app supports kg and g capacity units).
const eInGrams = (instrument) => Number(instrument?.e) * (instrument?.capacityUnit==="g" ? 1 : 1000);
// Effective n-range / Min limits for one class, honouring the coarse-e band of R 76-1 Table 3
// (Class II: e >= 0.1 g; Class III: e >= 5 g).
function classLimitsFor(cc, instrument){
  const eg = eInGrams(instrument);
  if(cc?.coarse && Number.isFinite(eg) && eg >= Number(cc.coarse.eFromG)){
    return { nMin:Number(cc.coarse.nMin), nMax:Number(cc.nMax), minCapacityInE:Number(cc.coarse.minCapacityInE) };
  }
  return { nMin:Number(cc.nMin), nMax:Number(cc.nMax), minCapacityInE:Number(cc.minCapacityInE) };
}
const classConfig = (accuracyClass, rule=null) => {
  const table = rule?.classTable || OIML_CLASS_TABLE;
  return table[accuracyClass] || table.III || OIML_CLASS_TABLE.III;
};

// MPE step tables (initial verification), expressed in verification scale intervals (e), keyed by
// the number of intervals ("load in e") represented by the test load. Per OIML R76-1:2006 Table 6
// the breakpoints DIFFER PER CLASS (Class I 50 000e/200 000e, II 5 000e/20 000e, III 500e/2 000e,
// IIII 50e/200e); they live in OIML_CLASS_TABLE[class].mpeSteps. In-service (post-installation,
// periodic) verification applies double the initial-verification MPE.
const IN_SERVICE_MULTIPLIER = 2;
const cloneRuleTable = () => JSON.parse(JSON.stringify(OIML_CLASS_TABLE));
const cloneMpeSteps = (cls="III") => JSON.parse(JSON.stringify((OIML_CLASS_TABLE[cls]||OIML_CLASS_TABLE.III).mpeSteps));
function stepsForClass(rule, accuracyClass){
  const own = rule?.classTable?.[accuracyClass]?.mpeSteps;
  if(Array.isArray(own) && own.length) return own;
  const builtIn = OIML_CLASS_TABLE[accuracyClass]?.mpeSteps;
  return builtIn || null; // unknown class -> null -> NaN MPE -> INVALID, never a silent PASS/FAIL
}
function ruleMpeInE(rule, loadInE, stage="initial", accuracyClass="III"){
  const steps = stepsForClass(rule, accuracyClass);
  if(!steps) return NaN;
  const n = Math.abs(Number(loadInE));
  const step = steps.find(s=>s.upTo===null || s.upTo===undefined || n<=Number(s.upTo));
  const base = Number(step?.mpeE);
  if(!Number.isFinite(base)) return NaN;
  return stage==="in-service" ? round(base*IN_SERVICE_MULTIPLIER,2) : base;
}

const OFFICIAL_REFERENCE_SOURCES = [
  {id:"SRC-OIML-R76", authority:"OIML", title:"OIML R 76-1 — Non-automatic weighing instruments", type:"International recommendation", edition:"R 76-1:2006 (Ed. 2017 reference record)", status:"REFERENCE — VERIFY APPLICABLE EDITION", url:"https://www.oiml.org/", use:"Technical basis for NAWI test procedures and rule configuration.", provenance:"Official OIML source domain. Exact edition/amendment must be confirmed before regulatory use."},
  {id:"SRC-OIML-PUB", authority:"OIML", title:"OIML Publications Library", type:"Standards library", edition:"Current publication catalogue", status:"OFFICIAL SOURCE DIRECTORY", url:"https://www.oiml.org/", use:"Locate the authoritative recommendation, amendment or edition applicable to a test."},
  {id:"SRC-DOCA-LM", authority:"Government of India — Department of Consumer Affairs", title:"Legal Metrology", type:"Government portal", edition:"Current portal content", status:"OFFICIAL GOVERNMENT SOURCE", url:"https://consumeraffairs.nic.in/", use:"Government-side legal-metrology references, notices and departmental material."},
  {id:"SRC-INDIA-CODE", authority:"Government of India — India Code", title:"Central legislation repository", type:"Government legislation portal", edition:"Current consolidated portal", status:"OFFICIAL GOVERNMENT SOURCE", url:"https://www.indiacode.nic.in/", use:"Verify the current text of applicable Central Acts and rules before regulatory use."},
];

const RULE_VERSIONS = [
  {
    id:"R76-CFG-2.1", version:"OIML R-76-1:2006 (Ed. 2017) — Class-Aware Configuration v2.1",
    effective:"2026-09-28", status:"ACTIVE",
    notes:"Per-class MPE breakpoints per R 76-1 Table 6 (I: 50 000e/200 000e; II: 5 000e/20 000e; III: 500e/2 000e; IIII: 50e/200e), coarse-e n-range/Min limits per Table 3 (Class II e>=0.1 g, Class III e>=5 g), initial and in-service (2x) stages. Supersedes v2.0, which applied the Class III breakpoints to every class.",
    classTable: cloneRuleTable(),
  },
  {
    id:"R76-CFG-2.0", version:"OIML R-76-1:2006 (Ed. 2017) — Class-Aware Configuration v2.0",
    effective:"2026-01-01", status:"SUPERSEDED",
    notes:"Superseded 2026-09-28 by v2.1. Defect: one 500e/2 000e MPE table was used for all classes (correct only for Class III). Results for Class III are unaffected; Class I, II and IIII results issued under v2.0 must be re-evaluated.",
    classTable: cloneRuleTable(),
  },
  {
    id:"R76-CFG-1.2", version:"OIML R-76-1:2006 (Ed. 2017) — Local Configuration v1.2",
    effective:"2024-01-01", status:"SUPERSEDED",
    notes:"Historical configuration retained for traceability. Superseded by v2.0 on 2026-01-01.",
    classTable: cloneRuleTable(),
  },
  {
    id:"R76-CFG-1.1", version:"OIML R-76-1:2006 (Ed. 2017) — Local Configuration v1.1",
    effective:"2022-06-01", status:"SUPERSEDED",
    notes:"Historical configuration retained for traceability. Superseded by v1.2 on 2024-01-01.",
    classTable: cloneRuleTable(),
  },
];
const activeRule = () => RULE_VERSIONS.find(r=>r.status==="ACTIVE");

// Zero-return tolerance, expressed as a fraction of e (commonly published R76 reference value;
// flagged, like OIML_CLASS_TABLE above, for a metrology reviewer to confirm against the exact
// edition/amendment in force). Doubled for in-service verification, same as the main MPE table.
const ZERO_RETURN_TOLERANCE_E = 0.25;
function evalZeroReturn(observedZero, instrument, stage){
  const e = instrument.e;
  const verificationStage = stage || instrument.verificationStage || "initial";
  const toleranceE = verificationStage==="in-service" ? round(ZERO_RETURN_TOLERANCE_E*IN_SERVICE_MULTIPLIER,3) : ZERO_RETURN_TOLERANCE_E;
  const toleranceAbs = round(toleranceE*e,4);
  const error = round(Number(observedZero),4);
  const result = Math.abs(error) <= toleranceAbs ? "PASS" : "FAIL";
  return {error, toleranceE, toleranceAbs, result, e, verificationStage, accuracyClass:instrument.accuracyClass, testLoad:0, observedValue:Number(observedZero), referenceValue:0};
}

// Structural validation of an instrument's declared class/Max/Min/e against OIML R76 limits.
// Used at registration, on the Test Readiness panel, and before submission (items 3 & 11).
function validateClassConfig(instrument, rule=null){
  const flags = [];
  const active = rule || activeRule();
  const table = active?.classTable || OIML_CLASS_TABLE;
  const cls = instrument.accuracyClass;
  if(!table[cls]){ flags.push({level:"bad", msg:`Unrecognized accuracy class "${cls}".`}); return flags; }
  const cc = classConfig(cls, active);
  const e = Number(instrument.e), max = Number(instrument.maxCapacity), min = Number(instrument.minCapacity);
  if(!(e>0)) flags.push({level:"bad", msg:"Verification scale interval (e) must be greater than zero."});
  if(!(max>min)) flags.push({level:"bad", msg:"Maximum capacity (Max) must be greater than Minimum capacity (Min)."});
  if(e>0 && max>0){
    const n = max/e;
    const lim = classLimitsFor(cc, instrument);
    if(n > lim.nMax) flags.push({level:"bad", msg:`Number of verification intervals n = Max/e = ${round(n,0)} exceeds the Class ${cls} maximum of ${lim.nMax}. Selected class is not compatible with this Max/e configuration.`});
    else if(n < lim.nMin) flags.push({level:"bad", msg:`Number of verification intervals n = Max/e = ${round(n,0)} is below the Class ${cls} minimum of ${lim.nMin} for this e. Selected class is not compatible with this Max/e configuration.`});
    const expectedMin = round(lim.minCapacityInE*e,6);
    if(min>=0 && min < expectedMin) flags.push({level:"bad", msg:`Registered Min (${min} ${instrument.capacityUnit||""}) is below the Class ${cls} minimum reference of ${lim.minCapacityInE}e = ${expectedMin} ${instrument.capacityUnit||""}.`});
  }
  if(flags.length===0) flags.push({level:"ok", msg:`Class ${cls} configuration valid (n = ${round(max/e,0)}, permitted range ${classLimitsFor(cc, instrument).nMin}–${classLimitsFor(cc, instrument).nMax}).`});
  return flags;
}

// Full OIML calculation record for one test-point comparison. Every field item 1 requires
// (class, Max, Min, e, n, test load, load-in-e, stage, applicable/actual MPE, observed,
// reference, error, PASS/FAIL) is carried on the returned object so it can be stored on the
// observation, displayed, and reproduced later by "Explain Calculation" / "Why PASS/FAIL".
function evalAccuracy(load, observed, instrument, rule, stage){
  const e = Number(instrument.e);
  if(!Number.isFinite(e) || e<=0 || !Number.isFinite(Number(load)) || !Number.isFinite(Number(observed))) return {result:"INVALID", error:null, mpeE:null, mpeVal:null, loadInE:null, ruleVersionId:rule?.id||null};
  const verificationStage = stage || instrument.verificationStage || "initial";
  const error = round(observed-load,4);
  const loadInE = round(load/e,2);
  const mpeE = ruleMpeInE(rule, loadInE, verificationStage, instrument.accuracyClass);
  const mpeVal = round(mpeE*e,4);
  const result = Math.abs(error) <= mpeVal ? "PASS" : "FAIL";
  return {
    error, mpeE, mpeVal, result, loadInE,
    accuracyClass: instrument.accuracyClass, maxCapacity: instrument.maxCapacity, minCapacity: instrument.minCapacity,
    e, n: round(instrument.maxCapacity/e,1), testLoad: load, verificationStage,
    observedValue: observed, referenceValue: load, ruleVersionId: rule.id,
  };
}
function evalRepeatability(readings, load, instrument, rule, stage){
  if(readings.length<2) return null;
  const e = Number(instrument.e);
  if(!Number.isFinite(e) || e<=0 || !Number.isFinite(Number(load))) return {result:"INVALID", range:null, allowed:null, mpeE:null, ruleVersionId:rule?.id||null};
  const verificationStage = stage || instrument.verificationStage || "initial";
  const max=Math.max(...readings), min=Math.min(...readings);
  const range = round(max-min,4);
  const loadInE = round(load/e,2);
  const mpeE = ruleMpeInE(rule, loadInE, verificationStage, instrument.accuracyClass);
  const allowed = round(mpeE*e,4);
  const result = range<=allowed ? "PASS":"FAIL";
  return {
    range, allowed, mpeE, result, max, min, e, loadInE, testLoad:load,
    accuracyClass: instrument.accuracyClass, maxCapacity: instrument.maxCapacity, n: round(instrument.maxCapacity/e,1),
    verificationStage, ruleVersionId: rule.id,
  };
}

/* ============================== TEST DEFINITIONS ============================== */
const TEST_DEFS = {
  ACC: {key:"ACC", name:"Accuracy (Weighing) Test", purpose:"Determine indication error at test loads across the weighing range and compare against the maximum permissible error (MPE).",
    requiredObservations:"Test load, observed indication at 5 points (Min, 25%, 50%, 75%, Max capacity)",
    requiredEquipment:"Certified test weights / substitution weights, calibrated reference", envRequirements:"Stable ambient temperature, draft-free"},
  REP: {key:"REP", name:"Repeatability Test", purpose:"Verify consistency of indication under repeated application of the same load.",
    requiredObservations:"5 consecutive readings at a fixed test load (~50% capacity)",
    requiredEquipment:"Certified test weight matching selected load", envRequirements:"Stable ambient conditions during full test sequence"},
  DIS: {key:"DIS", name:"Discrimination Test", purpose:"Confirm the instrument responds to a small additional mass at high load.",
    requiredObservations:"Indication before and after adding 1e at/near maximum capacity",
    requiredEquipment:"Small calibrated increment mass (1e)", envRequirements:"None beyond standard lab conditions"},
  ECC: {key:"ECC", name:"Eccentricity (Corner Load) Test", purpose:"Verify indication error remains within tolerance when the load is applied off-center on the load-receptor.",
    requiredObservations:"Indication at center and at 4 corner positions with a fixed load",
    requiredEquipment:"Certified test weight, corner-loading template", envRequirements:"Level, stable mounting surface"},
  SEN: {key:"SEN", name:"Sensitivity Test", purpose:"Confirm the instrument produces a perceptible, correctly-signed change in indication for the smallest applied increment, at a load chosen by the technician (not only near Max).",
    requiredObservations:"Indication before and after adding a 1e increment mass at the selected load", requiredEquipment:"Small calibrated increment mass (1e)", envRequirements:"None beyond standard lab conditions"},
  ZERO: {key:"ZERO", name:"Zero-Return Test", purpose:"Verify the indication returns to zero, within tolerance, after a load equal to Max capacity is applied and fully removed.",
    requiredObservations:"Indication after removing a full Max-capacity load", requiredEquipment:"Certified test weight(s) totalling Max capacity", envRequirements:"Stable ambient conditions"},
  CREEP: {key:"CREEP", name:"Creep Test", purpose:"Verify indication does not drift beyond tolerance while Max capacity remains applied over a sustained period (typically 30 minutes).",
    requiredObservations:"Indication immediately after applying Max load, and again after 30 minutes under that same load", requiredEquipment:"Certified test weight(s) totalling Max capacity, stopwatch/timer", envRequirements:"Stable ambient conditions for the full duration"},
  TARE: {key:"TARE", name:"Tare Performance Test", purpose:"Verify indication error remains within the applicable MPE when the tare function is used to zero out a container before weighing a net load.",
    requiredObservations:"Net indication with tare set, against the certified net test load", requiredEquipment:"Tare vessel/container, certified test weight for the net load", envRequirements:"None beyond standard lab conditions"},
  TEMP: {key:"TEMP", name:"Temperature Effect Test", purpose:"Verify indication under a fixed load does not drift beyond the applicable tolerance across a stated change in ambient temperature.",
    requiredObservations:"Indication at a reference temperature and again after a recorded temperature change, same fixed load", requiredEquipment:"Certified test weight, calibrated thermometer / environmental chamber", envRequirements:"Ambient temperature recorded at both readings"},
  WARM: {key:"WARM", name:"Warm-Up Test", purpose:"Verify the instrument's indication stabilizes within its stated warm-up time before testing begins.",
    requiredObservations:"Successive zero/no-load readings taken at intervals after power-on until stable", requiredEquipment:"Stopwatch/timer", envRequirements:"Instrument powered on from cold/off state"},
  VOLT: {key:"VOLT", name:"Voltage Variation Test", purpose:"Verify indication under a fixed load does not drift beyond the applicable tolerance across the rated supply-voltage range (electronic instruments only).",
    requiredObservations:"Indication at nominal supply voltage and again at a rated voltage extreme, same fixed load", requiredEquipment:"Certified test weight, variable-voltage supply", envRequirements:"Not applicable to non-electronic (mechanical) instruments"},
  SPAN: {key:"SPAN", name:"Span Stability Test", purpose:"Verify the instrument's span (full-scale calibration) remains within tolerance between two calibration checks separated in time.",
    requiredObservations:"Indication at Max load on two separate span-check occasions", requiredEquipment:"Certified test weight(s) totalling Max capacity", envRequirements:"None beyond standard lab conditions"},
};
function generateTestPlan(instrument){
  const items = [
    {...TEST_DEFS.ACC, status:"applicable"},
    {...TEST_DEFS.REP, status:"applicable"},
    {...TEST_DEFS.DIS, status:"applicable"},
    {...TEST_DEFS.ECC, status: instrument.maxCapacity>=1 ? "applicable" : "not-applicable"},
    {...TEST_DEFS.SEN, status:"applicable"},
    {...TEST_DEFS.ZERO, status:"applicable"},
    {...TEST_DEFS.CREEP, status:"applicable"},
    {...TEST_DEFS.TARE, status:"applicable"},
    {...TEST_DEFS.TEMP, status:"applicable"},
    {...TEST_DEFS.WARM, status:"applicable"},
    {...TEST_DEFS.VOLT, status:"applicable"},
    {...TEST_DEFS.SPAN, status:"applicable"},
  ];
  return items;
}

/* ============================== VALIDATION ============================== */
function validateObservation(instrument, testKey, {load, observed}, existingLoads=[]){
  const flags = [];
  if(load===""||load===null||isNaN(load)) return [{level:"bad", msg:"Test load is required."}];
  if(observed===""||observed===null||isNaN(observed)) return [{level:"bad", msg:"Observed value is required."}];
  const L=Number(load), O=Number(observed), e=Number(instrument.e);
  if(!(e>0)) flags.push({level:"bad", msg:"Instrument verification scale interval (e) is invalid."});
  if(L < 0) flags.push({level:"bad", msg:"Test load cannot be negative."});
  if(L > Number(instrument.maxCapacity)) flags.push({level:"bad", msg:`Entered test load (${L}) exceeds the registered maximum capacity (${instrument.maxCapacity} ${instrument.capacityUnit}).`});
  if(L < Number(instrument.minCapacity) && L!==0) flags.push({level:"warn", msg:`Entered test load is below the registered minimum capacity (${instrument.minCapacity} ${instrument.capacityUnit}).`});
  if(e>0 && L>0){
    const intervals=L/e;
    const aligned=Math.abs(intervals-Math.round(intervals))<1e-9;
    if(!aligned) flags.push({level:"bad", msg:`Test load ${L} ${instrument.capacityUnit} is not aligned to the verification scale interval e = ${e} ${instrument.capacityUnit}. Use a whole-number multiple of e.`});
  }
  if(O < 0) flags.push({level:"bad", msg:"Observed value cannot be negative."});
  if(existingLoads.includes(L)) flags.push({level:"warn", msg:"A reading already exists at this test load. Duplicate observation."});
  const deviationPct = Math.abs(O-L) / (Number(instrument.maxCapacity)||1);
  if(deviationPct > 0.05) flags.push({level:"warn", msg:"Observed value deviates unusually far from the entered test load. Verify data-acquisition setup."});
  if(flags.length===0) flags.push({level:"ok", msg:"Observation is structurally valid and aligned to e."});
  return flags;
}
// Required observation/repetition counts per test type, shared by progress bars, the
// readiness panel, and submission validation — one source of truth.
const EXPECTED_OBS = {ACC:5, REP:5, DIS:1, ECC:5, SEN:1, ZERO:1, CREEP:1, TARE:1, TEMP:1, WARM:3, VOLT:1, SPAN:1};

function numericValue(v){ return v!=="" && v!==null && v!==undefined && Number.isFinite(Number(v)); }
function pushBad(flags,msg){ flags.push({level:"bad",msg}); }
function pushWarn(flags,msg){ flags.push({level:"warn",msg}); }

// Test-specific validation. The validator checks the data shape produced by each test
// workspace before a test can enter review. It deliberately validates the prototype's
// configured workflow; it does not claim to replace clause-by-clause metrology review.
function validateTestSpecific(instrument, test, flags){
  const obs = test.observations || [];
  const key = test.testKey;
  const max = Number(instrument.maxCapacity);
  const min = Number(instrument.minCapacity);

  if(["ACC","REP"].includes(key)){
    obs.forEach((o,i)=>{
      if(!numericValue(o.load) || !numericValue(o.observed)) pushBad(flags,`${key} observation ${i+1} has a missing or non-numeric load/indication.`);
      else if(Number(o.load)<0 || Number(o.load)>max) pushBad(flags,`${key} observation ${i+1} load must be between 0 and Max capacity (${max} ${instrument.capacityUnit}).`);
      else if(Number(o.observed)<0) pushBad(flags,`${key} observation ${i+1} indication cannot be negative.`);
    });
    if(key==="ACC"){
      const loads=obs.map(o=>Number(o.load));
      if(new Set(loads).size!==loads.length) pushWarn(flags,"Accuracy test contains duplicate load points. Confirm that repeated points are intentional.");
    }
  }

  if(key==="DIS"){
    const o=obs[0];
    if(obs.length!==1) pushBad(flags,"Discrimination requires exactly one recorded check.");
    if(o){
      if(!numericValue(o.before)) pushBad(flags,"Discrimination requires a numeric indication before the 1e increment.");
      else if(Number(o.before)<0 || Number(o.before)>max) pushBad(flags,"Discrimination load must be within the instrument capacity range.");
      if(typeof o.changed!=="boolean") pushBad(flags,"Discrimination requires an explicit Yes/No indication-change result.");
    }
  }

  if(key==="ECC"){
    const expected=["Center","Corner 1","Corner 2","Corner 3","Corner 4"];
    const positions=obs.map(o=>o.position);
    if(obs.length!==5 || expected.some(p=>!positions.includes(p))) pushBad(flags,"Eccentricity requires Center and all four corner observations.");
    if(new Set(positions).size!==positions.length) pushBad(flags,"Eccentricity contains duplicate position observations.");
    obs.forEach((o,i)=>{
      if(!numericValue(o.observed)) pushBad(flags,`Eccentricity position ${i+1} has no valid numeric indication.`);
      else if(Number(o.observed)<0) pushBad(flags,"Eccentricity indications cannot be negative.");
      if(!numericValue(o.load) || Number(o.load)<=0 || Number(o.load)>max) pushBad(flags,"Eccentricity test load must be positive and within Max capacity.");
    });
  }

  if(key==="SEN"){
    const o=obs[0];
    if(obs.length!==1) pushBad(flags,"Sensitivity requires exactly one recorded check.");
    if(o){
      if(!numericValue(o.load) || Number(o.load)<=0 || Number(o.load)>max) pushBad(flags,"Sensitivity test load must be positive and within Max capacity.");
      if(!numericValue(o.before)) pushBad(flags,"Sensitivity requires a numeric indication before adding 1e.");
      if(typeof o.changed!=="boolean") pushBad(flags,"Sensitivity requires an explicit Yes/No indication-change result.");
    }
  }

  if(key==="WARM"){
    if(obs.length<3) pushBad(flags,"Warm-up requires at least 3 successive readings before submission.");
    obs.forEach((o,i)=>{ if(!numericValue(o.observed) || Number(o.observed)<0) pushBad(flags,`Warm-up reading ${i+1} must contain a valid non-negative indication.`); });
    if(obs.length>=3 && test.result!=="PASS") pushBad(flags,"Warm-up is not complete: the latest three readings have not demonstrated stability within one scale interval.");
  }

  if(["ZERO","CREEP","TARE","TEMP","VOLT","SPAN"].includes(key)){
    const o=obs[0];
    if(obs.length!==1) pushBad(flags,`${key} requires exactly one reference/observed reading pair.`);
    if(o){
      if(!numericValue(o.observed)) pushBad(flags,`${key} requires a numeric observed indication.`);
      if(key!=="ZERO" && !numericValue(o.load)) pushBad(flags,`${key} requires a numeric test/reference load.`);
      if(key!=="ZERO" && numericValue(o.load) && (Number(o.load)<=0 || Number(o.load)>max)) pushBad(flags,`${key} test/reference load must be positive and within Max capacity.`);
      if(key==="ZERO" && numericValue(o.load) && Number(o.load)!==0) pushBad(flags,"Zero-return reference load must be zero.");
      if(numericValue(o.observed) && Number(o.observed)<0) pushBad(flags,`${key} observed indication cannot be negative.`);
    }
  }

  // A calculated result must exist before review. This catches records created through
  // alternate capture paths that bypass the visible result badge.
  if(obs.length>0 && !["WARM"].includes(key) && obs.some(o=>o.result===undefined || o.result===null)){
    pushBad(flags,"One or more recorded observations do not contain a calculated compliance result.");
  }
}

// Calibration-status check for a reference/test equipment record.
function equipmentStatus(equip){
  const flags = [];
  if(!equip){ flags.push({level:"bad", msg:"No reference/test equipment selected for this test."}); return flags; }
  if(!equip.certNumber) flags.push({level:"bad", msg:"No calibration certificate number on record for this equipment."});
  if(!equip.traceability) flags.push({level:"warn", msg:"No traceability information on record for this equipment."});
  if(equip.validUntil){
    const expired = new Date(equip.validUntil) < new Date();
    if(expired) flags.push({level:"bad", msg:`Calibration for ${equip.equipmentId} expired on ${equip.validUntil}. This equipment cannot be used until recalibrated.`});
  } else flags.push({level:"warn", msg:"No calibration expiry date on record for this equipment."});
  if(flags.length===0) flags.push({level:"ok", msg:`${equip.equipmentId} — calibration valid until ${equip.validUntil}.`});
  return flags;
}

function validateTestSubmission(instrument, test, {envReadings=[], equipment=null}={}){
  const flags = [];
  const required = EXPECTED_OBS[test.testKey] || 1;
  const obsCount = (test.observations||[]).length;
  if(obsCount===0) pushBad(flags,"No observations recorded yet.");
  else if(obsCount < required) pushBad(flags,`${obsCount} of ${required} required ${test.testKey==="REP"?"repetitions":"test points"} recorded.`);
  else flags.push({level:"ok", msg:`${obsCount} required ${test.testKey==="REP"?"repetitions":"observation(s)"} recorded.`});

  validateTestSpecific(instrument,test,flags);

  if(envReadings.length===0) pushWarn(flags,"No environmental conditions recorded for this test.");
  else flags.push({level:"ok", msg:`${envReadings.length} environmental reading(s) linked.`});

  const eqFlags = equipmentStatus(equipment);
  flags.push(...eqFlags.map(f=>({...f,msg:`Reference equipment: ${f.msg}`})));
  const classFlags = validateClassConfig(instrument, RULE_VERSIONS.find(r=>r.id===test.ruleVersionId)||activeRule());
  flags.push(...classFlags.filter(f=>f.level==="bad").map(f=>({...f,msg:`Instrument configuration: ${f.msg}`})));

  const canSubmit = !flags.some(f=>f.level==="bad");
  return {flags, canSubmit};
}
function detectRepeatAnomaly(readings){
  if(readings.length<3) return null;
  const last3 = readings.slice(-3);
  if(last3.every(r=>r===last3[0])) return "Multiple consecutive readings are identical. Verify the measurement and data-acquisition setup.";
  return null;
}

/* ============================== MEASUREMENT FINGERPRINT ENGINE ============================== */
/* Transparent, statistical, history-based trend indicator. Never overrides the OIML rule engine. */
const LOAD_PCTS = [0, 25, 50, 75, 100];
const DEFAULT_FP_CONFIG = { minEvaluations: 2, changeThresholdE: 2, significantThresholdE: 5 };

function findMaxLoadObs(test, instrument){
  if(!test || !test.observations || test.observations.length===0) return null;
  const target = instrument.maxCapacity;
  return test.observations.reduce((best,o)=>{
    if(o.load===undefined || o.error===undefined) return best;
    if(!best) return o;
    return Math.abs(o.load-target) < Math.abs(best.load-target) ? o : best;
  }, null);
}
// Maps a set of {load, error, ...} observations onto the standardized 0/25/50/75/100% load
// points, using whichever recorded observation is closest to each target — never invented.
function toLoadPoints(observations, maxCapacity){
  if(!observations || observations.length===0 || !maxCapacity) return [];
  return LOAD_PCTS.map(pct=>{
    const target = (pct/100)*maxCapacity;
    let best=null, bestDist=Infinity;
    for(const o of observations){
      if(o.load===undefined || o.error===undefined) continue;
      const d = Math.abs(o.load-target);
      if(d<bestDist){ bestDist=d; best=o; }
    }
    if(!best) return null;
    // only accept a match if it's reasonably close to the target band (within 15% of capacity)
    if(bestDist > maxCapacity*0.15 && pct!==0) return null;
    return {pct, load:best.load, error:best.error, unit:best.unit, source:best.source, timestamp:best.timestamp};
  }).filter(Boolean);
}
function getFingerprintPoints(db, instrument){
  const history = (instrument.measurementHistory||[]).map(h=>({
    id:h.id, date:h.date, testKey:h.testKey, load:h.load, error:h.error, unit:h.unit,
    source:h.source, envTemp:h.envTemp, envHumidity:h.envHumidity, resultLabel:h.resultLabel,
    isDemo:!!h.isDemo, isCurrent:false, refTestId:h.refTestId||null,
    loadPoints: h.loadPoints || (h.load!==undefined ? [{pct:null, load:h.load, error:h.error, unit:h.unit, source:h.source, timestamp:h.date}] : []),
  }));
  const accTest = db.tests.find(t=>t.instrumentId===instrument.id && t.testKey==="ACC");
  const liveObs = findMaxLoadObs(accTest, instrument);
  let current = null;
  if(liveObs){
    const env = db.envReadings.find(e=>e.testId===accTest.id);
    current = {
      id:`live-${accTest.id}`, date:liveObs.timestamp, testKey:"ACC", load:liveObs.load, error:liveObs.error, unit:liveObs.unit,
      source:liveObs.source, envTemp:env?env.temperature:null, envHumidity:env?env.humidity:null,
      resultLabel:accTest.result||liveObs.result, isDemo:false, isCurrent:true, refTestId:accTest.id,
      loadPoints: toLoadPoints(accTest.observations, instrument.maxCapacity),
    };
  }
  const all = current ? [...history, current] : [...history];
  return all.sort((a,b)=> a.date.localeCompare(b.date));
}
function computeFingerprint(db, instrument, config){
  const cfg = config || db.fingerprintConfig || DEFAULT_FP_CONFIG;
  const points = getFingerprintPoints(db, instrument);
  if(points.length < cfg.minEvaluations){
    return {status:"INSUFFICIENT", points, current:points[points.length-1]||null, previous:null, baselineMean:null, changeAbs:null, e:instrument.e, cfg};
  }
  const current = points[points.length-1];
  const previous = points[points.length-2];
  const baselinePoints = points.slice(0, -1);
  const baselineMean = round(baselinePoints.reduce((s,p)=>s+p.error,0)/baselinePoints.length, 5);
  const baselineMin = round(Math.min(...baselinePoints.map(p=>p.error)),5);
  const baselineMax = round(Math.max(...baselinePoints.map(p=>p.error)),5);
  const changeAbs = round(Math.abs(current.error - previous.error), 5);
  const e = instrument.e;

  // Multi-point (load-range) comparison, where both current and previous have load-point arrays.
  let loadDiffs = [];
  let keyFinding = null;
  if(current.loadPoints?.length>0 && previous.loadPoints?.length>0){
    const prevByPct = Object.fromEntries(previous.loadPoints.filter(p=>p.pct!=null).map(p=>[p.pct,p]));
    loadDiffs = current.loadPoints.filter(p=>p.pct!=null && prevByPct[p.pct]!==undefined).map(p=>({
      pct:p.pct, currentError:p.error, previousError:prevByPct[p.pct].error, diff: round(p.error - prevByPct[p.pct].error, 5),
    }));
    if(loadDiffs.length>0){
      const worst = loadDiffs.reduce((a,b)=> Math.abs(b.diff)>Math.abs(a.diff)?b:a);
      const upper = loadDiffs.filter(d=>d.pct>=50);
      const lower = loadDiffs.filter(d=>d.pct<50);
      const upperAvg = upper.length ? upper.reduce((s,d)=>s+Math.abs(d.diff),0)/upper.length : 0;
      const lowerAvg = lower.length ? lower.reduce((s,d)=>s+Math.abs(d.diff),0)/lower.length : 0;
      if(Math.abs(worst.diff) > e){
        const rangeLabel = upperAvg>=lowerAvg
          ? (upper.length>1 ? `${Math.min(...upper.map(d=>d.pct))}–${Math.max(...upper.map(d=>d.pct))}%` : `${worst.pct}%`)
          : (lower.length>1 ? `${Math.min(...lower.map(d=>d.pct))}–${Math.max(...lower.map(d=>d.pct))}%` : `${worst.pct}%`);
        keyFinding = `Pattern change is most visible at ${rangeLabel} load (largest difference ${worst.diff>=0?"+":""}${worst.diff} ${current.unit} at ${worst.pct}% load).`;
      }
    }
  }

  let status = "STABLE";
  if(changeAbs > cfg.significantThresholdE*e) status = "SIGNIFICANT";
  else if(changeAbs > cfg.changeThresholdE*e) status = "CHANGE";
  return {status, points, current, previous, baselineMean, baselineMin, baselineMax, changeAbs, e, loadDiffs, keyFinding, cfg};
}
const FP_STATUS_META = {
  STABLE:{label:"STABLE", icon:"", color:"var(--seal)", bg:"var(--seal-bg)", border:"var(--seal-line)", desc:"Measurement behavior is consistent with historical observations.", consistency:"High"},
  CHANGE:{label:"PATTERN CHANGE", icon:"", color:"var(--amber)", bg:"var(--amber-bg)", border:"var(--amber-line)", desc:"Current behavior differs from historical observations.", consistency:"Moderate"},
  SIGNIFICANT:{label:"SIGNIFICANT CHANGE", icon:"", color:"var(--rose)", bg:"var(--rose-bg)", border:"var(--rose-line)", desc:"A stronger historical deviation pattern has been detected and requires review.", consistency:"Low"},
  INSUFFICIENT:{label:"INSUFFICIENT DATA", icon:"", color:"var(--ink-faint)", bg:"var(--surface-2)", border:"var(--line-strong)", desc:"Not enough historical evaluations exist to establish a reliable baseline.", consistency:"Not established"},
};

/* ============================== INITIAL LAB DATA ============================== */
function buildInitialDb(){
  const now = Date.now();
  const t = (mins)=> new Date(now - mins*60000).toISOString();

  const instrumentA = {
    id:"INST-A300", nawiId:"NAWI-000142", manufacturer:"Suvidha Scales & Systems Pvt. Ltd.", applicant:"Suvidha Scales & Systems Pvt. Ltd.",
    model:"SS-300B", serialNumber:"SS300B-2026-0142", instrumentType:"Electronic Bench Scale",
    accuracyClass:"III", maxCapacity:300, minCapacity:2, capacityUnit:"kg", e:0.1, nVerification:3000, verificationStage:"initial",
    display:"LCD Digital, 6-digit", firmware:"v4.2.1", dateSubmission:t(60*24*30), laboratory:"Regional Reference Standards Laboratory, Panipat",
    status:"Evaluation Completed", photos:{instrument:true, nameplate:true}, documents:["Calibration Certificate.pdf","Technical Datasheet.pdf"],
    isDemo:false, recordSource:"Prototype laboratory seed record", dataStatus:"Seeded — replace with verified laboratory record", createdAt:t(60*24*30),
    measurementHistory:[
      {id:"MH-A1", date:t(60*24*410), testKey:"ACC", load:300, error:0.03, unit:"kg", source:"manual", envTemp:24.1, envHumidity:50.8, resultLabel:"PASS", isDemo:false, refTestId:null,
        loadPoints:[{pct:0,load:0,error:0,unit:"kg",source:"manual"},{pct:25,load:75,error:0.01,unit:"kg",source:"manual"},{pct:50,load:150,error:-0.02,unit:"kg",source:"manual"},{pct:75,load:225,error:0.03,unit:"kg",source:"manual"},{pct:100,load:300,error:0.03,unit:"kg",source:"manual"}]},
      {id:"MH-A2", date:t(60*24*205), testKey:"ACC", load:300, error:0.01, unit:"kg", source:"iot", envTemp:24.3, envHumidity:51.2, resultLabel:"PASS", isDemo:false, refTestId:null,
        loadPoints:[{pct:0,load:0,error:0,unit:"kg",source:"iot"},{pct:25,load:75,error:0.015,unit:"kg",source:"iot"},{pct:50,load:150,error:-0.01,unit:"kg",source:"iot"},{pct:75,load:225,error:0.02,unit:"kg",source:"iot"},{pct:100,load:300,error:0.01,unit:"kg",source:"iot"}]},
    ],
  };
  const instrumentB = {
    id:"INST-B030", nawiId:"NAWI-000071", manufacturer:"Prakash Weighing Technologies Pvt. Ltd.", applicant:"Prakash Weighing Technologies Pvt. Ltd.",
    model:"PWT-P30", serialNumber:"PWTP30-2026-0071", instrumentType:"Electronic Precision Balance",
    accuracyClass:"III", maxCapacity:30, minCapacity:0.5, capacityUnit:"kg", e:0.005, nVerification:6000, verificationStage:"initial",
    display:"OLED Digital, 7-digit", firmware:"v2.0.6", dateSubmission:t(60*24*6), laboratory:"Regional Reference Standards Laboratory, Panipat",
    status:"Testing In Progress", photos:{instrument:true, nameplate:true}, documents:["Technical Datasheet.pdf"],
    isDemo:false, recordSource:"Prototype laboratory seed record", dataStatus:"Seeded — replace with verified laboratory record", createdAt:t(60*24*6),
    measurementHistory:[
      {id:"MH-B1", date:t(60*24*395), testKey:"ACC", load:30, error:0.002, unit:"kg", source:"manual", envTemp:24.0, envHumidity:50.5, resultLabel:"PASS", isDemo:false, refTestId:null,
        loadPoints:[{pct:0,load:0,error:0,unit:"kg",source:"manual"},{pct:25,load:7.5,error:0.0005,unit:"kg",source:"manual"},{pct:50,load:15,error:0.001,unit:"kg",source:"manual"},{pct:75,load:22.5,error:0.0015,unit:"kg",source:"manual"},{pct:100,load:30,error:0.002,unit:"kg",source:"manual"}]},
      {id:"MH-B2", date:t(60*24*190), testKey:"ACC", load:30, error:0.006, unit:"kg", source:"iot", envTemp:24.4, envHumidity:51.4, resultLabel:"PASS", isDemo:false, refTestId:null,
        loadPoints:[{pct:0,load:0,error:0,unit:"kg",source:"iot"},{pct:25,load:7.5,error:0.001,unit:"kg",source:"iot"},{pct:50,load:15,error:0.002,unit:"kg",source:"iot"},{pct:75,load:22.5,error:0.004,unit:"kg",source:"iot"},{pct:100,load:30,error:0.006,unit:"kg",source:"iot"}]},
    ],
  };

  const rule = activeRule();

  // Instrument A: all tests PASS + approved + report generated
  const a_acc_obs = [0,75,150,225,300].map((load,i)=>{
    const observed = round(load + (load===0?0:[0.02,-0.03,0.04,0.01,-0.02][i]),3);
    const calc = evalAccuracy(load, observed, instrumentA, rule);
    return {id:uid("OBS"), load, observed, unit:"kg", source:"manual", timestamp:t(60*24*20-i), ...calc};
  });
  const a_rep_readings = [150.02,150.01,150.03,150.01,150.02];
  const a_rep_calc = evalRepeatability(a_rep_readings, 150, instrumentA, rule);

  const tests = [
    {id:"TEST-A-ACC", instrumentId:"INST-A300", testKey:"ACC", testName:TEST_DEFS.ACC.name, status:"approved",
      observations:a_acc_obs, ruleVersionId:rule.id, remarks:"All observed indications within MPE across the tested range.",
      technician:"R. Sharma", reviewer:"Dr. N. Verma", approver:"S. Iyer (Authority)", result:"PASS",
      createdAt:t(60*24*20), reviewedAt:t(60*24*18), approvedAt:t(60*24*17)},
    {id:"TEST-A-REP", instrumentId:"INST-A300", testKey:"REP", testName:TEST_DEFS.REP.name, status:"approved",
      observations:a_rep_readings.map((v,i)=>({id:uid("OBS"), load:150, observed:v, unit:"kg", source: i<2?"manual":"iot", timestamp:t(60*24*20-i)})),
      ruleVersionId:rule.id, remarks:"Range within permissible repeatability limit.", repeatCalc:a_rep_calc,
      technician:"R. Sharma", reviewer:"Dr. N. Verma", approver:"S. Iyer (Authority)", result:"PASS",
      createdAt:t(60*24*20), reviewedAt:t(60*24*18), approvedAt:t(60*24*17)},
    {id:"TEST-A-DIS", instrumentId:"INST-A300", testKey:"DIS", testName:TEST_DEFS.DIS.name, status:"approved",
      observations:[{id:uid("OBS"), before:300.0, after:300.1, incrementE:1, changed:true, timestamp:t(60*24*19)}],
      ruleVersionId:rule.id, remarks:"Indication responded to added increment at maximum capacity.",
      technician:"R. Sharma", reviewer:"Dr. N. Verma", approver:"S. Iyer (Authority)", result:"PASS",
      createdAt:t(60*24*19), reviewedAt:t(60*24*18), approvedAt:t(60*24*17)},
    {id:"TEST-A-ECC", instrumentId:"INST-A300", testKey:"ECC", testName:TEST_DEFS.ECC.name, status:"approved",
      observations:[
        {id:uid("OBS"), position:"Center", load:150, observed:150.02, timestamp:t(60*24*19)},
        {id:uid("OBS"), position:"Corner 1", load:150, observed:150.08, timestamp:t(60*24*19)},
        {id:uid("OBS"), position:"Corner 2", load:150, observed:149.95, timestamp:t(60*24*19)},
        {id:uid("OBS"), position:"Corner 3", load:150, observed:150.06, timestamp:t(60*24*19)},
        {id:uid("OBS"), position:"Corner 4", load:150, observed:149.97, timestamp:t(60*24*19)},
      ], ruleVersionId:rule.id, remarks:"Maximum corner deviation within MPE at tested load.",
      technician:"R. Sharma", reviewer:"Dr. N. Verma", approver:"S. Iyer (Authority)", result:"PASS",
      createdAt:t(60*24*19), reviewedAt:t(60*24*18), approvedAt:t(60*24*17)},

    // Instrument B: in progress, one FAIL and one REVIEW
    {id:"TEST-B-ACC", instrumentId:"INST-B030", testKey:"ACC", testName:TEST_DEFS.ACC.name, status:"review",
      observations:[0,7.5,15,22.5,30].map((load,i)=>{
        const observed = round(load + (load===0?0:[0.001,-0.0015,0.002,0.0018,0.041][i]),4);
        const calc = evalAccuracy(load, observed, instrumentB, rule);
        return {id:uid("OBS"), load, observed, unit:"kg", source: i===4?"iot":"manual", timestamp:t(60*8-i), ...calc};
      }), ruleVersionId:rule.id, remarks:"Significant deviation observed at maximum capacity — recommend recalibration check before resubmission.",
      technician:"P. Nair", reviewer:"", approver:"", result:"FAIL", createdAt:t(60*8), reviewedAt:null, approvedAt:null},
    {id:"TEST-B-REP", instrumentId:"INST-B030", testKey:"REP", testName:TEST_DEFS.REP.name, status:"review",
      observations:[15.001,15.001,15.001,15.018,15.002].map((v,i)=>({id:uid("OBS"), load:15, observed:v, unit:"kg", source:"iot", timestamp:t(60*6-i)})),
      ruleVersionId:rule.id, remarks:"Environmental variation flagged during sequence; readings pattern requires reviewer attention.",
      technician:"P. Nair", reviewer:"", approver:"", result:"REVIEW",
      envFlag:true, createdAt:t(60*6), reviewedAt:null, approvedAt:null},
    {id:"TEST-B-DIS", instrumentId:"INST-B030", testKey:"DIS", testName:TEST_DEFS.DIS.name, status:"pending",
      observations:[], ruleVersionId:rule.id, remarks:"", technician:"P. Nair", reviewer:"", approver:"", result:null, createdAt:t(60*5)},
    {id:"TEST-B-ECC", instrumentId:"INST-B030", testKey:"ECC", testName:TEST_DEFS.ECC.name, status:"pending",
      observations:[], ruleVersionId:rule.id, remarks:"", technician:"P. Nair", reviewer:"", approver:"", result:null, createdAt:t(60*5)},
  ];

  const testPlans = [
    {id:"TP-A300", instrumentId:"INST-A300", tests:generateTestPlan(instrumentA), generatedAt:t(60*24*20), ruleVersionId:rule.id},
    {id:"TP-B030", instrumentId:"INST-B030", tests:generateTestPlan(instrumentB), generatedAt:t(60*8), ruleVersionId:rule.id},
  ];

  const envReadings = [
    ...[0,1,2,3,4].map(i=>({
      id:uid("ENV"), instrumentId:"INST-A300", testId:"TEST-A-ACC",
      timestamp:t(60*24*20-i*3), temperature:[24.1,24.2,24.2,24.3,24.1][i],
      humidity:[51.0,51.2,51.1,51.3,51.0][i], sensorId:"TEMP-01/HUM-01"
    })),
    ...[0,1,2].map(i=>({
      id:uid("ENV"), instrumentId:"INST-B030", testId:"TEST-B-REP",
      timestamp:t(60*6-i*2), temperature:[29.8,24.4,24.5][i],
      humidity:[52.0,52.2,52.1][i], sensorId:"TEMP-01/HUM-01"
    })),
  ];

  const reports = [
    {id:"RPT-1", reportNumber:"NAWI-2026-00087", instrumentId:"INST-A300", testIds:["TEST-A-ACC","TEST-A-REP","TEST-A-DIS","TEST-A-ECC"],
      generatedAt:t(60*24*17), status:"Issued", ruleVersionId:rule.id, integrityHash:reportContentFingerprint(instrumentA, tests.filter(x=>x.instrumentId==="INST-A300"), rule.id),
      technician:"R. Sharma", reviewer:"Dr. N. Verma", approver:"S. Iyer (Authority)"},
  ];

  const auditLog = [
    {id:uid("AUD"), timestamp:t(60*24*30), user:"R. Sharma (Technician)", action:"Instrument registered", entity:"Instrument", entityId:"INST-A300", prevValue:"—", newValue:"SS-300B / SS300B-2026-0142", reason:"New submission"},
    {id:uid("AUD"), timestamp:t(60*24*20), user:"R. Sharma (Technician)", action:"Test plan generated", entity:"TestPlan", entityId:"TP-A300", prevValue:"—", newValue:"4 applicable tests", reason:"Auto-generated from instrument parameters"},
    {id:uid("AUD"), timestamp:t(60*24*20), user:"R. Sharma (Technician)", action:"Observation recorded", entity:"Test", entityId:"TEST-A-ACC", prevValue:"—", newValue:"5 accuracy readings", reason:"Routine test execution"},
    {id:uid("AUD"), timestamp:t(60*24*18), user:"Dr. N. Verma (Examiner)", action:"Test reviewed", entity:"Test", entityId:"TEST-A-ACC", prevValue:"review", newValue:"approved", reason:"Data verified against MPE table"},
    {id:uid("AUD"), timestamp:t(60*24*17), user:"S. Iyer (Authority)", action:"Report approved & issued", entity:"Report", entityId:"RPT-1", prevValue:"draft", newValue:"Issued", reason:"All applicable tests approved"},
    {id:uid("AUD"), timestamp:t(60*8), user:"P. Nair (Technician)", action:"Instrument registered", entity:"Instrument", entityId:"INST-B030", prevValue:"—", newValue:"PWT-P30 / PWTP30-2026-0071", reason:"New submission"},
    {id:uid("AUD"), timestamp:t(60*8), user:"P. Nair (Technician)", action:"Observation recorded", entity:"Test", entityId:"TEST-B-ACC", prevValue:"—", newValue:"5 accuracy readings", reason:"Routine test execution; deviation at max load"},
    {id:uid("AUD"), timestamp:t(60*8), user:"P. Nair (Technician)", action:"Test submitted for review", entity:"Test", entityId:"TEST-B-ACC", prevValue:"in_progress", newValue:"review", reason:"FAIL result — escalated to Examiner"},
    {id:uid("AUD"), timestamp:t(60*6), user:"Smart IoT Test Bench", action:"Environmental anomaly flagged", entity:"Test", entityId:"TEST-B-REP", prevValue:"24.3°C", newValue:"29.8°C", reason:"Sudden ambient variation during repeatability sequence"},
  ];

  const iotDevices = [
    {id:"DEV-SCALE", name:"NAWI-TEST-001 (Weighing Device)", type:"weighing", status:"connected", value:0, unit:"kg"},
    {id:"DEV-TEMP", name:"Temperature Sensor TEMP-01", type:"temperature", status:"connected", value:24.4, unit:"°C"},
    {id:"DEV-HUM", name:"Humidity Sensor HUM-01", type:"humidity", status:"connected", value:51.6, unit:"%RH"},
    {id:"DEV-GW", name:"IoT Data Gateway", type:"gateway", status:"connected", value:null, unit:""},
  ];

  const referenceEquipment = [
    {id:"REF-EQ-1", equipmentId:"RSW-500-A", equipmentType:"Reference Standard Weight Set (E2 class, 1g–500kg)", capacity:"500 kg", resolution:"0.1 g",
      certNumber:"NPL/CAL/2025/4471", calibrationDate:"2025-04-10", validUntil:"2027-04-09", traceability:"National Physical Laboratory (NPL), India — traceable to national mass standard"},
    {id:"REF-EQ-2", equipmentId:"RSW-30-B", equipmentType:"Reference Standard Weight Set (F1 class, 1g–30kg)", capacity:"30 kg", resolution:"1 mg",
      certNumber:"NPL/CAL/2024/1187", calibrationDate:"2024-02-15", validUntil:"2026-02-14", traceability:"National Physical Laboratory (NPL), India — traceable to national mass standard"},
  ];

  const users = [
    {id:"U1", name:"R. Sharma", role:"Technician", email:"r.sharma@nawi-lab.gov.in", phone:"+91 98110 22345", designation:"Laboratory Technician"},
    {id:"U2", name:"P. Nair", role:"Technician", email:"p.nair@nawi-lab.gov.in", phone:"+91 98110 22346", designation:"Laboratory Technician"},
    {id:"U3", name:"Dr. N. Verma", role:"Examiner", email:"n.verma@nawi-lab.gov.in", phone:"+91 98110 22347", designation:"Senior Examiner"},
    {id:"U4", name:"S. Iyer", role:"Authority", email:"s.iyer@nawi-lab.gov.in", phone:"+91 98110 22348", designation:"Approving Authority"},
    {id:"U5", name:"A. Admin", role:"Admin", email:"admin@nawi-lab.gov.in", phone:"+91 98110 22349", designation:"System Administrator"},
    {id:"U6", name:"Guest Viewer", role:"Viewer", email:"viewer@nawi-lab.gov.in", phone:"—", designation:"Observer"},
  ];

  return { instruments:[instrumentA, instrumentB], testPlans, tests, envReadings, reports, auditLog, iotDevices, users, referenceEquipment, ruleVersions: RULE_VERSIONS, referenceSources: OFFICIAL_REFERENCE_SOURCES, fingerprintConfig:{...DEFAULT_FP_CONFIG} };
}

/* ============================== SMALL UI PRIMITIVES ============================== */
const ResultBadge = ({result}) => {
  if(!result) return <span className="n-badge n-badge-neutral">PENDING</span>;
  if(result==="PASS") return <span className="n-badge n-badge-pass"><CheckCircle2 size={12}/>PASS</span>;
  if(result==="FAIL") return <span className="n-badge n-badge-fail"><XCircle size={12}/>FAIL</span>;
  return <span className="n-badge n-badge-review"><AlertTriangle size={12}/>REVIEW</span>;
};
const StatusBadge = ({status}) => {
  const map = {
    pending:{cls:"n-badge-neutral", label:"PENDING"}, in_progress:{cls:"n-badge-navy", label:"IN PROGRESS"},
    review:{cls:"n-badge-review", label:"AWAITING REVIEW"}, approved:{cls:"n-badge-pass", label:"APPROVED"},
    rejected:{cls:"n-badge-fail", label:"REJECTED"}, applicable:{cls:"n-badge-neutral", label:"APPLICABLE"},
    "not-applicable":{cls:"n-badge-neutral", label:"N/A"},
  };
  const m = map[status] || {cls:"n-badge-neutral", label:status};
  return <span className={`n-badge ${m.cls}`}>{m.label}</span>;
};
const FlagIcon = ({level}) => level==="ok" ? <CheckCircle2 size={13} className="n-flag-valid"/> : level==="warn" ? <AlertTriangle size={13} className="n-flag-warn"/> : <XCircle size={13} className="n-flag-bad"/>;

/* ============================== ROLE-BASED ACCESS CONTROL ==============================
   Single source of truth for what each role may do (mirrors the Can/Cannot lists in the
   problem statement). Every state-changing handler in the app calls can(role, action) BEFORE
   it mutates — not just the JSX that decides which button to render — so an action taken
   through any path is checked against the actor's role, not only hidden from the menu.
   The authenticated server session is the identity source. The frontend matrix controls the
   visible workflow while the Phase-6 backend independently enforces write permissions. */
const PERMISSIONS = {
  Technician: ["instrument:create","instrument:edit-draft","test:record","test:mark-na","test:submit","evidence:upload"],
  Examiner:   ["test:review"],
  Authority:  ["test:review","report:approve-final"],
  Admin:      ["instrument:create","instrument:edit-draft","test:record","test:mark-na","test:submit","evidence:upload","test:review","report:approve-final","rule:manage","fingerprint:configure","user:manage"],
  Viewer:     [],
};
function can(role, action){ return (PERMISSIONS[role]||[]).includes(action); }

/* UI navigation policy; server-side RBAC remains authoritative for writes. */
const ROLE_NAV = {
  Technician: new Set(["dashboard","instruments","tests","intelligence","reports","repository","verify","iot","analytics","audit","references","documentation","security","profile"]),
  Examiner: new Set(["dashboard","instruments","tests","reports","repository","verify","iot","analytics","audit","references","documentation","security","profile"]),
  Authority: new Set(["dashboard","instruments","tests","reports","repository","verify","analytics","audit","references","documentation","security","profile"]),
  Admin: new Set(["dashboard","instruments","tests","intelligence","reports","repository","verify","iot","analytics","audit","users","rules","references","documentation","security","settings","profile"]),
  Viewer: new Set(["dashboard","instruments","tests","reports","repository","verify","analytics","references","documentation","security","profile"])
};
function canNavigate(role, page){ return ROLE_NAV[role]?.has(page) ?? false; }
// True once a test's measurements are locked from further Technician entry — submitted for
// review, sent back for correction is the only way back to an editable state, per item 17.
const isLockedForEntry = (test) => !["pending","in_progress"].includes(test.status);

/* ============================== EXPLAIN CALCULATION / WHY PASS-FAIL ==============================
   Renders the full OIML calculation trail for one evaluated point (accuracy or repeatability),
   built only from fields already stored on the observation/repeatCalc object at the time the rule
   engine ran. Doubles as "Why PASS?" / "Why FAIL?" (item 12) since the same breakdown proves the
   result. No text here is static — every value is read from the calc object passed in. */
function ExplainCalculation({calc, instrument, kind="accuracy"}){
  const [open, setOpen] = useState(false);
  if(!calc) return null;
  const unit = instrument?.capacityUnit || "";
  const stageLabel = calc.verificationStage==="in-service" ? "In-Service Verification" : "Initial Verification";
  return (
    <div style={{marginTop:open?6:0}}>
      <button className="n-btn n-btn-sm" onClick={()=>setOpen(o=>!o)} style={{fontSize:11}}>
        {open ? "Hide" : (calc.result==="FAIL" ? "Why FAIL?" : calc.result==="PASS" ? "Why PASS?" : "Explain Calculation")}
      </button>
      {open && (
        <div className="f-mono" style={{marginTop:6, padding:"10px 12px", background:"var(--surface-2)", borderRadius:8, fontSize:12, lineHeight:1.75}}>
          {kind==="accuracy" ? (
            <>
              <div>Reference load: {calc.referenceValue} {unit}</div>
              <div>Observed indication: {calc.observedValue} {unit}</div>
              <div>Error = Observed − Reference = {calc.error>=0?"+":""}{calc.error} {unit}</div>
              <div>Verification scale interval (e): {calc.e} {unit}</div>
              <div>Load in e: {calc.loadInE}e (n = Max/e = {calc.n})</div>
              <div>Accuracy class: {calc.accuracyClass}</div>
              <div>Verification stage: {stageLabel}</div>
              <div>Applicable MPE: ±{calc.mpeE}e</div>
              <div>MPE value: ±{calc.mpeVal} {unit}</div>
              <div>Absolute error: {Math.abs(calc.error)} {unit}</div>
              <div style={{marginTop:4}}>Comparison: {Math.abs(calc.error)} {"<="} {calc.mpeVal} → {Math.abs(calc.error)<=calc.mpeVal ? "true" : "false"}</div>
              <div style={{fontWeight:700, marginTop:2}}>Result: {calc.result}</div>
            </>
          ) : (
            <>
              <div>Test load: {calc.testLoad} {unit}</div>
              <div>Readings — min: {calc.min} {unit}, max: {calc.max} {unit}</div>
              <div>Range = Max − Min = {calc.range} {unit}</div>
              <div>Verification scale interval (e): {calc.e} {unit}</div>
              <div>Load in e: {calc.loadInE}e</div>
              <div>Accuracy class: {calc.accuracyClass}</div>
              <div>Verification stage: {stageLabel}</div>
              <div>Applicable MPE: ±{calc.mpeE}e</div>
              <div>Allowed range: ±{calc.allowed} {unit}</div>
              <div style={{marginTop:4}}>Comparison: {calc.range} {"<="} {calc.allowed} → {calc.range<=calc.allowed ? "true" : "false"}</div>
              <div style={{fontWeight:700, marginTop:2}}>Result: {calc.result}</div>
            </>
          )}
        </div>
      )}
    </div>
  );
}

/* ============================== QR VERIFICATION HELPERS ==============================
   QR codes carry only a non-sensitive report identifier. The helper is intentionally
   defined centrally so every on-screen/PDF QR uses the same value. */
function makeReportQrToken(reportId){
  const id=String(reportId||"").trim().toUpperCase();
  return id ? `METRA:R:${id}` : "METRA:R:UNKNOWN";
}
function parseMetraQrToken(raw){
  const value=String(raw||"").trim();
  const direct=value.match(/^METRA:R:([A-Z0-9_-]+)$/i);
  if(direct) return {reportId:direct[1].toUpperCase()};
  try{
    const u=new URL(value);
    const q=u.searchParams.get("report") || u.searchParams.get("reportId");
    if(q) return {reportId:q.trim().toUpperCase()};
  }catch{}
  return null;
}

/* ============================== QR CODE ==============================
   Uses qrcode.react (standard ISO/IEC 18004 encoder) so printed/PDF QR codes scan on phones.
   Payload stays non-sensitive: METRA:R:<REPORT_ID>. */
function QrCode({value, size=100}){
  return (
    <div style={{background:"#fff", padding:4, display:"inline-block", lineHeight:0}}>
      <QRCodeSVG value={value || "NAWI"} size={size-8} level="M" includeMargin={false}/>
    </div>
  );
}
// Rasterise the same QR to a PNG data URL for the PDF export.
function qrPngDataUrl(value, px=320){
  return new Promise((resolve, reject)=>{
    const svg=renderToStaticMarkup(<QRCodeSVG value={value||"NAWI"} size={px} level="M" includeMargin={true} bgColor="#ffffff" fgColor="#111111"/>);
    const img=new Image();
    img.onload=()=>{ const c=document.createElement("canvas"); c.width=c.height=px; const ctx=c.getContext("2d"); ctx.fillStyle="#fff"; ctx.fillRect(0,0,px,px); ctx.drawImage(img,0,0,px,px); resolve(c.toDataURL("image/png")); };
    img.onerror=reject;
    img.src="data:image/svg+xml;charset=utf-8,"+encodeURIComponent(svg);
  });
}

function Seal({status="COMPLIANT", size=76}){
  const color = status==="COMPLIANT" ? "var(--seal)" : status==="REVIEW REQUIRED" ? "var(--amber)" : status==="NON-COMPLIANT" ? "var(--rose)" : "var(--ink-faint)";
  return (
    <svg width={size} height={size} viewBox="0 0 100 100">
      <circle cx="50" cy="50" r="46" className="n-stamp-ring" stroke={color}/>
      <circle cx="50" cy="50" r="39" className="n-stamp-ring" stroke={color} strokeDasharray="2 3"/>
      <text x="50" y="42" textAnchor="middle" fontSize="8" fontFamily="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif" fontWeight="700" fill={color}>NAWI</text>
      <text x="50" y="54" textAnchor="middle" fontSize="6.2" fontFamily="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif" fontWeight="700" fill={color}>{status.split(" ")[0]}</text>
      <text x="50" y="63" textAnchor="middle" fontSize="6.2" fontFamily="-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif" fontWeight="700" fill={color}>{status.split(" ")[1]||""}</text>
    </svg>
  );
}

function Modal({title, onClose, children, width=560}){
  return (
    <div className="no-print" style={{position:"fixed", inset:0, background:"rgba(17,24,39,.45)", backdropFilter:"blur(6px)", WebkitBackdropFilter:"blur(6px)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:100}} onClick={onClose}>
      <div className="n-panel" style={{width, maxWidth:"92vw", maxHeight:"88vh", overflow:"auto", borderRadius:28, boxShadow:"0 25px 60px -15px rgba(0,0,0,.2), 0 1px 3px rgba(0,0,0,.05)"}} onClick={e=>e.stopPropagation()}>
        <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", padding:"18px 24px", borderBottom:"1px solid var(--line)"}}>
          <div className="f-display" style={{fontWeight:800, fontSize:17, color:"var(--ink)", letterSpacing:"-.02em"}}>{title}</div>
          <button className="n-btn n-btn-ghost n-btn-sm" onClick={onClose} style={{width:32, height:32, padding:0, borderRadius:"50%"}}><X size={15}/></button>
        </div>
        <div style={{padding:"24px 28px"}}>{children}</div>
      </div>
    </div>
  );
}

/* ============================== LEGAL / PRODUCT DISCLOSURE ============================== */
function LegalPage({kind, nav}){
  const isPrivacy = kind==="privacy";
  return (
    <div className="nawi-page">
      <PageHeader
        eyebrow="Product documents"
        title={isPrivacy ? "Privacy Policy" : "Terms of Service"}
        desc={isPrivacy ? "Draft privacy notice for review before production deployment." : "Draft terms for review before production deployment."}
        right={<button className="n-btn" onClick={()=>nav("dashboard")}><ChevronLeft size={14}/> Back to dashboard</button>}
      />
      <div className="n-panel" style={{padding:22, maxWidth:900}}>
        <div className="n-badge n-badge-review" style={{marginBottom:14}}>DRAFT FOR REVIEW</div>
        {isPrivacy ? (
          <>
            <h3 style={{fontSize:18,margin:"0 0 8px"}}>Privacy and data handling</h3>
            <p style={{fontSize:13,lineHeight:1.7,color:"var(--ink-dim)",marginTop:0}}>
              Authenticated sessions use the METRA backend repository for active records; IndexedDB is only an offline initialization fallback. The prototype still does not establish a production retention period, hosting provider, encryption policy, subprocessors, user-consent flow, or data-subject request process. Those details must be reviewed and completed before real laboratory or personal data is used.
            </p>
            <div className="n-hr" style={{margin:"18px 0"}}/>
            <h4 style={{fontSize:13,margin:"0 0 7px"}}>Data categories represented in the prototype</h4>
            <p style={{fontSize:13,lineHeight:1.7,color:"var(--ink-dim)",margin:0}}>
              Instrument identifiers, test observations, environmental readings, audit events, report metadata, equipment calibration records, and user profile fields are represented in the seeded prototype dataset. Production collection and access rules require a documented review.
            </p>
          </>
        ) : (
          <>
            <h3 style={{fontSize:18,margin:"0 0 8px"}}>Prototype terms</h3>
            <p style={{fontSize:13,lineHeight:1.7,color:"var(--ink-dim)",marginTop:0}}>
              METRA is presented here as a prototype and simulation environment. The application itself already identifies that it is not an OIML-certified system. Production terms covering service availability, permitted use, intellectual property, warranties, liability, support, account responsibilities, and governing law have not been supplied and must be approved by the business owner and legal counsel.
            </p>
            <div className="n-hr" style={{margin:"18px 0"}}/>
            <h4 style={{fontSize:13,margin:"0 0 7px"}}>Compliance notice</h4>
            <p style={{fontSize:13,lineHeight:1.7,color:"var(--ink-dim)",margin:0}}>
              The software can calculate and explain configured OIML R-76 rules, but its outputs should be reviewed against the exact applicable edition and laboratory procedures before certification or regulatory use.
            </p>
          </>
        )}
      </div>
    </div>
  );
}

/* ============================== SERVER PERSISTENCE ============================== */
// IndexedDB remains available only as a non-authoritative fallback for an offline prototype.
// Phase 6 uses the authenticated backend as the source of truth whenever a session exists.
const NAWI_DB_STORE = "smart-nawi-repository";
const NAWI_DB_KEY = "repository-v2-phase1";
function openNawiDb(){
  return new Promise((resolve,reject)=>{
    if(typeof indexedDB === "undefined"){ reject(new Error("IndexedDB unavailable")); return; }
    const req = indexedDB.open(NAWI_DB_STORE,1);
    req.onupgradeneeded = () => { const db=req.result; if(!db.objectStoreNames.contains("state")) db.createObjectStore("state"); };
    req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error||new Error("Could not open repository"));
  });
}
async function loadNawiRepository(){
  const db=await openNawiDb();
  return new Promise((resolve,reject)=>{ const tx=db.transaction("state","readonly"); const req=tx.objectStore("state").get(NAWI_DB_KEY); req.onsuccess=()=>resolve(req.result||null); req.onerror=()=>reject(req.error||new Error("Could not load repository")); });
}
async function saveNawiRepository(state){
  const db=await openNawiDb();
  return new Promise((resolve,reject)=>{ const tx=db.transaction("state","readwrite"); tx.objectStore("state").put(state,NAWI_DB_KEY); tx.oncomplete=()=>resolve(true); tx.onerror=()=>reject(tx.error||new Error("Could not save repository")); });
}

/* ============================== DECORATIVE ORBIT ============================== */
function DecorativeOrbit({ size = 320, style = {} }) {
  return (
    <div aria-hidden="true" style={{ position: "relative", width: size, height: size, pointerEvents: "none", userSelect: "none", ...style }}>
      {/* Outer Orbit (380px nominal) */}
      <div style={{
        position: "absolute", inset: 0, margin: "auto",
        width: size, height: size, borderRadius: "50%",
        border: "1px dashed rgba(17, 24, 39, 0.12)",
        animation: "orbitRotate 45s linear infinite"
      }}>
        <span style={{ position: "absolute", top: -4, left: "50%", width: 8, height: 8, borderRadius: "50%", background: "#111827", boxShadow: "0 0 8px rgba(0,0,0,0.15)" }} />
        <span style={{ position: "absolute", bottom: "16%", right: "12%", width: 6, height: 6, borderRadius: "50%", background: "var(--seal)" }} />
      </div>
      {/* Middle Orbit (260px nominal) */}
      <div style={{
        position: "absolute", inset: 0, margin: "auto",
        width: Math.round(size * 0.68), height: Math.round(size * 0.68), borderRadius: "50%",
        border: "1px dashed rgba(17, 24, 39, 0.14)",
        animation: "orbitRotateReverse 35s linear infinite"
      }}>
        <span style={{ position: "absolute", top: "20%", left: -3, width: 7, height: 7, borderRadius: "50%", background: "var(--amber)" }} />
        <span style={{ position: "absolute", bottom: "24%", right: "6%", width: 5, height: 5, borderRadius: "50%", background: "var(--ink-dim)" }} />
      </div>
      {/* Inner Orbit (140px nominal) */}
      <div style={{
        position: "absolute", inset: 0, margin: "auto",
        width: Math.round(size * 0.38), height: Math.round(size * 0.38), borderRadius: "50%",
        border: "1px dashed rgba(17, 24, 39, 0.16)",
        animation: "orbitRotate 25s linear infinite"
      }}>
        <span style={{ position: "absolute", top: -3, right: "24%", width: 6, height: 6, borderRadius: "50%", background: "#111827" }} />
      </div>
      {/* Center Icon */}
      <div style={{
        position: "absolute", inset: 0, margin: "auto",
        width: Math.round(size * 0.18), height: Math.round(size * 0.18), borderRadius: "50%",
        background: "#ffffff", border: "1px solid var(--line)",
        boxShadow: "0 4px 14px rgba(0,0,0,0.06)",
        display: "flex", alignItems: "center", justifyContent: "center",
        fontSize: Math.round(size * 0.08), fontWeight: 800, color: "var(--ink)"
      }}>
        M
      </div>
    </div>
  );
}

/* ============================== APP ============================== */
function LoginPage({ onLogin, loading, error }) {
  const [email, setEmail] = React.useState("");
  const [password, setPassword] = React.useState("");

  const submit = (event) => {
    event.preventDefault();
    if (!email.trim() || !password) return;
    onLogin(email.trim(), password);
  };

  return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"#f7f8f9",padding:24,boxSizing:"border-box",position:"relative",overflow:"hidden"}}>
      {/* Decorative ambient orbits in background */}
      <div style={{position:"absolute", right:"-60px", top:"-60px", pointerEvents:"none", opacity:0.7}}>
        <DecorativeOrbit size={380}/>
      </div>
      <div style={{position:"absolute", left:"-100px", bottom:"-100px", pointerEvents:"none", opacity:0.5}}>
        <DecorativeOrbit size={300}/>
      </div>

      <div style={{width:440,maxWidth:"100%",background:"#ffffff",borderRadius:28,padding:"40px 36px",boxShadow:"0 20px 60px -15px rgba(0,0,0,0.07), 0 1px 3px rgba(0,0,0,0.03)",border:"1px solid #e5e7eb",boxSizing:"border-box",position:"relative",zIndex:2}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:10,height:10,borderRadius:"50%",background:"#111827"}}/>
          <div style={{fontWeight:800,fontSize:24,color:"#111827",letterSpacing:"-.03em"}}>METRA</div>
        </div>
        <div style={{marginTop:6,color:"#9ca3af",fontSize:10,letterSpacing:".1em",fontWeight:700}}>MEASUREMENT FINGERPRINT PLATFORM</div>
        <h2 style={{margin:"26px 0 6px",color:"#111827",fontSize:22,fontWeight:750,letterSpacing:"-.02em"}}>Sign in</h2>
        <p style={{margin:"0 0 24px",color:"#6b7280",fontSize:14,lineHeight:1.5}}>Sign in with a registered account.</p>
        <form onSubmit={submit}>
          <label style={{display:"block",fontSize:12.5,fontWeight:650,color:"#374151",marginBottom:7}}>Email</label>
          <input type="email" value={email} onChange={e=>setEmail(e.target.value)} autoComplete="username" placeholder="name@nawi-lab.gov.in" disabled={loading}
            style={{width:"100%",boxSizing:"border-box",padding:"11px 14px",border:"1px solid #e5e7eb",borderRadius:14,marginBottom:16,fontSize:14,background:"#fff",color:"#111827",outline:"none",transition:"all .2s"}} />
          <label style={{display:"block",fontSize:12.5,fontWeight:650,color:"#374151",marginBottom:7}}>Password</label>
          <input type="password" value={password} onChange={e=>setPassword(e.target.value)} autoComplete="current-password" placeholder="Password" disabled={loading}
            style={{width:"100%",boxSizing:"border-box",padding:"11px 14px",border:"1px solid #e5e7eb",borderRadius:14,marginBottom:18,fontSize:14,background:"#fff",color:"#111827",outline:"none",transition:"all .2s"}} />
          {error && <div style={{background:"#fef2f2",border:"1px solid #fecaca",color:"#dc2626",padding:"11px 14px",borderRadius:12,fontSize:13,marginBottom:18}}>{error}</div>}
          <button type="submit" disabled={loading || !email.trim() || !password}
            style={{width:"100%",padding:"12px 18px",border:0,borderRadius:9999,background:loading?"#9ca3af":"#111827",color:"#ffffff",fontSize:14,fontWeight:650,cursor:loading?"default":"pointer",boxShadow:loading?"none":"0 4px 14px rgba(17,24,39,0.18)",transition:"all .2s"}}>
            {loading ? "Signing in…" : "Sign in"}
          </button>
        </form>
        <div style={{marginTop:22,paddingTop:18,borderTop:"1px solid #e5e7eb",fontSize:11.5,color:"#6b7280",lineHeight:1.6}}>
          Passwords are verified against the server's stored account records. See README.md for the demo account list.
        </div>
      </div>
    </div>
  );
}

export default function App(){
  const [db, setDb] = useState(buildInitialDb);
  const [repositoryReady, setRepositoryReady] = useState(false);
  const [repositoryStatus, setRepositoryStatus] = useState("Connecting to secure repository…");
  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState("");
  const [page, setPage] = useState("dashboard");
  const [selectedInstrument, setSelectedInstrument] = useState(null);
  const [selectedTest, setSelectedTest] = useState(null);
  const [selectedReport, setSelectedReport] = useState(null);
  const [testsFilter, setTestsFilter] = useState("all");
  const [role, setRole] = useState("Technician");
  const [currentUserName, setCurrentUserName] = useState("R. Sharma");
  const [syncStatus, setSyncStatus] = useState("synced");
  const [assistantOpen, setAssistantOpen] = useState(false);
  const [toast, setToast] = useState(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [searchFocused, setSearchFocused] = useState(false);
  const [searchHighlight, setSearchHighlight] = useState(-1);

  useEffect(()=>{
    let cancelled=false;
    naviApi("/api/auth/me").then(async data=>{
      if(cancelled) return;
      setSession(data.user); setCurrentUserName(data.user.name); setRole(data.user.role);
      try{
        const remote=await naviApi("/api/repository");
        if(remote?.repository?.instruments?.length){ setDb(remote.repository); setRepositoryStatus("Secure server repository connected"); }
        else {
          const local=await loadNawiRepository().catch(()=>null);
          const initial=local?.instruments?.length?local:buildInitialDb();
          setDb(initial);
          await naviApi("/api/repository",{method:"PUT",body:JSON.stringify({repository:initial})});
          setRepositoryStatus("Secure repository initialized");
        }
      }catch(err){ setRepositoryStatus(`Repository error: ${err.message}`); }
      setRepositoryReady(true); setAuthLoading(false);
    }).catch(()=>{ if(!cancelled){ setSession(null); setRepositoryReady(true); setAuthLoading(false); setRepositoryStatus("Sign in required"); } });
    return ()=>{cancelled=true;};
  },[]);

  useEffect(()=>{
    if(!repositoryReady || !session) return;
    const tm=setTimeout(()=>{
      naviApi("/api/repository",{method:"PUT",body:JSON.stringify({repository:db})})
        .then(()=>{setRepositoryStatus("Secure server repository saved"); setSyncStatus("synced");})
        .catch(err=>{setRepositoryStatus(`Server save failed: ${err.message}`); setSyncStatus("failed");});
    },350);
    return ()=>clearTimeout(tm);
  },[db,repositoryReady,session]);

  const doLogin = async (email,password) => {
    setLoginLoading(true); setLoginError("");
    try{
      const data=await naviApi("/api/auth/login",{method:"POST",body:JSON.stringify({email,password})});
      setSession(data.user); setCurrentUserName(data.user.name); setRole(data.user.role);
      const remote=await naviApi("/api/repository");
      setDb(remote.repository?.instruments?.length?remote.repository:buildInitialDb());
      setRepositoryReady(true); setRepositoryStatus("Secure server repository connected");
    }catch(err){ setLoginError(err.message); }
    finally{ setLoginLoading(false); }
  };
  const doLogout = async () => {
    try{ await naviApi("/api/auth/logout",{method:"POST",body:"{}"}); }catch{}
    setSession(null); setRepositoryReady(false); setRepositoryStatus("Signed out");
  };

  useEffect(()=>{ if(!toast) return; const tm=setTimeout(()=>setToast(null),3200); return ()=>clearTimeout(tm); },[toast]);
  const notify = (msg) => setToast(msg);

  const logAudit = (draft, action, entity, entityId, prevValue, newValue, reason) => {
    draft.auditLog = [{id:uid("AUD"), timestamp:nowISO(), user:`${currentUserName} (${role})`, action, entity, entityId, prevValue:String(prevValue), newValue:String(newValue), reason}, ...draft.auditLog];
    return draft;
  };
  const mutate = (fn) => setDb(prev => { const draft = JSON.parse(JSON.stringify(prev)); fn(draft); return draft; });

  const nav = (p, extra={}) => {
    if(!canNavigate(role,p)){ notify(`This workspace is not available to the ${role} role.`); return; }
    setPage(p);
    if(extra.instrumentId!==undefined) setSelectedInstrument(extra.instrumentId);
    if(extra.testId!==undefined) setSelectedTest(extra.testId);
    if(extra.reportId!==undefined) setSelectedReport(extra.reportId);
    if(p==="tests") setTestsFilter(extra.testsFilter || "all");
    if(p==="reports" && extra.reportId===undefined) setSelectedReport(null);
    if(p!=="tests" && extra.testId===undefined) setSelectedTest(null);
  };

  const searchResults = useMemo(()=>{
    const q = searchQuery.trim().toLowerCase();
    if(q.length<2) return [];
    const out = [];
    db.instruments.forEach(i=>{
      const haystack = `${i.manufacturer} ${i.model} ${i.serialNumber} ${i.id} ${i.nawiId||""}`.toLowerCase();
      if(haystack.includes(q)){
        const exact = [i.id, i.nawiId, i.serialNumber].filter(Boolean).some(v=>v.toLowerCase()===q);
        out.push({type:"Instrument", label:`${i.model} — ${i.serialNumber}`, sub:`${i.nawiId||""} · ${i.manufacturer}`, exact,
          action:()=>{ nav("instruments",{instrumentId:i.id}); notify(`Opened ${i.nawiId||i.model} — Digital Passport`); }});
      }
    });
    db.tests.forEach(t=>{
      const inst = db.instruments.find(i=>i.id===t.instrumentId);
      const haystack = `${t.testName} ${t.id} ${inst?.model||""}`.toLowerCase();
      if(haystack.includes(q)){
        const exact = t.id.toLowerCase()===q;
        out.push({type:"Test", label:t.testName, sub:`${inst?.model||"—"} — ${t.id}`, exact,
          action:()=>{ nav("tests",{testId:t.id}); notify(`Opened test ${t.id}`); }});
      }
    });
    db.reports.forEach(r=>{
      const inst = db.instruments.find(i=>i.id===r.instrumentId);
      const haystack = `${r.reportNumber} ${inst?.model||""}`.toLowerCase();
      if(haystack.includes(q)){
        const exact = r.reportNumber.toLowerCase()===q;
        out.push({type:"Report", label:r.reportNumber, sub:inst?.model||"—", exact,
          action:()=>{ nav("reports",{instrumentId:r.instrumentId, reportId:r.id}); notify(`Opened report ${r.reportNumber}`); }});
      }
    });
    // exact matches first, then shorter/closer labels
    out.sort((a,b)=> (b.exact-a.exact) || a.label.length-b.label.length);
    return out.slice(0,8);
  },[searchQuery, db]);

  useEffect(()=>{ setSearchHighlight(-1); },[searchQuery]);

  const runSearch = (idx) => {
    if(searchResults.length===0){ notify("No matching instrument, test or report found."); return; }
    const target = idx!=null && searchResults[idx] ? searchResults[idx] : searchResults[0];
    target.action();
    setSearchQuery(""); setSearchFocused(false); setSearchHighlight(-1);
  };
  const searchKeyDown = (e) => {
    if(e.key==="ArrowDown"){ e.preventDefault(); setSearchHighlight(h=>Math.min(h+1, searchResults.length-1)); }
    else if(e.key==="ArrowUp"){ e.preventDefault(); setSearchHighlight(h=>Math.max(h-1, 0)); }
    else if(e.key==="Enter"){ e.preventDefault(); runSearch(searchHighlight>=0?searchHighlight:null); }
    else if(e.key==="Escape"){ setSearchQuery(""); setSearchFocused(false); }
  };

  const NAV_ITEMS_PRIMARY = [
    {key:"dashboard", label:"Dashboard", icon:Home},
    {key:"instruments", label:"Instruments", icon:Package},
    {key:"tests", label:"Tests", icon:ClipboardList},
    {key:"intelligence", label:"Measurement Intelligence", icon:FlaskConical, badge:""},
    {key:"reports", label:"Reports", icon:FileText},
    {key:"repository", label:"Repository", icon:Database},
    {key:"verify", label:"Verify QR", icon:Camera},
  ];
  const NAV_ITEMS_SECONDARY = [
    {key:"iot", label:"IoT Test Bench", icon:Cpu},
    {key:"analytics", label:"Analytics", icon:BarChart3},
    {key:"audit", label:"Audit Trail", icon:ScrollText},
    {key:"users", label:"Users", icon:Users},
    {key:"rules", label:"OIML Rules", icon:GitBranch},
    {key:"references", label:"Official References", icon:Library},
    {key:"documentation", label:"Technical Documentation", icon:FileText},
    {key:"security", label:"Security", icon:ShieldCheck},
    {key:"settings", label:"Settings", icon:Settings},
  ];
  const NAV_ITEMS = [...NAV_ITEMS_PRIMARY, ...NAV_ITEMS_SECONDARY];
  const VISIBLE_PRIMARY = NAV_ITEMS_PRIMARY.filter(it=>canNavigate(role,it.key));
  const VISIBLE_SECONDARY = NAV_ITEMS_SECONDARY.filter(it=>canNavigate(role,it.key));

  const stats = useMemo(()=>{
    const tests = db.tests;
    const total = db.instruments.length;
    const inProgress = tests.filter(t=>t.status==="in_progress"||t.status==="pending").length;
    const awaitingReview = tests.filter(t=>t.status==="review").length;
    const completed = tests.filter(t=>t.status==="approved").length;
    const pass = tests.filter(t=>t.result==="PASS").length;
    const fail = tests.filter(t=>t.result==="FAIL").length;
    const review = tests.filter(t=>t.result==="REVIEW").length;
    const evaluated = pass+fail+review || 1;
    return {total, inProgress, awaitingReview, completed, pass, fail, review,
      passPct: Math.round(pass/evaluated*100), failPct: Math.round(fail/evaluated*100), reviewPct: Math.round(review/evaluated*100),
      reports: db.reports.length};
  },[db]);

  if(authLoading) return <LoginPage onLogin={()=>{}} loading error="Connecting to METRA authentication…"/>;
  if(!session) return <LoginPage onLogin={doLogin} loading={loginLoading} error={loginError}/>;

  return (
    <div className="nawi-root" style={{minHeight:"100vh", fontSize:14}}>
      <GlobalStyle/>
      <div className="no-print nawi-shell" style={{display:"flex", minHeight:"100vh"}}>
        {/* SIDEBAR */}
        <div className="nawi-sidebar">
          <div className="nawi-brand">
            <div className="nawi-brand-name">METRA</div>
            <div style={{color:"var(--ink-faint)", fontSize:9.5, letterSpacing:".09em", marginTop:4, fontWeight:700}}>MEASUREMENT FINGERPRINT PLATFORM</div>
          </div>
          <div className="nawi-nav">
            {VISIBLE_PRIMARY.map(it=>(
              <div key={it.key} className={`n-sidebar-item ${page===it.key?"active":""}`} onClick={()=>{ setSelectedInstrument(null); setSelectedTest(null); setSelectedReport(null); setTestsFilter("all"); nav(it.key); }}>
                <it.icon size={15}/> {it.label}
              </div>
            ))}
            <div className="nawi-support-label">SUPPORTING TOOLS</div>
            {VISIBLE_SECONDARY.map(it=>(
              <div key={it.key} className={`n-sidebar-item ${page===it.key?"active":""}`} onClick={()=>{ setSelectedInstrument(null); setSelectedTest(null); setSelectedReport(null); setTestsFilter("all"); nav(it.key); }}>
                <it.icon size={15}/> {it.label}
              </div>
            ))}
          </div>
          <div className="nawi-sidebar-note">
            Prototype / Simulation environment.<br/>Not an OIML-certified system.
          </div>
        </div>

        {/* MAIN */}
        <div className="nawi-main">
          {/* TOPBAR */}
          <div className="nawi-topbar">
            <div className="nawi-topbar-left" style={{display:"flex", alignItems:"center", gap:18}}>
              <div className="n-eyebrow">Regional Reference Standards Laboratory — Panipat</div>
              <div style={{display:"flex", alignItems:"center", gap:6, fontSize:12, color:"var(--ink-dim)"}}>
                {syncStatus==="synced" && <><span style={{width:8,height:8,borderRadius:"50%",background:"var(--seal)",boxShadow:"0 0 8px rgba(5,150,105,.4)"}}/><Wifi size={13} color="var(--seal)"/> Synced</>}
                {syncStatus==="pending" && <><span style={{width:8,height:8,borderRadius:"50%",background:"var(--amber)",boxShadow:"0 0 8px rgba(217,119,6,.4)"}}/><RefreshCw size={13} color="var(--amber)"/> Pending Sync</>}
                {syncStatus==="failed" && <><span style={{width:8,height:8,borderRadius:"50%",background:"var(--rose)",boxShadow:"0 0 8px rgba(220,38,38,.4)"}}/><WifiOff size={13} color="var(--rose)"/> Sync Failed</>}
                <span className="n-repo-status" title="Authenticated server repository">· {repositoryStatus}</span>
              </div>
            </div>
            <div className="nawi-topbar-right" style={{display:"flex", alignItems:"center", gap:14}}>
              <div style={{position:"relative"}}>
                <Search size={13} style={{position:"absolute", left:12, top:11, color:"var(--ink-faint)"}}/>
                <input className="n-input" style={{width:250, paddingLeft:34, paddingRight:searchQuery?30:12, borderRadius:9999, background:"var(--surface-2)"}} placeholder="Search NAWI ID, model, serial, test, report…"
                  value={searchQuery} onChange={e=>setSearchQuery(e.target.value)}
                  onKeyDown={searchKeyDown}
                  onFocus={()=>setSearchFocused(true)} onBlur={()=>setTimeout(()=>setSearchFocused(false),150)}/>
                {searchQuery && (
                  <X size={13} style={{position:"absolute", right:12, top:11, color:"var(--ink-faint)", cursor:"pointer"}}
                    onMouseDown={e=>{ e.preventDefault(); setSearchQuery(""); setSearchHighlight(-1); }}/>
                )}
                {searchFocused && searchQuery.trim().length>=2 && (
                  <div className="n-panel" style={{position:"absolute", top:38, left:0, width:340, zIndex:120, boxShadow:"0 16px 40px rgba(0,0,0,.12)", maxHeight:320, overflow:"auto", borderRadius:20}}>
                    {searchResults.length===0 && (
                      <div style={{padding:"14px 12px", fontSize:12, color:"var(--ink-faint)", textAlign:"center"}}>No matching instrument, test or report found.</div>
                    )}
                    {["Instrument","Test","Report"].map(cat=>{
                      const items = searchResults.filter(r=>r.type===cat);
                      if(items.length===0) return null;
                      return (
                        <div key={cat}>
                          <div style={{padding:"8px 14px 4px", fontSize:10, letterSpacing:".08em", color:"var(--ink-faint)", fontWeight:700, background:"var(--surface-2)"}}>{cat.toUpperCase()}S</div>
                          {items.map((r)=>{
                            const idx = searchResults.indexOf(r);
                            return (
                              <div key={idx}
                                style={{padding:"10px 14px", borderBottom:"1px solid var(--line)", cursor:"pointer", fontSize:12.5, background: idx===searchHighlight?"var(--surface-2)":"transparent", transition:"background .15s"}}
                                onMouseEnter={()=>setSearchHighlight(idx)}
                                onMouseDown={(e)=>{ e.preventDefault(); runSearch(idx); }}>
                                <strong>{r.label}</strong>
                                {r.sub && <div style={{color:"var(--ink-faint)", fontSize:11, marginTop:1}}>{r.sub}</div>}
                                <div style={{color:"var(--ink-dim)", fontSize:10.5, fontWeight:600, marginTop:3}}>View {cat} →</div>
                              </div>
                            );
                          })}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:7,border:"1px solid var(--line)",background:"var(--surface-2)",borderRadius:9999,padding:"6px 14px",fontSize:11.5}}>
                <span style={{fontWeight:700}}>{currentUserName}</span><span style={{color:"var(--ink-faint)"}}>· {role}</span>
              </div>
              <button className="n-btn n-btn-sm" onClick={doLogout}>Sign out</button>
              <div title="My Profile" onClick={()=>nav("profile")} style={{width:32, height:32, borderRadius:"50%", background:"var(--accent)", color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:700, cursor:"pointer", transition:"all .2s ease", boxShadow:"0 2px 8px rgba(0,0,0,0.12)"}}>
                {currentUserName.split(" ").map(w=>w[0]).slice(0,2).join("")}
              </div>
              <div className="f-mono" style={{fontSize:11.5, color:"var(--ink-faint)"}}>{new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"})}</div>
            </div>
          </div>

          {/* PAGE BODY */}
          <div className="n-scroll nawi-body" style={{flex:1, overflow:"auto"}}>
            {toast && <div style={{position:"fixed", top:20, right:24, zIndex:200, background:"var(--navy)", color:"#fff", padding:"12px 22px", borderRadius:9999, fontSize:13, fontWeight:600, boxShadow:"0 10px 30px rgba(17,24,39,.25)", display:"flex", alignItems:"center", gap:8}}>✓ {toast}</div>}

            {page==="dashboard" && <Dashboard db={db} stats={stats} nav={nav} mutate={mutate} logAudit={logAudit} notify={notify} role={role} currentUserName={currentUserName}/>}
            {page==="instruments" && <InstrumentsPage db={db} mutate={mutate} logAudit={logAudit} nav={nav} notify={notify} role={role} currentUserName={currentUserName} selectedInstrument={selectedInstrument}/>}
            {page==="tests" && !selectedTest && <TestsPage db={db} nav={nav} role={role} presetFilter={testsFilter}/>}
            {page==="tests" && selectedTest && <TestWorkspace db={db} mutate={mutate} logAudit={logAudit} testId={selectedTest} nav={nav} notify={notify} role={role} currentUserName={currentUserName} syncStatus={syncStatus}/>}
            {page==="iot" && <IotLabPage db={db} mutate={mutate} logAudit={logAudit} notify={notify}/>}
            {page==="intelligence" && <MeasurementIntelligencePage db={db} mutate={mutate} logAudit={logAudit} nav={nav} notify={notify} role={role}/>}
            {page==="reports" && <ReportsPage db={db} mutate={mutate} logAudit={logAudit} nav={nav} notify={notify} role={role} selectedInstrument={selectedInstrument} selectedReport={selectedReport}/>}
            {page==="repository" && <RepositoryPage db={db} nav={nav}/>}
            {page==="verify" && <VerifyQrPage db={db} nav={nav} notify={notify}/>}
            {page==="analytics" && <AnalyticsPage db={db} stats={stats} nav={nav}/>}
            {page==="audit" && <AuditPage db={db} nav={nav}/>}
            {page==="users" && <UsersPage db={db} role={role} currentUserName={currentUserName}/>}
            {page==="rules" && <RulesPage db={db} mutate={mutate} logAudit={logAudit} role={role} notify={notify}/>}
            {page==="references" && <OfficialReferencesPage db={db} nav={nav}/>}
            {page==="documentation" && <TechnicalDocumentationPage nav={nav}/>}
            {page==="security" && <SecurityPage db={db} role={role} currentUserName={currentUserName} nav={nav}/>}
            {page==="settings" && <SettingsPage syncStatus={syncStatus} setSyncStatus={setSyncStatus}/>}
            {page==="profile" && <ProfilePage db={db} mutate={mutate} logAudit={logAudit} currentUserName={currentUserName} setCurrentUserName={setCurrentUserName} role={role} notify={notify} nav={nav}/>}
            {page==="terms" && <LegalPage kind="terms" nav={nav}/>}
            {page==="privacy" && <LegalPage kind="privacy" nav={nav}/>}

            <div className="n-panel" style={{marginTop:48, marginBottom:20, padding:"28px 32px", background:"linear-gradient(135deg, #ffffff 0%, #f9fafb 100%)", position:"relative", overflow:"hidden"}}>
              <div style={{position:"absolute", right:-40, top:-40, opacity:0.35, pointerEvents:"none"}}>
                <DecorativeOrbit size={200}/>
              </div>
              <div style={{position:"relative", zIndex:2, maxWidth:720}}>
                <div className="n-eyebrow" style={{marginBottom:6}}>Digital Metrology Platform</div>
                <div style={{fontSize:20, fontWeight:800, color:"var(--ink)", letterSpacing:"-.02em", marginBottom:8}}>Continuous measurement traceability under OIML R 76-1.</div>
                <div style={{fontSize:13.5, color:"var(--ink-dim)", lineHeight:1.6}}>Verifiable digital audit trails, automated maximum permissible error calculations, and instant QR certificate verification for Non-Automatic Weighing Instruments.</div>
              </div>
            </div>

            <div className="nawi-footer">
              <span>METRA · Prototype / Simulation environment · Not an OIML-certified system.</span>
              <span><a href="#" onClick={(e)=>{e.preventDefault();nav("terms")}}>Terms of Service</a><span style={{margin:"0 8px"}}>·</span><a href="#" onClick={(e)=>{e.preventDefault();nav("privacy")}}>Privacy Policy</a></span>
            </div>
          </div>
        </div>
      </div>

      <AssistantWidget db={db} selectedTest={selectedTest} open={assistantOpen} setOpen={setAssistantOpen}/>
    </div>
  );
}

/* ============================== DASHBOARD ============================== */
function DashboardIntelligenceSection({db, nav}){
  const rows = db.instruments.map(inst=>({inst, fp:computeFingerprint(db, inst)}));
  const counts = {
    CHANGE: rows.filter(r=>r.fp.status==="CHANGE").length,
    SIGNIFICANT: rows.filter(r=>r.fp.status==="SIGNIFICANT").length,
    STABLE: rows.filter(r=>r.fp.status==="STABLE").length,
    INSUFFICIENT: rows.filter(r=>r.fp.status==="INSUFFICIENT").length,
  };
  const attention = rows.filter(r=>r.fp.status==="CHANGE"||r.fp.status==="SIGNIFICANT").slice(0,3);
  return (
    <div className="n-panel" style={{padding:"26px 28px", border:"1px solid rgba(229,231,235,0.9)", marginBottom:20}}>
      <div style={{display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:18, flexWrap:"wrap", gap:12}}>
        <div>
          <div className="f-display" style={{fontSize:20, fontWeight:800, letterSpacing:"-.02em", color:"var(--ink)"}}>Measurement Intelligence</div>
          <div style={{fontSize:13.5, color:"var(--ink-dim)", marginTop:4}}>Which instruments show changes in measurement behavior and need attention?</div>
        </div>
        <button className="n-btn n-btn-sm n-btn-primary" onClick={()=>nav("intelligence")}>Open Measurement Intelligence <ChevronRight size={13}/></button>
      </div>
      <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:20}}>
        <MetricBlock label="Pattern Change" value={counts.CHANGE} onClick={()=>nav("intelligence")}/>
        <MetricBlock label="Significant Change" value={counts.SIGNIFICANT} onClick={()=>nav("intelligence")}/>
        <MetricBlock label="Stable Instruments" value={counts.STABLE} onClick={()=>nav("intelligence")}/>
        <MetricBlock label="Building Baseline" value={counts.INSUFFICIENT} onClick={()=>nav("intelligence")}/>
      </div>
      {attention.length===0 ? (
        <div style={{fontSize:13, color:"var(--ink-faint)", padding:"10px 0"}}>No instruments currently show a detected pattern change. All fingerprints are stable or still building a baseline.</div>
      ) : (
        <div>
          {attention.map(({inst,fp})=>{
            const meta = FP_STATUS_META[fp.status];
            return (
              <div key={inst.id} className="n-row-hover" style={{display:"flex", justifyContent:"space-between", alignItems:"center", padding:"12px 10px", borderTop:"1px solid var(--line)", borderRadius:12}} onClick={()=>nav("instruments",{instrumentId:inst.id})}>
                <div>
                  <div style={{fontWeight:700, fontSize:13.5, color:"var(--ink)"}}>{inst.model} <span className="n-badge" style={{background:meta.bg, color:meta.color, borderColor:meta.border, marginLeft:6, fontSize:10.5}}>{meta.icon} {meta.label}</span></div>
                  <div style={{fontSize:12, color:"var(--ink-dim)", marginTop:3}}>Current evaluation: {fp.current.error>=0?"+":""}{fp.current.error} {fp.current.unit} · Previous evaluation: {fp.previous.error>=0?"+":""}{fp.previous.error} {fp.previous.unit}</div>
                </div>
                <button className="n-btn n-btn-sm">View Fingerprint</button>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
function RoleMetricRow({items}){
  return <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>{items.map((x,i)=><MetricBlock key={i} label={x.label} value={x.value} onClick={x.onClick} hint={x.hint}/>)}</div>;
}
function RoleActionBar({role,nav}){
  const actions={Technician:["instruments","tests","iot"],Examiner:["tests","reports","audit"],Authority:["reports","analytics","repository"],Admin:["users","security","audit"],Viewer:["instruments","reports","repository"]};
  return <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:22}}>{(actions[role]||actions.Viewer).map((p,i)=><button key={p} className={i===0?"n-btn n-btn-primary":"n-btn"} onClick={()=>nav(p,p==="tests"&&role==="Examiner"?{testsFilter:"review"}:undefined)}>{({instruments:"Instruments",tests:role==="Examiner"?"Review Queue":"Tests",iot:"IoT Lab",reports:"Reports",audit:"Audit Trail",analytics:"Compliance View",repository:"Repository",users:"Users & Roles",security:"Security"})[p]}</button>)}</div>;
}
function RoleDashboard({db,stats,nav,role,currentUserName}){
  const myTests=db.tests.filter(t=>t.technician===currentUserName);
  const review=db.tests.filter(t=>t.status==="review" || (t.status==="in_progress" && (t.observations||[]).length));
  const issued=db.reports.filter(r=>r.status==="Issued");
  const pendingApproval=db.reports.filter(r=>!r.approver || r.status!=="Issued");
  const titles={Technician:["Technician Workspace","Record observations, monitor assigned evaluations, and resolve test-readiness issues."],Examiner:["Reviewer Dashboard","Review submitted test evidence, compliance results, and reports awaiting examination."],Authority:["Authority Dashboard","Monitor laboratory compliance outcomes, approvals, issued reports, and traceable records."],Admin:["Laboratory Administration","Manage laboratory operations, users, security, repository integrity, and system activity."],Viewer:["Read-Only Dashboard","Read-only visibility into instruments, tests, reports, and repository records."]};
  let metrics,queue,label;
  if(role==="Technician"){metrics=[{label:"My Active Tests",value:myTests.filter(t=>t.status==="in_progress"||t.status==="pending").length,onClick:()=>nav("tests",{testsFilter:"pending"})},{label:"My Completed Tests",value:myTests.filter(t=>t.result).length,onClick:()=>nav("tests")},{label:"Instruments",value:db.instruments.length,onClick:()=>nav("instruments")},{label:"Data / IoT Alerts",value:db.tests.filter(t=>t.envFlag||t.anomalyMsg).length,onClick:()=>nav("iot")}];queue=myTests.slice(0,6);label="My Work Queue";}
  else if(role==="Examiner"){metrics=[{label:"Awaiting Review",value:stats.awaitingReview,onClick:()=>nav("tests",{testsFilter:"review"})},{label:"Tests Failed",value:stats.fail,onClick:()=>nav("tests",{testsFilter:"failed"})},{label:"Reports",value:db.reports.length,onClick:()=>nav("reports")},{label:"Audit Events",value:db.auditLog.length,onClick:()=>nav("audit")}];queue=review.slice(0,6);label="Review Queue";}
  else if(role==="Authority"){metrics=[{label:"Reports Issued",value:issued.length,onClick:()=>nav("reports")},{label:"Pending Approval",value:pendingApproval.length,onClick:()=>nav("reports")},{label:"Passed Tests",value:stats.pass,onClick:()=>nav("analytics")},{label:"Registered Instruments",value:db.instruments.length,onClick:()=>nav("instruments")}];queue=pendingApproval.slice(0,6);label="Approval & Report Queue";}
  else if(role==="Admin"){metrics=[{label:"Users",value:db.users.length,onClick:()=>nav("users")},{label:"Instruments",value:db.instruments.length,onClick:()=>nav("instruments")},{label:"Repository Records",value:db.reports.length,onClick:()=>nav("repository")},{label:"Audit Events",value:db.auditLog.length,onClick:()=>nav("audit")}];queue=db.auditLog.slice(0,6);label="Latest System Activity";}
  else {metrics=[{label:"Instruments",value:db.instruments.length,onClick:()=>nav("instruments")},{label:"Tests",value:db.tests.length,onClick:()=>nav("tests")},{label:"Reports",value:db.reports.length,onClick:()=>nav("reports")},{label:"Issued Reports",value:issued.length,onClick:()=>nav("reports")}];queue=db.reports.slice(0,6);label="Available Reports";}
  return <div><PageHeader eyebrow={`Role · ${role}`} title={titles[role]?.[0]||"Laboratory Dashboard"} desc={titles[role]?.[1]} right={<span className="n-badge n-badge-navy">{currentUserName}</span>}/><RoleActionBar role={role} nav={nav}/><RoleMetricRow items={metrics}/><div style={{display:"grid",gridTemplateColumns:"1.5fr 1fr",gap:16}}><div className="n-panel" style={{padding:"24px 26px"}}><div className="n-eyebrow" style={{marginBottom:14}}>{label}</div>{queue.length===0?<div style={{fontSize:13,color:"var(--ink-faint)",padding:"16px 0"}}>No records currently require attention.</div>:<table className="n-table"><thead><tr><th>Record</th><th>Status</th><th>Instrument</th><th></th></tr></thead><tbody>{queue.map((item,i)=>{const audit=role==="Admin";const inst=audit?null:db.instruments.find(x=>x.id===item.instrumentId);return <tr key={item.id||i}><td><strong>{audit?item.action:(item.testName||item.reportNumber||item.model||item.id)}</strong><div style={{fontSize:11,color:"var(--ink-faint)",marginTop:2}}>{audit?item.entityId:(item.nawiId||item.serialNumber||item.id)}</div></td><td>{audit?<span className="n-badge n-badge-neutral">EVENT</span>:item.result?<ResultBadge result={item.result}/>:item.status?<StatusBadge status={item.status}/>:<span className="n-badge n-badge-neutral">RECORD</span>}</td><td>{inst?.model||item.instrumentId||"—"}</td><td><button className="n-btn n-btn-sm" onClick={()=>audit?nav("audit"):item.reportNumber?nav("reports",{reportId:item.id}):item.testName?nav("tests",{testId:item.id}):nav("instruments",{instrumentId:item.id})}>Open</button></td></tr>})}</tbody></table>}</div><div className="n-panel" style={{padding:"24px 26px"}}><div className="n-eyebrow" style={{marginBottom:14}}>Role Scope</div><div style={{fontSize:13.5,lineHeight:1.7,color:"var(--ink-dim)"}}>{role==="Technician"&&<>Record test observations, environmental readings, evidence, and test progress.</>}{role==="Examiner"&&<>Examine submitted evidence, review compliance results, and work with reports and audit history.</>}{role==="Authority"&&<>Oversee compliance outcomes, approvals, issued reports, repository records, and traceability.</>}{role==="Admin"&&<>Manage users, roles, security, repository operations, and system-level audit visibility.</>}{role==="Viewer"&&<>Read-only access to operational records. Measurement entry, approval, and administration are unavailable.</>}</div><div className="n-hr" style={{margin:"20px 0"}}/><div className="n-eyebrow" style={{marginBottom:10}}>Current Identity</div><div className="n-kv"><dt>User</dt><dd>{currentUserName}</dd><dt>Role</dt><dd>{role}</dd><dt>Access Mode</dt><dd>{role==="Viewer"?"Read only":"Role controlled"}</dd></div></div></div></div>;
}
/* One-click demo seed (ported from Smart NAWI). Clearly labelled DEMO; only offered to roles that may register instruments. */
function runDemoEvaluation({mutate,logAudit,notify,nav}){
  const nawiId=nextNawiId(); const id=uid("INST");
  const demoInst={
    id, nawiId, manufacturer:"Demo Metrology Instruments", applicant:"Demo Metrology Instruments",
    model:"DemoScale-100", serialNumber:`DEMO-${Date.now().toString().slice(-6)}`, instrumentType:"Electronic Bench Scale",
    accuracyClass:"III", maxCapacity:100, minCapacity:1, capacityUnit:"kg", e:0.05, nVerification:2000,
    display:"LCD Digital, 6-digit", firmware:"v1.0-demo", dateSubmission:nowISO(), laboratory:"Regional Reference Standards Laboratory, Panipat",
    status:"Testing In Progress", photos:{instrument:true, nameplate:true}, documents:["Technical Datasheet.pdf"],
    isDemo:true, recordSource:"Demo seed", dataStatus:"DEMO / SIMULATION", createdAt:nowISO(),
    measurementHistory:[
      {id:"MH-DEMO1", date:new Date(Date.now()-1000*60*60*24*220).toISOString(), testKey:"ACC", load:50, error:0.02, unit:"kg", source:"manual", envTemp:24.1, envHumidity:50.6, resultLabel:"PASS", isDemo:true, refTestId:null,
        loadPoints:[{pct:0,load:0,error:0,unit:"kg",source:"manual"},{pct:25,load:25,error:0.01,unit:"kg",source:"manual"},{pct:50,load:50,error:0.02,unit:"kg",source:"manual"}]},
      {id:"MH-DEMO2", date:new Date(Date.now()-1000*60*60*24*60).toISOString(), testKey:"ACC", load:50, error:0.03, unit:"kg", source:"iot", envTemp:24.3, envHumidity:51.0, resultLabel:"PASS", isDemo:true, refTestId:null,
        loadPoints:[{pct:0,load:0,error:0,unit:"kg",source:"iot"},{pct:25,load:25,error:0.015,unit:"kg",source:"iot"},{pct:50,load:50,error:0.03,unit:"kg",source:"iot"}]},
    ],
  };
  const rule=activeRule(); const items=generateTestPlan(demoInst);
  const tp={id:uid("TP"), instrumentId:id, tests:items, generatedAt:nowISO(), ruleVersionId:rule.id};
  const mk=(key,name,extra={})=>({id:uid("TEST"), instrumentId:id, testKey:key, testName:name, status:"pending", observations:[], evidence:[], ruleVersionId:rule.id, remarks:"", technician:"", reviewer:"", approver:"", result:null, createdAt:nowISO(), ...extra});
  const accTest=mk("ACC",TEST_DEFS.ACC.name,{status:"in_progress",technician:"Demo Technician"});
  // Deterministic readings: the 50 kg point deviates more than the seeded history so the pattern-change scenario reproduces every run.
  [[0,0.00],[25,25.02],[50,50.15]].forEach(([load,observed])=>{
    const calc=evalAccuracy(load,observed,demoInst,rule);
    accTest.observations.push({id:uid("OBS"), load, observed, unit:"kg", source:"iot", timestamp:nowISO(), syncFlag:"synced", ...calc});
  });
  accTest.result=accTest.observations.some(o=>o.result==="FAIL")?"FAIL":"PASS";
  const others=items.filter(i=>i.status==="applicable" && i.key!=="ACC").map(i=>mk(i.key,i.name));
  mutate(d=>{
    d.instruments.push(demoInst); d.testPlans.push(tp); d.tests.push(accTest,...others);
    logAudit(d,"Instrument registered","Instrument",id,"—",`${nawiId} — DEMO instrument`,"Demo Mode — one-click demonstration seed");
    logAudit(d,"Test plan generated","TestPlan",tp.id,"—",`${items.filter(i=>i.status==="applicable").length} applicable tests`,"Demo Mode");
  });
  notify("Demo evaluation loaded — DEMO / SIMULATION MODE.");
  nav("instruments",{instrumentId:id});
}
function Dashboard({db,stats,nav,mutate,logAudit,notify,role,currentUserName}){
  return <div>
    {can(role,"instrument:create") && <div style={{display:"flex",justifyContent:"flex-end",marginBottom:10}}>
      <button className="n-btn n-btn-seal" onClick={()=>runDemoEvaluation({mutate,logAudit,notify,nav})}><FlaskConical size={14}/> Run Demo Evaluation</button>
    </div>}
    <RoleDashboard db={db} stats={stats} nav={nav} role={role} currentUserName={currentUserName}/>
  </div>;
}

function DashboardBase({db, stats, nav, mutate, logAudit, notify}){
  const recent = db.auditLog.slice(0,6);
  const iotAlerts = db.iotDevices.filter(d=>d.status!=="connected").length + db.tests.filter(t=>t.envFlag||t.anomalyMsg).length;
  const activeEvals = db.tests.filter(t=>t.status==="in_progress" || (t.status==="pending" && (t.observations||[]).length>0))
    .map(t=>{
      const inst = db.instruments.find(i=>i.id===t.instrumentId);
      const pct = Math.min(100, Math.round(((t.observations||[]).length / (EXPECTED_OBS[t.testKey]||1)) * 100));
      return {t, inst, pct};
    });

  return (
    <div>
      <PageHeader eyebrow="Overview" title="Laboratory Dashboard" desc="Every NAWI develops a historical measurement behavior. METRA surfaces changes in that behavior first — official compliance and lab operations follow below."
        right={<div className="n-badge n-badge-neutral">Live laboratory records</div>}/>

      <DashboardIntelligenceSection db={db} nav={nav}/>

      <div className="n-eyebrow" style={{margin:"22px 0 10px"}}>Laboratory Operations (Secondary)</div>
      <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:14}}>
        <MetricBlock label="Instruments Under Evaluation" value={stats.total} onClick={()=>nav("instruments")}/>
        <MetricBlock label="Tests In Progress" value={stats.inProgress} onClick={()=>nav("tests",{testsFilter:"pending"})}/>
        <MetricBlock label="Awaiting Review" value={stats.awaitingReview} onClick={()=>nav("tests",{testsFilter:"review"})} hint="Pending review queue"/>
        <MetricBlock label="Approved Tests" value={stats.completed} onClick={()=>nav("tests",{testsFilter:"approved"})}/>
      </div>
      <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:16}}>
        <MetricBlock label="Reports Issued" value={stats.reports} onClick={()=>nav("reports")}/>
        <MetricBlock label="IoT / Data Alerts" value={iotAlerts} onClick={()=>nav("iot")} hint="Disconnections & anomaly flags"/>
        <MetricBlock label="Failed Tests" value={stats.fail} onClick={()=>nav("tests",{testsFilter:"failed"})}/>
        <MetricBlock label="Total Registered Users" value={db.users.length} onClick={()=>nav("users")}/>
      </div>

      {activeEvals.length>0 && (
        <div className="n-panel" style={{padding:"24px 28px", marginBottom:20}}>
          <div className="n-eyebrow" style={{marginBottom:14}}>Active Evaluations</div>
          <div style={{border:"1px solid var(--line)", borderRadius:20, overflow:"hidden"}}>
            <table className="n-table">
              <thead><tr><th>Instrument</th><th>Current Test</th><th>Progress</th><th>Technician</th><th>Status</th><th></th></tr></thead>
              <tbody>
                {activeEvals.map(({t,inst,pct})=>(
                  <tr key={t.id}>
                    <td><strong>{inst?.model}</strong> <span style={{color:"var(--ink-faint)", fontSize:12}}>({inst?.nawiId||inst?.serialNumber})</span></td>
                    <td>{t.testName}</td>
                    <td style={{minWidth:140}}>
                      <div style={{display:"flex", alignItems:"center", gap:10}}>
                        <div style={{flex:1, height:8, background:"var(--surface-2)", borderRadius:9999, overflow:"hidden"}}>
                          <div style={{height:8, width:`${pct}%`, background:"linear-gradient(90deg, var(--seal), #10b981)", borderRadius:9999}}/>
                        </div>
                        <span className="f-mono" style={{fontSize:11.5, fontWeight:600}}>{pct}%</span>
                      </div>
                    </td>
                    <td>{t.technician||"—"}</td>
                    <td><StatusBadge status={t.status}/></td>
                    <td><button className="n-btn n-btn-sm" onClick={()=>nav("tests",{testId:t.id})}>Continue Test</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div style={{display:"grid", gridTemplateColumns:"1.3fr 1fr", gap:18}}>
        <div className="n-panel" style={{padding:"26px 28px"}}>
          <div className="n-eyebrow" style={{marginBottom:14}}>Compliance Outcomes (Evaluated Tests)</div>
          <div style={{display:"flex", gap:26, alignItems:"center"}}>
            <div style={{width:104, height:104, borderRadius:"50%", flexShrink:0, boxShadow:"0 4px 16px rgba(0,0,0,.06)", background:`conic-gradient(var(--seal) 0 ${stats.passPct}%, var(--rose) ${stats.passPct}% ${stats.passPct+stats.failPct}%, var(--amber) ${stats.passPct+stats.failPct}% 100%)`}}/>
            <div style={{fontSize:13.5}}>
              <div style={{display:"flex", alignItems:"center", gap:8, marginBottom:8}}><span style={{width:10,height:10,borderRadius:"50%",background:"var(--seal)",display:"inline-block"}}/> Pass — <strong>{stats.passPct}%</strong></div>
              <div style={{display:"flex", alignItems:"center", gap:8, marginBottom:8}}><span style={{width:10,height:10,borderRadius:"50%",background:"var(--rose)",display:"inline-block"}}/> Fail — <strong>{stats.failPct}%</strong></div>
              <div style={{display:"flex", alignItems:"center", gap:8}}><span style={{width:10,height:10,borderRadius:"50%",background:"var(--amber)",display:"inline-block"}}/> Review — <strong>{stats.reviewPct}%</strong></div>
            </div>
          </div>
          <div className="n-hr" style={{margin:"20px 0"}}/>
          <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12}}>
            <div className="n-eyebrow">Recent Activity</div>
            <span style={{fontSize:11.5, color:"var(--ink)", fontWeight:600, cursor:"pointer", transition:"color .15s"}} onClick={()=>nav("audit")}>View Audit Trail →</span>
          </div>
          {recent.map(a=>(
            <div key={a.id} className="n-row-hover" style={{display:"flex", justifyContent:"space-between", fontSize:12.5, padding:"8px 8px", borderBottom:"1px solid var(--line)", borderRadius:8}} onClick={()=>jumpToEntity(a, nav, db)}>
              <span><strong>{a.action}</strong> — {a.entityId} <span style={{color:"var(--ink-faint)"}}>by {a.user}</span></span>
              <span className="f-mono" style={{color:"var(--ink-faint)", fontSize:11.5}}>{fmtT(a.timestamp)}</span>
            </div>
          ))}
        </div>
        <div className="n-panel" style={{padding:"26px 28px"}}>
          <div className="n-eyebrow" style={{marginBottom:14}}>Smart Lab Status</div>
          {db.iotDevices.map(d=>(
            <div key={d.id} className="n-row-hover" style={{display:"flex", justifyContent:"space-between", padding:"10px 8px", borderBottom:"1px solid var(--line)", fontSize:13, borderRadius:8}} onClick={()=>nav("iot")}>
              <span>{d.name}</span>
              <span style={{display:"flex", alignItems:"center", gap:6, color: d.status==="connected"?"var(--seal)":"var(--rose)", fontWeight:650}}>
                <span style={{width:7,height:7,borderRadius:"50%", background: d.status==="connected"?"var(--seal)":"var(--rose)", boxShadow: d.status==="connected"?"0 0 6px rgba(5,150,105,.4)":"0 0 6px rgba(220,38,38,.4)"}}/>
                {d.status==="connected"?"Connected":"Disconnected"}
              </span>
            </div>
          ))}
          <button className="n-btn n-btn-sm" style={{marginTop:16, width:"100%"}} onClick={()=>nav("iot")}>Open IoT Lab <ChevronRight size={13}/></button>
          <div className="n-hr" style={{margin:"20px 0"}}/>
          <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", cursor:"pointer", padding:"6px 4px", borderRadius:12}} onClick={()=>nav("reports")}>
            <div>
              <div className="n-eyebrow" style={{marginBottom:6}}>Reports Issued</div>
              <div className="f-display" style={{fontSize:32, fontWeight:800, letterSpacing:"-.02em", color:"var(--ink)"}}>{stats.reports}</div>
            </div>
            <span style={{width:32, height:32, borderRadius:"50%", background:"var(--surface-2)", display:"inline-flex", alignItems:"center", justifyContent:"center"}}><ChevronRight size={16} color="var(--ink-dim)"/></span>
          </div>
        </div>
      </div>
    </div>
  );
}
function jumpToEntity(a, nav, db){
  if(a.entity==="Instrument") nav("instruments",{instrumentId:a.entityId});
  else if(a.entity==="Test"){ const t=db.tests.find(x=>x.id===a.entityId); if(t) nav("tests",{testId:t.id}); }
  else if(a.entity==="Report"){ const r=db.reports.find(x=>x.id===a.entityId); if(r) nav("reports",{instrumentId:r.instrumentId, reportId:r.id}); }
  else if(a.entity==="TestPlan"){ const tp=db.testPlans.find(x=>x.id===a.entityId); if(tp) nav("instruments",{instrumentId:tp.instrumentId}); }
  else nav("audit");
}
const MetricBlock = ({label, value, onClick, hint}) => (
  <div className="n-panel" style={{padding:"18px 20px", cursor: onClick?"pointer":"default", borderRadius:24, border:"1px solid rgba(229,231,235,0.9)", transition:"all .2s cubic-bezier(0.16,1,0.3,1)"}} onClick={onClick}
    onMouseEnter={e=>{ if(onClick) { e.currentTarget.style.borderColor="var(--ink)"; e.currentTarget.style.transform="translateY(-2px)"; } }}
    onMouseLeave={e=>{ if(onClick) { e.currentTarget.style.borderColor="rgba(229,231,235,0.9)"; e.currentTarget.style.transform="translateY(0)"; } }}>
    <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10}}>
      <div className="n-label" style={{fontSize:11, letterSpacing:".08em", color:"var(--ink-faint)", fontWeight:700}}>{label}</div>
      {onClick && <span style={{width:24, height:24, borderRadius:"50%", background:"var(--surface-2)", display:"inline-flex", alignItems:"center", justifyContent:"center"}}><ChevronRight size={13} color="var(--ink-dim)"/></span>}
    </div>
    <div className="f-display" style={{fontSize:30, fontWeight:800, letterSpacing:"-.025em", color:"var(--ink)"}}>{value}</div>
    {hint && <div style={{fontSize:11, color:"var(--ink-faint)", marginTop:4}}>{hint}</div>}
  </div>
);
const Breadcrumb = ({items}) => (
  <div style={{display:"flex", alignItems:"center", gap:8, fontSize:12, color:"var(--ink-faint)", marginBottom:12, flexWrap:"wrap"}}>
    {items.map((it,i)=>(
      <span key={i} style={{display:"flex", alignItems:"center", gap:8}}>
        {i>0 && <ChevronRight size={11} color="var(--ink-faint)"/>}
        {it.onClick ? <span onClick={it.onClick} style={{color:"var(--ink-dim)", fontWeight:600, cursor:"pointer", transition:"color .15s"}} onMouseEnter={e=>e.currentTarget.style.color="var(--ink)"} onMouseLeave={e=>e.currentTarget.style.color="var(--ink-dim)"}>{it.label}</span> : <span style={{color:"var(--ink)", fontWeight:700}}>{it.label}</span>}
      </span>
    ))}
  </div>
);
const PageHeader = ({eyebrow, title, desc, right, breadcrumb}) => (
  <div style={{marginBottom:24}}>
    {breadcrumb && <Breadcrumb items={breadcrumb}/>}
    <div style={{display:"flex", justifyContent:"space-between", alignItems:"flex-end", gap:16, flexWrap:"wrap"}}>
      <div>
        <div className="n-eyebrow" style={{marginBottom:6, fontSize:11, letterSpacing:".12em", color:"var(--ink-dim)", fontWeight:750}}>{eyebrow}</div>
        <div className="f-display" style={{fontSize:32, fontWeight:800, letterSpacing:"-.03em", color:"var(--ink)", lineHeight:1.15}}>{title}</div>
        {desc && <div style={{color:"var(--ink-dim)", fontSize:14, marginTop:6, maxWidth:680, lineHeight:1.6}}>{desc}</div>}
      </div>
      {right}
    </div>
  </div>
);

/* ============================== INSTRUMENTS ============================== */
function emptyInstrumentForm(){
  return {manufacturer:"", applicant:"", model:"", serialNumber:"", instrumentType:"Electronic Bench Scale",
    accuracyClass:"III", maxCapacity:"", minCapacity:"", capacityUnit:"kg", e:"", nVerification:"", verificationStage:"initial",
    display:"", firmware:"", laboratory:"Regional Reference Standards Laboratory, Panipat"};
}
function InstrumentsPage({db, mutate, logAudit, nav, notify, role, currentUserName, selectedInstrument}){
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState(emptyInstrumentForm());
  const [ocrOpen, setOcrOpen] = useState(false);
  const [ocrStage, setOcrStage] = useState("idle");
  const [ocrData, setOcrData] = useState(null);

  const set = (k,v) => setForm(f=>({...f,[k]:v}));

  const runOcr = () => {
    setOcrStage("scanning");
    setTimeout(()=>{
      setOcrData({
        manufacturer:"PrecisionWeigh Technologies", model:"PW-500E", serialNumber:"PW500-2026-001",
        maxCapacity:"500", minCapacity:"5", capacityUnit:"kg", accuracyClass:"III", e:"0.2",
      });
      setOcrStage("done");
    }, 1400);
  };
  const confirmOcr = () => {
    setForm(f=>({...f, ...ocrData, nVerification: ocrData.maxCapacity && ocrData.e ? String(Math.round(ocrData.maxCapacity/ocrData.e)) : f.nVerification}));
    setOcrOpen(false); setOcrStage("idle"); setOcrData(null);
    notify("Nameplate data applied to registration form after officer confirmation.");
  };

  const classFlags = (form.maxCapacity && form.minCapacity!=="" && form.e && form.accuracyClass)
    ? validateClassConfig({accuracyClass:form.accuracyClass, maxCapacity:Number(form.maxCapacity), minCapacity:Number(form.minCapacity||0), e:Number(form.e), capacityUnit:form.capacityUnit})
    : [];

  const submit = () => {
    if(!can(role,"instrument:create")){ notify("Your role does not permit registering instruments."); return; }
    if(!form.manufacturer||!form.model||!form.serialNumber||!form.maxCapacity||!form.e){ notify("Please complete the required fields before saving."); return; }
    if(classFlags.some(f=>f.level==="bad")){ notify("Resolve the OIML class-configuration error shown below before registering this instrument."); return; }
    const id = uid("INST");
    const nawiId = nextNawiId();
    const inst = {...form, id, nawiId, status:"Registered", maxCapacity:Number(form.maxCapacity), minCapacity:Number(form.minCapacity||0),
      e:Number(form.e), nVerification: form.nVerification ? Number(form.nVerification) : Math.round(Number(form.maxCapacity)/Number(form.e)),
      photos:{instrument:true, nameplate: !!ocrData}, documents:[], isDemo:false, recordSource:"Laboratory registration", dataStatus:"User-entered record", measurementHistory:[], dateSubmission:nowISO(), createdAt:nowISO()};
    mutate(d=>{ d.instruments.push(inst); logAudit(d,"Instrument registered","Instrument",id,"—",`${nawiId} — ${inst.model} / ${inst.serialNumber}`,"New submission via registration form"); });
    notify(`Instrument registered as ${nawiId}. Opening digital passport.`);
    setShowForm(false); setForm(emptyInstrumentForm());
    nav("instruments", {instrumentId:id});
  };

  const [detail, setDetail] = useState(selectedInstrument || null);
  useEffect(()=>{ if(selectedInstrument) setDetail(selectedInstrument); },[selectedInstrument]);

  if(detail){
    const inst = db.instruments.find(i=>i.id===detail);
    return <InstrumentPassport db={db} instrument={inst} onBack={()=>setDetail(null)} nav={nav} mutate={mutate} logAudit={logAudit} notify={notify} role={role}/>;
  }

  return (
    <div>
      <PageHeader eyebrow="Registration & Records" title="Instruments" desc="Register new NAWI submissions and access each instrument's digital passport."
        right={can(role,"instrument:create") ? <button className="n-btn n-btn-primary" onClick={()=>setShowForm(true)}><Plus size={14}/> Register Instrument</button> : null}/>

      <div className="n-panel">
        <table className="n-table">
          <thead><tr><th>NAWI ID</th><th>Manufacturer / Model</th><th>Serial No.</th><th>Class</th><th>Max Capacity</th><th>Status</th><th>Submitted</th><th></th></tr></thead>
          <tbody>
            {db.instruments.map(i=>(
              <tr key={i.id} className="n-row-hover" onClick={()=>setDetail(i.id)}>
                <td className="f-mono">{i.nawiId||"—"}</td>
                <td><strong>{i.manufacturer}</strong><br/><span style={{color:"var(--ink-faint)"}}>{i.model}</span></td>
                <td className="f-mono">{i.serialNumber}</td>
                <td>{i.accuracyClass}</td>
                <td>{i.maxCapacity} {i.capacityUnit}</td>
                <td><span className="n-badge n-badge-neutral">{i.status}</span></td>
                <td className="f-mono">{fmtDT(i.dateSubmission)}</td>
                <td><ChevronRight size={14} color="var(--ink-faint)"/></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {showForm && (
        <Modal title="Register New Instrument" onClose={()=>setShowForm(false)} width={720}>
          <div style={{display:"flex", justifyContent:"flex-end", marginBottom:12}}>
            <button className="n-btn n-btn-sm" onClick={()=>setOcrOpen(true)}><Camera size={13}/> Scan Instrument Nameplate</button>
          </div>
          <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:12}}>
            <Field label="Manufacturer Name *"><input className="n-input" value={form.manufacturer} onChange={e=>set("manufacturer",e.target.value)}/></Field>
            <Field label="Applicant"><input className="n-input" value={form.applicant} onChange={e=>set("applicant",e.target.value)}/></Field>
            <Field label="Model Number *"><input className="n-input" value={form.model} onChange={e=>set("model",e.target.value)}/></Field>
            <Field label="Serial Number *"><input className="n-input" value={form.serialNumber} onChange={e=>set("serialNumber",e.target.value)}/></Field>
            <Field label="Instrument Type">
              <select className="n-select" value={form.instrumentType} onChange={e=>set("instrumentType",e.target.value)}>
                <option>Electronic Bench Scale</option><option>Electronic Precision Balance</option><option>Platform Scale</option><option>Counter Scale</option>
              </select>
            </Field>
            <Field label="Accuracy Class">
              <select className="n-select" value={form.accuracyClass} onChange={e=>set("accuracyClass",e.target.value)}>
                <option>I</option><option>II</option><option>III</option><option>IIII</option>
              </select>
            </Field>
            <Field label="Verification Stage">
              <select className="n-select" value={form.verificationStage} onChange={e=>set("verificationStage",e.target.value)}>
                <option value="initial">Initial Verification</option><option value="in-service">In-Service Verification</option>
              </select>
            </Field>
            <Field label="Maximum Capacity (Max) *"><input className="n-input" type="number" value={form.maxCapacity} onChange={e=>set("maxCapacity",e.target.value)}/></Field>
            <Field label="Minimum Capacity (Min)"><input className="n-input" type="number" value={form.minCapacity} onChange={e=>set("minCapacity",e.target.value)}/></Field>
            <Field label="Capacity Unit">
              <select className="n-select" value={form.capacityUnit} onChange={e=>set("capacityUnit",e.target.value)}><option>kg</option><option>g</option></select>
            </Field>
            <Field label="Verification Scale Interval (e) *"><input className="n-input" type="number" step="0.001" value={form.e} onChange={e=>set("e",e.target.value)}/></Field>
            <Field label="No. of Verification Intervals (n)"><input className="n-input" type="number" value={form.nVerification} onChange={e=>set("nVerification",e.target.value)} placeholder="Auto = Max / e"/></Field>
          </div>
          {classFlags.length>0 && (
            <div style={{margin:"4px 0 10px", padding:"8px 10px", border:"1px solid var(--line)", borderRadius:8}}>
              <div style={{fontSize:11, fontWeight:700, color:"var(--ink-dim)", marginBottom:4}}>OIML Class Configuration Check</div>
              {classFlags.map((f,i)=>(
                <div key={i} style={{display:"flex", gap:6, alignItems:"flex-start", fontSize:11.5, marginBottom:2, color: f.level==="bad"?"var(--rose)":f.level==="warn"?"var(--amber)":"var(--seal)"}}>
                  <FlagIcon level={f.level}/> <span>{f.msg}</span>
                </div>
              ))}
            </div>
          )}
          <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:12}}>
            <Field label="Display"><input className="n-input" value={form.display} onChange={e=>set("display",e.target.value)}/></Field>
            <Field label="Firmware / Software Version"><input className="n-input" value={form.firmware} onChange={e=>set("firmware",e.target.value)}/></Field>
            <Field label="Laboratory" full><input className="n-input" value={form.laboratory} onChange={e=>set("laboratory",e.target.value)}/></Field>
          </div>
          <div style={{marginTop:14, padding:10, background:"var(--surface-2)", fontSize:12, color:"var(--ink-dim)"}}>
            <Upload size={13} style={{verticalAlign:"-2px", marginRight:5}}/>
            Instrument photograph, nameplate photograph, technical documents and calibration certificates can be attached from the instrument's digital passport after registration.
          </div>
          <div style={{display:"flex", justifyContent:"flex-end", gap:8, marginTop:16}}>
            <button className="n-btn" onClick={()=>setShowForm(false)}>Cancel</button>
            <button className="n-btn n-btn-primary" onClick={submit}>Save & Create Passport</button>
          </div>
        </Modal>
      )}

      {ocrOpen && (
        <Modal title="Scan Instrument Nameplate" onClose={()=>{setOcrOpen(false); setOcrStage("idle"); setOcrData(null);}} width={460}>
          {ocrStage==="idle" && (
            <div style={{textAlign:"center", padding:"20px 0"}}>
              <Camera size={34} color="var(--ink-faint)"/>
              <div style={{fontSize:12.5, color:"var(--ink-dim)", margin:"10px 0 16px"}}>Upload or capture a photograph of the instrument nameplate. Data will be extracted for officer verification — never saved automatically.</div>
              <button className="n-btn n-btn-primary" onClick={runOcr}><Upload size={13}/> Upload Nameplate Photo</button>
            </div>
          )}
          {ocrStage==="scanning" && (
            <div style={{textAlign:"center", padding:"30px 0", fontSize:13, color:"var(--ink-dim)"}}>
              <RefreshCw size={22} className="f-mono" style={{animation:"spin 1s linear infinite"}}/>
              <div style={{marginTop:10}}>Analyzing nameplate image…</div>
            </div>
          )}
          {ocrStage==="done" && ocrData && (
            <div>
              <div className="n-eyebrow" style={{marginBottom:8}}>AI Extracted Data — Requires Officer Verification</div>
              <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:14}}>
                {Object.entries(ocrData).map(([k,v])=>(
                  <Field key={k} label={k}><input className="n-input" value={v} onChange={e=>setOcrData(o=>({...o,[k]:e.target.value}))}/></Field>
                ))}
              </div>
              <div style={{fontSize:11.5, color:"var(--amber)", marginBottom:12}}><AlertTriangle size={12} style={{verticalAlign:"-2px"}}/> Verify every field against the physical nameplate before confirming.</div>
              <div style={{display:"flex", justifyContent:"flex-end", gap:8}}>
                <button className="n-btn" onClick={()=>{setOcrOpen(false); setOcrStage("idle");}}>Discard</button>
                <button className="n-btn n-btn-seal" onClick={confirmOcr}><Check size={13}/> Officer Verified — Apply to Form</button>
              </div>
            </div>
          )}
        </Modal>
      )}
    </div>
  );
}
const Field = ({label, children, full}) => (
  <div style={full?{gridColumn:"1 / -1"}:{}}>
    <label className="n-field-label">{label}</label>
    {children}
  </div>
);

/* ============================== INSTRUMENT PASSPORT ============================== */
function InstrumentPassport({db, instrument, onBack, nav, mutate, logAudit, notify, role}){
  const [planOpen, setPlanOpen] = useState(false);
  const plan = db.testPlans.find(p=>p.instrumentId===instrument.id);
  const tests = db.tests.filter(t=>t.instrumentId===instrument.id);
  const reports = db.reports.filter(r=>r.instrumentId===instrument.id);
  const report = reports[0];

  const overall = useMemo(()=>{
    if(tests.length===0) return "NOT STARTED";
    if(tests.some(t=>t.result==="FAIL")) return "NON-COMPLIANT";
    if(tests.some(t=>t.result==="REVIEW"||t.status==="review")) return "REVIEW REQUIRED";
    if(tests.every(t=>t.status==="approved"||t.status==="not-applicable")) return "COMPLIANT";
    return "REVIEW REQUIRED";
  },[tests]);

  const genPlan = () => {
    if(!can(role,"instrument:edit-draft")){ notify("Your role does not permit generating a test plan."); return; }
    const items = generateTestPlan(instrument);
    const tp = {id:uid("TP"), instrumentId:instrument.id, tests:items, generatedAt:nowISO(), ruleVersionId:activeRule().id};
    mutate(d=>{
      d.testPlans.push(tp);
      items.filter(i=>i.status==="applicable").forEach(i=>{
        d.tests.push({id:uid("TEST"), instrumentId:instrument.id, testKey:i.key, testName:i.name, status:"pending", observations:[], ruleVersionId:activeRule().id, remarks:"", technician:"", reviewer:"", approver:"", result:null, createdAt:nowISO()});
      });
      const inst = d.instruments.find(x=>x.id===instrument.id); inst.status = "Testing In Progress";
      logAudit(d,"Test plan generated","TestPlan",tp.id,"—",`${items.filter(i=>i.status==="applicable").length} applicable tests`,"Generated from instrument parameters (accuracy class, capacity)");
    });
    notify("Test plan generated from instrument parameters.");
    setPlanOpen(true);
  };

  return (
    <div>
      <Breadcrumb items={[{label:"Instruments", onClick:onBack}, {label:instrument.model}]}/>
      <button className="n-btn n-btn-sm n-btn-ghost" style={{marginBottom:12}} onClick={onBack}><ChevronLeft size={14}/> Back to Instruments</button>
      <div className="n-panel" style={{padding:20, marginBottom:16}}>
        <div style={{display:"flex", justifyContent:"space-between"}}>
          <div>
            <div style={{display:"flex", alignItems:"center", gap:8}}>
              <div className="n-eyebrow">NAWI Digital Passport</div>
            </div>
            <div className="f-display" style={{fontSize:20, fontWeight:700, margin:"4px 0"}}>{instrument.manufacturer} — {instrument.model}</div>
            <div className="f-mono" style={{fontSize:12, color:"var(--navy-dim)", fontWeight:600, marginBottom:6}}>{instrument.nawiId || "NAWI-ID pending"}</div>
            <div className="n-kv" style={{marginTop:10}}>
              <dt>Serial Number</dt><dd>{instrument.serialNumber}</dd>
              <dt>Accuracy Class</dt><dd>{instrument.accuracyClass}</dd>
              <dt>Max Capacity</dt><dd>{instrument.maxCapacity} {instrument.capacityUnit}</dd>
              <dt>Min Capacity</dt><dd>{instrument.minCapacity} {instrument.capacityUnit}</dd>
              <dt>e / n</dt><dd>{instrument.e} {instrument.capacityUnit} / {instrument.nVerification}</dd>
              <dt>Registered</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>{fmtDT(instrument.createdAt)}</dd>
              <dt>Laboratory</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>{instrument.laboratory}</dd>
              <dt>Status</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>{instrument.status}</dd>
              <dt>Record Source</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>{instrument.recordSource || "Laboratory registration"}</dd>
              <dt>Data Status</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>{instrument.dataStatus || "User-entered record"}</dd>
            </div>
          </div>
          <div style={{display:"flex", gap:20}}>
            <div style={{textAlign:"center"}}>
              <div style={{border:"1px solid var(--line)", padding:5, display:"inline-block", background:"#fff"}}>
                {report ? <QrCode value={makeReportQrToken(report.id)} size={92}/> : <QrCode value={`METRA:I:${String(instrument.nawiId || instrument.id).toUpperCase()}`} size={92}/>}
              </div>
              <div style={{fontSize:9.5, color:"var(--ink-faint)", marginTop:4, maxWidth:115}}>
                {report ? <>Scan to verify report <span className="f-mono">{report.reportNumber}</span>.</> : <>Instrument QR — report verification becomes available after a report is issued.</>}
              </div>
            </div>
            <div style={{textAlign:"center"}}>
              <Seal status={overall}/>
              <div style={{fontSize:10.5, color:"var(--ink-faint)", marginTop:4}}>Compliance Status</div>
            </div>
          </div>
        </div>
      </div>

      <MeasurementFingerprint db={db} instrument={instrument} nav={nav}/>

      <MeasurementEvolutionTimeline db={db} instrument={instrument}/>

      <EvaluationHistoryTable db={db} instrument={instrument} nav={nav}/>

      <div className="n-eyebrow" style={{margin:"20px 0 10px"}}>⚖️ Official OIML Compliance</div>

      {!plan && (
        <div className="n-panel" style={{padding:18, marginBottom:16, textAlign:"center"}}>
          <FlaskConical size={22} color="var(--navy-dim)"/>
          <div style={{margin:"8px 0 12px", fontSize:13, color:"var(--ink-dim)"}}>No test plan has been generated for this instrument yet.</div>
          <button className="n-btn n-btn-primary" onClick={genPlan}>Generate Test Plan</button>
        </div>
      )}

      {plan && (
        <div className="n-panel" style={{padding:18, marginBottom:16}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Applicable Tests (OIML R-76 — {activeRule().version})</div>
          <table className="n-table">
            <thead><tr><th>Test</th><th>Purpose</th><th>Equipment</th><th>Status</th><th>Result</th><th></th></tr></thead>
            <tbody>
              {plan.tests.map(pt=>{
                const test = tests.find(t=>t.testKey===pt.key);
                return (
                  <tr key={pt.key} className={test?"n-row-hover":""} onClick={()=>test && nav("tests",{testId:test.id})}>
                    <td><strong>{pt.name}</strong></td>
                    <td style={{maxWidth:260, color:"var(--ink-dim)"}}>{pt.purpose}</td>
                    <td style={{color:"var(--ink-dim)"}}>{pt.requiredEquipment}</td>
                    <td>{test ? <StatusBadge status={test.status}/> : <StatusBadge status={pt.status}/>}</td>
                    <td>{test && <ResultBadge result={test.result}/>}</td>
                    <td>{test && <ChevronRight size={14} color="var(--ink-faint)"/>}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {tests.length>0 && (
        <div className="n-panel" style={{padding:18, marginBottom:16}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Compliance Breakdown</div>
          <table className="n-table">
            <thead><tr><th>Category</th><th>Result</th><th>Reason</th></tr></thead>
            <tbody>
              {tests.map(t=>(
                <tr key={t.id}>
                  <td>{t.testName}</td>
                  <td><ResultBadge result={t.result}/></td>
                  <td style={{color:"var(--ink-dim)"}}>{t.remarks || (t.status==="pending" ? "Not yet executed." : "Pending evaluation.")}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <EvaluationChecklist instrument={instrument} plan={plan} tests={tests} report={report}/>

      <div className="n-eyebrow" style={{margin:"20px 0 10px"}}>IoT / Environmental History</div>
      <InstrumentEnvHistory db={db} instrument={instrument} tests={tests}/>

      <div className="n-eyebrow" style={{margin:"20px 0 10px"}}>Trust & Traceability</div>

      <div className="n-panel" style={{padding:18, marginBottom:16}}>
        <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10}}>
          <div className="n-eyebrow">Report History ({reports.length})</div>
          <button className="n-btn n-btn-sm" onClick={()=>nav("reports",{instrumentId:instrument.id})}>Open in Reports <ChevronRight size={13}/></button>
        </div>
        {reports.length===0 && <div style={{fontSize:12.5, color:"var(--ink-faint)"}}>No reports issued yet for this instrument. Complete and approve all applicable tests, then generate a report.</div>}
        {reports.map(r=>(
          <div key={r.id} className="n-row-hover" style={{display:"flex", justifyContent:"space-between", alignItems:"center", padding:"9px 4px", borderBottom:"1px solid var(--line)"}} onClick={()=>nav("reports",{instrumentId:instrument.id, reportId:r.id})}>
            <div>
              <div className="f-mono" style={{fontSize:12.5, fontWeight:600}}>{r.reportNumber}</div>
              <div style={{fontSize:11, color:"var(--ink-faint)"}}>Generated {fmtDT(r.generatedAt)} · Approved by {r.approver||"—"}</div>
            </div>
            <ChevronRight size={14} color="var(--ink-faint)"/>
          </div>
        ))}
      </div>

      <ChainOfCustody db={db} instrument={instrument} tests={tests} report={report}/>

      <div className="n-panel" style={{padding:18, marginBottom:16}}>
        <div className="n-eyebrow" style={{marginBottom:10}}>Documents</div>
        {instrument.documents.length===0 && <div style={{fontSize:12.5, color:"var(--ink-faint)"}}>No documents attached.</div>}
        {instrument.documents.map((d,i)=><div key={i} style={{fontSize:12.5, padding:"5px 0"}}>{d}</div>)}
      </div>

      <InstrumentAuditSummary db={db} instrument={instrument} nav={nav}/>
    </div>
  );
}
function InstrumentEnvHistory({db, instrument, tests}){
  const testIds = tests.map(t=>t.id);
  const readings = db.envReadings.filter(e=>testIds.includes(e.testId)).slice(0,6);
  return (
    <div className="n-panel" style={{padding:18, marginBottom:16}}>
      <div className="n-eyebrow" style={{marginBottom:10}}>Environmental Readings Linked to This Instrument's Tests</div>
      {readings.length===0 ? (
        <div style={{fontSize:12.5, color:"var(--ink-faint)"}}>No environmental readings linked yet. Capture from the IoT Test Bench during a test.</div>
      ) : (
        <table className="n-table">
          <thead><tr><th>Timestamp</th><th>Test</th><th>Temperature</th><th>Humidity</th><th>Sensor</th></tr></thead>
          <tbody>
            {readings.map(e=>{
              const test = tests.find(t=>t.id===e.testId);
              const anomalous = e.temperature>27.5;
              return (
                <tr key={e.id}>
                  <td className="f-mono">{fmtDT(e.timestamp)}</td>
                  <td>{test?.testName}</td>
                  <td className="f-mono" style={anomalous?{color:"var(--amber)", fontWeight:700}:{}}>{e.temperature}°C {anomalous && ""}</td>
                  <td className="f-mono">{e.humidity}%RH</td>
                  <td>{e.sensorId}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
    </div>
  );
}
function InstrumentAuditSummary({db, instrument, nav}){
  const relevantIds = new Set([instrument.id, ...db.tests.filter(t=>t.instrumentId===instrument.id).map(t=>t.id), ...db.reports.filter(r=>r.instrumentId===instrument.id).map(r=>r.id)]);
  const entries = db.auditLog.filter(a=>relevantIds.has(a.entityId)).slice(0,5);
  return (
    <div className="n-panel" style={{padding:18}}>
      <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10}}>
        <div className="n-eyebrow">Recent Audit Activity</div>
        <button className="n-btn n-btn-sm" onClick={()=>nav("audit")}>Open Full Audit Trail <ChevronRight size={13}/></button>
      </div>
      {entries.length===0 && <div style={{fontSize:12.5, color:"var(--ink-faint)"}}>No audit entries yet.</div>}
      {entries.map(a=>(
        <div key={a.id} style={{display:"flex", justifyContent:"space-between", fontSize:12, padding:"6px 4px", borderBottom:"1px solid var(--line)"}}>
          <span><strong>{a.action}</strong> <span style={{color:"var(--ink-faint)"}}>by {a.user}</span></span>
          <span className="f-mono" style={{color:"var(--ink-faint)"}}>{fmtDT(a.timestamp)}</span>
        </div>
      ))}
    </div>
  );
}
function EvaluationChecklist({instrument, plan, tests, report}){
  const steps = [
    {label:"Registration", state:"done"},
    {label:"Test Plan Generated", state: plan ? "done" : "pending"},
  ];
  if(plan){
    plan.tests.filter(pt=>pt.status==="applicable").forEach(pt=>{
      const t = tests.find(x=>x.testKey===pt.key);
      steps.push({label:pt.name, state: !t ? "pending" : (t.status==="approved"||t.status==="not-applicable") ? "done" : (t.status==="in_progress"||t.observations?.length>0) ? "active" : "pending"});
    });
  }
  steps.push({label:"Human Review", state: tests.some(t=>t.status==="review") ? "active" : tests.length>0 && tests.every(t=>t.status==="approved"||t.status==="not-applicable") ? "done" : "pending"});
  steps.push({label:"Report Generated", state: report ? "done" : "pending"});
  const doneCount = steps.filter(s=>s.state==="done").length;
  const pct = Math.round(doneCount/steps.length*100);
  return (
    <div className="n-panel" style={{padding:18, marginBottom:16}}>
      <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12}}>
        <div className="n-eyebrow">Evaluation Status</div>
        <div className="f-mono" style={{fontSize:12, color:"var(--ink-dim)"}}>{pct}% complete</div>
      </div>
      <div style={{height:6, background:"var(--surface-2)", marginBottom:14}}>
        <div style={{height:6, width:`${pct}%`, background:"var(--seal)"}}/>
      </div>
      <div style={{display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(180px,1fr))", gap:8}}>
        {steps.map((s,i)=>(
          <div key={i} style={{display:"flex", alignItems:"center", gap:7, fontSize:12}}>
            {s.state==="done" && <CheckCircle2 size={14} color="var(--seal)"/>}
            {s.state==="active" && <span style={{width:14,height:14,borderRadius:"50%", background:"var(--amber)", display:"inline-block", flexShrink:0}}/>}
            {s.state==="pending" && <span style={{width:14,height:14,borderRadius:"50%", border:"1.5px solid var(--line-strong)", display:"inline-block", flexShrink:0}}/>}
            <span style={{color: s.state==="pending"?"var(--ink-faint)":"var(--ink)", fontWeight: s.state==="done"?400:600}}>{s.label}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function ChainOfCustody({db, instrument, tests, report}){
  const log = db.auditLog.filter(a=>
    a.entityId===instrument.id ||
    tests.some(t=>t.id===a.entityId) ||
    (report && a.entityId===report.id) ||
    db.testPlans.find(tp=>tp.instrumentId===instrument.id)?.id===a.entityId
  );
  const findEvent = (matcher) => log.filter(matcher).sort((a,b)=>a.timestamp.localeCompare(b.timestamp))[0];

  const stages = [
    {label:"Received / Registered", event: findEvent(a=>a.action==="Instrument registered")},
    {label:"Test Plan Assigned", event: findEvent(a=>a.action==="Test plan generated")},
    {label:"Testing Commenced", event: findEvent(a=>a.action==="Observation recorded")},
    {label:"Submitted for Review", event: findEvent(a=>a.action==="Test submitted for review")},
    {label:"Reviewed / Approved", event: findEvent(a=>a.action==="Test approved")},
    {label:"Report Issued", event: findEvent(a=>a.action==="Report generated")},
  ];
  return (
    <div className="n-panel" style={{padding:18, marginBottom:16}}>
      <div className="n-eyebrow" style={{marginBottom:10}}>Digital Chain of Custody</div>
      <table className="n-table">
        <thead><tr><th>Stage</th><th>Status</th><th>Timestamp</th><th>Responsible User</th></tr></thead>
        <tbody>
          {stages.map((s,i)=>(
            <tr key={i}>
              <td>{s.label}</td>
              <td>{s.event ? <span className="n-badge n-badge-pass"><CheckCircle2 size={11}/> Completed</span> : <span className="n-badge n-badge-neutral">Pending</span>}</td>
              <td className="f-mono">{s.event ? fmtDT(s.event.timestamp) : "—"}</td>
              <td>{s.event ? s.event.user : "—"}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ---------- Measurement Fingerprint UI ---------- */
function FingerprintChart({points, width=280, height=90}){
  if(points.length===0) return null;
  const errors = points.map(p=>p.error);
  const min = Math.min(0, ...errors), max = Math.max(0, ...errors);
  const range = (max-min) || 1;
  const padX=16, padY=14;
  const stepX = points.length>1 ? (width-padX*2)/(points.length-1) : 0;
  const yFor = (v)=> height-padY - ((v-min)/range)*(height-padY*2);
  const xFor = (i)=> padX + i*stepX;
  const path = points.map((p,i)=>`${i===0?"M":"L"}${xFor(i)},${yFor(p.error)}`).join(" ");
  const zeroY = yFor(0);
  return (
    <svg width={width} height={height}>
      <line x1={padX} y1={zeroY} x2={width-padX} y2={zeroY} stroke="var(--line-strong)" strokeDasharray="2 2"/>
      <path d={path} fill="none" stroke="var(--navy-dim)" strokeWidth="1.6"/>
      {points.map((p,i)=>(
        <circle key={p.id} cx={xFor(i)} cy={yFor(p.error)} r={p.isCurrent?4.5:3.5}
          fill={p.isCurrent ? "var(--navy)" : "var(--seal)"} stroke="#fff" strokeWidth="1"/>
      ))}
    </svg>
  );
}
// Error/Deviation vs Load — compares current, previous and (where available) historical baseline
// across the standardized 0/25/50/75/100% load points actually recorded.
function FingerprintLoadChart({current, previous, baselinePoints, width=320, height=140}){
  const series = [];
  if(baselinePoints && baselinePoints.length>1) series.push({label:"Historical Baseline", color:"var(--line-strong)", pts:baselinePoints, dashed:true});
  if(previous?.loadPoints?.length>0) series.push({label:"Previous", color:"var(--seal)", pts:previous.loadPoints});
  if(current?.loadPoints?.length>0) series.push({label:"Current", color:"var(--navy)", pts:current.loadPoints});
  if(series.every(s=>s.pts.length===0)) return <div style={{fontSize:11.5, color:"var(--ink-faint)"}}>Not enough matching load points recorded to plot a comparison.</div>;

  const allErrors = series.flatMap(s=>s.pts.map(p=>p.error));
  const min = Math.min(0, ...allErrors), max = Math.max(0, ...allErrors);
  const range = (max-min) || 1;
  const padX=30, padY=18;
  const xFor = (pct)=> padX + (pct/100)*(width-padX*2);
  const yFor = (v)=> height-padY - ((v-min)/range)*(height-padY*2);
  const zeroY = yFor(0);
  return (
    <div>
      <svg width={width} height={height}>
        <line x1={padX} y1={zeroY} x2={width-padX} y2={zeroY} stroke="var(--line-strong)" strokeDasharray="2 2"/>
        {LOAD_PCTS.map(pct=>(
          <text key={pct} x={xFor(pct)} y={height-4} fontSize="8.5" textAnchor="middle" fill="var(--ink-faint)">{pct}%</text>
        ))}
        {series.map((s,si)=>{
          const sorted = [...s.pts].filter(p=>p.pct!=null).sort((a,b)=>a.pct-b.pct);
          if(sorted.length===0) return null;
          const path = sorted.map((p,i)=>`${i===0?"M":"L"}${xFor(p.pct)},${yFor(p.error)}`).join(" ");
          return (
            <g key={si}>
              <path d={path} fill="none" stroke={s.color} strokeWidth={s.dashed?1.2:1.8} strokeDasharray={s.dashed?"3 2":"none"}/>
              {sorted.map(p=><circle key={p.pct} cx={xFor(p.pct)} cy={yFor(p.error)} r="3" fill={s.color} stroke="#fff" strokeWidth="0.8"/>)}
            </g>
          );
        })}
      </svg>
      <div style={{display:"flex", gap:14, fontSize:10.5, color:"var(--ink-dim)", flexWrap:"wrap"}}>
        {series.map((s,i)=>(
          <span key={i} style={{display:"flex", alignItems:"center", gap:4}}>
            <span style={{width:10, height:2, background:s.color, display:"inline-block"}}/>{s.label}
          </span>
        ))}
      </div>
    </div>
  );
}
function FingerprintAnalysisModal({fp, instrument, onClose}){
  const meta = FP_STATUS_META[fp.status];
  const baselinePoints = fp.points.length>2 ? averageLoadPoints(fp.points.slice(0,-1)) : null;
  return (
    <Modal title="Measurement Pattern Analysis" onClose={onClose} width={520}>
      <div className="n-badge" style={{background:meta.bg, color:meta.color, borderColor:meta.border, marginBottom:14, fontSize:12, padding:"6px 12px"}}>{meta.icon} {meta.label}</div>
      <div style={{fontSize:12.5, color:"var(--ink-dim)", marginBottom:14}}>
        {fp.status==="INSUFFICIENT"
          ? "Not enough historical evaluations exist yet to establish a reliable measurement baseline for this instrument."
          : fp.status==="STABLE"
          ? "The current evaluation's deviation is consistent with this instrument's historical baseline — no meaningful change detected."
          : `Current evaluation shows a ${fp.status==="SIGNIFICANT"?"substantially larger":"larger"} deviation than the previous evaluation on record.`}
      </div>
      {fp.status!=="INSUFFICIENT" && (
        <>
          <div className="n-eyebrow" style={{marginBottom:6, fontSize:10}}>Error / Deviation vs Load</div>
          <FingerprintLoadChart current={fp.current} previous={fp.previous} baselinePoints={baselinePoints} width={470} height={150}/>
          {fp.keyFinding && (
            <div style={{fontSize:12, background:"var(--surface-2)", padding:"9px 12px", margin:"12px 0"}}>
              <strong>Key Finding:</strong> {fp.keyFinding}
            </div>
          )}
          <div className="n-kv" style={{background:"var(--surface-2)", padding:12, marginBottom:12}}>
            <dt>Previous evaluation</dt><dd>{fp.previous.error>=0?"+":""}{fp.previous.error} {fp.previous.unit} <span style={{color:"var(--ink-faint)"}}>({fmtDT(fp.previous.date)})</span></dd>
            <dt>Current evaluation</dt><dd>{fp.current.error>=0?"+":""}{fp.current.error} {fp.current.unit} <span style={{color:"var(--ink-faint)"}}>({fmtDT(fp.current.date)})</span></dd>
            <dt>Change (reference load)</dt><dd style={{color:meta.color, fontWeight:700}}>{fp.changeAbs>=0?"+":""}{fp.changeAbs} {fp.current.unit}</dd>
            <dt>Historical evaluations used</dt><dd>{fp.points.length}</dd>
            <dt>Historical baseline (mean)</dt><dd>{fp.baselineMean>=0?"+":""}{fp.baselineMean} {fp.current.unit}</dd>
            <dt>Historical range</dt><dd>{fp.baselineMin} to {fp.baselineMax} {fp.current.unit}</dd>
          </div>
          {(fp.previous.envTemp!=null && fp.current.envTemp!=null) && (
            <div style={{fontSize:11.5, color:"var(--ink-faint)", marginBottom:12}}>
              Environmental conditions recorded during the evaluations differ: previous {fp.previous.envTemp}°C / {fp.previous.envHumidity}%RH, current {fp.current.envTemp}°C / {fp.current.envHumidity}%RH. This is not asserted as the cause of the measurement change.
            </div>
          )}
          {fp.status!=="STABLE" && (
            <div style={{fontSize:12, background:meta.bg, color:meta.color, padding:"9px 12px", border:`1px solid ${meta.border}`}}>
              <strong>Recommendation:</strong> Review the instrument condition and test setup before the next evaluation.
            </div>
          )}
        </>
      )}
      <div style={{fontSize:10.5, color:"var(--ink-faint)", marginTop:14}}>This is a historical trend indicator only, computed from stored observations using the thresholds set in Historical Analysis Configuration. It does not alter or override the OIML R-76 rule engine's PASS/FAIL result, which requires human review for final approval.</div>
    </Modal>
  );
}
function averageLoadPoints(points){
  const byPct = {};
  points.forEach(p=>(p.loadPoints||[]).forEach(lp=>{ if(lp.pct==null) return; (byPct[lp.pct]=byPct[lp.pct]||[]).push(lp.error); }));
  return Object.entries(byPct).map(([pct,errs])=>({pct:Number(pct), error:round(errs.reduce((a,b)=>a+b,0)/errs.length,5)}));
}
function MeasurementFingerprint({db, instrument, nav}){
  const fp = useMemo(()=>computeFingerprint(db, instrument),[db, instrument]);
  const [analysisOpen, setAnalysisOpen] = useState(false);
  const meta = FP_STATUS_META[fp.status];
  const baselinePoints = fp.points.length>2 ? averageLoadPoints(fp.points.slice(0,-1)) : null;

  return (
    <div className="n-panel" style={{padding:20, marginBottom:16, borderLeft:`4px solid ${meta.color}`, boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
      <div style={{display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:14}}>
        <div>
          <div className="f-display" style={{fontSize:15, fontWeight:700, marginBottom:2}}>Measurement Fingerprint</div>
          <div style={{fontSize:11, color:"var(--ink-faint)"}}>Historical trend indicator — decision-support only, separate from the official OIML result</div>
        </div>
        <span className="n-badge" style={{background:meta.bg, color:meta.color, borderColor:meta.border, fontSize:12, padding:"5px 11px"}}>{meta.icon} {meta.label}</span>
      </div>

      {fp.status==="INSUFFICIENT" ? (
        <div style={{fontSize:12.5, color:"var(--ink-dim)", padding:"10px 0"}}>
          <strong>Insufficient historical data.</strong> Complete additional evaluations to establish a reliable measurement baseline.
          {fp.points.length===1 && <div style={{marginTop:6, fontSize:11.5, color:"var(--ink-faint)"}}>1 evaluation on record: {fp.points[0].error>=0?"+":""}{fp.points[0].error} {fp.points[0].unit} ({fmtDT(fp.points[0].date)})</div>}
        </div>
      ) : (
        <>
          {fp.status!=="STABLE" && (
            <div style={{background:meta.bg, border:`1px solid ${meta.border}`, padding:"9px 12px", marginBottom:14, fontSize:12}}>
              <strong style={{color:meta.color}}>MEASUREMENT PATTERN CHANGE DETECTED</strong>
              <div style={{color:meta.color, marginTop:3}}>Current measurement behavior differs from the instrument's previous evaluation pattern.</div>
              {fp.keyFinding && <div style={{color:meta.color, marginTop:4, fontSize:11.5}}>{fp.keyFinding}</div>}
            </div>
          )}
          <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:20}}>
            <div>
              <div style={{display:"flex", gap:24, marginBottom:10}}>
                <div>
                  <div className="n-label">Current</div>
                  <div className="f-mono" style={{fontSize:18, fontWeight:700}}>{fp.current.error>=0?"+":""}{fp.current.error} {fp.current.unit}</div>
                  <div style={{fontSize:10.5, color:"var(--ink-faint)"}}>{fp.current.source==="iot"?"IoT Captured":"✏️ Manual Entry"}</div>
                </div>
                <div>
                  <div className="n-label">Previous</div>
                  <div className="f-mono" style={{fontSize:18, fontWeight:700, color:"var(--ink-dim)"}}>{fp.previous.error>=0?"+":""}{fp.previous.error} {fp.previous.unit}</div>
                  <div style={{fontSize:10.5, color:"var(--ink-faint)"}}>{fmtDT(fp.previous.date)}</div>
                </div>
              </div>
              <div className="n-label" style={{marginBottom:4}}>Error vs Time</div>
              <FingerprintChart points={fp.points}/>
            </div>
            <div>
              <div className="n-label" style={{marginBottom:4}}>Error / Deviation vs Load</div>
              <FingerprintLoadChart current={fp.current} previous={fp.previous} baselinePoints={baselinePoints}/>
            </div>
          </div>
          <div className="n-hr" style={{margin:"14px 0"}}/>
          <div style={{display:"flex", justifyContent:"space-between", alignItems:"center"}}>
            <div style={{display:"flex", gap:18, fontSize:11.5, color:"var(--ink-dim)"}}>
              <span><strong>{fp.points.length}</strong> evaluations</span>
              <span>Historical Consistency: <strong style={{color:meta.color}}>{meta.consistency}</strong></span>
            </div>
            <button className="n-btn n-btn-sm" onClick={()=>setAnalysisOpen(true)}>View Analysis</button>
          </div>
        </>
      )}
      {analysisOpen && <FingerprintAnalysisModal fp={fp} instrument={instrument} onClose={()=>setAnalysisOpen(false)}/>}
    </div>
  );
}
function EvaluationHistoryTable({db, instrument, nav}){
  const fp = computeFingerprint(db, instrument);
  if(fp.points.length===0) return null;
  return (
    <div className="n-panel" style={{padding:18, marginBottom:16}}>
      <div className="n-eyebrow" style={{marginBottom:10}}>Evaluation History</div>
      <table className="n-table">
        <thead><tr><th>Date</th><th>Test</th><th>Measurement Error</th><th>Source</th><th>Status</th><th></th></tr></thead>
        <tbody>
          {[...fp.points].reverse().map(p=>(
            <tr key={p.id} className={p.refTestId?"n-row-hover":""} onClick={()=> p.refTestId && nav("tests",{testId:p.refTestId})}>
              <td className="f-mono">{fmtDT(p.date)}</td>
              <td>{TEST_DEFS[p.testKey]?.name || p.testKey}</td>
              <td className="f-mono">{p.error>=0?"+":""}{p.error} {p.unit}</td>
              <td>{p.source==="iot"?"IoT":"Manual"}</td>
              <td>{p.resultLabel ? <ResultBadge result={p.resultLabel}/> : "—"}</td>
              <td>{p.refTestId ? <ChevronRight size={13} color="var(--ink-faint)"/> : <span style={{fontSize:10, color:"var(--ink-faint)"}}>Reference only</span>}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
function MeasurementEvolutionTimeline({db, instrument}){
  const fp = computeFingerprint(db, instrument);
  if(fp.points.length===0) return null;
  const rows = fp.points.map((p,i)=>{
    if(i===0) return {...p, label:"Baseline Established", icon:""};
    const prev = fp.points[i-1];
    const change = Math.abs(p.error-prev.error);
    const e = instrument.e;
    if(change>5*e) return {...p, label:"Significant Change Detected", icon:""};
    if(change>2*e) return {...p, label:"Pattern Change Detected", icon:""};
    return {...p, label:"Stable Evaluation", icon:""};
  });
  return (
    <div className="n-panel" style={{padding:18, marginBottom:16}}>
      <div className="n-eyebrow" style={{marginBottom:12}}>Measurement Evolution</div>
      <div style={{display:"flex", flexDirection:"column", gap:0}}>
        {rows.map((r,i)=>(
          <div key={r.id} style={{display:"flex", gap:12, alignItems:"flex-start"}}>
            <div style={{display:"flex", flexDirection:"column", alignItems:"center"}}>
              <span style={{fontSize:16}}>{r.icon}</span>
              {i<rows.length-1 && <div style={{width:1, flex:1, minHeight:24, background:"var(--line-strong)"}}/>}
            </div>
            <div style={{paddingBottom:16}}>
              <div style={{fontSize:12.5, fontWeight:600}}>{r.label} {r.isCurrent && <span className="n-badge n-badge-navy" style={{fontSize:9, marginLeft:4}}>CURRENT</span>}</div>
              <div style={{fontSize:11, color:"var(--ink-faint)"}}>{fmtDT(r.date)} · {r.error>=0?"+":""}{r.error} {r.unit}</div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function CompareWithHistory({db, instrument, test, nav}){
  const fp = computeFingerprint(db, instrument);
  const meta = FP_STATUS_META[fp.status];
  const officialMeta = test.result==="PASS" ? {label:"PASS", icon:"", color:"var(--seal)"} : test.result==="FAIL" ? {label:"FAIL", icon:"", color:"var(--rose)"} : {label:"REVIEW", icon:"", color:"var(--amber)"};
  return (
    <div className="n-panel" style={{padding:18, marginBottom:16}}>
      <div className="n-eyebrow" style={{marginBottom:12}}>Historical Comparison — Human Review Bridge</div>
      <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:14}}>
        <div style={{padding:12, background:"var(--surface-2)"}}>
          <div className="n-label" style={{marginBottom:5}}>Official OIML Result</div>
          <div style={{fontSize:15, fontWeight:700, color:officialMeta.color}}>{officialMeta.icon} {officialMeta.label}</div>
          <div style={{fontSize:11, color:"var(--ink-faint)", marginTop:3}}>From the configured OIML R-76 rule engine — rule-based, not a trend indicator.</div>
        </div>
        <div style={{padding:12, background: fp.status==="INSUFFICIENT" ? "var(--surface-2)" : meta.bg}}>
          <div className="n-label" style={{marginBottom:5}}>Historical Measurement Behavior</div>
          <div style={{fontSize:15, fontWeight:700, color: fp.status==="INSUFFICIENT" ? "var(--ink-faint)" : meta.color}}>{meta.icon} {meta.label}</div>
          <div style={{fontSize:11, color:"var(--ink-faint)", marginTop:3}}>Decision-support only — from this instrument's measurement fingerprint.</div>
        </div>
      </div>

      {fp.status==="INSUFFICIENT" || !fp.previous ? (
        <div style={{fontSize:12.5, color:"var(--ink-faint)"}}>No comparable historical evaluation is currently available. Complete additional evaluations to establish a Measurement Fingerprint baseline.</div>
      ) : (
        <>
          <div style={{display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14}}>
            <div>
              <div className="n-label">Current Evaluation</div>
              <div className="f-mono" style={{fontSize:16, fontWeight:700}}>{fp.current.error>=0?"+":""}{fp.current.error} {fp.current.unit}</div>
            </div>
            <div>
              <div className="n-label">Previous Evaluation</div>
              <div className="f-mono" style={{fontSize:16, fontWeight:700, color:"var(--ink-dim)"}}>{fp.previous.error>=0?"+":""}{fp.previous.error} {fp.previous.unit}</div>
              <div style={{fontSize:10.5, color:"var(--ink-faint)"}}>{fmtDT(fp.previous.date)}</div>
            </div>
            <div>
              <div className="n-label">Historical Average / Range</div>
              <div className="f-mono" style={{fontSize:13}}>{fp.baselineMean>=0?"+":""}{fp.baselineMean} {fp.current.unit}</div>
              <div style={{fontSize:10.5, color:"var(--ink-faint)"}}>{fp.baselineMin} to {fp.baselineMax} {fp.current.unit}</div>
            </div>
          </div>
          {fp.keyFinding && <div style={{marginTop:10, fontSize:11.5, color:"var(--ink-dim)"}}>{fp.keyFinding}</div>}
          <div style={{marginTop:12, display:"flex", justifyContent:"space-between", alignItems:"center"}}>
            <div style={{fontSize:11.5, color:"var(--ink-faint)"}}>Official Result: <strong style={{color:officialMeta.color}}>{officialMeta.label} {officialMeta.icon}</strong> &nbsp;·&nbsp; Historical Behavior: <strong style={{color:meta.color}}>{meta.label} {meta.icon}</strong></div>
            <button className="n-btn n-btn-sm" onClick={()=>nav("instruments",{instrumentId:instrument.id})}>View Full Fingerprint</button>
          </div>
        </>
      )}
    </div>
  );
}

/* ============================== TESTS LIST ============================== */
function TestsPage({db, nav, role, presetFilter}){
  const [filter, setFilter] = useState(presetFilter||"all");
  const [q, setQ] = useState("");
  useEffect(()=>{ setFilter(presetFilter||"all"); },[presetFilter]);
  const tests = db.tests.filter(t=>{
    if(filter==="review" && t.status!=="review") return false;
    if(filter==="pending" && !(t.status==="pending"||t.status==="in_progress")) return false;
    if(filter==="approved" && t.status!=="approved") return false;
    if(filter==="failed" && t.result!=="FAIL") return false;
    const inst = db.instruments.find(i=>i.id===t.instrumentId);
    if(q && !(`${inst?.model} ${inst?.serialNumber} ${t.testName}`.toLowerCase().includes(q.toLowerCase()))) return false;
    return true;
  });
  const FILTER_LABELS = {all:"All Tests", pending:"Pending / In Progress", review:"Pending Reviews", approved:"Approved", failed:"Failed"};
  return (
    <div>
      <PageHeader eyebrow="Test Management" title={filter==="review" ? "Pending Reviews" : "Tests"}
        desc={filter==="review" ? "Tests awaiting Examiner or Authority review. Open a test to inspect observations, calculations and evidence before deciding." : "All test executions across registered instruments. Select a test to enter observations, review, or approve."}/>
      <div style={{display:"flex", gap:8, marginBottom:14, alignItems:"center", flexWrap:"wrap"}}>
        <div style={{position:"relative", flex:1, maxWidth:280}}>
          <Search size={13} style={{position:"absolute", left:9, top:10, color:"var(--ink-faint)"}}/>
          <input className="n-input" style={{paddingLeft:28}} placeholder="Search model, serial, test…" value={q} onChange={e=>setQ(e.target.value)}/>
        </div>
        {Object.entries(FILTER_LABELS).map(([k,l])=>(
          <button key={k} className="n-btn n-btn-sm" style={filter===k?{background:"var(--navy)", color:"#fff", borderColor:"var(--navy)"}:{}} onClick={()=>setFilter(k)}>{l}</button>
        ))}
      </div>
      <div className="n-panel">
        <table className="n-table">
          <thead><tr><th>Test</th><th>Instrument</th><th>Status</th><th>Result</th><th>Technician</th><th>Updated</th><th></th></tr></thead>
          <tbody>
            {tests.map(t=>{
              const inst = db.instruments.find(i=>i.id===t.instrumentId);
              return (
                <tr key={t.id} className="n-row-hover" onClick={()=>nav("tests",{testId:t.id})}>
                  <td><strong>{t.testName}</strong><div className="f-mono" style={{color:"var(--ink-faint)", fontSize:11}}>{t.id}</div></td>
                  <td>{inst?.model} <span style={{color:"var(--ink-faint)"}}>({inst?.serialNumber})</span></td>
                  <td><StatusBadge status={t.status}/></td>
                  <td><ResultBadge result={t.result}/></td>
                  <td>{t.technician||"—"}</td>
                  <td className="f-mono">{fmtDT(t.createdAt)}</td>
                  <td><ChevronRight size={14} color="var(--ink-faint)"/></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== TEST WORKSPACE ============================== */
function TestWorkspace({db, mutate, logAudit, testId, nav, notify, role, currentUserName, syncStatus}){
  const test = db.tests.find(t=>t.id===testId);
  const instrument = db.instruments.find(i=>i.id===test.instrumentId);
  const rule = RULE_VERSIONS.find(r=>r.id===test.ruleVersionId) || activeRule();
  const [loadInput, setLoadInput] = useState("");
  const [obsInput, setObsInput] = useState("");
  const [voiceText, setVoiceText] = useState("");
  const [voiceParsed, setVoiceParsed] = useState(null);
  const [remarksDraft, setRemarksDraft] = useState(test.remarks||"");
  const [measureSource, setMeasureSource] = useState("manual");
  const [naOpen, setNaOpen] = useState(false);
  const [naReason, setNaReason] = useState(test.naReason||"");
  const markNotApplicable = () => {
    if(!can(role,"test:mark-na")){ notify("Your role does not permit marking a test Not Applicable."); return; }
    if(!naReason.trim()){ notify("A reason is required to mark this test Not Applicable."); return; }
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      const prev = dt.status; dt.status = "not-applicable"; dt.naReason = naReason.trim(); dt.result = null; dt.technician = currentUserName;
      logAudit(d,"Test marked Not Applicable","Test",dt.id,prev,"not-applicable",naReason.trim());
    });
    notify("Test marked Not Applicable.");
  };
  const revertNotApplicable = () => {
    if(!can(role,"test:mark-na")){ notify("Your role does not permit reinstating this test."); return; }
    mutate(d=>{ const dt = d.tests.find(x=>x.id===test.id); const prev=dt.status; dt.status="pending"; dt.naReason=""; logAudit(d,"Test reinstated as applicable","Test",dt.id,prev,"pending","Not-Applicable marking reversed"); });
  };
  const envForTest = db.envReadings.filter(e=>e.testId===test.id);
  const latestEnv = envForTest[0];
  const plan = db.testPlans.find(p=>p.instrumentId===instrument.id);
  const allTests = db.tests.filter(t=>t.instrumentId===instrument.id);
  const scaleDevice = db.iotDevices.find(d=>d.id==="DEV-SCALE");
  const tempDevice = db.iotDevices.find(d=>d.id==="DEV-TEMP");
  const humDevice = db.iotDevices.find(d=>d.id==="DEV-HUM");
  const equipmentList = db.referenceEquipment || [];
  const equipment = equipmentList.find(eq=>eq.id===test.referenceEquipmentId)
    || equipmentList.find(eq=>parseFloat(eq.capacity)>=instrument.maxCapacity)
    || equipmentList[0] || null;
  const eqFlags = equipmentStatus(equipment);
  const classFlags = validateClassConfig(instrument, RULE_VERSIONS.find(r=>r.id===test.ruleVersionId)||activeRule());

  const readiness = [
    {label:"Instrument Details Complete", ok: !!(instrument.manufacturer && instrument.model && instrument.serialNumber)},
    {label:"Technical Parameters Valid", ok: !classFlags.some(f=>f.level==="bad"), note: classFlags.filter(f=>f.level==="bad").map(f=>f.msg).join(" ")},
    {label:"Accuracy Class Selected", ok: !!instrument.accuracyClass},
    {label:"OIML Rule Configuration Selected", ok: rule?.status==="ACTIVE", note: rule?.status!=="ACTIVE"?"Test is bound to a superseded rule version.":""},
    {label:"Reference Equipment Valid", ok: !eqFlags.some(f=>f.level==="bad"), note: eqFlags.filter(f=>f.level==="bad").map(f=>f.msg).join(" ")},
    {label:"Calibration Valid", ok: equipment && !eqFlags.some(f=>f.level==="bad" && f.msg.includes("Calibration")), note: equipment?`Valid until ${equipment.validUntil}`:"No equipment selected"},
    {label:"Environment Recorded", ok: envForTest.length>0, note: envForTest.length===0?"No environmental readings linked to this test yet":""},
    {label:"Applicable Test Plan", ok: !!plan},
    {label:"Required Documents", ok: instrument.documents.length>0, note: instrument.documents.length===0?"No documents attached to instrument yet":""},
  ];
  const readinessIssues = readiness.filter(r=>!r.ok);

  const submission = validateTestSubmission(instrument, test, {envReadings: envForTest, equipment});

  const existingLoads = (test.observations||[]).map(o=>o.load).filter(v=>v!==undefined);
  const flags = (loadInput!==""&&obsInput!=="") ? validateObservation(instrument, test.testKey, {load:Number(loadInput), observed:Number(obsInput)}, existingLoads) : [];
  const hasBad = flags.some(f=>f.level==="bad");

  const addAccuracyObs = (loadVal, obsVal, source="manual") => {
    if(!can(role,"test:record")){ notify("Your role does not permit recording measurements."); return; }
    if(isLockedForEntry(test)){ notify("This test has been submitted and its measurements are locked."); return; }
    const load = Number(loadVal), observed = Number(obsVal);
    const vf = validateObservation(instrument, test.testKey, {load, observed}, existingLoads);
    if(vf.some(f=>f.level==="bad")){ notify("Observation blocked — resolve validation errors first."); return; }
    const calc = evalAccuracy(load, observed, instrument, rule);
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      dt.observations.push({id:uid("OBS"), load, observed, unit:instrument.capacityUnit, source, timestamp:nowISO(), syncFlag: syncStatus==="synced"?"synced":"pending", ...calc});
      dt.status = "in_progress"; dt.technician = currentUserName;
      const results = dt.observations.map(o=>o.result);
      dt.result = results.includes("FAIL") ? "FAIL" : "PASS";
      logAudit(d,"Observation recorded","Test",dt.id,"—",`load ${load}${instrument.capacityUnit} → ${observed}${instrument.capacityUnit} (${calc.result})`, source==="iot"?"Captured via IoT weighing device":source==="voice"?"Captured via voice input":"Manual entry");
    });
    setLoadInput(""); setObsInput("");
  };

  const addRepeatReading = (val, source="manual") => {
    if(!can(role,"test:record")){ notify("Your role does not permit recording measurements."); return; }
    if(isLockedForEntry(test)){ notify("This test has been submitted and its measurements are locked."); return; }
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      const load = 0.5*instrument.maxCapacity;
      dt.observations.push({id:uid("OBS"), load, observed:Number(val), unit:instrument.capacityUnit, source, timestamp:nowISO(), syncFlag: syncStatus==="synced"?"synced":"pending"});
      dt.status = "in_progress"; dt.technician = currentUserName;
      const readings = dt.observations.map(o=>o.observed);
      const calc = evalRepeatability(readings, load, instrument, rule);
      dt.repeatCalc = calc; dt.result = calc?.result || null;
      const anomaly = detectRepeatAnomaly(readings);
      dt.anomalyMsg = anomaly;
      logAudit(d,"Observation recorded","Test",dt.id,"—",`repeat reading ${val}${instrument.capacityUnit}`, source==="iot"?"Captured via IoT weighing device":"Manual entry");
    });
  };

  const autoCapture5 = () => {
    const base = 0.5*instrument.maxCapacity;
    for(let i=0;i<5;i++){
      const noise = (Math.random()-0.5)*instrument.e*1.6;
      addRepeatReading(round(base+noise,4), "iot");
    }
  };

  const addEvidence = (type, file) => {
    if(!can(role,"evidence:upload")){ notify("Your role does not permit uploading evidence."); return; }
    if(!file) return;
    const maxBytes = 8 * 1024 * 1024;
    if(file.size > maxBytes){ notify("File is too large. Maximum supported size is 8 MB."); return; }
    const allowed = type === "document"
      ? ["application/pdf","application/msword","application/vnd.openxmlformats-officedocument.wordprocessingml.document","text/plain"]
      : ["image/jpeg","image/png","image/webp"];
    if(!allowed.includes(file.type)){
      notify(type === "document" ? "Use PDF, Word, or text documents." : "Use JPG, PNG, or WebP images."); return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      mutate(d=>{
        const dt=d.tests.find(x=>x.id===test.id); dt.evidence=dt.evidence||[];
        dt.evidence.push({id:uid("EVD"),type,label:file.name,originalName:file.name,mimeType:file.type,size:file.size,dataUrl:reader.result,addedAt:nowISO(),addedBy:currentUserName});
        logAudit(d,"Evidence attached","Test",dt.id,"—",file.name,`Uploaded ${type === "document" ? "supporting document" : "photograph"}`);
      });
      notify(`${file.name} attached to this test.`);
    };
    reader.onerror=()=>notify("The file could not be read. Please try again.");
    reader.readAsDataURL(file);
  };
  const removeEvidence = (id) => {
    if(!can(role,"evidence:upload")){ notify("Your role does not permit changing evidence."); return; }
    mutate(d=>{ const dt=d.tests.find(x=>x.id===test.id); const ev=(dt.evidence||[]).find(x=>x.id===id); dt.evidence=(dt.evidence||[]).filter(x=>x.id!==id); if(ev) logAudit(d,"Evidence removed","Test",dt.id,ev.originalName||ev.label,"—","Evidence removed from test record"); });
    notify("Evidence removed.");
  };
  const evidenceInputRef = useRef(null);
  const [evidenceType,setEvidenceType] = useState("photo-instrument");
  const openEvidencePicker = (type) => {
    if(!can(role,"evidence:upload")){ notify("Your role does not permit uploading evidence."); return; }
    setEvidenceType(type);
    if(evidenceInputRef.current){ evidenceInputRef.current.accept=type === "document" ? ".pdf,.doc,.docx,.txt" : "image/jpeg,image/png,image/webp"; evidenceInputRef.current.click(); }
  };
  const pendingSyncCount = test.observations.filter(o=>o.syncFlag==="pending").length;

  const submitReview = () => {
    if(!can(role,"test:submit")){ notify("Your role does not permit submitting tests for review."); return; }
    if(isLockedForEntry(test)){ notify("This test has already been submitted."); return; }
    if(!submission.canSubmit){ notify("Cannot submit — resolve the blocking validation errors first."); return; }
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      const prev = dt.status; dt.status = "review"; dt.remarks = remarksDraft;
      logAudit(d,"Test submitted for review","Test",dt.id,prev,"review", dt.result==="FAIL"?"Result FAIL — escalated to Examiner":"Routine submission for examiner review");
    });
    notify("Test submitted for review.");
  };
  const reviewerAction = (action) => {
    if(!can(role,"test:review")){ notify("Your role does not permit reviewing tests."); return; }
    if(test.status!=="review"){ notify("This test is not currently awaiting review."); return; }
    // NOTE: the problem statement's workflow separates Examiner technical approval from a
    // later, distinct Authority final approval. This prototype's single-stage review step
    // does not yet model that split (tracked separately, along with submission-locking, as a
    // remaining item) — today, whichever permitted role (Examiner/Authority/Admin) approves
    // a test is recorded as both reviewer and approver, and the role that acted is captured
    // in the audit trail for traceability.
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      const prev = dt.status;
      if(action==="approve"){ dt.status="approved"; dt.reviewer=currentUserName; dt.approver=currentUserName; dt.approvedAt=nowISO(); dt.reviewedAt=nowISO(); }
      if(action==="reject"){ dt.status="rejected"; dt.reviewer=currentUserName; dt.reviewedAt=nowISO(); }
      if(action==="correction"){ dt.status="in_progress"; dt.reviewer=currentUserName; dt.reviewedAt=nowISO(); }
      logAudit(d,`Test ${action==="approve"?"approved":action==="reject"?"rejected":"sent back for correction"} (by ${role})`,"Test",dt.id,prev,dt.status,remarksDraft||"Reviewer decision");
    });
    notify(`Test ${action==="approve"?"approved":action==="reject"?"rejected":"returned for correction"}.`);
  };

  const parseVoice = () => {
    const nums = voiceText.match(/[\d.]+/g);
    if(!nums || nums.length<2){ notify("Could not parse two numeric values from the phrase."); return; }
    setVoiceParsed({load:nums[0], observed:nums[1]});
  };
  const confirmVoice = () => {
    if(test.testKey==="ACC") addAccuracyObs(voiceParsed.load, voiceParsed.observed, "voice");
    else if(test.testKey==="REP") addRepeatReading(voiceParsed.observed, "voice");
    setVoiceParsed(null); setVoiceText("");
  };

  const canReview = role==="Examiner"||role==="Authority"||role==="Admin";
  const canSubmit = role==="Technician"||role==="Admin";
  const saveAndContinueLater = () => { notify("Progress saved. This test remains in progress and can be resumed anytime."); nav("tests"); };

  return (
    <div>
      <Breadcrumb items={[{label:"Tests", onClick:()=>nav("tests")}, {label:instrument.model, onClick:()=>nav("instruments",{instrumentId:instrument.id})}, {label:test.testName}]}/>
      <button className="n-btn n-btn-sm n-btn-ghost" style={{marginBottom:12}} onClick={()=>nav("tests")}><ChevronLeft size={14}/> Back to Tests</button>

      <div style={{display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:16}}>
        <div>
          <div className="n-eyebrow">{instrument.model} — {instrument.serialNumber}</div>
          <div className="f-display" style={{fontSize:20, fontWeight:700, margin:"3px 0"}}>{test.testName}</div>
          <div style={{fontSize:12, color:"var(--ink-dim)"}}>{TEST_DEFS[test.testKey].purpose}</div>
        </div>
        <div style={{textAlign:"right"}}>
          <StatusBadge status={test.status}/> <ResultBadge result={test.result}/>
          <div style={{fontSize:10.5, color:"var(--ink-faint)", marginTop:6}}>Rule: {rule.id}</div>
          {pendingSyncCount>0 && <div style={{fontSize:10.5, color:"var(--amber)", marginTop:3, fontWeight:600}}><RefreshCw size={11} style={{verticalAlign:"-1px"}}/> {pendingSyncCount} reading{pendingSyncCount>1?"s":""} pending sync</div>}
        </div>
      </div>

      {syncStatus!=="synced" && test.observations.length>0 && (
        <div className="n-panel" style={{padding:"10px 14px", marginBottom:14, borderColor:"var(--amber-line)", background:"var(--amber-bg)", display:"flex", gap:8, alignItems:"center"}}>
          {syncStatus==="pending" ? <RefreshCw size={16} color="var(--amber)"/> : <WifiOff size={16} color="var(--amber)"/>}
          <div style={{fontSize:12.5, color:"var(--amber)"}}><strong>Saved Locally — Pending Synchronization.</strong> Observations captured while offline are stored on this device and will sync automatically once connectivity is restored. Change status under Settings.</div>
        </div>
      )}

      <EvaluationChecklist instrument={instrument} plan={plan} tests={allTests} report={db.reports.find(r=>r.instrumentId===instrument.id)}/>

      {test.status==="pending" && test.observations.length===0 && (
        <div className="n-panel" style={{padding:16, marginBottom:16}}>
          <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom: readinessIssues.length?10:0}}>
            <div className="n-eyebrow">Test Readiness</div>
            {readinessIssues.length===0
              ? <span className="n-badge n-badge-pass">READY TO START TEST</span>
              : <span className="n-badge n-badge-review">{readinessIssues.length} ITEM{readinessIssues.length>1?"S":""} REQUIRE ATTENTION</span>}
          </div>
          <div style={{display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(200px,1fr))", gap:7}}>
            {readiness.map((r,i)=>(
              <div key={i} style={{display:"flex", alignItems:"center", gap:6, fontSize:12}} title={r.note}>
                {r.ok ? <CheckCircle2 size={13} color="var(--seal)"/> : <AlertTriangle size={13} color="var(--amber)"/>}
                <span style={{color: r.ok?"var(--ink)":"var(--amber)"}}>{r.label}</span>
              </div>
            ))}
          </div>
          {readinessIssues.length>0 && <div style={{fontSize:11.5, color:"var(--ink-faint)", marginTop:8}}>{readinessIssues.map(r=>r.note).filter(Boolean).join(" · ")}</div>}
        </div>
      )}

      {test.anomalyMsg && (
        <div className="n-panel" style={{padding:"10px 14px", marginBottom:14, borderColor:"var(--amber-line)", background:"var(--amber-bg)", display:"flex", gap:8, alignItems:"center"}}>
          <AlertTriangle size={16} color="var(--amber)"/>
          <div style={{fontSize:12.5, color:"var(--amber)"}}><strong>Observation Pattern Requires Review.</strong> {test.anomalyMsg}</div>
        </div>
      )}
      {(test.envFlag || (latestEnv && Math.abs(latestEnv.temperature-24)>3)) && (
        <div className="n-panel" style={{padding:"10px 14px", marginBottom:14, borderColor:"var(--amber-line)", background:"var(--amber-bg)", display:"flex", gap:8, alignItems:"center"}}>
          <Thermometer size={16} color="var(--amber)"/>
          <div style={{fontSize:12.5, color:"var(--amber)"}}><strong>Environmental Variation Detected.</strong> Significant ambient variation was recorded during this test. Review test conditions and validity before approval.</div>
        </div>
      )}

      <div style={{display:"grid", gridTemplateColumns:"1fr 300px", gap:16}}>
        <div>
          {/* OBSERVATION ENTRY */}
          <div className="n-panel" style={{padding:18, marginBottom:16}}>
            <div className="n-eyebrow" style={{marginBottom:10}}>Test Observations</div>

            {test.testKey==="ACC" && (
              <>
                <MeasurementSourceTabs source={measureSource} setSource={setMeasureSource}/>

                {measureSource==="manual" && (
                  <div style={{display:"flex", gap:8, alignItems:"flex-end", marginBottom:10, flexWrap:"wrap"}}>
                    <Field label={`Test Load (${instrument.capacityUnit})`}><input className="n-input" style={{width:120}} type="number" value={loadInput} onChange={e=>setLoadInput(e.target.value)}/></Field>
                    <Field label={`Observed Value (${instrument.capacityUnit})`}><input className="n-input" style={{width:140}} type="number" step="0.001" value={obsInput} onChange={e=>setObsInput(e.target.value)}/></Field>
                    <button className="n-btn n-btn-sm" disabled={hasBad||!loadInput||!obsInput} onClick={()=>addAccuracyObs(loadInput,obsInput,"manual")}>Add Reading</button>
                  </div>
                )}
                {measureSource==="iot" && (
                  <IotCapturePanel scaleDevice={scaleDevice} tempDevice={tempDevice} humDevice={humDevice}
                    onCapture={()=>{
                      const points = [0, instrument.minCapacity, 0.25*instrument.maxCapacity, 0.5*instrument.maxCapacity, 0.75*instrument.maxCapacity, instrument.maxCapacity];
                      const done = test.observations.map(o=>o.load);
                      const next = points.find(p=>!done.includes(round(p,3))) ?? points[points.length-1];
                      const val = round(next + (Math.random()-0.5)*instrument.e*1.4, 4);
                      addAccuracyObs(round(next,3), val, "iot");
                    }}/>
                )}
                {measureSource==="voice" && (
                  <VoiceInput voiceText={voiceText} setVoiceText={setVoiceText} onParse={parseVoice} parsed={voiceParsed} onConfirm={confirmVoice} onCancel={()=>setVoiceParsed(null)} example="Test load ten kilograms, observed value ten point zero one two kilograms."/>
                )}
                {measureSource==="manual" && flags.length>0 && (
                  <div style={{marginBottom:10}}>
                    {flags.map((f,i)=>(
                      <div key={i} style={{display:"flex", gap:6, alignItems:"center", fontSize:11.5, marginBottom:3, color: f.level==="bad"?"var(--rose)":f.level==="warn"?"var(--amber)":"var(--seal)"}}>
                        <FlagIcon level={f.level}/> {f.msg}
                      </div>
                    ))}
                  </div>
                )}
                <table className="n-table">
                  <thead><tr><th>Load</th><th>Observed</th><th>Error</th><th>MPE</th><th>Source</th><th>Result</th><th></th></tr></thead>
                  <tbody>
                    {test.observations.map(o=>(
                      <React.Fragment key={o.id}>
                        <tr>
                          <td className="f-mono">{o.load} {instrument.capacityUnit}</td>
                          <td className="f-mono">{o.observed} {instrument.capacityUnit}</td>
                          <td className="f-mono">{o.error>=0?"+":""}{o.error}</td>
                          <td className="f-mono">±{o.mpeVal}</td>
                          <td>{o.source==="iot" ? <><Radio size={11}/> IoT</> : o.source==="voice" ? <><Mic size={11}/> Voice</> : "Manual"}</td>
                          <td><ResultBadge result={o.result}/></td>
                          <td></td>
                        </tr>
                        <tr>
                          <td colSpan={7} style={{padding:"0 0 8px 0", border:"none"}}>
                            <ExplainCalculation calc={o} instrument={instrument} kind="accuracy"/>
                          </td>
                        </tr>
                      </React.Fragment>
                    ))}
                    {test.observations.length===0 && <tr><td colSpan={7} style={{textAlign:"center", color:"var(--ink-faint)", padding:16}}>No observations recorded yet.</td></tr>}
                  </tbody>
                </table>
              </>
            )}

            {test.testKey==="REP" && (
              <>
                <div style={{fontSize:12, color:"var(--ink-dim)", marginBottom:10}}>Fixed test load: <strong className="f-mono">{round(0.5*instrument.maxCapacity,3)} {instrument.capacityUnit}</strong> (50% of maximum capacity)</div>
                <MeasurementSourceTabs source={measureSource} setSource={setMeasureSource}/>
                {measureSource==="manual" && (
                  <div style={{display:"flex", gap:8, marginBottom:10}}>
                    <input className="n-input" style={{width:160}} type="number" step="0.001" placeholder="Observed value" value={obsInput} onChange={e=>setObsInput(e.target.value)}/>
                    <button className="n-btn n-btn-sm" disabled={!obsInput} onClick={()=>{addRepeatReading(obsInput); setObsInput("");}}>Add Reading</button>
                  </div>
                )}
                {measureSource==="iot" && (
                  <IotCapturePanel scaleDevice={scaleDevice} tempDevice={tempDevice} humDevice={humDevice}
                    onCapture={()=>{
                      const load = 0.5*instrument.maxCapacity;
                      const val = round(load + (Math.random()-0.5)*instrument.e*1.6, 4);
                      addRepeatReading(val, "iot");
                    }} multiAction={{label:"Auto-Capture 5 Readings", onClick:autoCapture5}}/>
                )}
                {measureSource==="voice" && (
                  <VoiceInput voiceText={voiceText} setVoiceText={setVoiceText} onParse={parseVoice} parsed={voiceParsed} onConfirm={confirmVoice} onCancel={()=>setVoiceParsed(null)} example="Reading value ten point zero zero two kilograms."/>
                )}
                <table className="n-table">
                  <thead><tr><th>#</th><th>Observed</th><th>Source</th><th>Timestamp</th></tr></thead>
                  <tbody>
                    {test.observations.map((o,i)=>(
                      <tr key={o.id}><td>{i+1}</td><td className="f-mono">{o.observed} {instrument.capacityUnit}</td><td>{o.source==="iot"?<><Radio size={11}/> IoT</>:"Manual"}</td><td className="f-mono">{fmtT(o.timestamp)}</td></tr>
                    ))}
                    {test.observations.length===0 && <tr><td colSpan={4} style={{textAlign:"center", color:"var(--ink-faint)", padding:16}}>No readings recorded yet.</td></tr>}
                  </tbody>
                </table>
                {test.repeatCalc && (
                  <div style={{marginTop:10, fontSize:12.5}}>
                    Range: <span className="f-mono">{test.repeatCalc.range}</span> {instrument.capacityUnit} · Allowed: <span className="f-mono">±{test.repeatCalc.allowed}</span> {instrument.capacityUnit}
                    <ExplainCalculation calc={test.repeatCalc} instrument={instrument} kind="repeatability"/>
                  </div>
                )}
              </>
            )}

            {test.testKey==="DIS" && <DiscriminationEntry test={test} instrument={instrument} mutate={mutate} logAudit={logAudit} currentUserName={currentUserName} role={role} notify={notify}/>}
            {test.testKey==="ECC" && <EccentricityEntry test={test} instrument={instrument} rule={rule} mutate={mutate} logAudit={logAudit} currentUserName={currentUserName} role={role} notify={notify}/>}
            {test.testKey==="SEN" && <SensitivityEntry test={test} instrument={instrument} mutate={mutate} logAudit={logAudit} currentUserName={currentUserName} role={role} notify={notify}/>}
            {test.testKey==="WARM" && <WarmUpEntry test={test} instrument={instrument} mutate={mutate} logAudit={logAudit} currentUserName={currentUserName} role={role} notify={notify}/>}
            {["ZERO","CREEP","TARE","TEMP","VOLT","SPAN"].includes(test.testKey) && <ToleranceTestEntry test={test} instrument={instrument} rule={rule} mutate={mutate} logAudit={logAudit} currentUserName={currentUserName} testKey={test.testKey} role={role} notify={notify}/>}
          </div>

          {/* COMPLIANCE EXPLANATION */}
          <div className="n-panel" style={{padding:18, marginBottom:16}}>
            <div className="n-eyebrow" style={{marginBottom:8}}>Compliance Result & Explanation</div>
            <div style={{display:"flex", alignItems:"center", gap:10, marginBottom:10}}>
              <ResultBadge result={test.result}/>
              <span style={{fontSize:12, color:"var(--ink-faint)"}}>evaluated under {rule.version}</span>
            </div>
            {(test.testKey==="ACC" && test.observations.length>0) && (()=>{
              const last = test.observations[test.observations.length-1];
              return (
                <div className="n-kv" style={{marginBottom:10, background:"var(--surface-2)", padding:10}}>
                  <dt>Observed Value</dt><dd>{last.observed} {instrument.capacityUnit}</dd>
                  <dt>Calculated Error</dt><dd>{last.error>=0?"+":""}{last.error} {instrument.capacityUnit}</dd>
                  <dt>Applicable Permissible Error</dt><dd>±{last.mpeVal} {instrument.capacityUnit}</dd>
                  <dt>Rule Version</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>{rule.id}</dd>
                </div>
              );
            })()}
            {(test.testKey==="REP" && test.repeatCalc) && (
              <div className="n-kv" style={{marginBottom:10, background:"var(--surface-2)", padding:10}}>
                <dt>Range (Max−Min)</dt><dd>{test.repeatCalc.range} {instrument.capacityUnit}</dd>
                <dt>Applicable Permissible Range</dt><dd>±{test.repeatCalc.allowed} {instrument.capacityUnit}</dd>
                <dt>Rule Version</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>{rule.id}</dd>
              </div>
            )}
            <div className="n-eyebrow" style={{marginBottom:5, fontSize:10}}>Why?</div>
            <div style={{fontSize:12.5, color:"var(--ink-dim)"}}>{explainResult(test, instrument)}</div>
          </div>

          {test.testKey==="ACC" && <CompareWithHistory db={db} instrument={instrument} test={test} nav={nav}/>}

          {/* EVIDENCE */}
          <div className="n-panel" style={{padding:18, marginBottom:16}}>
            <div className="n-eyebrow" style={{marginBottom:10}}>Evidence</div>
            <input ref={evidenceInputRef} type="file" style={{display:"none"}} aria-label="Upload evidence file" onChange={e=>{const f=e.target.files?.[0]; if(f) addEvidence(evidenceType,f); e.target.value="";}}/>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:12}}>
              <button className="n-btn n-btn-sm" onClick={()=>openEvidencePicker("photo-instrument")}>Instrument Photo</button>
              <button className="n-btn n-btn-sm" onClick={()=>openEvidencePicker("photo-setup")}>Setup Photo</button>
              <button className="n-btn n-btn-sm" onClick={()=>openEvidencePicker("document")}>Supporting Document</button>
            </div>
            {(!test.evidence || test.evidence.length===0) ? (
              <div style={{fontSize:12,color:"var(--ink-faint)"}}>No evidence attached yet. Upload photographs or supporting documents. Measurement observations and environmental readings remain linked automatically.</div>
            ) : (
              <div style={{display:"grid",gap:8}}>
                {test.evidence.map(ev=>(
                  <div key={ev.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,border:"1px solid var(--line)",padding:"9px 10px"}}>
                    <div style={{minWidth:0}}>
                      <div style={{fontSize:12.5,fontWeight:650,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{ev.originalName||ev.label}</div>
                      <div style={{fontSize:10.5,color:"var(--ink-faint)",marginTop:2}}>{ev.type === "document" ? "Supporting document" : ev.type === "photo-setup" ? "Test setup photograph" : "Instrument photograph"} · {ev.size ? `${Math.max(1,Math.round(ev.size/1024))} KB` : "legacy record"} · {fmtT(ev.addedAt)}</div>
                    </div>
                    <div style={{display:"flex",gap:6,flexShrink:0}}>
                      {ev.dataUrl && <a className="n-btn n-btn-sm" href={ev.dataUrl} target="_blank" rel="noreferrer">View</a>}
                      <button className="n-btn n-btn-sm" onClick={()=>removeEvidence(ev.id)}>Remove</button>
                    </div>
                  </div>
                ))}
              </div>
            )}
            <div style={{fontSize:11,color:"var(--ink-faint)",marginTop:10}}>Maximum 8 MB per file. Photos accept JPG, PNG, and WebP. Supporting documents accept PDF, Word, and text files.</div>
          </div>

          {/* REMARKS + WORKFLOW */}
          <div className="n-panel" style={{padding:18}}>
            <div className="n-eyebrow" style={{marginBottom:8}}>{canReview && test.status==="review" ? "Reviewer Decision" : "Remarks & Workflow"}</div>
            {canReview && test.status==="review" && (
              <div style={{fontSize:11.5, color:"var(--ink-dim)", marginBottom:10}}>Review both the official OIML result and the historical measurement behavior above, then record your decision. Example: "Official compliance achieved. Historical deviation increase reviewed; further monitoring recommended."</div>
            )}
            <textarea className="n-textarea" rows={3} value={remarksDraft} onChange={e=>setRemarksDraft(e.target.value)} placeholder="Add remarks / reviewer comment…"/>
            {canSubmit && test.status==="pending" && test.observations.length===0 && (
              <div style={{marginTop:10}}>
                {!naOpen ? (
                  <button className="n-btn n-btn-sm" onClick={()=>setNaOpen(true)}>Mark Not Applicable</button>
                ) : (
                  <div style={{padding:"8px 10px", border:"1px solid var(--line)", borderRadius:8}}>
                    <div style={{fontSize:11, fontWeight:700, color:"var(--ink-dim)", marginBottom:6}}>Mark "{test.testName}" Not Applicable — reason required</div>
                    <textarea className="n-textarea" rows={2} value={naReason} onChange={e=>setNaReason(e.target.value)} placeholder="e.g. Instrument is purely mechanical; voltage variation test does not apply."/>
                    <div style={{display:"flex", gap:8, marginTop:6}}>
                      <button className="n-btn n-btn-sm n-btn-primary" onClick={markNotApplicable}>Confirm Not Applicable</button>
                      <button className="n-btn n-btn-sm" onClick={()=>setNaOpen(false)}>Cancel</button>
                    </div>
                  </div>
                )}
              </div>
            )}
            {test.status==="not-applicable" && (
              <div style={{marginTop:10, padding:"8px 10px", border:"1px solid var(--line)", borderRadius:8, fontSize:12}}>
                <div style={{fontWeight:700, marginBottom:4}}>Marked Not Applicable</div>
                <div style={{color:"var(--ink-dim)", marginBottom:6}}>{test.naReason}</div>
                {canSubmit && <button className="n-btn n-btn-sm" onClick={revertNotApplicable}>Reinstate as Applicable</button>}
              </div>
            )}
            {canSubmit && (test.status==="in_progress"||test.status==="pending") && test.observations.length>0 && (
              <div style={{marginTop:12, padding:"8px 10px", border:"1px solid var(--line)", borderRadius:8}}>
                <div style={{fontSize:11, fontWeight:700, color:"var(--ink-dim)", marginBottom:4}}>Submission Validation</div>
                {submission.flags.map((f,i)=>(
                  <div key={i} style={{display:"flex", gap:6, alignItems:"flex-start", fontSize:11.5, marginBottom:2, color: f.level==="bad"?"var(--rose)":f.level==="warn"?"var(--amber)":"var(--seal)"}}>
                    <FlagIcon level={f.level}/><span>{f.msg}</span>
                  </div>
                ))}
              </div>
            )}
            <div style={{display:"flex", gap:8, marginTop:12, flexWrap:"wrap"}}>
              {canSubmit && (test.status==="in_progress"||test.status==="pending") && test.observations.length>0 &&
                <button className="n-btn n-btn-primary" onClick={submitReview} disabled={!submission.canSubmit} title={!submission.canSubmit?"Resolve blocking errors above before submitting.":""}>Submit for Review</button>}
              {canSubmit && (test.status==="in_progress"||test.status==="pending") &&
                <button className="n-btn" onClick={saveAndContinueLater}>Save & Continue Later</button>}
              {canReview && test.status==="review" && (
                <>
                  <button className="n-btn n-btn-seal" onClick={()=>reviewerAction("approve")}>Approve</button>
                  <button className="n-btn n-btn-rose" onClick={()=>reviewerAction("reject")}>Reject</button>
                  <button className="n-btn" onClick={()=>reviewerAction("correction")}>Return for Correction</button>
                </>
              )}
              {test.status==="approved" && <span style={{fontSize:12, color:"var(--seal)", fontWeight:600}}><CheckCircle2 size={14} style={{verticalAlign:"-2px"}}/> Approved by {test.approver}</span>}
              {test.status==="rejected" && <span style={{fontSize:12, color:"var(--rose)", fontWeight:600}}><XCircle size={14} style={{verticalAlign:"-2px"}}/> Rejected — technician should review and resubmit.</span>}
            </div>
          </div>
        </div>

        <div>
          <div className="n-panel" style={{padding:16, marginBottom:14}}>
            <div className="n-eyebrow" style={{marginBottom:8}}>Environmental Snapshot</div>
            {latestEnv ? (
              <>
                <div className="n-kv" style={{marginBottom:10}}>
                  <dt>Temperature</dt><dd>{latestEnv.temperature}°C</dd>
                  <dt>Humidity</dt><dd>{latestEnv.humidity}%RH</dd>
                  <dt>Sensor</dt><dd>{latestEnv.sensorId}</dd>
                  <dt>Captured</dt><dd>{fmtT(latestEnv.timestamp)}</dd>
                </div>
                {envForTest.length>1 && (
                  <>
                    <div className="n-hr" style={{margin:"8px 0"}}/>
                    <div className="n-label" style={{marginBottom:5}}>Recent History</div>
                    {envForTest.slice(0,4).map(e=>(
                      <div key={e.id} style={{display:"flex", justifyContent:"space-between", fontSize:11, padding:"3px 0", color: e.temperature>27.5?"var(--amber)":"var(--ink-dim)"}}>
                        <span className="f-mono">{fmtT(e.timestamp)}</span><span className="f-mono">{e.temperature}°C</span>
                      </div>
                    ))}
                  </>
                )}
              </>
            ) : <div style={{fontSize:12, color:"var(--ink-faint)"}}>No environmental readings linked yet. Use IoT Device capture above, or visit IoT Lab.</div>}
          </div>
          <div className="n-panel" style={{padding:16, marginBottom:14}}>
            <div className="n-eyebrow" style={{marginBottom:8}}>Reference Equipment</div>
            <div style={{fontSize:11.5, color:"var(--ink-faint)", marginBottom:8}}>{TEST_DEFS[test.testKey].requiredEquipment}</div>
            {canSubmit && (test.status==="in_progress"||test.status==="pending") && (
              <select className="n-select" style={{marginBottom:8, width:"100%"}} value={equipment?.id||""}
                onChange={e=>{ if(!can(role,"test:record")){ notify("Your role does not permit changing reference equipment."); return; } const val=e.target.value; mutate(d=>{ const dt=d.tests.find(x=>x.id===test.id); dt.referenceEquipmentId=val; logAudit(d,"Reference equipment linked","Test",dt.id,"—",val,"Selected for this test record"); }); }}>
                {equipmentList.map(eq=><option key={eq.id} value={eq.id}>{eq.equipmentId} — {eq.equipmentType}</option>)}
              </select>
            )}
            {equipment ? (
              <div className="n-kv" style={{marginBottom:4}}>
                <dt>Equipment ID</dt><dd className="f-mono">{equipment.equipmentId}</dd>
                <dt>Capacity</dt><dd className="f-mono">{equipment.capacity}</dd>
                <dt>Resolution</dt><dd className="f-mono">{equipment.resolution}</dd>
                <dt>Cal. Certificate</dt><dd className="f-mono">{equipment.certNumber}</dd>
                <dt>Cal. Valid Until</dt><dd className="f-mono">{equipment.validUntil}</dd>
              </div>
            ) : <div style={{fontSize:12, color:"var(--ink-faint)"}}>No reference equipment on record.</div>}
            {eqFlags.map((f,i)=>(
              <div key={i} style={{display:"flex", gap:5, alignItems:"flex-start", fontSize:11, marginTop:3, color: f.level==="bad"?"var(--rose)":f.level==="warn"?"var(--amber)":"var(--seal)"}}>
                <FlagIcon level={f.level}/><span>{f.msg}</span>
              </div>
            ))}
          </div>
          <div className="n-panel" style={{padding:16}}>
            <div className="n-eyebrow" style={{marginBottom:8}}>Review Chain</div>
            <div className="n-kv">
              <dt>Technician</dt><dd>{test.technician||"—"}</dd>
              <dt>Reviewer</dt><dd>{test.reviewer||"—"}</dd>
              <dt>Approver</dt><dd>{test.approver||"—"}</dd>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function explainResult(test, instrument){
  if(test.testKey==="ACC"){
    if(test.observations.length===0) return "No observations recorded yet.";
    const failing = test.observations.filter(o=>o.result==="FAIL");
    if(failing.length>0){
      const o = failing[0];
      return `At test load ${o.load} ${instrument.capacityUnit}, the observed indication error of ${o.error} ${instrument.capacityUnit} exceeds the permissible error of ±${o.mpeVal} ${instrument.capacityUnit} for this load band. Result: FAIL.`;
    }
    return `All recorded indication errors fall within the applicable maximum permissible error for their respective load bands. Result: PASS.`;
  }
  if(test.testKey==="REP"){
    if(!test.repeatCalc) return "Insufficient readings to compute repeatability.";
    return `The range between the highest (${test.repeatCalc.max}) and lowest (${test.repeatCalc.min}) of ${test.observations.length} readings is ${test.repeatCalc.range} ${instrument.capacityUnit}, against a permissible range of ±${test.repeatCalc.allowed} ${instrument.capacityUnit}. Result: ${test.repeatCalc.result}.`;
  }
  if(test.testKey==="DIS"){
    if(test.observations.length===0) return "Not yet executed.";
    const o = test.observations[0];
    return o.changed ? "The indication changed visibly upon addition of the 1e increment mass at high load, confirming adequate discrimination. Result: PASS." : "The indication did not change upon addition of the increment mass. Result: FAIL.";
  }
  if(test.testKey==="ECC"){
    if(test.observations.length===0) return "Not yet executed.";
    return test.remarks || "Awaiting evaluation.";
  }
  if(test.testKey==="SEN"){
    if(test.observations.length===0) return "Not yet executed.";
    const o = test.observations[0];
    return o.changed ? `The indication changed visibly upon addition of the 1e increment mass at ${o.load} ${instrument.capacityUnit}, confirming adequate sensitivity. Result: PASS.` : "The indication did not change upon addition of the increment mass. Result: FAIL.";
  }
  if(test.testKey==="ZERO"){
    if(test.observations.length===0) return "Not yet executed.";
    const o = test.observations[0];
    return `Zero-return error of ${o.error} ${instrument.capacityUnit} against a tolerance of ±${o.toleranceAbs} ${instrument.capacityUnit}. Result: ${test.result}.`;
  }
  if(["CREEP","TARE","TEMP","VOLT","SPAN"].includes(test.testKey)){
    if(test.observations.length===0) return "Not yet executed.";
    const o = test.observations[0];
    return `Observed error of ${o.error} ${instrument.capacityUnit} against the applicable MPE of ±${o.mpeVal} ${instrument.capacityUnit}. Result: ${test.result}.`;
  }
  if(test.testKey==="WARM"){
    return test.remarks || "Awaiting sufficient warm-up readings.";
  }
  return "";
}

function MeasurementSourceTabs({source, setSource}){
  const opts = [["manual","Manual"],["iot","IoT Device"],["voice","Voice"]];
  return (
    <div style={{display:"flex", gap:6, marginBottom:12}}>
      {opts.map(([k,l])=>(
        <button key={k} className="n-btn n-btn-sm" style={source===k?{background:"var(--navy)", color:"#fff", borderColor:"var(--navy)"}:{}} onClick={()=>setSource(k)}>
          {k==="iot" && <Radio size={12}/>}{k==="voice" && <Mic size={12}/>} {l}
        </button>
      ))}
    </div>
  );
}
function IotCapturePanel({scaleDevice, tempDevice, humDevice, onCapture, multiAction}){
  const connected = scaleDevice?.status==="connected";
  return (
    <div className="n-panel-2" style={{padding:12, marginBottom:12}}>
      <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8}}>
        <div style={{fontSize:12.5, fontWeight:600, display:"flex", alignItems:"center", gap:6}}>
          <Radio size={13}/> {scaleDevice?.name || "NAWI-TEST-001"} — {connected ? <span style={{color:"var(--seal)"}}>Connected </span> : <span style={{color:"var(--rose)"}}>Disconnected </span>}
        </div>
        <span className="n-badge n-badge-neutral">SIMULATION MODE</span>
      </div>
      <div style={{display:"flex", gap:20, marginBottom:10}}>
        <div><div className="n-label">Temperature</div><div className="f-mono" style={{fontSize:16}}>{tempDevice?.value ?? "—"}°C</div></div>
        <div><div className="n-label">Humidity</div><div className="f-mono" style={{fontSize:16}}>{humDevice?.value ?? "—"}%RH</div></div>
      </div>
      {!connected && <div style={{fontSize:11.5, color:"var(--rose)", marginBottom:8}}><AlertTriangle size={12} style={{verticalAlign:"-2px"}}/> Weighing device offline — automatic capture is disabled until reconnected in IoT Lab.</div>}
      <div style={{display:"flex", gap:8}}>
        <button className="n-btn n-btn-sm n-btn-primary" disabled={!connected} onClick={onCapture}>Capture Reading</button>
        {multiAction && <button className="n-btn n-btn-sm" disabled={!connected} onClick={multiAction.onClick}>{multiAction.label}</button>}
      </div>
    </div>
  );
}
function VoiceInput({voiceText, setVoiceText, onParse, parsed, onConfirm, onCancel, example}){
  return (
    <div className="n-panel-2" style={{padding:10, marginBottom:12}}>
      <div style={{display:"flex", gap:8, alignItems:"center"}}>
        <Mic size={14} color="var(--navy-dim)"/>
        <input className="n-input" style={{flex:1}} placeholder={`Simulate voice: "${example}"`} value={voiceText} onChange={e=>setVoiceText(e.target.value)}/>
        <button className="n-btn n-btn-sm" onClick={()=>setVoiceText(example)}>Use Example</button>
        <button className="n-btn n-btn-sm" disabled={!voiceText} onClick={onParse}>Transcribe</button>
      </div>
      {parsed && (
        <div style={{marginTop:8, fontSize:12, display:"flex", gap:10, alignItems:"center"}}>
          Parsed → Load: <strong className="f-mono">{parsed.load}</strong> Observed: <strong className="f-mono">{parsed.observed}</strong>
          <button className="n-btn n-btn-sm n-btn-seal" onClick={onConfirm}>Confirm & Save</button>
          <button className="n-btn n-btn-sm" onClick={onCancel}>Discard</button>
        </div>
      )}
    </div>
  );
}

function DiscriminationEntry({test, instrument, mutate, logAudit, currentUserName, role, notify}){
  const [before, setBefore] = useState(instrument.maxCapacity);
  const [changed, setChanged] = useState(null);
  const save = () => {
    if(!can(role,"test:record")){ notify?.("Your role does not permit recording measurements."); return; }
    if(isLockedForEntry(test)){ notify?.("This test has been submitted and its measurements are locked."); return; }
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      dt.observations = [{id:uid("OBS"), before:Number(before), after:Number(before)+instrument.e, incrementE:1, changed, timestamp:nowISO()}];
      dt.status = "in_progress"; dt.technician = currentUserName; dt.result = changed ? "PASS":"FAIL";
      logAudit(d,"Observation recorded","Test",dt.id,"—",`discrimination check @${before}${instrument.capacityUnit}: ${changed?"changed":"no change"}`,"Manual entry");
    });
  };
  return (
    <div>
      <div style={{display:"flex", gap:8, alignItems:"flex-end", marginBottom:10, flexWrap:"wrap"}}>
        <Field label={`Load Before Increment (${instrument.capacityUnit})`}><input className="n-input" style={{width:160}} type="number" value={before} onChange={e=>setBefore(e.target.value)}/></Field>
        <Field label="Indication Changed After Adding 1e?">
          <select className="n-select" style={{width:140}} value={changed===null?"":changed?"yes":"no"} onChange={e=>setChanged(e.target.value==="yes")}>
            <option value="">Select…</option><option value="yes">Yes</option><option value="no">No</option>
          </select>
        </Field>
        <button className="n-btn n-btn-sm" disabled={changed===null} onClick={save}>Record Result</button>
      </div>
      {test.observations.length>0 && (
        <div style={{fontSize:12.5}}>Before: <span className="f-mono">{test.observations[0].before}</span> → After: <span className="f-mono">{test.observations[0].after}</span> {instrument.capacityUnit} — <ResultBadge result={test.result}/></div>
      )}
    </div>
  );
}
function EccentricityEntry({test, instrument, rule, mutate, logAudit, currentUserName, role, notify}){
  const load = round(0.5*instrument.maxCapacity,3);
  const positions = ["Center","Corner 1","Corner 2","Corner 3","Corner 4"];
  const [vals, setVals] = useState(Object.fromEntries(positions.map(p=>[p,""])));
  const save = () => {
    if(!can(role,"test:record")){ notify?.("Your role does not permit recording measurements."); return; }
    if(isLockedForEntry(test)){ notify?.("This test has been submitted and its measurements are locked."); return; }
    const obs = positions.map(p=>({id:uid("OBS"), position:p, load, observed:Number(vals[p]), timestamp:nowISO()}));
    const errors = obs.map(o=>Math.abs(o.observed-load));
    const maxErr = Math.max(...errors);
    const mpeE = ruleMpeInE(rule, load/instrument.e, instrument.verificationStage||"initial", instrument.accuracyClass); const mpeVal = round(mpeE*instrument.e,4);
    const result = maxErr<=mpeVal ? "PASS":"FAIL";
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      dt.observations = obs; dt.status="in_progress"; dt.technician=currentUserName; dt.result=result;
      dt.remarks = `Maximum corner deviation ${round(maxErr,4)} ${instrument.capacityUnit} vs MPE ±${mpeVal} ${instrument.capacityUnit}.`;
      logAudit(d,"Observation recorded","Test",dt.id,"—",`5 eccentricity positions @ ${load}${instrument.capacityUnit}`,"Manual entry");
    });
  };
  return (
    <div>
      <div style={{fontSize:12, color:"var(--ink-dim)", marginBottom:10}}>Fixed test load: <strong className="f-mono">{load} {instrument.capacityUnit}</strong> applied at center and 4 corner positions.</div>
      <div style={{display:"grid", gridTemplateColumns:"repeat(5,1fr)", gap:8, marginBottom:10}}>
        {positions.map(p=>(
          <Field key={p} label={p}><input className="n-input" type="number" step="0.001" value={vals[p]} onChange={e=>setVals(v=>({...v,[p]:e.target.value}))}/></Field>
        ))}
      </div>
      <button className="n-btn n-btn-sm" disabled={positions.some(p=>!vals[p])} onClick={save}>Record Result</button>
      {test.observations.length>0 && <div style={{marginTop:10, fontSize:12.5}}>{test.remarks} <ResultBadge result={test.result}/></div>}
    </div>
  );
}

// Config for the six tests that share a "reference reading vs. later reading, compared against
// a tolerance" shape (item 8: Creep, Zero Return, Tare, Temperature Effect, Voltage Variation,
// Span Stability). CREEP/TARE/TEMP/VOLT/SPAN use the class-aware MPE at the relevant load as
// their tolerance (a documented simplification — the exact R76 clause for each may prescribe a
// different fraction; a metrology reviewer should confirm before certifying). ZERO uses the
// dedicated zero-return tolerance above.
const TOLERANCE_TEST_CONFIG = {
  CREEP: {loadPct:100, refLabel:"Indication immediately after applying Max load", obsLabel:"Indication after 30 minutes under Max load", note:"Hold Max load for 30 minutes between the two readings."},
  ZERO:  {loadPct:0,   refLabel:"Reference (zero)", obsLabel:"Indication after removing a full Max-capacity load", note:"Apply and fully remove Max load once before taking the return reading.", useZero:true},
  TARE:  {loadPct:50,  refLabel:"Certified net test load", obsLabel:"Net indication with tare set", note:"Zero the display with the tare container in place, then apply the certified net load."},
  TEMP:  {loadPct:100, refLabel:"Indication at reference temperature", obsLabel:"Indication after recorded temperature change", note:"Record ambient temperature at both readings; allow the instrument to stabilize."},
  VOLT:  {loadPct:100, refLabel:"Indication at nominal supply voltage", obsLabel:"Indication at rated voltage extreme", note:"Applicable to mains/battery-powered electronic instruments only."},
  SPAN:  {loadPct:100, refLabel:"Indication at first span check", obsLabel:"Indication at repeat span check", note:"Compare two span-calibration checks separated by the applicable interval."},
};
function ToleranceTestEntry({test, instrument, rule, mutate, logAudit, currentUserName, testKey, role, notify}){
  const cfg = TOLERANCE_TEST_CONFIG[testKey];
  const defaultLoad = round((cfg.loadPct/100)*instrument.maxCapacity,3);
  const [refVal, setRefVal] = useState(cfg.useZero ? 0 : defaultLoad);
  const [obsVal, setObsVal] = useState("");
  const save = () => {
    if(!can(role,"test:record")){ notify?.("Your role does not permit recording measurements."); return; }
    if(isLockedForEntry(test)){ notify?.("This test has been submitted and its measurements are locked."); return; }
    const calc = cfg.useZero ? evalZeroReturn(obsVal, instrument, instrument.verificationStage) : evalAccuracy(Number(refVal), Number(obsVal), instrument, rule);
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      dt.observations = [{id:uid("OBS"), load: cfg.useZero?0:Number(refVal), observed:Number(obsVal), unit:instrument.capacityUnit, source:"manual", timestamp:nowISO(), ...calc}];
      dt.status = "in_progress"; dt.technician = currentUserName; dt.result = calc.result;
      logAudit(d,"Observation recorded","Test",dt.id,"—",`${testKey} check — ${cfg.useZero?"zero-return":"reference "+refVal} → ${obsVal} ${instrument.capacityUnit} (${calc.result})`,"Manual entry");
    });
  };
  const o = test.observations?.[0];
  return (
    <div>
      <div style={{fontSize:11.5, color:"var(--ink-faint)", marginBottom:10}}>{cfg.note}</div>
      <div style={{display:"flex", gap:8, alignItems:"flex-end", marginBottom:10, flexWrap:"wrap"}}>
        {!cfg.useZero && <Field label={`${cfg.refLabel} (${instrument.capacityUnit})`}><input className="n-input" style={{width:180}} type="number" step="0.001" value={refVal} onChange={e=>setRefVal(e.target.value)}/></Field>}
        <Field label={`${cfg.obsLabel} (${instrument.capacityUnit})`}><input className="n-input" style={{width:180}} type="number" step="0.001" value={obsVal} onChange={e=>setObsVal(e.target.value)}/></Field>
        <button className="n-btn n-btn-sm" disabled={obsVal===""} onClick={save}>Record Result</button>
      </div>
      {o && (
        <div style={{fontSize:12.5}}>
          {cfg.useZero
            ? <>Zero-return error: <span className="f-mono">{o.error>=0?"+":""}{o.error}</span> {instrument.capacityUnit} (tolerance ±{o.toleranceAbs}) — <ResultBadge result={test.result}/></>
            : <>Error: <span className="f-mono">{o.error>=0?"+":""}{o.error}</span> {instrument.capacityUnit} (MPE ±{o.mpeVal}) — <ResultBadge result={test.result}/></>}
          <ExplainCalculation calc={o} instrument={instrument} kind="accuracy"/>
        </div>
      )}
    </div>
  );
}

function SensitivityEntry({test, instrument, mutate, logAudit, currentUserName, role, notify}){
  const [load, setLoad] = useState(round(0.5*instrument.maxCapacity,3));
  const [before, setBefore] = useState("");
  const [changed, setChanged] = useState(null);
  const save = () => {
    if(!can(role,"test:record")){ notify?.("Your role does not permit recording measurements."); return; }
    if(isLockedForEntry(test)){ notify?.("This test has been submitted and its measurements are locked."); return; }
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      dt.observations = [{id:uid("OBS"), load:Number(load), before:Number(before), after:Number(before)+instrument.e, incrementE:1, changed, timestamp:nowISO()}];
      dt.status = "in_progress"; dt.technician = currentUserName; dt.result = changed ? "PASS":"FAIL";
      logAudit(d,"Observation recorded","Test",dt.id,"—",`sensitivity check @${load}${instrument.capacityUnit}: ${changed?"changed":"no change"}`,"Manual entry");
    });
  };
  return (
    <div>
      <div style={{display:"flex", gap:8, alignItems:"flex-end", marginBottom:10, flexWrap:"wrap"}}>
        <Field label={`Test Load (${instrument.capacityUnit})`}><input className="n-input" style={{width:140}} type="number" step="0.001" value={load} onChange={e=>setLoad(e.target.value)}/></Field>
        <Field label={`Indication Before (${instrument.capacityUnit})`}><input className="n-input" style={{width:160}} type="number" step="0.001" value={before} onChange={e=>setBefore(e.target.value)}/></Field>
        <Field label="Indication Changed After Adding 1e?">
          <select className="n-select" style={{width:140}} value={changed===null?"":changed?"yes":"no"} onChange={e=>setChanged(e.target.value==="yes")}>
            <option value="">Select…</option><option value="yes">Yes</option><option value="no">No</option>
          </select>
        </Field>
        <button className="n-btn n-btn-sm" disabled={changed===null||before===""} onClick={save}>Record Result</button>
      </div>
      {test.observations.length>0 && (
        <div style={{fontSize:12.5}}>Load: <span className="f-mono">{test.observations[0].load}</span> · Before: <span className="f-mono">{test.observations[0].before}</span> → After: <span className="f-mono">{test.observations[0].after}</span> {instrument.capacityUnit} — <ResultBadge result={test.result}/></div>
      )}
    </div>
  );
}

function WarmUpEntry({test, instrument, mutate, logAudit, currentUserName, role, notify}){
  const WARM_UP_MINUTES = 15; // Typical default absent a manufacturer-stated warm-up time; editable per instrument in a future revision.
  const [reading, setReading] = useState("");
  const readings = test.observations || [];
  const add = () => {
    if(!can(role,"test:record")){ notify?.("Your role does not permit recording measurements."); return; }
    if(isLockedForEntry(test)){ notify?.("This test has been submitted and its measurements are locked."); return; }
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===test.id);
      dt.observations = [...(dt.observations||[]), {id:uid("OBS"), observed:Number(reading), timestamp:nowISO()}];
      dt.status = "in_progress"; dt.technician = currentUserName;
      const vals = dt.observations.map(o=>o.observed);
      if(vals.length>=3){
        const last3 = vals.slice(-3);
        const spread = round(Math.max(...last3)-Math.min(...last3),4);
        dt.result = spread <= instrument.e ? "PASS" : null;
        dt.remarks = `Spread of last 3 readings: ${spread} ${instrument.capacityUnit} vs e = ${instrument.e} ${instrument.capacityUnit}. ${dt.result==="PASS" ? "Indication stable — warm-up complete." : "Not yet stable — continue taking readings."}`;
      }
      logAudit(d,"Observation recorded","Test",dt.id,"—",`warm-up reading ${reading}${instrument.capacityUnit}`,"Manual entry");
    });
    setReading("");
  };
  return (
    <div>
      <div style={{fontSize:11.5, color:"var(--ink-faint)", marginBottom:10}}>Take successive no-load (zero) readings at intervals after power-on, for at least the stated warm-up time (default {WARM_UP_MINUTES} minutes absent a manufacturer figure). Stable when the last 3 readings span no more than one scale interval (e).</div>
      <div style={{display:"flex", gap:8, alignItems:"flex-end", marginBottom:10}}>
        <Field label={`Zero-Load Indication (${instrument.capacityUnit})`}><input className="n-input" style={{width:180}} type="number" step="0.001" value={reading} onChange={e=>setReading(e.target.value)}/></Field>
        <button className="n-btn n-btn-sm" disabled={reading===""} onClick={add}>Record Reading</button>
      </div>
      {readings.length>0 && (
        <table className="n-table">
          <thead><tr><th>#</th><th>Time</th><th>Indication</th></tr></thead>
          <tbody>{readings.map((r,i)=>(<tr key={r.id}><td>{i+1}</td><td className="f-mono">{fmtT(r.timestamp)}</td><td className="f-mono">{r.observed} {instrument.capacityUnit}</td></tr>))}</tbody>
        </table>
      )}
      {test.remarks && <div style={{marginTop:10, fontSize:12.5}}>{test.remarks} {test.result && <ResultBadge result={test.result}/>}</div>}
    </div>
  );
}

/* ============================== IOT LAB ============================== */
function IotLabPage({db, mutate, logAudit, notify}){
  const [tab, setTab] = useState("bench");
  const [tick, setTick] = useState(0);
  useEffect(()=>{ const iv = setInterval(()=>setTick(t=>t+1), 2500); return ()=>clearInterval(iv); },[]);

  const scale = db.iotDevices.find(d=>d.id==="DEV-SCALE");
  const temp = db.iotDevices.find(d=>d.id==="DEV-TEMP");
  const hum = db.iotDevices.find(d=>d.id==="DEV-HUM");
  const gw = db.iotDevices.find(d=>d.id==="DEV-GW");
  const activeTests = db.tests.filter(t=>t.status==="in_progress"||t.status==="pending");
  const [selTest, setSelTest] = useState(activeTests[0]?.id||"");

  useEffect(()=>{
    if(scale?.status!=="connected") return;
    mutate(d=>{
      const s=d.iotDevices.find(x=>x.id==="DEV-TEMP"); if(s) s.value = round(24.2+Math.random()*0.6,1);
      const h=d.iotDevices.find(x=>x.id==="DEV-HUM"); if(h) h.value = round(51+Math.random()*1.2,1);
    });
    // eslint-disable-next-line
  },[tick]);

  const toggleConnect = (id) => mutate(d=>{ const dev=d.iotDevices.find(x=>x.id===id); dev.status = dev.status==="connected"?"disconnected":"connected"; if(dev.status==="connected") dev.lastComm=nowISO(); });
  const pingDevice = (id) => { mutate(d=>{ const dev=d.iotDevices.find(x=>x.id===id); dev.lastComm=nowISO(); }); notify("Device connection tested — communication confirmed."); };
  const lastIotObs = useMemo(()=>{
    let latest=null;
    db.tests.forEach(t=>(t.observations||[]).forEach(o=>{ if(o.source==="iot" && o.observed!==undefined && (!latest || o.timestamp>latest.timestamp)) latest=o; }));
    return latest;
  },[db.tests]);

  const simulateReading = () => {
    const t = db.tests.find(x=>x.id===selTest);
    if(!t){ notify("Select a test to receive the captured reading."); return; }
    const inst = db.instruments.find(i=>i.id===t.instrumentId);
    mutate(d=>{
      const dt = d.tests.find(x=>x.id===t.id);
      if(t.testKey==="ACC"){
        const points = [0, inst.minCapacity, 0.25*inst.maxCapacity, 0.5*inst.maxCapacity, 0.75*inst.maxCapacity, inst.maxCapacity];
        const done = dt.observations.map(o=>o.load);
        const next = points.find(p=>!done.includes(round(p,3))) ?? points[points.length-1];
        const observed = round(next + (Math.random()-0.5)*inst.e*1.4, 4);
        const rule = RULE_VERSIONS.find(r=>r.id===dt.ruleVersionId)||activeRule();
        const calc = evalAccuracy(round(next,3), observed, inst, rule);
        dt.observations.push({id:uid("OBS"), load:round(next,3), observed, unit:inst.capacityUnit, source:"iot", timestamp:nowISO(), ...calc});
        dt.status="in_progress"; const results=dt.observations.map(o=>o.result); dt.result = results.includes("FAIL")?"FAIL":"PASS";
      } else if(t.testKey==="REP"){
        const load = 0.5*inst.maxCapacity;
        const observed = round(load + (Math.random()-0.5)*inst.e*1.6, 4);
        dt.observations.push({id:uid("OBS"), load, observed, unit:inst.capacityUnit, source:"iot", timestamp:nowISO()});
        const rule = RULE_VERSIONS.find(r=>r.id===dt.ruleVersionId)||activeRule();
        const readings = dt.observations.map(o=>o.observed);
        dt.repeatCalc = evalRepeatability(readings, load, inst, rule); dt.result = dt.repeatCalc?.result;
        dt.status="in_progress";
      }
      d.envReadings = [{id:uid("ENV"), instrumentId:inst.id, testId:t.id, timestamp:nowISO(), temperature:temp.value, humidity:hum.value, sensorId:"TEMP-01/HUM-01"}, ...d.envReadings];
      logAudit(d,"IoT reading captured","Test",t.id,"—","simulated device reading","Smart IoT Test Bench simulation");
    });
    notify("Simulated reading captured into the active test.");
  };

  const generateAnomaly = () => {
    const t = db.tests.find(x=>x.id===selTest);
    mutate(d=>{
      const tp = d.iotDevices.find(x=>x.id==="DEV-TEMP"); tp.value = round(tp.value+5.5,1);
      if(t){
        const dt = d.tests.find(x=>x.id===t.id); dt.envFlag = true;
        d.envReadings = [{id:uid("ENV"), instrumentId:dt.instrumentId, testId:t.id, timestamp:nowISO(), temperature:tp.value, humidity:hum.value, sensorId:"TEMP-01/HUM-01"}, ...d.envReadings];
        logAudit(d,"Environmental anomaly flagged","Test",dt.id,`${round(tp.value-5.5,1)}°C`,`${tp.value}°C`,"Sudden ambient variation detected by IoT gateway");
      }
    });
    notify("Anomaly injected — environmental variation flagged for the selected test.");
  };

  return (
    <div>
      <PageHeader eyebrow="Connected Devices & Sensors" title="Smart IoT Test Bench" desc="Simulated laboratory instrumentation for demonstration. Every captured reading flows into the instrument's test record and, in turn, its Measurement Fingerprint — IoT here is a reliable data source, not an isolated feature. Designed to integrate with certified laboratory equipment through an appropriate data interface."/>
      <div style={{display:"flex", gap:4, marginBottom:14, borderBottom:"1px solid var(--line)"}}>
        <div className={`n-tab ${tab==="bench"?"active":""}`} onClick={()=>setTab("bench")}>Device Bench</div>
        <div className={`n-tab ${tab==="env"?"active":""}`} onClick={()=>setTab("env")}>Environmental Log</div>
      </div>

      {tab==="bench" && (
        <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:16}}>
          <div className="n-panel" style={{padding:18}}>
            <div className="n-eyebrow" style={{marginBottom:10}}>Device Status</div>
            {[scale,temp,hum,gw].map(d=>(
              <div key={d.id} style={{padding:"9px 0", borderBottom:"1px solid var(--line)"}}>
                <div style={{display:"flex", justifyContent:"space-between", alignItems:"center"}}>
                  <div style={{fontSize:12.5}}>{d.name}</div>
                  <div style={{display:"flex", alignItems:"center", gap:10}}>
                    {d.value!==null && <span className="f-mono" style={{fontSize:12.5}}>{d.value} {d.unit}</span>}
                    <span className={`n-badge ${d.status==="connected"?"n-badge-pass":"n-badge-fail"}`}>{d.status==="connected"?"CONNECTED ":"OFFLINE "}</span>
                    <button className="n-btn n-btn-sm" onClick={()=>toggleConnect(d.id)}>{d.status==="connected"?"Disconnect":"Connect"}</button>
                  </div>
                </div>
                <div style={{display:"flex", justifyContent:"flex-end", alignItems:"center", gap:8, marginTop:4}}>
                  <span style={{fontSize:10.5, color:"var(--ink-faint)"}}>Last communication: {d.lastComm ? fmtT(d.lastComm) : "—"}</span>
                  <button className="n-btn n-btn-sm n-btn-ghost" style={{padding:"2px 7px", fontSize:11}} disabled={d.status!=="connected"} onClick={()=>pingDevice(d.id)}><Signal size={11}/> Test Connection</button>
                </div>
              </div>
            ))}
            <div style={{marginTop:12, fontSize:11, color:"var(--ink-faint)"}}>Simulation Mode — values are generated for demonstration, not sourced from certified hardware.</div>
          </div>

          <div className="n-panel" style={{padding:18}}>
            <div className="n-eyebrow" style={{marginBottom:10}}>Capture Control</div>
            <Field label="Send captured readings to test"><select className="n-select" value={selTest} onChange={e=>setSelTest(e.target.value)}>
              <option value="">Select an in-progress test…</option>
              {activeTests.map(t=>{const inst=db.instruments.find(i=>i.id===t.instrumentId); return <option key={t.id} value={t.id}>{inst.model} — {t.testName}</option>;})}
            </select></Field>
            <div style={{display:"flex", flexWrap:"wrap", gap:8, marginTop:14}}>
              <button className="n-btn" onClick={()=>toggleConnect("DEV-SCALE")}>{scale.status==="connected"?"Stop Test":"Start Test"}</button>
              <button className="n-btn n-btn-primary" disabled={!selTest || scale.status!=="connected"} onClick={simulateReading}><Radio size={13}/> Simulate Reading</button>
              <button className="n-btn n-btn-rose" onClick={()=>toggleConnect("DEV-SCALE")}>Disconnect Device</button>
              <button className="n-btn" disabled={!selTest} onClick={generateAnomaly}><AlertTriangle size={13}/> Generate Anomaly</button>
            </div>
            <div className="n-hr" style={{margin:"16px 0"}}/>
            <div className="n-eyebrow" style={{marginBottom:6}}>Live Readout</div>
            <div style={{display:"flex", gap:22}}>
              <div><div className="n-label">Weight (Last Capture)</div><div className="f-mono" style={{fontSize:20}}>{lastIotObs ? `${lastIotObs.observed} ${lastIotObs.unit||"kg"}` : "—"}</div></div>
              <div><div className="n-label">Temperature</div><div className="f-mono" style={{fontSize:20}}>{temp.value}°C</div></div>
              <div><div className="n-label">Humidity</div><div className="f-mono" style={{fontSize:20}}>{hum.value}%RH</div></div>
            </div>
          </div>
        </div>
      )}

      {tab==="env" && (
        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Environmental Readings</div>
          <table className="n-table">
            <thead><tr><th>Timestamp</th><th>Instrument</th><th>Test</th><th>Temperature</th><th>Humidity</th><th>Sensor</th></tr></thead>
            <tbody>
              {db.envReadings.map(e=>{
                const inst = db.instruments.find(i=>i.id===e.instrumentId);
                const test = db.tests.find(t=>t.id===e.testId);
                const anomalous = e.temperature>27.5;
                return (
                  <tr key={e.id}>
                    <td className="f-mono">{fmtDT(e.timestamp)}</td>
                    <td>{inst?.model}</td>
                    <td>{test?.testName}</td>
                    <td className="f-mono" style={anomalous?{color:"var(--amber)", fontWeight:700}:{}}>{e.temperature}°C {anomalous && <AlertTriangle size={11} style={{verticalAlign:"-2px"}}/>}</td>
                    <td className="f-mono">{e.humidity}%RH</td>
                    <td>{e.sensorId}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ============================== REPORTS ============================== */
function reportEligibility(db, inst){
  const allTests = db.tests.filter(t=>t.instrumentId===inst.id);
  const tests = allTests.filter(t=>t.status!=="not-applicable");
  let reason = "", eligible = false;
  if(allTests.length===0){ reason = "No test plan generated yet"; }
  else {
    const approved = allTests.filter(t=>t.status==="approved").length;
    const na = allTests.filter(t=>t.status==="not-applicable").length;
    const rejected = allTests.filter(t=>t.status==="rejected").length;
    if(approved+na === allTests.length && tests.length>0){ eligible = true; }
    else { reason = `${approved}/${allTests.length} tests approved${na?`, ${na} not applicable`:""}${rejected?`, ${rejected} rejected`:""} — complete review first`; }
  }
  return {tests, eligible, reason};
}

function ReportsPage({db, mutate, logAudit, nav, notify, role, selectedInstrument, selectedReport}){
  const [viewInstrumentId, setViewInstrumentId] = useState(selectedInstrument || null);
  const [viewReportId, setViewReportId] = useState(selectedReport || null);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [justGenerated, setJustGenerated] = useState(null);
  const [q, setQ] = useState("");

  useEffect(()=>{ if(selectedInstrument) setViewInstrumentId(selectedInstrument); },[selectedInstrument]);
  useEffect(()=>{ if(selectedReport) setViewReportId(selectedReport); },[selectedReport]);

  if(viewReportId){
    const r = db.reports.find(x=>x.id===viewReportId);
    const inst = db.instruments.find(i=>i.id===r.instrumentId);
    return <ReportDocument db={db} report={r} nav={nav} justGenerated={justGenerated===r.id} onBack={()=>{setViewReportId(null); setJustGenerated(null); setViewInstrumentId(inst.id);}}
      breadcrumb={[{label:"Reports", onClick:()=>{setViewReportId(null); setViewInstrumentId(null); setJustGenerated(null);}}, {label:inst.model, onClick:()=>{setViewReportId(null); setViewInstrumentId(inst.id); setJustGenerated(null);}}, {label:r.reportNumber}]}/>;
  }

  if(viewInstrumentId){
    const inst = db.instruments.find(i=>i.id===viewInstrumentId);
    const myReports = db.reports.filter(r=>r.instrumentId===inst.id);
    const elig = reportEligibility(db, inst);

    const generate = () => {
      if(!can(role,"report:approve-final")){ notify("Your role does not permit issuing the final report — this requires Authority or Admin approval."); return; }
      const num = `NAWI-2026-${String(1000+db.reports.length+87).slice(-5)}`;
      const ruleForReport = activeRule();
      const reportSnapshot = {
        instrument: JSON.parse(JSON.stringify(inst)),
        tests: JSON.parse(JSON.stringify(elig.tests)),
        envReadings: JSON.parse(JSON.stringify(db.envReadings.filter(e=>e.instrumentId===inst.id))),
        referenceEquipment: JSON.parse(JSON.stringify((db.referenceEquipment||[]).filter(eq=>elig.tests.some(t=>t.referenceEquipmentId===eq.id)))),
        rule: JSON.parse(JSON.stringify(ruleForReport)),
        capturedAt: nowISO()
      };
      const rpt = {id:uid("RPT"), reportNumber:num, instrumentId:inst.id, testIds:elig.tests.map(t=>t.id), generatedAt:nowISO(),
        status:"Issued", ruleVersionId:ruleForReport.id, integrityHash:reportContentFingerprint(inst, elig.tests, ruleForReport.id),
        technician:elig.tests[0]?.technician, reviewer:elig.tests.find(t=>t.reviewer)?.reviewer, approver:elig.tests.find(t=>t.approver)?.approver,
        snapshot: reportSnapshot};
      mutate(d=>{
        d.reports.push(rpt);
        const i = d.instruments.find(x=>x.id===inst.id); i.status="Evaluation Completed";
        logAudit(d,"Report generated","Report",rpt.id,"—",rpt.reportNumber,"All applicable tests approved");
      });
      notify(`Report ${num} generated successfully.`);
      setPreviewOpen(false);
      setJustGenerated(rpt.id);
      setViewReportId(rpt.id);
    };

    return (
      <div>
        <PageHeader breadcrumb={[{label:"Reports", onClick:()=>setViewInstrumentId(null)}, {label:inst.model}]}
          eyebrow="Report History" title={inst.model} desc={`${inst.manufacturer} · ${inst.serialNumber} · ${myReports.length} report${myReports.length!==1?"s":""} on file`}
          right={<button className="n-btn n-btn-sm n-btn-ghost" onClick={()=>setViewInstrumentId(null)}><ChevronLeft size={13}/> All Instruments</button>}/>

        <div className="n-panel" style={{padding:18, marginBottom:16}}>
          <div style={{display:"flex", justifyContent:"space-between", alignItems:"center"}}>
            <div>
              <div className="n-eyebrow" style={{marginBottom:6}}>Report Eligibility</div>
              <div style={{fontSize:12.5, color:"var(--ink-dim)"}}>
                {elig.eligible ? "All applicable tests are approved — ready to generate a new report." : (elig.reason || "No applicable tests yet.")}
              </div>
            </div>
            <div style={{display:"flex", gap:8}}>
              {!elig.eligible && <button className="n-btn n-btn-sm" onClick={()=>nav("instruments",{instrumentId:inst.id})}>Open Passport</button>}
              <button className="n-btn n-btn-primary" disabled={!elig.eligible || !can(role,"report:approve-final")} title={!can(role,"report:approve-final")?"Requires Authority or Admin role":""} onClick={()=>setPreviewOpen(true)}>Generate New Report</button>
            </div>
          </div>
        </div>

        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Reports for This Instrument</div>
          {myReports.length===0 && <div style={{fontSize:12.5, color:"var(--ink-faint)"}}>No reports issued yet for this instrument.</div>}
          <table className="n-table">
            <thead><tr><th>Report No.</th><th>Test Date</th><th>Status</th><th>Result</th><th>Rule Version</th><th>Reviewer</th><th>Actions</th></tr></thead>
            <tbody>
              {myReports.map(r=>{
                const rtests = db.tests.filter(t=>r.testIds.includes(t.id));
                const overallResult = rtests.some(t=>t.result==="FAIL") ? "FAIL" : rtests.some(t=>t.result==="REVIEW") ? "REVIEW" : "PASS";
                return (
                  <tr key={r.id}>
                    <td className="f-mono">{r.reportNumber}</td>
                    <td className="f-mono">{fmtDT(r.generatedAt)}</td>
                    <td><span className="n-badge n-badge-pass"><ShieldCheck size={11}/> Issued</span></td>
                    <td><ResultBadge result={overallResult}/></td>
                    <td style={{fontSize:11.5, color:"var(--ink-dim)"}}>{r.ruleVersionId}</td>
                    <td>{r.reviewer||"—"}</td>
                    <td style={{display:"flex", gap:6, flexWrap:"wrap"}}>
                      <button className="n-btn n-btn-sm" onClick={()=>setViewReportId(r.id)}>View Report</button>
                      <button className="n-btn n-btn-sm" onClick={()=>nav("instruments",{instrumentId:inst.id})}>View Instrument</button>
                      <button className="n-btn n-btn-sm" onClick={()=>nav("audit")}>Audit Trail</button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {previewOpen && (
          <Modal title="Report Preview" onClose={()=>setPreviewOpen(false)} width={520}>
            <div style={{marginBottom:14}}>
              {[
                ["Instrument Details", true], ["Laboratory Conditions", db.envReadings.some(e=>e.instrumentId===inst.id)],
                ["Test Results", elig.tests.length>0], ["Calculations", elig.tests.every(t=>t.observations?.length>0)],
                ["Evidence", true], ["Reviewer Sign-off", elig.tests.every(t=>t.reviewer)],
                ["Rule Version", true], ["Audit Information", true],
              ].map(([label,ok])=>(
                <div key={label} style={{display:"flex", alignItems:"center", gap:7, fontSize:12.5, padding:"5px 0"}}>
                  {ok ? <CheckCircle2 size={14} color="var(--seal)"/> : <AlertTriangle size={14} color="var(--amber)"/>}
                  {label}
                </div>
              ))}
            </div>
            <div style={{fontSize:11.5, color:"var(--ink-faint)", marginBottom:14}}>A unique report number and integrity fingerprint will be generated. This action is recorded in the audit trail.</div>
            <div style={{display:"flex", justifyContent:"flex-end", gap:8}}>
              <button className="n-btn" onClick={()=>setPreviewOpen(false)}>Cancel</button>
              <button className="n-btn n-btn-primary" onClick={generate}>Generate Report</button>
            </div>
          </Modal>
        )}
      </div>
    );
  }

  const rows = db.instruments.map(inst=>({inst, reports: db.reports.filter(r=>r.instrumentId===inst.id), elig: reportEligibility(db, inst)}))
    .filter(r=> !q || `${r.inst.model} ${r.inst.manufacturer} ${r.inst.serialNumber}`.toLowerCase().includes(q.toLowerCase()));

  return (
    <div>
      <PageHeader eyebrow="Standardized Output" title="Reports" desc="Select an instrument to view its full report history, or generate a new report once every applicable test is approved."/>
      <div style={{position:"relative", maxWidth:320, marginBottom:14}}>
        <Search size={13} style={{position:"absolute", left:9, top:10, color:"var(--ink-faint)"}}/>
        <input className="n-input" style={{paddingLeft:28}} placeholder="Search manufacturer, model, serial…" value={q} onChange={e=>setQ(e.target.value)}/>
      </div>
      <div style={{display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(260px,1fr))", gap:14}}>
        {rows.map(({inst, reports, elig})=>(
          <div key={inst.id} className="n-panel n-row-hover" style={{padding:16, cursor:"pointer"}} onClick={()=>setViewInstrumentId(inst.id)}>
            <div style={{fontWeight:700, fontSize:14}}>{inst.model}</div>
            <div style={{fontSize:12, color:"var(--ink-dim)"}}>{inst.manufacturer}</div>
            <div className="f-mono" style={{fontSize:11, color:"var(--ink-faint)", marginTop:2}}>{inst.serialNumber}</div>
            <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginTop:12}}>
              <div>
                <div className="f-display" style={{fontSize:20, fontWeight:700}}>{reports.length}</div>
                <div style={{fontSize:10.5, color:"var(--ink-faint)"}}>Report{reports.length!==1?"s":""}</div>
              </div>
              {elig.eligible ? <span className="n-badge n-badge-pass">Ready</span> : reports.length>0 ? <span className="n-badge n-badge-neutral">On File</span> : <span className="n-badge n-badge-review">In Progress</span>}
            </div>
            <button type="button" className="n-btn n-btn-sm" style={{width:"100%", marginTop:12, pointerEvents:"none"}} tabIndex={-1}>View Reports <ChevronRight size={13}/></button>
          </div>
        ))}
        {rows.length===0 && <div style={{gridColumn:"1 / -1", padding:24, textAlign:"center", color:"var(--ink-faint)", fontSize:12.5}}>No instruments match your search.</div>}
      </div>
    </div>
  );
}

function ReportFingerprintSection({db, instrument}){
  const fp = computeFingerprint(db, instrument);
  const meta = FP_STATUS_META[fp.status];
  return (
    <div style={{border:`1px solid ${meta.border}`, background:meta.bg, padding:14, marginBottom:18}}>
      <div className="n-eyebrow" style={{marginBottom:4, color:meta.color}}>Historical Measurement Observation</div>
      <div style={{fontSize:10.5, color:"var(--ink-faint)", marginBottom:10, fontStyle:"italic"}}>Historical trend analysis — decision-support information. Does not affect the official compliance results above.</div>
      <div className="n-kv">
        <dt>Historical comparison available</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>{fp.status==="INSUFFICIENT" ? "No" : "Yes"}</dd>
        <dt>Current trend indicator</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif', color:meta.color, fontWeight:700}}>{meta.icon} {meta.label}</dd>
        {fp.previous && <><dt>Previous evaluation reference</dt><dd>{fp.previous.error>=0?"+":""}{fp.previous.error} {fp.previous.unit} ({fmtDT(fp.previous.date)})</dd></>}
        {fp.status!=="STABLE" && fp.status!=="INSUFFICIENT" && <><dt>Review flag</dt><dd style={{color:meta.color}}>Recommended — review instrument condition and test setup</dd></>}
      </div>
    </div>
  );
}

function ReportDocument({db, report, onBack, breadcrumb, nav, justGenerated}){
  // Issued reports are immutable snapshots. Older seeded reports fall back to the live
  // repository for backwards compatibility.
  const snapshot = report.snapshot || null;
  const inst = snapshot?.instrument || db.instruments.find(i=>i.id===report.instrumentId);
  const tests = snapshot?.tests || db.tests.filter(t=>report.testIds.includes(t.id));
  const naTests = snapshot ? [] : db.tests.filter(t=>t.instrumentId===inst.id && t.status==="not-applicable");
  const env = (snapshot?.envReadings || db.envReadings.filter(e=>e.instrumentId===inst.id)).slice(0,6);
  const rule = snapshot?.rule || RULE_VERSIONS.find(r=>r.id===report.ruleVersionId) || activeRule();
  const equipmentUsed = snapshot?.referenceEquipment || Array.from(new Set(tests.map(t=>t.referenceEquipmentId).filter(Boolean)))
    .map(id=>(db.referenceEquipment||[]).find(eq=>eq.id===id)).filter(Boolean);
  const overallResult = tests.some(t=>t.result==="FAIL") ? "NON-COMPLIANT" : tests.some(t=>t.result==="REVIEW") ? "REVIEW REQUIRED" : "COMPLIANT";
  const allEvidence = tests.flatMap(t=>(t.evidence||[]).map(ev=>({...ev, testName:t.testName})));
  // Per-observation rows for one test: load, observed indication, error, applicable MPE and
  // pass/fail per point — the level of detail an OIML R-76 test report is expected to carry,
  // rather than only the one-line PASS/FAIL summary from explainResult().
  const observationRows = (t) => {
    const unit = inst.capacityUnit || "";
    if(["ZERO","CREEP","TARE","TEMP","VOLT","SPAN"].includes(t.testKey)){
      const o=t.observations?.[0]; if(!o) return [];
      return [[o.load!==undefined&&o.load!==null?`${o.load} ${unit}`:"—", `${o.observed} ${unit}`, o.error!==undefined?`${o.error>=0?"+":""}${o.error} ${unit}`:"—", o.mpeVal!==undefined?`±${o.mpeVal} ${unit}`:"—", o.result||"—"]];
    }
    if(["DIS","SEN"].includes(t.testKey)){
      const o=t.observations?.[0]; if(!o) return [];
      return [[o.load!==undefined&&o.load!==null?`${o.load} ${unit}`:"—", `${o.before} ${unit}`, o.changed?"Changed":"No change", "—", o.changed?"PASS":"FAIL"]];
    }
    if(t.testKey==="ECC"){
      return (t.observations||[]).map(o=>[o.position||"—", `${o.load} ${unit}`, `${o.observed} ${unit}`, o.mpeVal!==undefined?`±${o.mpeVal} ${unit}`:"—", o.result||"—"]);
    }
    if(t.testKey==="WARM"){
      return (t.observations||[]).map((o,i)=>[`Reading ${i+1}`, `${o.observed} ${unit}`, "—", "—", "—"]);
    }
    // ACC / REP and anything else with load+observed+error+mpeVal+result per point
    return (t.observations||[]).map(o=>[o.load!==undefined?`${o.load} ${unit}`:"—", `${o.observed} ${unit}`, o.error!==undefined?`${o.error>=0?"+":""}${o.error} ${unit}`:"—", o.mpeVal!==undefined?`±${o.mpeVal} ${unit}`:"—", o.result||"—"]);
  };
  const observationHeaders = (t) => t.testKey==="ECC" ? ["Position","Load","Observed","MPE","Result"]
    : ["DIS","SEN"].includes(t.testKey) ? ["Load","Before increment","Change","MPE","Result"]
    : t.testKey==="WARM" ? ["Reading","Observed","—","—","—"]
    : ["Load","Observed","Error","MPE","Result"];

  const escapeHtml = (value) => String(value ?? "—").replace(/[&<>"']/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  const downloadBlob = (blob, filename) => { const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download=filename; a.click(); setTimeout(()=>URL.revokeObjectURL(url),1000); };

  const printReport = () => {
    const previousTitle = document.title;
    document.title = `${report.reportNumber} — NAWI Test Report`;
    window.setTimeout(() => {
      window.print();
      window.setTimeout(() => { document.title = previousTitle; }, 500);
    }, 50);
  };

  const downloadPdf = async () => {
    try {
    const pdf = new jsPDF({unit:"mm", format:"a4"});
    // Same METRA QR token as on screen so the issued PDF stays independently verifiable.
    try { pdf.addImage(await qrPngDataUrl(makeReportQrToken(report.id)),"PNG",165,12,30,30); } catch (_) {}
    const margin=15, pageW=210, contentW=180;
    let y=18;
    const ensure=(h=8)=>{ if(y+h>282){pdf.addPage(); y=18;} };
    const text=(value,size=9,bold=false)=>{ensure(size*0.5+2); pdf.setFont("helvetica",bold?"bold":"normal"); pdf.setFontSize(size); const lines=pdf.splitTextToSize(String(value??"—"),contentW); pdf.text(lines,margin,y); y += lines.length*(size*0.42)+2;};
    const heading=(value)=>{ensure(10); pdf.setFont("helvetica","bold"); pdf.setFontSize(11); pdf.text(value,margin,y); y+=6;};
    const row=(label,value)=>{ensure(7); pdf.setFont("helvetica","bold"); pdf.setFontSize(8); pdf.text(`${label}:`,margin,y); pdf.setFont("helvetica","normal"); pdf.text(pdf.splitTextToSize(String(value??"—"),contentW-35),margin+30,y); y+=5;};
    pdf.setProperties({title:report.reportNumber,subject:"NAWI OIML R-76 Type-Evaluation Report",author:"Navi"});
    pdf.setFont("helvetica","bold"); pdf.setFontSize(16); pdf.text(inst?.laboratory||"NAVI — NAWI LABORATORY",margin,y); y+=7;
    text("NAWI Type-Evaluation Test Report — per OIML R-76",10); text(`Report No.: ${report.reportNumber}   Issued: ${fmtDT(report.generatedAt)}`,9); text(`Overall Result: ${overallResult}`,11,true); y+=2;
    heading("Instrument Details");
    row("Manufacturer",inst.manufacturer); row("Applicant",inst.applicant); row("Model",inst.model); row("Serial No.",inst.serialNumber); row("NAWI ID",inst.nawiId||inst.id); row("Accuracy Class",inst.accuracyClass); row("Max / Min",`${inst.maxCapacity} / ${inst.minCapacity} ${inst.capacityUnit}`); row("e / n",`${inst.e} ${inst.capacityUnit} / ${inst.nVerification}`); row("Verification Stage",inst.verificationStage||"initial");
    heading("Laboratory Conditions");
    if(env.length) env.forEach(e=>text(`${fmtT(e.timestamp)} — ${e.temperature}°C, ${e.humidity}%RH`,8)); else text("No linked environmental readings.",8);
    heading("Rule Version Applied"); text(rule?.version||report.ruleVersionId,9);
    heading("Reference Equipment");
    if(equipmentUsed.length) equipmentUsed.forEach(eq=>text(`${eq.equipmentId} — ${eq.equipmentType} — ${eq.certNumber} — valid until ${eq.validUntil}`,8)); else text("No reference equipment on record.",8);
    heading("Test Results");
    const drawObsTable = (headers, rows) => {
      if(!rows.length) return;
      const colW = contentW/headers.length;
      ensure(6); pdf.setFont("helvetica","bold"); pdf.setFontSize(7.5);
      headers.forEach((h,i)=>pdf.text(String(h),margin+i*colW,y)); y+=4;
      pdf.setFont("helvetica","normal");
      rows.forEach(r=>{ ensure(5); r.forEach((v,i)=>pdf.text(pdf.splitTextToSize(String(v),colW-2)[0]||"",margin+i*colW,y)); y+=4.2; });
      y+=2;
    };
    tests.forEach(t=>{
      ensure(14); text(t.testName,9,true);
      text(`Result: ${t.result||"—"} | ${explainResult(t,inst)} | Remarks: ${t.remarks||"—"}`,8);
      drawObsTable(observationHeaders(t), observationRows(t));
    });
    if(naTests.length){heading("Not Applicable Tests"); naTests.forEach(t=>text(`${t.testName} — ${t.naReason||"No reason recorded"}`,8));}
    heading("Sign-off"); row("Technician",report.technician); row("Reviewer",report.reviewer); row("Approving Authority",report.approver);
    heading("Traceability"); text(`${tests.reduce((n,t)=>n+(t.observations?.length||0),0)} measurement observations · ${env.length} environmental readings · ${allEvidence.length} evidence items`,8); row("Integrity fingerprint",report.integrityHash);
    text("This is a prototype/simulation output and does not constitute an OIML-certified test report.",8,false);
    downloadBlob(pdf.output("blob"),`${report.reportNumber}.pdf`);
    } catch(err) {
      window.alert(`PDF generation failed: ${err?.message || err}`);
    }
  };

  const downloadWord = async () => {
    try {
    const cell=(value,bold=false)=>new TableCell({children:[new Paragraph({children:[new TextRun({text:String(value??"—"),bold})]})]});
    const tableRows=(headers, rows)=>new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[new TableRow({children:headers.map(h=>cell(h,true))}),...rows.map(r=>new TableRow({children:r.map(v=>cell(v))}))]});
    const children=[
      new Paragraph({text:inst?.laboratory||"NAVI — NAWI LABORATORY",heading:HeadingLevel.TITLE}),
      new Paragraph({children:[new TextRun({text:"NAWI Type-Evaluation Test Report — per OIML R-76",bold:true}),new TextRun({text:`\nReport No.: ${report.reportNumber}\nIssued: ${fmtDT(report.generatedAt)}\nOverall Result: ${overallResult}`})]}),
      new Paragraph({text:"Instrument Details",heading:HeadingLevel.HEADING_1}),
      tableRows(["Field","Value"],[["Manufacturer",inst.manufacturer],["Applicant",inst.applicant],["Model",inst.model],["Serial No.",inst.serialNumber],["NAWI ID",inst.nawiId||inst.id],["Accuracy Class",inst.accuracyClass],["Max / Min",`${inst.maxCapacity} / ${inst.minCapacity} ${inst.capacityUnit}`],["e / n",`${inst.e} ${inst.capacityUnit} / ${inst.nVerification}`],["Verification Stage",inst.verificationStage||"initial"]]),
      new Paragraph({text:"Laboratory Conditions",heading:HeadingLevel.HEADING_1}),
      tableRows(["Timestamp","Temperature","Humidity"],env.length?env.map(e=>[fmtT(e.timestamp),`${e.temperature} °C`,`${e.humidity} %RH`]):[["—","—","—"]]),
      new Paragraph({text:"Rule Version Applied",heading:HeadingLevel.HEADING_1}), new Paragraph({text:rule?.version||report.ruleVersionId}),
      new Paragraph({text:"Reference Equipment",heading:HeadingLevel.HEADING_1}),
      tableRows(["Equipment ID","Type","Certificate","Valid Until"],equipmentUsed.length?equipmentUsed.map(eq=>[eq.equipmentId,eq.equipmentType,eq.certNumber,eq.validUntil]):[["—","—","—","—"]]),
      new Paragraph({text:"Test Results",heading:HeadingLevel.HEADING_1}),
      tableRows(["Test","Calculation","Result","Remarks"],tests.map(t=>[t.testName,explainResult(t,inst),t.result||"—",t.remarks||"—"])),
      new Paragraph({text:"Observation Detail",heading:HeadingLevel.HEADING_1}),
      ...tests.flatMap(t=>{
        const rows=observationRows(t);
        if(!rows.length) return [new Paragraph({children:[new TextRun({text:t.testName,bold:true})]}), new Paragraph({text:"No observations recorded."})];
        return [new Paragraph({children:[new TextRun({text:t.testName,bold:true})]}), tableRows(observationHeaders(t), rows)];
      }),
      new Paragraph({text:"Sign-off",heading:HeadingLevel.HEADING_1}),
      tableRows(["Technician","Reviewer","Approving Authority"],[[report.technician||"—",report.reviewer||"—",report.approver||"—"]]),
      new Paragraph({text:"Evidence & Traceability",heading:HeadingLevel.HEADING_1}),
      new Paragraph({text:`${tests.reduce((n,t)=>n+(t.observations?.length||0),0)} measurement observations · ${env.length} environmental readings · ${allEvidence.length} evidence items`}),
      new Paragraph({children:[new TextRun({text:"Integrity fingerprint: ",bold:true}),new TextRun({text:report.integrityHash||"—"})]}),
      new Paragraph({text:"This is a prototype/simulation output and does not constitute an OIML-certified test report."})
    ];
    const doc=new Document({sections:[{properties:{},children}]});
    downloadBlob(await Packer.toBlob(doc),`${report.reportNumber}.docx`);
    } catch(err) {
      window.alert(`Word report generation failed: ${err?.message || err}`);
    }
  };

  return (
    <div>
      {breadcrumb && <div className="no-print"><Breadcrumb items={breadcrumb}/></div>}
      {justGenerated && (
        <div className="no-print n-panel" style={{padding:"12px 16px", marginBottom:14, borderColor:"var(--seal-line)", background:"var(--seal-bg)", display:"flex", alignItems:"center", gap:10}}>
          <CheckCircle2 size={18} color="var(--seal)"/>
          <div style={{fontSize:12.5, color:"var(--seal)"}}><strong>Report Generated Successfully ✓</strong> — Report ID: <span className="f-mono">{report.reportNumber}</span></div>
        </div>
      )}
      <div className="no-print" style={{display:"flex", justifyContent:"space-between", marginBottom:12}}>
        <button className="n-btn n-btn-sm n-btn-ghost" onClick={onBack}><ChevronLeft size={14}/> Back</button>
        <div style={{display:"flex", gap:8}}>
          {nav && <button className="n-btn n-btn-sm" onClick={()=>nav("instruments",{instrumentId:inst.id})}>Return to Instrument</button>}
          <button className="n-btn n-btn-sm" onClick={downloadWord}><Download size={13}/> Download Word (.docx, editable)</button>
          <button className="n-btn n-btn-sm" onClick={downloadPdf}><Download size={13}/> Download PDF</button>
          <button className="n-btn n-btn-sm n-btn-primary" onClick={printReport}><Printer size={13}/> Print</button>
        </div>
      </div>

      <div className="n-panel print-area" style={{padding:28}}>
        <div style={{display:"flex", justifyContent:"space-between", borderBottom:"2px solid var(--navy)", paddingBottom:14, marginBottom:16}}>
          <div>
            <div className="f-display" style={{fontWeight:700, fontSize:17}}>{inst.laboratory}</div>
            <div style={{fontSize:11.5, color:"var(--ink-dim)"}}>NAWI Type-Evaluation Test Report — per OIML R-76</div>
            <div className="f-mono" style={{fontSize:12, marginTop:6}}>Report No.: {report.reportNumber} · Date: {fmtDT(report.generatedAt)}</div>
            {inst.nawiId && <div className="f-mono" style={{fontSize:11, color:"var(--ink-faint)", marginTop:2}}>Instrument ID: {inst.nawiId}</div>}
          </div>
          <div style={{display:"flex", alignItems:"center", gap:12}}>
            <div style={{textAlign:"center"}}>
              <div style={{border:"1px solid var(--line)", padding:4, background:"#fff"}}><QrCode value={makeReportQrToken(report.id)} size={78}/></div>
              <div style={{fontSize:8.5, color:"var(--ink-faint)", marginTop:3}}>Scan to verify</div>
            </div>
            <Seal status={overallResult} size={70}/>
          </div>
        </div>

        <div className="f-display" style={{fontSize:13, fontWeight:700, marginBottom:10, letterSpacing:".03em"}}>⚖️ OFFICIAL OIML R-76 EVALUATION</div>
        <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:20, marginBottom:18}}>
          <div>
            <div className="n-eyebrow" style={{marginBottom:6}}>Instrument Details</div>
            <div className="n-kv">
              <dt>Manufacturer</dt><dd>{inst.manufacturer}</dd><dt>Applicant</dt><dd>{inst.applicant}</dd>
              <dt>Model</dt><dd>{inst.model}</dd><dt>Serial No.</dt><dd>{inst.serialNumber}</dd>
              <dt>Accuracy Class</dt><dd>{inst.accuracyClass}</dd>
              <dt>Max / Min</dt><dd>{inst.maxCapacity} / {inst.minCapacity} {inst.capacityUnit}</dd>
              <dt>e / n</dt><dd>{inst.e} {inst.capacityUnit} / {inst.nVerification}</dd>
            </div>
          </div>
          <div>
            <div className="n-eyebrow" style={{marginBottom:6}}>Laboratory Conditions</div>
            {env.length ? env.map(e=>(<div key={e.id} className="n-kv" style={{marginBottom:4}}><dt>{fmtT(e.timestamp)}</dt><dd>{e.temperature}°C, {e.humidity}%RH</dd></div>)) : <div style={{fontSize:12, color:"var(--ink-faint)"}}>No linked readings.</div>}
            <div className="n-eyebrow" style={{margin:"10px 0 6px"}}>Rule Version Applied</div>
            <div style={{fontSize:12}}>{rule.version}</div>
          </div>
        </div>

        <div className="n-eyebrow" style={{marginBottom:6}}>Reference Equipment</div>
        <table className="n-table" style={{marginBottom:18}}>
          <thead><tr><th>Equipment ID</th><th>Type</th><th>Cal. Certificate</th><th>Valid Until</th></tr></thead>
          <tbody>
            {equipmentUsed.length===0 && <tr><td colSpan={4} style={{color:"var(--ink-faint)"}}>No reference equipment on record.</td></tr>}
            {equipmentUsed.map(eq=>(<tr key={eq.id}><td>{eq.equipmentId}</td><td>{eq.equipmentType}</td><td>{eq.certNumber}</td><td>{eq.validUntil}</td></tr>))}
          </tbody>
        </table>

        <div className="n-eyebrow" style={{marginBottom:6}}>Test Results</div>
        <table className="n-table" style={{marginBottom: naTests.length?4:18}}>
          <thead><tr><th>Test</th><th>Key Calculation</th><th>Result</th><th>Remarks</th></tr></thead>
          <tbody>
            {tests.map(t=>(
              <tr key={t.id}>
                <td>{t.testName}</td>
                <td style={{color:"var(--ink-dim)"}}>{explainResult(t, inst)}</td>
                <td><ResultBadge result={t.result}/></td>
                <td style={{color:"var(--ink-dim)"}}>{t.remarks||"—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {naTests.length>0 && (
          <table className="n-table" style={{marginBottom:18}}>
            <tbody>
              {naTests.map(t=>(<tr key={t.id}><td>{t.testName}</td><td colSpan={3} style={{color:"var(--ink-faint)"}}>Not Applicable — {t.naReason||"No reason recorded"}</td></tr>))}
            </tbody>
          </table>
        )}

        <ReportFingerprintSection db={db} instrument={inst}/>

        <div style={{display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:16, marginBottom:18}}>
          <div><div className="n-eyebrow" style={{marginBottom:4}}>Technician</div><div style={{fontSize:12.5}}>{report.technician||"—"}</div></div>
          <div><div className="n-eyebrow" style={{marginBottom:4}}>Reviewer</div><div style={{fontSize:12.5}}>{report.reviewer||"—"}</div></div>
          <div><div className="n-eyebrow" style={{marginBottom:4}}>Approving Authority</div><div style={{fontSize:12.5}}>{report.approver||"—"}</div></div>
        </div>

        <div className="n-eyebrow" style={{marginBottom:6}}>Evidence & Traceability</div>
        <div className="print-section" style={{display:"flex", flexWrap:"wrap", gap:7, marginBottom:12}}>
          {allEvidence.length===0 && <div style={{fontSize:11.5, color:"var(--ink-faint)"}}>No supplementary evidence attached.</div>}
          {allEvidence.map(ev=>(<span key={ev.id} className="n-badge n-badge-neutral">{ev.type === "photo-setup" ? "Setup photograph" : ev.type === "photo-instrument" ? "Instrument photograph" : "Supporting document"}: {ev.originalName||ev.label}</span>))}
          <span className="n-badge n-badge-neutral">{tests.reduce((n,t)=>n+(t.observations?.length||0),0)} measurement observations</span>
          <span className="n-badge n-badge-neutral">{env.length} environmental readings</span>
        </div>
        {allEvidence.length>0 && (
          <div className="print-section" style={{display:"grid", gridTemplateColumns:"repeat(2,minmax(0,1fr))", gap:14, marginBottom:18}}>
            {allEvidence.map(ev=> ev.dataUrl && ev.mimeType?.startsWith("image/") ? (
              <figure key={`img-${ev.id}`} style={{margin:0, border:"1px solid var(--line)", padding:10}}>
                <img className="print-evidence-image" src={ev.dataUrl} alt={ev.originalName||ev.label} style={{display:"block", width:"100%", maxHeight:220, objectFit:"contain"}}/>
                <figcaption style={{fontSize:10.5, color:"var(--ink-dim)", marginTop:6}}>{ev.type === "photo-setup" ? "Test Setup Photograph" : "Instrument Photograph"} · {ev.originalName||ev.label} · {ev.testName}</figcaption>
              </figure>
            ) : (
              <div key={`doc-${ev.id}`} style={{border:"1px solid var(--line)", padding:10, fontSize:11.5}}>
                <strong>Supporting document</strong><br/>{ev.originalName||ev.label}<br/><span style={{color:"var(--ink-faint)"}}>{ev.mimeType||"file"} · {ev.testName}</span>
              </div>
            ))}
          </div>
        )}

        <div className="n-hr" style={{margin:"14px 0"}}/>
        <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", fontSize:11}}>
          <div style={{color:"var(--ink-faint)"}}>This is a prototype/simulation output and does not constitute an OIML-certified test report.</div>
          <div className="f-mono" style={{display:"flex", alignItems:"center", gap:6}}><Fingerprint size={13}/> {report.integrityHash}</div>
        </div>
      </div>
    </div>
  );
}

/* ============================== QR VERIFICATION ============================== */
function VerifyQrPage({db, nav, notify}){
  const [token,setToken]=useState("");
  const [scanning,setScanning]=useState(false);
  const videoRef=useRef(null);
  const streamRef=useRef(null);
  const [scannerMessage,setScannerMessage]=useState("");

  const stopScanner=()=>{
    if(streamRef.current){ streamRef.current.getTracks().forEach(t=>t.stop()); streamRef.current=null; }
    setScanning(false);
  };
  useEffect(()=>()=>stopScanner(),[]);
  useEffect(()=>{
    if(scanning && videoRef.current && streamRef.current){
      videoRef.current.srcObject=streamRef.current;
      videoRef.current.play().catch(()=>{});
    }
  },[scanning]);

  const verify=async(raw)=>{
    const parsed=parseMetraQrToken(raw);
    if(!parsed){ setScannerMessage("Invalid METRA QR. Use a QR generated by a METRA report."); return; }
    try{
      /* Server-first verification: a browser copy is never treated as proof. */
      const remote=await naviApi(`/api/public/reports/${encodeURIComponent(parsed.reportId)}`);
      stopScanner(); setScannerMessage("");
      const status = remote.verification?.status || "VERIFIED";
      notify?.(`Report ${remote.report.reportNumber} ${status.toLowerCase()} by METRA public verification.`);
      nav("reports",{reportId:remote.report.id,instrumentId:remote.report.instrumentId});
      return;
    }catch(err){
      const local=db.reports.find(r=>String(r.id).toUpperCase()===parsed.reportId);
      if(local){
        setScannerMessage("The QR is readable, but the server verification endpoint is unavailable. The local copy is shown only as an offline reference — it is not a verified result.");
      }else{
        setScannerMessage(`QR is readable, but the referenced report could not be verified: ${err.message}`);
      }
    }
  };

  const startScanner=async()=>{
    setScannerMessage("");
    if(!("BarcodeDetector" in window)){
      setScannerMessage("Camera QR scanning is not supported by this browser. Paste the QR text below instead.");
      return;
    }
    try{
      const formats=await window.BarcodeDetector.getSupportedFormats();
      if(!formats.includes("qr_code")){ setScannerMessage("This browser does not expose QR detection. Paste the QR text below instead."); return; }
      const stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:"environment"}}});
      streamRef.current=stream;
      setScanning(true);
      requestAnimationFrame(async function loop(){
        if(!streamRef.current || !videoRef.current) return;
        try{
          const detector=new window.BarcodeDetector({formats:["qr_code"]});
          const codes=await detector.detect(videoRef.current);
          if(codes.length && codes[0].rawValue){ verify(codes[0].rawValue); return; }
        }catch{}
        if(streamRef.current) requestAnimationFrame(loop);
      });
    }catch(err){
      stopScanner();
      setScannerMessage(err?.name==="NotAllowedError" ? "Camera permission was denied. Allow camera access or paste the QR text below." : "Could not start the camera. Paste the QR text below instead.");
    }
  };

  return (
    <div>
      <PageHeader eyebrow="Report Verification" title="Verify a METRA QR" desc="Scan a report QR or enter its non-sensitive verification token. METRA resolves it against the authenticated server repository or the public server verification endpoint."/>
      <div style={{display:"grid",gridTemplateColumns:"minmax(280px,1fr) minmax(280px,1fr)",gap:14}}>
        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:8}}>Camera Scanner</div>
          <div style={{fontSize:12,color:"var(--ink-dim)",marginBottom:12}}>Camera access is used only in this browser to read the QR; the QR itself contains only a report identifier.</div>
          {scanning ? <video ref={videoRef} playsInline muted style={{width:"100%",maxHeight:330,background:"#111",objectFit:"cover"}}/> : <div style={{border:"1px dashed var(--line-strong)",padding:38,textAlign:"center",color:"var(--ink-faint)",fontSize:12}}>Camera preview appears here when scanning starts.</div>}
          <div style={{display:"flex",gap:8,marginTop:12}}>
            {!scanning ? <button className="n-btn n-btn-primary" onClick={startScanner}><Camera size={13}/> Start Camera Scan</button> : <button className="n-btn" onClick={stopScanner}>Stop Scanner</button>}
          </div>
        </div>
        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:8}}>Manual Verification</div>
          <div style={{fontSize:12,color:"var(--ink-dim)",marginBottom:10}}>If your phone/browser scanner shows the QR text, paste it here.</div>
          <input className="n-input f-mono" style={{width:"100%",marginBottom:8}} value={token} onChange={e=>setToken(e.target.value)} placeholder="METRA:R:RPT-1"/>
          <button className="n-btn n-btn-primary" disabled={!token.trim()} onClick={()=>verify(token)}>Verify Report</button>
          {scannerMessage && <div style={{marginTop:12,padding:10,border:"1px solid var(--amber-line)",background:"var(--amber-bg)",fontSize:12,color:"var(--amber)"}}>{scannerMessage}</div>}
          <div className="n-hr" style={{margin:"18px 0"}}/>
          <div className="n-eyebrow" style={{marginBottom:6}}>Verification Rules</div>
          <ul style={{margin:"0 0 0 18px",padding:0,fontSize:12,lineHeight:1.7,color:"var(--ink-dim)"}}>
            <li>QR payload contains no instrument measurements or personal data.</li>
            <li>The report is resolved against Navi's authenticated repository or public verification endpoint.</li>
            <li>Successful verification opens the stored report and its integrity fingerprint.</li>
          </ul>
        </div>
      </div>
    </div>
  );
}

/* ============================== REPOSITORY ============================== */
function RepositoryPage({db, nav}){
  const [q, setQ] = useState("");
  const [status, setStatus] = useState("all");
  const rows = db.instruments.map(inst=>{
    const tests = db.tests.filter(t=>t.instrumentId===inst.id);
    const report = db.reports.find(r=>r.instrumentId===inst.id);
    const derivedStatus = report ? "Approved" : tests.some(t=>t.status==="review") ? "Pending Review" : tests.some(t=>t.result==="FAIL") ? "Failed" : tests.length ? "In Progress" : "Completed";
    return {inst, tests, report, derivedStatus};
  }).filter(r=>{
    if(status!=="all" && r.derivedStatus!==status) return false;
    const s = `${r.inst.manufacturer} ${r.inst.model} ${r.inst.serialNumber} ${r.inst.nawiId||""} ${r.report?.reportNumber||""}`.toLowerCase();
    return !q || s.includes(q.toLowerCase());
  });
  return (
    <div>
      <PageHeader eyebrow="Searchable Storage" title="Instrument & Report Repository" desc="Search across NAWI IDs, manufacturers, models, serial numbers, report numbers and evaluation status."/>
      <div style={{display:"flex", gap:8, marginBottom:14}}>
        <div style={{position:"relative", flex:1, maxWidth:320}}>
          <Search size={13} style={{position:"absolute", left:9, top:10, color:"var(--ink-faint)"}}/>
          <input className="n-input" style={{paddingLeft:28}} placeholder="NAWI ID, manufacturer, model, serial, report no…" value={q} onChange={e=>setQ(e.target.value)}/>
        </div>
        <select className="n-select" style={{width:200}} value={status} onChange={e=>setStatus(e.target.value)}>
          {["all","Completed","In Progress","Pending Review","Failed","Approved"].map(s=><option key={s} value={s}>{s==="all"?"All Statuses":s}</option>)}
        </select>
      </div>
      <div className="n-panel">
        <table className="n-table">
          <thead><tr><th>NAWI ID</th><th>Manufacturer / Model</th><th>Serial</th><th>Class</th><th>Report No.</th><th>Status</th><th>Laboratory</th><th></th></tr></thead>
          <tbody>
            {rows.map(r=>(
              <tr key={r.inst.id} className="n-row-hover" onClick={()=>nav("instruments",{instrumentId:r.inst.id})}>
                <td className="f-mono">{r.inst.nawiId||"—"}</td>
                <td><strong>{r.inst.manufacturer}</strong><br/><span style={{color:"var(--ink-faint)"}}>{r.inst.model}</span></td>
                <td className="f-mono">{r.inst.serialNumber}</td>
                <td>{r.inst.accuracyClass}</td>
                <td className="f-mono">{r.report?.reportNumber||"—"}</td>
                <td><span className="n-badge n-badge-neutral">{r.derivedStatus}</span></td>
                <td style={{color:"var(--ink-dim)"}}>{r.inst.laboratory}</td>
                <td><ChevronRight size={14} color="var(--ink-faint)"/></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== ANALYTICS ============================== */
/* ============================== MEASUREMENT INTELLIGENCE (primary nav page) ============================== */
function MeasurementIntelligencePage({db, mutate, logAudit, nav, notify, role}){
  const [filter, setFilter] = useState("all");
  const rows = db.instruments.map(inst=>({inst, fp:computeFingerprint(db, inst)}));
  const counts = {
    CHANGE: rows.filter(r=>r.fp.status==="CHANGE").length,
    SIGNIFICANT: rows.filter(r=>r.fp.status==="SIGNIFICANT").length,
    STABLE: rows.filter(r=>r.fp.status==="STABLE").length,
    INSUFFICIENT: rows.filter(r=>r.fp.status==="INSUFFICIENT").length,
  };
  const attention = rows.filter(r=>r.fp.status==="CHANGE"||r.fp.status==="SIGNIFICANT")
    .sort((a,b)=> (b.fp.status==="SIGNIFICANT"?1:0) - (a.fp.status==="SIGNIFICANT"?1:0));
  const filtered = filter==="all" ? rows : rows.filter(r=>r.fp.status===filter);

  const [cfgOpen, setCfgOpen] = useState(false);
  const [cfgDraft, setCfgDraft] = useState(db.fingerprintConfig);
  const saveCfg = () => {
    if(!can(role,"fingerprint:configure")){ notify("Your role does not permit changing this configuration."); return; }
    mutate(d=>{
      const prev = {...d.fingerprintConfig};
      d.fingerprintConfig = {minEvaluations:Number(cfgDraft.minEvaluations), changeThresholdE:Number(cfgDraft.changeThresholdE), significantThresholdE:Number(cfgDraft.significantThresholdE)};
      logAudit(d,"Historical Analysis Configuration updated","FingerprintConfig","GLOBAL",JSON.stringify(prev),JSON.stringify(d.fingerprintConfig),"Administrator adjusted trend-detection thresholds");
    });
    notify("Historical Analysis Configuration updated.");
    setCfgOpen(false);
  };

  return (
    <div>
      <PageHeader eyebrow="Core Product Intelligence" title="Measurement Intelligence" desc="Every NAWI develops a historical measurement behavior. This view surfaces instruments whose current evaluation differs from their own history — a decision-support layer, separate from official OIML compliance."
        right={role==="Admin" && <button className="n-btn n-btn-sm" onClick={()=>{setCfgDraft(db.fingerprintConfig); setCfgOpen(true);}}><Settings size={13}/> Historical Analysis Configuration</button>}/>

      <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:20}}>
        <MetricBlock label="Pattern Change" value={counts.CHANGE} onClick={()=>setFilter("CHANGE")} hint="Instruments requiring review"/>
        <MetricBlock label="Significant Change" value={counts.SIGNIFICANT} onClick={()=>setFilter("SIGNIFICANT")} hint="Stronger deviation pattern"/>
        <MetricBlock label="Stable Instruments" value={counts.STABLE} onClick={()=>setFilter("STABLE")}/>
        <MetricBlock label="Building Baseline" value={counts.INSUFFICIENT} onClick={()=>setFilter("INSUFFICIENT")} hint="Insufficient historical data"/>
      </div>

      {attention.length>0 && (
        <div className="n-panel" style={{padding:18, marginBottom:20, borderLeft:"4px solid var(--amber)"}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Instruments Requiring Attention</div>
          {attention.map(({inst,fp})=>{
            const meta = FP_STATUS_META[fp.status];
            return (
              <div key={inst.id} className="n-row-hover" style={{display:"flex", justifyContent:"space-between", alignItems:"center", padding:"10px 6px", borderBottom:"1px solid var(--line)"}} onClick={()=>nav("instruments",{instrumentId:inst.id})}>
                <div>
                  <div style={{fontWeight:700, fontSize:13}}>{inst.model} <span style={{color:"var(--ink-faint)", fontWeight:400}}>({inst.nawiId})</span></div>
                  <div style={{fontSize:11.5, color:"var(--ink-dim)"}}>Current: {fp.current.error>=0?"+":""}{fp.current.error} {fp.current.unit} · Previous: {fp.previous.error>=0?"+":""}{fp.previous.error} {fp.previous.unit}</div>
                </div>
                <div style={{display:"flex", alignItems:"center", gap:10}}>
                  <span className="n-badge" style={{background:meta.bg, color:meta.color, borderColor:meta.border}}>{meta.icon} {meta.label}</span>
                  <button className="n-btn n-btn-sm">View Fingerprint</button>
                </div>
              </div>
            );
          })}
        </div>
      )}

      <div className="n-panel" style={{marginBottom:20}}>
        <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", padding:"14px 18px 0"}}>
          <div className="n-eyebrow">All Instruments {filter!=="all" && `— filtered: ${FP_STATUS_META[filter].label}`}</div>
          {filter!=="all" && <button className="n-btn n-btn-sm n-btn-ghost" onClick={()=>setFilter("all")}>Clear filter</button>}
        </div>
        <table className="n-table">
          <thead><tr><th>Instrument</th><th>Fingerprint Status</th><th>Current</th><th>Previous</th><th>Evaluations</th><th></th></tr></thead>
          <tbody>
            {filtered.map(({inst,fp})=>{
              const meta = FP_STATUS_META[fp.status];
              return (
                <tr key={inst.id} className="n-row-hover" onClick={()=>nav("instruments",{instrumentId:inst.id})}>
                  <td><strong>{inst.model}</strong><br/><span style={{color:"var(--ink-faint)"}}>{inst.nawiId}</span></td>
                  <td><span className="n-badge" style={{background:meta.bg, color:meta.color, borderColor:meta.border}}>{meta.icon} {meta.label}</span></td>
                  <td className="f-mono">{fp.current ? `${fp.current.error>=0?"+":""}${fp.current.error} ${fp.current.unit}` : "—"}</td>
                  <td className="f-mono" style={{color:"var(--ink-dim)"}}>{fp.previous ? `${fp.previous.error>=0?"+":""}${fp.previous.error} ${fp.previous.unit}` : "—"}</td>
                  <td>{fp.points.length}</td>
                  <td><ChevronRight size={14} color="var(--ink-faint)"/></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="n-panel" style={{padding:18}}>
        <div className="n-eyebrow" style={{marginBottom:12}}>How the Measurement Fingerprint Works</div>
        <div style={{display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(200px,1fr))", gap:14, fontSize:12}}>
          {[
            "Test observations are recorded (manual or IoT).",
            "The source — IoT or manual — is preserved with each reading.",
            "The OIML R-76 rule engine determines official PASS/FAIL/REVIEW.",
            "Evaluation data is stored in the instrument's history.",
            "Current behavior is compared with previous evaluations across the load range.",
            "The system identifies Stable / Pattern Change / Significant Change.",
            "A human reviewer interprets the decision-support information — the system never decides alone.",
          ].map((step,i)=>(
            <div key={i} style={{display:"flex", gap:8}}>
              <span className="f-mono" style={{color:"var(--navy-dim)", fontWeight:700}}>{i+1}</span>
              <span style={{color:"var(--ink-dim)"}}>{step}</span>
            </div>
          ))}
        </div>
      </div>

      {cfgOpen && (
        <Modal title="Historical Analysis Configuration" onClose={()=>setCfgOpen(false)} width={440}>
          <div style={{fontSize:11.5, color:"var(--ink-faint)", marginBottom:14}}>These thresholds control the Measurement Fingerprint trend indicator only. They are laboratory-configurable and are not official OIML R-76 requirements, which remain governed separately under OIML Rules.</div>
          <Field label="Minimum evaluations required for baseline"><input className="n-input" type="number" min="2" value={cfgDraft.minEvaluations} onChange={e=>setCfgDraft(c=>({...c,minEvaluations:e.target.value}))}/></Field>
          <div style={{height:10}}/>
          <Field label="Pattern change threshold (× verification interval e)"><input className="n-input" type="number" step="0.5" value={cfgDraft.changeThresholdE} onChange={e=>setCfgDraft(c=>({...c,changeThresholdE:e.target.value}))}/></Field>
          <div style={{height:10}}/>
          <Field label="Significant change threshold (× verification interval e)"><input className="n-input" type="number" step="0.5" value={cfgDraft.significantThresholdE} onChange={e=>setCfgDraft(c=>({...c,significantThresholdE:e.target.value}))}/></Field>
          <div style={{display:"flex", justifyContent:"flex-end", gap:8, marginTop:16}}>
            <button className="n-btn" onClick={()=>setCfgOpen(false)}>Cancel</button>
            <button className="n-btn n-btn-primary" onClick={saveCfg}>Save Configuration</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== ANALYTICS ============================== */
function AnalyticsPage({db, stats, nav}){
  const byFailCategory = {};
  db.tests.filter(t=>t.result==="FAIL").forEach(t=>{ byFailCategory[t.testName]=(byFailCategory[t.testName]||0)+1; });
  const avgTests = db.instruments.length ? round(db.tests.length/db.instruments.length,1) : 0;

  const fpRows = db.instruments.map(inst=>({inst, fp:computeFingerprint(db, inst)}));
  const fpCounts = {
    CHANGE: fpRows.filter(r=>r.fp.status==="CHANGE").length,
    SIGNIFICANT: fpRows.filter(r=>r.fp.status==="SIGNIFICANT").length,
    STABLE: fpRows.filter(r=>r.fp.status==="STABLE").length,
    INSUFFICIENT: fpRows.filter(r=>r.fp.status==="INSUFFICIENT").length,
  };
  const loadRangeCounts = {};
  fpRows.forEach(({fp})=>{
    if(fp.loadDiffs && fp.loadDiffs.length>0){
      const worst = fp.loadDiffs.reduce((a,b)=>Math.abs(b.diff)>Math.abs(a.diff)?b:a);
      const band = worst.pct>=50 ? "50–100% load" : "0–50% load";
      loadRangeCounts[band] = (loadRangeCounts[band]||0)+1;
    }
  });

  return (
    <div>
      <PageHeader eyebrow="Laboratory-Wide Statistics" title="Laboratory Measurement Intelligence" desc="Aggregate view across all instruments. Fingerprint trend metrics first; conventional test throughput statistics are secondary detail below."/>

      <div className="n-eyebrow" style={{marginBottom:10}}>Measurement Behavior Across the Fleet</div>
      <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:20}}>
        <MetricBlock label="Instruments with Pattern Changes" value={fpCounts.CHANGE} onClick={()=>nav("intelligence")}/>
        <MetricBlock label="Significant Historical Changes" value={fpCounts.SIGNIFICANT} onClick={()=>nav("intelligence")}/>
        <MetricBlock label="Stable Instruments" value={fpCounts.STABLE} onClick={()=>nav("intelligence")}/>
        <MetricBlock label="Building Baseline" value={fpCounts.INSUFFICIENT} onClick={()=>nav("intelligence")}/>
      </div>

      <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:20}}>
        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Load Range Where Changes Occur Most Often</div>
          {Object.keys(loadRangeCounts).length===0 && <div style={{fontSize:12, color:"var(--ink-faint)"}}>No pattern changes recorded across the fleet yet.</div>}
          {Object.entries(loadRangeCounts).map(([band,v])=>(
            <div key={band} style={{marginBottom:10}}>
              <div style={{display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:3}}><span>{band}</span><span className="f-mono">{v} instrument{v!==1?"s":""}</span></div>
              <div style={{height:8, background:"var(--surface-2)"}}><div style={{height:8, width:`${Math.max(v*30,10)}px`, background:"var(--amber)"}}/></div>
            </div>
          ))}
        </div>
        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Instruments Requiring Historical Review</div>
          {fpRows.filter(r=>r.fp.status==="CHANGE"||r.fp.status==="SIGNIFICANT").length===0 && <div style={{fontSize:12, color:"var(--ink-faint)"}}>None currently flagged.</div>}
          {fpRows.filter(r=>r.fp.status==="CHANGE"||r.fp.status==="SIGNIFICANT").map(({inst,fp})=>(
            <div key={inst.id} className="n-row-hover" style={{display:"flex", justifyContent:"space-between", fontSize:12.5, padding:"6px 4px", borderBottom:"1px solid var(--line)"}} onClick={()=>nav("instruments",{instrumentId:inst.id})}>
              <span>{inst.model}</span><ResultBadge result={fp.status==="SIGNIFICANT"?"FAIL":"REVIEW"}/>
            </div>
          ))}
        </div>
      </div>

      <div className="n-eyebrow" style={{marginBottom:10}}>Conventional Test Statistics (Secondary)</div>
      <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:16}}>
        <MetricBlock label="Total Evaluations" value={db.instruments.length} onClick={()=>nav("instruments")}/>
        <MetricBlock label="Avg. Tests / Instrument" value={avgTests}/>
        <MetricBlock label="Pending Tests" value={db.tests.filter(t=>t.status==="pending"||t.status==="in_progress").length} onClick={()=>nav("tests",{testsFilter:"pending"})}/>
        <MetricBlock label="Data-Quality Flags" value={db.tests.filter(t=>t.anomalyMsg||t.envFlag).length} onClick={()=>nav("iot")}/>
      </div>
      <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:16}}>
        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Pass / Fail / Review Distribution</div>
          {[["Pass",stats.pass,"var(--seal)"],["Fail",stats.fail,"var(--rose)"],["Review",stats.review,"var(--amber)"]].map(([l,v,c])=>(
            <div key={l} style={{marginBottom:10}}>
              <div style={{display:"flex", justifyContent:"space-between", fontSize:12, marginBottom:3}}><span>{l}</span><span className="f-mono">{v}</span></div>
              <div style={{height:8, background:"var(--surface-2)"}}><div style={{height:8, width:`${Math.max(v*22,v>0?10:0)}px`, background:c}}/></div>
            </div>
          ))}
        </div>
        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Most Frequent Failure Categories</div>
          {Object.keys(byFailCategory).length===0 && <div style={{fontSize:12, color:"var(--ink-faint)"}}>No failures recorded.</div>}
          {Object.entries(byFailCategory).map(([k,v])=>(
            <div key={k} style={{display:"flex", justifyContent:"space-between", fontSize:12.5, padding:"6px 0", borderBottom:"1px solid var(--line)"}}><span>{k}</span><span className="f-mono">{v}</span></div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ============================== AUDIT TRAIL ============================== */
function AuditPage({db, nav}){
  const [q, setQ] = useState("");
  const [entityFilter, setEntityFilter] = useState("all");
  const entities = ["all", ...Array.from(new Set(db.auditLog.map(a=>a.entity)))];
  const rows = db.auditLog.filter(a=>{
    if(entityFilter!=="all" && a.entity!==entityFilter) return false;
    return !q || `${a.user} ${a.action} ${a.entityId}`.toLowerCase().includes(q.toLowerCase());
  });
  return (
    <div>
      <PageHeader eyebrow="Traceability" title="Digital Audit Trail" desc="Every significant action across the platform, with user, timestamp, previous and new values. Rows link directly to the affected instrument, test or report."/>
      <div style={{display:"flex", gap:8, marginBottom:14, flexWrap:"wrap"}}>
        <div style={{position:"relative", flex:1, maxWidth:280}}>
          <Search size={13} style={{position:"absolute", left:9, top:10, color:"var(--ink-faint)"}}/>
          <input className="n-input" style={{paddingLeft:28}} placeholder="Search user, action, entity…" value={q} onChange={e=>setQ(e.target.value)}/>
        </div>
        {entities.map(e=>(
          <button key={e} className="n-btn n-btn-sm" style={entityFilter===e?{background:"var(--navy)", color:"#fff", borderColor:"var(--navy)"}:{}} onClick={()=>setEntityFilter(e)}>{e==="all"?"All":e}</button>
        ))}
      </div>
      <div className="n-panel">
        <table className="n-table">
          <thead><tr><th>Timestamp</th><th>User</th><th>Action</th><th>Entity</th><th>Previous</th><th>New</th><th>Reason</th><th></th></tr></thead>
          <tbody>
            {rows.map(a=>(
              <tr key={a.id} className={nav?"n-row-hover":""} onClick={()=>nav && jumpToEntity(a, nav, db)}>
                <td className="f-mono">{fmtDT(a.timestamp)}</td>
                <td>{a.user}</td>
                <td>{a.action}</td>
                <td className="f-mono">{a.entityId}</td>
                <td style={{color:"var(--ink-faint)"}}>{a.prevValue}</td>
                <td>{a.newValue}</td>
                <td style={{color:"var(--ink-dim)"}}>{a.reason}</td>
                <td>{nav && <ChevronRight size={13} color="var(--ink-faint)"/>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== USERS ============================== */
function UsersPage({db, role, currentUserName}){
  const ROLE_DESC = {
    Admin:"Full access — manage users, OIML rule versions, all data.",
    Technician:"Register instruments, generate test plans, record observations, submit for review.",
    Examiner:"Review observations and calculations; approve, reject, or request correction.",
    Authority:"Final approval of reports; view analytics and repository.",
    Viewer:"Read-only access across all modules.",
  };
  return (
    <div>
      <PageHeader eyebrow="Access Control" title="Users & Roles" desc="Role-based access controls which actions are available. The active identity comes from the authenticated server session."/>
      <div className="n-panel" style={{marginBottom:16}}>
        <table className="n-table">
          <thead><tr><th>Name</th><th>Role</th><th>Capabilities</th><th></th></tr></thead>
          <tbody>
            {db.users.map(u=>(
              <tr key={u.id}>
                <td>{u.name}</td>
                <td><span className="n-badge n-badge-navy">{u.role}</span></td>
                <td style={{color:"var(--ink-dim)"}}>{ROLE_DESC[u.role]}</td>
                <td>{u.name===currentUserName ? <span className="n-badge n-badge-pass">ACTIVE SESSION</span> : <span className="n-badge n-badge-neutral">SERVER MANAGED</span>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="n-panel" style={{padding:18}}>
        <div className="n-eyebrow" style={{marginBottom:8}}>Authenticated Session</div>
        <div style={{fontSize:14, fontWeight:600}}>{currentUserName} — {role}</div>
      </div>
    </div>
  );
}

/* ============================== RULES ============================== */
function RulesPage({db, mutate, logAudit, role, notify}){
  const [showAdd, setShowAdd] = useState(false);
  const [notes, setNotes] = useState("");
  const [versionLabel, setVersionLabel] = useState("");
  const [effective, setEffective] = useState(new Date().toISOString().slice(0,10));
  const [classDraft, setClassDraft] = useState(cloneRuleTable());
  const [mpeClass, setMpeClass] = useState("III");
  const [validationMsg, setValidationMsg] = useState("");
  const canEdit = can(role,"rule:manage");
  const active = db.ruleVersions.find(r=>r.status==="ACTIVE") || activeRule();
  const activeTable = active?.classTable || OIML_CLASS_TABLE;
  const activeSteps = stepsForClass(active, mpeClass) || [];
  const mpeDraft = classDraft[mpeClass]?.mpeSteps || [];

  const openEditor = () => {
    const source = active || activeRule();
    setVersionLabel(`${source.version} — Controlled Revision`);
    setEffective(new Date().toISOString().slice(0,10));
    setNotes("");
    setClassDraft(Object.fromEntries(Object.keys(OIML_CLASS_TABLE).map(k=>{ const c=JSON.parse(JSON.stringify((source.classTable||{})[k]||OIML_CLASS_TABLE[k])); c.mpeSteps=JSON.parse(JSON.stringify(stepsForClass(source,k))); return [k,c]; })));
    setMpeClass("III");
    setValidationMsg("");
    setShowAdd(true);
  };

  const validateDraft = () => {
    const errors=[];
    Object.entries(classDraft).forEach(([cls,c])=>{
      if(!Number.isFinite(Number(c.nMin)) || Number(c.nMin)<=0) errors.push(`${cls}: Min n must be greater than zero.`);
      if(!Number.isFinite(Number(c.nMax)) || Number(c.nMax)<=Number(c.nMin)) errors.push(`${cls}: Max n must be greater than Min n.`);
      if(!Number.isFinite(Number(c.minCapacityInE)) || Number(c.minCapacityInE)<=0) errors.push(`${cls}: minimum capacity in e must be greater than zero.`);
    });
    Object.entries(classDraft).forEach(([cls,c])=>{
      const steps=c.mpeSteps||[];
      if(!steps.length) errors.push(`Class ${cls}: at least one MPE step is required.`);
      steps.forEach((step,i)=>{
        if(!Number.isFinite(Number(step.mpeE)) || Number(step.mpeE)<0) errors.push(`Class ${cls} MPE step ${i+1}: MPE in e must be zero or greater.`);
        if(step.upTo!==null && (!Number.isFinite(Number(step.upTo)) || Number(step.upTo)<=0)) errors.push(`Class ${cls} MPE step ${i+1}: upper interval limit must be positive or Unlimited.`);
        if(i>0){
          const prev=steps[i-1].upTo===null?Infinity:Number(steps[i-1].upTo);
          const cur=step.upTo===null?Infinity:Number(step.upTo);
          if(cur<=prev) errors.push(`Class ${cls} MPE step ${i+1}: upper interval limit must increase monotonically.`);
        }
      });
      if(steps.length && steps[steps.length-1]?.upTo!==null) errors.push(`Class ${cls}: the final MPE step must use Unlimited as its upper interval limit.`);
    });
    return errors;
  };

  const publish = () => {
    if(!canEdit){ notify("Your role does not permit changing OIML rule configuration."); return; }
    const errors=validateDraft();
    if(errors.length){ setValidationMsg(errors.join(" ")); return; }
    const current = active || activeRule();
    const id=uid("RCFG");
    const v={
      id,
      version:versionLabel.trim() || `${current.version} — Revision ${new Date().toISOString().slice(0,10)}`,
      effective,
      status:"ACTIVE",
      notes:notes.trim() || "Administrator-issued controlled rule configuration update.",
      classTable:Object.fromEntries(Object.entries(classDraft).map(([k,c])=>[k,{...JSON.parse(JSON.stringify(c)),nMin:Number(c.nMin),nMax:Number(c.nMax),minCapacityInE:Number(c.minCapacityInE),mpeSteps:(c.mpeSteps||[]).map(x=>({upTo:(x.upTo===null||x.upTo===undefined)?null:Number(x.upTo),mpeE:Number(x.mpeE)}))}])),
    };
    mutate(d=>{
      d.ruleVersions=(d.ruleVersions||[]).map(r=>({...r,status:r.status==="ACTIVE"?"SUPERSEDED":r.status}));
      d.ruleVersions.unshift(v);
      logAudit(d,"OIML rule configuration published","RuleVersion",v.id,current?.version||"—",v.version,"Controlled administrator revision; applies to new evaluations only");
    });
    RULE_VERSIONS.forEach(r=>{ if(r.status==="ACTIVE") r.status="SUPERSEDED"; });
    RULE_VERSIONS.unshift(v);
    notify("Controlled rule revision published. New evaluations use it; existing tests and reports retain their recorded rule version.");
    setShowAdd(false);
  };

  const updateClass=(cls,key,value)=>setClassDraft(prev=>({...prev,[cls]:{...prev[cls],[key]:value}}));
  const setClassSteps=(fn)=>setClassDraft(prev=>({...prev,[mpeClass]:{...prev[mpeClass],mpeSteps:fn(prev[mpeClass]?.mpeSteps||[])}}));
  const updateMpe=(idx,key,value)=>setClassSteps(list=>list.map((x,i)=>i===idx?{...x,[key]:value}:x));
  const addMpe=()=>setClassSteps(list=>[...list,{upTo:null,mpeE:1.5}]);
  const removeMpe=(idx)=>setClassSteps(list=>list.length<=1?list:list.filter((_,i)=>i!==idx));

  return (
    <div>
      <PageHeader eyebrow="Controlled Configuration" title="OIML Rule Version Management"
        desc="Rule thresholds and class limits are stored as versioned data. Administrators can publish a controlled configuration revision without editing individual test components. Historical evaluations retain the rule version they used."
        right={canEdit && <button className="n-btn n-btn-primary" onClick={openEditor}><Plus size={14}/> Create Revision</button>}/>

      <div className="n-panel" style={{padding:18,marginBottom:16}}>
        <div className="n-eyebrow" style={{marginBottom:6}}>Active configuration</div>
        <div style={{fontSize:16,fontWeight:700}}>{active?.version||"No active rule"}</div>
        <div style={{fontSize:11.5,color:"var(--ink-faint)",marginTop:5}}>Effective {active?.effective||"—"} · {Object.keys(activeTable).length} accuracy classes · per-class MPE steps · in-service multiplier {IN_SERVICE_MULTIPLIER}×</div>
      </div>

      <div className="n-panel">
        <table className="n-table">
          <thead><tr><th>Version</th><th>Effective</th><th>Status</th><th>Configuration</th><th>Notes</th></tr></thead>
          <tbody>
            {(db.ruleVersions||[]).map(r=>(
              <tr key={r.id}>
                <td className="f-mono">{r.version}</td>
                <td className="f-mono">{r.effective}</td>
                <td>{r.status==="ACTIVE" ? <span className="n-badge n-badge-pass">ACTIVE</span> : <span className="n-badge n-badge-neutral">SUPERSEDED</span>}</td>
                <td style={{fontSize:11,color:"var(--ink-dim)"}}>{Object.keys(r.classTable||{}).length} classes · per-class MPE steps</td>
                <td style={{color:"var(--ink-dim)"}}>{r.notes}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="n-panel" style={{marginTop:16}}>
        <div className="n-eyebrow" style={{marginBottom:10}}>Active accuracy-class configuration</div>
        <table className="n-table">
          <thead><tr><th>Class</th><th>Min n</th><th>Max n</th><th>Min Capacity</th><th>Coarse-e band (R 76-1 Table 3)</th></tr></thead>
          <tbody>{Object.entries(activeTable).map(([k,c])=><tr key={k}><td className="f-mono">{k}</td><td className="f-mono">{Number(c.nMin).toLocaleString()}</td><td className="f-mono">{Number(c.nMax).toLocaleString()}</td><td className="f-mono">{c.minCapacityInE}e</td><td className="f-mono">{c.coarse?`e ≥ ${c.coarse.eFromG} g: n ≥ ${Number(c.coarse.nMin).toLocaleString()}, Min ${c.coarse.minCapacityInE}e`:"—"}</td></tr>)}</tbody>
        </table>
      </div>

      <div className="n-panel" style={{marginTop:16}}>
        <div className="n-eyebrow" style={{marginBottom:10}}>Active MPE configuration — Class {mpeClass}</div>
        <div style={{display:"flex",gap:6,marginBottom:10}}>{Object.keys(OIML_CLASS_TABLE).map(k=><button key={k} className={"n-btn n-btn-sm"+(k===mpeClass?" n-btn-primary":"")} onClick={()=>setMpeClass(k)}>Class {k}</button>)}</div>
        <table className="n-table">
          <thead><tr><th>Load in e, up to</th><th>Initial MPE</th><th>In-service MPE</th></tr></thead>
          <tbody>{activeSteps.map((s,i)=><tr key={i}><td className="f-mono">{s.upTo===null?"Unlimited":Number(s.upTo).toLocaleString()}</td><td className="f-mono">±{s.mpeE}e</td><td className="f-mono">±{round(Number(s.mpeE)*IN_SERVICE_MULTIPLIER,2)}e</td></tr>)}</tbody>
        </table>
        <div style={{fontSize:11,color:"var(--ink-faint)",marginTop:8}}>The calculation engine reads these versioned values at evaluation time. Changing a future configuration does not rewrite historical test results.</div>
      </div>

      {!canEdit && <div style={{fontSize:11.5,color:"var(--ink-faint)",marginTop:10}}>Only an Admin can create or publish a controlled rule revision.</div>}

      {showAdd && (
        <Modal title="Create Controlled OIML Rule Revision" onClose={()=>setShowAdd(false)}>
          <div style={{fontSize:11.5,color:"var(--ink-dim)",lineHeight:1.6,marginBottom:14}}>Edit the versioned configuration below. This prototype does not independently verify regulatory values. A metrology reviewer must confirm the exact applicable OIML edition before publishing a production rule set.</div>
          <Field label="Version label"><input className="n-input" value={versionLabel} onChange={e=>setVersionLabel(e.target.value)}/></Field>
          <Field label="Effective date"><input className="n-input" type="date" value={effective} onChange={e=>setEffective(e.target.value)}/></Field>
          <Field label="Change notes"><textarea className="n-textarea" rows={2} value={notes} onChange={e=>setNotes(e.target.value)}/></Field>

          <div className="n-eyebrow" style={{margin:"16px 0 8px"}}>Accuracy class limits</div>
          <div style={{overflowX:"auto"}}><table className="n-table"><thead><tr><th>Class</th><th>Min n</th><th>Max n</th><th>Min capacity in e</th></tr></thead><tbody>
            {Object.entries(classDraft).map(([cls,c])=><tr key={cls}><td className="f-mono">{cls}</td>
              {[["nMin","Min n"],["nMax","Max n"],["minCapacityInE","Min capacity"]].map(([key])=><td key={key}><input className="n-input" type="number" value={c[key]} onChange={e=>updateClass(cls,key,e.target.value)}/></td>)}
            </tr>)}
          </tbody></table></div>

          <div className="n-eyebrow" style={{margin:"16px 0 8px"}}>MPE steps — Class {mpeClass}</div>
          <div style={{display:"flex",gap:6,marginBottom:8}}>{Object.keys(OIML_CLASS_TABLE).map(k=><button key={k} className={"n-btn n-btn-sm"+(k===mpeClass?" n-btn-primary":"")} onClick={()=>setMpeClass(k)}>Class {k}</button>)}</div>
          <div style={{overflowX:"auto"}}><table className="n-table"><thead><tr><th>Load in e, up to</th><th>MPE in e</th><th>Action</th></tr></thead><tbody>
            {mpeDraft.map((step,i)=><tr key={i}><td>{step.upTo===null?<span className="f-mono">Unlimited</span>:<input className="n-input" type="number" value={step.upTo} onChange={e=>updateMpe(i,"upTo",e.target.value)}/>}</td><td><input className="n-input" type="number" step="0.01" value={step.mpeE} onChange={e=>updateMpe(i,"mpeE",e.target.value)}/></td><td><button className="n-btn n-btn-sm" onClick={()=>removeMpe(i)} disabled={mpeDraft.length<=1}>Remove</button></td></tr>)}
          </tbody></table></div>
          <button className="n-btn n-btn-sm" style={{marginTop:8}} onClick={addMpe}>Add MPE step</button>

          {validationMsg && <div style={{marginTop:12,padding:10,border:"1px solid var(--fail-line)",background:"var(--fail-bg)",color:"var(--fail)",fontSize:11.5,lineHeight:1.5}}>{validationMsg}</div>}
          <div style={{display:"flex",justifyContent:"flex-end",gap:8,marginTop:16}}><button className="n-btn" onClick={()=>setShowAdd(false)}>Cancel</button><button className="n-btn n-btn-primary" onClick={publish}>Publish Revision</button></div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== TECHNICAL DOCUMENTATION ============================== */
function TechnicalDocumentationPage({nav}){
  const sections = [
    {id:"architecture", title:"1. Software Architecture", body:[
      "METRA is implemented as a client-side React prototype. The application shell provides navigation, shared state, role-aware actions, reporting, audit history and rule configuration in one workspace.",
      "The current prototype keeps demonstration data in application state. A production deployment should move persistence, authentication, authorization enforcement and file storage to a secured backend service and database."
    ]},
    {id:"workflow", title:"2. Evaluation Workflow", body:[
      "Instrument registration → technical configuration → rule selection → reference-equipment and calibration checks → environmental recording → applicable-test plan → observation capture → calculation and validation → technician submission → review → report generation → repository/history.",
      "Submitted tests are locked for technician measurement entry. A permitted reviewer can approve, reject or return a test for correction, with actions recorded in the audit trail."
    ]},
    {id:"calculations", title:"3. Calculation Methodology", body:[
      "The active rule engine is centralized in the OIML R-76 configuration and calculation functions. Accuracy evaluation uses instrument accuracy class, maximum capacity, verification scale interval (e), load and observed indication to determine the applicable maximum permissible error (MPE).",
      "The prototype supports Class I, II, III and IIII configurations and distinguishes initial-verification and in-service stages. The current active configuration is versioned as OIML R-76-1:2006 (Ed. 2017) — Class-Aware Configuration v2.0.",
      "Specialized workflows calculate or validate repeatability, zero return, discrimination, eccentricity, sensitivity, creep, tare, temperature effect, warm-up, voltage variation and span stability according to the fields implemented for each workflow.",
      "The calculation explanation shown inside a test is derived from stored calculation fields so the displayed PASS/FAIL rationale follows the calculation that produced the result."
    ]},
    {id:"validation", title:"4. Validation & Compliance", body:[
      "Validation operates before submission and includes instrument configuration, accuracy class, rule configuration, reference equipment, calibration validity, environmental records, required observations and test-specific data checks.",
      "A test result is generated by the rule/calculation layer; the AI Test Assistant does not make compliance decisions. Final approval remains a human workflow action."
    ]},
    {id:"data", title:"5. Data Model", body:[
      "Core records include instruments, test plans, tests, observations, environmental readings, reference equipment, reports, audit entries, users and versioned OIML rule configurations.",
      "Evidence records can contain uploaded photographs or supporting documents associated with a test. Report generation can include evidence metadata and supported image attachments."
    ]},
    {id:"reporting", title:"6. Reporting & Repository", body:[
      "Reports are assembled from the instrument record, applicable tests, not-applicable tests, environmental readings, rule version, reference equipment, workflow users and evidence. The prototype provides editable Word (.docx) export, native PDF generation, browser printing, QR verification, and immutable report snapshots.",
      "The repository groups report history by instrument and supports search/retrieval through the application interface."
    ]},
    {id:"security", title:"7. Security & Access Model", body:[
      "Role permissions are defined for Technician, Examiner, Authority, Admin and Viewer. State-changing handlers check the active role before allowing permitted actions.",
      "Important prototype limitation: role selection is client-side demonstration logic and is not authentication. Production deployment requires authenticated identities, server-side authorization, secure sessions, encrypted transport, protected storage and an appropriate audit/security architecture."
    ]},
    {id:"deployment", title:"8. Deployment Framework", body:[
      "Prototype deployment can run as a web application. For production use, the recommended architecture is a browser client connected to an authenticated application/API layer, a transactional database for instruments/tests/reports/audit records, and controlled object storage for evidence files.",
      "Production operations should include backups, monitoring, access logging, environment-specific configuration, database migration/versioning and controlled release of OIML rule configurations."
    ]},
    {id:"revision", title:"9. OIML Revision Management", body:[
      "Rule versions are stored separately and historical versions are retained as superseded configurations for traceability. Tests retain the rule-version identifier used for evaluation.",
      "The current Rules module allows version records to be added, but the MPE/class tables are still code-level configuration. A future production release should provide controlled authoring, validation, approval and activation of rule parameters without editing individual test components."
    ]},
    {id:"limitations", title:"10. Prototype Status & Review Requirements", body:[
      "This document describes the implemented prototype architecture and does not certify compliance with every clause of any OIML R-76 edition.",
      "Before regulatory or certification use, the applicable OIML edition, national/legal-metrology procedure, laboratory procedure and every implemented test calculation should be independently reviewed and validated against authoritative requirements.",
      "The prototype does not currently provide production authentication/security infrastructure or a server-backed persistence layer. Digital signatures are optional in the problem statement and are not implemented as a production cryptographic signature service."
    ]}
  ];
  return (
    <div>
      <PageHeader eyebrow="Engineering" title="Technical Documentation" desc="Architecture, calculation methodology and deployment framework for the METRA prototype." right={<button className="n-btn" onClick={()=>nav("rules")}>Open OIML Rules <ChevronRight size={13}/></button>}/>
      <div className="n-panel" style={{padding:18, marginBottom:16, borderLeft:"3px solid var(--navy)"}}>
        <div className="n-eyebrow" style={{marginBottom:8}}>Document status</div>
        <div style={{fontSize:13, fontWeight:700}}>Prototype technical documentation · Review before production or regulatory use</div>
        <div style={{fontSize:11.5, color:"var(--ink-faint)", marginTop:5}}>Scope: current METRA implementation, its calculation/rule architecture, workflow, data model and deployment considerations.</div>
      </div>
      <div style={{display:"grid", gridTemplateColumns:"minmax(190px,240px) 1fr", gap:16, alignItems:"start"}}>
        <div className="n-panel" style={{padding:14, position:"sticky", top:12}}>
          <div className="n-eyebrow" style={{marginBottom:8}}>Contents</div>
          {sections.map(x=><a key={x.id} href={`#tech-${x.id}`} style={{display:"block", padding:"7px 4px", borderBottom:"1px solid var(--line)", color:"var(--ink-dim)", fontSize:11.5}}>{x.title.replace(/^\d+\. /,"")}</a>)}
        </div>
        <div>
          {sections.map(x=>(
            <section id={`tech-${x.id}`} key={x.id} className="n-panel" style={{padding:20, marginBottom:14, scrollMarginTop:20}}>
              <h2 style={{fontSize:15, margin:"0 0 10px", color:"var(--ink)"}}>{x.title}</h2>
              {x.body.map((p,i)=><p key={i} style={{fontSize:12.5, lineHeight:1.65, color:"var(--ink-dim)", margin:i===x.body.length-1?0:"0 0 10px"}}>{p}</p>)}
            </section>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ============================== OFFICIAL REFERENCES ============================== */
function OfficialReferencesPage({db, nav}){
  const refs = db.referenceSources || OFFICIAL_REFERENCE_SOURCES;
  return (
    <div>
      <PageHeader
        eyebrow="Source Registry"
        title="Official Government & Standards References"
        desc="A provenance layer for standards and government sources used by Navi. Reference records are kept separate from laboratory measurements and instrument test data."
        right={<button className="n-btn" onClick={()=>nav("rules")}>OIML Rules <ChevronRight size={13}/></button>}
      />

      <div className="n-panel" style={{padding:18,marginBottom:16,borderLeft:"3px solid var(--navy)"}}>
        <div className="n-eyebrow" style={{marginBottom:7}}>Data provenance rule</div>
        <div style={{fontSize:13,fontWeight:700}}>Official references inform configuration; they do not become laboratory observations.</div>
        <p style={{fontSize:12,lineHeight:1.65,color:"var(--ink-dim)",margin:"7px 0 0"}}>
          METRA does not fabricate government measurements, instrument specifications, approvals or regulatory decisions. A source is recorded with its authority, title, provenance and verification boundary. The exact applicable edition or amendment must be confirmed before certification or regulatory use.
        </p>
      </div>

      <div className="n-panel" style={{padding:0,overflow:"hidden"}}>
        <table className="n-table">
          <thead><tr><th>Authority</th><th>Reference</th><th>Type</th><th>Status</th><th>Use in Navi</th><th>Source</th></tr></thead>
          <tbody>
            {refs.map(r=>(
              <tr key={r.id}>
                <td><strong>{r.authority}</strong></td>
                <td><div style={{fontWeight:650}}>{r.title}</div><div className="f-mono" style={{fontSize:10,color:"var(--ink-faint)",marginTop:3}}>{r.id}</div><div style={{fontSize:10.5,color:"var(--ink-dim)",marginTop:3}}>{r.edition}</div></td>
                <td>{r.type}</td>
                <td><span className={`n-badge ${r.status.includes("VERIFY")?"n-badge-neutral":"n-badge-pass"}`}>{r.status}</span></td>
                <td style={{fontSize:11.5,color:"var(--ink-dim)",maxWidth:280}}>{r.use}<div style={{fontSize:10,color:"var(--ink-faint)",marginTop:4}}>{r.provenance}</div></td>
                <td><a className="n-btn n-btn-sm" href={r.url} target="_blank" rel="noreferrer">Open official source <ExternalLink size={11}/></a></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="n-panel" style={{padding:18,marginTop:16}}>
        <div className="n-eyebrow" style={{marginBottom:9}}>Reference governance</div>
        <div className="n-kv">
          <dt>Reference records</dt><dd>{refs.length}</dd>
          <dt>Official-source records</dt><dd>{refs.filter(r=>r.status.includes("OFFICIAL")).length}</dd>
          <dt>Technical recommendation record</dt><dd>{refs.some(r=>r.id==="SRC-OIML-R76")?"Present":"Missing"}</dd>
          <dt>Laboratory data separation</dt><dd>Maintained</dd>
          <dt>Production verification</dt><dd>Required before regulatory use</dd>
        </div>
      </div>
    </div>
  );
}

/* ============================== SECURITY CENTER ============================== */
function SecurityPage({db, role, currentUserName, nav}){
  const permissions = PERMISSIONS;
  const roleList = Object.keys(permissions);
  const auditCount = db.auditLog?.length || 0;
  const userCount = db.users?.length || 0;
  const activeRoleUsers = db.users?.filter(u=>u.role===role).length || 0;
  const privilegedRoles = ["Admin","Authority","Examiner"];
  const privilegedUsers = db.users?.filter(u=>privilegedRoles.includes(u.role)).length || 0;
  const currentUser = db.users?.find(u=>u.name===currentUserName);
  const auditFingerprint = hashStr((db.auditLog||[]).map(a=>`${a.id}|${a.timestamp}|${a.user}|${a.action}|${a.entityId}`).join("||")).slice(0,18);

  const controlRows = [
    {control:"Role permission checks", status:"ACTIVE IN PROTOTYPE", tone:"pass", detail:"Protected state-changing handlers check the active role against the central PERMISSIONS matrix."},
    {control:"Role-based navigation", status:"ACTIVE IN PROTOTYPE", tone:"pass", detail:"Dashboard actions and navigation are tailored to the active role."},
    {control:"Audit event capture", status:"ACTIVE IN PROTOTYPE", tone:"pass", detail:"Workflow changes are recorded with actor, action, entity, previous value, new value and reason."},
    {control:"Trusted authentication", status:"ACTIVE", tone:"pass", detail:"Login is handled by the METRA backend with an HttpOnly session cookie; the UI receives the authenticated user and role."},
    {control:"Server-side authorization", status:"ACTIVE", tone:"pass", detail:"Repository writes are checked by the backend against role permissions and blocked when unauthorized collections are changed."},
    {control:"Protected repository", status:"ACTIVE", tone:"pass", detail:"Authenticated sessions load and save the server repository. Browser storage is retained only as an offline initialization fallback."},
    {control:"HTTPS / secure transport", status:"DEPLOYMENT REQUIRED", tone:"warn", detail:"Production deployment must use HTTPS/TLS for application, API and file-transfer endpoints."},
    {control:"Session expiry / revocation", status:"BACKEND REQUIRED", tone:"warn", detail:"Production authentication should support expiry, logout/revocation, idle timeout and re-authentication for sensitive actions."},
  ];

  const roleDescriptions = {
    Technician:"Register instruments, record tests/evidence and submit work for review.",
    Examiner:"Review submitted test evidence and return or approve examination decisions.",
    Authority:"Perform final report approval and authority-level review.",
    Admin:"Manage users, rules and administrative controls.",
    Viewer:"Read-only operational visibility."
  };

  return (
    <div>
      <PageHeader
        eyebrow="Security"
        title="Security & Access Dashboard"
        desc="A transparent view of Navi's current role controls, audit coverage and production security boundary."
        right={<button className="n-btn" onClick={()=>nav("users")}>Users & Roles <ChevronRight size={13}/></button>}
      />

      <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:12,marginBottom:16}}>
        <MetricBlock label="Users" value={userCount}/>
        <MetricBlock label="Privileged users" value={privilegedUsers}/>
        <MetricBlock label="Audit events" value={auditCount}/>
        <MetricBlock label="Active role" value={role}/>
      </div>

      <div className="n-panel" style={{padding:18,marginBottom:16,borderLeft:"3px solid var(--amber)"}}>
        <div className="n-eyebrow" style={{marginBottom:7}}>Security boundary</div>
        <div style={{fontSize:13,fontWeight:700}}>Authenticated server session + server-enforced RBAC</div>
        <p style={{fontSize:12,lineHeight:1.65,color:"var(--ink-dim)",margin:"7px 0 0"}}>
          METRA authenticates the session on the backend, uses the returned server identity for the active role, and independently checks repository write permissions on the server. The remaining production boundary is deployment hardening (HTTPS, durable database/object storage and operational monitoring).
        </p>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1.15fr .85fr",gap:16,marginBottom:16}}>
        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Current session</div>
          <div className="n-kv">
            <dt>Authenticated user</dt><dd>{currentUserName}</dd>
            <dt>Role</dt><dd>{role}</dd>
            <dt>Directory match</dt><dd>{currentUser ? "Matched to user record" : "No matching user record"}</dd>
            <dt>Users with this role</dt><dd>{activeRoleUsers}</dd>
            <dt>Authentication</dt><dd>Backend session · HttpOnly cookie</dd>
          </div>
        </div>
        <div className="n-panel" style={{padding:18}}>
          <div className="n-eyebrow" style={{marginBottom:10}}>Audit integrity indicator</div>
          <div className="f-mono" style={{fontSize:19,fontWeight:700,letterSpacing:.5}}>{auditFingerprint || "—"}</div>
          <div style={{fontSize:11.5,color:"var(--ink-faint)",lineHeight:1.6,marginTop:7}}>
            Client display fingerprint for the currently loaded audit set. The authoritative audit append-only check is enforced by the backend; production should additionally use durable database controls and operational key management.
          </div>
          <button className="n-btn n-btn-sm" style={{marginTop:10}} onClick={()=>nav("audit")}>Open Audit Trail</button>
        </div>
      </div>

      <div className="n-panel" style={{marginBottom:16}}>
        <div style={{padding:18,borderBottom:"1px solid var(--line)"}}>
          <div className="n-eyebrow">Security controls</div>
        </div>
        <table className="n-table">
          <thead><tr><th>Control</th><th>Status</th><th>What Navi does now</th></tr></thead>
          <tbody>{controlRows.map(c=>(
            <tr key={c.control}>
              <td style={{fontWeight:650}}>{c.control}</td>
              <td><span className={`n-badge ${c.tone==="pass"?"n-badge-pass":"n-badge-review"}`}>{c.status}</span></td>
              <td style={{color:"var(--ink-dim)",lineHeight:1.5}}>{c.detail}</td>
            </tr>
          ))}</tbody>
        </table>
      </div>

      <div className="n-panel" style={{padding:18,marginBottom:16}}>
        <div className="n-eyebrow" style={{marginBottom:10}}>Role permission matrix</div>
        <div style={{overflowX:"auto"}}>
          <table className="n-table">
            <thead><tr><th>Permission</th>{roleList.map(r=><th key={r}>{r}</th>)}</tr></thead>
            <tbody>
              {[...new Set(roleList.flatMap(r=>permissions[r]))].map(action=>(
                <tr key={action}>
                  <td style={{fontFamily:"ui-monospace,SFMono-Regular,Menlo,monospace",fontSize:11.5}}>{action}</td>
                  {roleList.map(r=><td key={r}>{permissions[r].includes(action)?<span className="n-badge n-badge-pass">ALLOW</span>:<span className="n-badge n-badge-neutral">DENY</span>}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="n-panel" style={{padding:18}}>
        <div className="n-eyebrow" style={{marginBottom:10}}>Role responsibilities</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:10}}>
          {roleList.map(r=><div key={r} style={{border:"1px solid var(--line)",padding:13}}>
            <div style={{fontWeight:700,fontSize:12.5,marginBottom:5}}>{r}</div>
            <div style={{fontSize:11.5,color:"var(--ink-dim)",lineHeight:1.55}}>{roleDescriptions[r]}</div>
          </div>)}
        </div>
      </div>
    </div>
  );
}

/* ============================== SETTINGS ============================== */
function SettingsPage({syncStatus, setSyncStatus}){
  return (
    <div>
      <PageHeader eyebrow="System" title="Settings" desc="Laboratory configuration and offline-resilience status."/>
      <div className="n-panel" style={{padding:18, marginBottom:16}}>
        <div className="n-eyebrow" style={{marginBottom:10}}>Laboratory Information</div>
        <div className="n-kv"><dt>Name</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>Regional Reference Standards Laboratory, Panipat</dd><dt>Accreditation Ref.</dt><dd>NABL-CAL-XXXX</dd></div>
      </div>
      <div className="n-panel" style={{padding:18}}>
        <div className="n-eyebrow" style={{marginBottom:10}}>Offline / Resilient Testing</div>
        <div style={{fontSize:12.5, color:"var(--ink-dim)", marginBottom:12}}>Test data captured while offline is stored locally and synchronized automatically when connectivity returns.</div>
        <div style={{display:"flex", gap:8}}>
          <button className="n-btn n-btn-sm" onClick={()=>setSyncStatus("synced")}>Synced</button>
          <button className="n-btn n-btn-sm" onClick={()=>setSyncStatus("pending")}>Pending Sync</button>
          <button className="n-btn n-btn-sm" onClick={()=>setSyncStatus("failed")}>Sync Failed</button>
        </div>
        <div style={{fontSize:12, marginTop:10}}>Current: <strong>{syncStatus}</strong></div>
      </div>
    </div>
  );
}

/* ============================== PROFILE ============================== */
function ProfilePage({db, mutate, logAudit, currentUserName, setCurrentUserName, role, notify, nav}){
  const user = db.users.find(u=>u.name===currentUserName);
  const [form, setForm] = useState({name:user?.name||"", email:user?.email||"", phone:user?.phone||"", designation:user?.designation||""});
  const [saved, setSaved] = useState(false);
  const set = (k,v) => { setForm(f=>({...f,[k]:v})); setSaved(false); };

  const dirty = user && (form.name!==user.name || form.email!==user.email || form.phone!==user.phone || form.designation!==user.designation);

  const save = async () => {
    if(!form.name.trim()){ notify("Name cannot be empty."); return; }
    try{
      const remote=await naviApi("/api/auth/profile",{method:"PUT",body:JSON.stringify(form)});
      const updated=remote.user;
      mutate(d=>{
        const u = d.users.find(x=>x.id===updated.id);
        if(!u) return;
        u.name=updated.name; u.email=updated.email; u.phone=updated.phone; u.designation=updated.designation;
        logAudit(d,"Profile updated","User",u.id,currentUserName,updated.name,"Authenticated self-service profile update");
      });
      if(updated.name!==currentUserName) setCurrentUserName(updated.name);
      setSaved(true);
      notify("Profile updated successfully.");
    }catch(err){ notify(`Profile update failed: ${err.message}`); }
  };
  const cancel = () => { setForm({name:user.name, email:user.email||"", phone:user.phone||"", designation:user.designation||""}); setSaved(false); };

  if(!user) return null;

  return (
    <div>
      <PageHeader eyebrow="Account" title="My Profile" desc="Update your contact details. Your role and permissions can only be changed by an Administrator."/>
      <div style={{display:"grid", gridTemplateColumns:"220px 1fr", gap:20}}>
        <div className="n-panel" style={{padding:20, textAlign:"center", alignSelf:"start"}}>
          <div style={{width:72, height:72, borderRadius:"50%", background:"var(--navy)", color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", fontSize:22, fontWeight:700, margin:"0 auto 12px"}}>
            {form.name.split(" ").map(w=>w[0]).slice(0,2).join("")}
          </div>
          <div style={{fontWeight:700, fontSize:14}}>{form.name}</div>
          <div style={{fontSize:11.5, color:"var(--ink-faint)", marginTop:2}}>{form.designation||user.role}</div>
          <div style={{marginTop:10}}><span className="n-badge n-badge-navy">{role}</span></div>
        </div>

        <div className="n-panel" style={{padding:20}}>
          <div className="n-eyebrow" style={{marginBottom:14}}>Editable Details</div>
          <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:14}}>
            <Field label="Full Name"><input className="n-input" value={form.name} onChange={e=>set("name",e.target.value)}/></Field>
            <Field label="Designation"><input className="n-input" value={form.designation} onChange={e=>set("designation",e.target.value)}/></Field>
            <Field label="Email"><input className="n-input" type="email" value={form.email} onChange={e=>set("email",e.target.value)}/></Field>
            <Field label="Phone"><input className="n-input" value={form.phone} onChange={e=>set("phone",e.target.value)}/></Field>
          </div>
          <div className="n-hr" style={{margin:"6px 0 14px"}}/>
          <div className="n-eyebrow" style={{marginBottom:8}}>Access (Administrator-Controlled)</div>
          <div className="n-kv" style={{marginBottom:16}}>
            <dt>Role</dt><dd style={{fontFamily:'-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif'}}>{user.role}</dd>
            <dt>User ID</dt><dd>{user.id}</dd>
          </div>
          <div style={{display:"flex", gap:8, alignItems:"center"}}>
            <button className="n-btn n-btn-primary" disabled={!dirty} onClick={save}>Save Changes</button>
            <button className="n-btn" disabled={!dirty} onClick={cancel}>Cancel</button>
            {saved && !dirty && <span style={{fontSize:12, color:"var(--seal)", fontWeight:600}}><CheckCircle2 size={14} style={{verticalAlign:"-2px"}}/> Saved</span>}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================== AI ASSISTANT ============================== */
function AssistantWidget({db, selectedTest, open, setOpen}){
  const [msgs, setMsgs] = useState([{from:"bot", text:"I'm the AI Test Assistant. I can explain procedures, validation warnings and calculation results — compliance decisions always come from the OIML Rule Engine, never from me."}]);
  const [input, setInput] = useState("");
  const test = db.tests.find(t=>t.id===selectedTest);

  const respond = (q) => {
    const s = q.toLowerCase();
    let reply = "";
    if(s.includes("next") || s.includes("what should i")){
      reply = test ? `For ${test.testName}, enter the test load and observed indication; the system calculates the error and compares it to the permissible error automatically.` : "Open a test from the Tests module and I can guide you through the required observations.";
    } else if(s.includes("error") || s.includes("fail")){
      reply = test ? explainResult(test, db.instruments.find(i=>i.id===test.instrumentId)) : "Select a test and I'll explain its latest calculation.";
    } else if(s.includes("missing") || s.includes("required")){
      reply = test ? `Required for ${test.testName}: ${TEST_DEFS[test.testKey].requiredObservations}.` : "Tell me which test you're working on, or open it first.";
    } else if(s.includes("equipment")){
      reply = test ? TEST_DEFS[test.testKey].requiredEquipment : "Select a test to see its required equipment.";
    } else if(s.includes("rule") || s.includes("mpe") || s.includes("oiml")){
      reply = `Compliance is evaluated under ${activeRule().version}. See the OIML Rules module for the full version history.`;
    } else {
      reply = "I can help explain test procedures, validation warnings, required observations, or the current calculation result. Try asking 'what should I enter next?'";
    }
    setMsgs(m=>[...m, {from:"user", text:q}, {from:"bot", text:reply}]);
  };

  return (
    <div className="no-print" style={{position:"fixed", right:24, bottom:24, zIndex:150}}>
      {open && (
        <div className="n-panel" style={{width:340, height:440, display:"flex", flexDirection:"column", marginBottom:12, borderRadius:24, boxShadow:"0 20px 50px -10px rgba(0,0,0,0.15), 0 1px 3px rgba(0,0,0,0.05)", border:"1px solid var(--line)", overflow:"hidden"}}>
          <div style={{padding:"14px 18px", borderBottom:"1px solid var(--line)", display:"flex", justifyContent:"space-between", alignItems:"center", background:"var(--accent)", color:"#fff"}}>
            <div style={{display:"flex", alignItems:"center", gap:8, fontSize:13, fontWeight:750}}><Bot size={16}/> AI Test Assistant</div>
            <span style={{cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", width:24, height:24, borderRadius:"50%", background:"rgba(255,255,255,0.15)"}} onClick={()=>setOpen(false)}><X size={14}/></span>
          </div>
          <div className="n-scroll" style={{flex:1, overflow:"auto", padding:"14px 16px"}}>
            {msgs.map((m,i)=>(
              <div key={i} style={{marginBottom:10, display:"flex", justifyContent: m.from==="user"?"flex-end":"flex-start"}}>
                <div style={{maxWidth:"85%", fontSize:12.5, lineHeight:1.5, padding:"9px 14px", borderRadius:16, background: m.from==="user"?"var(--accent)":"var(--surface-2)", color: m.from==="user"?"#fff":"var(--ink)"}}>{m.text}</div>
              </div>
            ))}
          </div>
          <div style={{padding:"10px 14px", borderTop:"1px solid var(--line)", display:"flex", gap:8, background:"var(--surface)"}}>
            <input className="n-input" placeholder="Ask about this test…" value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&input){respond(input); setInput("");}}} style={{borderRadius:9999, fontSize:12.5}}/>
            <button className="n-btn n-btn-sm n-btn-primary" onClick={()=>{if(input){respond(input); setInput("");}}} style={{borderRadius:9999, padding:"0 14px"}}><Send size={13}/></button>
          </div>
        </div>
      )}
      <button className="n-btn n-btn-primary" style={{borderRadius:"50%", width:52, height:52, padding:0, justifyContent:"center", boxShadow:"0 8px 24px rgba(17,24,39,0.22)"}} onClick={()=>setOpen(o=>!o)}><Bot size={22}/></button>
    </div>
  );
}
