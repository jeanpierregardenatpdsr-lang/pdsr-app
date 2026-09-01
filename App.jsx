import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType, ImageRun, PageBreak, HeadingLevel } from "docx";
import { Home, Users, FileText, Calendar, AlertTriangle, BarChart2, LogOut, Menu, X, ChevronRight, Plus, Check, ChevronLeft, Search, Bell, RefreshCw, Moon, Sun, MapPin, Printer, Download, Package, Wrench, Settings, Clock, Send, Briefcase, ClipboardList, Target, UserCheck, BookOpen, CalendarDays, GraduationCap } from "lucide-react";

/* Chargement robuste de SheetJS : 3 CDN en cascade, script classique (window.XLSX), sans dependance npm */
let _xlsxPromise = null;
function loadXLSX(){
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (_xlsxPromise) return _xlsxPromise;
  const urls = [
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
    "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js",
    "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"
  ];
  _xlsxPromise = new Promise((resolve, reject) => {
    const tryNext = (i) => {
      if (i >= urls.length) { _xlsxPromise = null; reject(new Error("Librairie Excel inaccessible (verifier la connexion)")); return; }
      const s = document.createElement("script");
      s.src = urls[i];
      s.onload = () => { window.XLSX ? resolve(window.XLSX) : tryNext(i + 1); };
      s.onerror = () => { s.remove(); tryNext(i + 1); };
      document.head.appendChild(s);
    };
    tryNext(0);
  });
  return _xlsxPromise;
}



export class ErrorBoundary extends React.Component{constructor(p){super(p);this.state={hasError:false,error:null};}static getDerivedStateFromError(e){return{hasError:true,error:e};}componentDidCatch(e,i){console.error("PDSR Error:",e,i);}render(){if(this.state.hasError){return React.createElement("div",{style:{padding:40,textAlign:"center"}},React.createElement("h2",null,"Une erreur est survenue"),React.createElement("p",null,String(this.state.error)),React.createElement("button",{onClick:()=>{localStorage.removeItem("pdsr_data");window.location.reload();},style:{padding:"10px 20px",background:"#2c6fbb",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",marginTop:16}},"Recharger l'application"));}return this.props.children;}}
const APP_BUILD="2026-09-01-n";
const RESET_KEY="pdsr_reset";
const getLocalReset=()=>{try{return Number(localStorage.getItem(RESET_KEY)||0);}catch(e){return 0;}};
const setLocalReset=(v)=>{try{localStorage.setItem(RESET_KEY,String(v));}catch(e){}};
const LS_KEY="pdsr_data";
const loadLS=()=>{try{const d=JSON.parse(localStorage.getItem(LS_KEY));return d||null;}catch{return null;}};
const saveLS=(o)=>{try{const prev=JSON.parse(localStorage.getItem(LS_KEY))||{};localStorage.setItem(LS_KEY,JSON.stringify({...prev,...o,ts:Date.now()}));}catch{}};
const FB_URL="https://pdsr-app-default-rtdb.firebaseio.com";
const FB_SECRET="GhVuY7AXz2QeB8shFhV4SEZvQAyrrf0YjeOqleEw";
let _clkOff=0;
const _captureClock=(r)=>{try{const h=r&&r.headers&&r.headers.get("date");if(h){const st=Date.parse(h);if(!isNaN(st)){const off=st-Date.now();if(Math.abs(off-_clkOff)>2000)_clkOff=off;}}}catch(e){}};
const nowSrv=()=>new Date(Date.now()+_clkOff).toISOString();
const fbSet=async(p,d)=>{try{const r=await fetch(FB_URL+"/"+p+".json?auth="+FB_SECRET,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(d)});_captureClock(r);}catch(e){}};
const _curBundle=(()=>{try{const sc=document.querySelector('script[src*="/assets/"]');return sc?sc.getAttribute("src"):null;}catch(e){return null;}})();
const inPurgeRanges=(ranges,typeName,x)=>{if(!Array.isArray(ranges)||!ranges.length||!x)return false;const st=String(x.updatedAt||x.createdAt||x.horodatage||"");if(!st||!x.date)return false;/* sans horodatage ou sans date : on ne masque jamais */return ranges.some(r=>r&&Array.isArray(r.t)&&r.t.includes(typeName)&&r.from&&r.to&&r.ts&&x.date>=r.from&&x.date<=r.to&&st<String(r.ts));};
const fetchTO=(url,opts,ms)=>{const c=new AbortController();const t=setTimeout(()=>c.abort(),ms||15000);return fetch(url,{...(opts||{}),signal:c.signal}).finally(()=>clearTimeout(t));};
const fbGet=async(p)=>{try{let r=await fetch(FB_URL+"/"+p+".json?auth="+FB_SECRET);_captureClock(r);return await r.json();}catch(e){return null;}};

const LOGO="data:image/png;base64,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";
const C={gold:"var(--gold)",goldDark:"var(--goldDark)",goldLight:"var(--goldLight)",goldGlow:"var(--goldGlow)",orange:"var(--orange)",orangeLight:"var(--orangeLight)",sable:"var(--sable)",sableLight:"var(--sableLight)",sableDark:"var(--sableDark)",dark:"var(--ink)",mid:"var(--inkMid)",light:"var(--inkLight)",white:"var(--surface)",border:"var(--edge)",accent:"var(--accent)",success:"var(--success)",danger:"var(--danger)",info:"var(--info)",primary:"var(--info)",glass:"var(--glass)",goldA22:"var(--goldA22)",primaryA18:"var(--infoA18)",shadowSm:"var(--shSm)",shadowMd:"var(--shMd)",shadowLg:"var(--shLg)",shadowGold:"var(--shGold)"};
const THEME_CSS=`:root{--gold:#B8860B;--goldDark:#8B6508;--goldLight:#FDF3D0;--goldGlow:rgba(184,134,11,0.15);--orange:#E07B00;--orangeLight:#FFF0D0;--sable:#EAD29C;--sableLight:#FAF7EE;--sableDark:#D9BE80;--ink:#1A0E00;--inkMid:#54402A;--inkLight:#6E5735;--surface:#FFFFFF;--edge:#E8D5B0;--accent:#6A1B9A;--success:#2E7D32;--danger:#C62828;--info:#1565C0;--glass:rgba(255,255,255,0.85);--goldA22:rgba(184,134,11,0.13);--infoA18:rgba(21,101,192,0.09);--shSm:0 1px 2px rgba(26,14,0,0.05),0 1px 3px rgba(26,14,0,0.04);--shMd:0 2px 4px rgba(26,14,0,0.05),0 4px 12px rgba(26,14,0,0.06);--shLg:0 8px 28px rgba(26,14,0,0.10);--shGold:0 1px 2px rgba(139,101,8,0.18);--edgeSoft:#F0E4CC;}\n*,*::before,*::after{box-sizing:border-box;}\nbutton,select,input,textarea{font-family:inherit;}\nbutton{touch-action:manipulation;-webkit-tap-highlight-color:transparent;}\ninput,select,textarea{font-size:16px;}\nselect,input:not([type="checkbox"]):not([type="radio"]){min-height:44px;}\n:focus-visible{outline:2px solid var(--gold);outline-offset:2px;}\nbody{-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility;}\nh1,h2,h3{letter-spacing:-0.01em;}\ndiv,span,p,td,th,label{overflow-wrap:anywhere;}\n@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important;}}\nhtml[data-theme="dark"]{--gold:#D9A82E;--goldDark:#C79420;--goldLight:#33270E;--goldGlow:rgba(217,168,46,0.16);--orange:#F0922A;--orangeLight:#33220A;--sable:#4A3A1E;--sableLight:#14100A;--sableDark:#5C4A26;--ink:#F4E9D4;--inkMid:#C7B490;--inkLight:#93805E;--surface:#211A10;--edge:#3B301E;--accent:#B57BD6;--success:#66BB6A;--danger:#EF5350;--info:#64B5F6;--glass:rgba(33,26,16,0.88);--goldA22:rgba(217,168,46,0.14);--infoA18:rgba(100,181,246,0.10);--shSm:0 2px 8px rgba(0,0,0,0.35);--shMd:0 4px 20px rgba(0,0,0,0.45);--shLg:0 8px 32px rgba(0,0,0,0.55);--shGold:0 4px 16px rgba(217,168,46,0.15);}\nhtml,body{background:var(--sableLight);color:var(--ink);transition:background 0.25s ease,color 0.25s ease;max-width:100vw;overflow-x:hidden;}
*{box-sizing:border-box;}`;
(()=>{try{const st=document.createElement("style");st.id="pdsr-theme";st.textContent=THEME_CSS;document.head.appendChild(st);const t=localStorage.getItem("pdsr_theme");if(t)document.documentElement.dataset.theme=t;}catch(e){}})();
const USERS=[{id:1,email:"lmarcille1962@gmail.com",password:"1789",role:"directeur",name:"Laurent Marcille",initials:"LM",site:"Tous"},{id:2,email:"jeanpierregardenatpdsr@gmail.com",password:"pdsr2026",role:"chef_service",isAdmin:true,name:"Jean-Pierre Gardenat",initials:"JP",site:"Tous"},{id:3,email:"Omar Ngom",password:"pdsr2026",role:"chef_service",name:"Omar Ngom",initials:"ON",site:"Tous"},{id:4,login:"go",password:"go2026",role:"educateur",name:"Go",initials:"GO",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[1,8]},{id:5,login:"abdoulaye",password:"abdoulaye2026",role:"educateur",name:"Abdoulaye",initials:"AB",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[2]},{id:6,login:"khady",password:"khady2026",role:"educateur",name:"Khady",initials:"KH",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[3,7]},{id:7,login:"charline",password:"charline2026",role:"educateur",name:"Charline",initials:"CH",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[4,9]},{id:8,login:"momar",password:"momar2026",role:"educateur",name:"Momar",initials:"MO",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[5]},{id:9,login:"emile",password:"emile2026",role:"educateur",name:"Emile",initials:"EM",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[6]},{id:10,login:"lysa",password:"lysa2026",role:"educateur",name:"Lysa",initials:"LY",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[1,8]},{id:11,login:"babacar",password:"babacar2026",role:"educateur",name:"Babacar",initials:"BA",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[2]},{id:12,login:"diarra",password:"diarra2026",role:"educateur",name:"Diarra",initials:"DI",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[4,9]},{id:13,login:"prospere",password:"prospere2026",role:"educateur",name:"Prospère",initials:"PR",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[6]},{id:14,login:"am",password:"am2026",role:"educateur",name:"Am",initials:"AM",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[5]},{id:15,login:"yacine",password:"yacine2026",role:"educateur",name:"Yacine",initials:"YA",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[3,7]},{id:16,login:"sadibou",password:"sadibou2026",role:"educateur",name:"Sadibou",initials:"SA",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[14]},{id:17,login:"ngor",password:"ngor2026",role:"educateur",name:"Ngor",initials:"NG",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[12,16]},{id:18,login:"mahault",password:"mahault2026",role:"educateur",name:"Mahault",initials:"MA",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[15,18]},{id:19,login:"malang",password:"malang2026",role:"educateur",name:"Malang",initials:"MA",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[14]},{id:20,login:"lea",password:"lea2026",role:"educateur",name:"Léa",initials:"LÉ",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[10]},{id:21,login:"assane",password:"assane2026",role:"educateur",name:"Assane",initials:"AS",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[11,17]},{id:22,login:"rouge",password:"rouge2026",role:"educateur",name:"Rouge",initials:"RO",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[16]},{id:23,login:"luc",password:"luc2026",role:"educateur",name:"Luc",initials:"LU",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[15]},{id:24,login:"nicole",password:"nicole2026",role:"educateur",name:"Nicole",initials:"NI",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[11,12]},{id:25,login:"bacary",password:"bacary2026",role:"educateur",name:"Bakary",initials:"BA",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[14,18]},{id:26,login:"fatou",password:"fatou2026",role:"educateur",name:"Fatou",initials:"FA",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[10]},{id:27,login:"kourou",password:"kourou2026",role:"educateur",name:"Kourou",initials:"KO",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[17]}];
const JEUNES=[];
// URL de base de l'API de synthèse IA.
// - App hébergée sur Vercel : laisser "" (même origine).
// - App hébergée sur GitHub Pages : mettre l'URL du projet Vercel qui porte la fonction, ex: "https://pdsr-app.vercel.app"
const SYNTHESE_API_BASE="";
const MAJEURS=[];
const DJI_PLAN={"2026-04-01":{a:false,b:false},"2026-04-02":{a:false,b:true,n:"Arrivée – WARANG"},"2026-04-03":{a:false,b:true,n:"Départ vers RHO"},"2026-04-04":{a:false,b:true,n:"RHO"},"2026-04-05":{a:true,b:true,n:"RHO"},"2026-04-06":{a:true,b:true,n:"Rando Individuelle"},"2026-04-07":{a:true,b:true,n:"Rando Individuelle"},"2026-04-08":{a:true,b:true,n:"Arrivée DJILASS"},"2026-04-09":{a:false,b:true},"2026-04-10":{a:false,b:true},"2026-04-11":{a:false,b:true},"2026-04-12":{a:true,b:false},"2026-04-13":{a:true,b:false},"2026-04-14":{a:true,b:false},"2026-04-15":{a:true,b:false},"2026-04-16":{a:true,b:false},"2026-04-17":{a:true,b:false},"2026-04-18":{a:true,b:false},"2026-04-19":{a:false,b:true},"2026-04-20":{a:false,b:true},"2026-04-21":{a:false,b:true},"2026-04-22":{a:false,b:true},"2026-04-23":{a:false,b:true},"2026-04-24":{a:false,b:true},"2026-04-25":{a:false,b:true},"2026-04-26":{a:true,b:false},"2026-04-27":{a:true,b:false},"2026-04-28":{a:true,b:false},"2026-04-29":{a:true,b:false},"2026-04-30":{a:true,b:false},"2026-05-01":{a:true,b:false},"2026-05-02":{a:true,b:false},"2026-05-03":{a:false,b:true},"2026-05-04":{a:false,b:true},"2026-05-05":{a:false,b:true},"2026-05-06":{a:false,b:true},"2026-05-07":{a:false,b:true},"2026-05-08":{a:false,b:true},"2026-05-09":{a:false,b:true},"2026-05-10":{a:true,b:false},"2026-05-11":{a:true,b:false},"2026-05-12":{a:true,b:false},"2026-05-13":{a:true,b:false},"2026-05-14":{a:true,b:false},"2026-05-15":{a:true,b:false},"2026-05-16":{a:true,b:false},"2026-05-17":{a:false,b:true},"2026-05-18":{a:false,b:true},"2026-05-19":{a:false,b:true},"2026-05-20":{a:false,b:true},"2026-05-21":{a:false,b:true},"2026-05-22":{a:false,b:true},"2026-05-23":{a:false,b:true},"2026-05-24":{a:true,b:false},"2026-05-25":{a:true,b:false},"2026-05-26":{a:true,b:false},"2026-05-27":{a:true,b:false},"2026-05-28":{a:true,b:false},"2026-05-29":{a:true,b:false},"2026-05-30":{a:true,b:false},"2026-05-31":{a:false,b:true},"2026-06-01":{a:false,b:true},"2026-06-02":{a:false,b:true},"2026-06-03":{a:false,b:true},"2026-06-04":{a:false,b:true},"2026-06-05":{a:false,b:true},"2026-06-06":{a:false,b:true},"2026-06-07":{a:true,b:false},"2026-06-08":{a:true,b:false},"2026-06-09":{a:true,b:false},"2026-06-10":{a:true,b:false},"2026-06-11":{a:true,b:false},"2026-06-12":{a:true,b:false,n:"Lecture RMS 1/2"},"2026-06-13":{a:true,b:false},"2026-06-14":{a:false,b:true},"2026-06-15":{a:false,b:true},"2026-06-16":{a:false,b:true},"2026-06-17":{a:false,b:true},"2026-06-18":{a:false,b:true},"2026-06-19":{a:false,b:true,n:"Lecture RMS 2/2"},"2026-06-20":{a:false,b:true},"2026-06-21":{a:true,b:false},"2026-06-22":{a:true,b:false},"2026-06-23":{a:true,b:false},"2026-06-24":{a:true,b:false},"2026-06-25":{a:true,b:false},"2026-06-26":{a:true,b:false},"2026-06-27":{a:true,b:false},"2026-06-28":{a:false,b:true},"2026-06-29":{a:false,b:true},"2026-06-30":{a:false,b:true},"2026-07-01":{a:false,b:true},"2026-07-02":{a:false,b:true},"2026-07-03":{a:false,b:true},"2026-07-04":{a:false,b:true},"2026-07-05":{a:true,b:false},"2026-07-06":{a:true,b:false},"2026-07-07":{a:true,b:false},"2026-07-08":{a:true,b:false},"2026-07-09":{a:true,b:false},"2026-07-10":{a:true,b:false},"2026-07-11":{a:true,b:false},"2026-07-12":{a:false,b:true},"2026-07-13":{a:false,b:true},"2026-07-14":{a:false,b:true},"2026-07-15":{a:false,b:true},"2026-07-16":{a:false,b:true},"2026-07-17":{a:false,b:true},"2026-07-18":{a:false,b:true},"2026-07-19":{a:true,b:false},"2026-07-20":{a:true,b:false},"2026-07-21":{a:true,b:false},"2026-07-22":{a:true,b:false},"2026-07-23":{a:true,b:false},"2026-07-24":{a:true,b:false},"2026-07-25":{a:true,b:false},"2026-07-26":{a:false,b:true},"2026-07-27":{a:false,b:true},"2026-07-28":{a:false,b:true},"2026-07-29":{a:false,b:true},"2026-07-30":{a:false,b:true},"2026-07-31":{a:false,b:true},"2026-08-01":{a:false,b:true},"2026-08-02":{a:true,b:false},"2026-08-03":{a:true,b:false},"2026-08-04":{a:true,b:false},"2026-08-05":{a:true,b:false},"2026-08-06":{a:true,b:false},"2026-08-07":{a:true,b:false},"2026-08-08":{a:true,b:false},"2026-08-09":{a:false,b:true},"2026-08-10":{a:false,b:true},"2026-08-11":{a:false,b:true},"2026-08-12":{a:false,b:true},"2026-08-13":{a:false,b:true,n:"Lecture RFS 1/2"},"2026-08-14":{a:false,b:true},"2026-08-15":{a:false,b:true},"2026-08-16":{a:true,b:false},"2026-08-17":{a:true,b:false},"2026-08-18":{a:true,b:false},"2026-08-19":{a:true,b:false},"2026-08-20":{a:true,b:false,n:"Lecture RFS 2/2"},"2026-08-21":{a:true,b:false,n:"Fête Départ"},"2026-08-22":{a:true,b:false},"2026-08-23":{a:false,b:true},"2026-08-24":{a:false,b:true},"2026-08-25":{a:false,b:true},"2026-08-26":{a:false,b:true,n:"Gorée"},"2026-08-27":{a:false,b:true,n:"Warang"}};
const FAT_PLAN={"2026-03-12":{a:false,b:true,n:"Arrivée"},"2026-03-13":{a:false,b:true,n:"Départ vers RHO"},"2026-03-14":{a:true,b:true,n:"RHO"},"2026-03-15":{a:true,b:true,n:"Rando"},"2026-03-16":{a:true,b:true,n:"Rando"},"2026-03-17":{a:true,b:true,n:"Rando"},"2026-03-18":{a:true,b:true,n:"Arrivée FATICK"},"2026-03-19":{a:true,b:false},"2026-03-20":{a:true,b:false},"2026-03-21":{a:true,b:false},"2026-03-22":{a:false,b:true},"2026-03-23":{a:false,b:true},"2026-03-24":{a:false,b:true},"2026-03-25":{a:false,b:true},"2026-03-26":{a:false,b:true},"2026-03-27":{a:false,b:true},"2026-03-28":{a:false,b:true},"2026-03-29":{a:true,b:false},"2026-03-30":{a:true,b:false},"2026-03-31":{a:true,b:false},"2026-04-01":{a:true,b:false},"2026-04-02":{a:true,b:false},"2026-04-03":{a:true,b:false},"2026-04-04":{a:true,b:false},"2026-04-05":{a:false,b:true},"2026-04-06":{a:false,b:true},"2026-04-07":{a:false,b:true},"2026-04-08":{a:false,b:true},"2026-04-09":{a:false,b:true},"2026-04-10":{a:false,b:true},"2026-04-11":{a:false,b:true},"2026-04-12":{a:true,b:false},"2026-04-13":{a:true,b:false},"2026-04-14":{a:true,b:false},"2026-04-15":{a:true,b:false},"2026-04-16":{a:true,b:false},"2026-04-17":{a:true,b:false},"2026-04-18":{a:true,b:false},"2026-04-19":{a:false,b:true},"2026-04-20":{a:false,b:true},"2026-04-21":{a:false,b:true},"2026-04-22":{a:false,b:true},"2026-04-23":{a:false,b:true},"2026-04-24":{a:false,b:true},"2026-04-25":{a:false,b:true},"2026-04-26":{a:true,b:false},"2026-04-27":{a:true,b:false},"2026-04-28":{a:true,b:false},"2026-04-29":{a:true,b:false},"2026-04-30":{a:true,b:false},"2026-05-01":{a:true,b:false},"2026-05-02":{a:true,b:false},"2026-05-03":{a:false,b:true},"2026-05-04":{a:false,b:true},"2026-05-05":{a:false,b:true},"2026-05-06":{a:false,b:true},"2026-05-07":{a:false,b:true},"2026-05-08":{a:false,b:true},"2026-05-09":{a:false,b:true},"2026-05-10":{a:true,b:false,n:"Lecture RMS 1/2"},"2026-05-11":{a:true,b:false},"2026-05-12":{a:true,b:false},"2026-05-13":{a:true,b:false},"2026-05-14":{a:true,b:false},"2026-05-15":{a:true,b:false},"2026-05-16":{a:true,b:false},"2026-05-17":{a:false,b:true},"2026-05-18":{a:false,b:true},"2026-05-19":{a:false,b:true,n:"Lecture RMS 2/2"},"2026-05-20":{a:false,b:true},"2026-05-21":{a:false,b:true},"2026-05-22":{a:false,b:true},"2026-05-23":{a:false,b:true},"2026-05-24":{a:true,b:false},"2026-05-25":{a:true,b:false},"2026-05-26":{a:true,b:false},"2026-05-27":{a:true,b:false},"2026-05-28":{a:true,b:false},"2026-05-29":{a:true,b:false},"2026-05-30":{a:true,b:false},"2026-05-31":{a:false,b:true},"2026-06-01":{a:false,b:true},"2026-06-02":{a:false,b:true},"2026-06-03":{a:false,b:true},"2026-06-04":{a:false,b:true},"2026-06-05":{a:false,b:true},"2026-06-06":{a:false,b:true},"2026-06-07":{a:true,b:false},"2026-06-08":{a:true,b:false},"2026-06-09":{a:true,b:false},"2026-06-10":{a:true,b:false},"2026-06-11":{a:true,b:false},"2026-06-12":{a:true,b:false},"2026-06-13":{a:true,b:false},"2026-06-14":{a:false,b:true},"2026-06-15":{a:false,b:true},"2026-06-16":{a:false,b:true},"2026-06-17":{a:false,b:true},"2026-06-18":{a:false,b:true},"2026-06-19":{a:false,b:true},"2026-06-20":{a:false,b:true},"2026-06-21":{a:true,b:false},"2026-06-22":{a:true,b:false},"2026-06-23":{a:true,b:false},"2026-06-24":{a:true,b:false},"2026-06-25":{a:true,b:false},"2026-06-26":{a:true,b:false},"2026-06-27":{a:true,b:false},"2026-06-28":{a:false,b:true},"2026-06-29":{a:false,b:true},"2026-06-30":{a:false,b:true},"2026-07-01":{a:false,b:true},"2026-07-02":{a:false,b:true},"2026-07-03":{a:false,b:true},"2026-07-04":{a:false,b:true},"2026-07-05":{a:true,b:false},"2026-07-06":{a:true,b:false},"2026-07-07":{a:true,b:false},"2026-07-08":{a:true,b:false},"2026-07-09":{a:true,b:false},"2026-07-10":{a:true,b:false},"2026-07-11":{a:true,b:false},"2026-07-12":{a:false,b:true},"2026-07-13":{a:false,b:true},"2026-07-14":{a:false,b:true},"2026-07-15":{a:false,b:true},"2026-07-16":{a:false,b:true},"2026-07-17":{a:false,b:true},"2026-07-18":{a:false,b:true},"2026-07-19":{a:true,b:false},"2026-07-20":{a:true,b:false},"2026-07-21":{a:true,b:false},"2026-07-22":{a:true,b:false},"2026-07-23":{a:true,b:false},"2026-07-24":{a:true,b:false},"2026-07-25":{a:true,b:false},"2026-07-26":{a:false,b:true,n:"Lecture RFS 1/2"},"2026-07-27":{a:false,b:true},"2026-07-28":{a:false,b:true},"2026-07-29":{a:false,b:true},"2026-07-30":{a:false,b:true},"2026-07-31":{a:false,b:true},"2026-08-01":{a:false,b:true,n:"Lecture RFS 2/2"},"2026-08-02":{a:true,b:false},"2026-08-03":{a:true,b:false},"2026-08-04":{a:true,b:false,n:"Fête de départ"},"2026-08-05":{a:true,b:false},"2026-08-06":{a:true,b:false},"2026-08-07":{a:true,b:false},"2026-08-08":{a:true,b:false},"2026-08-09":{a:false,b:true,n:"Gorée"},"2026-08-10":{a:false,b:true,n:"Warang"},"2026-08-11":{a:false,b:true,n:"Départ"},"2026-08-12":{a:false,b:true}};
const localDay=(dt=new Date())=>{const x=new Date(dt.getTime()-dt.getTimezoneOffset()*60000);return x.toISOString().slice(0,10);};
const today=localDay();
const PDSR_LOGO="data:image/png;base64,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";
const DEM_TYPES=[{k:"conges",l:"Congés / absence"},{k:"acompte",l:"Acompte sur salaire"},{k:"rdv",l:"Rendez-vous"}];
const DEM_ST={en_attente:{l:"En attente",bg:"#FFF8E1",c:"#B8860B"},acceptee:{l:"Acceptée",bg:"#E8F5E9",c:"#2E7D32"},refusee:{l:"Refusée",bg:"#FFEBEE",c:"#C62828"},annulee:{l:"Annulée",bg:"#F5F5F5",c:"#8B7050"}};
const demTitre=(d)=>{if(d.type==="conges")return"Congés du "+(d.dateDebut?fmt(d.dateDebut):"?")+" au "+(d.dateFin?fmt(d.dateFin):"?");if(d.type==="acompte")return"Acompte de "+(d.montant||"?")+" €";return"Rendez-vous : "+(d.objet||"sans objet");};
const demJours=(a,b)=>{if(!a||!b)return 0;const d1=new Date(a+"T12:00:00"),d2=new Date(b+"T12:00:00");if(isNaN(d1)||isNaN(d2))return 0;return Math.max(0,Math.round((d2-d1)/86400000)+1);};
let _jspdfPromise=null;
function loadJsPDF(){
  if(window.jspdf&&window.jspdf.jsPDF)return Promise.resolve(window.jspdf.jsPDF);
  if(_jspdfPromise)return _jspdfPromise;
  const urls=["https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js","https://unpkg.com/jspdf@2.5.1/dist/jspdf.umd.min.js","https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"];
  _jspdfPromise=new Promise((resolve,reject)=>{const tryNext=(i)=>{if(i>=urls.length){_jspdfPromise=null;reject(new Error("Librairie PDF inaccessible (vérifier la connexion)"));return;}const s=document.createElement("script");s.src=urls[i];s.onload=()=>{(window.jspdf&&window.jspdf.jsPDF)?resolve(window.jspdf.jsPDF):tryNext(i+1);};s.onerror=()=>{s.remove();tryNext(i+1);};document.head.appendChild(s);};tryNext(0);});
  return _jspdfPromise;
}
const STAGE_ST=["Prévu","En cours","Terminé","Interrompu"];
const STAGE_STC={"Prévu":{bg:"#E3F2FD",c:"#1565C0"},"En cours":{bg:"#FFF8E1",c:"#B8860B"},"Terminé":{bg:"#E8F5E9",c:"#2E7D32"},"Interrompu":{bg:"#FFEBEE",c:"#C62828"}};
const APP_CRIT=[{k:"assiduite",l:"Assiduité et ponctualité"},{k:"comportement",l:"Comportement et savoir-être"},{k:"technique",l:"Acquisition des gestes professionnels"},{k:"autonomie",l:"Autonomie"},{k:"integration",l:"Intégration dans l'équipe"}];
const APP_NIV=["Insuffisant","En progression","Satisfaisant","Très satisfaisant"];
function StagesPanel({sujet,user,users,onUpdate,etabConfig}){
  const stages=(sujet.stages)||[];
  const[form,setForm]=useState(false);
  const[f,setF]=useState({intitule:"",structure:"",lieu:"",tuteur:"",telTuteur:"",referent:"",dateDebut:"",dateFin:"",heures:""});
  const educs=(users||[]).filter(u=>u.role==="educateur"||u.role==="coordinateur_site");
  const save=(arr)=>onUpdate(sujet.id,"stages",arr);
  const add=()=>{if(!f.intitule.trim()||!f.structure.trim()){alert("L'intitulé et la structure d'accueil sont obligatoires.");return;}save([...stages,{id:Date.now(),...f,statut:"Prévu",appreciations:[],bilan:"",creePar:(user&&user.name)||"",creeLe:isoToday()}]);setF({intitule:"",structure:"",lieu:"",tuteur:"",telTuteur:"",referent:"",dateDebut:"",dateFin:"",heures:""});setForm(false);};
  const patch=(id,p)=>save(stages.map(s=>s.id===id?{...s,...p}:s));
  const encadrant=user&&(user.role==="chef_service"||user.role==="directeur");
  const del=(id)=>{if(!encadrant){alert("La suppression d'un stage est réservée au chef de service et au directeur.");return;}if(confirm("Supprimer ce stage et toutes ses appréciations ?"))save(stages.filter(s=>s.id!==id));};
  const addApp=(st)=>{const a={id:Date.now(),date:isoToday(),par:(user&&user.name)||"",notes:{},texte:""};patch(st.id,{appreciations:[...((st.appreciations)||[]),a]});};
  const patchApp=(st,aid,p)=>patch(st.id,{appreciations:((st.appreciations)||[]).map(a=>a.id===aid?{...a,...p}:a)});
  const delApp=(st,aid)=>{const a=((st.appreciations)||[]).find(x=>x.id===aid);if(!encadrant&&a&&a.par&&a.par!==(user&&user.name)){alert("Vous ne pouvez supprimer que vos propres appréciations.");return;}if(confirm("Supprimer cette appréciation ?"))patch(st.id,{appreciations:((st.appreciations)||[]).filter(x=>x.id!==aid)});};
  const F=(k,l,type)=>(<div><label style={{...S.lbl}}>{l}</label><input type={type||"text"} style={{...S.inp}} value={f[k]} onChange={e=>setF(p=>({...p,[k]:e.target.value}))}/></div>);
  return(<div>
    {!form&&<button onClick={()=>setForm(true)} style={{...S.btnP,width:"100%",justifyContent:"center",marginBottom:12}}><Plus size={14}/>Ajouter un stage</button>}
    {form&&<div style={{...S.card,background:C.sableLight,marginBottom:12}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr",gap:8}}>{F("intitule","Intitulé du stage")}{F("structure","Structure d'accueil")}{F("lieu","Lieu / adresse")}</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:8}}>{F("tuteur","Tuteur sur site")}{F("telTuteur","Téléphone tuteur","tel")}</div>
      <div style={{marginTop:8}}><label style={{...S.lbl}}>Référent éducatif</label><select style={{...S.inp}} value={f.referent} onChange={e=>setF(p=>({...p,referent:e.target.value}))}><option value="">--</option>{educs.map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginTop:8}}>{F("dateDebut","Début","date")}{F("dateFin","Fin","date")}{F("heures","Heures/sem.","number")}</div>
      <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}><button onClick={add} style={{...S.btnP,flex:1,justifyContent:"center"}}><Check size={14}/>Enregistrer</button><button onClick={()=>setForm(false)} style={{...S.btnO,flex:1,justifyContent:"center"}}>Annuler</button></div>
    </div>}
    {stages.length===0&&<div style={{...S.card,fontSize:12,color:C.light}}>Aucun stage enregistré.</div>}
    {stages.slice().sort((a,b)=>String(b.dateDebut||"").localeCompare(String(a.dateDebut||""))).map(st=>{const sc=STAGE_STC[st.statut]||STAGE_STC["Prévu"];const nb=demJours(st.dateDebut,st.dateFin);return(<div key={st.id} style={{...S.card,marginBottom:10,borderLeft:"4px solid "+sc.c}}>
      <div style={{display:"flex",justifyContent:"space-between",gap:8,alignItems:"flex-start",flexWrap:"wrap"}}>
        <div style={{flex:1,minWidth:150}}><div style={{fontSize:14,fontWeight:800,color:C.dark}}>{st.intitule}</div><div style={{fontSize:11.5,color:C.mid}}>{st.structure}{st.lieu?" · "+st.lieu:""}</div></div>
        <select value={st.statut} onChange={e=>patch(st.id,{statut:e.target.value})} style={{...S.inp,width:"auto",fontSize:12,padding:"4px 8px",background:sc.bg,color:sc.c,fontWeight:700,border:"1px solid "+sc.c}}>{STAGE_ST.map(x=><option key={x}>{x}</option>)}</select>
        <button onClick={async()=>{try{await attestationStagePDF(st,sujet,etabConfig);}catch(err){alert("PDF impossible : "+(err&&err.message?err.message:err));}}} title="Attestation de stage" style={{...S.btnO,fontSize:11.5,padding:"5px 11px",minHeight:0}}><Download size={12}/>Attestation</button>
        {encadrant&&<button onClick={()=>del(st.id)} title="Supprimer le stage" style={{background:"none",border:"none",color:"#C62828",cursor:"pointer",fontWeight:800,fontSize:15}}>✕</button>}
      </div>
      <div style={{fontSize:11.5,color:C.mid,marginTop:6,lineHeight:1.6}}>
        {st.dateDebut||st.dateFin?<div>Du {st.dateDebut?fmt(st.dateDebut):"?"} au {st.dateFin?fmt(st.dateFin):"?"}{nb?" ("+nb+" j)":""}{st.heures?" · "+st.heures+" h/sem.":""}</div>:null}
        {st.tuteur&&<div>Tuteur : {st.tuteur}{st.telTuteur?" — "+st.telTuteur:""}</div>}
        {st.referent&&<div>Référent éducatif : {st.referent}</div>}
      </div>
      <div style={{marginTop:10,paddingTop:10,borderTop:"1px solid "+C.border}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6,flexWrap:"wrap"}}><div style={{fontSize:12,fontWeight:800,color:C.dark}}>Appréciations ({((st.appreciations)||[]).length})</div><button onClick={()=>addApp(st)} style={{...S.btnO,fontSize:12,padding:"4px 11px"}}><Plus size={12}/>Ajouter</button></div>
        {((st.appreciations)||[]).map(a=>(<div key={a.id} style={{padding:"9px 10px",background:"#f8f9fa",borderRadius:8,marginBottom:6}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6,flexWrap:"wrap"}}><div style={{fontSize:12,fontWeight:700,color:C.light}}>{fmt(a.date)} — {a.par}</div>{(encadrant||!a.par||a.par===(user&&user.name))&&<button onClick={()=>delApp(st,a.id)} style={{background:"none",border:"none",color:"#C62828",cursor:"pointer",fontWeight:800,fontSize:13}}>✕</button>}</div>
          {APP_CRIT.map(cr=>(<div key={cr.k} style={{marginBottom:5}}>
            <div style={{fontSize:12,fontWeight:700,color:C.mid,marginBottom:3}}>{cr.l}</div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>{APP_NIV.map(n=><button key={n} onClick={()=>patchApp(st,a.id,{notes:{...(a.notes||{}),[cr.k]:n}})} style={{padding:"3px 9px",borderRadius:14,border:"1.5px solid "+((a.notes||{})[cr.k]===n?C.gold:C.border),background:(a.notes||{})[cr.k]===n?C.gold:C.white,color:(a.notes||{})[cr.k]===n?C.white:C.mid,fontSize:11.5,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{n}</button>)}</div>
          </div>))}
          <textarea key={a.id+"t"} defaultValue={a.texte||""} onBlur={e=>patchApp(st,a.id,{texte:e.target.value})} placeholder="Commentaire du tuteur ou de l'éducateur" style={{...S.inp,minHeight:52,resize:"vertical",marginTop:4}}/>
        </div>))}
      </div>
      <div style={{marginTop:8}}><label style={{...S.lbl}}>Bilan de fin de stage</label><textarea key={st.id+"b"} defaultValue={st.bilan||""} onBlur={e=>patch(st.id,{bilan:e.target.value})} style={{...S.inp,minHeight:60,resize:"vertical"}}/></div>
    </div>);})}
  </div>);
}
function pdfNew(jsPDF){return new jsPDF({unit:"mm",format:"a4"});}
function pdfEntete(doc,etab,titre,sousTitre){
  const M=16,W=210;
  doc.setFillColor(184,146,42);doc.rect(0,0,W,24,"F");
  try{doc.addImage(PDSR_LOGO,"PNG",W-M-17,3.5,17,16.4);}catch(e){}
  doc.setTextColor(255,255,255);doc.setFont("helvetica","bold");doc.setFontSize(14);
  doc.text((etab&&etab.raisonSociale)||"Association PDSR",M,11);
  doc.setFont("helvetica","normal");doc.setFontSize(8.5);
  doc.text(((etab&&etab.sousTitre)||"")+"   FINESS "+((etab&&etab.finess)||""),M,17.5);
  doc.setTextColor(40,40,40);doc.setFont("helvetica","bold");doc.setFontSize(13);doc.text(titre,M,34);
  if(sousTitre){doc.setFont("helvetica","normal");doc.setFontSize(9);doc.setTextColor(120,110,90);doc.text(sousTitre,M,40);}
  return sousTitre?47:41;
}
const horoFR=()=>{const d=new Date();const p2=x=>String(x).padStart(2,"0");return p2(d.getDate())+"/"+p2(d.getMonth()+1)+"/"+d.getFullYear()+" à "+p2(d.getHours())+"h"+p2(d.getMinutes());};
function pdfPied(doc,mention,tirage){
  const n=doc.getNumberOfPages();
  for(let i=1;i<=n;i++){doc.setPage(i);doc.setFontSize(7.5);doc.setTextColor(140,140,140);
    doc.text(mention+" — page "+i+"/"+n,16,286);
    if(tirage)doc.text(tirage,16,289.5);}
}
function pdfEcrivain(doc,state){
  const M=16,LW=210-2*M;
  return (txt,{size=10,bold=false,gap=5,color=[40,40,40]}={})=>{
    doc.setFont("helvetica",bold?"bold":"normal");doc.setFontSize(size);doc.setTextColor(color[0],color[1],color[2]);
    doc.splitTextToSize(String(txt==null?"":txt),LW).forEach(pt=>{if(state.y>272){doc.addPage();state.y=16;}doc.text(pt,M,state.y);state.y+=gap;});
  };
}
const APP_NOTE=(a)=>APP_CRIT.map(c=>c.l+" : "+(((a.notes)||{})[c.k]||"—")).join("   ·   ");
async function attestationStagePDF(st,sujet,etab){
  const jsPDF=await loadJsPDF();const doc=pdfNew(jsPDF);
  const s={y:pdfEntete(doc,etab,"ATTESTATION DE STAGE","Séjour de remobilisation — "+(sujet.site||""))};
  const L=pdfEcrivain(doc,s);
  const rule=()=>{if(s.y>272){doc.addPage();s.y=16;}doc.setDrawColor(200,180,120);doc.line(16,s.y,194,s.y);s.y+=5;};
  rule();
  L("Nous soussignés, "+((etab&&etab.raisonSociale)||"Association PDSR")+", attestons que :",{size:10,gap:8});
  L((sujet.prenom||"")+" "+(sujet.nom||""),{size:16,bold:true,gap:7});
  L("a effectué un stage au sein de la structure suivante :",{size:10,gap:8});
  L("Structure : "+(st.structure||"—"),{size:11,bold:true,gap:5.5});
  if(st.lieu)L("Lieu : "+st.lieu,{size:10,gap:5});
  L("Intitulé : "+(st.intitule||"—"),{size:10,gap:5});
  L("Période : du "+(st.dateDebut?fmt(normDate(st.dateDebut)):"—")+" au "+(st.dateFin?fmt(normDate(st.dateFin)):"—")+(st.heures?"   ("+st.heures+" h/semaine)":""),{size:10,gap:5});
  if(st.tuteur)L("Tuteur : "+st.tuteur,{size:10,gap:5});
  if(st.referent)L("Référent éducatif : "+st.referent,{size:10,gap:5});
  L("Statut : "+(st.statut||"—"),{size:10,gap:8});
  rule();
  const apps=(st.appreciations)||[];
  L("ÉVALUATION",{size:11,bold:true,gap:6,color:[150,110,20]});
  if(!apps.length)L("Aucune appréciation enregistrée.",{size:10,gap:6});
  apps.forEach(a=>{
    L(fmt(a.date)+" — "+(a.par||""),{size:9,bold:true,gap:4.4,color:[120,110,90]});
    L(APP_NOTE(a),{size:9,gap:4.6});
    if(a.texte)L(a.texte,{size:10,gap:6});
    s.y+=1;});
  rule();
  L("BILAN DE FIN DE STAGE",{size:11,bold:true,gap:6,color:[150,110,20]});
  L(st.bilan||"Bilan non renseigné.",{size:10.5,gap:7});
  s.y+=10;
  L("Fait à "+((etab&&etab.ville)||"Fatick")+", le "+fmt(isoToday()),{size:10,gap:14});
  L("Le référent éducatif :"+"                                     "+"Le chef de service :",{size:10,gap:16});
  pdfPied(doc,"Attestation de stage — "+((etab&&etab.raisonSociale)||"Association PDSR"),"Tirée le "+horoFR());
  doc.save("attestation_stage_"+String((sujet.prenom||"")+"_"+(sujet.nom||"")).replace(/[^A-Za-z0-9]+/g,"_")+".pdf");
}
async function projetPDF(projet,sujet,etab,user,numEdition){
  const jsPDF=await loadJsPDF();const doc=pdfNew(jsPDF);
  const isMaj=sujet.id>=100;const DOM=isMaj?PROJ_DOM_MAJEUR:PROJ_DOM_MINEUR;
  const S2={y:pdfEntete(doc,etab,"PROJET PERSONNALISÉ",isMaj?"Contrat jeune majeur — art. L.222-5 CASF":"Art. L.311-3 et L.311-4 CASF — articulation avec le projet pour l'enfant")};
  const L=pdfEcrivain(doc,S2);
  const rule=()=>{if(S2.y>272){doc.addPage();S2.y=16;}doc.setDrawColor(200,180,120);doc.line(16,S2.y,194,S2.y);S2.y+=5;};
  rule();
  L("Bénéficiaire : "+(sujet.prenom||"")+" "+(sujet.nom||""),{bold:true});
  L("Site : "+(sujet.site||"—")+"     Entrée : "+(sujet.dateDebut?fmt(normDate(sujet.dateDebut)):"—")+"     Sortie : "+(sujet.dateFin?fmt(normDate(sujet.dateFin)):"—"));
  L("Projet élaboré le "+(projet.dateElaboration?fmt(projet.dateElaboration):"—")+" par "+(projet.creePar||"—"));
  L("DIPC remis le "+(projet.dipcRemisLe?fmt(projet.dipcRemisLe):"non renseigné"),{gap:7});
  rule();
  const p=projet.participation||{};
  L("PARTICIPATION DU BÉNÉFICIAIRE",{size:11,bold:true,gap:6,color:[150,110,20]});
  L("Entretien de co-construction : "+(p.dateEntretien?fmt(p.dateEntretien):"non renseigné"),{size:9.5});
  L("Avis et attentes exprimés :",{size:9,bold:true,gap:4.2});
  L(p.refus?("Refus de participer. "+(p.motifRefus||"")):(p.avisJeune||"—"),{size:9.5,gap:5.5});
  if(!isMaj){L("Avis des titulaires de l'autorité parentale :",{size:9,bold:true,gap:4.2});L(p.avisTitulaires||"—",{size:9.5,gap:6});}
  rule();
  L("OBJECTIFS",{size:11,bold:true,gap:6,color:[150,110,20]});
  const objs=projet.objectifs||[];
  if(!objs.length)L("Aucun objectif défini.",{size:9.5,gap:6});
  DOM.forEach(d=>{const list=objs.filter(o=>o.domaine===d.k);if(!list.length)return;
    L(d.l,{size:9.5,bold:true,gap:5,color:[100,80,20]});
    list.forEach(o=>{
      L("• "+o.titre+"   ["+(o.statut||"En cours")+"]",{size:9.5,gap:4.5});
      if(o.moyens)L("   Moyens : "+o.moyens,{size:9,gap:4.2,color:[90,90,90]});
      if(o.indicateur)L("   Indicateur : "+o.indicateur,{size:9,gap:4.2,color:[90,90,90]});
      if(o.referent||o.echeance)L("   Référent : "+(o.referent||"—")+"   Échéance : "+(o.echeance?fmt(o.echeance):"—"),{size:9,gap:5,color:[90,90,90]});
    });S2.y+=1;});
  const sansDom=objs.filter(o=>!o.domaine);
  if(sansDom.length){L("Objectifs non rattachés",{size:9.5,bold:true,gap:5,color:[150,40,40]});sansDom.forEach(o=>L("• "+o.titre,{size:9.5,gap:4.5}));}
  rule();
  L("RÉVISIONS ET AVENANTS",{size:11,bold:true,gap:6,color:[150,110,20]});
  const revs=projet.revisions||[];
  if(!revs.length)L("Aucune révision.",{size:9.5,gap:6});
  revs.forEach(r=>{L(fmt(r.date)+" — "+(r.par||"")+(r.type?" · "+r.type:""),{size:9,bold:true,gap:4.2});L(r.note||"",{size:9.5,gap:5.5});});
  const obs=projet.observations||[];
  if(obs.length){rule();L("OBSERVATIONS DE TERRAIN ("+obs.length+")",{size:11,bold:true,gap:6,color:[150,110,20]});
    obs.slice().sort((a,b)=>String(a.date).localeCompare(String(b.date))).forEach(o=>{L(fmt(o.date)+" — "+(o.par||"")+((o.domaine&&DOM.find(d=>d.k===o.domaine))?" · "+DOM.find(d=>d.k===o.domaine).l:""),{size:8.5,bold:true,gap:4,color:[120,110,90]});L(o.texte||"",{size:9.5,gap:5});});}
  const stg=(sujet.stages)||[];
  if(stg.length){rule();L("STAGES ET MISES EN SITUATION ("+stg.length+")",{size:11,bold:true,gap:6,color:[150,110,20]});
    stg.forEach(st=>{
      L((st.intitule||"Stage")+" — "+(st.structure||""),{size:10,bold:true,gap:4.6});
      L("Du "+(st.dateDebut?fmt(normDate(st.dateDebut)):"—")+" au "+(st.dateFin?fmt(normDate(st.dateFin)):"—")+"   ·   "+(st.statut||"—")+(st.tuteur?"   ·   tuteur : "+st.tuteur:""),{size:9,gap:4.6,color:[110,105,90]});
      ((st.appreciations)||[]).forEach(a=>{L("   "+fmt(a.date)+" — "+APP_NOTE(a),{size:8.5,gap:4,color:[95,95,95]});if(a.texte)L("   "+a.texte,{size:9,gap:4.6});});
      if(st.bilan)L("   Bilan : "+st.bilan,{size:9.5,gap:5.5});
      S2.y+=1;});}
  if(projet.bilan&&(projet.bilan.date||projet.bilan.texte)){rule();L("BILAN DE FIN DE SÉJOUR",{size:11,bold:true,gap:6,color:[150,110,20]});L("Rédigé le "+(projet.bilan.date?fmt(projet.bilan.date):"—")+" par "+(projet.bilan.par||"—"),{size:9,gap:5});L(projet.bilan.texte||"—",{size:9.5,gap:6});}
  rule();
  L("VALIDATION",{size:11,bold:true,gap:6,color:[150,110,20]});
  L(projet.statut==="valide"?("Validé le "+fmt(projet.valideLe)+" par "+(projet.validePar||"—")):"Document non validé — version de travail",{size:10,bold:true,gap:6,color:projet.statut==="valide"?[46,125,50]:[198,40,40]});
  L("Cette édition (n° "+(numEdition||1)+") reflète l'état du dossier au "+horoFR()+". Une édition ultérieure pourra différer si le dossier évolue.",{size:8.5,gap:5,color:[120,110,90]});
  const edPrec=(projet.editions||[]).slice(0,-1);
  if(edPrec.length)L("Éditions précédentes : "+edPrec.map(e=>"n° "+e.n+" le "+fmt(e.date)+(e.par?" par "+e.par:"")).join("   ·   "),{size:8,gap:5,color:[140,132,115]});
  (projet.reouvertures||[]).forEach(r=>L("Rouvert le "+fmt(r.date)+" par "+(r.par||"—")+(r.motif?" — "+r.motif:""),{size:8.5,gap:4,color:[120,110,90]}));
  S2.y+=8;L("Le référent éducatif :"+"                              "+"Le chef de service :",{size:9.5,gap:16});
  L("Le directeur :",{size:9.5,gap:6});
  pdfPied(doc,"Document confidentiel — projet personnalisé","Édition n° "+(numEdition||1)+" — tirée le "+horoFR()+" par "+((user&&user.name)||"—"));
  doc.save("projet_"+String((sujet.prenom||"")+"_"+(sujet.nom||"")).replace(/[^A-Za-z0-9]+/g,"_")+".pdf");
}
async function registrePDF(lignes,etab){
  const jsPDF=await loadJsPDF();const doc=pdfNew(jsPDF);
  const S2={y:pdfEntete(doc,etab,"REGISTRE DES PERSONNES ACCUEILLIES","Art. L. 331-2 du code de l'action sociale et des familles")};
  const L=pdfEcrivain(doc,S2);
  L("Registre destiné à être coté et paraphé (art. R. 331-5 CASF). Tenu en permanence à la disposition des autorités judiciaires et administratives compétentes. Toute personne appelée par ses fonctions à en prendre connaissance est tenue au secret professionnel (art. 226-13 du code pénal).",{size:8.5,gap:4.2,color:[120,110,90]});
  S2.y+=3;
  const cols=[[16,10],[26,62],[88,26],[114,32],[146,32],[178,16]];
  const head=["N°","Nom et prénom","Site","Date d'entrée","Date de sortie","Feuillet"];
  const ligne=(vals,bold)=>{
    if(S2.y>268){doc.addPage();S2.y=20;}
    doc.setFont("helvetica",bold?"bold":"normal");doc.setFontSize(8.5);doc.setTextColor(40,40,40);
    vals.forEach((v,i)=>doc.text(String(v==null?"":v),cols[i][0],S2.y,{maxWidth:cols[i][1]-2}));
    S2.y+=4.2;doc.setDrawColor(bold?180:228,bold?160:222,bold?110:212);doc.line(16,S2.y-2.8,194,S2.y-2.8);
  };
  ligne(head,true);
  lignes.forEach((l,i)=>ligne([i+1,(l.nom||"")+" "+(l.prenom||""),l.site||"—",l.entree?fmt(l.entree):"—",l.sortie?fmt(l.sortie):"—",""]));
  S2.y+=6;
  L("Observations des autorités et agents chargés du contrôle (art. L. 331-3 CASF)",{size:9,bold:true,gap:6,color:[150,110,20]});
  for(let i=0;i<4;i++){if(S2.y>268){doc.addPage();S2.y=20;}doc.setDrawColor(215,205,185);doc.line(16,S2.y,194,S2.y);S2.y+=8;}
  S2.y+=4;
  L("Coté et paraphé le .............................. par ..............................",{size:9,gap:10});
  L("Le directeur :",{size:9,gap:6});
  pdfPied(doc,"Registre L. 331-2 CASF — document confidentiel","Tiré le "+horoFR());
  doc.save("registre_personnes_accueillies_"+isoToday()+".pdf");
}
async function demandesPDF(liste,etab,titre){
  const jsPDF=await loadJsPDF();const doc=pdfNew(jsPDF);
  const st={y:pdfEntete(doc,etab,"DEMANDES DU PERSONNEL",titre||("Édité le "+fmt(isoToday())))};
  const L=pdfEcrivain(doc,st);
  const rule=()=>{if(st.y>272){doc.addPage();st.y=16;}doc.setDrawColor(200,180,120);doc.line(16,st.y,194,st.y);st.y+=5;};
  rule();
  if(!liste.length)L("Aucune demande.",{size:10});
  liste.forEach(d=>{
    const stt=(DEM_ST[d.statut]||DEM_ST.en_attente).l;
    L((d._nom?d._nom+" — ":"")+demTitre(d),{size:10.5,bold:true,gap:4.6});
    L((DEM_TYPES.find(t=>t.k===d.type)||{}).l+" · déposée le "+fmt(d.creeLe)+" · "+stt,{size:8.5,gap:4.4,color:[120,110,90]});
    if(d.interlocuteur)L("Avec : "+d.interlocuteur,{size:9,gap:4.2});
    if(d.motif)L(d.motif,{size:9.5,gap:4.6});
    if(d.decisionPar)L(stt+" par "+d.decisionPar+" le "+fmt(d.decisionLe)+(d.decisionNote?" — "+d.decisionNote:""),{size:9,gap:5,color:[90,90,90]});
    st.y+=2;if(st.y>272){doc.addPage();st.y=16;}doc.setDrawColor(230,225,215);doc.line(16,st.y,194,st.y);st.y+=5;
  });
  pdfPied(doc,"Document confidentiel — dossier du personnel","Tiré le "+horoFR());
  doc.save("demandes_personnel_"+isoToday()+".pdf");
}
const ENT_ECHELLE=["Insuffisant","À développer","Maîtrisé","Expert"];
const ENT_PERIODES=["1er semestre","2e semestre"];
const ENT_GRILLE=[
 {s:"Bilan de la période écoulée",q:[
  {k:"b1",t:"Objectifs fixés lors du précédent entretien",type:"texte",obl:true},
  {k:"b2",t:"Degré d'atteinte de ces objectifs",type:"texte",obl:true},
  {k:"b3",t:"Faits marquants de la période",type:"texte",obl:false}]},
 {s:"Compétences professionnelles",q:[
  {k:"c1",t:"Posture éducative et distance professionnelle",type:"echelle",obl:true},
  {k:"c2",t:"Respect du cadre, des procédures et des consignes",type:"echelle",obl:true},
  {k:"c3",t:"Capacité d'observation et qualité des écrits professionnels",type:"echelle",obl:true},
  {k:"c4",t:"Gestion des situations de tension et de conflit",type:"echelle",obl:true},
  {k:"c5",t:"Travail en équipe, transmissions et relève",type:"echelle",obl:true},
  {k:"c6",t:"Relation avec les familles, l'ASE et les partenaires",type:"echelle",obl:true},
  {k:"c7",t:"Autonomie et prise d'initiative",type:"echelle",obl:true},
  {k:"c8",t:"Assiduité et ponctualité",type:"echelle",obl:true},
  {k:"c9",t:"Contribution au projet personnalisé des jeunes",type:"echelle",obl:true},
  {k:"c10",t:"Commentaire général sur les compétences",type:"texte",obl:false}]},
 {s:"Conditions et charge de travail",q:[
  {k:"d1",t:"Charge de travail jugée soutenable",type:"oui_non",obl:true},
  {k:"d2",t:"Difficultés matérielles ou organisationnelles signalées",type:"texte",obl:false},
  {k:"d3",t:"Équilibre vie professionnelle / vie personnelle évoqué",type:"case",obl:true},
  {k:"d4",t:"Situation de travail à l'étranger : conditions de vie sur site abordées",type:"case",obl:true}]},
 {s:"Formation et évolution",q:[
  {k:"f1",t:"Formations suivies sur la période",type:"texte",obl:true},
  {k:"f2",t:"Besoins de formation identifiés",type:"texte",obl:true},
  {k:"f3",t:"Souhaits d'évolution professionnelle",type:"texte",obl:false},
  {k:"f4",t:"Information sur les dispositifs de formation délivrée au salarié",type:"case",obl:true}]},
 {s:"Objectifs pour la période à venir",q:[
  {k:"o1",t:"Objectif 1",type:"texte",obl:true},
  {k:"o2",t:"Objectif 2",type:"texte",obl:false},
  {k:"o3",t:"Objectif 3",type:"texte",obl:false},
  {k:"o4",t:"Moyens mis à disposition par l'établissement",type:"texte",obl:true}]},
 {s:"Expression du salarié",q:[
  {k:"s1",t:"Observations du salarié",type:"texte",obl:true},
  {k:"s2",t:"Le salarié a pris connaissance du compte rendu",type:"case",obl:true},
  {k:"s3",t:"Désaccord exprimé (préciser le cas échéant)",type:"texte",obl:false}]}];
const entChamps=()=>ENT_GRILLE.reduce((a,s)=>a.concat(s.q),[]);
const entManquants=(e)=>entChamps().filter(q=>{if(!q.obl)return false;const v=(e.reponses||{})[q.k];if(q.type==="case")return v!==true;return !String(v==null?"":v).trim();}).map(q=>q.t);
const entLibelle=(e)=>(e.periode||"")+" "+(e.annee||"")+" — "+(e.date?fmt(e.date):"sans date");
async function entretienPDF(ent,salarie,etab){
  const jsPDF=await loadJsPDF();
  const doc=new jsPDF({unit:"mm",format:"a4"});
  const M=16,W=210,LW=W-2*M;let y=M;
  const line=(txt,{size=10,bold=false,gap=5,color=[40,40,40]}={})=>{doc.setFont("helvetica",bold?"bold":"normal");doc.setFontSize(size);doc.setTextColor(color[0],color[1],color[2]);const parts=doc.splitTextToSize(String(txt==null?"":txt),LW);parts.forEach(pt=>{if(y>272){doc.addPage();y=M;}doc.text(pt,M,y);y+=gap;});};
  const rule=()=>{if(y>272){doc.addPage();y=M;}doc.setDrawColor(200,180,120);doc.line(M,y,W-M,y);y+=5;};
  y=pdfEntete(doc,etab,"COMPTE RENDU D'ENTRETIEN ANNUEL D'ÉVALUATION","Périodicité interne : deux entretiens par an");
  rule();
  line("Salarié : "+(salarie.name||""),{bold:true});
  line("Fonction : "+(salarie.role||"")+"     Site : "+(salarie.site||"—"));
  line("Période évaluée : "+(ent.periode||"")+" "+(ent.annee||""));
  line("Date de l'entretien : "+(ent.date?fmt(ent.date):"—"));
  line("Conduit par : "+(ent.evaluateur||"—"),{gap:7});
  rule();
  ENT_GRILLE.forEach(sec=>{
    line(sec.s.toUpperCase(),{size:11,bold:true,gap:6,color:[150,110,20]});
    sec.q.forEach(q=>{
      const v=(ent.reponses||{})[q.k];
      let aff;
      if(q.type==="case")aff=v===true?"Oui":"Non";
      else if(q.type==="oui_non")aff=v||"—";
      else if(q.type==="echelle")aff=v||"—";
      else aff=String(v==null||v===""?"—":v);
      line(q.t,{size:9,bold:true,gap:4.2});
      line(aff,{size:9.5,gap:5.5});
    });
    y+=2;
  });
  rule();
  line("Signatures",{size:11,bold:true,gap:8,color:[150,110,20]});
  line("Le salarié :"+" ".repeat(40)+"Le chef de service :",{size:9.5,gap:16});
  line("Le directeur :",{size:9.5,gap:6});
  const n=doc.getNumberOfPages();
  for(let i=1;i<=n;i++){doc.setPage(i);doc.setFontSize(8);doc.setTextColor(140,140,140);doc.text("Document confidentiel — dossier du personnel — page "+i+"/"+n,M,289);}
  doc.save("entretien_"+String(salarie.name||"salarie").replace(/[^A-Za-z0-9]+/g,"_")+"_"+(ent.annee||"")+"_"+String(ent.periode||"").replace(/[^A-Za-z0-9]+/g,"_")+".pdf");
}

const ETAB_DEFAULT={raisonSociale:"Association PDSR",sousTitre:"Promotion Des Séjours De Remobilisation",finess:"930032727",directeur:"Laurent Marcille",adresse:"",ville:"",tel:"",email:"",sites:["Fatick","Djilass"]};
const PROJ_DOM_MINEUR=[{k:"dev",l:"Développement, santé physique et psychique"},{k:"fam",l:"Relations avec la famille et les tiers"},{k:"sco",l:"Scolarité et vie sociale"}];
const PROJ_DOM_MAJEUR=[{k:"ins",l:"Insertion professionnelle et formation"},{k:"log",l:"Logement, ressources et autonomie"},{k:"san",l:"Santé et accès aux droits"},{k:"soc",l:"Réseau social et familial"}];
const OBJ_STATUTS=["En cours","Atteint","Partiellement atteint","Abandonné"];
const PROJ_CFG_DEFAULT={delaiDipc:15,delaiObjectifs:30,delaiRevision:75,delaiBilan:150};
const projCfg=(ec)=>({...PROJ_CFG_DEFAULT,...((ec||{}).projet||{})});
const isoToday=()=>new Date().toISOString().slice(0,10);
const normDate=(s)=>{if(!s)return"";const t=String(s).trim();if(/^\d{4}-\d{2}-\d{2}$/.test(t))return t;const m=t.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);return m?(m[3]+"-"+m[2]+"-"+m[1]):"";};
const addDays=(iso,n)=>{const b=normDate(iso);if(!b)return"";const d=new Date(b+"T12:00:00");if(isNaN(d.getTime()))return"";d.setDate(d.getDate()+Number(n||0));return d.toISOString().slice(0,10);};
const echStatut=(due,fait)=>{if(fait)return{l:"Fait",bg:"#E8F5E9",c:"#2E7D32"};if(!due)return{l:"Non calculable",bg:"#F5F5F5",c:"#8B7050"};if(due<isoToday())return{l:"En retard",bg:"#FFEBEE",c:"#C62828"};return{l:"À venir",bg:"#FFF8E1",c:"#B8860B"};};
const projEcheances=(j,p,ec)=>{const cfg=projCfg(ec);const d0=(j&&j.dateDebut)||"";const objs=(p&&p.objectifs)||[];const revs=(p&&p.revisions)||[];return[{k:"dipc",l:"DIPC remis",due:addDays(d0,cfg.delaiDipc),fait:(p&&p.dipcRemisLe)||""},{k:"obj",l:"Objectifs posés",due:addDays(d0,cfg.delaiObjectifs),fait:objs.length?(objs[0].creeLe||(p&&p.dateElaboration)||isoToday()):""},{k:"rev",l:"Révision mi-séjour",due:addDays(d0,cfg.delaiRevision),fait:revs.length?revs[revs.length-1].date:""},{k:"bil",l:"Bilan de sortie",due:addDays(d0,cfg.delaiBilan),fait:(p&&p.bilan&&p.bilan.date)||""}];};
const XL_COLS=["Prénom","Nom","Site","Tél parent","Tél jeune","Email ASE","Date début","Date fin"];
const deacc=(s)=>String(s==null?"":s).normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[.,;:_\-'\u2019()]/g," ").replace(/\s+/g," ").trim().toLowerCase();
const XL_MAP={"prenom":"prenom","nom":"nom","site":"site","tel parent":"telParent1","tel du parent":"telParent1","telephone parent":"telParent1","tel jeune":"telJeune","tel du jeune":"telJeune","telephone jeune":"telJeune","email ase":"emailASE","mail ase":"emailASE","e mail ase":"emailASE","courriel ase":"emailASE","date debut":"dateDebut","debut":"dateDebut","date entree":"dateDebut","date fin":"dateFin","fin":"dateFin","date sortie":"dateFin"};
const xlDate=(v)=>{if(v==null||v==="")return"";if(v instanceof Date&&!isNaN(v.getTime()))return new Date(v.getTime()-v.getTimezoneOffset()*60000).toISOString().slice(0,10);if(typeof v==="number"&&isFinite(v)){const d=new Date(Date.UTC(1899,11,30)+Math.round(v)*86400000);return isNaN(d.getTime())?"":d.toISOString().slice(0,10);}return normDate(v);};
const xlSite=(v)=>{const t=deacc(v);if(t.indexOf("dji")===0||t.indexOf("djilass")>=0)return"Djilass";if(t.indexOf("fat")===0||t.indexOf("fatick")>=0)return"Fatick";return"";};
const xlTel=(v)=>{if(v==null||v==="")return"";let t=String(v).replace(/[^0-9+]/g,"");if(t&&!t.startsWith("+")&&t.length===9)t="0"+t;return t;};
const xlRow=(r)=>{const o={};Object.keys(r||{}).forEach(k=>{const t=XL_MAP[deacc(k)];if(t)o[t]=r[k];});return o;};
const xlParse=(rows)=>{const out=[];const rej=[];const siteKo=[];rows.forEach((raw,i)=>{const r=xlRow(raw);const prenom=String(r.prenom||"").trim();if(!prenom){if(Object.values(r).some(v=>String(v||"").trim()))rej.push(i+2);return;}const site=xlSite(r.site);if(!site)siteKo.push(i+2);out.push({prenom,nom:String(r.nom||"").trim(),site:site||"Fatick",referentA:"",referentB:"",referentC:"",referentD:"",statut:"actif",telParent1:xlTel(r.telParent1),telJeune:xlTel(r.telJeune),emailASE:String(r.emailASE||"").trim(),dateDebut:xlDate(r.dateDebut),dateFin:xlDate(r.dateFin)});});return{out,rej,siteKo};};
const xlBilan=(res,exist)=>{const m=[res.out.length+" fiche(s) lue(s)."];if(res.rej.length)m.push("Ignorées (prénom vide) : ligne(s) "+res.rej.join(", ")+".");if(res.siteKo.length)m.push("Site non reconnu, placé sur Fatick : ligne(s) "+res.siteKo.join(", ")+".");if(exist.length)m.push("Doublons possibles (déjà présents) : "+exist.join(", ")+".");return m.join("\n");};
async function xlModele(nomFichier,onglet){try{const XLSX=await loadXLSX();const ws=XLSX.utils.aoa_to_sheet([XL_COLS]);ws["!cols"]=XL_COLS.map(()=>({wch:18}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,onglet);XLSX.writeFile(wb,nomFichier);}catch(err){alert("Impossible de générer le modèle : "+(err&&err.message?err.message:err));}}
const WD=["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"];
function getWeekDates(){const n=new Date(),m=new Date(n);m.setDate(n.getDate()-((n.getDay()+6)%7));return Array.from({length:7},(_,i)=>{const d=new Date(m);d.setDate(m.getDate()+i);return localDay(d);});}
const WEEKDATES=getWeekDates();
const INIT_RAPPORTS=[{id:1,jeuneId:2,date:today,observation:"Andy a participé activement aux ateliers maraîchers ce matin. Bonne dynamique de groupe."},{id:2,jeuneId:1,date:today,observation:"Chaïna a montré une bonne implication lors de l'atelier couture. Elle progresse bien."},{id:3,jeuneId:10,date:"2026-05-12",observation:"Imane a bien participé au chantier solidaire. Bonne ambiance générale."}];
const INIT_PRESENCES=(()=>{const data=[];JEUNES.forEach(j=>{WEEKDATES.forEach((date,i)=>{data.push({id:`${j.id}-${date}`,jeuneId:j.id,date,statut:i===5&&j.id===2?"Absent":"Présent"});});});return data;})();
const INIT_EV=[{id:1,jeuneId:14,date:"2026-05-10",titre:"Altercation verbale",description:"Dispute avec un pair lors du repas du soir. Tensions apaisées après médiation.",gravite:"Léger"},{id:2,jeuneId:5,date:"2026-05-08",titre:"Refus d'activité",description:"Kais a refusé de participer à l'atelier du matin. Entretien mené par l'éducateur.",gravite:"Moyen"}];
const GC={Léger:{bg:"#E8F5E9",text:"#2E7D32",dot:"#4CAF50"},Moyen:{bg:"#FFF8E1",text:"#F57F17",dot:"#FFC107"},Grave:{bg:"#FFEBEE",text:"#C62828",dot:"#F44336"}};
const SC={Présent:{bg:"#E8F5E9",text:"#2E7D32",icon:"✓"},Absent:{bg:"#FFEBEE",text:"#C62828",icon:"✗"},Retard:{bg:"#FFF8E1",text:"#E65100",icon:"◷"}};
const age=dob=>{const d=new Date(dob),n=new Date();let a=n.getFullYear()-d.getFullYear();if(n.getMonth()<d.getMonth()||(n.getMonth()===d.getMonth()&&n.getDate()<d.getDate()))a--;return a;};
const fmt=s=>{if(!s)return"";const[y,m,d]=s.split("-");return`${d}/${m}/${y}`;};
const S={card:{background:C.white,borderRadius:12,padding:"16px 18px",boxShadow:C.shadowSm,border:`1px solid ${C.border}`,marginBottom:12,transition:"box-shadow 0.2s ease"},inp:{width:"100%",padding:"11px 14px",borderRadius:9,border:`1px solid ${C.border}`,fontSize:16,minHeight:44,fontFamily:"'Nunito',sans-serif",color:C.dark,background:C.white,outline:"none",boxSizing:"border-box",transition:"all 0.2s ease"},lbl:{display:"block",fontSize:11.5,fontWeight:700,color:C.mid,marginBottom:6,letterSpacing:"0.04em",textTransform:"uppercase"},btnP:{background:C.goldDark,color:"#fff",border:"none",borderRadius:9,padding:"12px 18px",minHeight:44,fontWeight:700,fontSize:14,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:8,fontFamily:"'Nunito',sans-serif",boxShadow:C.shadowGold,transition:"all 0.2s ease"},btnO:{background:C.white,color:C.goldDark,border:`1px solid ${C.gold}`,borderRadius:9,padding:"11px 18px",minHeight:44,fontWeight:700,fontSize:14,cursor:"pointer",display:"inline-flex",alignItems:"center",justifyContent:"center",gap:8,fontFamily:"'Nunito',sans-serif",transition:"background 0.15s ease"},btnS:{background:C.white,color:C.dark,border:`1px solid ${C.border}`,borderRadius:9,padding:"11px 18px",minHeight:44,fontWeight:600,fontSize:14,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:8,fontFamily:"'Nunito',sans-serif",transition:"all 0.2s ease"}};
const Tag=({bg,text,children})=><span style={{background:bg,color:text,borderRadius:20,padding:"4px 12px",fontSize:12,fontWeight:700,letterSpacing:"0.02em",display:"inline-flex",alignItems:"center"}}>{children}</span>;

function Login({onLogin,users}){
  const[email,setEmail]=useState(""),[ pw,setPw]=useState(""),[ err,setErr]=useState("");
  const handle=()=>{const byId={};USERS.forEach(x=>{byId[x.id]={...x};});(loadLS()?.users||[]).forEach(x=>{byId[x.id]={...(byId[x.id]||{}),...x};});(users||[]).forEach(x=>{byId[x.id]={...(byId[x.id]||{}),...x};});const list=Object.values(byId);const u=list.find(x=>(x.email===email||x.login===email)&&x.password===pw);if(u&&u.disabled){setErr("Compte désactivé. Contactez votre chef de service.")}else{u?(setErr(""),onLogin(u)):setErr("Identifiants incorrects")}};
  return(<div style={{minHeight:"100vh",background:"linear-gradient(160deg,"+C.dark+" 0%,#2A1500 40%,#4A2800 70%,"+C.goldDark+" 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,animation:"fadeIn 0.6s ease",flexWrap:"wrap"}}>
    <div style={{width:"100%",maxWidth:400}}>
      <div style={{textAlign:"center",marginBottom:36}}>
        <img src={LOGO} alt="PDSR" style={{width:90,height:90,borderRadius:18,objectFit:"contain",background:"rgba(255,255,255,0.92)",padding:6,marginBottom:14,boxShadow:"0 4px 20px rgba(0,0,0,0.3)"}}/>
        <h1 style={{color:C.white,fontSize:26,fontWeight:900,margin:0}}>PDSR</h1>
        <p style={{color:"rgba(255,255,255,0.55)",margin:"5px 0 0",fontSize:13}}>Séjours de remobilisation — Sénégal</p>
      </div>
      <div style={{...S.card,borderRadius:22,marginBottom:0,padding:"28px",boxShadow:"0 8px 40px rgba(0,0,0,0.15)"}}>
        <h2 style={{fontSize:18,fontWeight:800,color:C.dark,margin:"0 0 20px"}}>Connexion</h2>
        <div style={{marginBottom:14}}><label style={{...S.lbl}}>Identifiant</label><input style={{...S.inp}} type="text" value={email} onChange={e=>setEmail(e.target.value)} placeholder="email ou identifiant" onKeyDown={e=>e.key==="Enter"&&handle()}/></div>
        <div style={{marginBottom:6}}><label style={{...S.lbl}}>Mot de passe</label><input style={{...S.inp}} type="password" value={pw} onChange={e=>setPw(e.target.value)} placeholder="••••••••" onKeyDown={e=>e.key==="Enter"&&handle()}/></div>
        {err&&<p style={{color:"#C62828",fontSize:13,margin:"6px 0"}}>{err}</p>}
        <button style={{...S.btnP,width:"100%",justifyContent:"center",marginTop:18,padding:"13px"}} onClick={handle}>Connexion</button>
      </div>
      
    </div>
  </div>);
}

const NAV=[
{id:"dashboard",label:"Tableau de bord",icon:Home,g:"Suivi"},
{id:"jeunes",label:"Jeunes",icon:Users,g:"Suivi"},
{id:"majeurs",label:"Majeurs",icon:GraduationCap,g:"Suivi"},
{id:"recherche",label:"Recherche",icon:Search,g:"Suivi"},
{id:"rapports",label:"Rapports journaliers",icon:FileText,g:"Quotidien"},
{id:"evenements",label:"Événements",icon:AlertTriangle,g:"Quotidien"},
{id:"transmissions",label:"Transmissions",icon:Send,g:"Quotidien"},
{id:"projets",label:"Projets personnalisés",icon:Target,g:"Accompagnement"},
{id:"agenda",label:"Agenda / RDV",icon:Clock,g:"Accompagnement"},
{id:"planning",label:"Planning",icon:CalendarDays,g:"Accompagnement"},
{id:"rapport-hebdo",label:"Rapport hebdo",icon:BarChart2,g:"Pilotage"},
{id:"rapport-site",label:"Rapport de site",icon:ClipboardList,g:"Pilotage"},
{id:"export",label:"Export Excel",icon:Download,g:"Pilotage"},
{id:"pres-educ",label:"Présences éducateurs",icon:UserCheck,g:"Équipe"},
{id:"espace-rh",label:"Espace éducateur",icon:Briefcase,g:"Équipe"},
{id:"intendance",label:"Intendance",icon:Package,g:"Équipe"},
{id:"admin",label:"Administration",icon:Settings,g:"Système"}];
const NAV_GROUPES=["Suivi","Quotidien","Accompagnement","Pilotage","Équipe","Système"];

const GlobalFX=()=><style>{`
@keyframes pgIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
@keyframes cardIn{from{opacity:0;transform:translateY(14px) scale(.985)}to{opacity:1;transform:none}}
@keyframes pulseUrg{0%,100%{box-shadow:0 0 0 0 rgba(198,40,40,.55)}50%{box-shadow:0 0 0 7px rgba(198,40,40,0)}}
.pg-anim{animation:pgIn .32s cubic-bezier(.22,1,.36,1)}
.pg-anim>div>div{animation:cardIn .4s cubic-bezier(.22,1,.36,1) backwards}
.pg-anim>div>div:nth-child(2){animation-delay:.05s}.pg-anim>div>div:nth-child(3){animation-delay:.1s}.pg-anim>div>div:nth-child(4){animation-delay:.15s}.pg-anim>div>div:nth-child(5){animation-delay:.2s}
button{transition:transform .14s ease,filter .14s ease,box-shadow .14s ease!important}
button:hover{filter:brightness(1.07)}
button:active{transform:scale(.96)}
input,select,textarea{transition:border-color .15s ease,box-shadow .15s ease!important}
input:focus,select:focus,textarea:focus{outline:none!important;border-color:#C8963E!important;box-shadow:0 0 0 3px rgba(200,150,62,.18)!important}
.pulse-urg{animation:pulseUrg 1.6s infinite}
::-webkit-scrollbar{width:9px;height:9px}::-webkit-scrollbar-thumb{background:#c9b899;border-radius:8px}::-webkit-scrollbar-track{background:transparent}
`}</style>;
function Sidebar({page,onNav,user,onLogout,open,onClose}){
  return(<>
    {open&&<div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:40}}/>}
    <aside style={{position:"fixed",top:0,left:open?0:-270,width:260,height:"100vh",background:"linear-gradient(180deg,"+C.dark+" 0%,#2A1500 100%)",zIndex:50,transition:"left 0.3s cubic-bezier(0.4,0,0.2,1)",display:"flex",flexDirection:"column",overflowY:"auto",boxShadow:open?"8px 0 32px rgba(0,0,0,0.3)":"none"}}>
      <div style={{padding:"22px 18px 18px",borderBottom:"1px solid rgba(255,255,255,0.09)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap"}}>
          <div style={{display:"flex",alignItems:"center",gap:11,flexWrap:"wrap"}}>
            <img src={LOGO} alt="PDSR" style={{width:40,height:40,borderRadius:9,objectFit:"contain",background:"rgba(255,255,255,0.9)",padding:3}}/>
            <div><div style={{color:C.white,fontWeight:900,fontSize:14}}>PDSR</div><div style={{color:"rgba(255,255,255,0.45)",fontSize:11.5,letterSpacing:"0.06em",textTransform:"uppercase"}}>Sénégal</div></div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"rgba(255,255,255,0.4)",cursor:"pointer",padding:4,display:"flex",flexWrap:"wrap"}}><X size={17}/></button>
        </div>
      </div>
      <div style={{padding:"14px 18px",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <div style={{width:34,height:34,borderRadius:9,background:`${C.gold}33`,display:"flex",alignItems:"center",justifyContent:"center",color:C.gold,fontWeight:800,fontSize:12}}>{user.initials}</div>
          <div><div style={{color:C.white,fontWeight:700,fontSize:13,display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>{user.name}{user.isAdmin&&<span style={{background:C.gold,color:C.dark,fontSize:11,fontWeight:900,letterSpacing:"0.06em",padding:"2px 6px",borderRadius:5,textTransform:"uppercase"}}>Admin</span>}{user.role==="coordinateur_site"&&<span style={{background:"#6A1B9A",color:"#fff",fontSize:11,fontWeight:900,letterSpacing:"0.06em",padding:"2px 6px",borderRadius:5,textTransform:"uppercase"}}>Coordo</span>}{user.role==="chef_service"&&!user.isAdmin&&<span style={{background:"#1565C0",color:"#fff",fontSize:11,fontWeight:900,letterSpacing:"0.06em",padding:"2px 6px",borderRadius:5,textTransform:"uppercase"}}>Chef</span>}{user.role==="directeur"&&<span style={{background:"#2E7D32",color:"#fff",fontSize:11,fontWeight:900,letterSpacing:"0.06em",padding:"2px 6px",borderRadius:5,textTransform:"uppercase"}}>Dir</span>}</div><div style={{color:C.sable,fontSize:12,fontWeight:600}}>{user.isAdmin?"Administrateur":user.role==="directeur"?"Directeur":user.role==="chef_service"?"Chef de service":user.role==="coordinateur_site"?`Coordinateur · ${user.site}`:`Éducateur · ${user.site}`}</div></div>
        </div>
      </div>
      <nav style={{flex:1,padding:"10px 10px"}}>
        {(()=>{const vis=NAV.filter(n=>{if(n.id==="admin"&&user.role!=="directeur"&&user.role!=="chef_service")return false;if(n.id==="export"&&user.role!=="directeur"&&user.role!=="chef_service"&&user.role!=="coordinateur_site")return false;if(user.isEducMajeur&&(n.id==="jeunes"||n.id==="planning"))return false;if(n.id==="rapport-hebdo"&&user.role!=="chef_service"&&user.role!=="directeur")return false;if(n.id==="rapport-site"&&user.role!=="coordinateur_site"&&user.role!=="chef_service"&&user.role!=="directeur")return false;if((n.id==="intendance"||n.id==="pres-educ")&&user.role!=="coordinateur_site"&&user.role!=="chef_service"&&user.role!=="directeur")return false;if(user.role==="educateur"&&!user.isEducMajeur&&n.id==="majeurs")return false;return true;});
        return NAV_GROUPES.map(g=>{const items=vis.filter(n=>n.g===g);if(!items.length)return null;return(<div key={g} style={{marginBottom:14}}>
          <div style={{fontSize:11,fontWeight:800,letterSpacing:"0.10em",textTransform:"uppercase",color:"rgba(234,210,156,0.55)",padding:"0 14px 7px"}}>{g}</div>
          {items.map(item=>{const Icon=item.icon;const active=page===item.id||page.startsWith(item.id+"-");return(<button key={item.id} onClick={()=>{onNav(item.id);onClose();}} style={{width:"100%",display:"flex",alignItems:"center",gap:12,padding:"12px 14px",minHeight:46,borderRadius:9,border:"none",borderLeft:active?"3px solid "+C.gold:"3px solid transparent",cursor:"pointer",fontFamily:"inherit",fontWeight:active?800:600,fontSize:14,marginBottom:2,background:active?"rgba(234,210,156,0.13)":"transparent",color:active?"#F6E7BE":"rgba(255,255,255,0.86)",textAlign:"left",transition:"background 0.15s ease"}}><Icon size={18} style={{flexShrink:0,opacity:active?1:0.75}}/><span style={{flex:1}}>{item.label}</span>{active&&<ChevronRight size={14}/>}</button>);})}
        </div>);});})()}
      </nav>
      <div style={{padding:"10px 10px 22px"}}><button onClick={onLogout} style={{width:"100%",display:"flex",alignItems:"center",gap:11,padding:"10px 13px",borderRadius:9,border:"none",cursor:"pointer",fontFamily:"inherit",fontWeight:600,fontSize:13,background:"rgba(244,67,54,0.12)",color:"#EF5350",flexWrap:"wrap"}}><LogOut size={17}/>Déconnexion</button></div>
    </aside>
  </>);
}

function Topbar({title,onMenu,onBack,unread,onBell,onRefresh,refreshing,lastSync,theme,onTheme}){
  return(<header style={{position:"sticky",top:0,zIndex:30,background:"rgba(255,255,255,0.97)",backdropFilter:"blur(12px)",WebkitBackdropFilter:"blur(12px)",borderBottom:"1px solid "+C.border,padding:"0 18px",height:58,display:"flex",alignItems:"center",gap:12,boxShadow:"0 2px 16px rgba(0,0,0,0.05)",flexWrap:"wrap"}}>
    {onBack?<button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",display:"flex",alignItems:"center",gap:5,color:C.gold,fontWeight:700,fontSize:13,fontFamily:"inherit",padding:0,flexWrap:"wrap"}}><ChevronLeft size={19}/>Retour</button>
    :<button onClick={onMenu} style={{background:"none",border:"none",cursor:"pointer",color:C.dark,padding:4,display:"flex",flexWrap:"wrap"}}><Menu size={21}/></button>}
    <h1 style={{fontSize:16,fontWeight:800,color:C.dark,margin:0,flex:1}}>{title}</h1>
    {onTheme&&<button onClick={onTheme} title={theme==="dark"?"Mode jour":"Mode nuit"} style={{background:"none",border:"none",cursor:"pointer",color:C.dark,padding:6,display:"flex",flexWrap:"wrap"}}>{theme==="dark"?<Sun size={19}/>:<Moon size={19}/>}</button>}
    {onRefresh&&<button onClick={onRefresh} disabled={refreshing} title={lastSync?("Dernière synchro "+String(lastSync.getHours()).padStart(2,"0")+":"+String(lastSync.getMinutes()).padStart(2,"0")):"Actualiser"} style={{background:"none",border:"none",cursor:refreshing?"default":"pointer",color:refreshing?C.gold:C.dark,padding:6,display:"flex",animation:refreshing?"spin 0.9s linear infinite":"none",flexWrap:"wrap"}}><RefreshCw size={19}/></button>}
    {onBell&&<button onClick={onBell} style={{position:"relative",background:"none",border:"none",cursor:"pointer",color:C.dark,padding:6,display:"flex",flexWrap:"wrap"}}><Bell size={20}/>{unread>0&&<span style={{position:"absolute",top:0,right:2,background:"#C62828",color:"#fff",fontSize:12,fontWeight:900,minWidth:15,height:15,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",padding:"0 3px",flexWrap:"wrap"}}>{unread>99?"99+":unread}</span>}</button>}
  </header>);
}

function Dashboard({user,rapports,presences,evenements,onNav,setSel,setPage,jeunes,agenda,majeurs,intendance,suiviEduc,users,projets,etabConfig}){
  const isMajEduc=user.role==="educateur"&&user.isEducMajeur;
  const pool=isMajEduc?(majeurs||MAJEURS):(jeunes||JEUNES);
  const vj=(user.role==="educateur"||user.role==="coordinateur_site")?pool.filter(j=>user.site==="Tous"||j.site===user.site):pool;
  const todayP=presences.filter(p=>p.date===today);
  const encadre=user.role==="chef_service"||user.role==="directeur";
  const pool7=[...(jeunes||[]),...(majeurs||[])];
  const dans7=(d)=>{if(!d)return false;const j=(new Date(d+"T12:00:00")-new Date(today+"T12:00:00"))/86400000;return j>=0&&j<=7;};
  const alertes=(()=>{const a=[];
    const retards=[];
    pool7.forEach(pj=>{const pr=(projets||[]).find(x=>String(x.jeuneId)===String(pj.id));projEcheances(pj,pr,etabConfig).forEach(e=>{if(!e.fait&&e.due&&e.due<today)retards.push({q:pj,l:e.l,due:e.due});});});
    if(retards.length)a.push({ic:"alerte",ton:"danger",t:retards.length+(retards.length>1?" échéances dépassées":" échéance dépassée"),s:retards.slice(0,3).map(r=>(r.q.prenom||"")+" — "+r.l).join(" · "),nav:"projets"});
    const sansDate=pool7.filter(pj=>!normDate(pj.dateDebut));
    if(sansDate.length)a.push({ic:"alerte",ton:"danger",t:sansDate.length+(sansDate.length>1?" fiches sans date d'entrée":" fiche sans date d'entrée"),s:"Aucune échéance ne peut être calculée · "+sansDate.slice(0,3).map(x=>x.prenom||"?").join(", "),nav:"jeunes"});
    const sansProjet=pool7.filter(pj=>!(projets||[]).some(x=>String(x.jeuneId)===String(pj.id)));
    if(sansProjet.length)a.push({ic:"cible",ton:"warning",t:sansProjet.length+(sansProjet.length>1?" jeunes sans projet personnalisé":" jeune sans projet personnalisé"),s:sansProjet.slice(0,3).map(x=>x.prenom||"?").join(", "),nav:"projets"});
    const vides=(projets||[]).filter(pr=>((pr.objectifs||[]).length===0));
    if(vides.length)a.push({ic:"cible",ton:"warning",t:vides.length+(vides.length>1?" projets sans objectif":" projet sans objectif"),s:"À compléter avec le jeune",nav:"projets"});
    if(encadre){const av=(projets||[]).filter(pr=>pr.statut!=="valide"&&((pr.objectifs||[]).length>0));
      if(av.length)a.push({ic:"valider",ton:"warning",t:av.length+(av.length>1?" projets à valider":" projet à valider"),s:"En attente de la direction",nav:"projets"});}
    const rj=(agenda||[]).filter(r=>r&&r.date===today);
    if(rj.length)a.push({ic:"rdv",ton:"accent",t:rj.length+(rj.length>1?" rendez-vous aujourd'hui":" rendez-vous aujourd'hui"),s:rj.slice(0,2).map(r=>(r.heure||"")+" "+(r.description||r.type||"")).join(" · "),nav:"agenda"});
    const fins=[],debuts=[],aSuivre=[];
    pool7.forEach(pj=>((pj.stages)||[]).forEach(st=>{const q=(pj.prenom||"")+" — "+(st.structure||st.intitule||"stage");
      if(st.statut==="En cours"&&dans7(normDate(st.dateFin)))fins.push(q);
      else if(st.statut==="Prévu"&&dans7(normDate(st.dateDebut)))debuts.push(q);
      else if(st.statut==="En cours"&&!((st.appreciations)||[]).length)aSuivre.push(q);}));
    if(fins.length)a.push({ic:"stage",ton:"warning",t:fins.length+(fins.length>1?" stages se terminent":" stage se termine")+" sous 7 jours",s:fins.slice(0,2).join(" · ")+" — bilan à rédiger",nav:"jeunes"});
    if(debuts.length)a.push({ic:"stage",ton:"accent",t:debuts.length+(debuts.length>1?" stages démarrent":" stage démarre")+" sous 7 jours",s:debuts.slice(0,2).join(" · "),nav:"jeunes"});
    if(aSuivre.length)a.push({ic:"stage",ton:"accent",t:aSuivre.length+(aSuivre.length>1?" stages sans appréciation":" stage sans appréciation"),s:aSuivre.slice(0,2).join(" · "),nav:"jeunes"});
    if(encadre){const dem=(users||[]).reduce((n,u)=>n+((u.demandes)||[]).filter(d=>d.statut==="en_attente").length,0);
      if(dem)a.push({ic:"demande",ton:"warning",t:dem+(dem>1?" demandes du personnel":" demande du personnel"),s:"En attente de décision",nav:"espace-rh"});}
    return a;})();
  const TON={danger:{c:"#C62828",bg:"#FFEBEE"},warning:{c:"#B8860B",bg:"#FFF8E1"},accent:{c:"#1565C0",bg:"#E3F2FD"}};
  const ICO={alerte:AlertTriangle,valider:Check,rdv:Clock,stage:Briefcase,demande:FileText,cible:Target};
  const rdvVis=(agenda||[]).filter(r=>r&&r.date).slice().sort((a,b)=>String(a.date+(a.heure||"")).localeCompare(String(b.date+(b.heure||""))));
  const rdvJour=rdvVis.filter(r=>r.date===today).length;
  const prochainRdv=rdvVis.find(r=>r.date>=today)||null;
  const myPresents=todayP.filter(p=>p.statut==="Présent"&&vj.some(j=>j.id===p.jeuneId)).length;
  const myRapports=(rapports||[]).filter(r=>r.date===today&&vj.some(j=>j.id===r.jeuneId)).length;
  const myGraves=(evenements||[]).filter(e=>e.gravite==="Grave"&&vj.some(j=>j.id===e.jeuneId)).length;
  const jeunesNav=isMajEduc?"majeurs":"jeunes";
  return(<div style={{padding:"20px 16px",maxWidth:800,margin:"0 auto",animation:"fadeIn 0.4s ease"}}>
    <p style={{color:C.light,fontSize:12,margin:"0 0 4px",letterSpacing:"0.02em"}}>{new Date().toLocaleDateString("fr-FR",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</p>
    <h1 style={{fontSize:24,fontWeight:900,color:C.dark,margin:"0 0 22px",letterSpacing:"-0.01em"}}>Bonjour, {user.name.split(" ")[0]} 👋</h1>
      <div style={{...S.card,marginBottom:14}}>
        <div style={{fontSize:12,fontWeight:800,color:C.light,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:alertes.length?10:0}}>À traiter aujourd'hui</div>
        {alertes.length===0&&<div style={{display:"flex",gap:8,alignItems:"center",fontSize:13,color:C.mid,flexWrap:"wrap"}}><Check size={16} color="#2E7D32"/>Rien en attente. Tout est à jour.</div>}
        {alertes.map((a,i)=>{const t=TON[a.ton]||TON.accent;const Ic=ICO[a.ic]||AlertTriangle;return(
          <div key={i} onClick={()=>onNav(a.nav)} style={{display:"flex",gap:11,alignItems:"flex-start",padding:"11px 0",borderTop:i?"1px solid "+C.border:"none",cursor:"pointer",flexWrap:"wrap"}}>
            <div style={{width:32,height:32,borderRadius:9,background:t.bg,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Ic size={17} color={t.c}/></div>
            <div style={{flex:1,minWidth:120}}>
              <div style={{fontSize:14,fontWeight:700,color:C.dark}}>{a.t}</div>
              {a.s&&<div style={{fontSize:12,color:C.mid,marginTop:1}}>{a.s}</div>}
            </div>
            <ChevronRight size={16} color={C.light} style={{marginTop:8,flexShrink:0}}/>
          </div>);})}
      </div>
    {(user.id===2||user.id===3)&&(()=>{
      const its=(intendance||[]).filter(Boolean);
      const bAtt=its.filter(i=>i.type==="besoin"&&i.statut==="en_attente");
      const bVal=its.filter(i=>i.type==="besoin"&&i.statut==="validee");
      const rOuv=its.filter(i=>i.type==="reparation"&&i.statut!=="repare");
      const rUrg=rOuv.filter(i=>i.urgence==="Urgente");
      const se=(suiviEduc||[]).filter(Boolean);
      const pAuj=se.filter(e=>e.kind==="presence"&&e.date===today);
      const absAuj=pAuj.filter(e=>e.statut&&e.statut!=="Présent");
      const rapEducRec=se.filter(e=>e.kind==="rapport").sort((a,b)=>(b.date||"").localeCompare(a.date||"")).slice(0,3);
      return(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:22}}>
        <div onClick={()=>onNav("intendance")} style={{...S.card,marginBottom:0,cursor:"pointer",borderLeft:"4px solid "+C.orange}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}><Package size={18} color={C.orange}/><span style={{fontWeight:800,fontSize:13,color:C.dark,textTransform:"uppercase",letterSpacing:"0.04em"}}>Intendance</span></div>
          <div style={{fontSize:13,color:C.mid,lineHeight:1.7}}>
            <div><b style={{color:bAtt.length?"#E65100":C.mid}}>{bAtt.length}</b> demande{bAtt.length>1?"s":""} en attente{bVal.length?<span> · <b style={{color:"#1565C0"}}>{bVal.length}</b> validée{bVal.length>1?"s":""} à livrer</span>:null}</div>
            <div><b style={{color:rOuv.length?"#C62828":C.mid}}>{rOuv.length}</b> réparation{rOuv.length>1?"s":""} ouverte{rOuv.length>1?"s":""}{rUrg.length?<span style={{color:"#C62828",fontWeight:800}}> dont {rUrg.length} URGENTE{rUrg.length>1?"S":""}</span>:null}</div>
            {bAtt.slice(0,2).map(b=><div key={b.id} style={{fontSize:12,color:C.light}}>• {b.categorie} — {b.site} (liv. {b.dateLivraison})</div>)}
          </div>
        </div>
        <div onClick={()=>onNav("pres-educ")} style={{...S.card,marginBottom:0,cursor:"pointer",borderLeft:"4px solid "+C.accent}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}><Users size={18} color={C.accent}/><span style={{fontWeight:800,fontSize:13,color:C.dark,textTransform:"uppercase",letterSpacing:"0.04em"}}>Éducateurs</span></div>
          <div style={{fontSize:13,color:C.mid,lineHeight:1.7}}>
            {(()=>{const cnt=st=>pAuj.filter(e=>e.statut===st).length;const P=cnt("Présent"),R=cnt("Repos"),A=cnt("Absent"),M=cnt("Maladie"),Rt=cnt("Retard"),Cg=cnt("Congé");
              return(<div>
                <div>Aujourd'hui : <b style={{color:"#2E7D32"}}>{P}</b> présent{P>1?"s":""} · <b style={{color:"#6A1B9A"}}>{R}</b> repos · <b style={{color:"#C62828"}}>{A}</b> absent{A>1?"s":""} · <b style={{color:"#E65100"}}>{M}</b> maladie{(Rt>0||Cg>0)?<span> · {Rt>0?<span><b style={{color:"#F9A825"}}>{Rt}</b> retard{Rt>1?"s":""}</span>:null}{Rt>0&&Cg>0?" · ":""}{Cg>0?<span><b style={{color:"#1565C0"}}>{Cg}</b> congé{Cg>1?"s":""}</span>:null}</span>:null}</div>
                {pAuj.length===0&&<div style={{color:"#E65100",fontWeight:800}}>Registre non renseigné aujourd'hui</div>}
              </div>);})()}
            {absAuj.slice(0,3).map(a=><div key={a.id} style={{fontSize:12,color:C.light}}>• {a.educName} — {a.statut}{a.note?" ("+a.note+")":""} · {a.site}</div>)}
            {rapEducRec.length>0&&<div style={{marginTop:4,fontSize:12,color:C.light}}>Derniers rapports : {rapEducRec.map(r=>r.educName+" ("+r.type+")").join(", ")}</div>}
          </div>
        </div>
      </div>);
    })()}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:20}}>
      {[{l:isMajEduc?"Majeurs suivis":"Jeunes suivis",v:vj.length,i:"👥",c:C.gold,bg:C.goldLight,nav:jeunesNav},{l:rdvJour>0?"RDV aujourd'hui":"Prochain RDV",v:rdvJour>0?String(rdvJour):(prochainRdv?fmt(prochainRdv.date):"—"),i:"📅",c:"#1565C0",bg:"#E3F2FD",nav:"agenda"},{l:"Rapports ce jour",v:myRapports,i:"📝",c:C.orange,bg:C.orangeLight,nav:"rapports"},{l:"Incidents graves",v:myGraves,i:"⚠️",c:"#C62828",bg:"#FFEBEE",nav:"evenements"}].map((s,i)=>(
        <div key={i} onClick={()=>s.nav&&setPage(s.nav)} style={{...S.card,display:"flex",alignItems:"center",gap:14,padding:"16px",marginBottom:0,cursor:s.nav?"pointer":"default",borderLeft:"4px solid "+s.c,animation:"fadeIn 0.4s ease "+(i*0.08)+"s both",opacity:s.nav?1:0.7,flexWrap:"wrap"}}>
          <div style={{width:44,height:44,borderRadius:12,background:s.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0,flexWrap:"wrap"}}>{s.i}</div>
          <div><div style={{fontSize:26,fontWeight:900,color:s.c,lineHeight:1}}>{s.v}</div><div style={{fontSize:11.5,color:C.light,fontWeight:700,marginTop:3,letterSpacing:"0.03em",textTransform:"uppercase"}}>{s.l}</div></div>
        </div>
      ))}
    </div>
    {user.role!=="educateur"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
      {[{l:"Site Fatick",n:(jeunes||JEUNES).filter(j=>j.site==="Fatick").length,e:"🏖️",c:C.gold},{l:"Site Djilass",n:(jeunes||JEUNES).filter(j=>j.site==="Djilass").length,e:"🌿",c:C.orange}].map((s,i)=>(
        <div key={i} onClick={()=>setPage("jeunes")} style={{...S.card,borderTop:"none",borderLeft:"4px solid "+s.c,padding:"16px",marginBottom:0,cursor:"pointer",animation:"fadeIn 0.4s ease 0.3s both"}}>
          <div style={{fontSize:24,marginBottom:4}}>{s.e}</div><div style={{fontWeight:900,fontSize:20,color:C.dark}}>{s.n} jeunes</div><div style={{fontSize:12,color:C.light,fontWeight:700,marginTop:2}}>{s.l}</div>
        </div>
      ))}
    </div>}
    {user.role!=="educateur"&&<div onClick={()=>setPage("majeurs")} style={{...S.card,display:"flex",alignItems:"center",gap:14,padding:"16px",marginBottom:20,cursor:"pointer",borderLeft:"4px solid #6A1B9A",animation:"fadeIn 0.4s ease 0.38s both",flexWrap:"wrap"}}>
      <div style={{width:44,height:44,borderRadius:12,background:"#F3E5F5",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0,flexWrap:"wrap"}}>🎓</div>
      <div><div style={{fontSize:26,fontWeight:900,color:"#6A1B9A",lineHeight:1}}>{(majeurs||MAJEURS).length}</div><div style={{fontSize:11.5,color:C.light,fontWeight:700,marginTop:3,letterSpacing:"0.03em",textTransform:"uppercase"}}>Majeurs suivis</div></div>
    </div>}
    {agenda&&agenda.length>0&&(()=>{const allJ=[...(jeunes||JEUNES),...(majeurs||MAJEURS)];const upcoming=agenda.filter(a=>{if(a.date<new Date().toISOString().slice(0,10))return false;const j=allJ.find(x=>String(x.id)===String(a.jeuneId));if(user.role==="educateur"){if(user.isEducMajeur&&j&&j.id<100)return false;if(!user.isEducMajeur&&j&&j.id>=100)return false;if(user.site!=="Tous"&&j&&j.site!==user.site)return false;}return true;}).sort((a,b)=>a.date.localeCompare(b.date)||((a.heure||"").localeCompare(b.heure||"")));const typeColors={referent:C.gold,medical:"#2E7D32",administratif:C.primary,scolaire:C.orange,autre:C.mid};return(<div onClick={()=>setPage("agenda")} style={{...S.card,marginTop:16,cursor:"pointer",borderLeft:"4px solid "+C.gold}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12,flexWrap:"wrap"}}><span style={{fontSize:18}}>📅</span><span style={{fontWeight:800,fontSize:15,color:C.dark}}>Prochains RDV</span><span style={{marginLeft:"auto",fontSize:12,color:C.light,fontWeight:600}}>{upcoming.length} à venir</span></div>{upcoming.slice(0,5).map(a=>{const j=allJ.find(x=>String(x.id)===String(a.jeuneId));const nom=j?(j.prenom+" "+j.nom):(a.jeuneNom||"—");const tc=typeColors[a.type]||C.mid;const [y,m,d]=(a.date||"").split("-");return(<div key={a.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 6px",borderRadius:8,marginBottom:4,background:"#f8f9fa",flexWrap:"wrap"}}><div style={{minWidth:44,textAlign:"center"}}><div style={{fontSize:13,fontWeight:800,color:C.primary}}>{d}/{m}</div><div style={{fontSize:11.5,color:C.mid,fontWeight:600}}>{a.heure||""}</div></div><div style={{flex:1}}><div style={{fontSize:13,fontWeight:700,color:C.dark}}>{nom}</div><div style={{display:"inline-block",fontSize:11.5,fontWeight:700,color:"#fff",background:tc,borderRadius:4,padding:"1px 6px",marginTop:2}}>{a.type||"rdv"}</div></div></div>);})}{upcoming.length===0&&<div style={{fontSize:12,color:C.light,textAlign:"center",padding:10}}>Aucun RDV à venir</div>}</div>);})()}
    
  </div>);
}

function JeunesList({user,jeunes,presences,onSelect,onNav,onUpdateJeune}){
  const[q,setQ]=useState(""),[ site,setSite]=useState("Tous");
  const pool=(jeunes&&jeunes.length!==undefined)?jeunes:JEUNES;const vj=(user.role==="educateur"||user.role==="coordinateur_site")?pool.filter(j=>user.site==="Tous"||j.site===user.site):pool;
  const vis=vj.filter(j=>{const m=`${j.prenom} ${j.nom}`.toLowerCase().includes(q.toLowerCase());const s=site==="Tous"||j.site===site;return m&&s;});
  return(<div style={{padding:"18px 14px",maxWidth:800,margin:"0 auto"}}>
    <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
      <div style={{flex:1,position:"relative"}}><Search size={15} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:C.light}}/><input style={{...S.inp,paddingLeft:35}} placeholder="Rechercher..." value={q} onChange={e=>setQ(e.target.value)}/></div>
      {user.role!=="educateur"&&<select style={{...S.inp,width:"auto",paddingLeft:12}} value={site} onChange={e=>setSite(e.target.value)}><option>Tous</option><option>Fatick</option><option>Djilass</option></select>}
    </div>
    {vis.map(j=>{const tp=presences.filter(p=>p.date===today&&p.jeuneId===j.id)[0];return(<div key={j.id} style={{...S.card,cursor:"pointer",animation:"fadeIn 0.3s ease"}} onClick={()=>{onSelect(j);onNav("jeune-detail");}}>
      <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
        <div style={{width:46,height:46,borderRadius:14,background:"linear-gradient(135deg,"+C.sable+","+C.goldLight+")",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:16,color:C.goldDark,flexShrink:0,boxShadow:"0 2px 8px rgba(184,134,11,0.15)",flexWrap:"wrap"}}>{j.prenom[0]}{j.nom?j.nom[0]:""}</div>
        <div style={{flex:1}}>
          <div style={{fontWeight:800,fontSize:15,color:C.dark}}>{j.prenom} {j.nom}</div>
          <div style={{fontSize:12,color:C.light,marginTop:1}}>{j.referentA} / {j.referentB} · <span style={{color:C.gold,fontWeight:700}}>{j.site}</span></div>
        </div>
        {tp&&<div style={{fontSize:12,fontWeight:700,padding:"4px 10px",borderRadius:8,background:SC[tp.statut]?.bg||C.sable,color:SC[tp.statut]?.text||C.mid}}>{tp.statut}</div>}
        <ChevronRight size={16} color={C.light}/>
      </div>
    </div>);})}
    {vis.length===0&&<div style={{textAlign:"center",padding:48,color:C.light,fontSize:13,fontWeight:600,opacity:0.7}}>Aucun jeune trouvé</div>}
  </div>);
}

function JeuneDetail({jeune,rapports,presences,evenements,user,onAddR,onAddE,onCP,onUpdateJeune,users,projets,onUpdateProjets,etabConfig}){
  const[tab,setTab]=useState("fiche");const[saved,setSaved]=useState(false);
  const jr=(rapports||[]).filter(r=>r.jeuneId===jeune.id).sort((a,b)=>b.date.localeCompare(a.date));
  const jp=presences.filter(p=>p.jeuneId===jeune.id&&WEEKDATES.includes(p.date));
  const je=(evenements||[]).filter(e=>e.jeuneId===jeune.id).sort((a,b)=>b.date.localeCompare(a.date));
  const tabs=["fiche","rapports","projet","stages","incidents"];
  return(<div style={{padding:"14px",maxWidth:700,margin:"0 auto"}}>
    <div style={{...S.card,background:C.sable,border:"none",marginBottom:14}}>
      <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
        <div style={{width:52,height:52,borderRadius:15,background:C.white,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:20,color:C.gold,flexWrap:"wrap"}}>{jeune.prenom[0]}{jeune.nom?jeune.nom[0]:""}</div>
        <div><div style={{fontSize:20,fontWeight:900,color:C.dark}}>{jeune.prenom} {jeune.nom}</div><div style={{fontSize:12,color:C.mid,marginTop:2}}>Réf: {jeune.referentA} / {jeune.referentB}</div><div style={{fontSize:12,color:C.gold,fontWeight:700,marginTop:2}}>Site {jeune.site}</div></div>
      </div>
    </div>
    <div style={{display:"flex",gap:5,marginBottom:14,overflowX:"auto",paddingBottom:4,flexWrap:"wrap"}}>
      {tabs.map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:"7px 14px",borderRadius:20,border:`1.5px solid ${tab===t?C.gold:C.border}`,background:tab===t?C.gold:C.white,color:tab===t?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap",textTransform:"capitalize"}}>{t}</button>)}
    </div>
    {tab==="fiche"&&<div style={{...S.card,marginBottom:12,borderLeft:"4px solid "+(jeune.besoinTraitement==="oui"?"#C62828":jeune.besoinTraitement==="non"?"#2E7D32":"#E65100")}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,flexWrap:"wrap"}}>
        <div style={{fontWeight:800,fontSize:14,color:C.dark}}>💊 Traitement médical</div>
        {(user.role==="directeur"||user.role==="chef_service")?<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{[["oui","Oui"],["non","Non"]].map(([v,l])=><button key={v} onClick={()=>onUpdateJeune&&onUpdateJeune(jeune.id,"besoinTraitement",v)} style={{padding:"6px 18px",borderRadius:8,border:"1.5px solid "+(jeune.besoinTraitement===v?(v==="oui"?"#C62828":"#2E7D32"):C.border),background:jeune.besoinTraitement===v?(v==="oui"?"#C62828":"#2E7D32"):C.white,color:jeune.besoinTraitement===v?"#fff":C.mid,fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>{l}</button>)}</div>:<span style={{fontWeight:800,fontSize:13,color:jeune.besoinTraitement==="oui"?"#C62828":jeune.besoinTraitement==="non"?"#2E7D32":"#E65100"}}>{jeune.besoinTraitement==="oui"?"Oui":jeune.besoinTraitement==="non"?"Non":"À renseigner"}</span>}
      </div>
      {!jeune.besoinTraitement&&<div style={{marginTop:8,fontSize:12,color:"#C62828",fontWeight:700}}>⚠ À renseigner obligatoirement.</div>}
      {jeune.besoinTraitement==="oui"&&<div style={{marginTop:8,fontSize:12,color:C.mid}}>Le suivi d'administration sera demandé à chaque rapport journalier.</div>}
    </div>}
    {tab==="fiche"&&<div style={{...S.card}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>{[["Prénom",jeune.prenom],["Nom",jeune.nom],["Référent A",jeune.referentA],["Référent B",jeune.referentB],["Référent C",jeune.referentC||""],["Référent D",jeune.referentD||""],["Site",jeune.site],["Statut",jeune.statut]].map(([k,v])=><div key={k}><div style={{fontSize:11.5,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>{k}</div><div style={{fontWeight:700,color:C.dark,fontSize:13}}>{v}</div></div>)}</div></div>}
    {tab==="fiche"&&(user.role==="directeur"||user.role==="chef_service")&&<div style={{...S.card,marginTop:12}}><div style={{fontWeight:700,color:C.dark,fontSize:14,marginBottom:10}}>Dossier du jeune</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><div style={{fontSize:11.5,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>Référent A</div><select style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune.referentA||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,"referentA",e.target.value)}><option value="">--</option>{USERS.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div><div><div style={{fontSize:11.5,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>Référent B</div><select style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune.referentB||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,"referentB",e.target.value)}><option value="">--</option>{USERS.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div><div><div style={{fontSize:11.5,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>Référent C</div><select style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune.referentC||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,"referentC",e.target.value)}><option value="">--</option>{USERS.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div><div><div style={{fontSize:11.5,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>Référent D</div><select style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune.referentD||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,"referentD",e.target.value)}><option value="">--</option>{USERS.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>{[["emailASE","Email éduc. ASE"],["telASE","Tél. éduc. ASE"],["telParent1","Tél. parent/tuteur 1"],["telParent2","Tél. parent/tuteur 2"],["traitement","Traitement"],["notesDossier","Notes dossier"]].map(([field,label])=><div key={field} style={{gridColumn:field==="traitement"||field==="notesDossier"?"1/-1":"auto"}}><div style={{fontSize:11.5,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>{label}</div>{field==="notesDossier"||field==="traitement"?<textarea style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13,fontFamily:"inherit",minHeight:50,resize:"vertical"}} value={jeune[field]||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,field,e.target.value)}/>:<input style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune[field]||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,field,e.target.value)}/>}</div>)}</div><div style={{textAlign:"center",marginTop:16}}><button onClick={()=>{setSaved(true);setTimeout(()=>setSaved(false),2000);}} style={{padding:"10px 32px",background:saved?"#27ae60":"#2c6fbb",color:"#fff",border:"none",borderRadius:8,fontWeight:700,fontSize:14,cursor:"pointer",transition:"all 0.3s"}}>{saved?"✓ Enregistré !":"Enregistrer les modifications"}</button></div></div>}
    {tab==="rapports"&&<div>{jr.length===0?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun rapport</div>:jr.map(r=><div key={r.id} style={{...S.card}}><div style={{fontSize:12,color:C.gold,fontWeight:700,marginBottom:5}}>{fmt(r.date)}{r.author&&<span style={{fontWeight:400,fontSize:12,color:C.light,marginLeft:8}}>par {r.author}</span>}</div><p style={{margin:0,fontSize:13,color:C.dark,lineHeight:1.6}}>{r.observation}</p></div>)}<button style={{...S.btnP,marginTop:8}} onClick={()=>onAddR(jeune)}><Plus size={15}/>Nouveau rapport</button></div>}
    {tab==="présences"&&<div><div style={{display:"flex",gap:4,marginBottom:10,flexWrap:"wrap"}}>{WEEKDATES.map((date,i)=>{const p=jp.find(p2=>p2.date===date);const st=p?.statut||"Présent";const next={Présent:"Absent",Absent:"Retard",Retard:"Présent"};const sc2=SC[st]||SC.Présent;return(<div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:3}}><div style={{fontSize:12,fontWeight:700,color:C.light}}>{WD[i]}</div><button onClick={()=>onCP(jeune.id,date,next[st])} style={{width:"100%",aspectRatio:"1",borderRadius:7,background:sc2.bg,border:"none",cursor:"pointer",color:sc2.text,fontWeight:800,fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",flexWrap:"wrap"}}>{sc2.icon}</button></div>);})}
    </div></div>}
    {tab==="projet"&&<ProjetsPersonnalises user={user} jeunes={[jeune]} majeurs={[]} projets={projets} onUpdate={onUpdateProjets} etabConfig={etabConfig} users={users} fixedId={jeune.id}/>}
      {tab==="stages"&&<StagesPanel sujet={jeune} user={user} users={users} onUpdate={onUpdateJeune} etabConfig={etabConfig}/>}
      {tab==="incidents"&&<div>{je.length===0?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun incident</div>:je.map(e=>{const gc=GC[e.gravite]||GC["Léger"];return(<div key={e.id} style={{...S.card,borderLeft:`4px solid ${gc.dot}`,animation:"cardEnter 0.3s ease-out"}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:5,flexWrap:"wrap"}}><div style={{fontWeight:800,color:C.dark}}>{e.titre}</div><span style={{...{background:gc.bg,color:gc.text,borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:700}}}>{e.gravite}</span></div><p style={{margin:0,fontSize:12,color:C.mid}}>{e.description}</p><div style={{fontSize:11.5,color:C.light,marginTop:4}}>{fmt(e.date)}{e.author&&" - par "+e.author}</div></div>);})}
    <button style={{...S.btnO,marginTop:8}} onClick={()=>onAddE(jeune)}><Plus size={15}/>Déclarer événement</button></div>}
  </div>);
}

function Rapports({user,rapports,onSave,onDelete,onUpdate,onPatch,majeurs,jeunes}){
  const allPool=[...(jeunes||JEUNES),...(majeurs||MAJEURS)];const vj=(user.role==="educateur"||user.role==="coordinateur_site")?(user.isEducMajeur?allPool.filter(j=>(j.id>=100)&&(user.site==="Tous"||j.site===user.site)):allPool.filter(j=>(j.id<100)&&(user.site==="Tous"||j.site===user.site))):allPool;
  const[jid,setJid]=useState(vj[0]?.id||"");
  const[date,setDate]=useState(today);
  const[obs,setObs]=useState("");
  const[typeContact,setTypeContact]=useState("journee");
  const[traitAdmin,setTraitAdmin]=useState("");
  const[saved,setSaved]=useState(false);
  const[rPage,setRPage]=useState(0);
  const[editTC,setEditTC]=useState(null);
  const[siteF,setSiteF]=useState("Tous");
  const[editObs,setEditObs]=useState(null);
  const[editText,setEditText]=useState("");
  const[showHist,setShowHist]=useState(null);
  const isEncadrement=user.role==="chef_service"||user.role==="directeur"||user.role==="coordinateur_site";
  const canEditR=(r)=>{if(isEncadrement)return true;if(user.role!=="educateur")return false;if(r.authorId!=null&&user.id!=null)return r.authorId===user.id;return !!r.author&&r.author===user.name;};
  const saveEdit=(r)=>{const txt=editText.trim();if(!txt){alert("L'observation ne peut pas être vide.");return;}if(txt===r.observation){setEditObs(null);return;}const ts=new Date().toISOString();if(onPatch)onPatch(r.id,{observation:txt,modifieLe:ts,modifiePar:user.name||"?",historique:[...(r.historique||[]),{date:ts,par:user.name||"?",ancienTexte:r.observation}]});setEditObs(null);setEditText("");};
  const PER_PAGE=20;const MAX_DISPLAY=300;
  const existingCount=rapports.filter(r=>r.jeuneId===+jid&&r.date===date).length;
  const TYPE_CONTACT_LABELS={"journee":"Journée du jeune","rdv_parents":"RDV téléphonique avec les parents","rdv_exterieur":"RDV téléphonique contact extérieur"};
  const selJeune=allPool.find(j=>j.id===+jid);
  const needTrait=selJeune&&selJeune.besoinTraitement==="oui";
  const handle=()=>{if(!obs.trim())return;if(needTrait&&!traitAdmin){alert("Ce jeune a un traitement : indiquez s'il a été administré.");return;}onSave({jeuneId:+jid,date,observation:obs,typeContact,traitementAdministre:needTrait?traitAdmin:undefined});setSaved(true);setObs("");setTypeContact("journee");setTraitAdmin("");setTimeout(()=>setSaved(false),2500);};
  return(<div style={{padding:"18px 14px",maxWidth:700,margin:"0 auto"}}>
    <div style={{...S.card}}>
      <div style={{marginBottom:14}}><label style={{...S.lbl}}>Jeune</label><select style={{...S.inp}} value={jid} onChange={e=>setJid(e.target.value)}><option value="">-- Sélectionner --</option>{vj.map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom} — {j.site}</option>)}</select></div>
      <div style={{marginBottom:14}}><label style={{...S.lbl}}>Date</label><input style={{...S.inp}} type="date" value={date} onChange={e=>setDate(e.target.value)} max={today}/></div>
      <div style={{marginBottom:14}}><label style={{...S.lbl}}>Type de contact</label><select style={{...S.inp}} value={typeContact} onChange={e=>setTypeContact(e.target.value)}><option value="journee">Journée du jeune</option><option value="rdv_parents">RDV téléphonique avec les parents</option><option value="rdv_exterieur">RDV téléphonique contact extérieur (en lien avec le jeune)</option></select></div>
      <div style={{marginBottom:14}}><label style={{...S.lbl}}>Observation</label><textarea style={{...S.inp,minHeight:110,resize:"vertical",lineHeight:1.6}} placeholder="Décrivez la journée du jeune, le contenu de l'échange..." value={obs} onChange={e=>setObs(e.target.value)}/></div>
      {needTrait&&<div style={{marginBottom:14,padding:"12px 14px",background:"#FFF5F5",border:"1.5px solid #F3C6C6",borderRadius:10}}><label style={{...S.lbl,color:"#C62828"}}>💊 Traitement administré aujourd'hui <span style={{fontWeight:800}}>(obligatoire)</span></label><div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:4}}>{[["totalite","En totalité","#2E7D32"],["moitie","De moitié","#E65100"],["aucun","Pas du tout","#C62828"]].map(([v,l,c])=><button key={v} onClick={()=>setTraitAdmin(v)} style={{flex:"1 1 30%",padding:"9px 8px",borderRadius:8,border:"1.5px solid "+(traitAdmin===v?c:C.border),background:traitAdmin===v?c:C.white,color:traitAdmin===v?"#fff":C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{l}</button>)}</div></div>}
      <button onClick={handle} style={{...S.btnP,width:"100%",justifyContent:"center",background:saved?"#66BB6A":"#2E7D32",boxShadow:"0 4px 14px rgba(46,125,50,0.35)"}}>{saved?<><Check size={17}/>Enregistré !</>:<><FileText size={17}/>Enregistrer le rapport</>}</button>
    </div>
    {(()=>{const allR=(rapports||[]).filter(r=>{if(!vj.some(j=>j.id===r.jeuneId))return false;if(siteF!=="Tous"){const j=allPool.find(j2=>j2.id===r.jeuneId);if(!j||j.site!==siteF)return false;}return true;}).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,MAX_DISPLAY);const totalPages=Math.ceil(allR.length/PER_PAGE);const pageR=allR.slice(rPage*PER_PAGE,(rPage+1)*PER_PAGE);return(<>
    {(user.role==="directeur"||user.role==="chef_service"||user.site==="Tous")&&<div style={{display:"flex",gap:7,margin:"18px 0 4px",flexWrap:"wrap"}}>{["Tous","Fatick","Djilass"].map(s=><button key={s} onClick={()=>{setSiteF(s);setRPage(0);}} style={{padding:"6px 14px",borderRadius:20,border:`1.5px solid ${siteF===s?C.gold:C.border}`,background:siteF===s?C.gold:C.white,color:siteF===s?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{s}</button>)}</div>}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"18px 0 10px",flexWrap:"wrap"}}><h3 style={{fontSize:14,fontWeight:800,color:C.dark,margin:0}}>Rapports ({allR.length})</h3>{totalPages>1&&<div style={{fontSize:12,color:C.mid}}>Page {rPage+1}/{totalPages}</div>}</div>
    {pageR.map(r=>{const j=allPool.find(j2=>j2.id===r.jeuneId)||{prenom:"?",nom:"",id:0};const tcLabels={"journee":"Journée du jeune","rdv_parents":"RDV tél. parents","rdv_exterieur":"RDV tél. contact ext."};return(<div key={r.id} style={{...S.card}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:5,flexWrap:"wrap"}}><div style={{fontWeight:800,fontSize:13,color:C.dark}}>{j?.prenom} {j?.nom}</div><div style={{fontSize:12,color:C.gold,fontWeight:700}}>{fmt(r.date)}</div></div>{r.traitementAdministre&&<div style={{display:"inline-block",marginBottom:5,fontSize:11.5,fontWeight:700,padding:"2px 9px",borderRadius:8,background:r.traitementAdministre==="totalite"?"#E8F5E9":r.traitementAdministre==="moitie"?"#FFF3E0":"#FFEBEE",color:r.traitementAdministre==="totalite"?"#2E7D32":r.traitementAdministre==="moitie"?"#E65100":"#C62828"}}>💊 Traitement : {r.traitementAdministre==="totalite"?"administré en totalité":r.traitementAdministre==="moitie"?"administré de moitié":"non administré"}</div>}{editTC===r.id?(user.role==="chef_service"||user.role==="directeur")&&<div style={{marginBottom:6,display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}><select style={{...S.inp,fontSize:12,padding:"4px 8px",flex:1}} value={r.typeContact||"journee"} onChange={e=>{onUpdate&&onUpdate(r.id,"typeContact",e.target.value);setEditTC(null);}}><option value="journee">Journée du jeune</option><option value="rdv_parents">RDV tél. parents</option><option value="rdv_exterieur">RDV tél. contact ext.</option></select><button onClick={()=>setEditTC(null)} style={{background:"none",border:"none",cursor:"pointer",color:C.light,fontSize:12}}>✕</button></div>:r.typeContact&&<div style={{fontSize:12,color:C.info||"#1565C0",fontWeight:600,marginBottom:4,background:"#E3F2FD",padding:"3px 8px",borderRadius:6,display:"inline-block",cursor:(user.role==="chef_service"||user.role==="directeur")?"pointer":undefined}} onClick={()=>{if(user.role==="chef_service"||user.role==="directeur")setEditTC(r.id);}}>{tcLabels[r.typeContact]||r.typeContact}{(user.role==="chef_service"||user.role==="directeur")&&<span style={{marginLeft:4,fontSize:12}}>✏️</span>}</div>}{editObs===r.id?<div style={{marginBottom:6}}><textarea autoFocus value={editText} onChange={e=>setEditText(e.target.value)} rows={4} style={{...S.inp,fontSize:12,resize:"vertical",width:"100%"}}/><div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap"}}><button onClick={()=>saveEdit(r)} style={{padding:"5px 14px",borderRadius:6,border:"none",background:"#2E7D32",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>✓ Enregistrer</button><button onClick={()=>{setEditObs(null);setEditText("");}} style={{padding:"5px 14px",borderRadius:6,border:"1px solid "+C.border,background:C.white,color:C.mid,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Annuler</button></div></div>:<p style={{margin:0,fontSize:12,color:C.mid,lineHeight:1.5}}>{r.observation}</p>}{r.author&&<div style={{fontSize:11.5,color:C.light,marginTop:4,fontStyle:"italic"}}>Rédigé par {r.author}{(r.horodatage||r.createdAt)?(" le "+(r.horodatage||new Date(r.createdAt).toISOString()).replace("T"," à ").slice(0,19)):""}</div>}{r.modifieLe&&<div style={{fontSize:11.5,color:C.goldDark,marginTop:2,fontStyle:"italic"}}>✏️ Modifié le {r.modifieLe.replace("T"," à ").slice(0,19)} par {r.modifiePar}{isEncadrement&&(r.historique||[]).length>0&&<span onClick={()=>setShowHist(showHist===r.id?null:r.id)} style={{marginLeft:8,cursor:"pointer",textDecoration:"underline",fontWeight:700}}>{showHist===r.id?"masquer les versions":"voir "+(r.historique||[]).length+" version(s) antérieure(s)"}</span>}</div>}{isEncadrement&&showHist===r.id&&(r.historique||[]).slice().reverse().map((h,i)=><div key={i} style={{marginTop:6,padding:"7px 10px",background:C.goldLight,borderRadius:8,border:"1px dashed "+C.gold}}><div style={{fontSize:12,fontWeight:700,color:C.goldDark,marginBottom:2}}>Version remplacée le {h.date.replace("T"," à ").slice(0,19)} — modifiée par {h.par}</div><div style={{fontSize:12,color:C.mid,lineHeight:1.4}}>{h.ancienTexte}</div></div>)}<div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap"}}>{canEditR(r)&&editObs!==r.id&&<button onClick={(e)=>{e.stopPropagation();setEditObs(r.id);setEditText(r.observation||"");setShowHist(null);}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>✏️ Modifier</button>}{(user.role==="chef_service"||user.role==="directeur")&&<button onClick={(e)=>{e.stopPropagation();if(confirm("Supprimer ce rapport ?"))onDelete(r.id);}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Supprimer</button>}</div></div>);})}
    {totalPages>1&&<div style={{display:"flex",justifyContent:"center",gap:6,marginTop:14,flexWrap:"wrap"}}>{rPage>0&&<button onClick={()=>setRPage(rPage-1)} style={{...S.btnO,padding:"6px 14px",fontSize:12}}><ChevronLeft size={14}/>Précédent</button>}{Array.from({length:totalPages},(_,i)=>i).filter(i=>i===0||i===totalPages-1||Math.abs(i-rPage)<=2).map((i,idx,arr)=>{const prev=arr[idx-1];const gap=prev!==undefined&&i-prev>1;return(<React.Fragment key={i}>{gap&&<span style={{padding:"6px 4px",color:C.light}}>…</span>}<button onClick={()=>setRPage(i)} style={{padding:"6px 12px",borderRadius:8,border:i===rPage?"2px solid "+C.gold:"1px solid #ddd",background:i===rPage?C.goldLight:"#fff",color:i===rPage?C.goldDark:C.mid,fontWeight:i===rPage?800:500,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{i+1}</button></React.Fragment>);})}{rPage<totalPages-1&&<button onClick={()=>setRPage(rPage+1)} style={{...S.btnO,padding:"6px 14px",fontSize:12}}>Suivant<ChevronRight size={14}/></button>}</div>}
    </>);})()}
  </div>);
}

function Presences({user,presences,onCP,onDelP,jeunes}){
  const[gomme,setGomme]=useState(false);
  const pool=(jeunes&&jeunes.length!==undefined)?jeunes:JEUNES;const vj=(user.role==="educateur"||user.role==="coordinateur_site")?pool.filter(j=>(j.id<100)&&(user.site==="Tous"||j.site===user.site)):pool;
  const[site,setSite]=useState("Tous");
  const vis=site==="Tous"?vj:vj.filter(j=>j.site===site);
  const next={Présent:"Absent",Absent:"Retard",Retard:"Présent"};
  return(<div style={{padding:"18px 14px",maxWidth:900,margin:"0 auto"}}>
    <div style={{display:"flex",justifyContent:"flex-end",marginBottom:10,flexWrap:"wrap"}}><button onClick={()=>setGomme(g=>!g)} style={{padding:"6px 14px",borderRadius:20,border:"1.5px solid "+(gomme?"#C62828":C.border),background:gomme?"#FFEBEE":C.white,color:gomme?"#C62828":C.mid,fontWeight:800,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{gomme?"✕ Mode suppression actif":"Supprimer des présences"}</button></div>
    {gomme&&<div style={{fontSize:12,color:"#C62828",fontWeight:700,marginBottom:10,padding:"8px 10px",background:"#FFEBEE",borderRadius:8}}>Touchez une case renseignée pour effacer la saisie. Les cases vides sont sans effet.</div>}
    {user.role!=="educateur"&&<div style={{display:"flex",gap:7,marginBottom:14,flexWrap:"wrap"}}>{["Tous","Fatick","Djilass"].map(s=><button key={s} onClick={()=>setSite(s)} style={{padding:"6px 14px",borderRadius:20,border:`1.5px solid ${site===s?C.gold:C.border}`,background:site===s?C.gold:C.white,color:site===s?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{s}</button>)}</div>}
    <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"separate",borderSpacing:"0 6px",minWidth:500}}>
      <thead><tr><th style={{textAlign:"left",fontSize:12,fontWeight:700,color:C.light,padding:"0 8px 4px",textTransform:"uppercase"}}>Jeune</th>{WEEKDATES.map((d,i)=><th key={i} style={{fontSize:11.5,fontWeight:700,color:C.light,textAlign:"center",padding:"0 2px 4px"}}>{WD[i]}<br/>{d.slice(8)}</th>)}<th style={{fontSize:12,fontWeight:700,color:C.light,textAlign:"center",padding:"0 8px 4px"}}>Total</th></tr></thead>
      <tbody>{vis.map(j=>{const jp=presences.filter(p=>p.jeuneId===j.id&&WEEKDATES.includes(p.date));return(<tr key={j.id}>
        <td style={{padding:"8px",background:C.white,borderRadius:"10px 0 0 10px",border:`1px solid ${C.border}`,borderRight:"none",fontWeight:700,fontSize:13,color:C.dark,whiteSpace:"nowrap"}}>{j.prenom} <span style={{color:C.light}}>{j.nom.slice(0,1)}.</span> <span style={{fontSize:11.5,color:C.gold}}>{j.site}</span></td>
        {WEEKDATES.map((date,i)=>{const p=jp.find(p2=>p2.date===date);const st=p?.statut||"Présent";const sc2=SC[st]||SC.Présent;return(<td key={i} style={{background:C.white,border:`1px solid ${C.border}`,borderLeft:"none",borderRight:"none",textAlign:"center",padding:"4px 2px"}}><button onClick={()=>{if(gomme){if(st&&onDelP)onDelP(j.id,date);}else onCP(j.id,date,next[st]);}} style={{width:30,height:30,borderRadius:7,background:gomme&&st?"#FFEBEE":sc2.bg,border:gomme&&st?"1px dashed #C62828":"none",cursor:"pointer",color:gomme&&st?"#C62828":sc2.text,fontWeight:800,fontSize:13}}>{gomme&&st?"✕":sc2.icon}</button></td>);})}
        <td style={{padding:"8px",background:C.white,borderRadius:"0 10px 10px 0",border:`1px solid ${C.border}`,borderLeft:"none",textAlign:"center",fontWeight:800,fontSize:13,color:"#2E7D32"}}>{jp.filter(p=>p.statut==="Présent").length}/7</td>
      </tr>);})}
      </tbody>
    </table></div>
  </div>);
}

function Evenements({user,evenements,onAdd,onDelete,majeurs,onUpdateAll,jeunes}){
  const allPool=[...(jeunes||JEUNES),...(majeurs||MAJEURS)];const vj=(user.role==="educateur"||user.role==="coordinateur_site")?(user.isEducMajeur?allPool.filter(j=>(j.id>=100)&&(user.site==="Tous"||j.site===user.site)):allPool.filter(j=>(j.id<100)&&(user.site==="Tous"||j.site===user.site))):allPool;
  const[jid,setJid]=useState(vj[0]?.id||"");
  const[titre,setTitre]=useState(""),[ desc,setDesc]=useState(""),[ grav,setGrav]=useState("Léger"),[ date2,setDate2]=useState(today),[categ,setCateg]=useState("jeune"),[typeEv,setTypeEv]=useState("incident");
  const[eig,setEig]=useState(false),[ eigOpen,setEigOpen]=useState(null);
  const[editEv,setEditEv]=useState(null),[ edTitre,setEdTitre]=useState(""),[ edDesc,setEdDesc]=useState(""),[ edGrav,setEdGrav]=useState("Léger"),[ edType,setEdType]=useState("incident");
  const isEncadrementE=user.role==="chef_service"||user.role==="directeur"||user.role==="coordinateur_site";
  const canEditE=(e)=>{if(isEncadrementE)return true;if(user.role!=="educateur")return false;if(e.authorId!=null&&user.id!=null)return e.authorId===user.id;return !!e.author&&e.author===user.name;};
  const startEditE=(e)=>{setEditEv(e.id);setEdTitre(e.titre||"");setEdDesc(e.description||"");setEdGrav(GC[e.gravite]?e.gravite:"Léger");setEdType(e.type||"incident");};
  const saveEditE=(e)=>{const t=edTitre.trim();if(!t){alert("Le titre ne peut pas être vide.");return;}if(t===e.titre&&edDesc.trim()===(e.description||"")&&edGrav===e.gravite&&edType===(e.type||"incident")){setEditEv(null);return;}const ts=new Date().toISOString();const updated=(evenements||[]).map(x=>x.id===e.id?{...x,titre:t,description:edDesc.trim(),gravite:edGrav,type:edType,modifieLe:ts,modifiePar:user.name||"?",historique:[...(x.historique||[]),{date:ts,par:user.name||"?",ancienTitre:x.titre,ancienTexte:x.description||"",ancienneGravite:x.gravite,ancienType:x.type||"incident"}]}:x);if(onUpdateAll)onUpdateAll(updated);setEditEv(null);};
  const updEig=(id,patch)=>{const updated=(evenements||[]).map(x=>x.id===id?{...x,eigData:{...(x.eigData||{}),...patch}}:x);if(onUpdateAll)onUpdateAll(updated);};
  const toggleEig=(id,val)=>{const updated=(evenements||[]).map(x=>x.id===id?{...x,eig:val,eigData:val?(x.eigData||{destinataires:"",dateTransmission:"",accuseReception:false,statutCloture:"En cours"}):x.eigData}:x);if(onUpdateAll)onUpdateAll(updated);if(!val)setEigOpen(null);};
  const[open2,setOpen2]=useState(false),[ saved,setSaved]=useState(false),[ fg,setFg]=useState("Tous"),[ fSite,setFSite]=useState("Tous");
  const canSeeEduc=user.role==="chef_service"||user.role==="directeur";
  const vis=(evenements||[]).filter(e=>{if(e.categorie==="educateur"&&!canSeeEduc)return false;const ok=e.categorie==="educateur"||vj.some(j=>j.id===e.jeuneId);const fok=fg==="Tous"||e.gravite===fg;const j=allPool.find(j2=>j2.id===e.jeuneId);const sok=fSite==="Tous"||(j?j.site===fSite:e.site===fSite);return ok&&fok&&sok;});
  const handle=()=>{if(!titre.trim())return;const evSite=(categ==="jeune"||categ==="jeune_pro")?(allPool.find(j=>j.id===+jid)?.site||null):(user.site&&user.site!=="Tous"?user.site:null);onAdd({jeuneId:(categ==="jeune"||categ==="jeune_pro")?+jid:null,date:date2,titre,description:desc,gravite:grav,type:typeEv,categorie:categ,site:evSite,eig,eigData:eig?{destinataires:"",dateTransmission:"",accuseReception:false,statutCloture:"En cours"}:null});setSaved(true);setTitre("");setDesc("");setEig(false);setTimeout(()=>{setSaved(false);setOpen2(false);},2000);};
  return(<div style={{padding:"18px 14px",maxWidth:700,margin:"0 auto"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap"}}><h2 style={{fontSize:18,fontWeight:900,color:C.dark,margin:0,letterSpacing:"-0.01em"}}>Événements indésirables</h2><button onClick={()=>exportIncidentsXLSX(evenements,jeunes)} style={{background:C.primary,color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:13,fontWeight:600}}>Exporter CSV</button></div>
      <button style={{...S.btnO,fontSize:12,padding:"8px 14px"}} onClick={()=>setOpen2(!open2)}><Plus size={14}/>Déclarer</button>
    </div>
    {open2&&<div style={{...S.card,borderLeft:`4px solid ${C.orange}`,marginBottom:14}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={{...S.lbl}}>Concerne</label><select style={{...S.inp}} value={categ} onChange={e=>setCateg(e.target.value)}><option value="jeune">Un jeune</option><option value="jeune_pro">Jeune à professionnel</option><option value="educateur">Entre éducateurs</option></select></div><div><label style={{...S.lbl}}>Type</label><select style={{...S.inp}} value={typeEv} onChange={e=>setTypeEv(e.target.value)}><option value="incident">Incident</option><option value="plainte">Plainte</option><option value="reclamation">Reclamation</option></select></div></div>
      {(categ==="jeune"||categ==="jeune_pro")&&<div style={{marginBottom:10}}><label style={{...S.lbl}}>Jeune</label><select style={{...S.inp}} value={jid} onChange={e=>setJid(e.target.value)}>{vj.map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom}{j.id>=100?" (majeur)":""}</option>)}</select></div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
        <div><label style={{...S.lbl}}>Date</label><input style={{...S.inp}} type="date" value={date2} onChange={e=>setDate2(e.target.value)} max={today}/></div>
        <div><label style={{...S.lbl}}>Gravité</label><select style={{...S.inp}} value={grav} onChange={e=>setGrav(e.target.value)}><option>Léger</option><option>Moyen</option><option>Grave</option></select></div>
      </div>
      <div style={{marginBottom:10}}><label style={{...S.lbl}}>Qualification</label><select style={{...S.inp}} value={eig?"oui":"non"} onChange={e=>setEig(e.target.value==="oui")}><option value="non">Événement indésirable simple</option><option value="oui">EIG — signalement obligatoire aux autorités (art. L331-8-1 CASF)</option></select></div>
      <div style={{marginBottom:10}}><label style={{...S.lbl}}>Titre</label><input style={{...S.inp}} value={titre} onChange={e=>setTitre(e.target.value)} placeholder="Ex: Altercation verbale"/></div>
      <div style={{marginBottom:10}}><label style={{...S.lbl}}>Description</label><textarea style={{...S.inp,minHeight:80,resize:"vertical"}} value={desc} onChange={e=>setDesc(e.target.value)} placeholder="Décrivez brièvement l'incident..."/></div>
      <button onClick={handle} style={{...S.btnO,width:"100%",justifyContent:"center",background:saved?"#4CAF50":undefined}}>{saved?<><Check size={17}/>Déclaré !</>:<><AlertTriangle size={17}/>Enregistrer</>}</button>
    </div>}
    {(user.role==="directeur"||user.role==="chef_service"||user.site==="Tous")&&<div style={{display:"flex",gap:7,marginBottom:8,flexWrap:"wrap"}}>{["Tous","Fatick","Djilass"].map(s=><button key={s} onClick={()=>setFSite(s)} style={{padding:"6px 14px",borderRadius:20,border:`1.5px solid ${fSite===s?C.gold:C.border}`,background:fSite===s?C.gold:C.white,color:fSite===s?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{s}</button>)}</div>}
    <div style={{display:"flex",gap:7,marginBottom:14,flexWrap:"wrap"}}>{["Tous","Léger","Moyen","Grave"].map(g=><button key={g} onClick={()=>setFg(g)} style={{padding:"6px 16px",borderRadius:24,border:`1.5px solid ${fg===g?C.gold:C.border}`,background:fg===g?`linear-gradient(135deg,${C.goldLight},#F5E6B0)`:C.white,color:fg===g?C.goldDark:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",transition:"all 0.2s ease",boxShadow:fg===g?C.shadowGold:"none"}}>{g}</button>)}</div>
    {vis.sort((a,b)=>b.date.localeCompare(a.date)).map(e=>{const j=allPool.find(j2=>j2.id===e.jeuneId)||{prenom:"?",nom:"",id:0};const gc=GC[e.gravite]||GC["Léger"];return(<div key={e.id} style={{...S.card,borderLeft:`4px solid ${gc.dot}`,animation:"cardEnter 0.3s ease-out"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:7,flexWrap:"wrap"}}><div><div style={{fontWeight:800,fontSize:14,color:C.dark}}>{e.titre}</div><div style={{fontSize:12,color:C.light,marginTop:1}}>{j?.prenom} {j?.nom} · {fmt(e.date)}</div></div><div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>{(()=>{const TY={incident:{l:"Incident",bg:"#ECEFF1",c:"#546E7A"},plainte:{l:"Plainte",bg:"#FFF3E0",c:"#E65100"},reclamation:{l:"Réclamation",bg:"#E3F2FD",c:"#1565C0"}};const t=TY[e.type]||TY.incident;return<span style={{background:t.bg,color:t.c,borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:700}}>{t.l}</span>;})()}<span style={{background:gc.bg,color:gc.text,borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:700}}>{e.gravite}</span></div></div>{editEv===e.id?<div style={{marginBottom:6}}><label style={{...S.lbl}}>Titre</label><input value={edTitre} onChange={ev2=>setEdTitre(ev2.target.value)} style={{...S.inp,fontSize:12,marginBottom:6}}/><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:6}}><div><label style={{...S.lbl}}>Type</label><select value={edType} onChange={ev2=>setEdType(ev2.target.value)} style={{...S.inp,fontSize:12}}><option value="incident">Incident</option><option value="plainte">Plainte</option><option value="reclamation">Réclamation</option></select></div><div><label style={{...S.lbl}}>Gravité</label><select value={edGrav} onChange={ev2=>setEdGrav(ev2.target.value)} style={{...S.inp,fontSize:12}}><option>Léger</option><option>Moyen</option><option>Grave</option></select></div></div><label style={{...S.lbl}}>Description</label><textarea value={edDesc} onChange={ev2=>setEdDesc(ev2.target.value)} rows={3} style={{...S.inp,fontSize:12,resize:"vertical",width:"100%"}}/><div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap"}}><button onClick={()=>saveEditE(e)} style={{padding:"5px 14px",borderRadius:6,border:"none",background:"#2E7D32",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>✓ Enregistrer</button><button onClick={()=>setEditEv(null)} style={{padding:"5px 14px",borderRadius:6,border:"1px solid "+C.border,background:C.white,color:C.mid,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Annuler</button></div></div>:<p style={{margin:0,fontSize:12,color:C.mid,lineHeight:1.5}}>{e.description}</p>}{e.author&&<div style={{fontSize:11.5,color:C.light,marginTop:4,fontStyle:"italic"}}>Rédigé par {e.author}{e.horodatage&&(" le "+e.horodatage.replace("T"," à "))}</div>}{e.modifieLe&&<div style={{fontSize:11.5,color:C.goldDark,marginTop:2,fontStyle:"italic"}}>✏️ Modifié le {e.modifieLe.replace("T"," à ").slice(0,19)} par {e.modifiePar}</div>}{canEditE(e)&&editEv!==e.id&&<button onClick={()=>{startEditE(e);}} style={{marginTop:6,padding:"4px 10px",borderRadius:6,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>✏️ Modifier</button>}{e.numeroSuivi&&<div style={{fontSize:11.5,color:C.gold,marginTop:2,fontWeight:700}}>N° suivi: {e.numeroSuivi}</div>}{e.eig&&<div style={{marginTop:8,padding:"8px 10px",background:"#FFEBEE",border:"1.5px solid #C62828",borderRadius:8}}><div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}><span style={{background:"#C62828",color:"#fff",borderRadius:4,padding:"2px 8px",fontSize:11.5,fontWeight:800,letterSpacing:"0.05em"}}>EIG</span><span style={{fontSize:12,fontWeight:700,color:"#C62828"}}>{e.eigData?.statutCloture||"En cours"}</span>{e.eigData?.dateTransmission?<span style={{fontSize:11.5,color:"#555"}}>Transmis le {e.eigData.dateTransmission}{e.eigData.accuseReception?" · AR reçu":" · AR en attente"}</span>:<span style={{fontSize:11.5,color:"#C62828",fontWeight:700}}>⚠ Non transmis aux autorités</span>}{e.eigData?.destinataires&&<span style={{fontSize:11.5,color:"#555"}}>→ {e.eigData.destinataires}</span>}</div>{canSeeEduc&&eigOpen===e.id&&<div style={{marginTop:8,display:"grid",gap:6}}><div><label style={{fontSize:11.5,fontWeight:700,color:"#555"}}>Destinataires (CD, PJJ, parquet…)</label><input style={{...S.inp,fontSize:12}} defaultValue={e.eigData?.destinataires||""} onBlur={ev2=>updEig(e.id,{destinataires:ev2.target.value})} placeholder="Ex: CD93 — Cellule de recueil, PJJ"/></div><div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}><div><label style={{fontSize:11.5,fontWeight:700,color:"#555",display:"block"}}>Date de transmission</label><input type="date" style={{...S.inp,fontSize:12,width:"auto"}} value={e.eigData?.dateTransmission||""} onChange={ev2=>updEig(e.id,{dateTransmission:ev2.target.value})}/></div><label style={{display:"flex",alignItems:"center",gap:5,fontSize:12,fontWeight:700,color:"#555",cursor:"pointer",marginTop:14,flexWrap:"wrap"}}><input type="checkbox" checked={!!e.eigData?.accuseReception} onChange={ev2=>updEig(e.id,{accuseReception:ev2.target.checked})} style={{accentColor:"#C62828"}}/>Accusé de réception</label><div><label style={{fontSize:11.5,fontWeight:700,color:"#555",display:"block"}}>Statut</label><select style={{...S.inp,fontSize:12,width:"auto"}} value={e.eigData?.statutCloture||"En cours"} onChange={ev2=>updEig(e.id,{statutCloture:ev2.target.value})}><option>En cours</option><option>Clôturé</option></select></div></div></div>}</div>}{(user.role==="chef_service"||user.role==="directeur"||user.role==="coordinateur_site")&&<div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap"}}>{<button onClick={()=>{const num=prompt(e.numeroSuivi?"Modifier le n° suivi (actuel: "+e.numeroSuivi+") :":"Numéro de suivi à attribuer :",e.numeroSuivi||"");if(num!==null&&num.trim()){const updated=(evenements||[]).map(x=>x.id===e.id?{...x,numeroSuivi:num.trim()}:x);if(onUpdateAll)onUpdateAll(updated);}}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{e.numeroSuivi?"Modifier N° suivi":"Attribuer N° suivi"}</button>}{!e.eig&&<button onClick={()=>{if(confirm("Qualifier cet événement en EIG (signalement obligatoire aux autorités) ?"))toggleEig(e.id,true);}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:C.white,color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Qualifier EIG</button>}{e.eig&&<button onClick={()=>setEigOpen(eigOpen===e.id?null:e.id)} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:eigOpen===e.id?"#C62828":"#FFEBEE",color:eigOpen===e.id?"#fff":"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{eigOpen===e.id?"Fermer suivi EIG":"Suivi EIG"}</button>}{e.eig&&<button onClick={()=>{if(confirm("Retirer la qualification EIG ? Le suivi sera conservé."))toggleEig(e.id,false);}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #9E9E9E",background:"#F5F5F5",color:"#757575",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Retirer EIG</button>}{onDelete&&<button onClick={()=>{if(confirm("Supprimer cet événement ?"))onDelete(e.id);}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Supprimer</button>}</div>}</div>);})}
    {vis.length===0&&<div style={{textAlign:"center",padding:48,color:C.light,fontSize:13,fontWeight:600,opacity:0.7}}>Aucun événement</div>}
  </div>);
}

function RapportHebdo({user,rapports,presences,evenements,jeunes,majeurs,onSaveHebdo,sejourConfig}){
  const allJeunes=[...(jeunes||JEUNES),...(majeurs||MAJEURS)];
  const mj=user.role==="coordinateur_site"?allJeunes.filter(j=>(user.site==="Tous"||j.site===user.site)&&(user.isEducMajeur?j.id>=100:j.id<100)):user.role==="educateur"?allJeunes.filter(j=>user.site==="Tous"||j.site===user.site):allJeunes;
  const[site,setSite]=useState("Djilass");
  const siteJeunes=mj.filter(j=>j.site===site);
  const[selJeune,setSelJeune]=useState("");
  const getISOWeekNum=(dt)=>{const d=new Date(Date.UTC(dt.getFullYear(),dt.getMonth(),dt.getDate()));d.setUTCDate(d.getUTCDate()+4-(d.getUTCDay()||7));const y1=new Date(Date.UTC(d.getUTCFullYear(),0,1));return Math.ceil((((d-y1)/86400000)+1)/7);};
  const DEFAUT_DEBUT={Djilass:"2026-03-30",Fatick:"2026-03-16"};
  const getDebut=(s)=>(sejourConfig&&sejourConfig[s]&&sejourConfig[s].dateDebut)||DEFAUT_DEBUT[s]||DEFAUT_DEBUT.Djilass;
  const calcSiteWeek=(s)=>{const d0=new Date(getDebut(s)+"T00:00:00");const diff=Math.floor((Date.now()-d0.getTime())/86400000);const sw=Math.floor(diff/7)+1;return String(sw>0?sw:1).padStart(2,"0");};
  const weekRange=(w,s)=>{const d0=new Date(getDebut(s)+"T00:00:00");const st=new Date(d0);st.setDate(d0.getDate()+(parseInt(w,10)-1)*7);const en=new Date(st);en.setDate(st.getDate()+6);const iso=d=>d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");return{start:iso(st),end:iso(en)};};
  const[weekNum,setWeekNum]=useState(()=>calcSiteWeek("Djilass"));
  useEffect(()=>{setWeekNum(calcSiteWeek(site));},[site,sejourConfig]);
  const[groupText,setGroupText]=useState("");
  const[persoTexts,setPersoTexts]=useState({});
  const[destTexts,setDestTexts]=useState({});
  const deriveDest=(j)=>{const em=(j&&j.emailASE||"").trim();if(!em||!em.includes("@"))return"";const local=em.split("@")[0].replace(/[0-9_-]+/g,".");const parts=local.split(".").filter(Boolean);if(parts.length>=2){return parts[0][0].toUpperCase()+" "+parts[parts.length-1].toUpperCase();}if(parts.length===1&&parts[0].length>2){return parts[0][0].toUpperCase()+" "+parts[0].slice(1).toUpperCase();}return"";};
  const autoFillDest=(js)=>{setDestTexts(p=>{const n={...p};js.forEach(j=>{if(!(n[j.id]||"").trim()){const d=deriveDest(j);if(d)n[j.id]=d;}});return n;});};
  const[hebdoStatuts,setHebdoStatuts]=useState({});
  const canValider=user.role==="directeur"||user.role==="chef_service"||user.role==="coordinateur_site";
  const stKey=(jId)=>site+"_S"+weekNum+"_"+jId;
  const getStatut=(jId)=>hebdoStatuts[stKey(jId)]||{statut:"brouillon"};
  const setStatut=(jId,statut)=>{const v={statut,par:user?.name||"?",le:new Date().toISOString().slice(0,16).replace("T"," ")};setHebdoStatuts(prev=>{const n={...prev,[stKey(jId)]:v};const saved=loadLS()||{};if(!saved.hebdo)saved.hebdo={};saved.hebdo.statuts=n;try{localStorage.setItem(LS_KEY,JSON.stringify({...saved,hebdo:saved.hebdo}));}catch(e){}fbSet("hebdo",saved.hebdo);return n;});};
  const ST_BADGE={brouillon:{label:"Brouillon",bg:"#ECEFF1",c:"#546E7A"},valide:{label:"Validé",bg:"#E8F5E9",c:"#2E7D32"},transmis:{label:"Transmis ASE",bg:"#FFF8E1",c:"#B8860B"}};
  const[preview,setPreview]=useState(false);
  const[sending,setSending]=useState(false);
  const[sent,setSent]=useState(false);
  const siteCode=site==="Djilass"?"DJI":"FAT";
  const getFileName=(j)=>siteCode+"-RH-"+j.prenom.substring(0,3).toUpperCase()+"-S"+weekNum;
  const refA=(jId)=>{const j=allJeunes.find(x=>x.id===jId);if(!j)return"";const u=USERS.find(u=>u.name===j.referentA);return u?u.name:j.referentA||""};
  const refB=(jId)=>{const j=allJeunes.find(x=>x.id===jId);if(!j)return"";const u=USERS.find(u=>u.name===j.referentB);return u?u.name:j.referentB||""};
  
  useEffect(()=>{
    const saved=loadLS();
    if(saved&&saved.hebdo){
      if(saved.hebdo[site+"_group"])setGroupText(saved.hebdo[site+"_group"]);
      if(saved.hebdo.perso)setPersoTexts(saved.hebdo.perso);
      if(saved.hebdo.dest)setDestTexts(saved.hebdo.dest);
      if(saved.hebdo.statuts)setHebdoStatuts(saved.hebdo.statuts);
    }
  },[site]);

  const saveHebdoData=()=>{
    const saved=loadLS()||{};
    if(!saved.hebdo)saved.hebdo={};
    saved.hebdo[site+"_group"]=groupText;
    saved.hebdo.perso={...persoTexts};
    saved.hebdo.dest={...destTexts};
    saved.hebdo.statuts={...hebdoStatuts};
    try{localStorage.setItem(LS_KEY,JSON.stringify({...saved,hebdo:saved.hebdo}));}catch(e){}
    fbSet("hebdo",saved.hebdo);
  };

  const[iaLoading,setIaLoading]=useState(false);
  const[iaProgress,setIaProgress]=useState("");
  const getISOWeekStr=(dateStr)=>{const d=new Date(dateStr);d.setHours(0,0,0,0);d.setDate(d.getDate()+3-(d.getDay()+6)%7);const w1=new Date(d.getFullYear(),0,4);return String(1+Math.round(((d-w1)/86400000-3+(w1.getDay()+6)%7)/7)).padStart(2,"0");};
  const syntheseIA=async()=>{
    if(iaLoading)return;
    if(!confirm("Lancer la synthèse IA des rapports journaliers et événements de la semaine "+weekNum+" pour le site de "+site+" ?\n\nLes parties personnelles existantes seront remplacées."))return;
    setIaLoading(true);
    const{start:wkS,end:wkE}=weekRange(weekNum,site);
    const results={};const errors=[];let done=0;
    const cible=siteJeunes.filter(j=>{if(getStatut(j.id).statut!=="brouillon")return false;const jr=(rapports||[]).filter(r=>r.jeuneId===j.id&&r.date&&getISOWeekStr(r.date)===isoTarget);const je=(evenements||[]).filter(e=>e.jeuneId===j.id&&e.date&&getISOWeekStr(e.date)===isoTarget);return jr.length>0||je.length>0;});
    if(cible.length===0){alert("Aucun rapport ni événement trouvé pour la semaine "+weekNum+".");setIaLoading(false);setIaProgress("");return;}
    for(const j of cible){
      setIaProgress(j.prenom+" "+(j.nom||"")+" ("+(done+1)+"/"+cible.length+")");
      const jRapports=(rapports||[]).filter(r=>r.jeuneId===j.id&&r.date&&getISOWeekStr(r.date)===isoTarget).sort((a,b)=>a.date.localeCompare(b.date)).map(r=>({date:r.date,typeContact:r.typeContact||"journee",observation:r.observation||""}));
      const jEvenements=(evenements||[]).filter(e=>e.jeuneId===j.id&&e.date&&getISOWeekStr(e.date)===isoTarget).sort((a,b)=>a.date.localeCompare(b.date)).map(e=>({date:e.date,titre:e.titre||"",description:e.description||"",gravite:e.gravite||"",type:e.type||""}));
      try{
        const resp=await fetch(SYNTHESE_API_BASE+"/api/synthese-hebdo",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({jeune:{prenom:j.prenom,nom:j.nom},semaine:weekNum,site,rapports:jRapports,evenements:jEvenements})});
        if(!resp.ok)throw new Error("HTTP "+resp.status);
        const data=await resp.json();
        if(data&&data.synthese)results[j.id]=data.synthese;else throw new Error("Réponse vide");
      }catch(err){errors.push(j.prenom+" "+(j.nom||"")+" : "+err.message);}
      done++;
    }
    if(Object.keys(results).length>0){setPersoTexts(p=>({...p,...results}));autoFillDest(cible.filter(j=>results[j.id]));setTimeout(saveHebdoData,200);}
    setIaLoading(false);setIaProgress("");
    alert("Synthèse IA terminée : "+Object.keys(results).length+"/"+cible.length+" jeunes."+(errors.length?"\n\nErreurs :\n"+errors.join("\n"):""));
  };
  const compileWeekRapports=()=>{const getISOWeek=(dateStr)=>{const d=new Date(dateStr);d.setHours(0,0,0,0);d.setDate(d.getDate()+3-(d.getDay()+6)%7);const w1=new Date(d.getFullYear(),0,4);return String(1+Math.round(((d-w1)/86400000-3+(w1.getDay()+6)%7)/7)).padStart(2,"0")};const compiled={};let total=0;const{start:wkS,end:wkE}=weekRange(weekNum,site);siteJeunes.forEach(j=>{if(getStatut(j.id).statut!=="brouillon")return;const jRapports=(rapports||[]).filter(r=>{if(r.jeuneId!==j.id||!r.date)return false;return r.date>=wkS&&r.date<=wkE});if(jRapports.length>0){jRapports.sort((a,b)=>a.date.localeCompare(b.date));compiled[j.id]=jRapports.map(r=>{const dt=new Date(r.date);const dayName=dt.toLocaleDateString("fr-FR",{weekday:"long"});return dayName.charAt(0).toUpperCase()+dayName.slice(1)+" ("+r.date+") : "+r.observation}).join("\n\n");total+=jRapports.length}});setPersoTexts(p=>({...p,...compiled}));autoFillDest(siteJeunes.filter(j=>compiled[j.id]));setTimeout(saveHebdoData,200);alert("Compilation terminée : "+Object.keys(compiled).length+" jeunes, "+total+" rapports trouvés pour la semaine "+weekNum)};
const generateDocx=async(jeune)=>{console.log("generateDocx called for",jeune.prenom);const logoB64=LOGO.split(",")[1];const logoBin=atob(logoB64);const logoBuffer=new Uint8Array(logoBin.length);for(let i=0;i<logoBin.length;i++)logoBuffer[i]=logoBin.charCodeAt(i);
    const fileName=getFileName(jeune);
    const ra=refA(jeune.id);
    const rb=refB(jeune.id);
    const refs=[ra,rb].filter(Boolean).join(" – ");
    const border={style:BorderStyle.SINGLE,size:1,color:"999999"};
    const borders={top:border,bottom:border,left:border,right:border};
    const doc=new Document({
      styles:{default:{document:{run:{font:"Arial",size:22}}}},
      sections:[{
        properties:{page:{size:{width:11906,height:16838},margin:{top:1200,right:1200,bottom:1200,left:1200}}},
        headers:{default:new Header({children:[
          new Table({width:{size:9506,type:WidthType.DXA},columnWidths:[3000,6506],rows:[
            new TableRow({children:[
              new TableCell({borders,width:{size:3000,type:WidthType.DXA},verticalAlign:"center",children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({data:logoBuffer,transformation:{width:120,height:100},type:"png",altText:{title:"Logo PDSR",description:"Logo Association PDSR",name:"logo"}})]}),new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:100},children:[new TextRun({text:"Association PDSR",bold:true,size:22,font:"Arial"})]})]}),
              new TableCell({borders,width:{size:6506,type:WidthType.DXA},children:[
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Rapport Hebdomadaire",bold:true,size:28,font:"Arial"})]}),
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:(jeune.nom||"").toUpperCase()+" "+jeune.prenom,bold:true,size:24,font:"Arial"})]}),
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Semaine "+parseInt(weekNum,10),size:22,font:"Arial"})]}),
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Groupe de "+site.toUpperCase(),bold:true,size:22,font:"Arial"})]})
              ]})
            ]})
          ]})
        ]})},
        footers:{default:new Footer({children:[
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:100},children:[new TextRun({text:"Association PDSR, 28 rue rouget de Lisle 93160 Noisy le Grand",size:16,font:"Arial"})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Secrétariat : 28 Cité de la plante 16100 Chateaubernard",size:16,font:"Arial"})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"associationpdsr@gmail.com / tél : 06 24 75 34 31 - 05 17 22 59 33",size:16,font:"Arial"})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"SIRET : 835 155 466 00013  APE : 88999",size:16,font:"Arial"})]})
        ]})},
        children:[
          new Paragraph({spacing:{after:200},children:[new TextRun({text:"A l’attention de "+((destTexts[jeune.id]||"").trim()||deriveDest(jeune)||"…"),italics:true,size:22})]}),
          new Paragraph({spacing:{before:300,after:200},children:[new TextRun({text:"Cette semaine sur le groupe :",bold:true,size:24,underline:{}})]}),
          ...groupText.split("\n").map(line=>new Paragraph({spacing:{after:100},children:[new TextRun({text:line,size:22})]})),
          new Paragraph({spacing:{before:400,after:200},children:[new TextRun({text:"La semaine de "+(jeune.nom||"").toUpperCase()+" "+jeune.prenom+" :",bold:true,size:24,underline:{}})]}),
          ...(persoTexts[jeune.id]||"").split("\n").map(line=>new Paragraph({spacing:{after:100},children:[new TextRun({text:line,size:22})]})),
          new Paragraph({spacing:{before:500},border:{top:{style:BorderStyle.SINGLE,size:4,color:"999999"}},children:[]}),
          new Paragraph({spacing:{before:200},children:[new TextRun({text:refs,bold:true,size:22,font:"Arial"})]}),
          new Paragraph({children:[new TextRun({text:"Educateurs spécialisés",italics:true,size:20,font:"Arial"})]}),
        ]
      }]
    });
    console.log("Doc created, calling Packer.toBlob...");const blob=await Packer.toBlob(doc);console.log("Blob created:",blob);
    return{blob,fileName};
  };

  const handleDownload=async(jeune)=>{
    try{
      const{blob,fileName}=await generateDocx(jeune);
      const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=fileName+".docx";document.body.appendChild(a);a.click();document.body.removeChild(a);setTimeout(()=>URL.revokeObjectURL(url),5000);
    }catch(err){console.error("Word error:",err);alert("Erreur Word: "+err.message);}
  };

  const handlePrint=async(jeune)=>{
    setSending(true);setSent(false);
    try{
      const{blob,fileName}=await generateDocx(jeune);
      var dlUrl=URL.createObjectURL(blob);var dlA=document.createElement("a");dlA.href=dlUrl;dlA.download=fileName+".docx";document.body.appendChild(dlA);dlA.click();document.body.removeChild(dlA);setTimeout(function(){URL.revokeObjectURL(dlUrl)},5000);
      setSent(true);
      window.open("mailto:lmarcille1962@gmail.com,jeanpierregardenatpdsr@gmail.com?subject="+encodeURIComponent(fileName)+"&body="+encodeURIComponent("Bonjour,\n\nVeuillez trouver ci-joint le rapport hebdomadaire "+fileName+".\n\nCordialement,\nAssociation PDSR"),"_blank");
    }catch(e){alert("Erreur: "+e.message);}
    setSending(false);
  };

  const handlePrintAll=async()=>{
    setSending(true);
    for(const j of siteJeunes){
      try{
        const{blob,fileName}=await generateDocx(j);
        var dlUrl2=URL.createObjectURL(blob);var dlA2=document.createElement("a");dlA2.href=dlUrl2;dlA2.download=fileName+".docx";document.body.appendChild(dlA2);dlA2.click();document.body.removeChild(dlA2);setTimeout(function(){URL.revokeObjectURL(dlUrl2)},5000);
      }catch(e){console.error(e);}
    }
    setSending(false);
    window.open("mailto:lmarcille1962@gmail.com,jeanpierregardenatpdsr@gmail.com?subject="+encodeURIComponent("Rapports hebdo "+site+" S"+weekNum)+"&body="+encodeURIComponent("Bonjour,\n\nVeuillez trouver ci-joints les rapports hebdomadaires de "+site+" semaine S"+weekNum+".\n\nCordialement,\nAssociation PDSR"),"_blank");
  };

  return(<div style={{padding:"1rem"}}>
    <h2 style={{fontSize:"1.3rem",fontWeight:700,marginBottom:"1rem",color:C.gold}}>Rapports Hebdomadaires</h2>
    <div style={{display:"flex",gap:"1rem",marginBottom:"1rem",flexWrap:"wrap"}}>
      <div><label style={{fontWeight:600}}>Site : </label><select value={site} onChange={e=>{setSite(e.target.value);setSelJeune("");setWeekNum(calcSiteWeek(e.target.value));}} style={{padding:"0.4rem",borderRadius:6,border:"1px solid #ccc"}}><option>Djilass</option><option>Fatick</option></select></div>
      <div><label style={{fontWeight:600}}>Semaine : </label><input type="text" value={weekNum} onChange={e=>setWeekNum(e.target.value.replace(/[^0-9]/g,"").substring(0,2))} style={{width:60,padding:"0.4rem",borderRadius:6,border:"1px solid #ccc"}} placeholder="01"/>{(()=>{if(!weekNum||!parseInt(weekNum,10))return null;const{start,end}=weekRange(weekNum,site);return<span style={{fontSize:12,color:C.light,marginLeft:8}}>du {fmt(start)} au {fmt(end)}</span>;})()}</div>
    <button onClick={compileWeekRapports} style={{background:C.goldDark,color:"#fff",border:"none",borderRadius:6,padding:"0.4rem 1rem",cursor:"pointer",fontWeight:600,fontSize:"0.85rem",display:"flex",alignItems:"center",gap:"0.3rem",flexWrap:"wrap"}} title="Compiler automatiquement les rapports journaliers de cette semaine"><FileText size={14}/> Compiler la semaine</button>
    {(user.role==="directeur"||user.role==="chef_service"||user.role==="coordinateur_site")&&<button onClick={syntheseIA} disabled={iaLoading} style={{background:iaLoading?"#9E9E9E":C.primary,color:"#fff",border:"none",borderRadius:6,padding:"0.4rem 1rem",cursor:iaLoading?"wait":"pointer",fontWeight:600,fontSize:"0.85rem",display:"flex",alignItems:"center",gap:"0.3rem",flexWrap:"wrap"}} title="Synthétiser par IA les rapports journaliers et événements de la semaine en un texte hebdomadaire par jeune"><BarChart2 size={14}/> {iaLoading?("Synthèse... "+iaProgress):"Synthèse IA"}</button>}
    </div>

    {(user.role!=="educateur"||user.role==="coordinateur_site")&&<div style={{background:"#f9f9f9",border:"1px solid #ddd",borderRadius:8,padding:"1rem",marginBottom:"1rem"}}>
      <h3 style={{fontWeight:600,marginBottom:"0.5rem",color:C.goldDark}}>Partie Groupe ({site})</h3>
      <p style={{fontSize:"0.85rem",color:"#666",marginBottom:"0.5rem"}}>Ce texte sera identique pour tous les jeunes de {site}</p>
      <textarea value={groupText} onChange={e=>setGroupText(e.target.value)} onBlur={saveHebdoData} rows={6} style={{width:"100%",padding:"0.5rem",borderRadius:6,border:"1px solid #ccc",fontFamily:"Arial",fontSize:"0.9rem"}} placeholder="Cette semaine, le groupe a..."/>
    </div>}

    <h3 style={{fontWeight:600,marginBottom:"0.5rem",color:C.goldDark}}>Parties individuelles</h3>
    {siteJeunes.map(j=><div key={j.id} style={{background:selJeune===String(j.id)?"#fff8e1":"#fff",border:"1px solid "+(selJeune===String(j.id)?C.gold:"#ddd"),borderRadius:8,padding:"0.8rem",marginBottom:"0.5rem",cursor:"pointer"}} onClick={()=>setSelJeune(String(j.id))}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap"}}>
        <span style={{fontWeight:600,display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>{j.prenom} {j.nom} <span style={{fontSize:"0.8rem",color:"#888"}}>({getFileName(j)})</span>{(()=>{const st=getStatut(j.id);const b=ST_BADGE[st.statut]||ST_BADGE.brouillon;return<span title={st.par?(b.label+" par "+st.par+" le "+st.le):b.label} style={{background:b.bg,color:b.c,borderRadius:12,padding:"2px 10px",fontSize:11.5,fontWeight:800}}>{b.label}</span>;})()}</span>
        <div style={{display:"flex",gap:"0.5rem",flexWrap:"wrap"}}>
          <button onClick={e=>{e.stopPropagation();handleDownload(j)}} style={{background:C.gold,color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.8rem"}} title="Télécharger Word"><Download size={14}/> Word</button>
          <button onClick={async e=>{e.stopPropagation();try{const{blob:b2,fileName:fn2}=await generateDocx(j);const url2=URL.createObjectURL(b2);const w=window.open(url2);if(!w){const a=document.createElement("a");a.href=url2;a.download=fn2+".docx";a.click();}setTimeout(()=>URL.revokeObjectURL(url2),10000);}catch(err){alert("Erreur: "+err.message)}}} style={{background:"#1565C0",color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.8rem"}} title="Ouvrir pour impression"><FileText size={14}/> PDF</button>
          <button onClick={e=>{e.stopPropagation();handlePrint(j)}} style={{background:C.orange,color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.8rem"}} title="Imprimer et envoyer"><Printer size={14}/> Imprimer</button>{canValider&&(()=>{const st=getStatut(j.id).statut;return(<>{st==="brouillon"&&<button onClick={e=>{e.stopPropagation();setStatut(j.id,"valide");}} style={{background:"#2E7D32",color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.8rem",fontWeight:700}} title="Valider le rapport hebdomadaire">✓ Valider</button>}{st==="valide"&&<button onClick={e=>{e.stopPropagation();setStatut(j.id,"transmis");}} style={{background:C.goldDark,color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.8rem",fontWeight:700}} title="Marquer comme transmis au référent ASE">→ Transmis</button>}{st!=="brouillon"&&<button onClick={e=>{e.stopPropagation();if(confirm("Repasser ce rapport en brouillon ?"))setStatut(j.id,"brouillon");}} style={{background:"#ECEFF1",color:"#546E7A",border:"none",borderRadius:6,padding:"0.3rem 0.6rem",cursor:"pointer",fontSize:"0.8rem"}} title="Repasser en brouillon">↺</button>}</>);})()}
        </div>
      </div>
      {selJeune===String(j.id)&&<div style={{marginTop:"0.5rem"}}>
        <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6,flexWrap:"wrap"}}><label style={{fontSize:12,fontWeight:700,color:"#555"}}>À l'attention de :</label><input value={destTexts[j.id]||""} readOnly={getStatut(j.id).statut!=="brouillon"} onChange={e=>setDestTexts(p=>({...p,[j.id]:e.target.value}))} onBlur={saveHebdoData} placeholder={"Auto : "+("…")+" (généré depuis l'email ASE à la compilation)"} style={{flex:1,minWidth:180,padding:"0.35rem 0.5rem",borderRadius:6,border:"1px solid #ccc",fontSize:"0.85rem",fontFamily:"inherit"}}/></div>{getStatut(j.id).statut!=="brouillon"&&<div style={{fontSize:12,color:"#2E7D32",fontWeight:700,marginBottom:4}}>🔒 Rapport {ST_BADGE[getStatut(j.id).statut].label.toLowerCase()} — repasser en brouillon pour modifier</div>}<textarea readOnly={getStatut(j.id).statut!=="brouillon"} value={persoTexts[j.id]||""} onChange={e=>setPersoTexts(p=>({...p,[j.id]:e.target.value}))} onBlur={saveHebdoData} rows={5} style={{width:"100%",padding:"0.5rem",borderRadius:6,border:"1px solid #ccc",fontFamily:"Arial",fontSize:"0.9rem"}} placeholder={"La semaine de "+j.prenom+"..."}/>
        {preview&&<div style={{marginTop:"0.5rem",background:C.white,border:"1px solid #ddd",borderRadius:6,padding:"1rem"}}>
          <div style={{textAlign:"center",fontWeight:700,fontSize:"1.1rem",borderBottom:"2px solid "+C.gold,paddingBottom:"0.5rem",marginBottom:"1rem"}}>
            <div>Association PDSR</div><div>Rapport Hebdomadaire</div><div>{j.prenom} {j.nom}</div><div>Semaine S{weekNum} - Groupe de {site.toUpperCase()}</div>
          </div>
          <p style={{fontStyle:"italic",marginBottom:"1rem"}}>A l’attention de Mme Eynac Céline, Mr Bossu Sylvain et Mme Souchon Sylvia</p>
          <h4 style={{fontWeight:700,textDecoration:"underline",marginBottom:"0.5rem"}}>Cette semaine sur le groupe :</h4>
          <p style={{whiteSpace:"pre-wrap",marginBottom:"1rem"}}>{groupText||"(non renseigné)"}</p>
          <h4 style={{fontWeight:700,textDecoration:"underline",marginBottom:"0.5rem"}}>La semaine de {j.nom} {j.prenom} :</h4>
          <p style={{whiteSpace:"pre-wrap",marginBottom:"1rem"}}>{persoTexts[j.id]||"(non renseigné)"}</p>
          <div style={{borderTop:"1px solid #ddd",paddingTop:"0.5rem",textAlign:"center",fontSize:"0.85rem",color:"#666"}}>
            <div style={{fontWeight:600}}>{[refA(j.id),refB(j.id)].filter(Boolean).join(" – ")}</div>
            <div style={{fontStyle:"italic"}}>Éducateurs spécialisés</div>
            <div>Association PDSR - associationpdsr@gmail.com</div>
          </div>
        </div>}
        <div style={{display:"flex",gap:"0.5rem",marginTop:"0.5rem",flexWrap:"wrap"}}>
          <button onClick={()=>setPreview(!preview)} style={{background:preview?"#999":C.goldDark,color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.85rem"}}>{preview?"Fermer aperçu":"Prévisualiser"}</button>
        </div>
      </div>}
    </div>)}

    <div style={{display:"flex",gap:"1rem",marginTop:"1rem",borderTop:"1px solid #ddd",paddingTop:"1rem",flexWrap:"wrap"}}>
      <button onClick={handlePrintAll} disabled={sending} style={{background:C.orange,color:"#fff",border:"none",borderRadius:8,padding:"0.6rem 1.5rem",cursor:"pointer",fontWeight:600,opacity:sending?0.6:1}}><Printer size={16}/> {sending?"Génération...":"Imprimer tous les rapports ("+site+")"}</button>
    </div>
    {sent&&<p style={{color:"green",marginTop:"0.5rem"}}>Rapports générés avec succès !</p>}
  </div>);
}

function ProjetsPersonnalises({user,jeunes,majeurs,projets,onUpdate,etabConfig,users,fixedId}){
  const allPool=[...(jeunes||JEUNES),...(majeurs||MAJEURS)];
  const vj=(user.role==="educateur"||user.role==="coordinateur_site")?allPool.filter(j=>(user.site==="Tous"||j.site===user.site)&&(user.isEducMajeur?j.id>=100:j.id<100)):allPool;
  const[siteF,setSiteF]=useState("Tous");
  const[selId,setSelId]=useState(fixedId!=null?String(fixedId):"");
  const[oDom,setODom]=useState("");const[oTitre,setOTitre]=useState("");const[oMoyens,setOMoyens]=useState("");const[oRef,setORef]=useState("");const[oInd,setOInd]=useState("");const[oEch,setOEch]=useState("");const[showObj,setShowObj]=useState(false);
  const[obsTxt,setObsTxt]=useState("");const[obsDom,setObsDom]=useState("");
  const[revType,setRevType]=useState("Révision mi-séjour");const[revNote,setRevNote]=useState("");const[showRev,setShowRev]=useState(false);
  const vis=vj.filter(j=>(siteF==="Tous"||j.site===siteF)&&j.statut!=="inactif"&&j.statut!=="archivé");
  const sel=allPool.find(j=>String(j.id)===String(selId));
  const isMaj=!!sel&&sel.id>=100;
  const DOM=isMaj?PROJ_DOM_MAJEUR:PROJ_DOM_MINEUR;
  const domLabel=(k)=>{const d=[...PROJ_DOM_MINEUR,...PROJ_DOM_MAJEUR].find(x=>x.k===k);return d?d.l:"Non rattaché";};
  const projet=(projets||[]).find(p=>String(p.jeuneId)===String(selId));
  const educs=(users||USERS).filter(u=>u.role==="educateur"||u.role==="coordinateur_site");
  const upd=(patch)=>{if(projet){onUpdate(prev=>prev.map(p=>p.id===projet.id?{...p,...patch}:p));}else{onUpdate(prev=>[...(prev||[]),{id:Date.now(),jeuneId:+selId,dateElaboration:isoToday(),creePar:user?.name||"?",dipcRemisLe:"",participation:{dateEntretien:"",avisJeune:"",avisTitulaires:"",refus:false,motifRefus:""},objectifs:[],revisions:[],bilan:{date:"",par:"",texte:""},...patch}]);}};
  const estDir=user.role==="directeur"||user.role==="chef_service";
  const verrou=!!projet&&projet.statut==="valide"&&!estDir;
  const valider=()=>{if(!confirm("Valider ce projet ? Les éducateurs ne pourront plus le modifier."))return;upd({statut:"valide",valideLe:isoToday(),validePar:user?.name||"?"});};
  const rouvrir=()=>{const m=prompt("Motif de la réouverture")||"";upd({statut:"brouillon",reouvertures:[...((projet&&projet.reouvertures)||[]),{date:isoToday(),par:user?.name||"?",motif:m}]});};
  const addObs=()=>{if(!obsTxt.trim())return;upd({observations:[...((projet&&projet.observations)||[]),{id:Date.now(),date:isoToday(),par:user?.name||"?",domaine:obsDom,texte:obsTxt.trim()}]});setObsTxt("");setObsDom("");};
  const delObs=(oid)=>{if(confirm("Supprimer cette observation ?"))upd({observations:((projet&&projet.observations)||[]).filter(o=>o.id!==oid)});};
  const part=(projet&&projet.participation)||{};
  const updPart=(k,v)=>upd({participation:{...part,[k]:v}});
  const addObj=()=>{if(!oTitre.trim()){alert("Le libellé de l'objectif est obligatoire.");return;}if(!oDom){alert("Rattachez l'objectif à un domaine de vie.");return;}upd({objectifs:[...((projet&&projet.objectifs)||[]),{id:Date.now(),domaine:oDom,titre:oTitre.trim(),moyens:oMoyens.trim(),referent:oRef,indicateur:oInd.trim(),echeance:oEch,statut:"En cours",creeLe:isoToday(),creePar:user?.name||"?"}]});setODom("");setOTitre("");setOMoyens("");setORef("");setOInd("");setOEch("");setShowObj(false);};
  const updObj=(oid,patch)=>upd({objectifs:((projet&&projet.objectifs)||[]).map(o=>o.id===oid?{...o,...patch}:o)});
  const delObj=(oid)=>{if(confirm("Supprimer cet objectif ?"))upd({objectifs:((projet&&projet.objectifs)||[]).filter(o=>o.id!==oid)});};
  const addRev=()=>{if(!revNote.trim()){alert("Le contenu de la révision est obligatoire.");return;}upd({revisions:[...((projet&&projet.revisions)||[]),{id:Date.now(),date:isoToday(),par:user?.name||"?",type:revType,note:revNote.trim()}]});setRevNote("");setShowRev(false);};
  const OBJ_ST={"En cours":{bg:"#E3F2FD",c:"#1565C0"},"Atteint":{bg:"#E8F5E9",c:"#2E7D32"},"Partiellement atteint":{bg:"#FFF8E1",c:"#B8860B"},"Abandonné":{bg:"#FFEBEE",c:"#C62828"}};
  const ech=sel?projEcheances(sel,projet,etabConfig):[];
  const partTracee=!!(part.dateEntretien&&((part.avisJeune||"").trim()||part.refus));
  return(<div style={{padding:fixedId?"0":"18px 14px",maxWidth:760,margin:"0 auto"}}>
    {!fixedId&&user.role!=="educateur"&&<div style={{display:"flex",gap:7,marginBottom:14,flexWrap:"wrap"}}>{["Tous","Fatick","Djilass"].map(s=><button key={s} onClick={()=>{setSiteF(s);setSelId("");}} style={{padding:"6px 14px",borderRadius:20,border:`1.5px solid ${siteF===s?C.gold:C.border}`,background:siteF===s?C.gold:C.white,color:siteF===s?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{s}</button>)}</div>}
    {!fixedId&&<div style={{...S.card,marginBottom:14}}>
      <label style={{...S.lbl}}>Bénéficiaire</label>
      <select style={{...S.inp}} value={selId} onChange={e=>setSelId(e.target.value)}><option value="">-- Sélectionner --</option>{vis.map(j=>{const p=(projets||[]).find(x=>String(x.jeuneId)===String(j.id));const late=p?projEcheances(j,p,etabConfig).some(x=>!x.fait&&x.due&&x.due<isoToday()):false;return<option key={j.id} value={j.id}>{j.prenom} {j.nom} — {j.site}{!p?" (aucun projet)":(late?" ⚠ échéance dépassée":" ✓")}</option>;})}</select>
    </div>}
    {sel&&<div style={{...S.card,borderLeft:"4px solid "+C.gold}}>
      <div style={{marginBottom:10}}>
        <div style={{fontWeight:900,fontSize:16,color:C.dark}}>Projet personnalisé — {sel.prenom} {sel.nom}</div>
        <div style={{fontSize:12,color:C.light,marginTop:2}}>{isMaj?"Jeune majeur — contrat jeune majeur (art. L.222-5 CASF)":"Mineur — articulation avec le projet pour l'enfant (art. L.223-1-1 CASF)"}{projet?" · Élaboré le "+projet.dateElaboration+" par "+(projet.creePar||"?"):""}</div>
        {!normDate(sel.dateDebut)&&<div style={{fontSize:12,color:"#C62828",fontWeight:700,marginTop:4}}>⚠ Date d'entrée non renseignée : les échéances ne peuvent pas être calculées.</div>}
      </div>
      {!projet&&<button onClick={()=>upd({})} style={{...S.btnP,justifyContent:"center"}}><Plus size={14}/>Créer le projet personnalisé</button>}
      {projet&&<>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:12,padding:"9px 11px",borderRadius:9,background:projet.statut==="valide"?"#E8F5E9":"#FFF8E1"}}>
          <div style={{flex:1,minWidth:150,fontSize:11.5,fontWeight:800,color:projet.statut==="valide"?"#2E7D32":"#B8860B"}}>{projet.statut==="valide"?("Validé le "+fmt(projet.valideLe)+" par "+(projet.validePar||"?")):"Version de travail — non validée"}{verrou?" · lecture seule":""}</div>
          {estDir&&projet.statut!=="valide"&&<button onClick={valider} style={{...S.btnP,fontSize:11.5,padding:"6px 13px"}}><Check size={13}/>Valider</button>}
          {estDir&&projet.statut==="valide"&&<button onClick={rouvrir} style={{...S.btnO,fontSize:11.5,padding:"6px 13px"}}>Rouvrir</button>}
          <button onClick={async()=>{const nEd=((projet.editions||[]).length)+1;try{await projetPDF(projet,sel,etabConfig,user,nEd);upd({editions:[...((projet.editions)||[]),{n:nEd,date:isoToday(),par:user?.name||"?",statut:projet.statut||"brouillon"}]});}catch(err){alert("PDF impossible : "+(err&&err.message?err.message:err));}}} style={{...S.btnO,fontSize:11.5,padding:"6px 13px"}}><Download size={13}/>PDF</button>
        </div>
        {(projet.editions||[]).length>0&&<div style={{fontSize:11.5,color:C.light,marginBottom:10}}>Éditions PDF : {(projet.editions||[]).slice(-4).map(e=>"n° "+e.n+" le "+fmt(e.date)+" ("+(e.statut==="valide"?"validé":"travail")+")").join("  ·  ")}</div>}
        {(projet.reouvertures||[]).length>0&&<div style={{fontSize:11.5,color:C.light,marginBottom:10}}>{(projet.reouvertures||[]).map((r,i)=><div key={i}>Rouvert le {fmt(r.date)} par {r.par}{r.motif?" — "+r.motif:""}</div>)}</div>}
        {!verrou&&<div style={{...S.card,background:C.sableLight,marginBottom:14}}>
          <div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:2}}>Observation express</div>
          <div style={{fontSize:11.5,color:C.light,marginBottom:8}}>Une ligne suffit. C'est ce que vous voyez au quotidien qui nourrit le projet.</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>{DOM.map(d=><button key={d.k} onClick={()=>setObsDom(obsDom===d.k?"":d.k)} style={{padding:"5px 11px",borderRadius:15,border:"1.5px solid "+(obsDom===d.k?C.gold:C.border),background:obsDom===d.k?C.gold:C.white,color:obsDom===d.k?C.white:C.mid,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{d.l}</button>)}</div>
          <textarea value={obsTxt} onChange={e=>setObsTxt(e.target.value)} placeholder="Ce que j'ai observé aujourd'hui…" style={{...S.inp,minHeight:56,resize:"vertical"}}/>
          <button onClick={addObs} style={{...S.btnP,width:"100%",justifyContent:"center",marginTop:8}}><Plus size={14}/>Enregistrer l'observation</button>
        </div>}
        {(projet.observations||[]).length>0&&<div style={{marginBottom:14}}>
          <div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:6}}>Observations ({(projet.observations||[]).length})</div>
          {(projet.observations||[]).slice().sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0,12).map(o=>(<div key={o.id} style={{padding:"8px 10px",background:"#f8f9fa",borderRadius:8,marginBottom:6}}>
            <div style={{display:"flex",justifyContent:"space-between",gap:6,flexWrap:"wrap"}}><div style={{fontSize:11.5,fontWeight:700,color:C.light}}>{fmt(o.date)} — {o.par}{o.domaine&&DOM.find(d=>d.k===o.domaine)?" · "+DOM.find(d=>d.k===o.domaine).l:""}</div>{!verrou&&<button onClick={()=>delObs(o.id)} style={{background:"none",border:"none",color:"#C62828",cursor:"pointer",fontWeight:800,fontSize:12}}>✕</button>}</div>
            <div style={{fontSize:12.5,color:C.dark,marginTop:3,whiteSpace:"pre-wrap"}}>{o.texte}</div>
          </div>))}
        </div>}
        <div style={{marginBottom:14}}>
          <div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:6}}>Échéances du séjour</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{ech.map(e=>{const st=echStatut(e.due,e.fait);return<div key={e.k} style={{flex:"1 1 46%",minWidth:150,padding:"7px 10px",borderRadius:8,background:st.bg,border:"1px solid "+st.c+"33"}}><div style={{fontSize:12,fontWeight:800,color:C.dark}}>{e.l}</div><div style={{fontSize:11.5,color:C.mid}}>{e.due?"Échéance "+fmt(e.due):"—"}</div><div style={{fontSize:11.5,fontWeight:800,color:st.c,marginTop:2}}>{st.l}</div></div>;})}</div>
        </div>
        <div style={{marginBottom:14,paddingTop:12,borderTop:"1px solid "+C.border}}>
          <div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:6}}>Document individuel de prise en charge</div>
          <label style={{...S.lbl}}>Remis au jeune / titulaires le</label>
          <input disabled={verrou} type="date" style={{...S.inp}} value={projet.dipcRemisLe||""} onChange={e=>upd({dipcRemisLe:e.target.value})}/>
          <div style={{fontSize:11.5,color:C.light,marginTop:4}}>Art. D.311 CASF : remise dans les 15 jours suivant l'admission.</div>
        </div>
        <div style={{marginBottom:14,paddingTop:12,borderTop:"1px solid "+C.border}}>
          <div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:2}}>Participation du jeune {partTracee?<span style={{fontSize:11.5,color:"#2E7D32"}}>✓ tracée</span>:<span style={{fontSize:11.5,color:"#C62828"}}>non tracée</span>}</div>
          <div style={{fontSize:11.5,color:C.light,marginBottom:8}}>Art. L.311-3 7° CASF : participation directe à la conception et à la mise en œuvre du projet.</div>
          <label style={{...S.lbl}}>Date de l'entretien de co-construction</label>
          <input disabled={verrou} type="date" style={{...S.inp,marginBottom:10}} value={part.dateEntretien||""} onChange={e=>updPart("dateEntretien",e.target.value)}/>
          <label style={{...S.lbl}}>Avis et attentes exprimés par le jeune</label>
          <textarea key={selId+"-aj"} defaultValue={part.avisJeune||""} onBlur={e=>updPart("avisJeune",e.target.value)} placeholder="Retranscrire les mots du jeune" style={{...S.inp,minHeight:80,resize:"vertical",marginBottom:10}}/>
          {!isMaj&&<><label style={{...S.lbl}}>Avis des titulaires de l'autorité parentale</label>
          <textarea key={selId+"-at"} defaultValue={part.avisTitulaires||""} onBlur={e=>updPart("avisTitulaires",e.target.value)} style={{...S.inp,minHeight:60,resize:"vertical",marginBottom:10}}/></>}
          <label style={{display:"flex",alignItems:"center",gap:7,fontSize:12,fontWeight:700,color:C.dark,cursor:"pointer",flexWrap:"wrap"}}><input type="checkbox" checked={!!part.refus} onChange={e=>updPart("refus",e.target.checked)} style={{accentColor:C.gold}}/>Le jeune a refusé de participer</label>
          {part.refus&&<><label style={{...S.lbl,marginTop:8}}>Motif du refus / démarches engagées</label><textarea key={selId+"-mr"} defaultValue={part.motifRefus||""} onBlur={e=>updPart("motifRefus",e.target.value)} style={{...S.inp,minHeight:50,resize:"vertical"}}/></>}
        </div>
        <div style={{paddingTop:12,borderTop:"1px solid "+C.border}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4,flexWrap:"wrap"}}><div style={{fontWeight:800,fontSize:13,color:C.dark}}>Objectifs ({(projet.objectifs||[]).length})</div><button disabled={verrou} onClick={()=>setShowObj(!showObj)} style={{...S.btnO,fontSize:12,padding:"5px 12px"}}><Plus size={12}/>Objectif</button></div>
          <div style={{fontSize:11.5,color:C.light,marginBottom:8}}>{isMaj?"Domaines du contrat jeune majeur (trame interne)":"Domaines de vie du référentiel PPE — art. D.223-12 CASF"}</div>
          {showObj&&<div style={{...S.card,background:C.sableLight,marginBottom:10}}>
            <label style={{...S.lbl}}>Domaine de vie</label>
            <select style={{...S.inp,marginBottom:8}} value={oDom} onChange={e=>setODom(e.target.value)}><option value="">-- Choisir --</option>{DOM.map(d=><option key={d.k} value={d.k}>{d.l}</option>)}</select>
            <label style={{...S.lbl}}>Objectif (formulé avec le jeune)</label>
            <input style={{...S.inp,marginBottom:8}} value={oTitre} onChange={e=>setOTitre(e.target.value)}/>
            <label style={{...S.lbl}}>Moyens / prestations mobilisés</label>
            <textarea style={{...S.inp,minHeight:56,resize:"vertical",marginBottom:8}} value={oMoyens} onChange={e=>setOMoyens(e.target.value)}/>
            <label style={{...S.lbl}}>Indicateur d'évaluation</label>
            <input style={{...S.inp,marginBottom:8}} value={oInd} onChange={e=>setOInd(e.target.value)} placeholder="À quoi verra-t-on que l'objectif est atteint ?"/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <div><label style={{...S.lbl}}>Professionnel référent</label><select style={{...S.inp}} value={oRef} onChange={e=>setORef(e.target.value)}><option value="">--</option>{educs.map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div>
              <div><label style={{...S.lbl}}>Échéance</label><input type="date" style={{...S.inp}} value={oEch} onChange={e=>setOEch(e.target.value)}/></div>
            </div>
            <button onClick={addObj} style={{...S.btnP,width:"100%",justifyContent:"center",marginTop:10}}><Check size={14}/>Ajouter l'objectif</button>
          </div>}
          {DOM.map(d=>{const objs=(projet.objectifs||[]).filter(o=>o.domaine===d.k);if(!objs.length)return null;return(<div key={d.k} style={{marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:800,color:C.goldDark,textTransform:"uppercase",letterSpacing:"0.04em",marginBottom:4}}>{d.l}</div>
            {objs.map(o=>{const st=OBJ_ST[o.statut]||OBJ_ST["En cours"];const late=o.echeance&&o.statut==="En cours"&&o.echeance<isoToday();return(<div key={o.id} style={{padding:"9px 11px",background:"#f8f9fa",borderRadius:8,marginBottom:6}}>
              <div style={{display:"flex",gap:8,alignItems:"flex-start",flexWrap:"wrap"}}>
                <div style={{flex:1,minWidth:150}}><div style={{fontSize:13,fontWeight:700,color:C.dark}}>{o.titre}</div>{o.echeance&&<div style={{fontSize:11.5,color:late?"#C62828":C.light,fontWeight:late?800:600}}>Échéance : {fmt(o.echeance)}{late?" — dépassée":""}</div>}</div>
                <select value={o.statut} onChange={e=>updObj(o.id,{statut:e.target.value})} style={{...S.inp,width:"auto",fontSize:12,padding:"4px 8px",background:st.bg,color:st.c,fontWeight:700,border:"1px solid "+st.c}}>{OBJ_STATUTS.map(s=><option key={s}>{s}</option>)}</select>
                <button onClick={()=>delObj(o.id)} style={{background:"none",border:"none",color:"#C62828",cursor:"pointer",fontWeight:800,fontSize:14}}>✕</button>
              </div>
              {o.moyens&&<div style={{fontSize:12,color:C.mid,marginTop:4}}><b>Moyens :</b> {o.moyens}</div>}
              {o.indicateur&&<div style={{fontSize:12,color:C.mid}}><b>Indicateur :</b> {o.indicateur}</div>}
              {o.referent&&<div style={{fontSize:11.5,color:C.light,marginTop:2}}>Référent : {o.referent}</div>}
            </div>);})}
          </div>);})}
          {(projet.objectifs||[]).filter(o=>!o.domaine).length>0&&<div style={{marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:800,color:"#C62828",textTransform:"uppercase",marginBottom:4}}>Non rattachés à un domaine</div>
            {(projet.objectifs||[]).filter(o=>!o.domaine).map(o=>(<div key={o.id} style={{padding:"9px 11px",background:"#FFF5F5",borderRadius:8,marginBottom:6,display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:140,fontSize:13,fontWeight:700,color:C.dark}}>{o.titre}</div>
              <select value="" onChange={e=>e.target.value&&updObj(o.id,{domaine:e.target.value})} style={{...S.inp,width:"auto",fontSize:12,padding:"4px 8px"}}><option value="">Rattacher à…</option>{DOM.map(d=><option key={d.k} value={d.k}>{d.l}</option>)}</select>
              <button onClick={()=>delObj(o.id)} style={{background:"none",border:"none",color:"#C62828",cursor:"pointer",fontWeight:800,fontSize:14}}>✕</button>
            </div>))}
          </div>}
          {(projet.objectifs||[]).length===0&&<div style={{fontSize:12,color:C.light,padding:"8px 0"}}>Aucun objectif défini.</div>}
        </div>
        <div style={{paddingTop:12,borderTop:"1px solid "+C.border}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,flexWrap:"wrap"}}><div style={{fontWeight:800,fontSize:13,color:C.dark}}>Révisions / avenants ({(projet.revisions||[]).length})</div><button disabled={verrou} onClick={()=>setShowRev(!showRev)} style={{...S.btnO,fontSize:12,padding:"5px 12px"}}><Plus size={12}/>Révision</button></div>
          {showRev&&<div style={{...S.card,background:C.sableLight,marginBottom:10}}>
            <label style={{...S.lbl}}>Nature</label>
            <select style={{...S.inp,marginBottom:8}} value={revType} onChange={e=>setRevType(e.target.value)}><option>Avenant initial (objectifs et prestations)</option><option>Révision mi-séjour</option><option>Révision exceptionnelle</option><option>Actualisation annuelle</option></select>
            <label style={{...S.lbl}}>Contenu</label>
            <textarea style={{...S.inp,minHeight:70,resize:"vertical"}} value={revNote} onChange={e=>setRevNote(e.target.value)}/>
            <button onClick={addRev} style={{...S.btnP,width:"100%",justifyContent:"center",marginTop:10}}><Check size={14}/>Enregistrer</button>
          </div>}
          {(projet.revisions||[]).slice().reverse().map(r=><div key={r.id} style={{padding:"8px 10px",background:C.goldLight,borderRadius:8,marginBottom:6}}><div style={{fontSize:11.5,fontWeight:700,color:C.goldDark}}>{fmt(r.date)} — {r.par}{r.type?" · "+r.type:""}</div><div style={{fontSize:12,color:C.dark,marginTop:2}}>{r.note}</div></div>)}
          {(projet.revisions||[]).length===0&&<div style={{fontSize:12,color:C.light,padding:"4px 0"}}>Aucune révision.</div>}
        </div>
        <div style={{paddingTop:12,marginTop:12,borderTop:"1px solid "+C.border}}>
          <div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:2}}>Bilan de fin de séjour</div>
          <div style={{fontSize:11.5,color:C.light,marginBottom:8}}>Alimente le rapport de situation transmis au département (art. L.223-5 CASF).</div>
          <label style={{...S.lbl}}>Date</label>
          <input type="date" style={{...S.inp,marginBottom:8}} value={(projet.bilan&&projet.bilan.date)||""} onChange={e=>upd({bilan:{...(projet.bilan||{}),date:e.target.value,par:user?.name||"?"}})}/>
          <label style={{...S.lbl}}>Synthèse</label>
          <textarea key={selId+"-bil"} defaultValue={(projet.bilan&&projet.bilan.texte)||""} onBlur={e=>upd({bilan:{...(projet.bilan||{}),texte:e.target.value,par:user?.name||"?"}})} style={{...S.inp,minHeight:100,resize:"vertical"}}/>
        </div>
      </>}
    </div>}
  </div>);
}

function Planning({djiPlan,fatPlan,site,user,onUpdate}){
  const[selSite,setSelSite]=useState(site==="Tous"?"Fatick":site);
  const plan=selSite==="Fatick"?fatPlan:djiPlan;
  const[month,setMonth]=useState(()=>{const d=new Date();return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0");});
  const[editDay,setEditDay]=useState(null);
  const[editNote,setEditNote]=useState("");
  const canEdit=user&&(user.role==="chef_service"||user.role==="directeur");
  const WD=["Dim","Lun","Mar","Mer","Jeu","Ven","Sam"];
  const y=parseInt(month.split("-")[0]),m=parseInt(month.split("-")[1]);
  const first=new Date(y,m-1,1),last=new Date(y,m,0);
  const days=[];for(let d=1;d<=last.getDate();d++){const dt=new Date(y,m-1,d);const key=dt.toISOString().slice(0,10);const info=plan[key]||{};days.push({d,wd:dt.getDay(),key,a:info.a||false,b:info.b||false,n:info.n||"",v:info.v||0});}
  const prev=()=>{let nm=m-1,ny=y;if(nm<1){nm=12;ny--;}setMonth(ny+"-"+String(nm).padStart(2,"0"));};
  const next=()=>{let nm=m+1,ny=y;if(nm>12){nm=1;ny++;}setMonth(ny+"-"+String(nm).padStart(2,"0"));};
  const MN=["","Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
  const toggleDay=(key,field)=>{if(!canEdit)return;const cur=plan[key]||{};const updated={...cur,[field]:!cur[field]};onUpdate(selSite,key,updated);};
  const toggleVac=(key)=>{if(!canEdit)return;const cur=plan[key]||{};const updated={...cur,v:cur.v?0:1};onUpdate(selSite,key,updated);};
  const saveNote=(key)=>{const cur=plan[key]||{};const updated={...cur,n:editNote};onUpdate(selSite,key,updated);setEditDay(null);setEditNote("");};
  const setFullWeek=(weekStart,a,b)=>{if(!canEdit)return;for(let i=0;i<7;i++){const dt=new Date(weekStart);dt.setDate(dt.getDate()+i);if(dt.getDay()===0)continue;const key=dt.toISOString().slice(0,10);const cur=plan[key]||{};onUpdate(selSite,key,{...cur,a,b});}};
  const getWeekMonday=(dateStr)=>{const d=new Date(dateStr);const day=d.getDay();const diff=d.getDate()-day+(day===0?-6:1);return new Date(d.setDate(diff));};
  return(<div>{site==="Tous"&&<div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}><button onClick={()=>setSelSite("Fatick")} style={{padding:"6px 16px",borderRadius:8,border:"none",background:selSite==="Fatick"?C.gold:"#eee",color:selSite==="Fatick"?"#fff":"#333",fontWeight:600,cursor:"pointer"}}>Fatick</button><button onClick={()=>setSelSite("Djilass")} style={{padding:"6px 16px",borderRadius:8,border:"none",background:selSite==="Djilass"?C.gold:"#eee",color:selSite==="Djilass"?"#fff":"#333",fontWeight:600,cursor:"pointer"}}>Djilass</button></div>}
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap"}}><button onClick={prev} style={{background:C.gold,color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:15,fontWeight:700}}>◀</button><h2 style={{fontSize:18,fontWeight:900,color:C.dark,margin:0}}>{MN[m]} {y} — Planning {selSite}</h2><button onClick={next} style={{background:C.gold,color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:15,fontWeight:700}}>▶</button></div>
    {canEdit&&<div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
      <div style={{fontSize:12,color:C.light,padding:"6px 0"}}>Cliquez sur un jour pour modifier. Raccourcis semaine :</div>
    </div>}
    <div style={{display:"flex",gap:4,marginBottom:8,fontSize:11.5,color:C.light,flexWrap:"wrap"}}>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:C.sableLight,border:"1px solid #ddd",display:"inline-block"}}/> Aucune</span>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:"#fef3c7",border:"1px solid #ddd",display:"inline-block"}}/> Éq. A</span>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:"#dbeafe",border:"1px solid #ddd",display:"inline-block"}}/> Éq. B</span>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:"#e0e7ff",border:"1px solid #ddd",display:"inline-block"}}/> A+B</span>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:"#FFF3E0",border:"1px solid #ddd",display:"inline-block"}}/> Vacances</span>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4}}>{WD.map(w=><div key={w} style={{textAlign:"center",fontWeight:700,fontSize:12,color:C.dark,padding:4}}>{w}</div>)}{Array.from({length:first.getDay()}).map((_,i)=><div key={"e"+i}/>)}{days.map(day=>{const bg=day.v?"#FFF3E0":day.a&&day.b?"#e0e7ff":day.a?"#fef3c7":day.b?"#dbeafe":"#f3f4f6";const today=day.key===new Date().toISOString().slice(0,10);const isEdit=editDay===day.key;return(<div key={day.key} onClick={()=>{if(!canEdit)return;if(editDay===day.key){setEditDay(null);}else{setEditDay(day.key);setEditNote(day.n);}}} style={{background:bg,borderRadius:8,padding:6,minHeight:70,border:today?"2px solid "+C.gold:isEdit?"2px solid "+C.accent:"1px solid #e5e7eb",position:"relative",cursor:canEdit?"pointer":"default",transition:"all 0.15s"}}>
      <div style={{fontWeight:700,fontSize:13,color:C.dark}}>{day.d}</div>
      {day.a&&<div style={{fontSize:11.5,color:C.gold,fontWeight:600}}>Éq. A</div>}
      {day.b&&<div style={{fontSize:11.5,color:"#2563eb",fontWeight:600}}>Éq. B</div>}
      {day.v?<div style={{fontSize:12,color:"#E65100",fontWeight:700}}>VACANCES</div>:null}
      {day.n&&!isEdit&&<div style={{fontSize:12,color:"#6b7280",marginTop:2}}>{day.n}</div>}
      {isEdit&&canEdit&&<div onClick={e=>e.stopPropagation()} style={{position:"absolute",top:0,left:0,right:0,bottom:0,background:"rgba(255,255,255,0.97)",borderRadius:8,padding:6,zIndex:10,display:"flex",flexDirection:"column",gap:3}}>
        <div style={{fontSize:12,fontWeight:700,color:C.dark}}>Jour {day.d}</div>
        <div style={{display:"flex",gap:3}}>
          <button onClick={()=>toggleDay(day.key,"a")} style={{fontSize:12,padding:"2px 6px",borderRadius:4,border:day.a?"1px solid "+C.gold:"1px solid #ccc",background:day.a?"#fef3c7":"#fff",color:C.dark,cursor:"pointer",fontWeight:600}}>A{day.a?" ✓":""}</button>
          <button onClick={()=>toggleDay(day.key,"b")} style={{fontSize:12,padding:"2px 6px",borderRadius:4,border:day.b?"1px solid #2563eb":"1px solid #ccc",background:day.b?"#dbeafe":"#fff",color:C.dark,cursor:"pointer",fontWeight:600}}>B{day.b?" ✓":""}</button>
          <button onClick={()=>toggleVac(day.key)} style={{fontSize:12,padding:"2px 6px",borderRadius:4,border:day.v?"1px solid #E65100":"1px solid #ccc",background:day.v?"#FFF3E0":"#fff",color:C.dark,cursor:"pointer",fontWeight:600}}>Vac{day.v?" ✓":""}</button>
        </div>
        <input value={editNote} onChange={e=>setEditNote(e.target.value)} onClick={e=>e.stopPropagation()} placeholder="Note..." style={{fontSize:12,padding:"2px 4px",border:"1px solid #ccc",borderRadius:3,width:"100%"}}/>
        <div style={{display:"flex",gap:3}}>
          <button onClick={()=>saveNote(day.key)} style={{fontSize:12,padding:"2px 6px",borderRadius:3,background:C.gold,color:"#fff",border:"none",cursor:"pointer",fontWeight:600,flex:1}}>OK</button>
          <button onClick={()=>{const mon=getWeekMonday(day.key);setFullWeek(mon,true,false);}} style={{fontSize:11,padding:"2px 4px",borderRadius:3,background:"#fef3c7",color:C.dark,border:"1px solid "+C.gold,cursor:"pointer",fontWeight:600}} title="Sem. A">Sem A</button>
          <button onClick={()=>{const mon=getWeekMonday(day.key);setFullWeek(mon,false,true);}} style={{fontSize:11,padding:"2px 4px",borderRadius:3,background:"#dbeafe",color:C.dark,border:"1px solid #2563eb",cursor:"pointer",fontWeight:600}} title="Sem. B">Sem B</button>
        </div>
      </div>}
    </div>);})}</div>
  </div>);
}

function exportIncidentsXLSX(evenements,jeunes){const rows=[["Date","Jeune","Titre","Description","Gravité","Horodatage","N° Suivi","Catégorie"]];(evenements||[]).forEach(ev=>{const j=jeunes.find(j2=>j2.id===ev.jeuneId);rows.push([ev.date,j?(j.prenom+" "+(j.nom||"")):("ID:"+ev.jeuneId),ev.titre,ev.description,ev.gravite||"normal"]);});const bom="﻿";const csv=rows.map(r=>r.map(c=>'"'+String(c).replace(/"/g,'""')+'"').join(";")).join("\n");const blob=new Blob([bom+csv],{type:"text/csv;charset=utf-8"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="incidents_pdsr_"+new Date().toISOString().slice(0,10)+".csv";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}


function ArchivesSejours({currentUser}){
  const allowed=currentUser&&(currentUser.role==="directeur"||currentUser.role==="chef_service");
  const[idx,setIdx]=useState(null);const[busy,setBusy]=useState("");const[err,setErr]=useState("");
  const load=async()=>{try{const r=await fetchTO(FB_URL+"/archivesFiles/index.json?auth="+FB_SECRET,null,15000);const v=await r.json();setIdx(v?Object.values(v).sort((a,b)=>(b.date||"").localeCompare(a.date||"")):[]);}catch(e){setErr("Chargement de la liste impossible — vérifiez la connexion");setIdx([]);}};
  useEffect(()=>{if(allowed)load();},[]);
  const onFile=async(e)=>{const f=e.target.files&&e.target.files[0];e.target.value="";if(!f)return;setErr("");
    if(f.size>5*1024*1024){setErr("Fichier trop volumineux : "+(f.size/1048576).toFixed(1)+" Mo (maximum 5 Mo)");return;}
    setBusy("Envoi de "+f.name+"…");
    try{
      const b64=await new Promise((res,rej)=>{const rd=new FileReader();rd.onload=()=>res(String(rd.result).split(",")[1]);rd.onerror=()=>rej(new Error("lecture"));rd.readAsDataURL(f);});
      const id="arc_"+Date.now();
      const meta={id,name:f.name,size:f.size,type:f.type||"application/octet-stream",date:new Date().toISOString(),by:currentUser.name||""};
      let r=await fetchTO(FB_URL+"/archivesFiles/blobs/"+id+".json?auth="+FB_SECRET,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify({...meta,data:b64})},60000);
      if(!r.ok)throw new Error("HTTP "+r.status);
      r=await fetchTO(FB_URL+"/archivesFiles/index/"+id+".json?auth="+FB_SECRET,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(meta)},15000);
      if(!r.ok)throw new Error("HTTP "+r.status);
      setBusy("");load();
    }catch(ex){setBusy("");setErr("Envoi échoué — réessayez");}
  };
  const dl=async(m)=>{setBusy("Téléchargement de "+m.name+"…");setErr("");
    try{const r=await fetchTO(FB_URL+"/archivesFiles/blobs/"+m.id+".json?auth="+FB_SECRET,null,60000);const v=await r.json();if(!v||!v.data)throw new Error("introuvable");
      const bin=atob(v.data);const arr=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);
      const blob=new Blob([arr],{type:m.type||"application/octet-stream"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=m.name;a.click();URL.revokeObjectURL(a.href);setBusy("");
    }catch(ex){setBusy("");setErr("Téléchargement impossible");}};
  const del=async(m)=>{if(!confirm("Supprimer définitivement « "+m.name+" » ?"))return;setBusy("Suppression…");
    try{await fetchTO(FB_URL+"/archivesFiles/blobs/"+m.id+".json?auth="+FB_SECRET,{method:"DELETE"},20000);await fetchTO(FB_URL+"/archivesFiles/index/"+m.id+".json?auth="+FB_SECRET,{method:"DELETE"},15000);setBusy("");load();}catch(ex){setBusy("");setErr("Suppression impossible");}};
  const fmtSize=(n)=>n>1048576?(n/1048576).toFixed(1)+" Mo":Math.round((n||0)/1024)+" Ko";
  const fmtD=(iso)=>{if(!iso)return"";const d=new Date(iso);const p=n=>String(n).padStart(2,"0");return p(d.getDate())+"/"+p(d.getMonth()+1)+"/"+d.getFullYear();};
  if(!allowed)return <div style={S.card}>Section réservée au directeur et aux chefs de service.</div>;
  return(<div style={S.card}>
    <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12,flexWrap:"wrap"}}>
      <div style={{fontWeight:800,color:C.dark,fontSize:15}}>Archives des séjours précédents</div>
      <label style={{marginLeft:"auto",...S.btnP,cursor:"pointer",padding:"8px 16px",fontSize:12.5}}>Ajouter un fichier<input type="file" accept=".xlsx,.xls,.csv,.pdf,.docx" onChange={onFile} style={{display:"none"}}/></label>
    </div>
    <div style={{fontSize:11.5,color:C.light,marginBottom:10}}>Formats Excel, CSV, PDF ou Word · 5 Mo maximum · visible uniquement par le directeur et les chefs de service.</div>
    {busy&&<div style={{fontSize:12.5,fontWeight:700,color:C.goldDark,marginBottom:8}}>{busy}</div>}
    {err&&<div style={{fontSize:12.5,fontWeight:700,color:C.danger,marginBottom:8}}>{err}</div>}
    {idx===null&&<div style={{color:C.light,fontSize:13}}>Chargement…</div>}
    {idx&&idx.length===0&&!err&&<div style={{color:C.light,fontSize:13}}>Aucune archive. Ajoutez le premier fichier pour démarrer.</div>}
    {idx&&idx.map(m=>(
      <div key={m.id} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:"1px solid "+C.border,flexWrap:"wrap"}}>
        <div style={{flex:1,minWidth:160}}>
          <div style={{fontWeight:800,fontSize:13,color:C.dark,wordBreak:"break-all"}}>{m.name}</div>
          <div style={{fontSize:12,color:C.light,fontWeight:600}}>{fmtSize(m.size)} · déposé le {fmtD(m.date)} par {m.by}</div>
        </div>
        <button onClick={()=>dl(m)} style={{padding:"6px 13px",borderRadius:8,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontWeight:800,fontSize:11.5,cursor:"pointer",fontFamily:"inherit"}}>Télécharger</button>
        <button onClick={()=>del(m)} style={{padding:"6px 13px",borderRadius:8,border:"1px solid "+C.danger,background:"transparent",color:C.danger,fontWeight:800,fontSize:11.5,cursor:"pointer",fontFamily:"inherit"}}>Supprimer</button>
      </div>))}
  </div>);
}

function Admin({djiPlan,fatPlan,onBulkPlan,users,jeunes,onUpdateUsers,onUpdateJeunes,loginLogs,onRefresh,appMajeurs,onUpdateMajeurs,deletionLogs,onPurgeLogs,onPurgeDeletionLogs,onResetGlobal,rapports,evenements,sejourConfig,onUpdateSejours,presences,onChangeP,agenda,onUpdateAgenda,projets,rapportsSite,onUpdateRapportsSite,onDeleteRapport,onUpdateRapport,onDeleteEvenement,onUpdateEvenements,currentUser,isAdmin,onViewAs,onForcePush,onForcePull,onCheckIntegrity,onBackup,onRestore,etabConfig,onUpdateEtab,onArchiveSejour}){
  const[tab,setTab]=useState("educs");const[planSite,setPlanSite]=useState("Fatick");const[planD1,setPlanD1]=useState("");const[planD2,setPlanD2]=useState("");const[planMode,setPlanMode]=useState("rotation");const[planBascule,setPlanBascule]=useState("0");const[planDebut,setPlanDebut]=useState("a");const[planEcraser,setPlanEcraser]=useState(false);
const[entSel,setEntSel]=useState("");const[entOpen,setEntOpen]=useState(null);const[logTab,setLogTab]=useState("connexions");
  const[opFilter,setOpFilter]=useState("");const[opSite,setOpSite]=useState("Tous");const[editRap,setEditRap]=useState(null);const[editRapText,setEditRapText]=useState("");const[editRapDate,setEditRapDate]=useState("");const[editRapJeune,setEditRapJeune]=useState("");
  const[integrity,setIntegrity]=useState(null);const[viewAsId,setViewAsId]=useState("");
  const[statSite,setStatSite]=useState("Tous");const[fiche360Id,setFiche360Id]=useState("");const[clotureLabel,setClotureLabel]=useState("");
  const[srMode,setSrMode]=useState("semaine");const[srWeek,setSrWeek]=useState("");const[srJeune,setSrJeune]=useState("");const[srSite,setSrSite]=useState("Tous");
  const ec=etabConfig||{};
  const[newPrenom,setNewPrenom]=useState("");
  const[newNom,setNewNom]=useState("");
  const[newSite,setNewSite]=useState("Fatick");
  const[newTelP1,setNewTelP1]=useState("");
  const[newTelJ,setNewTelJ]=useState("");
  const[newEmailASE,setNewEmailASE]=useState("");
  const[newDateD,setNewDateD]=useState("");
  const[newDateF,setNewDateF]=useState("");
  const[showAddJeune,setShowAddJeune]=useState(false);
  const[showAddMajeur,setShowAddMajeur]=useState(false);
  const[newMajPrenom,setNewMajPrenom]=useState("");const[newMajNom,setNewMajNom]=useState("");const[newMajSite,setNewMajSite]=useState("Fatick");const[newMajTelP,setNewMajTelP]=useState("");const[newMajTelJ,setNewMajTelJ]=useState("");const[newMajEmail,setNewMajEmail]=useState("");const[newMajDateD,setNewMajDateD]=useState("");const[newMajDateF,setNewMajDateF]=useState("");
  const[newRole,setNewRole]=useState("educateur");
 const[newType,setNewType]=useState("jour");
 const[newSection,setNewSection]=useState("mineurs");
  const[editJeune,setEditJeune]=useState(null);
  const educs=users.filter(u=>u.role==="educateur");
  const genLogin=(nom,prenom)=>(nom+prenom).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^a-z]/g,"");
  const addEduc=()=>{if(!newPrenom.trim()||!newNom.trim())return;const login=genLogin(newNom,newPrenom);const id=Math.max(...users.map(u=>u.id))+1;const isEduc=newRole==="educateur"||newRole==="coordinateur_site";onUpdateUsers([...users,{id,login,password:login+"2026",role:newRole,name:(newPrenom+(newNom?" "+newNom:"")).trim(),site:isEduc?newSite:"Tous",type:newType,section:newSection,isEducMajeur:isEduc&&newSection==="majeurs",initials:(newPrenom.substring(0,1)+(newNom?newNom.substring(0,1):newPrenom.substring(1,2))).toUpperCase(),assignedIds:[],equipe:isEduc?"A":undefined}]);setNewPrenom("");setNewNom("");setNewType("jour");setNewSection("mineurs");setNewRole("educateur");};
  const resetPwd=(u)=>{const np=(u.login||"compte")+"2026";onUpdateUsers(users.map(x=>x.id===u.id?{...x,password:np}:x));};
  const setUserField=(id,field,val)=>onUpdateUsers(users.map(x=>x.id===id?{...x,[field]:val}:x));
  const roleLabel=(r)=>r==="directeur"?"Directeur":r==="chef_service"?"Chef de service":r==="coordinateur_site"?"Coordinateur":"Éducateur";
  // Domaine 3 — données opérationnelles
  const opPool=[...(jeunes||[]),...(appMajeurs||MAJEURS)];
  const opName=(id)=>{const j=opPool.find(x=>String(x.id)===String(id));return j?(j.prenom+" "+(j.nom||"")):("ID:"+id);};
  const opSiteOf=(id)=>{const j=opPool.find(x=>String(x.id)===String(id));return j?j.site:"";};
  const matchOp=(id,extra)=>{const nm=opName(id).toLowerCase();const st=opSite==="Tous"||opSiteOf(id)===opSite;const q=!opFilter||nm.includes(opFilter.toLowerCase())||(extra||"").toLowerCase().includes(opFilter.toLowerCase());return st&&q;};
  const SITES=["Fatick","Djilass"];
  const STATUTS=["actif","sorti","archivé"];
  const removeEduc=(id)=>{if(!confirm("Supprimer cet éducateur ?"))return;onUpdateUsers(users.filter(u=>u.id!==id));const updated=jeunes.map(j=>j.educateurId===id?{...j,educateurId:null}:j);onUpdateJeunes(updated);};
  const toggleEduc=(id)=>{onUpdateUsers(users.map(u=>u.id===id?{...u,disabled:!u.disabled}:u))};
  const addJeune=()=>{if(!newPrenom.trim())return;const id=Math.max(...jeunes.map(j=>j.id),0)+1;onUpdateJeunes([...jeunes,{id,prenom:newPrenom,nom:newNom,site:newSite,educateurId:null,referentA:"",referentB:"",referentC:"",referentD:"",statut:"actif",telParent1:newTelP1,telJeune:newTelJ,emailASE:newEmailASE,dateDebut:newDateD,dateFin:newDateF}]);setNewPrenom("");setNewNom("");setNewTelP1("");setNewTelJ("");setNewEmailASE("");setNewDateD("");setNewDateF("");setShowAddJeune(false);};
  const assignJeune=(jeuneId,educName)=>{onUpdateJeunes(jeunes.map(j=>j.id===jeuneId?{...j,referentA:educName||""}:j));const educ=users.find(u=>u.name===educName);if(educ&&!educ.assignedIds?.includes(jeuneId)){onUpdateUsers(users.map(u=>u.name===educName?{...u,assignedIds:[...(u.assignedIds||[]),jeuneId]}:u.assignedIds?.includes(jeuneId)?{...u,assignedIds:u.assignedIds.filter(i=>i!==jeuneId)}:u));}else if(!educName){onUpdateUsers(users.map(u=>u.assignedIds?.includes(jeuneId)?{...u,assignedIds:u.assignedIds.filter(i=>i!==jeuneId)}:u));}};
  const removeJeune=(id)=>{if(!confirm("Supprimer ce jeune ?"))return;onUpdateJeunes((jeunes||[]).filter(j=>j.id!==id));onUpdateUsers(users.map(u=>u.assignedIds?{...u,assignedIds:u.assignedIds.filter(i=>i!==id)}:u));};
  return(<div style={{padding:"18px 14px",maxWidth:800,margin:"0 auto"}}>
    <h2 style={{fontSize:18,fontWeight:900,color:C.dark,margin:"0 0 14px"}}>Administration</h2>
    {[{g:"Comptes & accès",items:[{k:"educs",l:"Équipe"},{k:"creds",l:"Identifiants"}]},{g:"Bénéficiaires",items:[{k:"jeunes",l:"Jeunes"},{k:"majeurs",l:"Majeurs"}]},{g:"Données opérationnelles",items:[{k:"op-rapports",l:"Rapports"},{k:"op-presences",l:"Présences"},{k:"op-incidents",l:"Incidents / EIG"},{k:"op-agenda",l:"Agenda"},{k:"op-projets",l:"Projets"},{k:"op-rsite",l:"Rapports de site"}]},{g:"Pilotage",items:[{k:"alertes",l:"Alertes / Qualité"},{k:"stats",l:"Statistiques"},{k:"suivi-rapports",l:"Suivi rapports"},{k:"fiche360",l:"Fiche 360"}]},{g:"Système",items:[{k:"config",l:"Établissement"},{k:"registre",l:"Registre L.331-2"},{k:"planning-cfg",l:"Planning"},{k:"projets-cfg",l:"Projet personnalisé"},{k:"entretiens",l:"Entretiens individuels"},{k:"sejours",l:"Séjours"},{k:"archives",l:"Archives"},{k:"logs",l:"Logs"},{k:"modifs",l:"Modifications"},...(isAdmin?[{k:"maintenance",l:"Maintenance"}]:[])]}].map(grp=>(<div key={grp.g} style={{marginBottom:10}}>
      <div style={{fontSize:12,fontWeight:800,color:C.light,letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:5}}>{grp.g}</div>
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{grp.items.map(t=><button key={t.k} onClick={()=>setTab(t.k)} style={{padding:"6px 14px",borderRadius:20,border:`1.5px solid ${tab===t.k?C.gold:C.border}`,background:tab===t.k?C.gold:C.white,color:tab===t.k?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{t.l}</button>)}</div>
    </div>))}
    <div style={{height:6}}/>
    {tab==="educs"&&<div>
      <div style={{...S.card,borderLeft:`4px solid ${C.gold}`,marginBottom:14}}>
        <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 10px",color:C.dark}}>Ajouter un compte</h3>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <div><label style={{...S.lbl}}>Prénom</label><input style={{...S.inp}} value={newPrenom} onChange={e=>setNewPrenom(e.target.value)} placeholder="Prénom"/></div>
          <div><label style={{...S.lbl}}>Nom</label><input style={{...S.inp}} value={newNom} onChange={e=>setNewNom(e.target.value)} placeholder="Nom"/></div>
        </div>
        <div style={{marginBottom:8}}><label style={{...S.lbl}}>Rôle</label><select style={{...S.inp}} value={newRole} onChange={e=>setNewRole(e.target.value)}><option value="educateur">Éducateur</option><option value="coordinateur_site">Coordinateur de site</option><option value="chef_service">Chef de service</option><option value="directeur">Directeur</option></select></div>
        {(newRole==="educateur"||newRole==="coordinateur_site")&&<><div style={{marginBottom:10}}><label style={{...S.lbl}}>Site</label><select style={{...S.inp}} value={newSite} onChange={e=>setNewSite(e.target.value)}><option>Fatick</option><option>Djilass</option></select></div>
 <div><div style={{fontSize:12,fontWeight:700,color:C.light,marginBottom:4}}>Type</div><select style={{...S.inp}} value={newType} onChange={e=>setNewType(e.target.value)}><option value="jour">Jour</option><option value="nuit">Nuit</option></select></div>
 <div style={{marginTop:6}}><div style={{fontSize:12,fontWeight:700,color:C.light,marginBottom:4}}>Section</div><select style={{...S.inp}} value={newSection} onChange={e=>setNewSection(e.target.value)}><option value="mineurs">Mineurs</option><option value="majeurs">Majeurs</option></select></div></>}
        {newPrenom&&newNom&&<div style={{fontSize:12,color:C.mid,margin:"8px 0"}}>Identifiant généré : <strong>{genLogin(newNom,newPrenom)}</strong> / Mot de passe : <strong>{genLogin(newNom,newPrenom)}2026</strong></div>}
        <button onClick={addEduc} style={{...S.btnP,width:"100%",justifyContent:"center"}}><Plus size={14}/>Ajouter le compte</button>
      </div>
      {educs.map(u=><div key={u.id} style={{...S.card,marginBottom:8,opacity:u.disabled?0.6:1,borderLeft:u.disabled?"3px solid #C62828":"3px solid transparent"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap"}}>
          <div><div style={{fontWeight:800,fontSize:14,color:C.dark}}>{u.name}</div><div style={{fontSize:12,color:C.light}}>{u.site} · {u.login} · {u.type==="nuit"?"Nuit":"Jour"} · Éq. {u.equipe||"?"} · {u.section==="majeurs"?"Majeurs":"Mineurs"} · {jeunes.filter(j=>j.referentA===u.name||j.referentB===u.name||j.referentC===u.name||j.referentD===u.name).length} jeunes</div></div>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:5,marginTop:8}}>
          <button onClick={()=>toggleEduc(u.id)} style={{padding:"4px 10px",borderRadius:6,border:u.disabled?"1px solid #2E7D32":"1px solid #E65100",background:u.disabled?"#E8F5E9":"#FFF3E0",color:u.disabled?"#2E7D32":"#E65100",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{u.disabled?"Activer":"Désactiver"}</button>
          <button onClick={()=>{const up=users.map(x=>x.id===u.id?{...x,isEducMajeur:!x.isEducMajeur,section:x.isEducMajeur?"mineurs":"majeurs"}:x);onUpdateUsers(up);}} style={{padding:"4px 10px",borderRadius:6,border:u.isEducMajeur?"1px solid #1565C0":"1px solid #9E9E9E",background:u.isEducMajeur?"#E3F2FD":"#F5F5F5",color:u.isEducMajeur?"#1565C0":"#757575",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{u.isEducMajeur?"Éduc Majeur":"Standard"}</button>
          <button onClick={()=>{const newEq=u.equipe==="A"?"B":"A";const up=users.map(x=>x.id===u.id?{...x,equipe:newEq}:x);onUpdateUsers(up);}} style={{padding:"4px 10px",borderRadius:6,border:u.equipe==="A"?"1px solid #2E7D32":"1px solid #1565C0",background:u.equipe==="A"?"#E8F5E9":"#E3F2FD",color:u.equipe==="A"?"#2E7D32":"#1565C0",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Éq. {u.equipe||"?"}</button>
          <button onClick={()=>{const up=users.map(x=>x.id===u.id?{...x,role:x.role==="coordinateur_site"?"educateur":"coordinateur_site"}:x);onUpdateUsers(up);}} style={{padding:"4px 10px",borderRadius:6,border:u.role==="coordinateur_site"?"1px solid #6A1B9A":"1px solid #9E9E9E",background:u.role==="coordinateur_site"?"#F3E5F5":"#F5F5F5",color:u.role==="coordinateur_site"?"#6A1B9A":"#757575",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{u.role==="coordinateur_site"?"Coordinateur · retirer":"Nommer coordinateur"}</button>
          <select value={u.site||"Fatick"} onChange={e=>setUserField(u.id,"site",e.target.value)} style={{padding:"4px 8px",borderRadius:6,border:"1px solid "+C.border,background:C.white,color:C.dark,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{SITES.map(s=><option key={s} value={s}>{s}</option>)}</select>
          <button onClick={()=>resetPwd(u)} style={{padding:"4px 10px",borderRadius:6,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}} title="Réinitialise le mot de passe à login+2026">Réinit. MDP</button>
          <button onClick={()=>removeEduc(u.id)} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>
        </div>
      </div>)}
      {(()=>{const staff=users.filter(u=>u.role!=="educateur");return staff.length>0&&<div style={{marginTop:18}}>
        <div style={{fontSize:12,fontWeight:800,color:C.light,letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:8}}>Comptes encadrants</div>
        {staff.map(u=><div key={u.id} style={{...S.card,marginBottom:8,padding:"12px 14px",opacity:u.disabled?0.6:1,borderLeft:"3px solid "+C.accent}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <div><div style={{fontWeight:800,fontSize:14,color:C.dark}}>{u.name}</div><div style={{fontSize:12,color:C.light}}>{roleLabel(u.role)} · {u.email||u.login}{u.site&&u.site!=="Tous"?" · "+u.site:""}</div></div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
              {u.role==="coordinateur_site"&&<button onClick={()=>{if(!confirm("Retirer le rôle de coordinateur à "+u.name+" ? Le compte redevient éducateur et conserve son historique."))return;onUpdateUsers((users||[]).map(x=>x.id===u.id?{...x,role:"educateur",updatedAt:new Date().toISOString()}:x));}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #6A1B9A",background:"#F3E5F5",color:"#6A1B9A",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Retirer coordinateur</button>}
              <button onClick={()=>toggleEduc(u.id)} style={{padding:"4px 10px",borderRadius:6,border:u.disabled?"1px solid #2E7D32":"1px solid #E65100",background:u.disabled?"#E8F5E9":"#FFF3E0",color:u.disabled?"#2E7D32":"#E65100",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{u.disabled?"Activer":"Désactiver"}</button>
              <button onClick={()=>resetPwd(u)} style={{padding:"4px 10px",borderRadius:6,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Réinit. MDP</button>
              {u.role!=="directeur"&&<button onClick={()=>{if(confirm("Supprimer le compte de "+u.name+" ?"))onUpdateUsers(users.filter(x=>x.id!==u.id));}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>}
            </div>
          </div>
        </div>)}
      </div>;})()}
    </div>}
    {tab==="jeunes"&&<div>
      <div style={{...S.card,borderLeft:`4px solid ${C.gold}`,marginBottom:14}}>
        {!showAddJeune&&<button onClick={()=>setShowAddJeune(true)} style={{...S.btnP,marginBottom:12}}><Plus size={14}/>Ajouter un jeune</button>}
        {showAddJeune&&<div style={{...S.card,borderLeft:`4px solid ${C.gold}`,marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap"}}><h3 style={{fontSize:13,fontWeight:800,margin:0,color:C.dark}}>Ajouter un jeune</h3><button onClick={()=>setShowAddJeune(false)} style={{background:"none",border:"none",cursor:"pointer",color:C.mid}}><X size={16}/></button></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
            <div><label style={{...S.lbl}}>Prénom</label><input style={{...S.inp}} value={newPrenom} onChange={e=>setNewPrenom(e.target.value)} placeholder="Prénom"/></div>
            <div><label style={{...S.lbl}}>Nom</label><input style={{...S.inp}} value={newNom} onChange={e=>setNewNom(e.target.value)} placeholder="Nom"/></div>
            <div><label style={{...S.lbl}}>Tél. parent</label><input style={{...S.inp}} value={newTelP1} onChange={e=>setNewTelP1(e.target.value)} placeholder="06 ..."/></div>
            <div><label style={{...S.lbl}}>Tél. jeune</label><input style={{...S.inp}} value={newTelJ} onChange={e=>setNewTelJ(e.target.value)} placeholder="07 ..."/></div>
            <div><label style={{...S.lbl}}>Email ASE</label><input style={{...S.inp}} value={newEmailASE} onChange={e=>setNewEmailASE(e.target.value)} placeholder="email@ase.fr"/></div>
            <div><label style={{...S.lbl}}>Site</label><select style={{...S.inp}} value={newSite} onChange={e=>setNewSite(e.target.value)}><option>Fatick</option><option>Djilass</option></select></div>
            <div><label style={{...S.lbl}}>Date début</label><input type="date" style={{...S.inp}} value={newDateD} onChange={e=>setNewDateD(e.target.value)}/></div>
            <div><label style={{...S.lbl}}>Date fin</label><input type="date" style={{...S.inp}} value={newDateF} onChange={e=>setNewDateF(e.target.value)}/></div>
          </div>
          <button onClick={addJeune} style={{...S.btnP,width:"100%",justifyContent:"center"}}><Plus size={14}/>Ajouter</button>
        </div>}
      </div>
      <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:12,flexWrap:"wrap"}}>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <label style={{...S.btnP,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6}}><Download size={14}/>Importer Excel<input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={async(e)=>{const f=e.target.files[0];if(!f)return;try{const XLSX=await loadXLSX();const wb=XLSX.read(await f.arrayBuffer(),{cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{defval:""});const res=xlParse(rows);if(!res.out.length){alert("Aucune fiche exploitable.\n\nColonnes attendues : "+XL_COLS.join(", ")+"\n\nUtilisez le bouton « Modèle Excel » pour partir du bon fichier.");e.target.value="";return;}const cur=(jeunes||[]);const exist=res.out.filter(n=>cur.some(x=>deacc(x.prenom)===deacc(n.prenom)&&deacc(x.nom)===deacc(n.nom))).map(n=>n.prenom+" "+n.nom);if(confirm(xlBilan(res,exist)+"\n\nImporter "+res.out.length+" jeune(s) ?")){const base=Math.max(0,...(jeunes||[]).map(j=>Number(j.id)||0));onUpdateJeunes([...(jeunes||[]),...res.out.map((j,i)=>({...j,id:base+1+i}))]);}}catch(err){alert("Import impossible : "+(err&&err.message?err.message:err));}e.target.value="";}}/></label>
          <button onClick={()=>xlModele("modele_import_jeunes.xlsx","Jeunes")} style={{...S.btnO,display:"inline-flex",alignItems:"center",gap:6}}><Download size={14}/>Modèle Excel</button>
          </div>
          <span style={{fontSize:11.5,color:C.mid}}>Colonnes: Prénom, Nom, Site, Tél parent, Tél jeune, Email ASE, Date début, Date fin</span>
        </div>
            {jeunes.map(j=>{return(<div key={j.id} style={{...S.card,marginBottom:8}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6,flexWrap:"wrap"}}>
          <div><div style={{fontWeight:800,fontSize:14,color:C.dark}}>{j.prenom} {j.nom}</div><div style={{fontSize:12,color:C.light}}>{j.site} · Réf: {[j.referentA,j.referentB,j.referentC,j.referentD].filter(Boolean).join(", ")||"Non assigné"}</div></div>
          <button onClick={()=>removeJeune(j.id)} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
          {["referentA","referentB","referentC","referentD"].map((rf,i)=><div key={rf} style={{display:"flex",gap:4,alignItems:"center",marginBottom:2,flexWrap:"wrap"}}><span style={{fontSize:11.5,fontWeight:700,color:C.mid,minWidth:32}}>Réf {String.fromCharCode(65+i)}</span><select style={{...S.inp,flex:1,padding:"3px 6px",fontSize:12}} value={j[rf]||""} onChange={e=>onUpdateJeunes(jeunes.map(x=>x.id===j.id?{...x,[rf]:e.target.value}:x))}><option value="">--</option>{educs.filter(u=>u.site===j.site).map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div>)}
        </div>
        <div style={{display:"flex",gap:8,alignItems:"flex-end",flexWrap:"wrap",marginTop:8,paddingTop:8,borderTop:"1px solid "+C.border}}>
          <div><div style={{fontSize:12,fontWeight:800,color:C.light,textTransform:"uppercase",marginBottom:2}}>Statut</div><select value={j.statut||"actif"} onChange={e=>onUpdateJeunes(jeunes.map(x=>x.id===j.id?{...x,statut:e.target.value}:x))} style={{padding:"4px 8px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontWeight:700,fontFamily:"inherit",color:j.statut==="actif"||!j.statut?"#2E7D32":j.statut==="sorti"?"#E65100":C.mid,background:C.white}}>{STATUTS.map(s=><option key={s} value={s}>{s}</option>)}</select></div>
          <div><div style={{fontSize:12,fontWeight:800,color:C.light,textTransform:"uppercase",marginBottom:2}}>Site</div><select value={j.site} onChange={e=>onUpdateJeunes(jeunes.map(x=>x.id===j.id?{...x,site:e.target.value}:x))} style={{padding:"4px 8px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontWeight:700,fontFamily:"inherit",color:C.dark,background:C.white}}>{SITES.map(s=><option key={s} value={s}>{s}</option>)}</select></div>
          <div><div style={{fontSize:12,fontWeight:800,color:C.light,textTransform:"uppercase",marginBottom:2}}>Entrée</div><input type="date" value={j.dateDebut||""} onChange={e=>onUpdateJeunes(jeunes.map(x=>x.id===j.id?{...x,dateDebut:e.target.value}:x))} style={{padding:"3px 6px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontFamily:"inherit"}}/></div>
          <div><div style={{fontSize:12,fontWeight:800,color:C.light,textTransform:"uppercase",marginBottom:2}}>Sortie</div><input type="date" value={j.dateFin||""} onChange={e=>onUpdateJeunes(jeunes.map(x=>x.id===j.id?{...x,dateFin:e.target.value}:x))} style={{padding:"3px 6px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontFamily:"inherit"}}/></div>
        </div>
      </div>);})}
    </div>}

    {tab==="majeurs"&&<div>
        <div style={{fontWeight:700,fontSize:15,marginBottom:12}}>Gestion des Majeurs</div>
        {!showAddMajeur&&<button onClick={()=>setShowAddMajeur(true)} style={{...S.btnP,marginBottom:12}}><Plus size={14}/>Ajouter un majeur</button>}
        {showAddMajeur&&<div style={{...S.card,borderLeft:`4px solid ${C.gold}`,marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap"}}><h3 style={{fontSize:13,fontWeight:800,margin:0,color:C.dark}}>Ajouter un majeur</h3><button onClick={()=>setShowAddMajeur(false)} style={{background:"none",border:"none",cursor:"pointer",color:C.mid}}><X size={16}/></button></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
            <div><label style={{...S.lbl}}>Prénom</label><input style={{...S.inp}} value={newMajPrenom} onChange={e=>setNewMajPrenom(e.target.value)} placeholder="Prénom"/></div>
            <div><label style={{...S.lbl}}>Nom</label><input style={{...S.inp}} value={newMajNom} onChange={e=>setNewMajNom(e.target.value)} placeholder="Nom"/></div>
            <div><label style={{...S.lbl}}>Tél. parent</label><input style={{...S.inp}} value={newMajTelP} onChange={e=>setNewMajTelP(e.target.value)} placeholder="06 ..."/></div>
            <div><label style={{...S.lbl}}>Tél. jeune</label><input style={{...S.inp}} value={newMajTelJ} onChange={e=>setNewMajTelJ(e.target.value)} placeholder="07 ..."/></div>
            <div><label style={{...S.lbl}}>Email ASE</label><input style={{...S.inp}} value={newMajEmail} onChange={e=>setNewMajEmail(e.target.value)} placeholder="email@ase.fr"/></div>
            <div><label style={{...S.lbl}}>Site</label><select style={{...S.inp}} value={newMajSite} onChange={e=>setNewMajSite(e.target.value)}><option>Fatick</option><option>Djilass</option></select></div>
            <div><label style={{...S.lbl}}>Date début</label><input type="date" style={{...S.inp}} value={newMajDateD} onChange={e=>setNewMajDateD(e.target.value)}/></div>
            <div><label style={{...S.lbl}}>Date fin</label><input type="date" style={{...S.inp}} value={newMajDateF} onChange={e=>setNewMajDateF(e.target.value)}/></div>
          </div>
          <button onClick={()=>{if(!newMajPrenom.trim())return;const id=Math.max(...(appMajeurs||MAJEURS).map(j=>j.id),99)+1;const nm=[...(appMajeurs||MAJEURS),{id,prenom:newMajPrenom,nom:newMajNom,site:newMajSite,referentA:"",referentB:"",referentC:"",referentD:"",statut:"actif",telParent1:newMajTelP,telJeune:newMajTelJ,emailASE:newMajEmail,dateDebut:newMajDateD,dateFin:newMajDateF}];onUpdateMajeurs(null,null,null,nm);setNewMajPrenom("");setNewMajNom("");setNewMajTelP("");setNewMajTelJ("");setNewMajEmail("");setNewMajDateD("");setNewMajDateF("");setShowAddMajeur(false);}} style={{...S.btnP,width:"100%",justifyContent:"center"}}><Plus size={14}/>Ajouter</button>
        </div>}
        <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:12,flexWrap:"wrap"}}>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <label style={{...S.btnP,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6}}><Download size={14}/>Importer Excel<input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={async(e)=>{const f=e.target.files[0];if(!f)return;try{const XLSX=await loadXLSX();const wb=XLSX.read(await f.arrayBuffer(),{cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws,{defval:""});const res=xlParse(rows);const cur=(appMajeurs||MAJEURS||[]);if(!res.out.length){alert("Aucune fiche exploitable.\n\nColonnes attendues : "+XL_COLS.join(", ")+"\n\nUtilisez le bouton « Modèle Excel » pour partir du bon fichier.");e.target.value="";return;}const exist=res.out.filter(n=>cur.some(x=>deacc(x.prenom)===deacc(n.prenom)&&deacc(x.nom)===deacc(n.nom))).map(n=>n.prenom+" "+n.nom);if(confirm(xlBilan(res,exist)+"\n\nImporter "+res.out.length+" majeur(s) ?")){const base=Math.max(99,...cur.map(j=>Number(j.id)||0));onUpdateMajeurs(null,null,null,[...cur,...res.out.map((j,i)=>({...j,id:base+1+i}))]);}}catch(err){alert("Import impossible : "+(err&&err.message?err.message:err));}e.target.value="";}}/></label>
          <button onClick={()=>xlModele("modele_import_majeurs.xlsx","Majeurs")} style={{...S.btnO,display:"inline-flex",alignItems:"center",gap:6}}><Download size={14}/>Modèle Excel</button>
          </div>
          <span style={{fontSize:11.5,color:C.mid}}>Colonnes: Prénom, Nom, Site, Tél parent, Tél jeune, Email ASE, Date début, Date fin</span>
        </div>
        {(appMajeurs||MAJEURS).map(m=><div key={m.id} style={{...S.card,marginBottom:8}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap"}}>
            <div><div style={{fontWeight:700,color:C.dark}}>{m.prenom} {m.nom||""}</div><div style={{fontSize:12,color:C.light}}>{m.site} | Réf: {[m.referentA,m.referentB,m.referentC,m.referentD].filter(Boolean).join(", ")||"Aucun"}</div></div>
            <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
              {["referentA","referentB","referentC","referentD"].map((rf,i)=><select key={rf} value={m[rf]||""} onChange={e=>{if(onUpdateMajeurs)onUpdateMajeurs(m.id,rf,e.target.value);}} style={{...S.inp,width:100,fontSize:12}}><option value="">{"Réf "+(i+1)}</option>{users.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select>)}
              <button onClick={()=>{if(confirm("Supprimer ce majeur ?"))onUpdateMajeurs(null,null,null,(appMajeurs||MAJEURS).filter(x=>x.id!==m.id));}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"flex-end",flexWrap:"wrap",marginTop:8,paddingTop:8,borderTop:"1px solid "+C.border}}>
            <div><div style={{fontSize:12,fontWeight:800,color:C.light,textTransform:"uppercase",marginBottom:2}}>Statut</div><select value={m.statut||"actif"} onChange={e=>onUpdateMajeurs(m.id,"statut",e.target.value)} style={{padding:"4px 8px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontWeight:700,fontFamily:"inherit",color:m.statut==="actif"||!m.statut?"#2E7D32":m.statut==="sorti"?"#E65100":C.mid,background:C.white}}>{STATUTS.map(s=><option key={s} value={s}>{s}</option>)}</select></div>
            <div><div style={{fontSize:12,fontWeight:800,color:C.light,textTransform:"uppercase",marginBottom:2}}>Site</div><select value={m.site} onChange={e=>onUpdateMajeurs(m.id,"site",e.target.value)} style={{padding:"4px 8px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontWeight:700,fontFamily:"inherit",color:C.dark,background:C.white}}>{SITES.map(s=><option key={s} value={s}>{s}</option>)}</select></div>
            <div><div style={{fontSize:12,fontWeight:800,color:C.light,textTransform:"uppercase",marginBottom:2}}>Entrée</div><input type="date" value={m.dateDebut||""} onChange={e=>onUpdateMajeurs(m.id,"dateDebut",e.target.value)} style={{padding:"3px 6px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontFamily:"inherit"}}/></div>
            <div><div style={{fontSize:12,fontWeight:800,color:C.light,textTransform:"uppercase",marginBottom:2}}>Sortie</div><input type="date" value={m.dateFin||""} onChange={e=>onUpdateMajeurs(m.id,"dateFin",e.target.value)} style={{padding:"3px 6px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontFamily:"inherit"}}/></div>
          </div>
        </div>)}
      </div>}  {tab==="creds"&&<div>
      <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 14px",color:C.dark}}>Identifiants de connexion</h3>
      {users.map(u=><div key={u.id} style={{...S.card,marginBottom:6,padding:"10px 14px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap"}}>
          <div><div style={{fontWeight:700,fontSize:13,color:C.dark}}>{u.name}</div><div style={{fontSize:12,color:C.light}}>{u.role==="educateur"?"Éducateur · "+u.site:u.role==="coordinateur_site"?"Coordinateur · "+u.site:u.role==="chef_service"?"Chef de service":"Directeur"}</div></div>
          <div style={{textAlign:"right"}}><div style={{fontSize:12,fontWeight:700,color:C.gold}}>{u.email||u.login}</div><div style={{display:"flex",alignItems:"center",gap:5,justifyContent:"flex-end",marginTop:3,flexWrap:"wrap"}}><span style={{fontSize:11.5,color:C.light}}>mdp</span><input value={u.password||""} onChange={e=>setUserField(u.id,"password",e.target.value)} style={{width:120,padding:"3px 7px",border:"1px solid "+C.border,borderRadius:6,fontSize:12,fontFamily:"inherit",color:C.mid}}/></div></div>
        </div>
      </div>)}
    </div>}

    {tab==="archives"&&<ArchivesSejours currentUser={currentUser}/>}
    {tab==="sejours"&&(()=>{
      const calcW=(d)=>{if(!d)return"—";const d0=new Date(d+"T00:00:00");if(isNaN(d0))return"—";const diff=Math.floor((Date.now()-d0.getTime())/86400000);if(diff<0)return"démarre le "+d;const sw=Math.floor(diff/7)+1;return"semaine "+String(sw).padStart(2,"0")+" en cours";};
      return(<div>
        <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>Paramétrage des séjours</div>
        <div style={{fontSize:12,color:C.light,marginBottom:14}}>Date de début de séjour par site (lundi de la semaine 1). Le numéro de semaine du rapport hebdomadaire, la compilation et la synthèse IA s'appuient sur ces dates. À mettre à jour à chaque nouveau séjour.</div>
        {["Fatick","Djilass"].map(s=>{const d=(sejourConfig&&sejourConfig[s]&&sejourConfig[s].dateDebut)||"";return(<div key={s} style={{...S.card,marginBottom:10,display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
          <div style={{minWidth:90,fontWeight:900,color:C.dark,fontSize:15}}>{s}</div>
          <div><label style={{...S.lbl}}>Début du séjour (semaine 1)</label><input type="date" value={d} onChange={e=>onUpdateSejours&&onUpdateSejours(s,e.target.value)} style={{...S.inp,width:"auto"}}/></div>
          <div style={{background:C.goldLight,color:C.goldDark,borderRadius:20,padding:"5px 14px",fontSize:12,fontWeight:800}}>{calcW(d)}</div>
        </div>);})}
        <div style={{fontSize:12,color:C.light,marginTop:6}}>⚠ Modifier une date en cours de séjour renumérote les semaines : les rapports hebdo déjà validés restent attachés à leur ancien numéro de semaine.</div>
      </div>);
    })()}
    {tab==="modifs"&&(()=>{
      const poolAll=[...(jeunes||[]),...(appMajeurs||[])];
      const jName=(id)=>{const j=poolAll.find(x=>x.id===id);return j?(j.prenom+" "+(j.nom||"")):"—";};
      const entries=[
        ...(rapports||[]).flatMap(r=>(r.historique||[]).map((h,i)=>({kind:"Rapport",date:h.date,par:h.par,jeune:jName(r.jeuneId),avant:h.ancienTexte,apres:(i===(r.historique||[]).length-1)?r.observation:(r.historique[i+1]||{}).ancienTexte,auteurOriginal:r.author||"?"}))),
        ...(evenements||[]).flatMap(e=>(e.historique||[]).map((h,i)=>({kind:"Événement",date:h.date,par:h.par,jeune:e.jeuneId?jName(e.jeuneId):"(équipe)",avant:(h.ancienTitre||"")+(h.ancienTexte?" — "+h.ancienTexte:""),apres:(i===(e.historique||[]).length-1)?((e.titre||"")+(e.description?" — "+e.description:"")):(((e.historique[i+1]||{}).ancienTitre||"")+((e.historique[i+1]||{}).ancienTexte?" — "+(e.historique[i+1]||{}).ancienTexte:"")),auteurOriginal:e.author||"?"})))
      ].sort((a,b)=>(b.date||"").localeCompare(a.date||""));
      return(<div>
        <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>Journal des modifications</div>
        <div style={{fontSize:12,color:C.light,marginBottom:12}}>{entries.length} modification(s) sur les rapports journaliers et événements. L'ancien contenu est conservé intégralement.</div>
        {entries.length===0&&<div style={{...S.card,textAlign:"center",color:C.light}}>Aucune modification enregistrée</div>}
        {entries.slice(0,200).map((m,i)=><div key={i} style={{...S.card,marginBottom:8,borderLeft:"4px solid "+(m.kind==="Rapport"?C.gold:C.orange)}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:6,marginBottom:6}}>
            <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}><span style={{background:m.kind==="Rapport"?C.goldLight:C.orangeLight,color:m.kind==="Rapport"?C.goldDark:C.orange,borderRadius:10,padding:"2px 10px",fontSize:11.5,fontWeight:800}}>{m.kind}</span><span style={{fontWeight:800,fontSize:13,color:C.dark}}>{m.jeune}</span></div>
            <div style={{fontSize:11.5,color:C.light}}>{(m.date||"").replace("T"," à ").slice(0,19)} — modifié par <b>{m.par}</b> (rédigé par {m.auteurOriginal})</div>
          </div>
          <div style={{display:"grid",gap:6}}>
            <div style={{padding:"6px 9px",background:"#FFEBEE",borderRadius:7}}><div style={{fontSize:12,fontWeight:800,color:"#C62828",marginBottom:2}}>AVANT</div><div style={{fontSize:12,color:C.mid,lineHeight:1.4}}>{m.avant||"(vide)"}</div></div>
            <div style={{padding:"6px 9px",background:"rgba(102,187,106,0.15)",borderRadius:7}}><div style={{fontSize:12,fontWeight:800,color:"#2E7D32",marginBottom:2}}>APRÈS</div><div style={{fontSize:12,color:C.mid,lineHeight:1.4}}>{m.apres||"(vide)"}</div></div>
          </div>
        </div>)}
      </div>);
    })()}
    {tab==="logs"&&<div>
      <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>{[{k:"connexions",l:"Connexions"},{k:"suppressions",l:"Suppressions"}].map(t=><button key={t.k} onClick={()=>setLogTab(t.k)} style={{padding:"5px 14px",borderRadius:16,border:`1px solid ${logTab===t.k?C.gold:C.border}`,background:logTab===t.k?C.goldLight:C.white,color:logTab===t.k?C.dark:C.mid,fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{t.l} ({t.k==="connexions"?(loginLogs||[]).length:(deletionLogs||[]).length})</button>)}</div>
      {logTab==="connexions"&&<div>
        <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
          <button onClick={()=>{const bom="﻿";const rows=[["Date","Utilisateur","Rôle"],...(loginLogs||[]).map(l=>[new Date(l.date).toLocaleString("fr-FR"),l.user,l.role])];const csv=rows.map(r=>r.map(c=>'"'+String(c||"").replace(/"/g,'""')+'"').join(";")).join("\n");const blob=new Blob([bom+csv],{type:"text/csv;charset=utf-8"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="logs_connexion_"+new Date().toISOString().slice(0,10)+".csv";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}} style={{padding:"5px 12px",borderRadius:8,border:"1px solid "+C.gold,background:C.goldLight,color:C.dark,fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>Exporter CSV</button>
          <button onClick={()=>{if(onRefresh)onRefresh();}} style={{padding:"5px 12px",borderRadius:8,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>Actualiser</button>
          <button onClick={()=>{if(!currentUser||!currentUser.isAdmin){alert("Seul l'administrateur peut purger.");return;}if(confirm("Purger tous les logs de connexion ? (irréversible)"))onPurgeLogs();}} style={{padding:"5px 12px",borderRadius:8,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>Purger</button>
        </div>
        <div style={{...S.card,maxHeight:400,overflowY:"auto"}}>
          {(loginLogs||[]).length===0?<div style={{textAlign:"center",color:C.light,padding:20}}>Aucune connexion</div>:
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Date/Heure</th><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Utilisateur</th><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Rôle</th></tr></thead>
            <tbody>{[...(loginLogs||[])].sort((a,b)=>(b.date||"").localeCompare(a.date||"")).slice(0,500).map(l=><tr key={l.id}><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12}}>{new Date(l.date).toLocaleString("fr-FR")}</td><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12,fontWeight:600}}>{l.user}</td><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12}}>{l.role}</td></tr>)}</tbody>
          </table>}
        </div>
      </div>}
      {logTab==="suppressions"&&<div>
        <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
          <button onClick={()=>{const bom="﻿";const rows=[["Date suppression","Par","Type","Date item","Jeune","Détail"],...(deletionLogs||[]).map(l=>[new Date(l.date).toLocaleString("fr-FR"),l.by,l.type,l.itemDate,l.jeune,(l.detail||"")])];const csv=rows.map(r=>r.map(c=>'"'+String(c||"").replace(/"/g,'""')+'"').join(";")).join("\n");const blob=new Blob([bom+csv],{type:"text/csv;charset=utf-8"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="logs_suppressions_"+new Date().toISOString().slice(0,10)+".csv";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}} style={{padding:"5px 12px",borderRadius:8,border:"1px solid "+C.gold,background:C.goldLight,color:C.dark,fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>Exporter CSV</button>
          <button onClick={()=>{if(!currentUser||!currentUser.isAdmin){alert("Seul l'administrateur peut purger.");return;}if(confirm("Purger tous les logs de suppression ? (irréversible)"))onPurgeDeletionLogs();}} style={{padding:"5px 12px",borderRadius:8,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>Purger</button>
        </div>
        <div style={{...S.card,maxHeight:400,overflowY:"auto"}}>
          {(deletionLogs||[]).length===0?<div style={{textAlign:"center",color:C.light,padding:20}}>Aucune suppression enregistrée</div>:
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Date</th><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Par</th><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Type</th><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Jeune</th><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Détail</th></tr></thead>
            <tbody>{(deletionLogs||[]).slice(0,200).map(l=><tr key={l.id}><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12}}>{new Date(l.date).toLocaleString("fr-FR")}</td><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12,fontWeight:600}}>{l.by}</td><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12}}><span style={{padding:"2px 8px",borderRadius:10,background:l.type==="rapport"?"#E3F2FD":"#FFF3E0",fontSize:11.5,fontWeight:600}}>{l.type}</span></td><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12}}>{l.jeune}</td><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12,color:C.mid,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{l.detail}</td></tr>)}</tbody>
          </table>}
        </div>
      </div>}
    </div>}

    {tab.startsWith("op-")&&<div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
      <div style={{flex:1,minWidth:160,position:"relative"}}><Search size={15} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:C.light}}/><input style={{...S.inp,paddingLeft:35}} placeholder="Rechercher (nom, texte…)" value={opFilter} onChange={e=>setOpFilter(e.target.value)}/></div>
      <select style={{...S.inp,width:"auto"}} value={opSite} onChange={e=>setOpSite(e.target.value)}><option>Tous</option><option>Fatick</option><option>Djilass</option></select>
    </div>}

    {tab==="op-rapports"&&(()=>{const list=(rapports||[]).filter(r=>matchOp(r.jeuneId,r.observation)).sort((a,b)=>(b.date||"").localeCompare(a.date||""));return(<div>
      <div style={{fontSize:12,color:C.light,marginBottom:10}}>{list.length} rapport(s) journalier(s). Modification et suppression centralisées — toute suppression est tracée dans les logs.</div>
      {list.length===0&&<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun rapport</div>}
      {list.map(r=><div key={r.id} style={{...S.card,marginBottom:8}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5,gap:8,flexWrap:"wrap"}}><div><span style={{fontWeight:800,fontSize:13,color:C.dark}}>{opName(r.jeuneId)}</span><span style={{fontSize:12,color:C.gold,fontWeight:700,marginLeft:8}}>{fmt(r.date)}</span>{r.author&&<span style={{fontSize:11.5,color:C.light,marginLeft:6}}>par {r.author}</span>}</div><div style={{display:"flex",gap:5,flexWrap:"wrap"}}>{editRap===r.id?<><button onClick={()=>{onUpdateRapport&&onUpdateRapport(r.id,{observation:editRapText,date:editRapDate||r.date,jeuneId:editRapJeune?Number(editRapJeune):r.jeuneId});setEditRap(null);}} style={{padding:"4px 10px",borderRadius:6,border:"none",background:"#2E7D32",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>✓</button><button onClick={()=>setEditRap(null)} style={{padding:"4px 10px",borderRadius:6,border:"1px solid "+C.border,background:C.white,color:C.mid,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Annuler</button></>:<><button onClick={()=>{setEditRap(r.id);setEditRapText(r.observation||"");setEditRapDate(r.date||"");setEditRapJeune(r.jeuneId);}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Modifier</button><button onClick={()=>{if(confirm("Supprimer ce rapport ?"))onDeleteRapport&&onDeleteRapport(r.id);}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button></>}</div></div>
        {editRap===r.id?<div>{isAdmin&&<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6,flexWrap:"wrap"}}><span style={{fontSize:11.5,fontWeight:800,color:C.light,textTransform:"uppercase"}}>Jeune</span><select value={editRapJeune} onChange={e=>setEditRapJeune(e.target.value)} style={{padding:"4px 8px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontFamily:"inherit",flex:1,minWidth:160}}>{opPool.slice().sort((a,b)=>a.prenom.localeCompare(b.prenom)).map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom} · {j.site}</option>)}</select></div>}<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6,flexWrap:"wrap"}}><span style={{fontSize:11.5,fontWeight:800,color:C.light,textTransform:"uppercase"}}>Date du rapport</span><input type="date" value={editRapDate} onChange={e=>setEditRapDate(e.target.value)} style={{padding:"4px 8px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontFamily:"inherit"}}/></div><textarea value={editRapText} onChange={e=>setEditRapText(e.target.value)} rows={3} style={{...S.inp,width:"100%",resize:"vertical",fontSize:12}}/></div>:<p style={{margin:0,fontSize:12,color:C.mid,lineHeight:1.5}}>{r.observation}</p>}
      </div>)}
    </div>);})()}

    {tab==="op-presences"&&(()=>{const pool=opPool.filter(j=>matchOp(j.id)&&j.statut!=="archivé");return(<div>
      <div style={{fontSize:12,color:C.light,marginBottom:10}}>Correction des présences sur la semaine en cours. Touchez une case pour basculer Présent → Absent → Retard.</div>
      <div style={{...S.card,padding:"10px 8px",overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",minWidth:420}}>
          <thead><tr><th style={{textAlign:"left",padding:6,fontSize:11.5,color:C.light}}>Bénéficiaire</th>{WEEKDATES.map((d,i)=><th key={d} style={{padding:4,fontSize:11.5,color:C.light,textAlign:"center"}}>{WD[i]}</th>)}</tr></thead>
          <tbody>{pool.map(j=>(<tr key={j.id}><td style={{padding:6,fontSize:12,fontWeight:700,color:C.dark,whiteSpace:"nowrap"}}>{j.prenom} {j.nom}</td>{WEEKDATES.map(date=>{const p=(presences||[]).find(p2=>p2.jeuneId===j.id&&p2.date===date);const st=p?.statut||"Présent";const nx={Présent:"Absent",Absent:"Retard",Retard:"Présent"};const sc2=SC[st]||SC.Présent;return(<td key={date} style={{padding:3,textAlign:"center"}}><button onClick={()=>onChangeP&&onChangeP(j.id,date,nx[st])} style={{width:30,height:30,borderRadius:7,background:sc2.bg,border:"none",cursor:"pointer",color:sc2.text,fontWeight:800,fontSize:13}}>{sc2.icon}</button></td>);})}</tr>))}</tbody>
        </table>
        {pool.length===0&&<div style={{textAlign:"center",color:C.light,padding:16,fontSize:12}}>Aucun bénéficiaire</div>}
      </div>
    </div>);})()}

    {tab==="op-incidents"&&(()=>{const list=(evenements||[]).filter(e=>matchOp(e.jeuneId,(e.titre||"")+" "+(e.description||""))).sort((a,b)=>(b.date||"").localeCompare(a.date||""));const setEig=(id,val)=>onUpdateEvenements&&onUpdateEvenements((evenements||[]).map(x=>x.id===id?{...x,eig:val,eigData:x.eigData||{}}:x));const setEigData=(id,patch)=>onUpdateEvenements&&onUpdateEvenements((evenements||[]).map(x=>x.id===id?{...x,eigData:{...(x.eigData||{}),...patch}}:x));const setNum=(id,v)=>onUpdateEvenements&&onUpdateEvenements((evenements||[]).map(x=>x.id===id?{...x,numeroSuivi:v}:x));return(<div>
      <div style={{fontSize:12,color:C.light,marginBottom:10}}>Registre consolidé des événements indésirables. Qualifier en EIG, attribuer un n° de suivi, suivre la transmission aux autorités.</div>
      {list.length===0&&<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun incident</div>}
      {list.map(e=>{const gc=GC[e.gravite]||GC["Léger"];return(<div key={e.id} style={{...S.card,marginBottom:8,borderLeft:"4px solid "+gc.dot}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,flexWrap:"wrap"}}><div><div style={{fontWeight:800,fontSize:13,color:C.dark}}>{e.titre}</div><div style={{fontSize:12,color:C.light,display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>{isAdmin?<select value={e.jeuneId} onChange={ev=>onUpdateEvenements&&onUpdateEvenements((evenements||[]).map(x=>x.id===e.id?{...x,jeuneId:Number(ev.target.value)}:x))} style={{padding:"3px 6px",borderRadius:6,border:"1px solid "+C.border,fontSize:12,fontFamily:"inherit",maxWidth:200}}>{opPool.slice().sort((a,b)=>a.prenom.localeCompare(b.prenom)).map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom} · {j.site}</option>)}</select>:<span>{opName(e.jeuneId)}</span>}<span>· {fmt(e.date)}</span></div></div><div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}><span style={{background:gc.bg,color:gc.text,borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:700}}>{e.gravite}</span>{e.eig&&<span style={{background:"#C62828",color:"#fff",borderRadius:4,padding:"2px 8px",fontSize:11.5,fontWeight:800}}>EIG</span>}</div></div>
        <p style={{margin:"6px 0",fontSize:12,color:C.mid,lineHeight:1.5}}>{e.description}</p>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginTop:6,paddingTop:8,borderTop:"1px solid "+C.border}}>
          <div style={{display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}><span style={{fontSize:11.5,fontWeight:700,color:C.light}}>N° suivi</span><input defaultValue={e.numeroSuivi||""} onBlur={ev2=>setNum(e.id,ev2.target.value)} placeholder="—" style={{width:90,padding:"3px 6px",border:"1px solid "+C.border,borderRadius:6,fontSize:12,fontFamily:"inherit"}}/></div>
          <button onClick={()=>setEig(e.id,!e.eig)} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:e.eig?"#C62828":"#fff",color:e.eig?"#fff":"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{e.eig?"Retirer EIG":"Qualifier EIG"}</button>
          {onDeleteEvenement&&<button onClick={()=>{if(confirm("Supprimer cet événement ?"))onDeleteEvenement(e.id);}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>}
        </div>
        {e.eig&&<div style={{marginTop:8,padding:"8px 10px",background:"#FFF5F5",border:"1px solid #F3C6C6",borderRadius:8,display:"flex",gap:8,flexWrap:"wrap",alignItems:"flex-end"}}>
          <div><div style={{fontSize:12,fontWeight:800,color:"#C62828",marginBottom:2}}>Destinataires</div><input defaultValue={e.eigData?.destinataires||""} onBlur={ev2=>setEigData(e.id,{destinataires:ev2.target.value})} placeholder="CD, PJJ, parquet…" style={{padding:"3px 6px",border:"1px solid "+C.border,borderRadius:6,fontSize:12,fontFamily:"inherit",minWidth:150}}/></div>
          <div><div style={{fontSize:12,fontWeight:800,color:"#C62828",marginBottom:2}}>Transmission</div><input type="date" value={e.eigData?.dateTransmission||""} onChange={ev2=>setEigData(e.id,{dateTransmission:ev2.target.value})} style={{padding:"3px 6px",border:"1px solid "+C.border,borderRadius:6,fontSize:12,fontFamily:"inherit"}}/></div>
          <label style={{display:"flex",alignItems:"center",gap:5,fontSize:12,fontWeight:700,color:"#C62828",cursor:"pointer",flexWrap:"wrap"}}><input type="checkbox" checked={!!e.eigData?.accuseReception} onChange={ev2=>setEigData(e.id,{accuseReception:ev2.target.checked})} style={{accentColor:"#C62828"}}/>AR reçu</label>
          <div><div style={{fontSize:12,fontWeight:800,color:"#C62828",marginBottom:2}}>Statut</div><select value={e.eigData?.statutCloture||"En cours"} onChange={ev2=>setEigData(e.id,{statutCloture:ev2.target.value})} style={{padding:"3px 6px",border:"1px solid "+C.border,borderRadius:6,fontSize:12,fontFamily:"inherit"}}><option>En cours</option><option>Clôturé</option></select></div>
        </div>}
      </div>);})}
    </div>);})()}

    {tab==="op-agenda"&&(()=>{const list=(agenda||[]).filter(a=>matchOp(a.jeuneId,(a.jeuneNom||"")+" "+(a.type||"")+" "+(a.notes||""))).sort((a,b)=>(b.date||"").localeCompare(a.date||""));return(<div>
      <div style={{fontSize:12,color:C.light,marginBottom:10}}>{list.length} rendez-vous. Suppression centralisée.</div>
      {list.length===0&&<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun rendez-vous</div>}
      {list.map(a=><div key={a.id} style={{...S.card,marginBottom:6,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap"}}>
        <div><div style={{fontWeight:800,fontSize:13,color:C.dark}}>{a.jeuneId?opName(a.jeuneId):(a.jeuneNom||"—")}</div><div style={{fontSize:12,color:C.light}}>{fmt(a.date)}{a.heure?" · "+a.heure:""}{a.type?" · "+a.type:""}{a.lieu?" · "+a.lieu:""}</div></div>
        <button onClick={()=>{if(confirm("Supprimer ce RDV ?"))onUpdateAgenda&&onUpdateAgenda((agenda||[]).filter(x=>x.id!==a.id));}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>
      </div>)}
    </div>);})()}

    {tab==="entretiens"&&(()=>{
      const peut=currentUser&&(currentUser.role==="directeur"||currentUser.role==="chef_service");
      if(!peut)return <div style={{...S.card,fontSize:12,color:C.mid}}>Onglet réservé à la direction et aux chefs de service.</div>;
      const staff=(users||[]).filter(u=>u.role!=="jeune");
      const sal=staff.find(u=>String(u.id)===String(entSel));
      const liste=((sal&&sal.entretiens)||[]).slice().sort((a,b)=>String(b.annee+b.periode).localeCompare(String(a.annee+a.periode)));
      const cur=liste.find(e=>e.id===entOpen);
      const save=(patch)=>{if(!sal)return;onUpdateUsers(prev=>(prev||[]).map(u=>u.id===sal.id?{...u,entretiens:((u.entretiens)||[]).map(e=>e.id===entOpen?{...e,...patch}:e)}:u));};
      const saveRep=(k,v)=>{const e=cur||{};save({reponses:{...(e.reponses||{}),[k]:v}});};
      const creer=()=>{if(!sal)return;const id=Date.now();const nouv={id,annee:new Date().getFullYear(),periode:ENT_PERIODES[0],date:isoToday(),evaluateur:(currentUser&&currentUser.name)||"",statut:"brouillon",reponses:{},creeLe:isoToday()};onUpdateUsers(prev=>(prev||[]).map(u=>u.id===sal.id?{...u,entretiens:[...((u.entretiens)||[]),nouv]}:u));setEntOpen(id);};
      const supprimer=(id)=>{if(!confirm("Supprimer ce compte rendu ?"))return;onUpdateUsers(prev=>(prev||[]).map(u=>u.id===sal.id?{...u,entretiens:((u.entretiens)||[]).filter(e=>e.id!==id)}:u));if(entOpen===id)setEntOpen(null);};
      const cloturer=()=>{const m=entManquants(cur);if(m.length){alert("Champs obligatoires non renseignés :\n\n- "+m.join("\n- "));return;}save({statut:"cloture",clotureLe:isoToday(),cloturePar:(currentUser&&currentUser.name)||""});};
      const verrou=cur&&cur.statut==="cloture";
      return(<div>
        <div style={{...S.card,marginBottom:12}}>
          <label style={{...S.lbl}}>Salarié</label>
          <select style={{...S.inp}} value={entSel} onChange={e=>{setEntSel(e.target.value);setEntOpen(null);}}><option value="">-- Sélectionner --</option>{staff.map(u=>{const n=((u.entretiens)||[]).filter(x=>x.annee===new Date().getFullYear()).length;return<option key={u.id} value={u.id}>{u.name} — {u.role}{" ("+n+"/2 cette année)"}</option>;})}</select>
          <div style={{fontSize:11.5,color:C.light,marginTop:6}}>Deux entretiens attendus par année civile.</div>
        </div>
        {sal&&!cur&&<div style={{...S.card}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,flexWrap:"wrap"}}><div style={{fontWeight:800,fontSize:13,color:C.dark}}>Historique ({liste.length})</div><button onClick={creer} style={{...S.btnP,fontSize:12,padding:"6px 14px"}}><Plus size={13}/>Nouvel entretien</button></div>
          {liste.length===0&&<div style={{fontSize:12,color:C.light}}>Aucun entretien enregistré.</div>}
          {liste.map(e=>(<div key={e.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,padding:"9px 10px",borderTop:"1px solid "+C.border,flexWrap:"wrap"}}>
            <div style={{minWidth:0,flex:1}}><div style={{fontSize:13,fontWeight:700,color:C.dark}}>{entLibelle(e)}</div><div style={{fontSize:11.5,color:C.light}}>Conduit par {e.evaluateur||"—"}</div></div>
            <span style={{fontSize:11.5,fontWeight:800,padding:"2px 8px",borderRadius:6,background:e.statut==="cloture"?"#E8F5E9":"#FFF8E1",color:e.statut==="cloture"?"#2E7D32":"#B8860B"}}>{e.statut==="cloture"?"Clôturé":"Brouillon"}</span>
            <button onClick={()=>setEntOpen(e.id)} style={{...S.btnO,fontSize:12,padding:"5px 12px"}}>Ouvrir</button>
            <button onClick={()=>supprimer(e.id)} style={{background:"none",border:"none",color:"#C62828",cursor:"pointer",fontWeight:800,fontSize:14}}>✕</button>
          </div>))}
        </div>}
        {sal&&cur&&<div style={{...S.card,borderLeft:"4px solid "+C.gold}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,marginBottom:10,flexWrap:"wrap"}}>
            <div><div style={{fontWeight:900,fontSize:15,color:C.dark}}>{sal.name}</div><div style={{fontSize:12,color:C.light}}>{entLibelle(cur)} · {verrou?"clôturé le "+fmt(cur.clotureLe):"brouillon"}</div></div>
            <button onClick={()=>setEntOpen(null)} style={{...S.btnO,fontSize:12,padding:"5px 12px"}}>Retour</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
            <div><label style={{...S.lbl}}>Période</label><select disabled={verrou} style={{...S.inp}} value={cur.periode} onChange={e=>save({periode:e.target.value})}>{ENT_PERIODES.map(p=><option key={p}>{p}</option>)}</select></div>
            <div><label style={{...S.lbl}}>Année</label><input disabled={verrou} type="number" style={{...S.inp}} value={cur.annee} onChange={e=>save({annee:parseInt(e.target.value,10)||new Date().getFullYear()})}/></div>
            <div><label style={{...S.lbl}}>Date</label><input disabled={verrou} type="date" style={{...S.inp}} value={cur.date||""} onChange={e=>save({date:e.target.value})}/></div>
            <div><label style={{...S.lbl}}>Conduit par</label><input disabled={verrou} style={{...S.inp}} value={cur.evaluateur||""} onChange={e=>save({evaluateur:e.target.value})}/></div>
          </div>
          {ENT_GRILLE.map(sec=>(<div key={sec.s} style={{paddingTop:12,borderTop:"1px solid "+C.border,marginBottom:4}}>
            <div style={{fontSize:12,fontWeight:800,color:C.goldDark,textTransform:"uppercase",letterSpacing:"0.04em",marginBottom:8}}>{sec.s}</div>
            {sec.q.map(q=>{const v=(cur.reponses||{})[q.k];const manque=q.obl&&(q.type==="case"?v!==true:!String(v==null?"":v).trim());return(<div key={q.k} style={{marginBottom:10}}>
              <label style={{...S.lbl,color:manque?"#C62828":undefined}}>{q.t}{q.obl?" *":""}</label>
              {q.type==="texte"&&<textarea disabled={verrou} key={cur.id+q.k} defaultValue={v||""} onBlur={e=>saveRep(q.k,e.target.value)} style={{...S.inp,minHeight:56,resize:"vertical",borderColor:manque?"#C62828":undefined}}/>}
              {q.type==="echelle"&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{ENT_ECHELLE.map(n=><button key={n} disabled={verrou} onClick={()=>saveRep(q.k,n)} style={{padding:"5px 11px",borderRadius:16,border:"1.5px solid "+(v===n?C.gold:C.border),background:v===n?C.gold:C.white,color:v===n?C.white:C.mid,fontSize:12,fontWeight:700,cursor:verrou?"default":"pointer",fontFamily:"inherit"}}>{n}</button>)}</div>}
              {q.type==="oui_non"&&<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{["Oui","Non"].map(n=><button key={n} disabled={verrou} onClick={()=>saveRep(q.k,n)} style={{padding:"5px 16px",borderRadius:16,border:"1.5px solid "+(v===n?C.gold:C.border),background:v===n?C.gold:C.white,color:v===n?C.white:C.mid,fontSize:12,fontWeight:700,cursor:verrou?"default":"pointer",fontFamily:"inherit"}}>{n}</button>)}</div>}
              {q.type==="case"&&<label style={{display:"flex",alignItems:"center",gap:8,fontSize:12.5,color:C.dark,cursor:verrou?"default":"pointer",flexWrap:"wrap"}}><input type="checkbox" disabled={verrou} checked={v===true} onChange={e=>saveRep(q.k,e.target.checked)} style={{accentColor:C.gold,width:17,height:17}}/>Oui</label>}
            </div>);})}
          </div>))}
          <div style={{display:"flex",gap:8,flexWrap:"wrap",paddingTop:12,borderTop:"1px solid "+C.border}}>
            {!verrou&&<button onClick={cloturer} style={{...S.btnP,flex:1,justifyContent:"center",minWidth:150}}><Check size={14}/>Clôturer l'entretien</button>}
            {verrou&&currentUser.role==="directeur"&&<button onClick={()=>{if(confirm("Rouvrir ce compte rendu ?"))save({statut:"brouillon",reouvertLe:isoToday(),reouvertPar:currentUser.name});}} style={{...S.btnO,flex:1,justifyContent:"center",minWidth:150}}>Rouvrir</button>}
            <button onClick={async()=>{try{await entretienPDF(cur,sal,etabConfig);}catch(err){alert("PDF impossible : "+(err&&err.message?err.message:err));}}} style={{...S.btnO,flex:1,justifyContent:"center",minWidth:150}}><Download size={14}/>PDF</button>
          </div>
          {!verrou&&entManquants(cur).length>0&&<div style={{fontSize:12,color:"#C62828",fontWeight:700,marginTop:8}}>{entManquants(cur).length} champ(s) obligatoire(s) restant(s).</div>}
        </div>}
      </div>);})()}

    {tab==="planning-cfg"&&(()=>{
      const plan=planSite==="Djilass"?(djiPlan||{}):(fatPlan||{});
      const JOURS=[["0","dimanche"],["1","lundi"],["2","mardi"],["3","mercredi"],["4","jeudi"],["5","vendredi"],["6","samedi"]];
      const jours=()=>{const out=[];if(!planD1||!planD2)return out;let d=new Date(planD1+"T12:00:00");const fin=new Date(planD2+"T12:00:00");if(isNaN(d)||isNaN(fin)||d>fin)return out;
        while(d<=fin){out.push({key:d.toISOString().slice(0,10),wd:d.getDay()});d=new Date(d.getTime()+86400000);}return out;};
      const liste=jours();
      const calcul=()=>{const res={};let eq=planDebut;let init=false;
        liste.forEach((j,idx)=>{
          if(planMode==="rotation"){ if(idx>0&&String(j.wd)===planBascule)eq=(eq==="a"?"b":"a"); if(!init)init=true;
            res[j.key]={a:eq==="a",b:eq==="b"}; }
          else if(planMode==="deux")res[j.key]={a:true,b:true};
          else if(planMode==="a")res[j.key]={a:true,b:false};
          else if(planMode==="b")res[j.key]={a:false,b:true};
          else res[j.key]={a:false,b:false};
        });return res;};
      const proj=calcul();
      const aEcrire=Object.keys(proj).filter(k=>{const ex=plan[k];if(!ex)return true;if(planEcraser)return true;return !(ex.n||ex.a||ex.b||ex.v);});
      const proteges=Object.keys(proj).length-aEcrire.length;
      const appliquer=()=>{
        if(!aEcrire.length){alert("Aucun jour à écrire sur cette période.");return;}
        if(!confirm("Appliquer le modèle sur "+aEcrire.length+" jour(s) du site "+planSite+" ?"+(proteges?"\n\n"+proteges+" jour(s) déjà renseignés seront conservés.":"")))return;
        const maj={};aEcrire.forEach(k=>{maj[k]={...(plan[k]||{}),...proj[k]};});
        onBulkPlan(planSite,maj);
        alert(aEcrire.length+" jour(s) mis à jour.");
      };
      const apercu=liste.slice(0,14);
      return(<div>
        <div style={{...S.card,borderLeft:"4px solid "+C.gold,marginBottom:12}}>
          <h3 style={{fontSize:13.5,fontWeight:800,margin:"0 0 4px",color:C.dark}}>Générer le planning</h3>
          <div style={{fontSize:11.5,color:C.mid,marginBottom:12}}>Remplit les jours d'une période selon un modèle de rotation. Les jours déjà renseignés sont conservés, sauf si vous cochez l'écrasement.</div>
          <label style={{...S.lbl}}>Site</label>
          <div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap"}}>{["Fatick","Djilass"].map(x=><button key={x} onClick={()=>setPlanSite(x)} style={{padding:"7px 16px",borderRadius:20,border:"1.5px solid "+(planSite===x?C.gold:C.border),background:planSite===x?C.gold:C.white,color:planSite===x?C.white:C.mid,fontWeight:700,fontSize:12.5,cursor:"pointer",fontFamily:"inherit"}}>{x}</button>)}</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
            <div><label style={{...S.lbl}}>Du</label><input type="date" style={{...S.inp}} value={planD1} onChange={e=>setPlanD1(e.target.value)}/></div>
            <div><label style={{...S.lbl}}>Au</label><input type="date" style={{...S.inp}} value={planD2} onChange={e=>setPlanD2(e.target.value)}/></div>
          </div>
          <label style={{...S.lbl}}>Modèle</label>
          <select style={{...S.inp,marginBottom:12}} value={planMode} onChange={e=>setPlanMode(e.target.value)}>
            <option value="rotation">Rotation des équipes</option>
            <option value="deux">Les deux équipes tous les jours</option>
            <option value="a">Équipe A seule</option>
            <option value="b">Équipe B seule</option>
            <option value="vide">Aucune équipe (vider)</option>
          </select>
          {planMode==="rotation"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
            <div><label style={{...S.lbl}}>Bascule le</label><select style={{...S.inp}} value={planBascule} onChange={e=>setPlanBascule(e.target.value)}>{JOURS.map(j=><option key={j[0]} value={j[0]}>{j[1]}</option>)}</select></div>
            <div><label style={{...S.lbl}}>Commence par</label><select style={{...S.inp}} value={planDebut} onChange={e=>setPlanDebut(e.target.value)}><option value="a">Équipe A</option><option value="b">Équipe B</option></select></div>
          </div>}
          <label style={{display:"flex",alignItems:"center",gap:9,fontSize:13,fontWeight:700,color:C.dark,cursor:"pointer",padding:"9px 11px",borderRadius:8,background:planEcraser?"#FFEBEE":"transparent",border:"1.5px solid "+(planEcraser?"#C62828":C.border)}}>
            <input type="checkbox" checked={planEcraser} onChange={e=>setPlanEcraser(e.target.checked)} style={{accentColor:"#C62828",width:17,height:17}}/>Écraser les jours déjà renseignés</label>
        </div>
        {liste.length>0&&<div style={{...S.card,marginBottom:12}}>
          <div style={{fontSize:12.5,fontWeight:800,color:C.dark,marginBottom:8}}>Aperçu — {liste.length} jour(s), {aEcrire.length} à écrire{proteges?", "+proteges+" conservé(s)":""}</div>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>{apercu.map(j=>{const q=proj[j.key];const lab=q.a&&q.b?"A+B":q.a?"A":q.b?"B":"—";const bg=q.a&&q.b?"#e0e7ff":q.a?"#fef3c7":q.b?"#dbeafe":C.sableLight;const gar=aEcrire.indexOf(j.key)<0;
            return(<div key={j.key} style={{minWidth:56,padding:"6px 4px",borderRadius:7,background:gar?C.sableLight:bg,border:"1px solid "+(gar?"#ddd":"transparent"),textAlign:"center",opacity:gar?0.45:1}}>
              <div style={{fontSize:11,color:C.light}}>{fmt(j.key).slice(0,5)}</div>
              <div style={{fontSize:12.5,fontWeight:800,color:C.dark}}>{gar?"—":lab}</div>
            </div>);})}</div>
          {liste.length>14&&<div style={{fontSize:11,color:C.light,marginTop:6}}>Les 14 premiers jours sont affichés.</div>}
        </div>}
        <button onClick={appliquer} disabled={!liste.length} style={{...S.btnP,width:"100%",justifyContent:"center"}}><Check size={15}/>Appliquer au planning</button>
      </div>);})()}

    {tab==="registre"&&(()=>{
      const lignes=[...(jeunes||[]),...(appMajeurs||[])].map(p=>({nom:p.nom||"",prenom:p.prenom||"",site:p.site||"",entree:normDate(p.dateDebut),sortie:normDate(p.dateFin)})).sort((a,b)=>String(a.entree||"9").localeCompare(String(b.entree||"9"))||String(a.nom).localeCompare(String(b.nom)));
      const sansEntree=lignes.filter(l=>!l.entree).length;
      return(<div>
        <div style={{...S.card,borderLeft:"4px solid "+C.gold,marginBottom:12}}>
          <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 4px",color:C.dark}}>Registre des personnes accueillies</h3>
          <div style={{fontSize:11.5,color:C.mid,lineHeight:1.55}}>Art. L. 331-2 CASF : identité des personnes séjournant dans l'établissement, date d'entrée et date de sortie. Coté et paraphé selon l'art. R. 331-5. Tenu en permanence à la disposition des autorités judiciaires et administratives. Les agents de contrôle le signent et y consignent leurs observations (art. L. 331-3).</div>
          {sansEntree>0&&<div style={{fontSize:11.5,color:"#C62828",fontWeight:700,marginTop:8}}>{sansEntree} fiche(s) sans date d'entrée : le registre est incomplet au regard de l'article L. 331-2.</div>}
          <button onClick={async()=>{try{await registrePDF(lignes,etabConfig);}catch(err){alert("PDF impossible : "+(err&&err.message?err.message:err));}}} style={{...S.btnP,marginTop:12}}><Download size={14}/>Éditer le registre (PDF)</button>
        </div>
        <div style={{...S.card,padding:"6px 4px",overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",minWidth:460}}>
            <thead><tr style={{background:C.sableLight}}>{["N°","Nom et prénom","Site","Entrée","Sortie"].map(h=><th key={h} style={{fontSize:11,fontWeight:800,color:C.mid,textAlign:"left",padding:"6px 8px",textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>{lignes.map((l,i)=>(<tr key={i} style={{borderTop:"1px solid "+C.border}}>
              <td style={{padding:"6px 8px",fontSize:12,color:C.light}}>{i+1}</td>
              <td style={{padding:"6px 8px",fontSize:12.5,fontWeight:700,color:C.dark}}>{l.nom} {l.prenom}</td>
              <td style={{padding:"6px 8px",fontSize:12,color:C.mid}}>{l.site||"—"}</td>
              <td style={{padding:"6px 8px",fontSize:12,color:l.entree?C.mid:"#C62828",fontWeight:l.entree?600:800}}>{l.entree?fmt(l.entree):"manquante"}</td>
              <td style={{padding:"6px 8px",fontSize:12,color:C.mid}}>{l.sortie?fmt(l.sortie):"—"}</td>
            </tr>))}</tbody>
          </table>
          {lignes.length===0&&<div style={{fontSize:12,color:C.light,padding:"10px 8px"}}>Aucune personne accueillie enregistrée.</div>}
        </div>
      </div>);})()}

    {tab==="projets-cfg"&&(()=>{const cfg=projCfg(etabConfig);const setCfg=(k,v)=>{const n=parseInt(v,10);if(onUpdateEtab)onUpdateEtab(prev=>({...prev,projet:{...projCfg(prev),[k]:isNaN(n)?0:n}}));};const pool=opPool.filter(j=>matchOp(j.id)&&j.statut!=="archivé"&&j.statut!=="inactif");const FIELDS=[["delaiDipc","Remise du DIPC","Art. D.311 CASF : 15 jours maximum"],["delaiObjectifs","Pose des objectifs","Avenant : 6 mois maximum au sens du décret, à raccourcir sur un séjour court"],["delaiRevision","Révision de mi-séjour","Pratique interne"],["delaiBilan","Bilan de fin de séjour","Pratique interne"]];return(<div>
      <div style={{...S.card,borderLeft:"4px solid "+C.gold,marginBottom:14}}>
        <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 4px",color:C.dark}}>Échéances du projet personnalisé</h3>
        <div style={{fontSize:12,color:C.light,marginBottom:10}}>Exprimées en jours à compter de la date d'entrée du bénéficiaire. Les séjours PDSR durant environ cinq mois, ces délais doivent rester inférieurs à la durée du séjour.</div>
        {FIELDS.map(([k,l,h])=><div key={k} style={{marginBottom:10}}>
          <label style={{...S.lbl}}>{l}</label>
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}><input type="number" min="1" style={{...S.inp,width:100}} value={cfg[k]} onChange={e=>setCfg(k,e.target.value)}/><span style={{fontSize:12,color:C.mid}}>jours après l'entrée</span></div>
          <div style={{fontSize:11.5,color:C.light,marginTop:2}}>{h}</div>
        </div>)}
      </div>
      <div style={{...S.card,padding:"6px 4px",overflowX:"auto"}}>
        <div style={{fontSize:13,fontWeight:800,color:C.dark,padding:"8px 8px 4px"}}>Conformité ({pool.length})</div>
        <table style={{width:"100%",borderCollapse:"collapse",minWidth:560}}>
          <thead><tr style={{background:C.sableLight}}>{["Bénéficiaire","DIPC","Objectifs","Révision","Bilan","Participation","Obj."].map(h=><th key={h} style={{fontSize:11.5,fontWeight:800,color:C.mid,textAlign:"left",padding:"6px 8px",textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
          <tbody>{pool.map(j=>{const p=(projets||[]).find(x=>String(x.jeuneId)===String(j.id));const e=projEcheances(j,p,etabConfig);const pa=(p&&p.participation)||{};const tracee=!!(pa.dateEntretien&&((pa.avisJeune||"").trim()||pa.refus));const nObj=p?(p.objectifs||[]).length:0;const nSans=p?(p.objectifs||[]).filter(o=>!o.domaine).length:0;return(<tr key={j.id} style={{borderTop:"1px solid "+C.border}}>
            <td style={{padding:"6px 8px",fontSize:12,fontWeight:700,color:C.dark,whiteSpace:"nowrap"}}>{j.prenom} {j.nom}<div style={{fontSize:12,color:C.light,fontWeight:600}}>{j.site}{!p?" · aucun projet":""}</div></td>
            {e.map(x=>{const st=echStatut(x.due,x.fait);return<td key={x.k} style={{padding:"6px 8px"}}><span style={{fontSize:11.5,fontWeight:800,color:st.c,background:st.bg,borderRadius:6,padding:"2px 7px",whiteSpace:"nowrap"}}>{st.l}</span></td>;})}
            <td style={{padding:"6px 8px"}}><span style={{fontSize:11.5,fontWeight:800,color:tracee?"#2E7D32":"#C62828"}}>{tracee?"Tracée":"Non tracée"}</span></td>
            <td style={{padding:"6px 8px",fontSize:12,fontWeight:700,color:C.dark}}>{nObj}{nSans>0&&<span style={{color:"#C62828",fontSize:11.5,fontWeight:800}}> ({nSans} sans domaine)</span>}</td>
          </tr>);})}</tbody>
        </table>
      </div>
    </div>);})()}

    {tab==="op-projets"&&(()=>{const pool=opPool.filter(j=>matchOp(j.id)&&j.statut!=="archivé");const today2=new Date().toISOString().slice(0,10);return(<div>
      <div style={{fontSize:12,color:C.light,marginBottom:10}}>Vue d'ensemble des projets personnalisés : objectifs définis et échéances dépassées.</div>
      <div style={{...S.card,padding:"6px 4px",overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",minWidth:380}}>
          <thead><tr><th style={{textAlign:"left",padding:7,fontSize:11.5,color:C.light}}>Bénéficiaire</th><th style={{padding:7,fontSize:11.5,color:C.light}}>Objectifs</th><th style={{padding:7,fontSize:11.5,color:C.light}}>En retard</th><th style={{padding:7,fontSize:11.5,color:C.light}}>État</th></tr></thead>
          <tbody>{pool.map(j=>{const p=(projets||[]).find(x=>String(x.jeuneId)===String(j.id));const objs=(p&&p.objectifs)||[];const retard=objs.filter(o=>o.echeance&&o.echeance<today2&&o.statut!=="Atteint").length;const ok=objs.length>0;return(<tr key={j.id}><td style={{padding:7,fontSize:12,fontWeight:700,color:C.dark}}>{j.prenom} {j.nom}</td><td style={{padding:7,textAlign:"center",fontSize:12,color:C.dark}}>{objs.length}</td><td style={{padding:7,textAlign:"center",fontSize:12,fontWeight:700,color:retard>0?"#C62828":C.light}}>{retard||"—"}</td><td style={{padding:7,textAlign:"center"}}><span style={{fontSize:11.5,fontWeight:700,padding:"2px 8px",borderRadius:10,background:ok?(retard>0?"#FFEBEE":"#E8F5E9"):"#F5F5F5",color:ok?(retard>0?"#C62828":"#2E7D32"):C.light}}>{ok?(retard>0?"À revoir":"À jour"):"Aucun projet"}</span></td></tr>);})}</tbody>
        </table>
        {pool.length===0&&<div style={{textAlign:"center",color:C.light,padding:16,fontSize:12}}>Aucun bénéficiaire</div>}
      </div>
    </div>);})()}

    {tab==="op-rsite"&&(()=>{const list=(rapportsSite||[]).slice().sort((a,b)=>(b.date||"").localeCompare(a.date||""));return(<div>
      <div style={{fontSize:12,color:C.light,marginBottom:10}}>{list.length} rapport(s) de site.</div>
      {list.length===0&&<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun rapport de site</div>}
      {list.map(r=><div key={r.id} style={{...S.card,marginBottom:6,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap"}}>
        <div><div style={{fontWeight:800,fontSize:13,color:C.dark}}>{r.site||"—"}</div><div style={{fontSize:12,color:C.light}}>{fmt(r.date)}{r.auteur?" · "+r.auteur:""}</div></div>
        <button onClick={()=>{if(confirm("Supprimer ce rapport de site ?"))onUpdateRapportsSite&&onUpdateRapportsSite((rapportsSite||[]).filter(x=>x.id!==r.id));}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>
      </div>)}
    </div>);})()}

    {tab==="alertes"&&(()=>{
      const actifs=opPool.filter(j=>j.statut!=="archivé"&&j.statut!=="sorti");
      const lastRap={};(rapports||[]).forEach(r=>{const d=r.date||"";if(!lastRap[r.jeuneId]||d>lastRap[r.jeuneId])lastRap[r.jeuneId]=d;});
      const sansRef=actifs.filter(j=>!j.referentA&&!j.referentB);
      const sansRapport=actifs.filter(j=>{const d=lastRap[j.id];if(!d)return true;return (new Date(today)-new Date(d))/86400000>3;});
      const joursEcoules=WEEKDATES.filter(d=>d<=today);
      const presManq=actifs.map(j=>({j,manque:joursEcoules.filter(d=>!(presences||[]).some(p=>p.jeuneId===j.id&&p.date===d)).length})).filter(x=>x.manque>0);
      const eigProb=(evenements||[]).filter(e=>e.eig&&(!e.eigData||!e.eigData.dateTransmission));
      const projRetard=[];actifs.forEach(j=>{const p=(projets||[]).find(x=>String(x.jeuneId)===String(j.id));const objs=(p&&p.objectifs)||[];const r=objs.filter(o=>o.echeance&&o.echeance<today&&o.statut!=="Atteint");if(r.length)projRetard.push(j.prenom+" "+j.nom+" — "+r.length+" obj.");});
      const blocks=[
        {t:"Bénéficiaires sans référent",col:"#C62828",items:sansRef.map(j=>j.prenom+" "+j.nom)},
        {t:"Sans rapport depuis plus de 3 jours",col:"#E65100",items:sansRapport.map(j=>j.prenom+" "+j.nom+(lastRap[j.id]?" (dernier "+fmt(lastRap[j.id])+")":" (aucun)"))},
        {t:"Présences manquantes cette semaine",col:"#E65100",items:presManq.map(x=>x.j.prenom+" "+x.j.nom+" — "+x.manque+" j")},
        {t:"EIG non transmis aux autorités",col:"#C62828",items:eigProb.map(e=>opName(e.jeuneId)+" — "+(e.titre||""))},
        {t:"Projets avec objectifs en retard",col:"#E65100",items:projRetard}
      ];
      const total=blocks.reduce((s,b)=>s+b.items.length,0);
      return(<div>
        <div style={{...S.card,marginBottom:14,display:"flex",alignItems:"center",gap:14,borderLeft:"4px solid "+(total>0?"#C62828":"#2E7D32"),flexWrap:"wrap"}}>
          <div style={{fontSize:34,fontWeight:900,color:total>0?"#C62828":"#2E7D32",lineHeight:1}}>{total}</div>
          <div><div style={{fontWeight:800,fontSize:15,color:C.dark}}>{total>0?"point(s) d'attention":"Tout est à jour"}</div><div style={{fontSize:12,color:C.light}}>Contrôles automatiques de qualité de saisie et de conformité.</div></div>
        </div>
        {blocks.map(b=><div key={b.t} style={{...S.card,marginBottom:10,borderLeft:"4px solid "+(b.items.length>0?b.col:C.border),opacity:b.items.length>0?1:0.55}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:b.items.length>0?8:0,flexWrap:"wrap"}}><span style={{background:b.items.length>0?b.col:C.sable,color:b.items.length>0?"#fff":C.mid,minWidth:26,height:26,borderRadius:8,display:"inline-flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:13}}>{b.items.length}</span><span style={{fontWeight:800,fontSize:13,color:C.dark}}>{b.t}</span></div>
          {b.items.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:6}}>{b.items.slice(0,50).map((it,i)=><span key={i} style={{fontSize:12,background:C.sableLight,color:C.mid,padding:"3px 9px",borderRadius:8,fontWeight:600}}>{it}</span>)}</div>}
        </div>)}
      </div>);
    })()}

    {tab==="stats"&&(()=>{
      const pool=opPool.filter(j=>statSite==="Tous"||j.site===statSite);const ids=new Set(pool.map(j=>j.id));
      const wp=(presences||[]).filter(p=>WEEKDATES.includes(p.date)&&ids.has(p.jeuneId));
      const pres=wp.filter(p=>p.statut==="Présent").length,abs=wp.filter(p=>p.statut==="Absent").length,ret=wp.filter(p=>p.statut==="Retard").length;const totP=wp.length||1;
      const evs=(evenements||[]).filter(e=>ids.has(e.jeuneId));
      const gravCounts=["Léger","Moyen","Grave"].map(g=>({label:g,n:evs.filter(e=>e.gravite===g).length,col:(GC[g]||{}).dot}));
      const typeCounts=[["incident","Incident","#546E7A"],["plainte","Plainte","#E65100"],["reclamation","Réclamation","#1565C0"]].map(([k,l,col])=>({label:l,n:evs.filter(e=>(e.type||"incident")===k).length,col}));
      const rapByAuthor={};(rapports||[]).forEach(r=>{const a=r.author||"Sans auteur";rapByAuthor[a]=(rapByAuthor[a]||0)+1;});
      const rapRows=Object.entries(rapByAuthor).map(([label,n])=>({label,n})).sort((x,y)=>y.n-x.n).slice(0,12);
      const educs2=users.filter(u=>u.role==="educateur"&&(statSite==="Tous"||u.site===statSite));
      const chargeRows=educs2.map(u=>({label:u.name,n:opPool.filter(j=>[j.referentA,j.referentB,j.referentC,j.referentD].includes(u.name)).length})).sort((x,y)=>y.n-x.n);
      const bars=(rows,color)=>{const m=Math.max(1,...rows.map(a=>a.n));return(<div style={{display:"grid",gap:6}}>{rows.length===0?<div style={{fontSize:12,color:C.light}}>Aucune donnée</div>:rows.map((r,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}><div style={{width:115,fontSize:12,fontWeight:700,color:C.mid,textAlign:"right",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.label}</div><div style={{flex:1,background:C.sableLight,borderRadius:6,height:16}}><div style={{width:(r.n/m*100)+"%",background:r.col||color||C.gold,height:"100%",borderRadius:6}}/></div><div style={{width:28,fontSize:12,fontWeight:800,color:C.dark}}>{r.n}</div></div>)}</div>);};
      return(<div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,marginBottom:12,flexWrap:"wrap"}}>
          <button onClick={()=>{const bom="﻿";const rows=[["Statistiques PDSR","Site: "+statSite,"Généré le "+new Date().toLocaleString("fr-FR")],[],["PRÉSENCES — semaine en cours"],["Indicateur","Valeur"],["Taux de présence (%)",Math.round(pres/totP*100)],["Présent",pres],["Absent",abs],["Retard",ret],["Total enregistrements",wp.length],[],["INCIDENTS PAR GRAVITÉ"],["Gravité","Nombre"],...gravCounts.map(g=>[g.label,g.n]),[],["INCIDENTS PAR CATÉGORIE"],["Catégorie","Nombre"],...typeCounts.map(t=>[t.label,t.n]),[],["RAPPORTS RÉDIGÉS PAR ÉDUCATEUR"],["Éducateur","Rapports"],...rapRows.map(r=>[r.label,r.n]),[],["CHARGE — BÉNÉFICIAIRES SUIVIS PAR ÉDUCATEUR"],["Éducateur","Bénéficiaires"],...chargeRows.map(r=>[r.label,r.n])];const csv=rows.map(r=>r.map(c=>'"'+String(c==null?"":c).replace(/"/g,'""')+'"').join(";")).join("\n");const blob=new Blob([bom+csv],{type:"text/csv;charset=utf-8"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="statistiques_pdsr_"+(statSite==="Tous"?"global":statSite.toLowerCase())+"_"+new Date().toISOString().slice(0,10)+".csv";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}} style={{...S.btnP}}><Download size={14}/>Exporter (CSV)</button>
          <select value={statSite} onChange={e=>setStatSite(e.target.value)} style={{...S.inp,width:"auto"}}><option>Tous</option><option>Fatick</option><option>Djilass</option></select>
        </div>
        <div style={{...S.card,marginBottom:12}}>
          <div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:10}}>Présences — semaine en cours</div>
          <div style={{display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
            <div style={{textAlign:"center"}}><div style={{fontSize:30,fontWeight:900,color:"#2E7D32",lineHeight:1}}>{Math.round(pres/totP*100)}%</div><div style={{fontSize:11.5,color:C.light,fontWeight:700,textTransform:"uppercase"}}>Taux de présence</div></div>
            <div style={{flex:1,minWidth:160}}>{[["Présent",pres,"#2E7D32"],["Absent",abs,"#C62828"],["Retard",ret,"#E65100"]].map(([l,n,c])=><div key={l} style={{display:"flex",alignItems:"center",gap:8,marginBottom:5,flexWrap:"wrap"}}><div style={{width:70,fontSize:12,fontWeight:700,color:C.mid}}>{l}</div><div style={{flex:1,background:C.sableLight,borderRadius:6,height:14}}><div style={{width:(n/totP*100)+"%",background:c,height:"100%",borderRadius:6}}/></div><div style={{width:28,fontSize:12,fontWeight:800,color:C.dark}}>{n}</div></div>)}</div>
          </div>
        </div>
        <div style={{...S.card,marginBottom:12}}><div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:10}}>Incidents par gravité</div>{bars(gravCounts)}</div>
        <div style={{...S.card,marginBottom:12}}><div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:10}}>Incidents par catégorie</div>{bars(typeCounts)}</div>
        <div style={{...S.card,marginBottom:12}}><div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:10}}>Rapports rédigés par éducateur</div>{bars(rapRows,C.orange)}</div>
        <div style={{...S.card,marginBottom:12}}><div style={{fontWeight:800,fontSize:13,color:C.dark,marginBottom:10}}>Charge : bénéficiaires suivis par éducateur</div>{bars(chargeRows,C.info)}</div>
      </div>);
    })()}

    {tab==="fiche360"&&(()=>{
      const sel=opPool.find(j=>String(j.id)===String(fiche360Id));
      return(<div>
        <div style={{marginBottom:12}}><select value={fiche360Id} onChange={e=>setFiche360Id(e.target.value)} style={{...S.inp}}><option value="">Choisir un bénéficiaire…</option>{opPool.slice().sort((a,b)=>a.prenom.localeCompare(b.prenom)).map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom} · {j.site}{j.id>=100?" (majeur)":""}</option>)}</select></div>
        {!sel?<div style={{...S.card,textAlign:"center",color:C.light}}>Sélectionnez un bénéficiaire pour voir sa fiche complète.</div>:(()=>{
          const jr=(rapports||[]).filter(r=>String(r.jeuneId)===String(sel.id)).sort((a,b)=>(b.date||"").localeCompare(a.date||""));
          const je=(evenements||[]).filter(e=>String(e.jeuneId)===String(sel.id)).sort((a,b)=>(b.date||"").localeCompare(a.date||""));
          const ja=(agenda||[]).filter(a=>String(a.jeuneId)===String(sel.id)).sort((a,b)=>(b.date||"").localeCompare(a.date||""));
          const jp=(projets||[]).find(x=>String(x.jeuneId)===String(sel.id));const objs=(jp&&jp.objectifs)||[];
          return(<div>
            <div style={{...S.card,background:C.sable,border:"none",marginBottom:12}}><div style={{fontSize:18,fontWeight:900,color:C.dark}}>{sel.prenom} {sel.nom}</div><div style={{fontSize:12,color:C.mid,marginTop:2}}>{sel.site} · Réf : {[sel.referentA,sel.referentB,sel.referentC,sel.referentD].filter(Boolean).join(", ")||"aucun"} · Statut {sel.statut||"actif"}</div>{(sel.dateDebut||sel.dateFin)&&<div style={{fontSize:12,color:C.gold,fontWeight:700,marginTop:2}}>Séjour {sel.dateDebut||"?"} → {sel.dateFin||"?"}</div>}</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:12}}>{[["Rapports",jr.length,C.gold],["Incidents",je.length,"#C62828"],["RDV",ja.length,C.info],["Objectifs",objs.length,C.accent]].map(([l,n,c])=><div key={l} style={{...S.card,marginBottom:0,padding:"12px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:900,color:c,lineHeight:1}}>{n}</div><div style={{fontSize:12,color:C.light,fontWeight:700,textTransform:"uppercase",marginTop:3}}>{l}</div></div>)}</div>
            <div style={{...S.card,marginBottom:10}}><div style={{fontWeight:800,fontSize:13,marginBottom:8,color:C.dark}}>Présences — semaine</div><div style={{display:"flex",gap:5,flexWrap:"wrap"}}>{WEEKDATES.map((d,i)=>{const p=(presences||[]).find(p2=>p2.jeuneId===sel.id&&p2.date===d);const st=p?.statut||"—";const sc2=SC[st]||{bg:C.sableLight,text:C.light,icon:"·"};return(<div key={d} style={{flex:1,textAlign:"center"}}><div style={{fontSize:12,color:C.light,fontWeight:700}}>{WD[i]}</div><div style={{marginTop:3,aspectRatio:"1",borderRadius:7,background:sc2.bg,color:sc2.text,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:13,flexWrap:"wrap"}}>{sc2.icon}</div></div>);})}</div></div>
            <div style={{...S.card,marginBottom:10}}><div style={{fontWeight:800,fontSize:13,marginBottom:8,color:C.dark}}>Derniers rapports</div>{jr.length===0?<div style={{fontSize:12,color:C.light}}>Aucun rapport</div>:jr.slice(0,6).map(r=><div key={r.id} style={{padding:"6px 0",borderBottom:"1px solid "+C.border}}><div style={{fontSize:12,color:C.gold,fontWeight:700}}>{fmt(r.date)}{r.author?" · "+r.author:""}</div><div style={{fontSize:12,color:C.mid,lineHeight:1.4}}>{r.observation}</div></div>)}</div>
            <div style={{...S.card,marginBottom:10}}><div style={{fontWeight:800,fontSize:13,marginBottom:8,color:C.dark}}>Incidents</div>{je.length===0?<div style={{fontSize:12,color:C.light}}>Aucun incident</div>:je.map(e=>{const gc=GC[e.gravite]||GC["Léger"];return(<div key={e.id} style={{padding:"6px 0",borderBottom:"1px solid "+C.border}}><div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}><span style={{background:gc.bg,color:gc.text,borderRadius:10,padding:"1px 8px",fontSize:11.5,fontWeight:700}}>{e.gravite}</span><span style={{fontWeight:700,fontSize:12,color:C.dark}}>{e.titre}</span>{e.eig&&<span style={{background:"#C62828",color:"#fff",borderRadius:4,padding:"1px 6px",fontSize:12,fontWeight:800}}>EIG</span>}<span style={{fontSize:11.5,color:C.light,marginLeft:"auto"}}>{fmt(e.date)}</span></div></div>);})}</div>
            <div style={{...S.card,marginBottom:10}}><div style={{fontWeight:800,fontSize:13,marginBottom:8,color:C.dark}}>Projet personnalisé</div>{objs.length===0?<div style={{fontSize:12,color:C.light}}>Aucun objectif défini</div>:objs.map(o=><div key={o.id} style={{display:"flex",gap:8,alignItems:"center",padding:"4px 0",flexWrap:"wrap"}}><span style={{fontSize:12,color:C.dark,fontWeight:600,flex:1}}>{o.titre}</span>{o.echeance&&<span style={{fontSize:11.5,color:o.echeance<today&&o.statut!=="Atteint"?"#C62828":C.light,fontWeight:700}}>{fmt(o.echeance)}</span>}<span style={{fontSize:11.5,fontWeight:700,padding:"1px 8px",borderRadius:10,background:o.statut==="Atteint"?"#E8F5E9":C.sableLight,color:o.statut==="Atteint"?"#2E7D32":C.mid}}>{o.statut||"En cours"}</span></div>)}</div>
            <div style={{...S.card}}><div style={{fontWeight:800,fontSize:13,marginBottom:8,color:C.dark}}>Rendez-vous</div>{ja.length===0?<div style={{fontSize:12,color:C.light}}>Aucun RDV</div>:ja.slice(0,8).map(a=><div key={a.id} style={{display:"flex",gap:8,alignItems:"center",padding:"4px 0",borderBottom:"1px solid "+C.border,flexWrap:"wrap"}}><span style={{fontSize:12,color:C.dark,fontWeight:700}}>{fmt(a.date)}{a.heure?" · "+a.heure:""}</span><span style={{fontSize:12,color:C.light}}>{a.type||""}{a.with?" · "+a.with:""}</span></div>)}</div>
          </div>);
        })()}
      </div>);
    })()}

    {tab==="config"&&(()=>{const set=(f,v)=>onUpdateEtab&&onUpdateEtab(prev=>({...prev,[f]:v}));const fields=[["raisonSociale","Raison sociale"],["sousTitre","Sous-titre / objet"],["finess","N° FINESS"],["directeur","Directeur (signataire)"],["adresse","Adresse"],["ville","Code postal / Ville"],["tel","Téléphone"],["email","Email"]];return(<div>
      <div style={{fontWeight:700,fontSize:15,marginBottom:4,color:C.dark}}>Identité de l'établissement</div>
      <div style={{fontSize:12,color:C.light,marginBottom:14}}>Ces informations identifient l'association. Elles sont enregistrées et synchronisées, et serviront de source pour les documents générés.</div>
      <div style={{...S.card}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>{fields.map(([f,l])=><div key={f} style={{gridColumn:(f==="raisonSociale"||f==="sousTitre"||f==="adresse")?"1/-1":"auto"}}><label style={{...S.lbl}}>{l}</label><input style={{...S.inp}} value={ec[f]||""} onChange={e=>set(f,e.target.value)} placeholder={l}/></div>)}</div>
      </div>
      <div style={{fontSize:12,color:C.mid,marginTop:10,padding:"10px 12px",background:C.goldLight,borderRadius:8,border:"1px solid "+C.border}}>ℹ️ La propagation automatique de ces valeurs dans tous les modèles de documents (en-têtes, signatures) est une étape distincte : aujourd'hui certains documents conservent encore les valeurs d'origine codées. Dis-le-moi quand tu veux que je rebranche les générateurs sur cette config.</div>
    </div>);})()}

    {tab==="suivi-rapports"&&(()=>{
      const wkStart=(ds)=>{if(!ds)return"";const d=new Date(ds+"T00:00:00");const day=(d.getDay()+6)%7;d.setDate(d.getDate()-day);return localDay(d);};
      const wkLabel=(ws)=>{if(!ws)return"—";const d=new Date(ws+"T00:00:00");const e=new Date(d);e.setDate(d.getDate()+6);const f=x=>("0"+x.getDate()).slice(-2)+"/"+("0"+(x.getMonth()+1)).slice(-2);return f(d)+" → "+f(e);};
      const TYPES=[["journee","Journée"],["rdv_parents","RDV parents"],["rdv_exterieur","RDV ext."]];
      const pool=opPool.filter(j=>srSite==="Tous"||j.site===srSite);const ids=new Set(pool.map(j=>j.id));
      const rs=(rapports||[]).filter(r=>ids.has(r.jeuneId));
      const weeks=[...new Set(rs.map(r=>wkStart(r.date)).filter(Boolean))].sort().reverse();
      const curWeek=srWeek||weeks[0]||"";const curJeune=srJeune||(pool[0]&&pool[0].id)||"";
      const countBy=(arr)=>{const o={total:arr.length};TYPES.forEach(([k])=>o[k]=arr.filter(r=>(r.typeContact||"journee")===k).length);return o;};
      const exp=(rows,head,fname)=>{const bom="﻿";const csv=[head,...rows].map(r=>r.map(c=>'"'+String(c==null?"":c).replace(/"/g,'""')+'"').join(";")).join("\n");const b=new Blob([bom+csv],{type:"text/csv;charset=utf-8"});const u=URL.createObjectURL(b);const a=document.createElement("a");a.href=u;a.download=fname;document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(u);};
      const cell={padding:"7px 6px",fontSize:12,textAlign:"center",borderBottom:"1px solid "+C.border};const th={padding:"7px 6px",fontSize:11.5,fontWeight:800,color:C.light,textTransform:"uppercase"};
      return(<div>
        <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
          {[["semaine","Par semaine"],["matrice","Toutes les semaines"],["jeune","Par jeune"]].map(([k,l])=><button key={k} onClick={()=>setSrMode(k)} style={{padding:"6px 14px",borderRadius:20,border:"1.5px solid "+(srMode===k?C.gold:C.border),background:srMode===k?C.gold:C.white,color:srMode===k?"#fff":C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{l}</button>)}
          <select value={srSite} onChange={e=>setSrSite(e.target.value)} style={{...S.inp,width:"auto",marginLeft:"auto"}}><option>Tous</option><option>Fatick</option><option>Djilass</option></select>
        </div>
        {srMode==="semaine"&&<div>
          <div style={{display:"flex",gap:8,marginBottom:10,alignItems:"center",flexWrap:"wrap"}}>
            <select value={curWeek} onChange={e=>setSrWeek(e.target.value)} style={{...S.inp,width:"auto"}}>{weeks.length===0&&<option value="">—</option>}{weeks.map(w=><option key={w} value={w}>Semaine du {wkLabel(w)}</option>)}</select>
            <button onClick={()=>{const rows=pool.map(j=>{const c=countBy(rs.filter(r=>r.jeuneId===j.id&&wkStart(r.date)===curWeek));return[j.prenom+" "+(j.nom||""),c.total,c.journee,c.rdv_parents,c.rdv_exterieur];});exp(rows,["Jeune","Total","Journée","RDV parents","RDV ext."],"suivi_rapports_semaine_"+curWeek+".csv");}} style={{...S.btnO,marginLeft:"auto"}}><Download size={14}/>Export CSV</button>
          </div>
          <div style={{...S.card,padding:"4px 6px",overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",minWidth:380}}><thead><tr><th style={{...th,textAlign:"left"}}>Jeune</th><th style={th}>Total</th>{TYPES.map(([k,l])=><th key={k} style={th}>{l}</th>)}</tr></thead><tbody>{pool.map(j=>{const c=countBy(rs.filter(r=>r.jeuneId===j.id&&wkStart(r.date)===curWeek));return(<tr key={j.id}><td style={{...cell,textAlign:"left",fontWeight:700,color:C.dark,whiteSpace:"nowrap"}}>{j.prenom} {j.nom}</td><td style={{...cell,fontWeight:800,color:c.total===0?"#C62828":C.dark}}>{c.total}</td>{TYPES.map(([k])=><td key={k} style={{...cell,color:C.mid}}>{c[k]||"—"}</td>)}</tr>);})}</tbody></table></div>
        </div>}
        {srMode==="jeune"&&<div>
          <div style={{display:"flex",gap:8,marginBottom:10,alignItems:"center",flexWrap:"wrap"}}>
            <select value={curJeune} onChange={e=>setSrJeune(e.target.value)} style={{...S.inp,width:"auto"}}>{pool.map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom}</option>)}</select>
            <button onClick={()=>{const rows=weeks.map(w=>{const c=countBy(rs.filter(r=>String(r.jeuneId)===String(curJeune)&&wkStart(r.date)===w));return["Semaine du "+wkLabel(w),c.total,c.journee,c.rdv_parents,c.rdv_exterieur];});exp(rows,["Semaine","Total","Journée","RDV parents","RDV ext."],"suivi_rapports_jeune.csv");}} style={{...S.btnO,marginLeft:"auto"}}><Download size={14}/>Export CSV</button>
          </div>
          <div style={{...S.card,padding:"4px 6px",overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",minWidth:360}}><thead><tr><th style={{...th,textAlign:"left"}}>Semaine</th><th style={th}>Total</th>{TYPES.map(([k,l])=><th key={k} style={th}>{l}</th>)}</tr></thead><tbody>{weeks.map(w=>{const c=countBy(rs.filter(r=>String(r.jeuneId)===String(curJeune)&&wkStart(r.date)===w));return(<tr key={w}><td style={{...cell,textAlign:"left",fontWeight:700,color:C.dark,whiteSpace:"nowrap"}}>{wkLabel(w)}</td><td style={{...cell,fontWeight:800,color:c.total===0?"#C62828":C.dark}}>{c.total}</td>{TYPES.map(([k])=><td key={k} style={{...cell,color:C.mid}}>{c[k]||"—"}</td>)}</tr>);})}{weeks.length===0&&<tr><td colSpan={5} style={{...cell,color:C.light}}>Aucun rapport</td></tr>}</tbody></table></div>
        </div>}
        {srMode==="matrice"&&<div>
          <div style={{display:"flex",justifyContent:"flex-end",marginBottom:10,flexWrap:"wrap"}}><button onClick={()=>{const head=["Jeune",...weeks.map(w=>wkLabel(w)),"Total"];const rows=pool.map(j=>{const per=weeks.map(w=>rs.filter(r=>r.jeuneId===j.id&&wkStart(r.date)===w).length);return[j.prenom+" "+(j.nom||""),...per,per.reduce((a,b)=>a+b,0)];});exp(rows,head,"suivi_rapports_matrice.csv");}} style={{...S.btnO}}><Download size={14}/>Export CSV</button></div>
          <div style={{...S.card,padding:"4px 6px",overflowX:"auto"}}><table style={{borderCollapse:"collapse",minWidth:Math.max(360,160+weeks.length*60)}}><thead><tr><th style={{...th,textAlign:"left"}}>Jeune</th>{weeks.map(w=><th key={w} style={th}>{wkLabel(w)}</th>)}<th style={th}>Total</th></tr></thead><tbody>{pool.map(j=>{const per=weeks.map(w=>rs.filter(r=>r.jeuneId===j.id&&wkStart(r.date)===w).length);const tot=per.reduce((a,b)=>a+b,0);return(<tr key={j.id}><td style={{...cell,textAlign:"left",fontWeight:700,color:C.dark,whiteSpace:"nowrap"}}>{j.prenom} {j.nom}</td>{per.map((n,i)=><td key={i} style={{...cell,color:n===0?C.sable:C.dark,fontWeight:n>0?700:400}}>{n||"·"}</td>)}<td style={{...cell,fontWeight:800,color:C.gold}}>{tot}</td></tr>);})}{weeks.length===0&&<tr><td style={{...cell,color:C.light}}>Aucun rapport</td></tr>}</tbody></table></div>
          <div style={{fontSize:11.5,color:C.light,marginTop:6}}>Chaque case = total de rapports cette semaine-là. Le détail par type est dans « Par semaine » et « Par jeune ».</div>
        </div>}
      </div>);
    })()}

    {tab==="maintenance"&&isAdmin&&(()=>{const educList=users.filter(u=>u.role==="educateur"||u.role==="coordinateur_site");return(<div>
      <div style={{...S.card,borderLeft:"4px solid #C62828",marginBottom:14}}>
        <div style={{fontSize:11.5,fontWeight:800,color:C.light,marginBottom:6,letterSpacing:"0.05em"}}>VERSION {APP_BUILD}</div>
        <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 4px",color:"#C62828"}}>Réinitialisation de fin de séjour</h3>
        <div style={{fontSize:12,color:C.mid,marginBottom:10,lineHeight:1.55}}>Efface jeunes, majeurs, rapports, incidents, présences, agenda et projets, sur le serveur et sur <b>tous les appareils</b>. Les postes restés ouverts se videront et se rechargeront d'eux-mêmes à leur prochaine synchronisation, sans manipulation de leur part. Les comptes utilisateurs, les documents RH et la configuration sont conservés.</div>
        <button onClick={()=>{const c=prompt("Cette action est irréversible et s'applique à tous les appareils.\n\nTapez REINITIALISER pour confirmer.");if(!currentUser||!currentUser.isAdmin){alert("Seul l'administrateur peut réinitialiser.");return;}if(c==="REINITIALISER"&&onResetGlobal)onResetGlobal();else if(c!==null)alert("Confirmation incorrecte. Rien n'a été fait.");}} style={{padding:"9px 16px",borderRadius:9,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontWeight:800,fontSize:12.5,cursor:"pointer",fontFamily:"inherit"}}>Réinitialiser pour la nouvelle saison</button>
      </div>
      <div style={{...S.card,borderLeft:"4px solid "+C.accent,marginBottom:14}}>
        <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 4px",color:C.dark}}>👁 Voir comme</h3>
        <div style={{fontSize:12,color:C.light,marginBottom:10}}>Ouvre l'app telle que la voit un éducateur, pour vérifier ses accès et son périmètre. Aperçu — évitez de modifier des données pendant l'impersonation.</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <select value={viewAsId} onChange={e=>setViewAsId(e.target.value)} style={{...S.inp,flex:1,minWidth:180}}><option value="">Choisir un compte…</option>{educList.map(u=><option key={u.id} value={u.id}>{u.name} · {u.role==="coordinateur_site"?"Coord.":"Éduc."} {u.site}{u.isEducMajeur?" · Majeurs":""}</option>)}</select>
          <button onClick={()=>{const u=users.find(x=>String(x.id)===String(viewAsId));if(u&&onViewAs)onViewAs(u);}} disabled={!viewAsId} style={{...S.btnP}}>Voir comme</button>
        </div>
      </div>
      <div style={{...S.card,borderLeft:"4px solid "+C.gold,marginBottom:14}}>
        <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 4px",color:C.dark}}>💾 Sauvegarde & restauration</h3>
        <div style={{fontSize:12,color:C.light,marginBottom:10}}>Export complet de toutes les données (JSON) et réimport. À faire avant toute manipulation risquée.</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <button onClick={()=>{try{const data=onBackup&&onBackup();const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="sauvegarde_pdsr_"+new Date().toISOString().slice(0,10)+".json";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}catch(e){alert("Échec de l'export");}}} style={{...S.btnP}}><Download size={14}/>Exporter (JSON)</button>
          <label style={{...S.btnS,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6}}>Restaurer…<input type="file" accept=".json" style={{display:"none"}} onChange={async(e)=>{const f=e.target.files[0];if(!f)return;if(!confirm("Restaurer ces données ? L'état actuel sera remplacé."))return;try{const txt=await f.text();const data=JSON.parse(txt);onRestore&&onRestore(data);alert("Restauration effectuée.");}catch(err){alert("Fichier invalide.");}e.target.value="";}}/></label>
        </div>
      </div>
      <div style={{...S.card,borderLeft:"4px solid "+C.info,marginBottom:14}}>
        <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 4px",color:C.dark}}>☁ Synchronisation cloud</h3>
        <div style={{fontSize:12,color:C.light,marginBottom:10}}>Forcer l'envoi des données locales vers le cloud, ou recharger depuis le cloud. Utile en cas de connexion instable.</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <button onClick={()=>onForcePush&&onForcePush()} style={{...S.btnP}}>↑ Envoyer au cloud</button>
          <button onClick={()=>{if(confirm("Recharger depuis le cloud ? Les modifications locales non synchronisées seront perdues."))onForcePull&&onForcePull();}} style={{...S.btnS}}>↓ Recharger du cloud</button>
          <button onClick={async()=>{setIntegrity("loading");const r=onCheckIntegrity&&await onCheckIntegrity();setIntegrity(r||"error");}} style={{...S.btnO}}>Vérifier l'intégrité</button>
        </div>
        {integrity==="loading"&&<div style={{fontSize:12,color:C.light,marginTop:10}}>Vérification…</div>}
        {integrity==="error"&&<div style={{fontSize:12,color:C.danger,marginTop:10}}>Impossible de joindre le cloud.</div>}
        {Array.isArray(integrity)&&<div style={{marginTop:12,overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",minWidth:320}}><thead><tr><th style={{textAlign:"left",padding:6,fontSize:11.5,color:C.light}}>Donnée</th><th style={{padding:6,fontSize:11.5,color:C.light}}>Local</th><th style={{padding:6,fontSize:11.5,color:C.light}}>Cloud</th><th style={{padding:6,fontSize:11.5,color:C.light}}>État</th></tr></thead><tbody>{integrity.map(r=>{const ok=r.local===r.cloud;return(<tr key={r.label}><td style={{padding:6,fontSize:12,fontWeight:700,color:C.dark}}>{r.label}</td><td style={{padding:6,textAlign:"center",fontSize:12}}>{r.local}</td><td style={{padding:6,textAlign:"center",fontSize:12}}>{r.cloud}</td><td style={{padding:6,textAlign:"center"}}><span style={{fontSize:14}}>{ok?"✓":"⚠"}</span></td></tr>);})}</tbody></table><div style={{fontSize:11.5,color:C.light,marginTop:6}}>✓ = local et cloud alignés · ⚠ = écart (pensez à synchroniser)</div></div>}
      </div>
      <div style={{...S.card,borderLeft:"4px solid "+C.orange,marginBottom:14}}>
        <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 4px",color:C.dark}}>📦 Clôture de séjour</h3>
        <div style={{fontSize:12,color:C.light,marginBottom:10}}>Archive l'ensemble des données du séjour (cloud + fichier JSON horodaté) pour repartir proprement sur la session suivante. L'archive est conservée ; les données actuelles ne sont pas effacées automatiquement.</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <input value={clotureLabel} onChange={e=>setClotureLabel(e.target.value)} placeholder="Libellé (ex : Séjour Fatick mars-août 2026)" style={{...S.inp,flex:1,minWidth:200}}/>
          <button onClick={()=>{if(confirm("Archiver le séjour actuel ?"))onArchiveSejour&&onArchiveSejour(clotureLabel);}} style={{...S.btnP}}><Download size={14}/>Archiver le séjour</button>
        </div>
      </div>
      <div style={{...S.card,borderLeft:"4px solid #C62828"}}>
        <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 4px",color:C.dark}}>🔑 Réinitialisation des mots de passe</h3>
        <div style={{fontSize:12,color:C.light,marginBottom:10}}>Réinitialise le mot de passe de tous les comptes éducateurs au format <b>identifiant + 2026</b>. Les comptes encadrants ne sont pas touchés.</div>
        <button onClick={()=>{if(confirm("Réinitialiser le MDP de tous les éducateurs ?"))onUpdateUsers(users.map(u=>u.role==="educateur"?{...u,password:(u.login||"compte")+"2026"}:u));}} style={{padding:"10px 18px",borderRadius:10,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>Réinitialiser tous les MDP éducateurs</button>
        <div style={{fontSize:11.5,color:C.light,marginTop:8}}>Note : forcer la déconnexion à distance n'est pas possible côté application (pas de session serveur). Pour bloquer un accès immédiatement, désactivez le compte dans « Équipe ».</div>
      </div>
    </div>);})()}
  </div>);
}

function AgendaPage({agenda,setAgenda,jeunes,majeurs,users,user}){
 const[showForm,setShowForm]=useState(false); const[editId,setEditId]=useState(null);
 const[rdvJeune,setRdvJeune]=useState("");
 const[rdvDate,setRdvDate]=useState(new Date().toISOString().slice(0,10));
 const[rdvHeure,setRdvHeure]=useState("09:00");
 const[rdvType,setRdvType]=useState("educateur");
 const[rdvDesc,setRdvDesc]=useState("");
 const[rdvWith,setRdvWith]=useState("");
 const allJ=[...jeunes,...majeurs];
 const sortedAgenda=(agenda||[]).filter(a=>{const j=allJ.find(x=>String(x.id)===String(a.jeuneId));if(user.role==="educateur"){if(user.isEducMajeur){if(j&&j.id<100)return false;}else{if(j&&j.id>=100)return false;}if(user.site!=="Tous"&&j&&j.site!==user.site)return false;}return true;}).sort((a,b)=>(a.date+a.heure).localeCompare(b.date+b.heure)).reverse();
 const addRdv=()=>{if(!rdvJeune||!rdvDate)return alert("Jeune et date requis");const jid=Number(rdvJeune);if(editId){setAgenda(p=>p.map(a=>a.id===editId?{...a,jeuneId:jid,jeuneNom:(allJ.find(j=>String(j.id)===String(rdvJeune))||{}).prenom||"",date:rdvDate,heure:rdvHeure,type:rdvType,description:rdvDesc,with:rdvWith}:a));setEditId(null);}else{setAgenda(p=>[...p,{id:Date.now(),jeuneId:jid,jeuneNom:(allJ.find(j=>String(j.id)===String(rdvJeune))||{}).prenom||"",date:rdvDate,heure:rdvHeure,type:rdvType,description:rdvDesc,with:rdvWith,createdBy:user.name}]);}setShowForm(false);setRdvDesc("");setRdvWith("");};
 const delRdv=id=>{setDeletionLogs(prev=>[...prev,{id:"dl_"+Date.now()+"_agenda_"+id,type:"agenda",origId:id,date:new Date().toISOString()}]);setAgenda(p=>p.filter(r=>r.id!==id));};
 const getJeuneName=id=>{const j=allJ.find(j2=>String(j2.id)===String(id));return j?(j.prenom+" "+(j.nom||"")):"?";};
 const typeColors={educateur:"#3498db",referent:"#2ecc71",medical:"#e74c3c",juridique:"#9b59b6",autre:"#f39c12"};
 return(<div>
  <div style={{...S.card,marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap"}}><div><div style={{fontWeight:700,fontSize:16,color:C.dark}}>Agenda / Rendez-vous</div><div style={{fontSize:12,color:C.light}}>{(agenda||[]).length} RDV enregistrés</div></div><button onClick={()=>{setEditId(null);setRdvJeune("");setRdvDate(new Date().toISOString().slice(0,10));setRdvHeure("09:00");setRdvType("educateur");setRdvDesc("");setRdvWith("");setShowForm(!showForm);}} style={{...S.btn}}><Plus size={14}/> Nouveau RDV</button></div>
  {showForm&&<div style={{...S.card,marginBottom:12}}>
   <div style={{fontWeight:700,marginBottom:8}}>{editId?"Modifier le rendez-vous":"Nouveau rendez-vous"}</div>
   <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
    <div><div style={{fontSize:12,fontWeight:700,color:C.light,marginBottom:4}}>Jeune</div><select value={rdvJeune} onChange={e=>setRdvJeune(e.target.value)} style={{...S.inp}}><option value="">Choisir...</option>{allJ.filter(j=>{if(user.role==="educateur"){if(user.isEducMajeur&&j.id<100)return false;if(!user.isEducMajeur&&j.id>=100)return false;if(user.site!=="Tous"&&j.site!==user.site)return false;}return true;}).map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom||""}</option>)}</select></div>
    <div><div style={{fontSize:12,fontWeight:700,color:C.light,marginBottom:4}}>Type</div><select value={rdvType} onChange={e=>setRdvType(e.target.value)} style={{...S.inp}}><option value="educateur">Éducateur</option><option value="referent">Référent ASE</option><option value="medical">Médical</option><option value="juridique">Juridique</option><option value="autre">Autre</option></select></div>
    <div><div style={{fontSize:12,fontWeight:700,color:C.light,marginBottom:4}}>Date</div><input type="date" value={rdvDate} onChange={e=>setRdvDate(e.target.value)} style={{...S.inp}}/></div>
    <div><div style={{fontSize:12,fontWeight:700,color:C.light,marginBottom:4}}>Heure</div><input type="time" value={rdvHeure} onChange={e=>setRdvHeure(e.target.value)} style={{...S.inp}}/></div>
   </div>
   <div style={{marginTop:8}}><div style={{fontSize:12,fontWeight:700,color:C.light,marginBottom:4}}>Avec qui / Détails</div><input value={rdvWith} onChange={e=>setRdvWith(e.target.value)} placeholder="Nom du professionnel..." style={{...S.inp,marginBottom:6}}/></div>
   <div><div style={{fontSize:12,fontWeight:700,color:C.light,marginBottom:4}}>Description</div><textarea value={rdvDesc} onChange={e=>setRdvDesc(e.target.value)} placeholder="Détails du RDV..." rows={2} style={{...S.inp,marginBottom:6}}/></div>
   <div style={{display:"flex",gap:6,flexWrap:"wrap"}}><button onClick={addRdv} style={{...S.btn}}>Valider</button><button onClick={()=>{setShowForm(false);setEditId(null);}} style={{...S.btnO}}>Annuler</button></div>
  </div>}
  {sortedAgenda.length===0?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun rendez-vous</div>:sortedAgenda.map(r=><div key={r.id} style={{...S.card,marginBottom:8,borderLeft:"4px solid "+(typeColors[r.type]||"#ccc")}}>
   <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap"}}>
    <div><div style={{fontWeight:700,color:C.dark}}>{getJeuneName(r.jeuneId)}</div><div style={{fontSize:12,color:C.light}}>{r.date} à {r.heure} — <span style={{color:typeColors[r.type]||"#999",fontWeight:600}}>{r.type}</span>{r.with?" — "+r.with:""}</div>{r.description&&<div style={{fontSize:12,color:C.mid,marginTop:4}}>{r.description}</div>}{(user.role==="directeur"||user.role==="chef_service")&&<div style={{marginTop:6,borderTop:"1px solid #eee",paddingTop:6}}><div style={{fontSize:12,fontWeight:700,color:C.primary,marginBottom:3}}>CR du RDV</div><textarea value={r.cr||""} onChange={e=>{setAgenda(p=>p.map(a=>a.id===r.id?{...a,cr:e.target.value}:a));}} placeholder="Saisir le compte-rendu..." rows={2} style={{...S.inp,fontSize:12,width:"100%"}}/></div>}</div>
    <div style={{display:"flex",flexDirection:"column",gap:4}}>{user.role==="chef_service"&&<button onClick={()=>delRdv(r.id)} style={{background:"none",border:"none",color:"#e74c3c",cursor:"pointer",fontSize:14}} title="Supprimer">🗑️</button>}<button onClick={()=>{setEditId(r.id);setRdvJeune(String(r.jeuneId));setRdvDate(r.date);setRdvHeure(r.heure||"09:00");setRdvType(r.type||"educateur");setRdvDesc(r.description||"");setRdvWith(r.with||"");setShowForm(true);}} style={{background:"none",border:"none",color:C.primary,cursor:"pointer",fontSize:14}} title="Modifier">✏️</button></div>
   </div>
  </div>)}
 </div>);
}

function MajeurDetail({majeur,rapports,presences,evenements,user,onBack,onAddR,onAddE,onCP,users,addR,addE,onUpdateMajeur,projets,onUpdateProjets,etabConfig}){
 const[tab,setTab]=useState("fiche");
 const[showFormR,setShowFormR]=useState(false);
 const[showFormE,setShowFormE]=useState(false);
 const[rDate,setRDate]=useState(new Date().toISOString().slice(0,10));
 const[rObs,setRObs]=useState("");
 const[rTC,setRTC]=useState("journee");
 const[eDate,setEDate]=useState(new Date().toISOString().slice(0,10));
 const[eTitre,setETitre]=useState("");
 const[eDesc,setEDesc]=useState("");
 const[eGrav,setEGrav]=useState("Léger");
 const[eType,setEType]=useState("incident");
 const[eEig,setEEig]=useState(false);
 const mr=(rapports||[]).filter(r=>r.jeuneId===majeur.id).sort((a,b)=>b.date.localeCompare(a.date));
 const me=(evenements||[]).filter(e=>e.jeuneId===majeur.id).sort((a,b)=>b.date.localeCompare(a.date));
 const mp=Object.entries(presences||{}).filter(([k])=>k.startsWith("2")).reduce((acc,[date,v])=>{if(v[majeur.id])acc[date]=v[majeur.id];return acc;},{});
 const tabs=["fiche","rapports","projet","stages","incidents","presences"];
 const TC_LABELS={journee:"Journée du jeune",rdv_parents:"RDV tél. parents",rdv_exterieur:"RDV tél. contact ext."};
 const submitR=()=>{if(!rObs.trim())return alert("Observation requise");addR&&addR({jeuneId:majeur.id,date:rDate,observation:rObs.trim(),typeContact:rTC});setRObs("");setRTC("journee");setShowFormR(false);};
 const submitE=()=>{if(!eTitre.trim())return alert("Titre requis");addE&&addE({jeuneId:majeur.id,date:eDate,titre:eTitre.trim(),description:eDesc.trim(),gravite:eGrav,type:eType,categorie:"jeune",site:majeur.site||null,eig:eEig,eigData:eEig?{destinataires:"",dateTransmission:"",accuseReception:false,statutCloture:"En cours"}:null});setETitre("");setEDesc("");setEEig(false);setShowFormE(false);};
 return(<div style={{padding:"14px",maxWidth:700,margin:"0 auto"}}>
 <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12,cursor:"pointer",flexWrap:"wrap"}} onClick={onBack}><ChevronLeft size={18}/><span style={{fontWeight:700,color:C.dark}}>Retour Majeurs</span></div>
 <div style={{...S.card,background:C.sable,border:"none",marginBottom:14}}>
   <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
     <div style={{width:52,height:52,borderRadius:15,background:C.white,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:20,color:C.gold,flexWrap:"wrap"}}>{majeur.prenom[0]}{majeur.nom?majeur.nom[0]:""}</div>
     <div><div style={{fontSize:20,fontWeight:900,color:C.dark}}>{majeur.prenom} {majeur.nom||""} <span style={{fontSize:12,fontWeight:800,color:"#6A1B9A",background:"#F3E5F5",borderRadius:10,padding:"2px 8px",verticalAlign:"middle"}}>MAJEUR</span></div><div style={{fontSize:12,color:C.mid,marginTop:2}}>Réf: {majeur.referentA||"Non affecté"}{majeur.referentB?" / "+majeur.referentB:""}</div><div style={{fontSize:12,color:C.gold,fontWeight:700,marginTop:2}}>Site {majeur.site} · {majeur.dateDebut||"?"} → {majeur.dateFin||"?"}</div></div>
   </div>
 </div>
 <div style={{display:"flex",gap:5,marginBottom:14,overflowX:"auto",paddingBottom:4,flexWrap:"wrap"}}>{tabs.map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:"7px 14px",borderRadius:20,border:`1.5px solid ${tab===t?C.gold:C.border}`,background:tab===t?C.gold:C.white,color:tab===t?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap",textTransform:"capitalize"}}>{t==="presences"?"Présences":t}</button>)}</div>
 {tab==="fiche"&&<div style={{...S.card}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>{[["Prénom",majeur.prenom],["Nom",majeur.nom||"-"],["Site",majeur.site],["Statut",majeur.statut],["Début",majeur.dateDebut||"-"],["Fin",majeur.dateFin||"-"],["Email ASE",majeur.emailASE||"-"],["Tel parent",majeur.telParent1||"-"],["Tel jeune",majeur.telJeune||"-"],["Référent A",majeur.referentA||"-"],["Référent B",majeur.referentB||"-"]].map(([k,v])=><div key={k}><div style={{fontSize:11.5,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>{k}</div><div style={{fontWeight:700,color:C.dark,fontSize:13}}>{v}</div></div>)}</div></div>}
 {tab==="rapports"&&<div>{showFormR?<div style={{...S.card,borderLeft:`4px solid ${C.gold}`,marginBottom:8}}><div style={{fontWeight:800,marginBottom:8,color:C.dark}}>Nouveau rapport</div><label style={{...S.lbl}}>Date</label><input type="date" value={rDate} onChange={e=>setRDate(e.target.value)} style={{...S.inp,marginBottom:8}}/><label style={{...S.lbl}}>Type de contact</label><select value={rTC} onChange={e=>setRTC(e.target.value)} style={{...S.inp,marginBottom:8}}><option value="journee">Journée du jeune</option><option value="rdv_parents">RDV tél. parents</option><option value="rdv_exterieur">RDV tél. contact ext.</option></select><label style={{...S.lbl}}>Observation</label><textarea value={rObs} onChange={e=>setRObs(e.target.value)} placeholder="Observation..." rows={4} style={{...S.inp,marginBottom:8,resize:"vertical"}}/><div style={{display:"flex",gap:6,flexWrap:"wrap"}}><button onClick={submitR} style={{...S.btnP,background:"#2E7D32",borderColor:"#2E7D32"}}>Enregistrer</button><button onClick={()=>setShowFormR(false)} style={{...S.btnO}}>Annuler</button></div></div>:null}{mr.length===0&&!showFormR?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun rapport</div>:mr.map(r=><div key={r.id} style={{...S.card,marginBottom:8}}><div style={{fontSize:12,color:C.gold,fontWeight:700,marginBottom:5}}>{fmt(r.date)} · {TC_LABELS[r.typeContact]||"Journalier"}{r.author&&<span style={{fontWeight:400,fontSize:12,color:C.light,marginLeft:8}}>par {r.author}</span>}</div><p style={{margin:0,fontSize:13,color:C.dark,lineHeight:1.6}}>{r.observation}</p></div>)}{!showFormR&&<button onClick={()=>setShowFormR(true)} style={{...S.btnP,width:"100%",marginTop:8,justifyContent:"center"}}><Plus size={15}/>Nouveau rapport</button>}</div>}
 {tab==="projet"&&<ProjetsPersonnalises user={user} jeunes={[]} majeurs={[majeur]} projets={projets} onUpdate={onUpdateProjets} etabConfig={etabConfig} users={users} fixedId={majeur.id}/>}
      {tab==="stages"&&<StagesPanel sujet={majeur} user={user} users={users} onUpdate={onUpdateMajeur} etabConfig={etabConfig}/>}
      {tab==="incidents"&&<div>{showFormE?<div style={{...S.card,borderLeft:`4px solid ${C.orange}`,marginBottom:8}}><div style={{fontWeight:800,marginBottom:8,color:C.dark}}>Déclarer un événement</div><label style={{...S.lbl}}>Date</label><input type="date" value={eDate} onChange={e=>setEDate(e.target.value)} style={{...S.inp,marginBottom:8}}/><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}><div><label style={{...S.lbl}}>Type</label><select value={eType} onChange={e=>setEType(e.target.value)} style={{...S.inp}}><option value="incident">Incident</option><option value="plainte">Plainte</option><option value="reclamation">Réclamation</option></select></div><div><label style={{...S.lbl}}>Gravité</label><select value={eGrav} onChange={e=>setEGrav(e.target.value)} style={{...S.inp}}><option>Léger</option><option>Moyen</option><option>Grave</option></select></div></div><label style={{...S.lbl}}>Qualification</label><select value={eEig?"oui":"non"} onChange={e=>setEEig(e.target.value==="oui")} style={{...S.inp,marginBottom:8}}><option value="non">Événement indésirable simple</option><option value="oui">EIG — signalement obligatoire (art. L331-8-1 CASF)</option></select><label style={{...S.lbl}}>Titre</label><input value={eTitre} onChange={e=>setETitre(e.target.value)} placeholder="Titre" style={{...S.inp,marginBottom:8}}/><label style={{...S.lbl}}>Description</label><textarea value={eDesc} onChange={e=>setEDesc(e.target.value)} placeholder="Description..." rows={3} style={{...S.inp,marginBottom:8,resize:"vertical"}}/><div style={{display:"flex",gap:6,flexWrap:"wrap"}}><button onClick={submitE} style={{...S.btnP}}>Enregistrer</button><button onClick={()=>setShowFormE(false)} style={{...S.btnO}}>Annuler</button></div></div>:null}{me.length===0&&!showFormE?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun incident</div>:me.map(e=>{const gc=GC[e.gravite]||GC["Léger"];return(<div key={e.id} style={{...S.card,marginBottom:8,borderLeft:`4px solid ${gc.dot}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap"}}><span style={{fontWeight:800,color:C.dark}}>{e.titre}</span><div style={{display:"flex",gap:5,flexWrap:"wrap"}}>{e.eig&&<span style={{background:"#C62828",color:"#fff",borderRadius:4,padding:"2px 8px",fontSize:11.5,fontWeight:800}}>EIG</span>}<span style={{background:gc.bg,color:gc.text,borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:700}}>{GC[e.gravite]?e.gravite:"Léger"}</span></div></div><div style={{fontSize:12,color:C.mid,marginTop:4}}>{e.description}</div><div style={{fontSize:11.5,color:C.light,marginTop:4}}>{fmt(e.date)}{e.author&&" — par "+e.author}</div></div>);})}{!showFormE&&<button onClick={()=>setShowFormE(true)} style={{...S.btnO,width:"100%",marginTop:8,justifyContent:"center"}}><Plus size={15}/>Déclarer un événement</button>}</div>}
 {tab==="presences"&&<div>{Object.keys(mp).length===0?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucune présence enregistrée</div>:<div style={{...S.card}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><th style={{textAlign:"left",padding:6,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Date</th><th style={{textAlign:"center",padding:6,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Matin</th><th style={{textAlign:"center",padding:6,borderBottom:"2px solid #ddd",fontSize:12,color:C.light}}>Après-midi</th></tr></thead><tbody>{Object.entries(mp).sort(([a],[b])=>b.localeCompare(a)).map(([d,v])=><tr key={d}><td style={{padding:6,borderBottom:"1px solid #eee",fontSize:12}}>{d}</td><td style={{textAlign:"center",padding:6,borderBottom:"1px solid #eee"}}>{v.a?"✅":"❌"}</td><td style={{textAlign:"center",padding:6,borderBottom:"1px solid #eee"}}>{v.b?"✅":"❌"}</td></tr>)}</tbody></table></div>}</div>}
 </div>);
}

function RapportSite({user,rapportsSite,onSave,onDelete}){
  const[date,setDate]=useState(new Date().toISOString().slice(0,10));
  const[obs,setObs]=useState("");
  const[saved,setSaved]=useState(false);
  const canPickSite=user.site==="Tous"||(user.role==="chef_service"||user.role==="directeur");
  const[selSite,setSelSite]=useState(user.site==="Tous"?"Djilass":user.site);
  const site=canPickSite?selSite:(user.site||"Fatick");
  const existing=(rapportsSite||[]).filter(r=>r.site===site).sort((a,b)=>b.date.localeCompare(a.date));
  const todayR=existing.find(r=>r.date===date);
  useEffect(()=>{if(todayR)setObs(todayR.observation||"");},[date,todayR?.id]);
  const handleSave=()=>{if(!obs.trim())return;const now=new Date();const ts=now.toISOString();const r={id:todayR?todayR.id:Date.now(),site,date,observation:obs.trim(),author:user.name||"?",createdAt:todayR?.createdAt||ts,updatedAt:ts,horodatage:ts};onSave(r);setSaved(true);setTimeout(()=>setSaved(false),2000);};
  const exportSiteReport=async(r)=>{try{const{Document,Packer,Paragraph,TextRun,AlignmentType}=await import("docx");const doc=new Document({sections:[{properties:{},children:[new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:"RAPPORT DE SITE",bold:true,size:32,font:"Arial"})]}),new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:300},children:[new TextRun({text:"Association PDSR",size:24,font:"Arial",color:"B88608"})]}),new Paragraph({spacing:{after:100},children:[new TextRun({text:"Site : ",bold:true,size:22,font:"Arial"}),new TextRun({text:r.site,size:22,font:"Arial"})]}),new Paragraph({spacing:{after:100},children:[new TextRun({text:"Date : ",bold:true,size:22,font:"Arial"}),new TextRun({text:r.date,size:22,font:"Arial"})]}),new Paragraph({spacing:{after:100},children:[new TextRun({text:"Auteur : ",bold:true,size:22,font:"Arial"}),new TextRun({text:r.author||"?",size:22,font:"Arial"})]}),new Paragraph({spacing:{after:100},children:[new TextRun({text:"Horodatage : ",bold:true,size:22,font:"Arial"}),new TextRun({text:(r.horodatage||r.createdAt||"").replace("T"," ").slice(0,19),size:22,font:"Arial"})]}),new Paragraph({spacing:{before:300,after:100},children:[new TextRun({text:"Observation :",bold:true,size:24,font:"Arial",underline:{}})]}),new Paragraph({spacing:{after:200},children:[new TextRun({text:r.observation||"",size:22,font:"Arial"})]})]}]});const blob=await Packer.toBlob(doc);const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="RapportSite_"+r.site+"_"+r.date+".docx";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}catch(e){console.error("Export error:",e);alert("Erreur export: "+e.message);}};
  return(<div style={{padding:"18px 14px",maxWidth:700,margin:"0 auto"}}>
    <h2 style={{fontSize:18,fontWeight:900,color:C.dark,margin:"0 0 14px"}}>Rapport de site{site?" — "+site:""}</h2>
    <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"center",flexWrap:"wrap"}}>
      {canPickSite&&<><label style={{fontWeight:600,fontSize:13}}>Site :</label><select value={selSite} onChange={e=>setSelSite(e.target.value)} style={{padding:"6px 10px",borderRadius:8,border:"1.5px solid "+C.border,fontSize:13,fontFamily:"inherit"}}><option value="Djilass">Djilass</option><option value="Fatick">Fatick</option></select></>}
      <label style={{fontWeight:600,fontSize:13}}>Date :</label>
      <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{padding:"6px 10px",borderRadius:8,border:"1.5px solid "+C.border,fontSize:13}}/>
    </div>
    <div style={{...S.card,borderLeft:"4px solid #6A1B9A",marginBottom:14}}>
      <div style={{fontWeight:700,marginBottom:8,color:"#6A1B9A"}}>Observation du jour</div>
      <textarea value={obs} onChange={e=>setObs(e.target.value)} rows={8} style={{width:"100%",padding:"0.5rem",borderRadius:8,border:"1.5px solid "+C.border,fontFamily:"'Nunito',sans-serif",fontSize:"0.9rem",resize:"vertical"}} placeholder={"Décrivez la journée sur le camp "+site+"..."}/>
      <button onClick={handleSave} style={{marginTop:8,background:saved?"#4CAF50":"linear-gradient(135deg,#6A1B9A,#4A148C)",color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",cursor:"pointer",fontWeight:700,fontSize:13,fontFamily:"inherit"}}>{saved?"✔ Enregistré !":"Enregistrer"}</button>
    </div>
    <h3 style={{fontWeight:700,fontSize:15,marginBottom:10,color:C.dark}}>Historique — {site}</h3>
    {existing.length===0&&<div style={{textAlign:"center",padding:20,color:C.light}}>Aucun rapport de site pour {site}</div>}
    {existing.map(r=><div key={r.id} style={{...S.card,marginBottom:10,borderLeft:"3px solid #6A1B9A",position:"relative"}}>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:4,alignItems:"center",flexWrap:"wrap"}}>
        <span style={{fontWeight:700,color:C.dark}}>{r.date}</span>
        <span style={{fontSize:12,color:C.light}}>{r.author}</span>
      </div>
      <p style={{margin:"6px 0",fontSize:13,color:C.mid,whiteSpace:"pre-wrap",lineHeight:1.6}}>{r.observation}</p>
      {(r.horodatage||r.createdAt)&&<div style={{fontSize:11.5,color:C.light,marginTop:4}}>Horodatage : {(r.horodatage||r.createdAt||"").replace("T"," à ").slice(0,19)}</div>}
      <div style={{display:"flex",gap:8,marginTop:8,flexWrap:"wrap"}}>
        <button onClick={()=>exportSiteReport(r)} style={{padding:"5px 14px",borderRadius:7,border:"1px solid #6A1B9A",background:"#F3E5F5",color:"#6A1B9A",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Exporter Word</button>
        {(user.role==="chef_service"||user.role==="directeur")&&onDelete&&<button onClick={()=>{if(confirm("Supprimer ce rapport de site ?"))onDelete(r.id);}} style={{padding:"5px 14px",borderRadius:7,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Supprimer</button>}
      </div>
    </div>)}
  </div>);
}
function ExportPage({rapports,evenements,agenda,jeunes,majeurs,rapportsSite,onPurge,sejourConfig,purgeRanges,onCancelRange,peutPurger}){
const[dateFrom,setDateFrom]=useState("");
const[dateTo,setDateTo]=useState("");
const[exporting,setExporting]=useState(false);
const[done,setDone]=useState(false);
const[chkRapportsJ,setChkRapportsJ]=useState(true);
const[chkRapportsH,setChkRapportsH]=useState(true);
const[chkRapportsSite,setChkRapportsSite]=useState(true);
const[chkRdv,setChkRdv]=useState(true);
const[chkEvenements,setChkEvenements]=useState(true);
const[chkStages,setChkStages]=useState(true);
const[filterJournee,setFilterJournee]=useState(true);
const[filterTelParents,setFilterTelParents]=useState(true);
const[filterTelExt,setFilterTelExt]=useState(true);
const allJ=[...(jeunes||JEUNES),...(majeurs||MAJEURS)];
const jName=(id)=>{const j=allJ.find(x=>x.id===id);return j?(j.prenom+" "+(j.nom||"")):"ID:"+id;};
const jSite=(id)=>{const j=allJ.find(x=>x.id===id);return j?(j.site||""):"";};
const SEJ_DEBUT={Djilass:(sejourConfig&&sejourConfig.Djilass&&sejourConfig.Djilass.dateDebut)||"2026-03-30",Fatick:(sejourConfig&&sejourConfig.Fatick&&sejourConfig.Fatick.dateDebut)||"2026-03-16"};
const sejWeek=(dateStr,site)=>{const d0=new Date((SEJ_DEBUT[site]||SEJ_DEBUT.Djilass)+"T00:00:00");const d=new Date(dateStr+"T00:00:00");if(isNaN(d)||isNaN(d0))return"";const sw=Math.floor((d-d0)/604800000)+1;return sw>0?("S"+String(sw).padStart(2,"0")):"avant séjour";};
const allChecked=chkRapportsJ&&chkRapportsH&&chkRapportsSite&&chkRdv&&chkEvenements;
const toggleAll=()=>{const v=!allChecked;setChkRapportsJ(v);setChkRapportsH(v);setChkRapportsSite(v);setChkRdv(v);setChkEvenements(v);};
const TC_LABELS={"journee":"Journée du jeune","tel_parents":"RDV tél. parents","tel_ext":"RDV tél. contact extérieur"};
const fR=(rapports||[]).filter(r=>r.date>=dateFrom&&r.date<=dateTo);
const fRFiltered=fR.filter(r=>{const tc=r.typeContact||"journee";if(tc==="journee"&&!filterJournee)return false;if(tc==="tel_parents"&&!filterTelParents)return false;if(tc==="tel_ext"&&!filterTelExt)return false;return true;});
const doExport=async()=>{
if(!dateFrom||!dateTo)return alert("Sélectionnez une période");
if(!chkRapportsJ&&!chkRapportsH&&!chkRapportsSite&&!chkRdv&&!chkEvenements&&!chkStages)return alert("Sélectionnez au moins un type de données");
setExporting(true);
try{
const XLSX=await loadXLSX();
const wb=XLSX.utils.book_new();
if(chkRapportsJ){
  const sorted=fRFiltered.sort((a,b)=>a.date.localeCompare(b.date));
  const rowsR=[["Date","Site","Semaine séjour","Jeune","Type de contact","Observation","Horodatage","Auteur"]];
  sorted.forEach(r=>{const st=jSite(r.jeuneId);rowsR.push([r.date,st,sejWeek(r.date,st),jName(r.jeuneId),TC_LABELS[r.typeContact||"journee"]||r.typeContact||"Journée",r.observation||"",r.horodatage||r.createdAt||"",r.author||""]);});
  const ws=XLSX.utils.aoa_to_sheet(rowsR);
  ws["!cols"]=[{wch:12},{wch:10},{wch:12},{wch:25},{wch:28},{wch:60},{wch:20},{wch:20}];
  ws["!autofilter"]={ref:"A1:H"+rowsR.length};
  XLSX.utils.book_append_sheet(wb,ws,"Rapports journaliers");
}
if(chkRapportsH){
  const getISOWeek=(dateStr)=>{const d=new Date(dateStr);d.setHours(0,0,0,0);d.setDate(d.getDate()+3-(d.getDay()+6)%7);const w1=new Date(d.getFullYear(),0,4);return String(1+Math.round(((d-w1)/86400000-3+(w1.getDay()+6)%7)/7)).padStart(2,"0")};
  const weekData={};
  fR.forEach(r=>{const st=jSite(r.jeuneId);const wk=sejWeek(r.date,st);const jn=jName(r.jeuneId);const key=wk+"|"+st;if(!weekData[key])weekData[key]={};if(!weekData[key][jn])weekData[key][jn]=[];weekData[key][jn].push(r);});
  const rowsH=[["Semaine séjour","Site","Jeune","Nb rapports","Résumé semaine"]];
  Object.keys(weekData).sort().forEach(key=>{const[wk,st]=key.split("|");Object.keys(weekData[key]).sort().forEach(jn=>{const rs=weekData[key][jn];const summary=rs.sort((a,b)=>a.date.localeCompare(b.date)).map(r=>{const dt=new Date(r.date);const dn=dt.toLocaleDateString("fr-FR",{weekday:"short"});return dn+": "+(r.observation||"").substring(0,80);}).join(" | ");rowsH.push([wk,st,jn,rs.length,summary]);});});
  const ws=XLSX.utils.aoa_to_sheet(rowsH);
  ws["!cols"]=[{wch:13},{wch:10},{wch:25},{wch:12},{wch:80}];
  ws["!autofilter"]={ref:"A1:E"+rowsH.length};
  XLSX.utils.book_append_sheet(wb,ws,"Rapports hebdos");
}
if(chkRapportsSite){
  const fRS=(rapportsSite||[]).filter(r=>r.date>=dateFrom&&r.date<=dateTo).sort((a,b)=>a.date.localeCompare(b.date));
  const rowsRS=[["Date","Site","Auteur","Contenu"]];
  fRS.forEach(r=>rowsRS.push([r.date,r.site||"",r.author||"",r.observation||r.contenu||r.text||""]));
  const ws=XLSX.utils.aoa_to_sheet(rowsRS);
  ws["!cols"]=[{wch:12},{wch:15},{wch:20},{wch:80}];
  ws["!autofilter"]={ref:"A1:D"+rowsRS.length};
  XLSX.utils.book_append_sheet(wb,ws,"Rapports de site");
}
if(chkRdv){
  const fA=(agenda||[]).filter(a=>a.date>=dateFrom&&a.date<=dateTo).sort((a,b)=>a.date.localeCompare(b.date));
  const rowsA=[["Date","Heure","Site","Jeune","Type RDV","Lieu","Interlocuteur","Notes","Créé par"]];
  fA.forEach(a=>rowsA.push([a.date,a.heure||"",a.jeuneId?jSite(a.jeuneId):"",a.jeuneNom||jName(a.jeuneId),a.type||"",a.lieu||"",a.interlocuteur||"",a.notes||"",a.createdBy||""]));
  const ws=XLSX.utils.aoa_to_sheet(rowsA);
  ws["!cols"]=[{wch:12},{wch:8},{wch:10},{wch:25},{wch:15},{wch:25},{wch:25},{wch:50},{wch:20}];
  ws["!autofilter"]={ref:"A1:I"+rowsA.length};
  XLSX.utils.book_append_sheet(wb,ws,"Rendez-vous");
}
if(chkEvenements){
  const fE=(evenements||[]).filter(e=>e.date>=dateFrom&&e.date<=dateTo).sort((a,b)=>a.date.localeCompare(b.date));
  const rowsE=[["Date","Site","Jeune","Type","Titre","Description","Gravité","EIG","Transmission EIG","N° Suivi","Catégorie","Horodatage","Auteur"]];
  fE.forEach(e=>{const st=e.jeuneId?jSite(e.jeuneId):(e.site||"");rowsE.push([e.date,st,e.jeuneId?jName(e.jeuneId):"(équipe)",e.type||"événement",e.titre||"",e.description||"",e.gravite||"",e.eig?"OUI":"",e.eig?(e.eigData&&e.eigData.dateTransmission?("Transmis le "+e.eigData.dateTransmission+(e.eigData.accuseReception?" (AR reçu)":" (AR en attente)")):"NON TRANSMIS"):"",e.numeroSuivi||"",e.categorie||"",e.horodatage||"",e.author||""]);});
  const ws=XLSX.utils.aoa_to_sheet(rowsE);
  ws["!cols"]=[{wch:12},{wch:10},{wch:25},{wch:18},{wch:25},{wch:60},{wch:12},{wch:8},{wch:26},{wch:14},{wch:18},{wch:20},{wch:20}];
  ws["!autofilter"]={ref:"A1:M"+rowsE.length};
  XLSX.utils.book_append_sheet(wb,ws,"Événements indésirables");
}
if(chkStages){
  const rowsS=[["Jeune","Site","Intitulé","Structure d'accueil","Lieu","Tuteur","Tél. tuteur","Référent éducatif","Début","Fin","H/sem.","Statut","Nb appréciations","Dernière appréciation","Assiduité","Comportement","Gestes pro.","Autonomie","Intégration","Commentaire du tuteur","Bilan de fin de stage"]];
  [...(jeunes||[]),...(majeurs||[])].forEach(pj=>((pj.stages)||[]).forEach(st=>{
    const apps=(st.appreciations)||[];const last=apps.length?apps[apps.length-1]:null;const n=(last&&last.notes)||{};
    rowsS.push([(pj.prenom||"")+" "+(pj.nom||""),pj.site||"",st.intitule||"",st.structure||"",st.lieu||"",st.tuteur||"",st.telTuteur||"",st.referent||"",normDate(st.dateDebut),normDate(st.dateFin),st.heures||"",st.statut||"",apps.length,last?last.date:"",n.assiduite||"",n.comportement||"",n.technique||"",n.autonomie||"",n.integration||"",last?(last.texte||""):"",st.bilan||""]);}));
  const ws=XLSX.utils.aoa_to_sheet(rowsS);
  ws["!cols"]=[{wch:24},{wch:10},{wch:24},{wch:26},{wch:20},{wch:18},{wch:14},{wch:18},{wch:12},{wch:12},{wch:8},{wch:12},{wch:8},{wch:12},{wch:16},{wch:16},{wch:16},{wch:16},{wch:16},{wch:50},{wch:60}];
  ws["!autofilter"]={ref:"A1:U"+rowsS.length};
  ws["!freeze"]={xSplit:1,ySplit:1};
  XLSX.utils.book_append_sheet(wb,ws,"Stages");
}
{
  const nomsFeuilles=wb.SheetNames.slice();
  const syn=[["EXPORT PDSR"],[],["Période",fmt(dateFrom)+" au "+fmt(dateTo)],["Édité le",fmt(isoToday())],["Sites","Fatick et Djilass"],[],["Feuille","Lignes de données"]];
  nomsFeuilles.forEach(n=>{const f=wb.Sheets[n];const r=XLSX.utils.decode_range(f["!ref"]||"A1:A1");syn.push([n,Math.max(0,r.e.r)]);});
  syn.push([]);syn.push(["Chaque feuille a un filtre sur la ligne d'en-tête et la première ligne figée."]);
  const wsSyn=XLSX.utils.aoa_to_sheet(syn);
  wsSyn["!cols"]=[{wch:30},{wch:34}];
  XLSX.utils.book_append_sheet(wb,wsSyn,"Sommaire");
  wb.SheetNames=["Sommaire",...nomsFeuilles];
  nomsFeuilles.forEach(n=>{const f=wb.Sheets[n];if(!f["!freeze"])f["!freeze"]={xSplit:0,ySplit:1};});
}
XLSX.writeFile(wb,"PDSR_Export_"+dateFrom+"_"+dateTo+".xlsx");
setDone(true);
}catch(err){alert("Erreur export: "+err.message);}
setExporting(false);
};
const doPurge0=()=>{
  const doPurge=()=>{if(!peutPurger){alert("Seul l'administrateur peut purger.");return;}doPurge0();};
if(!dateFrom||!dateTo)return alert("Sélectionnez une période");
const scope={rapports:chkRapportsJ,evenements:chkEvenements,agenda:chkRdv,rapportsSite:chkRapportsSite};
if(!scope.rapports&&!scope.evenements&&!scope.agenda&&!scope.rapportsSite)return alert("Cochez au moins une catégorie à purger.");
const inP=(arr)=>(arr||[]).filter(x=>x&&x.date>=dateFrom&&x.date<=dateTo).length;
const parts=[];
if(scope.rapports)parts.push(inP(rapports)+" rapport(s) journalier(s)");
if(scope.evenements)parts.push(inP(evenements)+" événement(s)");
if(scope.agenda)parts.push(inP(agenda)+" RDV");
if(scope.rapportsSite)parts.push(inP(rapportsSite)+" rapport(s) de site");
if(!confirm("Purger UNIQUEMENT :\n• "+parts.join("\n• ")+"\n\nPériode du "+dateFrom+" au "+dateTo+". Le reste n'est pas touché. Action irréversible. Confirmer ?"))return;
onPurge(dateFrom,dateTo,scope);
setDone(false);
alert("Purge effectuée : "+parts.join(", ")+".");
};
const Chk=({checked,onChange,label,count})=><label style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",background:checked?"#FDF3D0":"#f9f9f9",border:checked?"1.5px solid "+C.gold:"1.5px solid #e5e7eb",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600,color:checked?C.dark:"#666",transition:"all 0.15s",flexWrap:"wrap"}}><input type="checkbox" checked={checked} onChange={onChange} style={{accentColor:C.gold,width:16,height:16}}/>{label}{count!==undefined&&<span style={{fontSize:12,color:C.light,fontWeight:400}}>({count})</span>}</label>;
return(<div style={{maxWidth:600,margin:"0 auto"}}>
<h2 style={{fontSize:18,fontWeight:900,color:C.dark,marginBottom:16}}>Export Excel avancé</h2>
<div style={{...S.card,marginBottom:16}}>
<div style={{fontWeight:700,fontSize:14,color:C.dark,marginBottom:12}}>Période</div>
<div style={{display:"flex",gap:12,marginBottom:16,flexWrap:"wrap"}}>
<div style={{flex:1}}><label style={{fontSize:12,fontWeight:700,color:C.light,display:"block",marginBottom:4}}>Du</label><input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} style={{...S.inp,width:"100%"}}/></div>
<div style={{flex:1}}><label style={{fontSize:12,fontWeight:700,color:C.light,display:"block",marginBottom:4}}>Au</label><input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} style={{...S.inp,width:"100%"}}/></div>
</div>
</div>
<div style={{...S.card,marginBottom:16}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap"}}>
  <div style={{fontWeight:700,fontSize:14,color:C.dark}}>Données à exporter</div>
  <button onClick={toggleAll} style={{fontSize:12,fontWeight:700,color:C.gold,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>{allChecked?"Tout désélectionner":"Tout sélectionner"}</button>
</div>
<div style={{display:"flex",flexDirection:"column",gap:8}}>
  <Chk checked={chkRapportsJ} onChange={()=>setChkRapportsJ(!chkRapportsJ)} label="Rapports journaliers" count={dateFrom&&dateTo?fRFiltered.length:undefined}/>
  {chkRapportsJ&&<div style={{marginLeft:28,display:"flex",flexWrap:"wrap",gap:6}}>
    <label style={{display:"flex",alignItems:"center",gap:4,fontSize:12,color:C.dark,cursor:"pointer",flexWrap:"wrap"}}><input type="checkbox" checked={filterJournee} onChange={()=>setFilterJournee(!filterJournee)} style={{accentColor:C.gold}}/>Journée du jeune</label>
    <label style={{display:"flex",alignItems:"center",gap:4,fontSize:12,color:C.dark,cursor:"pointer",flexWrap:"wrap"}}><input type="checkbox" checked={filterTelParents} onChange={()=>setFilterTelParents(!filterTelParents)} style={{accentColor:C.gold}}/>RDV tél. parents</label>
    <label style={{display:"flex",alignItems:"center",gap:4,fontSize:12,color:C.dark,cursor:"pointer",flexWrap:"wrap"}}><input type="checkbox" checked={filterTelExt} onChange={()=>setFilterTelExt(!filterTelExt)} style={{accentColor:C.gold}}/>RDV tél. contact ext.</label>
  </div>}
  <Chk checked={chkRapportsH} onChange={()=>setChkRapportsH(!chkRapportsH)} label="Rapports hebdos (compilation)" count={dateFrom&&dateTo?fR.length:undefined}/>
  <Chk checked={chkRapportsSite} onChange={()=>setChkRapportsSite(!chkRapportsSite)} label="Rapports de site" count={dateFrom&&dateTo?(rapportsSite||[]).filter(r=>r.date>=dateFrom&&r.date<=dateTo).length:undefined}/>
  <Chk checked={chkRdv} onChange={()=>setChkRdv(!chkRdv)} label="Rendez-vous" count={dateFrom&&dateTo?(agenda||[]).filter(a=>a.date>=dateFrom&&a.date<=dateTo).length:undefined}/>
  <Chk checked={chkEvenements} onChange={()=>setChkEvenements(!chkEvenements)} label="Événements indésirables" count={dateFrom&&dateTo?(evenements||[]).filter(e=>e.date>=dateFrom&&e.date<=dateTo).length:undefined}/>
  <Chk checked={chkStages} onChange={()=>setChkStages(!chkStages)} label="Stages et bilans" count={[...(jeunes||[]),...(majeurs||[])].reduce((n,p2)=>n+((p2.stages)||[]).length,0)}/>
</div>
</div>
<div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
<button onClick={doExport} disabled={exporting} style={{flex:1,background:C.gold,color:"#fff",border:"none",borderRadius:8,padding:"12px 20px",fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{exporting?"Export en cours...":"Exporter en Excel (.xlsx)"}</button>
</div>
{done&&<div style={{marginTop:16,padding:12,background:"rgba(102,187,106,0.15)",borderRadius:8,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap"}}>
<span style={{color:"#2E7D32",fontWeight:700,fontSize:13}}>Export terminé !</span>
{peutPurger&&<button onClick={doPurge} style={{background:"#C62828",color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Purger cette période</button>}
{!peutPurger&&<span style={{fontSize:11.5,color:C.light,fontWeight:600}}>La purge est réservée à l'administrateur.</span>}
</div>}
{Array.isArray(purgeRanges)&&purgeRanges.length>0&&<div style={{...S.card,marginTop:16,borderLeft:"4px solid #C62828"}}>
  <div style={{fontWeight:800,color:C.dark,fontSize:14,marginBottom:4}}>Purges actives</div>
  <div style={{fontSize:11.5,color:C.light,marginBottom:10}}>Les éléments datés dans ces plages et antérieurs à la purge sont masqués sur tous les appareils. Annuler une purge réaffiche ce qui existe encore dans le cloud.</div>
  {purgeRanges.map(r=>{const L={rapport:"Rapports",evenement:"Événements",agenda:"RDV",rapportSite:"Rapports de site"};const fr=(r.from||"").split("-").reverse().join("/");const to=(r.to||"").split("-").reverse().join("/");return(
    <div key={r.ts} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:"1px solid "+C.border,flexWrap:"wrap"}}>
      <div style={{flex:1,minWidth:180}}>
        <div style={{fontWeight:800,fontSize:12.5,color:C.dark}}>{fr} → {to}</div>
        <div style={{fontSize:12,color:C.light,fontWeight:600}}>{(r.t||[]).map(x=>L[x]||x).join(", ")} · purgé le {(r.ts||"").slice(0,10).split("-").reverse().join("/")}</div>
      </div>
      <button disabled={!peutPurger} title={peutPurger?"":"Réservé à l'administrateur"} onClick={()=>{if(!peutPurger){alert("Seul l'administrateur peut annuler une purge.");return;}if(confirm("Annuler cette purge ? Les éléments encore présents dans le cloud réapparaîtront."))onCancelRange(r.ts);}} style={{padding:"6px 13px",borderRadius:8,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontWeight:800,fontSize:11.5,cursor:"pointer",fontFamily:"inherit"}}>Annuler cette purge</button>
    </div>);})}
</div>}
</div>);
}

function SignaturePad({onValidate,onCancel}){
  const ref=useRef(null);const drawing=useRef(false);const last=useRef(null);const touched=useRef(false);
  const pos=(e)=>{const c=ref.current;const r=c.getBoundingClientRect();const t=e.touches&&e.touches[0]?e.touches[0]:e;return{x:(t.clientX-r.left)*(c.width/r.width),y:(t.clientY-r.top)*(c.height/r.height)};};
  const start=(e)=>{e.preventDefault();drawing.current=true;touched.current=true;last.current=pos(e);};
  const move=(e)=>{if(!drawing.current)return;e.preventDefault();const c=ref.current;const ctx=c.getContext("2d");const p=pos(e);ctx.strokeStyle="#1a1a1a";ctx.lineWidth=2.5;ctx.lineCap="round";ctx.beginPath();ctx.moveTo(last.current.x,last.current.y);ctx.lineTo(p.x,p.y);ctx.stroke();last.current=p;};
  const end=()=>{drawing.current=false;};
  const clear=()=>{const c=ref.current;c.getContext("2d").clearRect(0,0,c.width,c.height);touched.current=false;};
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:80,display:"flex",alignItems:"center",justifyContent:"center",padding:16,flexWrap:"wrap"}}>
    <div style={{background:C.white,borderRadius:16,padding:18,width:"100%",maxWidth:440,boxShadow:C.shadowLg}}>
      <div style={{fontWeight:800,fontSize:15,color:C.dark,marginBottom:4}}>Signature</div>
      <div style={{fontSize:12,color:C.light,marginBottom:10}}>Signez dans le cadre avec le doigt ou la souris.</div>
      <canvas ref={ref} width={400} height={180} style={{width:"100%",height:180,border:"1.5px dashed "+C.gold,borderRadius:10,background:C.sableLight,touchAction:"none",cursor:"crosshair"}} onMouseDown={start} onMouseMove={move} onMouseUp={end} onMouseLeave={end} onTouchStart={start} onTouchMove={move} onTouchEnd={end}/>
      <div style={{display:"flex",gap:8,marginTop:12,flexWrap:"wrap"}}>
        <button onClick={clear} style={{...S.btnO,flex:1,justifyContent:"center"}}>Effacer</button>
        <button onClick={onCancel} style={{...S.btnS,flex:1,justifyContent:"center"}}>Annuler</button>
        <button onClick={()=>{if(!touched.current){alert("Veuillez signer d'abord.");return;}onValidate(ref.current.toDataURL("image/png"));}} style={{...S.btnP,flex:1,justifyContent:"center"}}><Check size={15}/>Valider</button>
      </div>
    </div>
  </div>);
}

function Intendance({user,items,onSave,onDelete}){
  const isChef=user.role==="chef_service"||user.role==="directeur";
  const[tab,setTab]=useState("besoins");
  const[site,setSite]=useState(isChef?"Tous":user.site);
  const mySite=isChef?site:user.site;
  const[cat,setCat]=useState("Produits d'hygiène");const[desc,setDesc]=useState("");const[qte,setQte]=useState("");const[dLiv,setDLiv]=useState("");const[err,setErr]=useState("");
  const[objR,setObjR]=useState("");const[descR,setDescR]=useState("");const[urg,setUrg]=useState("Normale");
  const CATS=["Produits d'hygiène","Poches d'eau","Régie","Alimentation","Matériel pédagogique","Autre besoin"];
  const BST={en_attente:{l:"En attente",bg:"#FFF3E0",c:"#E65100"},validee:{l:"Validée",bg:"#E3F2FD",c:"#1565C0"},livree:{l:"Livrée",bg:"#E8F5E9",c:"#2E7D32"},refusee:{l:"Refusée",bg:"#FFEBEE",c:"#C62828"}};
  const RST={a_reparer:{l:"À réparer",bg:"#FFEBEE",c:"#C62828"},en_cours:{l:"En cours",bg:"#FFF3E0",c:"#E65100"},repare:{l:"Réparé",bg:"#E8F5E9",c:"#2E7D32"}};
  const list=(items||[]).filter(i=>i&&(mySite==="Tous"||i.site===mySite));
  const besoins=list.filter(i=>i.type==="besoin").sort((a,b)=>(b.createdAt||"").localeCompare(a.createdAt||""));
  const reps=list.filter(i=>i.type==="reparation").sort((a,b)=>(b.createdAt||"").localeCompare(a.createdAt||""));
  const submitBesoin=()=>{
    setErr("");
    if(!desc.trim()){setErr("Décrivez le besoin.");return;}
    if(!dLiv){setErr("Indiquez la date de livraison souhaitée.");return;}
    const diff=new Date(dLiv+"T00:00:00").getTime()-Date.now();
    if(diff<48*3600*1000){setErr("Anticipation insuffisante : la date de livraison doit être au moins 48 h à l'avance.");return;}
    const siteVal=isChef?(site==="Tous"?"Fatick":site):user.site;
    onSave({id:"int_"+Date.now(),type:"besoin",site:siteVal,categorie:cat,description:desc.trim(),quantite:qte.trim(),dateLivraison:dLiv,statut:"en_attente",auteur:user.name,createdAt:new Date().toISOString()});
    setDesc("");setQte("");setDLiv("");
  };
  const submitRep=()=>{
    setErr("");
    if(!objR.trim()){setErr("Indiquez l'objet à réparer.");return;}
    const siteVal=isChef?(site==="Tous"?"Fatick":site):user.site;
    onSave({id:"int_"+Date.now(),type:"reparation",site:siteVal,objet:objR.trim(),description:descR.trim(),urgence:urg,statut:"a_reparer",auteur:user.name,createdAt:new Date().toISOString()});
    setObjR("");setDescR("");setUrg("Normale");
  };
  const chip=(st,map)=><span style={{background:map[st]?.bg||"#eee",color:map[st]?.c||"#555",fontSize:12,fontWeight:800,padding:"3px 10px",borderRadius:8}}>{map[st]?.l||st}</span>;
  const minDate=new Date(Date.now()+48*3600*1000).toISOString().slice(0,10);
  return(<div>
    <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}}>
      <h2 style={{margin:0,fontSize:20,color:C.dark,fontWeight:800}}>Intendance</h2>
      {isChef&&<select value={site} onChange={e=>setSite(e.target.value)} style={{...S.inp,width:"auto",padding:"8px 12px"}}><option>Tous</option><option>Djilass</option><option>Fatick</option></select>}
      {!isChef&&<span style={{fontSize:13,color:C.light,fontWeight:700}}>Site : {user.site}</span>}
    </div>
    <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
      <button onClick={()=>setTab("besoins")} style={{...(tab==="besoins"?S.btnP:S.btnS)}}><Package size={16}/>Besoins ({besoins.length})</button>
      <button onClick={()=>setTab("reparations")} style={{...(tab==="reparations"?S.btnP:S.btnS)}}><Wrench size={16}/>Réparations ({reps.length})</button>
    </div>
    {err&&<div style={{background:"#FFEBEE",color:"#C62828",fontWeight:700,fontSize:13,padding:"10px 14px",borderRadius:10,marginBottom:12}}>{err}</div>}
    {tab==="besoins"&&<div>
      <div style={S.card}>
        <div style={{fontWeight:800,color:C.dark,fontSize:15,marginBottom:12}}>Nouvelle demande</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={S.lbl}>Catégorie</label><select value={cat} onChange={e=>setCat(e.target.value)} style={S.inp}>{CATS.map(c=><option key={c}>{c}</option>)}</select></div>
          <div><label style={S.lbl}>Quantité</label><input value={qte} onChange={e=>setQte(e.target.value)} placeholder="ex : 20 poches, 5 bidons…" style={S.inp}/></div>
        </div>
        <div style={{marginBottom:10}}><label style={S.lbl}>Description du besoin</label><textarea value={desc} onChange={e=>setDesc(e.target.value)} rows={3} style={{...S.inp,resize:"vertical"}} placeholder="Détaillez le besoin (produits, montant régie, etc.)"/></div>
        <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
          <div><label style={S.lbl}>Date de livraison souhaitée (min. 48 h)</label><input type="date" min={minDate} value={dLiv} onChange={e=>setDLiv(e.target.value)} style={S.inp}/></div>
          <button onClick={submitBesoin} style={S.btnP}><Plus size={16}/>Envoyer la demande</button>
        </div>
      </div>
      {besoins.map(b=><div key={b.id} style={{...S.card,padding:"14px 18px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:10,flexWrap:"wrap"}}>
          <div style={{flex:1,minWidth:200}}>
            <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:4}}><span style={{fontWeight:800,color:C.dark,fontSize:14}}>{b.categorie}</span>{chip(b.statut,BST)}<span style={{fontSize:12,color:C.light,fontWeight:700}}>{b.site}</span></div>
            <div style={{fontSize:13,color:C.mid,marginBottom:4}}>{b.description}{b.quantite?" — "+b.quantite:""}</div>
            <div style={{fontSize:12,color:C.light}}>Livraison souhaitée : <b>{b.dateLivraison}</b> · demandé par {b.auteur} le {(b.createdAt||"").slice(0,10)}</div>
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {isChef&&b.statut==="en_attente"&&<><button onClick={()=>onSave({...b,statut:"validee"})} style={{...S.btnS,padding:"7px 12px",fontSize:12,color:"#1565C0",borderColor:"#1565C0"}}>Valider</button><button onClick={()=>onSave({...b,statut:"refusee"})} style={{...S.btnS,padding:"7px 12px",fontSize:12,color:"#C62828",borderColor:"#C62828"}}>Refuser</button></>}
            {isChef&&b.statut==="validee"&&<button onClick={()=>onSave({...b,statut:"livree"})} style={{...S.btnS,padding:"7px 12px",fontSize:12,color:"#2E7D32",borderColor:"#2E7D32"}}>Marquer livrée</button>}
            {(isChef||b.auteur===user.name)&&<button onClick={()=>{if(confirm("Supprimer cette demande ?"))onDelete(b.id);}} style={{...S.btnS,padding:"7px 12px",fontSize:12}}>Suppr.</button>}
          </div>
        </div>
      </div>)}
      {besoins.length===0&&<div style={{textAlign:"center",color:C.light,fontSize:13,padding:20}}>Aucune demande pour l'instant.</div>}
    </div>}
    {tab==="reparations"&&<div>
      <div style={S.card}>
        <div style={{fontWeight:800,color:C.dark,fontSize:15,marginBottom:12}}>Signaler une réparation</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
          <div><label style={S.lbl}>Objet</label><input value={objR} onChange={e=>setObjR(e.target.value)} placeholder="ex : toilettes bloc B, lit chambre 4…" style={S.inp}/></div>
          <div><label style={S.lbl}>Urgence</label><select value={urg} onChange={e=>setUrg(e.target.value)} style={S.inp}><option>Normale</option><option>Urgente</option></select></div>
        </div>
        <div style={{marginBottom:10}}><label style={S.lbl}>Détail</label><textarea value={descR} onChange={e=>setDescR(e.target.value)} rows={2} style={{...S.inp,resize:"vertical"}} placeholder="Décrivez le problème…"/></div>
        <button onClick={submitRep} style={S.btnP}><Plus size={16}/>Signaler</button>
      </div>
      {reps.map(r=><div key={r.id} style={{...S.card,padding:"14px 18px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:10,flexWrap:"wrap"}}>
          <div style={{flex:1,minWidth:200}}>
            <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:4}}><span style={{fontWeight:800,color:C.dark,fontSize:14}}>{r.objet}</span>{chip(r.statut,RST)}{r.urgence==="Urgente"&&<span className="pulse-urg" style={{background:"#C62828",color:"#fff",fontSize:11.5,fontWeight:900,padding:"2px 8px",borderRadius:6}}>URGENT</span>}<span style={{fontSize:12,color:C.light,fontWeight:700}}>{r.site}</span></div>
            {r.description&&<div style={{fontSize:13,color:C.mid,marginBottom:4}}>{r.description}</div>}
            <div style={{fontSize:12,color:C.light}}>Signalé par {r.auteur} le {(r.createdAt||"").slice(0,10)}</div>
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {r.statut==="a_reparer"&&<button onClick={()=>onSave({...r,statut:"en_cours"})} style={{...S.btnS,padding:"7px 12px",fontSize:12,color:"#E65100",borderColor:"#E65100"}}>En cours</button>}
            {r.statut!=="repare"&&<button onClick={()=>onSave({...r,statut:"repare"})} style={{...S.btnS,padding:"7px 12px",fontSize:12,color:"#2E7D32",borderColor:"#2E7D32"}}>Réparé</button>}
            {(isChef||r.auteur===user.name)&&<button onClick={()=>{if(confirm("Supprimer ce signalement ?"))onDelete(r.id);}} style={{...S.btnS,padding:"7px 12px",fontSize:12}}>Suppr.</button>}
          </div>
        </div>
      </div>)}
      {reps.length===0&&<div style={{textAlign:"center",color:C.light,fontSize:13,padding:20}}>Rien à réparer pour l'instant.</div>}
    </div>}
  </div>);
}
function PresEduc({user,users,entries,onSave,onDelete}){
  const isChef=user.role==="chef_service"||user.role==="directeur";
  const today=new Date().toISOString().slice(0,10);
  const[date,setDate]=useState(today);
  const[site,setSite]=useState(isChef?"Fatick":user.site);
  const mySite=isChef?site:user.site;
  const educs=(users||[]).filter(u=>u&&u.role==="educateur"&&!u.disabled&&(u.site===mySite||u.site==="Tous"));
  const STATUTS=["Présent","Absent","Maladie","Retard","Congé","Repos"];
  const SC={"Présent":"#2E7D32","Absent":"#C62828","Maladie":"#E65100","Retard":"#F9A825","Congé":"#1565C0","Repos":"#6A1B9A"};
  const presList=(entries||[]).filter(e=>e&&e.kind==="presence");
  const[regMois,setRegMois]=useState(today.slice(0,7));
  const regJours=useMemo(()=>{const[y,m]=regMois.split("-").map(Number);const n=new Date(y,m,0).getDate();return Array.from({length:n},(_,i)=>regMois+"-"+String(i+1).padStart(2,"0"));},[regMois]);
  const regIdx=useMemo(()=>{const idx={};presList.forEach(p=>{if(p.site===mySite&&p.date&&p.date.slice(0,7)===regMois)idx[p.date+"|"+p.educName]=p;});return idx;},[presList,mySite,regMois]);
  const exportRegCSV=()=>{const head=["Éducateur",...regJours.map(d=>d.slice(8)),"Présent","Absent/Maladie"];const rows=educs.map(ed=>{const line=regJours.map(d=>{const x=regIdx[d+"|"+ed.name];return x?x.statut:"";});return[ed.name,...line,line.filter(x=>x==="Présent").length,line.filter(x=>x==="Absent"||x==="Maladie").length];});const csv=[head,...rows].map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(";")).join("\n");const blob=new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="registre_educateurs_"+mySite+"_"+regMois+".csv";a.click();URL.revokeObjectURL(a.href);};

  const getP=(name)=>presList.find(p=>p.date===date&&p.educName===name&&p.site===mySite);
  const setP=(name,statut)=>{const ex=getP(name);onSave({id:ex?ex.id:"se_"+date+"_"+mySite+"_"+name.replace(/\W/g,""),kind:"presence",date,educName:name,site:mySite,statut,note:ex?ex.note:"",auteur:user.name,createdAt:ex?ex.createdAt:new Date().toISOString()});};
  const setNote=(name,note)=>{const ex=getP(name);if(!ex)return;onSave({...ex,note});};
  const[rEduc,setREduc]=useState("");const[rType,setRType]=useState("Maladie");const[rDate,setRDate]=useState(today);const[rTexte,setRTexte]=useState("");const[rErr,setRErr]=useState("");
  const rapports=(entries||[]).filter(e=>e&&e.kind==="rapport"&&(mySite==="Tous"||e.site===mySite)).sort((a,b)=>(b.date||"").localeCompare(a.date||""));
  const submitR=()=>{setRErr("");if(!rEduc){setRErr("Choisissez un éducateur.");return;}if(!rTexte.trim()){setRErr("Rédigez le rapport.");return;}onSave({id:"se_r_"+Date.now(),kind:"rapport",date:rDate,educName:rEduc,site:mySite,type:rType,note:rTexte.trim(),auteur:user.name,createdAt:new Date().toISOString()});setRTexte("");};
  return(<div>
    <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}}>
      <h2 style={{margin:0,fontSize:20,color:C.dark,fontWeight:800}}>Présences éducateurs</h2>
      <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{...S.inp,width:"auto",padding:"8px 12px"}}/>
      {isChef&&<select value={site} onChange={e=>setSite(e.target.value)} style={{...S.inp,width:"auto",padding:"8px 12px"}}><option>Djilass</option><option>Fatick</option></select>}
      {!isChef&&<span style={{fontSize:13,color:C.light,fontWeight:700}}>Site : {user.site}</span>}
    </div>
    <div style={S.card}>
      <div style={{fontWeight:800,color:C.dark,fontSize:15,marginBottom:12}}>Registre du {date.split("-").reverse().join("/")}</div>
      {educs.length===0&&<div style={{color:C.light,fontSize:13}}>Aucun éducateur rattaché à ce site.</div>}
      {educs.map(ed=>{const p=getP(ed.name);return(<div key={ed.id} style={{borderBottom:"1px solid "+C.border,padding:"10px 0"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <span style={{fontWeight:800,color:C.dark,fontSize:14,minWidth:110}}>{ed.name}</span>
          {STATUTS.map(st=><button key={st} onClick={()=>setP(ed.name,st)} style={{padding:"5px 11px",borderRadius:8,fontSize:12,fontWeight:800,cursor:"pointer",fontFamily:"inherit",border:"1.5px solid "+(p&&p.statut===st?SC[st]:C.border),background:p&&p.statut===st?SC[st]:"#fff",color:p&&p.statut===st?"#fff":C.mid}}>{st}</button>)}
        </div>
        {p&&p.statut&&p.statut!=="Présent"&&<input value={p.note||""} onChange={e=>setNote(ed.name,e.target.value)} placeholder="Précision (motif, heure…)" style={{...S.inp,marginTop:8,padding:"8px 12px",fontSize:13}}/>}
      </div>);})}
    </div>
    <div style={S.card}>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,flexWrap:"wrap"}}>
        <div style={{fontWeight:800,color:C.dark,fontSize:15}}>Registre mensuel — {mySite}</div>
        <input type="month" value={regMois} onChange={e=>setRegMois(e.target.value)} style={{...S.inp,width:"auto",padding:"7px 10px",fontSize:13}}/>
        <button onClick={exportRegCSV} style={{marginLeft:"auto",padding:"6px 12px",borderRadius:8,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontWeight:700,fontSize:11.5,cursor:"pointer",fontFamily:"inherit"}}>Exporter CSV</button>
      </div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
        {STATUTS.map(st=><span key={st} style={{fontSize:12,fontWeight:800,display:"inline-flex",alignItems:"center",gap:4}}><span style={{width:10,height:10,borderRadius:3,background:SC[st],display:"inline-block"}}/> {st}</span>)}
      </div>
      <div style={{overflowX:"auto",WebkitOverflowScrolling:"touch",maxWidth:"100%"}}>
        <table style={{borderCollapse:"collapse",fontSize:12,minWidth:regJours.length*26+220}}>
          <thead><tr>
            <th style={{position:"sticky",left:0,background:C.white,textAlign:"left",padding:"6px 8px",borderBottom:"2px solid "+C.border,color:C.mid,zIndex:1}}>Éducateur</th>
            {regJours.map(d=><th key={d} style={{padding:"4px 2px",borderBottom:"2px solid "+C.border,color:C.light,fontWeight:700,minWidth:22,textAlign:"center"}}>{parseInt(d.slice(8),10)}</th>)}
            <th style={{padding:"4px 6px",borderBottom:"2px solid "+C.border,color:C.success,fontWeight:800}}>P</th>
            <th style={{padding:"4px 6px",borderBottom:"2px solid "+C.border,color:C.danger,fontWeight:800}}>A</th>
          </tr></thead>
          <tbody>
            {educs.map(ed=>{const line=regJours.map(d=>regIdx[d+"|"+ed.name]||null);const nbP=line.filter(x=>x&&x.statut==="Présent").length;const nbA=line.filter(x=>x&&(x.statut==="Absent"||x.statut==="Maladie")).length;return(
              <tr key={ed.id}>
                <td style={{position:"sticky",left:0,background:C.white,padding:"5px 8px",fontWeight:700,color:C.dark,whiteSpace:"nowrap",borderBottom:"1px solid "+C.border,zIndex:1}}>{ed.name}</td>
                {line.map((x,i)=><td key={i} title={x?(x.statut+(x.note?" — "+x.note:"")):""} style={{textAlign:"center",borderBottom:"1px solid "+C.border,padding:"3px 1px"}}>{x?<span style={{display:"inline-block",width:16,height:16,lineHeight:"16px",borderRadius:4,background:SC[x.statut]||C.border,color:"#fff",fontSize:11.5,fontWeight:900}}>{x.statut==="Repos"?"Re":x.statut[0]}</span>:<span style={{color:C.border}}>·</span>}</td>)}
                <td style={{textAlign:"center",fontWeight:800,color:C.success,borderBottom:"1px solid "+C.border}}>{nbP}</td>
                <td style={{textAlign:"center",fontWeight:800,color:C.danger,borderBottom:"1px solid "+C.border}}>{nbA}</td>
              </tr>);})}
          </tbody>
        </table>
      </div>
      <div style={{fontSize:12,color:C.light,marginTop:8}}>Un point (·) signifie qu'aucun statut n'a été coché ce jour-là. Survolez une case pour lire le motif.</div>
    </div>
    <div style={S.card}>
      <div style={{fontWeight:800,color:C.dark,fontSize:15,marginBottom:12}}>Rapport concernant un éducateur</div>
      {rErr&&<div style={{background:"#FFEBEE",color:"#C62828",fontWeight:700,fontSize:13,padding:"10px 14px",borderRadius:10,marginBottom:10}}>{rErr}</div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
        <div><label style={S.lbl}>Éducateur</label><select value={rEduc} onChange={e=>setREduc(e.target.value)} style={S.inp}><option value="">--</option>{educs.map(ed=><option key={ed.id} value={ed.name}>{ed.name}</option>)}</select></div>
        <div><label style={S.lbl}>Type</label><select value={rType} onChange={e=>setRType(e.target.value)} style={S.inp}><option>Maladie</option><option>Problème</option><option>Incident</option><option>Comportement</option><option>Autre</option></select></div>
        <div><label style={S.lbl}>Date</label><input type="date" value={rDate} onChange={e=>setRDate(e.target.value)} style={S.inp}/></div>
      </div>
      <div style={{marginBottom:10}}><label style={S.lbl}>Rapport</label><textarea value={rTexte} onChange={e=>setRTexte(e.target.value)} rows={4} style={{...S.inp,resize:"vertical"}} placeholder="Décrivez la situation (maladie, problème, incident…)"/></div>
      <button onClick={submitR} style={S.btnP}><Plus size={16}/>Enregistrer le rapport</button>
    </div>
    {rapports.map(r=><div key={r.id} style={{...S.card,padding:"14px 18px"}}>
      <div style={{display:"flex",justifyContent:"space-between",gap:10,flexWrap:"wrap"}}>
        <div style={{flex:1,minWidth:200}}>
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:4}}><span style={{fontWeight:800,color:C.dark,fontSize:14}}>{r.educName}</span><span style={{background:"#FFF3E0",color:"#E65100",fontSize:12,fontWeight:800,padding:"3px 10px",borderRadius:8}}>{r.type}</span><span style={{fontSize:12,color:C.light,fontWeight:700}}>{r.date} · {r.site}</span></div>
          <div style={{fontSize:13,color:C.mid,whiteSpace:"pre-wrap"}}>{r.note}</div>
          <div style={{fontSize:12,color:C.light,marginTop:4}}>Rédigé par {r.auteur}</div>
        </div>
        {(isChef||r.auteur===user.name)&&<button onClick={()=>{if(confirm("Supprimer ce rapport ?"))onDelete(r.id);}} style={{...S.btnS,padding:"7px 12px",fontSize:12,alignSelf:"flex-start"}}>Suppr.</button>}
      </div>
    </div>)}
  </div>);
}
function EspaceRH({user,docs,users,onAdd,onSign,onDelete,onUpdateUsers,etabConfig}){
  const[rhTab,setRhTab]=useState("docs");const[demForm,setDemForm]=useState(false);const[demType,setDemType]=useState("conges");const[demMotif,setDemMotif]=useState("");const[demD1,setDemD1]=useState("");const[demD2,setDemD2]=useState("");const[demMontant,setDemMontant]=useState("");const[demObjet,setDemObjet]=useState("");const[demQui,setDemQui]=useState("");
  const[cat,setCat]=useState("Document");const[dest,setDest]=useState(user.role==="educateur"?user.name:"Tous");const[signId,setSignId]=useState(null);const[uploading,setUploading]=useState(false);
  const isEnc=user.role!=="educateur";
  const educs=users.filter(u=>u.role==="educateur");
  const visible=docs.filter(d=>isEnc||d.destinataire==="Tous"||d.destinataire===user.name||d.deposeParId===user.id).sort((a,b)=>(b.date||"").localeCompare(a.date||""));
  const upload=(e)=>{const f=e.target.files[0];if(!f)return;if(f.size>4*1024*1024){alert("Fichier trop volumineux (max 4 Mo).");e.target.value="";return;}setUploading(true);const reader=new FileReader();reader.onload=()=>{onAdd({id:Date.now(),name:f.name,type:f.type,size:f.size,dataUrl:reader.result,categorie:cat,destinataire:dest,deposePar:user.name,deposeParId:user.id,date:new Date().toISOString(),signatures:[]});setUploading(false);};reader.onerror=()=>{alert("Erreur de lecture du fichier.");setUploading(false);};reader.readAsDataURL(f);e.target.value="";};
  const dlDoc=(d)=>{const a=document.createElement("a");a.href=d.dataUrl;a.download=d.name;document.body.appendChild(a);a.click();document.body.removeChild(a);};
  const fmtDate=(s)=>{try{return new Date(s).toLocaleString("fr-FR",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"});}catch{return s;}};
  const roleLbl=(r)=>r==="directeur"?"Directeur":r==="chef_service"?"Chef de service":r==="coordinateur_site"?"Coordinateur":"Éducateur";
  const enAtt=(users||[]).reduce((n,u)=>n+((u.demandes)||[]).filter(d=>d.statut==="en_attente").length,0);
  return(<div style={{padding:"18px 14px",maxWidth:720,margin:"0 auto"}}>
    <div style={{display:"flex",gap:7,marginBottom:14,flexWrap:"wrap"}}>{[{k:"docs",l:"Documents"},{k:"demandes",l:"Demandes"}].map(t=><button key={t.k} onClick={()=>setRhTab(t.k)} style={{padding:"7px 15px",borderRadius:20,border:"1.5px solid "+(rhTab===t.k?C.gold:C.border),background:rhTab===t.k?C.gold:C.white,color:rhTab===t.k?C.white:C.mid,fontWeight:700,fontSize:12.5,cursor:"pointer",fontFamily:"inherit"}}>{t.l}{t.k==="demandes"&&isEnc&&enAtt>0?" ("+enAtt+")":""}</button>)}</div>
      {rhTab==="demandes"&&(()=>{
        const estDir=user.role==="directeur"||user.role==="chef_service";
        const moi=(users||[]).find(u=>u.id===user.id)||user;
        const toutes=estDir?(users||[]).flatMap(u=>((u.demandes)||[]).map(d=>({...d,_uid:u.id,_nom:u.name}))):((moi.demandes)||[]).map(d=>({...d,_uid:moi.id,_nom:moi.name}));
        const liste=toutes.slice().sort((a,b)=>String(b.creeLe||"").localeCompare(String(a.creeLe||"")));
        const patch=(uid,id,p)=>onUpdateUsers(prev=>(prev||[]).map(u=>u.id===uid?{...u,demandes:((u.demandes)||[]).map(d=>d.id===id?{...d,...p}:d)}:u));
        const creer=()=>{
          if(demType==="conges"&&(!demD1||!demD2)){alert("Indiquez les dates de début et de fin.");return;}
          if(demType==="acompte"&&!String(demMontant).trim()){alert("Indiquez le montant souhaité.");return;}
          if(demType==="rdv"&&!demObjet.trim()){alert("Indiquez l'objet du rendez-vous.");return;}
          const d={id:Date.now(),type:demType,statut:"en_attente",creeLe:isoToday(),motif:demMotif.trim(),dateDebut:demD1,dateFin:demD2,montant:demType==="acompte"?String(demMontant).trim():"",objet:demObjet.trim(),interlocuteur:demQui.trim()};
          onUpdateUsers(prev=>(prev||[]).map(u=>u.id===user.id?{...u,demandes:[...((u.demandes)||[]),d]}:u));
          setDemMotif("");setDemD1("");setDemD2("");setDemMontant("");setDemObjet("");setDemQui("");setDemForm(false);
        };
        return(<div>
          {!estDir&&<div style={{marginBottom:12}}>
            {!demForm&&<button onClick={()=>setDemForm(true)} style={{...S.btnP,width:"100%",justifyContent:"center"}}><Plus size={14}/>Nouvelle demande</button>}
            {demForm&&<div style={{...S.card,background:C.sableLight}}>
              <label style={{...S.lbl}}>Type de demande</label>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:10}}>{DEM_TYPES.map(t=><button key={t.k} onClick={()=>setDemType(t.k)} style={{padding:"6px 13px",borderRadius:16,border:"1.5px solid "+(demType===t.k?C.gold:C.border),background:demType===t.k?C.gold:C.white,color:demType===t.k?C.white:C.mid,fontSize:11.5,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{t.l}</button>)}</div>
              {demType==="conges"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                <div><label style={{...S.lbl}}>Du</label><input type="date" style={{...S.inp}} value={demD1} onChange={e=>setDemD1(e.target.value)}/></div>
                <div><label style={{...S.lbl}}>Au</label><input type="date" style={{...S.inp}} value={demD2} onChange={e=>setDemD2(e.target.value)}/></div>
              </div>}
              {demType==="conges"&&demD1&&demD2&&<div style={{fontSize:12,color:C.mid,marginBottom:8}}>{demJours(demD1,demD2)} jour(s) calendaire(s).</div>}
              {demType==="acompte"&&<div style={{marginBottom:8}}><label style={{...S.lbl}}>Montant souhaité (€)</label><input type="number" min="0" style={{...S.inp}} value={demMontant} onChange={e=>setDemMontant(e.target.value)}/></div>}
              {demType==="rdv"&&<><div style={{marginBottom:8}}><label style={{...S.lbl}}>Objet</label><input style={{...S.inp}} value={demObjet} onChange={e=>setDemObjet(e.target.value)}/></div>
              <div style={{marginBottom:8}}><label style={{...S.lbl}}>Avec qui</label><input style={{...S.inp}} value={demQui} onChange={e=>setDemQui(e.target.value)} placeholder="Chef de service, directeur…"/></div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                <div><label style={{...S.lbl}}>Créneau souhaité</label><input type="date" style={{...S.inp}} value={demD1} onChange={e=>setDemD1(e.target.value)}/></div>
                <div><label style={{...S.lbl}}>Autre créneau</label><input type="date" style={{...S.inp}} value={demD2} onChange={e=>setDemD2(e.target.value)}/></div>
              </div></>}
              <label style={{...S.lbl}}>Motif / précisions</label>
              <textarea style={{...S.inp,minHeight:60,resize:"vertical"}} value={demMotif} onChange={e=>setDemMotif(e.target.value)}/>
              <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
                <button onClick={creer} style={{...S.btnP,flex:1,justifyContent:"center"}}><Check size={14}/>Envoyer</button>
                <button onClick={()=>setDemForm(false)} style={{...S.btnO,flex:1,justifyContent:"center"}}>Annuler</button>
              </div>
            </div>}
          </div>}
          {estDir&&<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,marginBottom:10,flexWrap:"wrap"}}>
            <div style={{fontSize:12,color:C.light}}>Demandes de l'ensemble du personnel. {liste.filter(d=>d.statut==="en_attente").length} en attente.</div>
            <button onClick={async()=>{try{await demandesPDF(liste,etabConfig);}catch(err){alert("PDF impossible : "+(err&&err.message?err.message:err));}}} style={{...S.btnO,fontSize:12,padding:"5px 12px"}}><Download size={13}/>Exporter en PDF</button>
          </div>}
          {liste.length===0&&<div style={{...S.card,fontSize:12,color:C.light}}>Aucune demande.</div>}
          {liste.map(d=>{const st=DEM_ST[d.statut]||DEM_ST.en_attente;return(<div key={d._uid+"-"+d.id} style={{...S.card,marginBottom:8,borderLeft:"4px solid "+st.c}}>
            <div style={{display:"flex",justifyContent:"space-between",gap:8,alignItems:"flex-start",flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:150}}>
                <div style={{fontSize:13.5,fontWeight:800,color:C.dark}}>{demTitre(d)}</div>
                <div style={{fontSize:12,color:C.light}}>{estDir?d._nom+" · ":""}{(DEM_TYPES.find(t=>t.k===d.type)||{}).l} · déposée le {fmt(d.creeLe)}</div>
              </div>
              <span style={{fontSize:11.5,fontWeight:800,padding:"3px 9px",borderRadius:6,background:st.bg,color:st.c,whiteSpace:"nowrap"}}>{st.l}</span>
            </div>
            {d.interlocuteur&&<div style={{fontSize:11.5,color:C.mid,marginTop:4}}>Avec : {d.interlocuteur}</div>}
            {d.motif&&<div style={{fontSize:12,color:C.dark,marginTop:6,whiteSpace:"pre-wrap"}}>{d.motif}</div>}
            {d.decisionPar&&<div style={{fontSize:12,color:st.c,marginTop:6,fontWeight:700}}>{st.l} par {d.decisionPar} le {fmt(d.decisionLe)}{d.decisionNote?" — "+d.decisionNote:""}</div>}
            {estDir&&d.statut==="en_attente"&&<div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
              <button onClick={()=>{const n=prompt("Commentaire (facultatif)")||"";patch(d._uid,d.id,{statut:"acceptee",decisionPar:user.name,decisionLe:isoToday(),decisionNote:n});}} style={{...S.btnP,flex:1,justifyContent:"center",minWidth:110}}><Check size={13}/>Accepter</button>
              <button onClick={()=>{const n=prompt("Motif du refus")||"";patch(d._uid,d.id,{statut:"refusee",decisionPar:user.name,decisionLe:isoToday(),decisionNote:n});}} style={{...S.btnO,flex:1,justifyContent:"center",minWidth:110}}>Refuser</button>
            </div>}
            {!estDir&&d.statut==="en_attente"&&<button onClick={()=>{if(confirm("Annuler cette demande ?"))patch(d._uid,d.id,{statut:"annulee"});}} style={{...S.btnO,marginTop:10,fontSize:12,padding:"5px 12px"}}>Annuler ma demande</button>}
          </div>);})}
        </div>);})()}
    {rhTab==="docs"&&<div>
    <div style={{...S.card,marginBottom:14}}>
      <div style={{fontWeight:800,fontSize:15,color:C.dark,marginBottom:4}}>Déposer un document</div>
      <div style={{fontSize:12,color:C.light,marginBottom:12}}>Contrat de travail, attestation, note de service… (PDF ou image, max 4 Mo).</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        <div><label style={{...S.lbl}}>Type de document</label><select style={{...S.inp}} value={cat} onChange={e=>setCat(e.target.value)}><option>Document</option><option>Contrat de travail</option><option>Avenant</option><option>Attestation</option><option>Note de service</option><option>Fiche de paie</option></select></div>
        <div><label style={{...S.lbl}}>Destinataire</label><select style={{...S.inp}} value={dest} onChange={e=>setDest(e.target.value)} disabled={!isEnc}><option value="Tous">Tous</option>{educs.map(u=><option key={u.id} value={u.name}>{u.name}</option>)}{!isEnc&&<option value={user.name}>{user.name}</option>}</select></div>
      </div>
      <label style={{...S.btnP,width:"100%",justifyContent:"center",cursor:"pointer",opacity:uploading?0.6:1}}>{uploading?"Chargement…":<><FileText size={15}/>Choisir un fichier à déposer</>}<input type="file" accept="application/pdf,image/*" style={{display:"none"}} onChange={upload} disabled={uploading}/></label>
    </div>
    <h3 style={{fontSize:14,fontWeight:800,color:C.dark,margin:"4px 0 10px"}}>Documents ({visible.length})</h3>
    {visible.length===0&&<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun document</div>}
    {visible.map(d=>{const signed=(d.signatures||[]).length>0;return(<div key={d.id} style={{...S.card,marginBottom:10,borderLeft:"4px solid "+(signed?"#2E7D32":C.gold)}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,flexWrap:"wrap"}}>
        <div style={{flex:1,minWidth:180}}>
          <div style={{fontWeight:800,fontSize:13,color:C.dark}}>{d.name}</div>
          <div style={{fontSize:12,color:C.light,marginTop:2}}>{d.categorie} · pour {d.destinataire} · déposé par {d.deposePar} · {fmtDate(d.date)}</div>
        </div>
        <span style={{fontSize:11.5,fontWeight:800,padding:"3px 9px",borderRadius:8,background:signed?"#E8F5E9":"#FFF3E0",color:signed?"#2E7D32":"#E65100"}}>{signed?"✓ Signé":"À signer"}</span>
      </div>
      {signed&&<div style={{marginTop:8,paddingTop:8,borderTop:"1px solid "+C.border}}>{(d.signatures||[]).map((s,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:10,marginBottom:6,flexWrap:"wrap"}}><img src={s.image} alt="signature" style={{height:42,border:"1px solid "+C.border,borderRadius:6,background:C.white}}/><div style={{fontSize:11.5,color:C.mid}}><b>{s.nom}</b> ({roleLbl(s.role)})<br/>signé le {fmtDate(s.date)}</div></div>)}</div>}
      <div style={{display:"flex",gap:6,marginTop:10,flexWrap:"wrap"}}>
        <button onClick={()=>dlDoc(d)} style={{...S.btnS}}><Download size={14}/>Télécharger</button>
        <button onClick={()=>setSignId(d.id)} style={{...S.btnP}}>✍️ Signer</button>
        {(isEnc||d.deposeParId===user.id)&&<button onClick={()=>{if(confirm("Supprimer ce document ?"))onDelete(d.id);}} style={{padding:"8px 14px",borderRadius:10,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>Supprimer</button>}
      </div>
    </div>);})}
    {signId!=null&&<SignaturePad onCancel={()=>setSignId(null)} onValidate={(img)=>{onSign(signId,{nom:user.name,role:user.role,date:new Date().toISOString(),image:img});setSignId(null);}}/>}
    <div style={{fontSize:11.5,color:C.light,marginTop:14,padding:"10px 12px",background:C.sableLight,borderRadius:8}}>ℹ️ La signature recueillie ici est une <b>signature simple</b> (tracé manuscrit numérisé + nom, rôle et horodatage). Elle a valeur de preuve d'accord interne. Pour un contrat de travail à pleine valeur juridique opposable, une signature électronique qualifiée via un prestataire certifié reste recommandée.</div>
    </div>}
  </div>);
}


const NOTIF_META={eig:{ic:"⚠️",c:"#C62828",l:"EIG"},evenement:{ic:"❗",c:"#E65100",l:"Événement"},intendance_urgente:{ic:"🔧",c:"#C62828",l:"Intendance urgente"},intendance:{ic:"📦",c:"#1565C0",l:"Intendance"},rapport_site:{ic:"📋",c:"#2E7D32",l:"Rapport de site"},transmission:{ic:"📝",c:"#6A1B9A",l:"Transmission"}};
function NotifPanel({notifs,user,onClose,onReadAll,onOpen}){
  const list=[...(notifs||[])].sort((a,b)=>(b.date||"").localeCompare(a.date||""));
  const fmt=(iso)=>{if(!iso)return"";const d=new Date(iso);const p=n=>String(n).padStart(2,"0");return p(d.getDate())+"/"+p(d.getMonth()+1)+" "+p(d.getHours())+":"+p(d.getMinutes());};
  return(<>
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.35)",zIndex:70}}/>
    <div style={{position:"fixed",top:0,right:0,width:"min(360px,92vw)",height:"100vh",background:C.white,zIndex:71,boxShadow:"-8px 0 32px rgba(0,0,0,0.25)",display:"flex",flexDirection:"column"}}>
      <div style={{padding:"16px 16px 12px",borderBottom:"1px solid "+C.border,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
        <Bell size={18} color={C.gold}/><span style={{fontWeight:800,fontSize:15,color:C.dark,flex:1}}>Notifications</span>
        <button onClick={onReadAll} style={{background:"none",border:"1px solid "+C.border,borderRadius:8,padding:"4px 10px",fontSize:12,fontWeight:700,color:C.mid,cursor:"pointer",fontFamily:"inherit"}}>Tout marquer lu</button>
        <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:C.mid,padding:4,display:"flex",flexWrap:"wrap"}}><X size={17}/></button>
      </div>
      <div style={{flex:1,overflowY:"auto",padding:"10px 12px"}}>
        {list.length===0&&<div style={{textAlign:"center",color:C.light,fontSize:13,padding:30}}>Aucune notification</div>}
        {list.map(n=>{const m=NOTIF_META[n.type]||{ic:"🔔",c:C.mid,l:n.type};const lu=(n.readBy||[]).includes(user.id);return(
          <div key={n.id} onClick={()=>onOpen&&onOpen(n)} style={{borderLeft:"4px solid "+m.c,background:lu?C.sableLight:C.goldLight,borderRadius:8,padding:"10px 12px",marginBottom:8,cursor:"pointer"}}>
            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:3,flexWrap:"wrap"}}>
              <span style={{fontSize:14}}>{m.ic}</span>
              <span style={{fontSize:12,fontWeight:800,color:m.c}}>{n.titre||m.l}</span>
              {!lu&&<span style={{width:8,height:8,borderRadius:4,background:"#C62828",marginLeft:"auto"}}/>}
            </div>
            {n.message&&<div style={{fontSize:12.5,color:C.dark,marginBottom:4,whiteSpace:"pre-wrap"}}>{n.message}</div>}
            <div style={{fontSize:12,fontWeight:800,color:m.c,marginBottom:2}}>Voir le détail →</div>
            <div style={{fontSize:12,color:C.light,fontWeight:600}}>{fmt(n.date)}{n.site?" · "+n.site:""}{n.author?" · "+n.author:""}</div>
          </div>);})}
      </div>
    </div>
  </>);
}

function Transmissions({user,items,onAdd,onMarkLu}){
  const SITES=["Djilass","Fatick"];
  const defSite=user.site&&user.site!=="Tous"?user.site:SITES[0];
  const[site,setSite]=useState(defSite);
  const[txt,setTxt]=useState("");
  const[filter,setFilter]=useState(user.site&&user.site!=="Tous"?user.site:"Tous");
  const fmt=(iso)=>{if(!iso)return"";const d=new Date(iso);const p=n=>String(n).padStart(2,"0");return p(d.getDate())+"/"+p(d.getMonth()+1)+"/"+d.getFullYear()+" "+p(d.getHours())+":"+p(d.getMinutes());};
  const list=(items||[]).filter(t=>filter==="Tous"||t.site===filter).sort((a,b)=>(b.date||"").localeCompare(a.date||""));
  const[reunion,setReunion]=useState(false);
  const submit=()=>{if(!txt.trim())return;onAdd({site,texte:txt.trim(),type:reunion?"reunion_hebdo":"transmission"});setTxt("");setReunion(false);};
  return(<div>
    <div style={{...S.card,marginBottom:14}}>
      <div style={{fontWeight:800,fontSize:15,color:C.dark,marginBottom:10}}>Nouvelle consigne</div>
      <div style={{display:"flex",gap:8,marginBottom:8,flexWrap:"wrap"}}>
        {SITES.map(x=><button key={x} onClick={()=>setSite(x)} style={{flex:1,padding:"8px 0",borderRadius:8,border:"1.5px solid "+(site===x?C.gold:C.border),background:site===x?C.goldA22:"#fff",fontWeight:700,fontSize:13,color:site===x?C.goldDark:C.mid,cursor:"pointer",fontFamily:"inherit"}}>{x}</button>)}
      </div>
      <label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,fontWeight:700,color:C.dark,cursor:"pointer",marginBottom:8,padding:"8px 10px",borderRadius:8,background:reunion?C.goldLight:"transparent",border:"1.5px solid "+(reunion?C.gold:C.border),flexWrap:"wrap"}}><input type="checkbox" checked={reunion} onChange={e=>setReunion(e.target.checked)} style={{accentColor:C.gold,width:17,height:17}}/>Réunion hebdomadaire d'équipe</label>
      <textarea value={txt} onChange={e=>setTxt(e.target.value)} rows={3} placeholder="Consigne courte pour la relève (traitement donné, incident, visite prévue…)" style={{width:"100%",boxSizing:"border-box",border:"1.5px solid "+C.border,borderRadius:10,padding:10,fontSize:13.5,fontFamily:"inherit",resize:"vertical",marginBottom:8}}/>
      <button onClick={submit} disabled={!txt.trim()} style={{width:"100%",padding:"10px 0",borderRadius:10,border:"none",background:txt.trim()?C.gold:C.border,color:"#fff",fontWeight:800,fontSize:13.5,cursor:txt.trim()?"pointer":"default",fontFamily:"inherit"}}>Consigner ({site})</button>
    </div>
    <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
      {["Tous",...SITES].map(x=><button key={x} onClick={()=>setFilter(x)} style={{flex:1,padding:"7px 0",borderRadius:8,border:"1.5px solid "+(filter===x?C.primary:C.border),background:filter===x?C.primaryA18:"#fff",fontWeight:700,fontSize:12.5,color:filter===x?C.primary:C.mid,cursor:"pointer",fontFamily:"inherit"}}>{x}</button>)}
    </div>
    {list.length===0&&<div style={{...S.card,textAlign:"center",color:C.light,fontSize:13}}>Aucune transmission</div>}
    {list.map(t=>{const luParMoi=(t.luPar||[]).some(l=>l&&l.name===user.name);return(
      <div key={t.id} style={{...S.card,marginBottom:10,borderLeft:"4px solid "+(t.site==="Djilass"?C.gold:"#6A1B9A")}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6,flexWrap:"wrap"}}>
          <span style={{fontSize:12,fontWeight:900,color:"#fff",background:t.site==="Djilass"?C.gold:"#6A1B9A",borderRadius:5,padding:"2px 8px"}}>{t.site}</span>
          <span style={{fontSize:12,fontWeight:700,color:C.dark}}>{t.author}</span>
          <span style={{fontSize:12,color:C.light,fontWeight:600,marginLeft:"auto"}}>{fmt(t.date)}</span>
        </div>
        <div style={{fontSize:13.5,color:C.dark,whiteSpace:"pre-wrap",marginBottom:8}}>{t.type==="reunion_hebdo"&&<span style={{display:"inline-block",fontSize:11.5,fontWeight:800,color:"#5B3E90",background:"#EDE7F6",borderRadius:6,padding:"2px 8px",marginRight:6,verticalAlign:"middle"}}>RÉUNION HEBDO</span>}{t.texte}</div>
        <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
          {(t.luPar||[]).map((l,i)=><span key={i} title={fmt(l.date)} style={{fontSize:12,fontWeight:700,color:"#2E7D32",background:"rgba(102,187,106,0.15)",borderRadius:5,padding:"2px 7px"}}>✓ {l.name}</span>)}
          {!luParMoi&&<button onClick={()=>onMarkLu(t.id)} style={{marginLeft:"auto",padding:"4px 12px",borderRadius:8,border:"1.5px solid #2E7D32",background:C.white,color:"#2E7D32",fontWeight:800,fontSize:11.5,cursor:"pointer",fontFamily:"inherit"}}>Marquer lu</button>}
        </div>
      </div>);})}
  </div>);
}

function RechercheGlobale({user,rapports,evenements,transmissions,rapportsSite,intendance,jeunes,majeurs}){
  const[q,setQ]=useState("");
  const allP=[...(jeunes||[]),...(majeurs||[])];
  const nomJ=(jid)=>{const j=allP.find(x=>String(x.id)===String(jid));return j?(j.prenom+" "+(j.nom||"")):"";};
  const canSee=(jid)=>{const j=allP.find(x=>String(x.id)===String(jid));if(!j)return user.role!=="educateur";if(user.role==="educateur"){if(user.isEducMajeur&&j.id<100)return false;if(!user.isEducMajeur&&j.id>=100)return false;}if((user.role==="educateur"||user.role==="coordinateur_site")&&user.site&&user.site!=="Tous"&&j.site!==user.site)return false;return true;};
  const needle=q.trim().toLowerCase();
  const hit=(...fields)=>fields.some(f=>f&&String(f).toLowerCase().includes(needle));
  const res=useMemo(()=>{
    if(needle.length<2)return null;
    const r={};
    r.rapports=(rapports||[]).filter(x=>canSee(x.jeuneId)&&hit(x.observation,x.author,nomJ(x.jeuneId))).slice(-60).reverse();
    r.evenements=(evenements||[]).filter(x=>canSee(x.jeuneId)&&hit(x.titre,x.description,x.author,nomJ(x.jeuneId))).slice(-60).reverse();
    r.transmissions=(transmissions||[]).filter(x=>hit(x.texte,x.author,x.site)).slice(-60).reverse();
    r.rapportsSite=(rapportsSite||[]).filter(x=>hit(x.site,x.observation,x.contenu,x.texte,x.observations,x.author,x.auteur)).slice(-30).reverse();
    r.intendance=(intendance||[]).filter(x=>hit(x.objet,x.description,x.auteur,x.site)).slice(-30).reverse();
    return r;
  },[needle,rapports,evenements,transmissions,rapportsSite,intendance]);
  const Section=({titre,items,render})=>!items||items.length===0?null:(
    <div style={{marginBottom:16}}>
      <div style={{fontWeight:800,fontSize:13,color:C.goldDark,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:8}}>{titre} ({items.length})</div>
      {items.map(render)}
    </div>);
  const line=(key,top,body,meta)=>(
    <div key={key} style={{...S.card,padding:"10px 14px",marginBottom:8}}>
      <div style={{fontSize:12.5,fontWeight:800,color:C.dark,marginBottom:2}}>{top}</div>
      {body&&<div style={{fontSize:12.5,color:C.mid,whiteSpace:"pre-wrap",maxHeight:80,overflow:"hidden"}}>{body}</div>}
      <div style={{fontSize:12,color:C.light,fontWeight:600,marginTop:4}}>{meta}</div>
    </div>);
  return(<div>
    <div style={{...S.card,marginBottom:14,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
      <Search size={18} color={C.gold}/>
      <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Rechercher un jeune, un mot, un auteur…" autoFocus style={{flex:1,border:"none",outline:"none",fontSize:14.5,fontFamily:"inherit",background:"transparent"}}/>
      {q&&<button onClick={()=>setQ("")} style={{background:"none",border:"none",cursor:"pointer",color:C.mid,display:"flex",padding:2,flexWrap:"wrap"}}><X size={15}/></button>}
    </div>
    {!res&&<div style={{textAlign:"center",color:C.light,fontSize:13,padding:20}}>Saisissez au moins 2 caractères — la recherche couvre rapports journaliers, événements, transmissions, rapports de site et intendance.</div>}
    {res&&<>
      <Section titre="Rapports journaliers" items={res.rapports} render={x=>line("r"+x.id,nomJ(x.jeuneId)||("ID "+x.jeuneId),x.observation,(x.date||"")+" · "+(x.author||""))}/>
      <Section titre="Événements indésirables" items={res.evenements} render={x=>line("e"+x.id,(x.titre||"Événement")+(nomJ(x.jeuneId)?" — "+nomJ(x.jeuneId):""),x.description,(x.date||"")+" · "+(x.author||"")+(x.gravite?" · "+x.gravite:""))}/>
      <Section titre="Transmissions" items={res.transmissions} render={x=>line("t"+x.id,(x.site||"")+" — "+(x.author||""),x.texte,(x.date||"").slice(0,10))}/>
      <Section titre="Rapports de site" items={res.rapportsSite} render={x=>line("s"+x.id,(x.site||"Site")+(x.semaine?" · "+x.semaine:""),x.observation||x.contenu||x.texte||x.observations||"",(x.date||x.createdAt||"").slice(0,10)+" · "+(x.author||x.auteur||""))}/>
      <Section titre="Intendance" items={res.intendance} render={x=>line("i"+x.id,(x.objet||"Demande")+(x.site?" · "+x.site:""),x.description,(x.createdAt||"").slice(0,10)+" · "+(x.auteur||"")+(x.urgence?" · "+x.urgence:""))}/>
      {res.rapports.length===0&&res.evenements.length===0&&res.transmissions.length===0&&res.rapportsSite.length===0&&res.intendance.length===0&&<div style={{textAlign:"center",color:C.light,fontSize:13,padding:20}}>Aucun résultat pour « {q} »</div>}
    </>}
  </div>);
}

export default function App(){
  const[user,setUser]=useState(null),[ page,setPage]=useState("dashboard"),[ open,setOpen]=useState(false),[ sel,setSel]=useState(null);
  const[viewAs,setViewAs]=useState(null);
  const[syncMsg,setSyncMsg]=useState(null);const[cloudReady,setCloudReady]=useState(false);const[updateReady,setUpdateReady]=useState(false);
  const lsData=loadLS();
  const[rapports,setRapports0]=useState(Array.isArray(lsData?.rapports)?lsData.rapports:INIT_RAPPORTS),[ presences,setPresences0]=useState(Array.isArray(lsData?.presences)?lsData.presences:INIT_PRESENCES),[ evenements,setEvenements0]=useState(Array.isArray(lsData?.evenements)?lsData.evenements:INIT_EV);
  const[appUsers,setAppUsers0]=useState(()=>Array.isArray(lsData&&lsData.users)&&lsData.users.length?lsData.users:USERS);
  const setAppUsers=(v)=>setAppUsers0(prev=>{const nx=typeof v==="function"?v(prev):v;const ts=new Date().toISOString();const pm={};(prev||[]).forEach(x=>{if(x&&x.id!=null)pm[x.id]=x;});return(nx||[]).map(x=>{if(!x||x.id==null)return x;const old=pm[x.id];return(old&&JSON.stringify(old)===JSON.stringify(x))?x:{...x,updatedAt:ts};});});
  const[appJeunes,setAppJeunes]=useState(()=>Array.isArray(lsData&&lsData.jeunes)&&lsData.jeunes.length?lsData.jeunes:JEUNES);
 const[rapportsSite,setRapportsSite0]=useState(Array.isArray(lsData?.rapportsSite)?lsData.rapportsSite:[]);const[intendance,setIntendance0]=useState(Array.isArray(lsData?.intendance)?lsData.intendance:[]);const[suiviEduc,setSuiviEduc0]=useState(Array.isArray(lsData?.suiviEduc)?lsData.suiviEduc:[]);const[purgeMarks,setPurgeMarks]=useState(lsData?.purgeMarks||{});const[notifs,setNotifs0]=useState(Array.isArray(lsData?.notifs)?lsData.notifs:[]);const[transmissions,setTransmissions0]=useState(Array.isArray(lsData?.transmissions)?lsData.transmissions:[]);const[notifOpen,setNotifOpen]=useState(false);const[theme,setTheme]=useState(()=>{try{return localStorage.getItem("pdsr_theme")||"light";}catch(e){return"light";}});useEffect(()=>{document.documentElement.dataset.theme=theme;try{localStorage.setItem("pdsr_theme",theme);}catch(e){}},[theme]);const[ptr,setPtr]=useState(0);const mainRef=useRef(null);const ptrStart=useRef(null);
  const[agenda,setAgenda0]=useState(lsData?.agenda||[]);
  const _stampDiff=(prev,arr,keyf)=>{const kf=keyf||(x=>x&&x.id);const pm={};(Array.isArray(prev)?prev:[]).forEach(x=>{const k=x?kf(x):null;if(k!=null){const{updatedAt,...rest}=x;pm[k]=JSON.stringify(rest);}});return(Array.isArray(arr)?arr:[]).map(x=>{const k=x?kf(x):null;if(k==null)return x;const{updatedAt,...rest}=x;return pm[k]===JSON.stringify(rest)?x:{...x,updatedAt:nowSrv()};});};
  const _mkStamped=(setter,keyf)=>(next)=>setter(prev=>_stampDiff(prev,typeof next==="function"?next(prev):next,keyf));
  const setRapports=_mkStamped(setRapports0),setEvenements=_mkStamped(setEvenements0),setAgenda=_mkStamped(setAgenda0),setIntendance=_mkStamped(setIntendance0),setSuiviEduc=_mkStamped(setSuiviEduc0),setRapportsSite=_mkStamped(setRapportsSite0),setNotifs=_mkStamped(setNotifs0),setTransmissions=_mkStamped(setTransmissions0),setPresences=_mkStamped(setPresences0,x=>x&&x.jeuneId!=null&&x.date?x.jeuneId+"|"+x.date:null);
 const[loginLogs,setLoginLogs]=useState(lsData?.loginLogs||[]);
const[deletionLogs,setDeletionLogs]=useState(lsData?.deletionLogs||[]);
const[appMajeurs,setAppMajeurs]=useState(lsData?.majeurs||MAJEURS);
const[appDjiPlan,setAppDjiPlan]=useState(lsData?.djiPlan||DJI_PLAN);
const[appFatPlan,setAppFatPlan]=useState(lsData?.fatPlan||FAT_PLAN);
const[projets,setProjets]=useState(Array.isArray(lsData?.projets)?lsData.projets:[]);
const[sejourConfig,setSejourConfig]=useState(lsData?.sejourConfig||{Djilass:{dateDebut:"2026-03-30"},Fatick:{dateDebut:"2026-03-16"}});
const[etabConfig,setEtabConfig]=useState({...ETAB_DEFAULT,...(lsData?.etabConfig||{})});
const[docs,setDocs]=useState(()=>{try{return JSON.parse(localStorage.getItem("pdsr_docs"))||[];}catch{return[];}});
const docsReadOK=useRef(false);const docsWasNonEmpty=useRef(false);
useEffect(()=>{let cancelled=false,tries=0;const att=async()=>{tries++;let ok=false,val=null;try{const r=await fetch(FB_URL+"/documents.json?auth="+FB_SECRET);if(r.ok){val=await r.json();ok=true;}}catch(e){}if(cancelled)return;if(ok){const remote=Array.isArray(val)?val:val?Object.values(val):[];if(remote.length){docsWasNonEmpty.current=true;setDocs(prev=>{const byId={};[...prev,...remote].forEach(d=>{if(d&&d.id!=null)byId[d.id]=d;});return Object.values(byId);});}docsReadOK.current=true;}else if(tries<6){setTimeout(att,3000);}};att();return()=>{cancelled=true;};},[]);
useEffect(()=>{try{localStorage.setItem("pdsr_docs",JSON.stringify(docs));}catch{}if(!docsReadOK.current)return;if(docs.length===0&&docsWasNonEmpty.current)return;if(docs.length)docsWasNonEmpty.current=true;if(window._docsTimer)clearTimeout(window._docsTimer);window._docsTimer=setTimeout(()=>{fbSet("documents",docs);},1500);},[docs]);
  // Temps réel : récupère et fusionne les nouveautés toutes les 5 s (affichage live sans recharger)
  useEffect(()=>{if(!user)return;const sameIds=(a,b)=>{if(a.length!==b.length)return true;const byId={};a.forEach(x=>{if(x&&x.id!=null)byId[x.id]=x;});return!b.every(x=>x&&byId[x.id]===x);};const tick=async()=>{if(!fbReadOK.current)return;if(window._tickBusy||document.visibilityState==="hidden")return;window._tickBusy=true;try{let rem=null;try{const r=await fetchTO(FB_URL+"/data.json?auth="+FB_SECRET,null,15000);if(r.ok)rem=await r.json();}catch(e){}if(!rem)return;const tomb=new Set();((rem.deletionLogs)||[]).forEach(t=>{if(t&&t.origId!=null)tomb.add(t.type+":"+t.origId);});const merge=(prev,remoteArr,typeName)=>{const m={};(Array.isArray(prev)?prev:[]).forEach(x=>{if(x&&x.id!=null)m[x.id]=x;});const ra=Array.isArray(remoteArr)?remoteArr:(remoteArr&&typeof remoteArr==="object"?Object.values(remoteArr):[]);ra.forEach(x=>{if(!x||x.id==null)return;const cur=m[x.id];if(!cur){m[x.id]=x;return;}if((x.updatedAt||"")>(cur.updatedAt||""))m[x.id]=x;});return Object.values(m).filter(x=>!tomb.has(typeName+":"+x.id)&&!inPurgeRanges(((rem.purgeMarks||{}).ranges),typeName,x));};setRapports0(prev=>{const n=merge(prev,rem.rapports,"rapport");return sameIds(prev||[],n)?n:prev;});setEvenements0(prev=>{const n=merge(prev,rem.evenements,"evenement");return sameIds(prev||[],n)?n:prev;});setAgenda0(prev=>{const n=merge(prev,rem.agenda,"agenda");return sameIds(prev||[],n)?n:prev;});setIntendance0(prev=>{const n=merge(prev,rem.intendance,"intendance");return sameIds(prev||[],n)?n:prev;});setSuiviEduc0(prev=>{const n=merge(prev,rem.suiviEduc,"suiviEduc");return sameIds(prev||[],n)?n:prev;});if(rem.purgeMarks){setPurgeMarks(prev=>({...prev,...rem.purgeMarks}));const pm=rem.purgeMarks;if(pm.loginLogs)setLoginLogs(prev=>{const n=(prev||[]).filter(x=>(x&&x.date||"")>pm.loginLogs);return n.length===(prev||[]).length?prev:n;});if(pm.deletionLogs)setDeletionLogs(prev=>{const n=(prev||[]).filter(x=>(x&&x.date||"")>pm.deletionLogs);return n.length===(prev||[]).length?prev:n;});}if(rem.presences){setPresences0(prev=>{const pm={};const add=(a)=>{(Array.isArray(a)?a:(a&&typeof a==="object"?Object.values(a):[])).forEach(p=>{if(!p||p.jeuneId==null||!p.date)return;const k=p.jeuneId+"|"+p.date;const cur=pm[k];pm[k]=(!cur||(p.updatedAt||"")>(cur.updatedAt||""))?p:cur;});};add(prev);add(rem.presences);const n=Object.values(pm);return(Array.isArray(prev)&&prev.length===n.length)?prev:n;});}if(rem.users&&user&&user.role!=="directeur"&&user.role!=="chef_service"&&!user.isAdmin){const fbu=(Array.isArray(rem.users)?rem.users:Object.values(rem.users)).filter(Boolean);setAppUsers(prev=>{const next=fbu.map(fu=>{const base=USERS.find(x=>x.id===fu.id)||fu;return{...base,...fu};});const changed=next.length!==(prev||[]).length||next.some(n=>{const p=(prev||[]).find(x=>x.id===n.id);return!p||p.role!==n.role||p.site!==n.site||p.disabled!==n.disabled;});return changed?next:prev;});}}finally{window._tickBusy=false;}};const iv=setInterval(tick,15000);return()=>clearInterval(iv);},[user]);
  useEffect(()=>{const h=()=>{if(document.visibilityState==="hidden"&&window._fbTimer){clearTimeout(window._fbTimer);window._fbTimer=null;if(window._fbFlush)window._fbFlush();}};document.addEventListener("visibilitychange",h);window.addEventListener("pagehide",h);return()=>{document.removeEventListener("visibilitychange",h);window.removeEventListener("pagehide",h);};},[]);
  useEffect(()=>{if(!user)return;if(!(user.role==="directeur"||user.role==="chef_service"||user.isAdmin))return;if(!fbReadOK.current)return;if(!Array.isArray(appUsers)||appUsers.length===0)return;if(window._usersTimer)clearTimeout(window._usersTimer);window._usersTimer=setTimeout(()=>{fbSet("data/users",appUsers);},700);},[appUsers,user]);

// Firebase sync
const fbSkip=useRef(false);
const fbLoaded=useRef(false);
const fbWasNonEmpty=useRef(false);
const fbReadOK=useRef(false);
const purgeIntent=useRef(false);
const toArr=(v)=>!v?[]:Array.isArray(v)?v.filter(Boolean):Object.values(v).filter(Boolean);
useEffect(()=>{if(fbSkip.current){fbSkip.current=false;return;}if(!user)return;if(!fbLoaded.current)return;if(!fbReadOK.current)return;const presN=presences?(Array.isArray(presences)?presences.length:Object.keys(presences).length):0;const localVide=(!rapports||rapports.length===0)&&(!evenements||evenements.length===0)&&presN===0;if(localVide&&!fbWasNonEmpty.current){return;}if(window._fbTimer)clearTimeout(window._fbTimer);const _doSave=async()=>{window._fbTimer=null;
  try{
  let remote=null;try{const r=await fetchTO(FB_URL+"/data.json?auth="+FB_SECRET,null,15000);if(r.ok)remote=await r.json();}catch(e){}
  const tomb=new Set();[...(deletionLogs||[]),...((remote&&remote.deletionLogs)||[])].forEach(t=>{if(t&&t.origId!=null)tomb.add(t.type+":"+t.origId);});
  const rpm={...((remote&&remote.purgeMarks)||{}),...purgeMarks};
  const union=(loc,rem,typeName)=>{const m={};(Array.isArray(rem)?rem:[]).forEach(x=>{if(x&&x.id!=null)m[x.id]=x;});(Array.isArray(loc)?loc:[]).forEach(x=>{if(!x||x.id==null)return;const cur=m[x.id];if(!cur){m[x.id]=x;return;}m[x.id]=((cur.updatedAt||"")>(x.updatedAt||""))?cur:x;});return Object.values(m).filter(x=>!tomb.has(typeName+":"+x.id)&&!inPurgeRanges(rpm.ranges,typeName,x));};
  const mergedR=union(rapports,remote&&remote.rapports,"rapport");
  const mergedE=union(evenements,remote&&remote.evenements,"evenement");
  const mergedA=union(agenda,remote&&remote.agenda,"agenda");
  const pm={};const addP=(arr)=>{(Array.isArray(arr)?arr:(arr&&typeof arr==="object"?Object.values(arr):[])).forEach(p=>{if(!p||p.jeuneId==null||!p.date)return;const k=p.jeuneId+"|"+p.date;const cur=pm[k];pm[k]=(!cur||(p.updatedAt||"")>=(cur.updatedAt||""))?p:cur;});};addP(remote&&remote.presences);addP(Array.isArray(presences)?presences:[]);const mergedP=Array.isArray(presences)&&!Array.isArray(remote&&remote.presences)?presences:Object.values(pm);
  const mergedDL=union(deletionLogs,remote&&remote.deletionLogs,"__none__").filter(x=>!rpm.deletionLogs||(x&&x.date||"")>rpm.deletionLogs);const llW=(()=>{const seen=new Set();const out=[];for(const x of[...(loginLogs||[]),...toArr(remote&&remote.loginLogs)]){if(!x||x.id==null||seen.has(x.id))continue;seen.add(x.id);out.push(x);}return out.filter(x=>!rpm.loginLogs||(x&&x.date||"")>rpm.loginLogs).sort((a,b)=>(b.date||"").localeCompare(a.date||"")).slice(0,800);})();
  const mergedInt=union(intendance,remote&&remote.intendance,"intendance");
  const mergedSE=union(suiviEduc,remote&&remote.suiviEduc,"suiviEduc");
  const mergedRS=union(rapportsSite,remote&&remote.rapportsSite,"rapportSite");
  const mergedNo=union(notifs,remote&&remote.notifs,"notif");
  const mergedTr=union(transmissions,remote&&remote.transmissions,"transmission");
  const remoteUsersArr=Array.isArray(remote&&remote.users)?remote.users:(remote&&remote.users)?Object.values(remote.users):null;
  const isRosterAdmin=user&&(user.role==="directeur"||user.role==="chef_service"||user.isAdmin);
  const mergedU=(()=>{const m={};(remoteUsersArr||[]).forEach(x=>{if(x&&x.id!=null)m[x.id]=x;});(Array.isArray(appUsers)?appUsers:[]).forEach(x=>{if(!x||x.id==null)return;const cur=m[x.id];if(!cur){if(x.updatedAt)m[x.id]=x;return;}if((x.updatedAt||"")>(cur.updatedAt||""))m[x.id]=x;});return Object.values(m);})();
  const usersToWrite=isRosterAdmin?mergedU:(remoteUsersArr||appUsers);
  const data={rapports:mergedR,presences:mergedP,evenements:mergedE,jeunes:appJeunes,users:usersToWrite,agenda:mergedA,loginLogs:llW,majeurs:appMajeurs,rapportsSite:mergedRS,intendance:mergedInt,suiviEduc:mergedSE,notifs:mergedNo,transmissions:mergedTr,purgeMarks:rpm,djiPlan:appDjiPlan,fatPlan:appFatPlan,deletionLogs:mergedDL,projets,sejourConfig,etabConfig};
  const dataVide=mergedR.length===0&&mergedE.length===0;
  if(dataVide&&fbWasNonEmpty.current&&!purgeIntent.current){console.warn("[PDSR] Envoi vide bloqué.");return;}
  if(mergedR.length||mergedE.length)fbWasNonEmpty.current=true;
  let _wok=false;
  for(let _a=0;_a<3&&!_wok;_a++){try{const _r=await fetch(FB_URL+"/data.json?auth="+FB_SECRET,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(data)});if(_r.ok){_wok=true;_captureClock(_r);}else{await new Promise(x=>setTimeout(x,1200));}}catch(e){await new Promise(x=>setTimeout(x,1200));}}
  if(!_wok){setSyncMsg("Écriture cloud échouée — saisies conservées sur cet appareil, nouvel essai automatique");setTimeout(()=>setSyncMsg(""),4000);window._fbTimer=setTimeout(_doSave,8000);return;}
  if(dataVide&&purgeIntent.current){purgeIntent.current=false;fbWasNonEmpty.current=false;}
  if(!dataVide)fbSet("backups/jour_"+today,{...data,_ts:new Date().toISOString()});
  if(mergedR.length>(rapports||[]).length||mergedE.length>(evenements||[]).length){fbSkip.current=true;setRapports0(mergedR);setEvenements0(mergedE);setAgenda0(mergedA);}
  }catch(_e){console.error("[PDSR] Sauvegarde cloud interrompue :",_e);try{setSyncMsg("Synchronisation interrompue \u2014 saisies conservees sur cet appareil");setTimeout(()=>setSyncMsg(""),6000);}catch(_){}}
};window._fbFlush=_doSave;window._fbTimer=setTimeout(_doSave,800);},[rapports,presences,evenements,appUsers,appJeunes,agenda,loginLogs,appMajeurs,rapportsSite,intendance,suiviEduc,purgeMarks,appDjiPlan,appFatPlan,deletionLogs,projets,sejourConfig,notifs,transmissions]);
useEffect(()=>{const t=setTimeout(()=>{saveLS({rapports,presences,evenements,users:appUsers,jeunes:appJeunes,agenda,loginLogs,majeurs:appMajeurs,rapportsSite,djiPlan:appDjiPlan,fatPlan:appFatPlan,deletionLogs,projets,sejourConfig,etabConfig,intendance,suiviEduc,notifs,transmissions,purgeMarks});},400);return()=>clearTimeout(t);},[rapports,presences,evenements,appUsers,appJeunes,agenda,loginLogs,appMajeurs,rapportsSite,appDjiPlan,appFatPlan,deletionLogs,projets,sejourConfig,etabConfig,intendance,suiviEduc,notifs,transmissions,purgeMarks]);
const loadFb=(d)=>{if(!d){return;}
  const rs=Number(d.resetStamp||0);
  if(rs&&rs>getLocalReset()){
    setLocalReset(rs);
    try{localStorage.removeItem(LS_KEY);}catch(e){}
    fbSkip.current=true;
    window.location.reload();
    return;
  }const cnt=(v)=>!v?0:(Array.isArray(v)?v.filter(Boolean).length:Object.keys(v).length);const remoteR=cnt(d.rapports),remoteE=cnt(d.evenements);const localR=(rapports||[]).length,localE=(evenements||[]).length;const purgedRecently=d.purgeMarks&&d.purgeMarks.lastPurge&&(Date.now()-new Date(d.purgeMarks.lastPurge).getTime())<86400000;if(remoteR===0&&remoteE===0&&(localR>0||localE>0)&&!purgedRecently){console.warn("[PDSR] Cloud vide ignoré : on conserve les données locales et on répare le cloud.");fbWasNonEmpty.current=true;fbReadOK.current=true;fbLoaded.current=true;const heal={rapports,presences,evenements,jeunes:appJeunes,users:appUsers,agenda,loginLogs,majeurs:appMajeurs,rapportsSite,intendance,suiviEduc,notifs,transmissions,djiPlan:appDjiPlan,fatPlan:appFatPlan,deletionLogs,projets,sejourConfig,etabConfig};fbSet("data",heal);fbSet("backups/jour_"+new Date().toISOString().slice(0,10),{...heal,_ts:new Date().toISOString()});return;}fbSkip.current=true;const remoteNonEmpty=remoteR||remoteE;if(remoteNonEmpty)fbWasNonEmpty.current=true;const tombL=new Set();toArr(d.deletionLogs).forEach(t=>{if(t&&t.origId!=null)tombL.add(t.type+":"+t.origId);});const mrg=(remArr,prev,type)=>{const m={};(remArr||[]).forEach(x=>{if(x&&x.id!=null)m[x.id]=x;});(prev||[]).forEach(x=>{if(!x||x.id==null)return;const cur=m[x.id];if(!cur){m[x.id]=x;return;}if((x.updatedAt||"")>(cur.updatedAt||""))m[x.id]=x;});return Object.values(m).filter(x=>!tombL.has(type+":"+x.id)&&!inPurgeRanges(((d.purgeMarks||{}).ranges)||((purgeMarks||{}).ranges),type,x));};if(d.rapports&&(remoteR>0||localR===0))setRapports0(prev=>mrg(toArr(d.rapports),prev,"rapport"));if(d.presences)setPresences0(Array.isArray(d.presences)?d.presences.filter(Boolean):(typeof d.presences==="object"?Object.values(d.presences).filter(Boolean):INIT_PRESENCES));if(d.evenements&&(remoteE>0||localE===0))setEvenements0(prev=>mrg(toArr(d.evenements),prev,"evenement"));if(d.jeunes){const jArr=toArr(d.jeunes);setAppJeunes(jArr.map(fj=>{const base=JEUNES.find(x=>x.id===fj.id)||fj;return{...base,...fj};}));}if(d.agenda)setAgenda0(prev=>mrg(toArr(d.agenda),prev,"agenda"));const pmk=d.purgeMarks||{};if(d.purgeMarks)setPurgeMarks(prev=>({...prev,...d.purgeMarks}));if(d.loginLogs)setLoginLogs(prev=>{const seen=new Set();const out=[];for(const x of[...toArr(d.loginLogs),...(prev||[])]){if(!x||x.id==null||seen.has(x.id))continue;seen.add(x.id);out.push(x);}return out.filter(x=>!pmk.loginLogs||(x&&x.date||"")>pmk.loginLogs).sort((a,b)=>(b.date||"").localeCompare(a.date||"")).slice(0,800);});if(d.users)setAppUsers0(prev=>{const fb=toArr(d.users);const pm={};(prev||[]).forEach(x=>{if(x&&x.id!=null)pm[x.id]=x;});const out=fb.map(fu=>{const loc=pm[fu.id];if(loc&&(loc.updatedAt||"")>(fu.updatedAt||""))return loc;const base=USERS.find(x=>x.id===fu.id)||{};return{...base,...fu};});const seen=new Set(out.map(x=>x.id));(prev||[]).forEach(x=>{if(x&&x.id!=null&&!seen.has(x.id)&&x.updatedAt)out.push(x);});return out;});if(d.majeurs)setAppMajeurs(toArr(d.majeurs));if(d.rapportsSite)setRapportsSite0(prev=>mrg(toArr(d.rapportsSite),prev,"rapportSite"));if(d.intendance)setIntendance0(prev=>mrg(toArr(d.intendance),prev,"intendance"));if(d.suiviEduc)setSuiviEduc0(prev=>mrg(toArr(d.suiviEduc),prev,"suiviEduc"));if(d.notifs)setNotifs0(prev=>mrg(toArr(d.notifs),prev,"notif"));if(d.transmissions)setTransmissions0(prev=>mrg(toArr(d.transmissions),prev,"transmission"));if(d.djiPlan&&typeof d.djiPlan==="object")setAppDjiPlan(prev=>({...DJI_PLAN,...d.djiPlan}));if(d.fatPlan&&typeof d.fatPlan==="object")setAppFatPlan(prev=>({...FAT_PLAN,...d.fatPlan}));if(d.deletionLogs)setDeletionLogs(prev=>{const seen=new Set();const out=[];for(const x of[...toArr(d.deletionLogs),...(prev||[])]){if(!x||x.id==null||seen.has(x.id))continue;seen.add(x.id);out.push(x);}return out.filter(x=>!pmk.deletionLogs||(x&&x.date||"")>pmk.deletionLogs);});if(d.projets)setProjets(toArr(d.projets));if(d.sejourConfig)setSejourConfig(d.sejourConfig);if(d.etabConfig)setEtabConfig(prev=>({...ETAB_DEFAULT,...prev,...d.etabConfig}));fbLoaded.current=true;};
const[refreshing,setRefreshing]=useState(false);
const[lastSync,setLastSync]=useState(null);
const refreshAll=useCallback(async(auto)=>{
  const silent=auto===true;
  if(refreshing)return;
  setRefreshing(true);if(!silent)setSyncMsg("Synchronisation…");
  try{
    const hadPending=!!window._fbTimer;
    if(window._fbTimer){clearTimeout(window._fbTimer);window._fbTimer=null;}
    if(window._fbFlush&&fbReadOK.current&&(!silent||hadPending)){try{await window._fbFlush();}catch(e){}}
    const r=await fetchTO(FB_URL+"/data.json?auth="+FB_SECRET,null,20000);
    if(!r.ok)throw new Error("HTTP "+r.status);
    const val=await r.json();
    if(val){loadFb(val);fbReadOK.current=true;fbLoaded.current=true;setCloudReady(true);}
    setLastSync(new Date());if(!silent){setSyncMsg("À jour");setTimeout(()=>setSyncMsg(""),1500);}
  }catch(e){if(!silent){setSyncMsg("Échec de synchronisation — données locales conservées");setTimeout(()=>setSyncMsg(""),3500);}}
  setRefreshing(false);
},[refreshing,loadFb]);
useEffect(()=>{
  if(!_curBundle)return;
  let stop=false;
  const check=async()=>{try{const r=await fetchTO(window.location.origin+"/index.html?u="+Date.now(),{cache:"no-store"},10000);if(!r.ok)return;const h=await r.text();const m=h.match(/src="(\/assets\/[^"]+\.js)"/);if(m&&m[1]&&m[1]!==_curBundle&&!stop)setUpdateReady(true);}catch(e){}};
  const iv=setInterval(check,300000);
  const vh=()=>{if(document.visibilityState==="visible")check();};
  document.addEventListener("visibilitychange",vh);
  check();
  return()=>{stop=true;clearInterval(iv);document.removeEventListener("visibilitychange",vh);};
},[]);
useEffect(()=>{const h=()=>{if(document.visibilityState==="visible"&&user)refreshAll(true);};document.addEventListener("visibilitychange",h);return()=>document.removeEventListener("visibilitychange",h);},[refreshAll,user]);
useEffect(()=>{let cancelled=false;let tries=0;
  const attempt=async()=>{tries++;let ok=false,val=null;
    try{const r=await fetch(FB_URL+"/data.json?auth="+FB_SECRET);if(r.ok){val=await r.json();ok=true;}}catch(e){}
    if(cancelled)return;
    if(ok){if(val)loadFb(val);fbReadOK.current=true;fbLoaded.current=true;setCloudReady(true);}
    else{if(tries>=8)fbLoaded.current=true;/* fbReadOK reste faux : aucune écriture tant que la base n'est pas jointe */setTimeout(attempt,tries<8?3000:10000);}
  };
  attempt();return()=>{cancelled=true;};
},[]);
// Migration : normalise les jeuneId de l'agenda en numérique (corrige les anciens enregistrements en chaîne)
useEffect(()=>{setAgenda(prev=>{if(!Array.isArray(prev))return prev;let changed=false;const next=prev.map(a=>{if(a&&typeof a.jeuneId==="string"&&a.jeuneId!==""&&!isNaN(Number(a.jeuneId))){changed=true;return{...a,jeuneId:Number(a.jeuneId)};}return a;});return changed?next:prev;});},[]);
// Données globales : sauvegarde, restauration, synchronisation
const collectData=()=>({rapports,presences,evenements,jeunes:appJeunes,users:appUsers,agenda,loginLogs,majeurs:appMajeurs,rapportsSite,intendance,suiviEduc,notifs,transmissions,djiPlan:appDjiPlan,fatPlan:appFatPlan,deletionLogs,projets,sejourConfig,etabConfig,_exportTs:new Date().toISOString()});
const restoreData=(d)=>{if(!d)return;
  const dedup=(arr,keyf)=>{const seen=new Set();const out=[];for(const x of arr){if(!x)continue;const k=keyf(x);if(seen.has(k))continue;seen.add(k);out.push(x);}return out;};
  const inR=d.rapports?toArr(d.rapports):[];const inE=d.evenements?toArr(d.evenements):[];
  const inA=(d.agenda?toArr(d.agenda):[]).map(a=>a&&typeof a.jeuneId==="string"&&!isNaN(Number(a.jeuneId))?{...a,jeuneId:Number(a.jeuneId)}:a);
  const mergedR=dedup([...(rapports||[]),...inR],r=>(r.horodatage||r.id||"")+"|"+r.jeuneId+"|"+((r.observation||"").slice(0,40)));
  const mergedE=dedup([...(evenements||[]),...inE],e=>(e.horodatage||e.id||"")+"|"+e.jeuneId+"|"+(e.titre||""));
  const mergedA=dedup([...(agenda||[]),...inA],a=>(a.id||"")+"|"+a.jeuneId+"|"+(a.date||"")+"|"+(a.heure||""));
  setRapports(mergedR);setEvenements(mergedE);setAgenda(mergedA);
  if(d.presences)setPresences0(Array.isArray(d.presences)?d.presences.filter(Boolean):(typeof d.presences==="object"?Object.values(d.presences).filter(Boolean):[]));
  if(d.jeunes&&toArr(d.jeunes).length)setAppJeunes(toArr(d.jeunes));
  if(d.users&&toArr(d.users).length)setAppUsers(toArr(d.users));
  if(d.majeurs&&toArr(d.majeurs).length)setAppMajeurs(toArr(d.majeurs));
  if(d.rapportsSite&&toArr(d.rapportsSite).length)setRapportsSite(toArr(d.rapportsSite));
  if(d.djiPlan&&typeof d.djiPlan==="object")setAppDjiPlan(d.djiPlan);
  if(d.fatPlan&&typeof d.fatPlan==="object")setAppFatPlan(d.fatPlan);
  if(d.deletionLogs)setDeletionLogs(prev=>{const seen=new Set();const out=[];for(const x of[...toArr(d.deletionLogs),...(prev||[])]){if(!x||x.id==null||seen.has(x.id))continue;seen.add(x.id);out.push(x);}return out.filter(x=>!pmk.deletionLogs||(x&&x.date||"")>pmk.deletionLogs);});
  if(d.projets&&toArr(d.projets).length)setProjets(toArr(d.projets));
  if(d.sejourConfig)setSejourConfig(d.sejourConfig);
  if(d.etabConfig)setEtabConfig(prev=>({...ETAB_DEFAULT,...prev,...d.etabConfig}));
  fbWasNonEmpty.current=true;fbReadOK.current=true;fbLoaded.current=true;
  const data={rapports:mergedR,presences:(d.presences||presences),evenements:mergedE,jeunes:(d.jeunes&&toArr(d.jeunes).length?toArr(d.jeunes):appJeunes),users:(d.users&&toArr(d.users).length?toArr(d.users):appUsers),agenda:mergedA,loginLogs,majeurs:(d.majeurs&&toArr(d.majeurs).length?toArr(d.majeurs):appMajeurs),rapportsSite:(d.rapportsSite&&toArr(d.rapportsSite).length?toArr(d.rapportsSite):rapportsSite),intendance:(d.intendance&&toArr(d.intendance).length?toArr(d.intendance):intendance),suiviEduc:(d.suiviEduc&&toArr(d.suiviEduc).length?toArr(d.suiviEduc):suiviEduc),djiPlan:appDjiPlan,fatPlan:appFatPlan,deletionLogs,projets:(d.projets&&toArr(d.projets).length?toArr(d.projets):projets),sejourConfig:(d.sejourConfig||sejourConfig),etabConfig};
  fbSet("data",data);fbSet("backups/restauration_"+new Date().toISOString().slice(0,19).replace(/[:T]/g,"-"),{...data,_ts:new Date().toISOString()});
};
const forcePush=async()=>{setSyncMsg("Envoi vers le cloud…");try{await fbSet("data",collectData());setSyncMsg("✓ Données envoyées au cloud");}catch(e){setSyncMsg("✗ Échec de l'envoi");}setTimeout(()=>setSyncMsg(null),3500);};
const forcePull=async()=>{setSyncMsg("Récupération depuis le cloud…");try{const d=await fbGet("data");if(d){restoreData(d);setSyncMsg("✓ Données rechargées depuis le cloud");}else setSyncMsg("Aucune donnée distante");}catch(e){setSyncMsg("✗ Échec de la récupération");}setTimeout(()=>setSyncMsg(null),3500);};
const checkIntegrity=async()=>{try{const d=await fbGet("data")||{};const loc=collectData();const keys=[["rapports","Rapports"],["presences","Présences"],["evenements","Événements"],["jeunes","Jeunes"],["majeurs","Majeurs"],["agenda","Agenda"],["projets","Projets"],["rapportsSite","Rapports de site"]];const cnt=v=>!v?0:Array.isArray(v)?v.length:Object.keys(v).length;return keys.map(([k,l])=>({label:l,local:cnt(loc[k]),cloud:cnt(d[k])}));}catch(e){return null;}};

  if(!user)return<Login users={appUsers} onLogin={u=>{const base=USERS.find(x=>x.id===u.id||x.email===u.email);if(base&&base.isAdmin)u={...u,isAdmin:true};if(u.role==="educateur"||u.role==="coordinateur_site"){const pool=u.isEducMajeur?[...(appMajeurs||MAJEURS)]:[...(appJeunes||JEUNES)];u.assignedIds=pool.filter(j=>j.referentA===u.name||j.referentB===u.name).map(j=>j.id);}setLoginLogs(prev=>[{id:Date.now(),user:u.name||u.login,role:u.role,date:new Date().toISOString(),ts:Date.now()},...prev].slice(0,500));setUser(u);setPage("dashboard");}}/>;
  const liveUser=(()=>{if(!user)return user;const cu=(appUsers||[]).find(x=>x.id===user.id);if(!cu)return user;return{...user,role:cu.role||user.role,site:cu.site||user.site,isAdmin:user.isAdmin||cu.isAdmin,disabled:cu.disabled};})();
  const effUser=viewAs||liveUser;
  const isImpersonating=!!viewAs;
  const pushNotif=(type,titre,message,site)=>{const ts=new Date().toISOString();setNotifs(p=>[...(p||[]),{id:Date.now()+Math.floor(Math.random()*1e4),type,titre,message:message||"",site:site||"",date:ts,updatedAt:ts,author:(user&&user.name)||"",readBy:[]}].slice(-400));};
  const addR=({jeuneId,date,observation,typeContact,traitementAdministre})=>{const now=new Date();const pad2=n=>String(n).padStart(2,"0");const ts=now.getFullYear()+"-"+pad2(now.getMonth()+1)+"-"+pad2(now.getDate())+"T"+pad2(now.getHours())+":"+pad2(now.getMinutes())+":"+pad2(now.getSeconds());setRapports(p=>[...p,{id:Date.now(),jeuneId,date,observation,typeContact:typeContact||"journee",traitementAdministre:traitementAdministre||null,createdAt:ts,horodatage:ts,author:user?.name||"?",authorId:user?.id??null}]);};
  const logDeletion=(type,item)=>{const allJ=[...JEUNES,...MAJEURS,...(appMajeurs||[])];const j=allJ.find(x=>x.id===item.jeuneId);setDeletionLogs(prev=>[{id:Date.now(),origId:item.id,type,date:new Date().toISOString(),by:user?.name||"?",itemDate:item.date||"?",jeune:j?(j.prenom+" "+(j.nom||"")):("ID:"+item.jeuneId),detail:type==="rapport"?(item.observation||"").substring(0,100):(item.titre||item.description||"").substring(0,100)},...prev].slice(0,1000));};
  const delR=(id)=>{const item=rapports.find(r=>r.id===id);if(item)logDeletion("rapport",item);setRapports(p=>p.filter(r=>r.id!==id));};
  const delE=(id)=>{const item=evenements.find(e=>e.id===id);if(item)logDeletion("evenement",item);setEvenements(p=>p.filter(e=>e.id!==id));};
  const addE=ev=>{const now=new Date();const pad2=n=>String(n).padStart(2,"0");const ts=now.getFullYear()+"-"+pad2(now.getMonth()+1)+"-"+pad2(now.getDate())+"T"+pad2(now.getHours())+":"+pad2(now.getMinutes())+":"+pad2(now.getSeconds());const newEv={id:Date.now(),...ev,author:user?.name||"?",authorId:user?.id??null,createdAt:ts,horodatage:ts};setEvenements(p=>[...p,newEv]);try{const _p=[...(appJeunes||[]),...(appMajeurs||[])];const _j=_p.find(x=>x&&x.id===ev.jeuneId);const _nom=_j?(_j.prenom+" "+(_j.nom||"")):"";fetch("/api/notify",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({titre:ev.titre||"",description:ev.description||"",gravite:ev.gravite||"",eig:!!ev.eig,site:ev.site||(_j&&_j.site)||"",jeune:_nom,author:(user&&user.name)||"",date:ev.date||""})}).catch(()=>{});const _eig=!!ev.eig||/grave|critique/i.test(ev.gravite||"");pushNotif(_eig?"eig":"evenement",_eig?"EIG déclaré":"Événement indésirable",(_nom?_nom+" — ":"")+(ev.titre||""),ev.site||(_j&&_j.site)||"");}catch(e){}try{const j=(appJeunes||[]).find(j2=>j2.id===ev.jeuneId);const nom=j?(j.prenom+" "+(j.nom||"")):("ID:"+ev.jeuneId);const bom="﻿";const rows=[["Date","Jeune","Titre","Description","Gravité"],[ev.date||"",nom,ev.titre||"",ev.description||"",ev.gravite||"normal"]];const csv=rows.map(r=>r.map(c=>'"'+String(c).replace(/"/g,'""')+'"').join(";")).join("\n");const blob=new Blob([bom+csv],{type:"text/csv;charset=utf-8"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="incident_"+((ev.date||"").replace(/-/g,""))+"_"+nom.replace(/\s/g,"_")+".csv";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}catch(e){console.error("Auto CSV export error:",e);}};
  const changeP=(jeuneId,date,statut)=>setPresences(p=>[...p.filter(p2=>!(p2.jeuneId===jeuneId&&p2.date===date)),{id:`${jeuneId}-${date}`,jeuneId,date,statut}]);
  const delP=(jeuneId,date)=>setPresences(p=>(p||[]).filter(x=>!(x.jeuneId===jeuneId&&x.date===date)));
  const visibleNotifs=(notifs||[]).filter(n=>{if(effUser.role==="directeur"||effUser.role==="chef_service")return true;if(effUser.site&&effUser.site!=="Tous"&&n.site&&n.site!==effUser.site)return false;return true;});
  const TITLES={transmissions:"Cahier de transmissions",recherche:"Recherche globale",dashboard:"Tableau de bord",jeunes:"Jeunes","jeune-detail":sel?`${sel.prenom} ${sel.nom}`:"Fiche",rapports:"Rapports journaliers",presences:"Présences",evenements:"Événements indésirables","rapport-hebdo":"Rapport hebdomadaire",projets:"Projets personnalisés","rapport-site":"Rapport de site","espace-rh":"Espace éducateur"};
  return(<div style={{fontFamily:"'Nunito',sans-serif",background:`linear-gradient(180deg,${C.sableLight} 0%,#F5EFE0 40%,#EDE4D0 100%)`,minHeight:"100vh",display:"flex",flexDirection:"column"}}>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap');*{box-sizing:border-box;-webkit-tap-highlight-color:transparent;-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale}html{scroll-behavior:smooth}body{margin:0}button{position:relative;transition:all 0.25s cubic-bezier(0.34,1.56,0.64,1);cursor:pointer}button:hover:not(:disabled){transform:translateY(-2px) scale(1.02);filter:brightness(1.06);box-shadow:0 6px 16px rgba(0,0,0,0.14)}button:active:not(:disabled){transform:translateY(1px) scale(0.95);transition-duration:0.08s;animation:ripple 0.45s ease-out}button:focus-visible{outline:2px solid ${C.gold};outline-offset:2px}button:disabled{transform:none!important;box-shadow:none!important;filter:grayscale(0.4) opacity(0.7);cursor:not-allowed}@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}input,select,textarea{transition:all 0.2s ease}input:focus,select:focus,textarea:focus{border-color:${C.gold}!important;outline:none;box-shadow:0 0 0 3px ${C.goldGlow}}::selection{background:${C.goldLight};color:${C.dark}}::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:${C.sableDark};border-radius:99px}::-webkit-scrollbar-thumb:hover{background:${C.gold}}@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
@keyframes fadeIn{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}@keyframes slideIn{from{opacity:0;transform:translateX(-16px)}to{opacity:1;transform:translateX(0)}}@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}@keyframes countUp{from{opacity:0;transform:scale(0.8)}to{opacity:1;transform:scale(1)}}@keyframes shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}@keyframes cardEnter{from{opacity:0;transform:translateY(16px) scale(0.97)}to{opacity:1;transform:translateY(0) scale(1)}}@keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-4px)}}@keyframes ripple{0%{box-shadow:0 0 0 0 ${C.goldGlow}}100%{box-shadow:0 0 0 12px rgba(184,134,11,0)}}select{cursor:pointer;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%238B7050' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='m6 9 6 6 6-6'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:32px!important}`}</style>
    {isImpersonating&&<div style={{position:"fixed",top:0,left:0,right:0,zIndex:60,background:"linear-gradient(90deg,"+C.gold+","+C.goldDark+")",color:"#fff",padding:"8px 14px",display:"flex",alignItems:"center",justifyContent:"center",gap:12,fontSize:13,fontWeight:700,boxShadow:"0 2px 12px rgba(0,0,0,0.2)",flexWrap:"wrap"}}><span>👁 Vue de <b>{effUser.name}</b> ({effUser.role==="educateur"?"Éducateur"+(effUser.site&&effUser.site!=="Tous"?" · "+effUser.site:""):effUser.role==="coordinateur_site"?"Coordinateur":effUser.role==="directeur"?"Directeur":"Chef de service"})</span><button onClick={()=>{setViewAs(null);setPage("admin");setSel(null);}} style={{background:C.white,color:C.goldDark,border:"none",borderRadius:8,padding:"4px 12px",fontWeight:800,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>Quitter</button></div>}
    {updateReady&&<div onClick={async()=>{try{if(window._fbTimer){clearTimeout(window._fbTimer);window._fbTimer=null;}if(window._fbFlush&&fbReadOK.current)await window._fbFlush();}catch(e){}window.location.reload();}} style={{position:"fixed",top:0,left:0,right:0,zIndex:80,background:"#1B5E20",color:"#fff",padding:"11px 14px",fontSize:13.5,fontWeight:800,textAlign:"center",cursor:"pointer",boxShadow:"0 2px 12px rgba(0,0,0,0.35)"}}>Nouvelle version disponible — toucher ici pour mettre à jour</div>}
    {user&&!cloudReady&&<div style={{position:"fixed",top:0,left:0,right:0,zIndex:75,background:"#B26A00",color:"#fff",padding:"7px 14px",fontSize:12.5,fontWeight:700,textAlign:"center"}}>Connexion à la base en cours — vos saisies sont conservées sur cet appareil et partiront automatiquement</div>}
    {syncMsg&&<div style={{position:"fixed",bottom:18,left:"50%",transform:"translateX(-50%)",zIndex:70,background:C.dark,color:"#fff",padding:"10px 18px",borderRadius:12,fontSize:13,fontWeight:700,boxShadow:C.shadowLg,animation:"fadeIn 0.3s ease"}}>{syncMsg}</div>}
    <div style={{height:isImpersonating?34:0}}/>
    <Sidebar page={page} onNav={setPage} user={effUser} onLogout={()=>{setViewAs(null);setUser(null);}} open={open} onClose={()=>setOpen(false)}/>
    <div style={{flex:1,display:"flex",flexDirection:"column"}}>
      <Topbar title={TITLES[page]||"PDSR"} onMenu={()=>setOpen(true)} onBack={page==="jeune-detail"?()=>setPage("jeunes"):undefined} unread={visibleNotifs.filter(n=>!(n.readBy||[]).includes(effUser.id)).length} onBell={()=>setNotifOpen(o=>!o)} onRefresh={refreshAll} refreshing={refreshing} lastSync={lastSync} theme={theme} onTheme={()=>setTheme(t=>t==="dark"?"light":"dark")}/>
      {notifOpen&&<NotifPanel notifs={visibleNotifs} user={effUser} onClose={()=>setNotifOpen(false)} onOpen={(n)=>{const dest={eig:"evenements",evenement:"evenements",intendance:"intendance",intendance_urgente:"intendance",rapport_site:"rapport-site",transmission:"transmissions"}[n.type]||"dashboard";const ts=new Date().toISOString();setNotifs(p=>(p||[]).map(x=>x.id===n.id&&!(x.readBy||[]).includes(effUser.id)?{...x,readBy:[...(x.readBy||[]),effUser.id],updatedAt:ts}:x));setNotifOpen(false);setPage(dest);}} onReadAll={()=>{const ts=new Date().toISOString();setNotifs(p=>(p||[]).map(n=>(n.readBy||[]).includes(effUser.id)?n:{...n,readBy:[...(n.readBy||[]),effUser.id],updatedAt:ts}));}}/>}
      <main ref={mainRef} onTouchStart={e=>{const el=mainRef.current;if(el&&el.scrollTop<=0)ptrStart.current=e.touches[0].clientY;else ptrStart.current=null;}} onTouchMove={e=>{if(ptrStart.current==null)return;const d=e.touches[0].clientY-ptrStart.current;if(d>0)setPtr(Math.min(d*0.5,90));}} onTouchEnd={()=>{if(ptr>=60)refreshAll();setPtr(0);ptrStart.current=null;}} style={{flex:1,overflowY:"auto",overflowX:"hidden",position:"relative",width:"100%",maxWidth:"100vw",boxSizing:"border-box"}}>
        {ptr>0&&<div style={{height:ptr,display:"flex",alignItems:"center",justifyContent:"center",color:C.gold,transition:"height 0.15s",flexWrap:"wrap"}}><RefreshCw size={20} style={{transform:"rotate("+(ptr*4)+"deg)"}}/><span style={{marginLeft:8,fontSize:12,fontWeight:700}}>{ptr>=60?"Relâchez pour actualiser":"Tirez pour actualiser"}</span></div>}
        <GlobalFX/><div key={page} className="pg-anim" style={{maxWidth:"100%",overflowX:"hidden"}}>
        {page==="dashboard"&&<Dashboard projets={projets} etabConfig={etabConfig} setPage={setPage} user={effUser} rapports={rapports} presences={presences} evenements={evenements} onNav={setPage} setSel={setSel} jeunes={appJeunes} agenda={agenda} majeurs={appMajeurs} intendance={intendance} suiviEduc={suiviEduc} users={appUsers}/>}
        {page==="jeunes"&&<JeunesList user={effUser} jeunes={appJeunes} presences={presences} onSelect={setSel} onNav={setPage} onUpdateJeune={(id,field,val)=>{setAppJeunes(prev=>prev.map(j=>j.id===id?{...j,[field]:val}:j));}}/>}
        {page==="majeurs"&&<div><div style={{...S.card,marginBottom:12}}><div style={{fontWeight:700,fontSize:16,color:C.dark,marginBottom:12}}>Jeunes Majeurs</div><div style={{fontSize:12,color:C.light,marginBottom:8}}>Section des jeunes majeurs</div></div>{(appMajeurs||MAJEURS).map(m=><div key={m.id} onClick={()=>{setSel(m);setPage("majeur-detail");}} style={{...S.card,marginBottom:8,cursor:"pointer",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}><div style={{width:36,height:36,borderRadius:18,background:C.primary,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:13,flexWrap:"wrap"}}>{m.prenom[0]}{(m.nom||"")[0]||""}</div><div><div style={{fontWeight:700,color:C.dark,fontSize:14}}>{m.prenom} {m.nom}</div><div style={{fontSize:12,color:C.light}}>{m.site} | {m.dateDebut} - {m.dateFin}</div></div></div>)}</div>}
        {page==="majeur-detail"&&sel&&<MajeurDetail projets={projets} onUpdateProjets={setProjets} etabConfig={etabConfig} majeur={(appMajeurs||[]).find(m=>m.id===sel.id)||sel} onUpdateMajeur={(id,field,val)=>setAppMajeurs(prev=>(prev||[]).map(m=>m.id===id?{...m,[field]:val}:m))} rapports={rapports} presences={presences} evenements={evenements} user={effUser} onBack={()=>setPage("majeurs")} onAddR={j=>{setSel(j);setPage("rapports");}} onAddE={j=>{setSel(j);setPage("evenements");}} onCP={changeP} users={appUsers} addR={r=>{addR(r);}} addE={ev=>{addE(ev);}}/>}
        {page==="jeune-detail"&&sel&&<JeuneDetail projets={projets} onUpdateProjets={setProjets} etabConfig={etabConfig} jeune={appJeunes.find(j=>j.id===sel.id)||sel} rapports={rapports} presences={presences} evenements={evenements} user={effUser} onAddR={j=>{setSel(j);setPage("rapports");}} onAddE={j=>{setSel(j);setPage("evenements");}} onCP={changeP} users={appUsers} onUpdateJeune={(id,field,val)=>{setAppJeunes(prev=>prev.map(j=>j.id===id?{...j,[field]:val}:j));}}/>}
        {page==="rapports"&&<Rapports user={effUser} rapports={rapports} jeunes={appJeunes} onSave={addR} onDelete={delR} onUpdate={(id,field,val)=>setRapports(p=>p.map(r=>r.id===id?{...r,[field]:val}:r))} onPatch={(id,patch)=>setRapports(p=>p.map(r=>r.id===id?{...r,...patch}:r))} majeurs={appMajeurs}/>}
        
        {page==="evenements"&&<Evenements user={effUser} evenements={evenements} onAdd={addE} onDelete={delE} majeurs={appMajeurs} jeunes={appJeunes} onUpdateAll={setEvenements}/>}
        {page==="agenda"&&<AgendaPage agenda={agenda} setAgenda={setAgenda} jeunes={appJeunes} majeurs={MAJEURS} users={appUsers} user={effUser}/>}
        {page==="transmissions"&&<Transmissions user={effUser} items={transmissions} onAdd={({site,texte,type})=>{const ts=new Date().toISOString();setTransmissions(p=>[...(p||[]),{id:Date.now(),site,texte,type:type||"transmission",author:(effUser&&effUser.name)||"",date:ts,updatedAt:ts,luPar:[]}]);pushNotif("transmission","Nouvelle transmission",String(texte).slice(0,100),site);}} onMarkLu={(id)=>{const ts=new Date().toISOString();setTransmissions(p=>(p||[]).map(t=>t.id===id?{...t,luPar:[...(t.luPar||[]),{name:(effUser&&effUser.name)||"",date:ts}],updatedAt:ts}:t));}}/>}
        {page==="recherche"&&<RechercheGlobale user={effUser} rapports={rapports} evenements={evenements} transmissions={transmissions} rapportsSite={rapportsSite} intendance={intendance} jeunes={appJeunes} majeurs={appMajeurs}/>}
        {page==="intendance"&&(effUser.role==="coordinateur_site"||effUser.role==="chef_service"||effUser.role==="directeur")&&<Intendance user={effUser} items={intendance} onSave={r=>{const isNew=!(intendance||[]).some(x=>x.id===r.id);setIntendance(prev=>{const idx=prev.findIndex(x=>x.id===r.id);if(idx>=0){const cp=[...prev];cp[idx]=r;return cp;}return[...prev,r];});if(isNew)pushNotif(r.urgence==="Urgente"?"intendance_urgente":"intendance",r.urgence==="Urgente"?"Demande d’intendance URGENTE":"Demande d’intendance",(r.objet||"")+(r.description?" — "+String(r.description).slice(0,80):""),r.site||"");}} onDelete={id=>{setDeletionLogs(prev=>[...prev,{id:"dl_"+Date.now(),type:"intendance",origId:id,date:new Date().toISOString()}]);setIntendance(prev=>prev.filter(x=>x.id!==id));}}/>}
        {page==="pres-educ"&&(effUser.role==="coordinateur_site"||effUser.role==="chef_service"||effUser.role==="directeur")&&<PresEduc user={effUser} users={appUsers} entries={suiviEduc} onSave={r=>{setSuiviEduc(prev=>{const idx=prev.findIndex(x=>x.id===r.id);if(idx>=0){const cp=[...prev];cp[idx]=r;return cp;}return[...prev,r];});}} onDelete={id=>{setDeletionLogs(prev=>[...prev,{id:"dl_"+Date.now(),type:"suiviEduc",origId:id,date:new Date().toISOString()}]);setSuiviEduc(prev=>prev.filter(x=>x.id!==id));}}/>}
        {page==="rapport-site"&&(effUser.role==="coordinateur_site"||effUser.role==="chef_service"||effUser.role==="directeur")&&<RapportSite user={effUser} rapportsSite={rapportsSite} onSave={r=>{const isNew=!(rapportsSite||[]).some(x=>x.id===r.id);setRapportsSite(prev=>{const idx=prev.findIndex(x=>x.id===r.id);if(idx>=0){const cp=[...prev];cp[idx]=r;return cp;}return[...prev,r];});if(isNew)pushNotif("rapport_site","Rapport de site déposé",(r.site||"")+((r.semaine||r.periode)?" · "+(r.semaine||r.periode):""),r.site||"");}} onDelete={id=>{setDeletionLogs(prev=>[...prev,{id:"dl_"+Date.now()+"_rapportSite_"+id,type:"rapportSite",origId:id,date:new Date().toISOString()}]);setRapportsSite(prev=>prev.filter(x=>x.id!==id));}}/>}
        {page==="export"&&(effUser.role==="directeur"||effUser.role==="chef_service"||effUser.role==="coordinateur_site")&&<ExportPage peutPurger={!!effUser.isAdmin} purgeRanges={(purgeMarks&&purgeMarks.ranges)||[]} onCancelRange={(ts)=>{if(!effUser||!effUser.isAdmin){alert("Seul l'administrateur peut annuler une purge.");return;}setPurgeMarks(prev=>({...prev,ranges:(prev.ranges||[]).filter(r=>r.ts!==ts)}));}} sejourConfig={sejourConfig} rapports={rapports} evenements={evenements} agenda={agenda} jeunes={appJeunes} majeurs={appMajeurs} rapportsSite={rapportsSite} onPurge={(from,to,scope)=>{const sc=scope||{rapports:true,evenements:true,agenda:true};const ts=nowSrv();purgeIntent.current=true;setPurgeMarks(prev=>({...prev,lastPurge:ts,ranges:[...(prev&&prev.ranges||[]),{t:[sc.rapports&&"rapport",sc.evenements&&"evenement",sc.agenda&&"agenda",sc.rapportsSite&&"rapportSite"].filter(Boolean),from,to,ts}].slice(-15)}));const base=Date.now();const mk=(arr,type)=>(arr||[]).filter(x=>x&&x.date>=from&&x.date<=to&&x.id!=null).map(x=>({id:"dl_"+base+"_"+type+"_"+x.id,type,origId:x.id,date:ts}));const tombs=[];if(sc.rapports)tombs.push(...mk(rapports,"rapport"));if(sc.evenements)tombs.push(...mk(evenements,"evenement"));if(sc.agenda)tombs.push(...mk(agenda,"agenda"));if(sc.rapportsSite)tombs.push(...mk(rapportsSite,"rapportSite"));if(tombs.length)setDeletionLogs(prev=>[...prev,...tombs]);if(sc.rapports)setRapports(p=>p.filter(r=>r.date<from||r.date>to));if(sc.evenements)setEvenements(p=>p.filter(e=>e.date<from||e.date>to));if(sc.agenda)setAgenda(p=>p.filter(a=>a.date<from||a.date>to));if(sc.rapportsSite)setRapportsSite(p=>p.filter(r=>r.date<from||r.date>to));}}/>}
      {page==="admin"&&(effUser.role==="directeur"||effUser.role==="chef_service")&&<Admin djiPlan={appDjiPlan} fatPlan={appFatPlan} onBulkPlan={(siteName,maj)=>{if(siteName==="Djilass")setAppDjiPlan(prev=>({...prev,...maj}));else setAppFatPlan(prev=>({...prev,...maj}));}} currentUser={effUser} isAdmin={effUser.isAdmin} onRefresh={refreshAll} etabConfig={etabConfig} onUpdateEtab={setEtabConfig} onArchiveSejour={async(label)=>{setSyncMsg("Archivage du séjour…");try{const snap={...collectData(),archivedAt:new Date().toISOString(),label:label||("Séjour "+today)};await fbSet("archives/"+Date.now(),snap);const blob=new Blob([JSON.stringify(snap,null,2)],{type:"application/json"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="archive_sejour_"+today+".json";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);setSyncMsg("✓ Séjour archivé (cloud + fichier)");}catch(e){setSyncMsg("✗ Échec de l'archivage");}setTimeout(()=>setSyncMsg(null),4000);}} onViewAs={(u)=>{setViewAs(u);setPage("dashboard");setSel(null);setOpen(false);}} onForcePush={forcePush} onForcePull={forcePull} onCheckIntegrity={checkIntegrity} onBackup={collectData} onRestore={restoreData} rapports={rapports} evenements={evenements} sejourConfig={sejourConfig} onUpdateSejours={(s,d)=>setSejourConfig(p=>({...p,[s]:{...(p&&p[s]||{}),dateDebut:d}}))} users={appUsers} jeunes={appJeunes} onUpdateUsers={setAppUsers} onUpdateJeunes={setAppJeunes} loginLogs={loginLogs} appMajeurs={appMajeurs} onUpdateMajeurs={(id,field,val,fullArr)=>{if(fullArr){setAppMajeurs(fullArr);}else{setAppMajeurs(prev=>(prev||MAJEURS).map(m=>m.id===id?{...m,[field]:val}:m));}}} deletionLogs={deletionLogs} onResetGlobal={async()=>{
   const stamp=Date.now();
   setAppJeunes([]);setAppMajeurs([]);setRapports([]);setEvenements([]);setPresences([]);setProjets([]);setAgenda([]);
   try{localStorage.removeItem(LS_KEY);}catch(e){}
   setLocalReset(stamp);
   try{
     await fbSet("data/jeunes",[]);await fbSet("data/majeurs",[]);await fbSet("data/rapports",[]);
     await fbSet("data/evenements",[]);await fbSet("data/presences",[]);await fbSet("data/projets",[]);await fbSet("data/agenda",[]);
     const iso=new Date().toISOString();
     setPurgeMarks(prev=>({...prev,lastPurge:iso}));
     await fbSet("data/purgeMarks/lastPurge",iso);
     await fbSet("data/resetStamp",stamp);
   }catch(e){alert("La réinitialisation serveur a échoué : "+(e&&e.message?e.message:e)+"\nRien ne garantit que les autres appareils soient purgés.");return;}
   alert("Réinitialisation effectuée. Les autres appareils se videront à leur prochaine synchronisation.");
   window.location.reload();
 }} onPurgeLogs={()=>{setPurgeMarks(p=>({...p,loginLogs:nowSrv()}));setLoginLogs([]);}} onPurgeDeletionLogs={()=>{setPurgeMarks(p=>({...p,deletionLogs:nowSrv()}));setDeletionLogs([]);}} presences={presences} onChangeP={changeP} agenda={agenda} onUpdateAgenda={setAgenda} projets={projets} rapportsSite={rapportsSite} onUpdateRapportsSite={setRapportsSite} onDeleteRapport={delR} onUpdateRapport={(id,patch)=>setRapports(p=>p.map(r=>r.id===id?{...r,...patch}:r))} onDeleteEvenement={delE} onUpdateEvenements={setEvenements}/>}
        {page==="rapport-hebdo"&&(effUser.role==="chef_service"||effUser.role==="directeur")&&<RapportHebdo user={effUser} rapports={rapports} presences={presences} evenements={evenements} jeunes={appJeunes} majeurs={appMajeurs} sejourConfig={sejourConfig}/>}
        {page==="projets"&&<ProjetsPersonnalises user={effUser} jeunes={appJeunes} majeurs={appMajeurs} projets={projets} onUpdate={setProjets} etabConfig={etabConfig} users={appUsers}/>}
      {page==="planning"&&<Planning djiPlan={appDjiPlan} fatPlan={appFatPlan} site={effUser.site} user={effUser} onUpdate={(siteName,key,data)=>{if(siteName==="Djilass")setAppDjiPlan(prev=>({...prev,[key]:data}));else setAppFatPlan(prev=>({...prev,[key]:data}));}}/>}
        {page==="espace-rh"&&<EspaceRH user={effUser} docs={docs} users={appUsers} onUpdateUsers={setAppUsers} etabConfig={etabConfig} onAdd={(d)=>setDocs(p=>[...p,d])} onSign={(id,sig)=>setDocs(p=>p.map(x=>x.id===id?{...x,signatures:[...(x.signatures||[]),sig]}:x))} onDelete={(id)=>setDocs(p=>p.filter(x=>x.id!==id))}/>}
      </div></main>
    </div>
  </div>);
}