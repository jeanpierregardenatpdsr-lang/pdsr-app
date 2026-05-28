
import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType, ImageRun, PageBreak, HeadingLevel } from "docx";
import { Home, Users, FileText, Calendar, AlertTriangle, BarChart2, LogOut, Menu, X, ChevronRight, Plus, Check, ChevronLeft, Search, MapPin, Printer, Download } from "lucide-react";


export class ErrorBoundary extends React.Component{constructor(p){super(p);this.state={hasError:false,error:null};}static getDerivedStateFromError(e){return{hasError:true,error:e};}componentDidCatch(e,i){console.error("PDSR Error:",e,i);}render(){if(this.state.hasError){return React.createElement("div",{style:{padding:40,textAlign:"center"}},React.createElement("h2",null,"Une erreur est survenue"),React.createElement("p",null,String(this.state.error)),React.createElement("button",{onClick:()=>{localStorage.removeItem("pdsr_data");window.location.reload();},style:{padding:"10px 20px",background:"#2c6fbb",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",marginTop:16}},"Recharger l'application"));}return this.props.children;}}
const LS_KEY="pdsr_data";
const loadLS=()=>{try{const d=JSON.parse(localStorage.getItem(LS_KEY));return d||null;}catch{return null;}};
const saveLS=(rapports,presences,evenements,users,jeunes,agenda,loginLogs,majeurs,rapportsSite,djiPlan,fatPlan)=>{try{localStorage.setItem(LS_KEY,JSON.stringify({rapports,presences,evenements,users,jeunes,agenda,loginLogs,majeurs,rapportsSite,djiPlan,fatPlan,ts:Date.now()}));}catch{}};
const FB_URL="https://pdsr-app-default-rtdb.firebaseio.com";
const FB_SECRET="GhVuY7AXz2QeB8shFhV4SEZvQAyrrf0YjeOqleEw";
const fbSet=async(p,d)=>{try{await fetch(FB_URL+"/"+p+".json?auth="+FB_SECRET,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(d)});}catch(e){}};
const fbGet=async(p)=>{try{let r=await fetch(FB_URL+"/"+p+".json?auth="+FB_SECRET);return await r.json();}catch(e){return null;}};

const LOGO="data:image/png;base64,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";
const C={gold:"#B8860B",goldDark:"#8B6508",goldLight:"#FDF3D0",orange:"#E07B00",orangeLight:"#FFF0D0",sable:"#EAD29C",sableLight:"#FAF7EE",sableDark:"#D9BE80",dark:"#1A0E00",mid:"#5C4020",light:"#8B7050",white:"#FFFFFF",border:"#E8D5B0",accent:"#6A1B9A",success:"#2E7D32",danger:"#C62828",info:"#1565C0"};
const USERS=[{id:1,email:"lmarcille1962@gmail.com",password:"1789",role:"directeur",name:"Laurent Marcille",initials:"LM",site:"Tous"},{id:2,email:"jeanpierregardenatpdsr@gmail.com",password:"pdsr2026",role:"chef_service",name:"Jean-Pierre Gardenat",initials:"JP",site:"Tous"},{id:3,email:"Omar Ngom",password:"pdsr2026",role:"chef_service",name:"Omar Ngom",initials:"ON",site:"Tous"},{id:4,login:"go",password:"go2026",role:"educateur",name:"Go",initials:"GO",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[1,8]},{id:5,login:"abdoulaye",password:"abdoulaye2026",role:"educateur",name:"Abdoulaye",initials:"AB",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[2]},{id:6,login:"khady",password:"khady2026",role:"educateur",name:"Khady",initials:"KH",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[3,7]},{id:7,login:"charline",password:"charline2026",role:"educateur",name:"Charline",initials:"CH",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[4,9]},{id:8,login:"momar",password:"momar2026",role:"educateur",name:"Momar",initials:"MO",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[5]},{id:9,login:"emile",password:"emile2026",role:"educateur",name:"Emile",initials:"EM",site:"Fatick",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[6]},{id:10,login:"lysa",password:"lysa2026",role:"educateur",name:"Lysa",initials:"LY",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[1,8]},{id:11,login:"babacar",password:"babacar2026",role:"educateur",name:"Babacar",initials:"BA",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[2]},{id:12,login:"diarra",password:"diarra2026",role:"educateur",name:"Diarra",initials:"DI",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[4,9]},{id:13,login:"prospere",password:"prospere2026",role:"educateur",name:"Prospère",initials:"PR",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[6]},{id:14,login:"am",password:"am2026",role:"educateur",name:"Am",initials:"AM",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[5]},{id:15,login:"yacine",password:"yacine2026",role:"educateur",name:"Yacine",initials:"YA",site:"Fatick",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[3,7]},{id:16,login:"sadibou",password:"sadibou2026",role:"educateur",name:"Sadibou",initials:"SA",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[14]},{id:17,login:"ngor",password:"ngor2026",role:"educateur",name:"Ngor",initials:"NG",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[12,16]},{id:18,login:"mahault",password:"mahault2026",role:"educateur",name:"Mahault",initials:"MA",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[15,18]},{id:19,login:"malang",password:"malang2026",role:"educateur",name:"Malang",initials:"MA",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[14]},{id:20,login:"lea",password:"lea2026",role:"educateur",name:"Léa",initials:"LÉ",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[10]},{id:21,login:"assane",password:"assane2026",role:"educateur",name:"Assane",initials:"AS",site:"Djilass",equipe:"A",section:"mineurs",isEducMajeur:false,assignedIds:[11,17]},{id:22,login:"rouge",password:"rouge2026",role:"educateur",name:"Rouge",initials:"RO",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[16]},{id:23,login:"luc",password:"luc2026",role:"educateur",name:"Luc",initials:"LU",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[15]},{id:24,login:"nicole",password:"nicole2026",role:"educateur",name:"Nicole",initials:"NI",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[11,12]},{id:25,login:"bacary",password:"bacary2026",role:"educateur",name:"Bakary",initials:"BA",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[14,18]},{id:26,login:"fatou",password:"fatou2026",role:"educateur",name:"Fatou",initials:"FA",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[10]},{id:27,login:"kourou",password:"kourou2026",role:"educateur",name:"Kourou",initials:"KO",site:"Djilass",equipe:"B",section:"mineurs",isEducMajeur:false,assignedIds:[17]}];
const JEUNES=[{id:1,prenom:"Narondya",nom:"PIERRE",site:"Fatick",referentA:"Go",referentB:"Lysa",statut:"actif"},{id:2,prenom:"Andy",nom:"BRAUN",site:"Fatick",referentA:"Abdoulaye",referentB:"Babacar",statut:"actif"},{id:3,prenom:"Djeneba",nom:"DRAME",site:"Fatick",referentA:"Khady",referentB:"Yacine",statut:"actif"},{id:4,prenom:"Hafssa",nom:"AMINE",site:"Fatick",referentA:"Charline",referentB:"Diarra",statut:"actif"},{id:5,prenom:"Kais",nom:"MISSAOUI",site:"Fatick",referentA:"Momar",referentB:"Am",statut:"actif"},{id:6,prenom:"Djanis",nom:"MEDICO",site:"Fatick",referentA:"Emile",referentB:"Prospère",statut:"actif"},{id:7,prenom:"Louiza",nom:"TOUHARI",site:"Fatick",referentA:"Khady",referentB:"Yacine",statut:"actif"},{id:8,prenom:"Linda",nom:"LAMBLIN",site:"Fatick",referentA:"Go",referentB:"Lysa",statut:"actif"},{id:9,prenom:"Delinda",nom:"BADOURI",site:"Fatick",referentA:"Charline",referentB:"Diarra",statut:"actif"},{id:10,prenom:"Imane",nom:"CAMARA",site:"Djilass",referentA:"Lea",referentB:"Fatou",statut:"actif"},{id:11,prenom:"Venerenda",nom:"HAINAUT",site:"Djilass",referentA:"Assane",referentB:"Nicole",statut:"actif"},{id:12,prenom:"Oceane",nom:"PERLIER",site:"Djilass",referentA:"Ngor",referentB:"Nicole",statut:"actif"},{id:14,prenom:"Kyllian",nom:"PULIDO DIAZ MAROTO",site:"Djilass",referentA:"Sadibou",referentB:"Malang",statut:"actif"},{id:15,prenom:"Mathis",nom:"MARTINEZ",site:"Djilass",referentA:"Mahault",referentB:"Luc",statut:"actif"},{id:16,prenom:"Franck",nom:"MERCIER",site:"Djilass",referentA:"Ngor",referentB:"Rouge",statut:"actif"},{id:17,prenom:"Jimmy",nom:"DISDIER",site:"Djilass",referentA:"Assane",referentB:"Kourou",statut:"actif"},{id:18,prenom:"Wesley",nom:"CARTERET",site:"Djilass",referentA:"Mahault",referentB:"Bakary",statut:"actif"}];
const MAJEURS=[{id:100,prenom:"Sarah",nom:"HELAIMIA",site:"Fatick",referentA:"",referentB:"",statut:"actif",telParent1:"06 11 97 11 22",telJeune:"07 73 90 42 49",emailASE:"stravaglini@vosges.fr",dateDebut:"23/03/2026",dateFin:"30/06/2026"},{id:101,prenom:"Chaina",nom:"ADJABI",site:"Fatick",referentA:"Malang",referentB:"Luc",statut:"actif",emailASE:"b.pouget@sauvegarde2savoie.fr",dateDebut:"11/05/2026",dateFin:"28/08/2026"},{id:102,prenom:"Fatima",nom:"EL HAOUATE",site:"Fatick",referentA:"",referentB:"",statut:"actif",emailASE:"kseri@seinesaintdenis.fr",dateDebut:"13/05/2026",dateFin:"28/08/2026"},{id:103,prenom:"Charline",nom:"LUTONADIO NGENGE",site:"Fatick",referentA:"",referentB:"",statut:"actif",emailASE:"mlimentour@seinesaintdenis.fr",dateDebut:"01/04/2026",dateFin:"12/08/2026"}];
const DJI_PLAN={"2026-04-01":{a:false,b:false},"2026-04-02":{a:false,b:true,n:"Arrivée – WARANG"},"2026-04-03":{a:false,b:true,n:"Départ vers RHO"},"2026-04-04":{a:false,b:true,n:"RHO"},"2026-04-05":{a:true,b:true,n:"RHO"},"2026-04-06":{a:true,b:true,n:"Rando Individuelle"},"2026-04-07":{a:true,b:true,n:"Rando Individuelle"},"2026-04-08":{a:true,b:true,n:"Arrivée DJILASS"},"2026-04-09":{a:true,b:false},"2026-04-10":{a:true,b:false},"2026-04-11":{a:true,b:false},"2026-04-12":{a:true,b:false},"2026-04-13":{a:false,b:true},"2026-04-14":{a:false,b:true},"2026-04-15":{a:false,b:true},"2026-04-16":{a:false,b:true},"2026-04-17":{a:false,b:true},"2026-04-18":{a:false,b:true},"2026-04-19":{a:false,b:true},"2026-04-20":{a:true,b:false},"2026-04-21":{a:true,b:false},"2026-04-22":{a:true,b:false},"2026-04-23":{a:true,b:false},"2026-04-24":{a:true,b:false},"2026-04-25":{a:true,b:false},"2026-04-26":{a:true,b:false},"2026-04-27":{a:false,b:true},"2026-04-28":{a:false,b:true},"2026-04-29":{a:false,b:true},"2026-04-30":{a:false,b:true},"2026-05-01":{a:false,b:true},"2026-05-02":{a:false,b:true},"2026-05-03":{a:false,b:true},"2026-05-04":{a:true,b:false},"2026-05-05":{a:true,b:false},"2026-05-06":{a:true,b:false},"2026-05-07":{a:true,b:false},"2026-05-08":{a:true,b:false},"2026-05-09":{a:true,b:false},"2026-05-10":{a:true,b:false},"2026-05-11":{a:false,b:true},"2026-05-12":{a:false,b:true},"2026-05-13":{a:false,b:true},"2026-05-14":{a:false,b:true},"2026-05-15":{a:false,b:true},"2026-05-16":{a:false,b:true},"2026-05-17":{a:false,b:true},"2026-05-18":{a:true,b:false},"2026-05-19":{a:true,b:false},"2026-05-20":{a:true,b:false},"2026-05-21":{a:true,b:false},"2026-05-22":{a:true,b:false},"2026-05-23":{a:true,b:false},"2026-05-24":{a:true,b:false},"2026-05-25":{a:false,b:true},"2026-05-26":{a:false,b:true},"2026-05-27":{a:false,b:true},"2026-05-28":{a:false,b:true},"2026-05-29":{a:false,b:true},"2026-05-30":{a:false,b:true},"2026-05-31":{a:false,b:true},"2026-06-01":{a:true,b:false},"2026-06-02":{a:true,b:false},"2026-06-03":{a:true,b:false},"2026-06-04":{a:true,b:false},"2026-06-05":{a:true,b:false},"2026-06-06":{a:true,b:false},"2026-06-07":{a:true,b:false},"2026-06-08":{a:false,b:true},"2026-06-09":{a:false,b:true},"2026-06-10":{a:false,b:true},"2026-06-11":{a:false,b:true},"2026-06-12":{a:false,b:true,n:"Lecture RMS 1/2"},"2026-06-13":{a:false,b:true},"2026-06-14":{a:false,b:true},"2026-06-15":{a:true,b:false},"2026-06-16":{a:true,b:false},"2026-06-17":{a:true,b:false},"2026-06-18":{a:true,b:false},"2026-06-19":{a:true,b:false,n:"Lecture RMS 2/2"},"2026-06-20":{a:true,b:false},"2026-06-21":{a:true,b:false},"2026-06-22":{a:false,b:true},"2026-06-23":{a:false,b:true},"2026-06-24":{a:false,b:true},"2026-06-25":{a:false,b:true},"2026-06-26":{a:false,b:true},"2026-06-27":{a:false,b:true},"2026-06-28":{a:false,b:true},"2026-06-29":{a:true,b:false},"2026-06-30":{a:true,b:false},"2026-07-01":{a:true,b:false},"2026-07-02":{a:true,b:false},"2026-07-03":{a:true,b:false},"2026-07-04":{a:true,b:false},"2026-07-05":{a:true,b:false},"2026-07-06":{a:false,b:true},"2026-07-07":{a:false,b:true},"2026-07-08":{a:false,b:true},"2026-07-09":{a:false,b:true},"2026-07-10":{a:false,b:true},"2026-07-11":{a:false,b:true},"2026-07-12":{a:false,b:true},"2026-07-13":{a:true,b:false},"2026-07-14":{a:true,b:false},"2026-07-15":{a:true,b:false},"2026-07-16":{a:true,b:false},"2026-07-17":{a:true,b:false},"2026-07-18":{a:true,b:false},"2026-07-19":{a:true,b:false},"2026-07-20":{a:false,b:true},"2026-07-21":{a:false,b:true},"2026-07-22":{a:false,b:true},"2026-07-23":{a:false,b:true},"2026-07-24":{a:false,b:true},"2026-07-25":{a:false,b:true},"2026-07-26":{a:false,b:true},"2026-07-27":{a:true,b:false},"2026-07-28":{a:true,b:false},"2026-07-29":{a:true,b:false},"2026-07-30":{a:true,b:false},"2026-07-31":{a:true,b:false},"2026-08-01":{a:true,b:false},"2026-08-02":{a:true,b:false},"2026-08-03":{a:false,b:true},"2026-08-04":{a:false,b:true},"2026-08-05":{a:false,b:true},"2026-08-06":{a:false,b:true},"2026-08-07":{a:false,b:true},"2026-08-08":{a:false,b:true},"2026-08-09":{a:false,b:true},"2026-08-10":{a:true,b:false},"2026-08-11":{a:true,b:false},"2026-08-12":{a:true,b:false},"2026-08-13":{a:true,b:false,n:"Lecture RFS 1/2"},"2026-08-14":{a:true,b:false},"2026-08-15":{a:true,b:false},"2026-08-16":{a:true,b:false},"2026-08-17":{a:false,b:true},"2026-08-18":{a:false,b:true},"2026-08-19":{a:false,b:true},"2026-08-20":{a:false,b:true,n:"Lecture RFS 2/2"},"2026-08-21":{a:false,b:true,n:"Fête Départ"},"2026-08-22":{a:false,b:true},"2026-08-23":{a:false,b:true},"2026-08-24":{a:true,b:false},"2026-08-25":{a:true,b:false},"2026-08-26":{a:true,b:false,n:"Gorée"},"2026-08-27":{a:true,b:false,n:"Warang"},"2026-08-28":{a:true,b:false,n:"Départ"}};
const FAT_PLAN={"2026-03-12":{a:false,b:true,n:"Arrivée"},"2026-03-13":{a:false,b:true,n:"Départ vers RHO"},"2026-03-14":{a:true,b:true,n:"RHO"},"2026-03-15":{a:true,b:true,n:"Rando"},"2026-03-16":{a:true,b:true,n:"Rando"},"2026-03-17":{a:true,b:true,n:"Rando"},"2026-03-18":{a:true,b:true,n:"Arrivée FATICK"},"2026-03-19":{a:true,b:false},"2026-03-20":{a:true,b:false},"2026-03-21":{a:true,b:false},"2026-03-22":{a:true,b:false},"2026-03-23":{a:false,b:true},"2026-03-24":{a:false,b:true},"2026-03-25":{a:false,b:true},"2026-03-26":{a:false,b:true},"2026-03-27":{a:false,b:true},"2026-03-28":{a:false,b:true},"2026-03-29":{a:false,b:true},"2026-03-30":{a:true,b:false},"2026-03-31":{a:true,b:false},"2026-04-01":{a:true,b:false},"2026-04-02":{a:true,b:false},"2026-04-03":{a:true,b:false},"2026-04-04":{a:true,b:false},"2026-04-05":{a:true,b:false},"2026-04-06":{a:false,b:true},"2026-04-07":{a:false,b:true},"2026-04-08":{a:false,b:true},"2026-04-09":{a:false,b:true},"2026-04-10":{a:false,b:true},"2026-04-11":{a:false,b:true},"2026-04-12":{a:false,b:true},"2026-04-13":{a:true,b:false},"2026-04-14":{a:true,b:false},"2026-04-15":{a:true,b:false},"2026-04-16":{a:true,b:false},"2026-04-17":{a:true,b:false},"2026-04-18":{a:true,b:false},"2026-04-19":{a:true,b:false},"2026-04-20":{a:false,b:true},"2026-04-21":{a:false,b:true},"2026-04-22":{a:false,b:true},"2026-04-23":{a:false,b:true},"2026-04-24":{a:false,b:true},"2026-04-25":{a:false,b:true},"2026-04-26":{a:false,b:true},"2026-04-27":{a:true,b:false},"2026-04-28":{a:true,b:false},"2026-04-29":{a:true,b:false},"2026-04-30":{a:true,b:false},"2026-05-01":{a:true,b:false},"2026-05-02":{a:true,b:false},"2026-05-03":{a:true,b:false},"2026-05-04":{a:false,b:true},"2026-05-05":{a:false,b:true},"2026-05-06":{a:false,b:true},"2026-05-07":{a:false,b:true},"2026-05-08":{a:false,b:true},"2026-05-09":{a:false,b:true,n:"Lecture RMS 1/2"},"2026-05-10":{a:false,b:true},"2026-05-11":{a:true,b:false},"2026-05-12":{a:true,b:false},"2026-05-13":{a:true,b:false},"2026-05-14":{a:true,b:false},"2026-05-15":{a:true,b:false},"2026-05-16":{a:true,b:false},"2026-05-17":{a:true,b:false},"2026-05-18":{a:false,b:true,n:"Lecture RMS 2/2"},"2026-05-19":{a:false,b:true},"2026-05-20":{a:false,b:true},"2026-05-21":{a:false,b:true},"2026-05-22":{a:false,b:true},"2026-05-23":{a:false,b:true},"2026-05-24":{a:false,b:true},"2026-05-25":{a:true,b:false},"2026-05-26":{a:true,b:false},"2026-05-27":{a:true,b:false},"2026-05-28":{a:true,b:false},"2026-05-29":{a:true,b:false},"2026-05-30":{a:true,b:false},"2026-05-31":{a:true,b:false},"2026-06-01":{a:false,b:true},"2026-06-02":{a:false,b:true},"2026-06-03":{a:false,b:true},"2026-06-04":{a:false,b:true},"2026-06-05":{a:false,b:true},"2026-06-06":{a:false,b:true},"2026-06-07":{a:false,b:true},"2026-06-08":{a:true,b:false},"2026-06-09":{a:true,b:false},"2026-06-10":{a:true,b:false},"2026-06-11":{a:true,b:false},"2026-06-12":{a:true,b:false},"2026-06-13":{a:true,b:false},"2026-06-14":{a:true,b:false},"2026-06-15":{a:false,b:true},"2026-06-16":{a:false,b:true},"2026-06-17":{a:false,b:true},"2026-06-18":{a:false,b:true},"2026-06-19":{a:false,b:true},"2026-06-20":{a:false,b:true},"2026-06-21":{a:false,b:true},"2026-06-22":{a:true,b:false},"2026-06-23":{a:true,b:false},"2026-06-24":{a:true,b:false},"2026-06-25":{a:true,b:false},"2026-06-26":{a:true,b:false},"2026-06-27":{a:true,b:false},"2026-06-28":{a:true,b:false},"2026-06-29":{a:false,b:true},"2026-06-30":{a:false,b:true},"2026-07-01":{a:false,b:true},"2026-07-02":{a:false,b:true},"2026-07-03":{a:false,b:true},"2026-07-04":{a:false,b:true},"2026-07-05":{a:false,b:true},"2026-07-06":{a:true,b:false},"2026-07-07":{a:true,b:false},"2026-07-08":{a:true,b:false},"2026-07-09":{a:true,b:false},"2026-07-10":{a:true,b:false},"2026-07-11":{a:true,b:false},"2026-07-12":{a:true,b:false},"2026-07-13":{a:false,b:true},"2026-07-14":{a:false,b:true},"2026-07-15":{a:false,b:true},"2026-07-16":{a:false,b:true},"2026-07-17":{a:false,b:true},"2026-07-18":{a:false,b:true},"2026-07-19":{a:false,b:true},"2026-07-20":{a:true,b:false},"2026-07-21":{a:true,b:false},"2026-07-22":{a:true,b:false},"2026-07-23":{a:true,b:false},"2026-07-24":{a:true,b:false},"2026-07-25":{a:true,b:false},"2026-07-26":{a:true,b:false,n:"Lecture RFS 1/2"},"2026-07-27":{a:false,b:true},"2026-07-28":{a:false,b:true},"2026-07-29":{a:false,b:true},"2026-07-30":{a:false,b:true},"2026-07-31":{a:false,b:true},"2026-08-01":{a:false,b:true,n:"Lecture RFS 2/2"},"2026-08-02":{a:false,b:true},"2026-08-03":{a:true,b:false},"2026-08-04":{a:true,b:false,n:"Fête de départ"},"2026-08-05":{a:true,b:false},"2026-08-06":{a:true,b:false},"2026-08-07":{a:true,b:false},"2026-08-08":{a:true,b:false},"2026-08-09":{a:true,b:false,n:"Gorée"},"2026-08-10":{a:false,b:true,n:"Warang"},"2026-08-11":{a:false,b:true,n:"Départ"}};
const today=new Date().toISOString().slice(0,10);
const WD=["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"];
function getWeekDates(){const n=new Date(),m=new Date(n);m.setDate(n.getDate()-((n.getDay()+6)%7));return Array.from({length:7},(_,i)=>{const d=new Date(m);d.setDate(m.getDate()+i);return d.toISOString().slice(0,10);});}
const WEEKDATES=getWeekDates();
const INIT_RAPPORTS=[{id:1,jeuneId:2,date:today,observation:"Andy a particip\u00e9 activement aux ateliers mara\u00eechers ce matin. Bonne dynamique de groupe."},{id:2,jeuneId:1,date:today,observation:"Cha\u00efna a montr\u00e9 une bonne implication lors de l\u0027atelier couture. Elle progresse bien."},{id:3,jeuneId:10,date:"2026-05-12",observation:"Imane a bien particip\u00e9 au chantier solidaire. Bonne ambiance g\u00e9n\u00e9rale."}];
const INIT_PRESENCES=(()=>{const data=[];JEUNES.forEach(j=>{WEEKDATES.forEach((date,i)=>{data.push({id:`${j.id}-${date}`,jeuneId:j.id,date,statut:i===5&&j.id===2?"Absent":"Présent"});});});return data;})();
const INIT_EV=[{id:1,jeuneId:14,date:"2026-05-10",titre:"Altercation verbale",description:"Dispute avec un pair lors du repas du soir. Tensions apais\u00e9es apr\u00e8s m\u00e9diation.",gravite:"L\u00e9ger"},{id:2,jeuneId:5,date:"2026-05-08",titre:"Refus d\u0027activit\u00e9",description:"Kais a refus\u00e9 de participer \u00e0 l\u0027atelier du matin. Entretien men\u00e9 par l\u0027\u00e9ducateur.",gravite:"Moyen"}];
const GC={Léger:{bg:"#E8F5E9",text:"#2E7D32",dot:"#4CAF50"},Moyen:{bg:"#FFF8E1",text:"#F57F17",dot:"#FFC107"},Grave:{bg:"#FFEBEE",text:"#C62828",dot:"#F44336"}};
const SC={Présent:{bg:"#E8F5E9",text:"#2E7D32",icon:"✓"},Absent:{bg:"#FFEBEE",text:"#C62828",icon:"✗"},Retard:{bg:"#FFF8E1",text:"#E65100",icon:"◷"}};
const age=dob=>{const d=new Date(dob),n=new Date();let a=n.getFullYear()-d.getFullYear();if(n.getMonth()<d.getMonth()||(n.getMonth()===d.getMonth()&&n.getDate()<d.getDate()))a--;return a;};
const fmt=s=>{if(!s)return"";const[y,m,d]=s.split("-");return`${d}/${m}/${y}`;};
const S={card:{background:C.white,borderRadius:16,padding:"18px",boxShadow:"0 2px 16px rgba(0,0,0,0.06)",border:`1px solid ${C.border}`,marginBottom:12,transition:"box-shadow 0.2s ease, transform 0.15s ease"},inp:{width:"100%",padding:"11px 14px",borderRadius:10,border:`1.5px solid ${C.border}`,fontSize:14,fontFamily:"'Nunito',sans-serif",color:C.dark,background:C.white,outline:"none",boxSizing:"border-box"},lbl:{display:"block",fontSize:11,fontWeight:700,color:C.mid,marginBottom:5,letterSpacing:"0.05em",textTransform:"uppercase"},btnP:{background:`linear-gradient(135deg,${C.gold},${C.goldDark})`,color:C.white,border:"none",borderRadius:11,padding:"11px 20px",fontWeight:700,fontSize:14,cursor:"pointer",display:"flex",alignItems:"center",gap:7,fontFamily:"'Nunito',sans-serif"},btnO:{background:`linear-gradient(135deg,${C.orange},#C06800)`,color:C.white,border:"none",borderRadius:11,padding:"10px 18px",fontWeight:700,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",gap:7,fontFamily:"'Nunito',sans-serif"},btnS:{background:C.sableLight,color:C.dark,border:`1px solid ${C.border}`,borderRadius:11,padding:"10px 18px",fontWeight:600,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",gap:7,fontFamily:"'Nunito',sans-serif"}};
const Tag=({bg,text,children})=><span style={{background:bg,color:text,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>{children}</span>;

function Login({onLogin}){
  const[email,setEmail]=useState(""),[ pw,setPw]=useState(""),[ err,setErr]=useState("");
  const handle=()=>{const u=(loadLS()?.users||USERS).find(u=>(u.email===email||u.login===email)&&u.password===pw);if(u&&u.disabled){setErr("Compte désactivé. Contactez votre chef de service.")}else{u?(setErr(""),onLogin(u)):setErr("Identifiants incorrects")}};
  return(<div style={{minHeight:"100vh",background:"linear-gradient(160deg,"+C.dark+" 0%,#2A1500 40%,#4A2800 70%,"+C.goldDark+" 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,animation:"fadeIn 0.6s ease"}}>
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

const NAV=[{id:"dashboard",label:"Tableau de bord",icon:Home},{id:"jeunes",label:"Jeunes",icon:Users},{id:"majeurs",label:"Majeurs",icon:Users},{id:"rapports",label:"Rapports journaliers",icon:FileText},{id:"presences",label:"Présences",icon:Calendar},{id:"evenements",label:"Événements",icon:AlertTriangle},{id:"rapport-hebdo",label:"Rapport hebdo",icon:BarChart2},{id:"agenda",label:"Agenda / RDV",icon:Calendar},{id:"export",label:"Export Excel",icon:Download},{id:"admin",label:"Administration",icon:Users},{id:"planning",label:"Planning",icon:Calendar},{id:"rapport-site",label:"Rapport de site",icon:FileText}];

function Sidebar({page,onNav,user,onLogout,open,onClose}){
  return(<>
    {open&&<div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:40}}/>}
    <aside style={{position:"fixed",top:0,left:open?0:-270,width:260,height:"100vh",background:"linear-gradient(180deg,"+C.dark+" 0%,#2A1500 100%)",zIndex:50,transition:"left 0.3s cubic-bezier(0.4,0,0.2,1)",display:"flex",flexDirection:"column",overflowY:"auto",boxShadow:open?"8px 0 32px rgba(0,0,0,0.3)":"none"}}>
      <div style={{padding:"22px 18px 18px",borderBottom:"1px solid rgba(255,255,255,0.09)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:11}}>
            <img src={LOGO} alt="PDSR" style={{width:40,height:40,borderRadius:9,objectFit:"contain",background:"rgba(255,255,255,0.9)",padding:3}}/>
            <div><div style={{color:C.white,fontWeight:900,fontSize:14}}>PDSR</div><div style={{color:"rgba(255,255,255,0.45)",fontSize:10,letterSpacing:"0.06em",textTransform:"uppercase"}}>Sénégal</div></div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"rgba(255,255,255,0.4)",cursor:"pointer",padding:4,display:"flex"}}><X size={17}/></button>
        </div>
      </div>
      <div style={{padding:"14px 18px",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:34,height:34,borderRadius:9,background:`${C.gold}33`,display:"flex",alignItems:"center",justifyContent:"center",color:C.gold,fontWeight:800,fontSize:12}}>{user.initials}</div>
          <div><div style={{color:C.white,fontWeight:700,fontSize:13}}>{user.name}</div><div style={{color:C.sable,fontSize:11,fontWeight:600}}>{user.role==="directeur"?"Directeur":user.role==="chef_service"?"Chef de service":`Éducateur · ${user.site}`}</div></div>
        </div>
      </div>
      <nav style={{flex:1,padding:"10px 10px"}}>
        {NAV.filter(n=>{if((n.id==="admin"||n.id==="export")&&user.role!=="directeur"&&user.role!=="chef_service"&&user.role!=="coordinateur_site")return false;if(user.isEducMajeur&&(n.id==="jeunes"||n.id==="presences"||n.id==="planning"))return false;if(n.id==="rapport-site"&&user.role!=="coordinateur_site"&&user.role!=="chef_service"&&user.role!=="directeur")return false;if(user.role==="educateur"&&!user.isEducMajeur&&n.id==="majeurs")return false;return true;}).map(item=>{const Icon=item.icon;const active=page===item.id||page.startsWith(item.id+"-");return(<button key={item.id} onClick={()=>{onNav(item.id);onClose();}} style={{width:"100%",display:"flex",alignItems:"center",gap:11,padding:"11px 14px",borderRadius:10,border:"none",cursor:"pointer",fontFamily:"inherit",fontWeight:active?700:500,fontSize:13,marginBottom:3,background:active?"linear-gradient(135deg,"+C.gold+"33,"+C.gold+"11)":"transparent",color:active?C.sable:"rgba(255,255,255,0.55)",textAlign:"left",transition:"all 0.15s ease",letterSpacing:"0.01em"}}><Icon size={17}/>{item.label}{active&&<ChevronRight size={13} style={{marginLeft:"auto"}}/>}</button>);})}
      </nav>
      <div style={{padding:"10px 10px 22px"}}><button onClick={onLogout} style={{width:"100%",display:"flex",alignItems:"center",gap:11,padding:"10px 13px",borderRadius:9,border:"none",cursor:"pointer",fontFamily:"inherit",fontWeight:600,fontSize:13,background:"rgba(244,67,54,0.12)",color:"#EF5350"}}><LogOut size={17}/>Déconnexion</button></div>
    </aside>
  </>);
}

function Topbar({title,onMenu,onBack}){
  return(<header style={{position:"sticky",top:0,zIndex:30,background:"rgba(255,255,255,0.97)",backdropFilter:"blur(12px)",WebkitBackdropFilter:"blur(12px)",borderBottom:"1px solid "+C.border,padding:"0 18px",height:58,display:"flex",alignItems:"center",gap:12,boxShadow:"0 2px 16px rgba(0,0,0,0.05)"}}>
    {onBack?<button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",display:"flex",alignItems:"center",gap:5,color:C.gold,fontWeight:700,fontSize:13,fontFamily:"inherit",padding:0}}><ChevronLeft size={19}/>Retour</button>
    :<button onClick={onMenu} style={{background:"none",border:"none",cursor:"pointer",color:C.dark,padding:4,display:"flex"}}><Menu size={21}/></button>}
    <h1 style={{fontSize:16,fontWeight:800,color:C.dark,margin:0,flex:1}}>{title}</h1>
  </header>);
}

function Dashboard({user,rapports,presences,evenements,onNav,setSel,setPage,jeunes,agenda,majeurs}){
  const isMajEduc=user.role==="educateur"&&user.isEducMajeur;
  const pool=isMajEduc?(majeurs||MAJEURS):(jeunes||JEUNES);
  const vj=user.role==="educateur"?pool.filter(j=>user.site==="Tous"||j.site===user.site):pool;
  const todayP=presences.filter(p=>p.date===today);
  const myPresents=todayP.filter(p=>p.statut==="Présent"&&vj.some(j=>j.id===p.jeuneId)).length;
  const myRapports=(rapports||[]).filter(r=>r.date===today&&vj.some(j=>j.id===r.jeuneId)).length;
  const myGraves=(evenements||[]).filter(e=>e.gravite==="Grave"&&vj.some(j=>j.id===e.jeuneId)).length;
  const jeunesNav=isMajEduc?"majeurs":"jeunes";
  return(<div style={{padding:"20px 16px",maxWidth:800,margin:"0 auto",animation:"fadeIn 0.4s ease"}}>
    <p style={{color:C.light,fontSize:12,margin:"0 0 4px",letterSpacing:"0.02em"}}>{new Date().toLocaleDateString("fr-FR",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</p>
    <h1 style={{fontSize:24,fontWeight:900,color:C.dark,margin:"0 0 22px",letterSpacing:"-0.01em"}}>Bonjour, {user.name.split(" ")[0]} 👋</h1>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:20}}>
      {[{l:isMajEduc?"Majeurs suivis":"Jeunes suivis",v:vj.length,i:"👥",c:C.gold,bg:C.goldLight,nav:jeunesNav},{l:"Présents aujourd'hui",v:`${myPresents}/${vj.length}`,i:"✔️",c:"#2E7D32",bg:"#E8F5E9",nav:isMajEduc?null:"presences"},{l:"Rapports ce jour",v:myRapports,i:"📝",c:C.orange,bg:C.orangeLight,nav:"rapports"},{l:"Incidents graves",v:myGraves,i:"⚠️",c:"#C62828",bg:"#FFEBEE",nav:"evenements"}].map((s,i)=>(
        <div key={i} onClick={()=>s.nav&&setPage(s.nav)} style={{...S.card,display:"flex",alignItems:"center",gap:14,padding:"16px",marginBottom:0,cursor:s.nav?"pointer":"default",borderLeft:"4px solid "+s.c,animation:"fadeIn 0.4s ease "+(i*0.08)+"s both",opacity:s.nav?1:0.7}}>
          <div style={{width:44,height:44,borderRadius:12,background:s.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{s.i}</div>
          <div><div style={{fontSize:26,fontWeight:900,color:s.c,lineHeight:1}}>{s.v}</div><div style={{fontSize:10,color:C.light,fontWeight:700,marginTop:3,letterSpacing:"0.03em",textTransform:"uppercase"}}>{s.l}</div></div>
        </div>
      ))}
    </div>
    {user.role!=="educateur"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:20}}>
      {[{l:"Site Fatick",n:JEUNES.filter(j=>j.site==="Fatick").length,e:"🏖️",c:C.gold},{l:"Site Djilass",n:JEUNES.filter(j=>j.site==="Djilass").length,e:"🌿",c:C.orange}].map((s,i)=>(
        <div key={i} onClick={()=>setPage("jeunes")} style={{...S.card,borderTop:"none",borderLeft:"4px solid "+s.c,padding:"16px",marginBottom:0,cursor:"pointer",animation:"fadeIn 0.4s ease 0.3s both"}}>
          <div style={{fontSize:24,marginBottom:4}}>{s.e}</div><div style={{fontWeight:900,fontSize:20,color:C.dark}}>{s.n} jeunes</div><div style={{fontSize:11,color:C.light,fontWeight:700,marginTop:2}}>{s.l}</div>
        </div>
      ))}
    </div>}
    {agenda&&agenda.length>0&&(()=>{const allJ=[...(jeunes||JEUNES),...(majeurs||MAJEURS)];const upcoming=agenda.filter(a=>{if(a.date<new Date().toISOString().slice(0,10))return false;const j=allJ.find(x=>String(x.id)===String(a.jeuneId));if(user.role==="educateur"){if(user.isEducMajeur&&j&&j.id<100)return false;if(!user.isEducMajeur&&j&&j.id>=100)return false;if(user.site!=="Tous"&&j&&j.site!==user.site)return false;}return true;}).sort((a,b)=>a.date.localeCompare(b.date)||((a.heure||"").localeCompare(b.heure||"")));const typeColors={referent:C.gold,medical:"#2E7D32",administratif:C.primary,scolaire:C.orange,autre:C.mid};return(<div onClick={()=>setPage("agenda")} style={{...S.card,marginTop:16,cursor:"pointer",borderLeft:"4px solid "+C.gold}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}><span style={{fontSize:18}}>📅</span><span style={{fontWeight:800,fontSize:15,color:C.dark}}>Prochains RDV</span><span style={{marginLeft:"auto",fontSize:11,color:C.light,fontWeight:600}}>{upcoming.length} à venir</span></div>{upcoming.slice(0,5).map(a=>{const j=allJ.find(x=>String(x.id)===String(a.jeuneId));const nom=j?(j.prenom+" "+j.nom):(a.jeuneNom||"—");const tc=typeColors[a.type]||C.mid;const [y,m,d]=(a.date||"").split("-");return(<div key={a.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 6px",borderRadius:8,marginBottom:4,background:"#f8f9fa"}}><div style={{minWidth:44,textAlign:"center"}}><div style={{fontSize:13,fontWeight:800,color:C.primary}}>{d}/{m}</div><div style={{fontSize:10,color:C.mid,fontWeight:600}}>{a.heure||""}</div></div><div style={{flex:1}}><div style={{fontSize:13,fontWeight:700,color:C.dark}}>{nom}</div><div style={{display:"inline-block",fontSize:10,fontWeight:700,color:"#fff",background:tc,borderRadius:4,padding:"1px 6px",marginTop:2}}>{a.type||"rdv"}</div></div></div>);})}{upcoming.length===0&&<div style={{fontSize:12,color:C.light,textAlign:"center",padding:10}}>Aucun RDV à venir</div>}</div>);})()}
    
  </div>);
}

function JeunesList({user,jeunes,presences,onSelect,onNav,onUpdateJeune}){
  const[q,setQ]=useState(""),[ site,setSite]=useState("Tous");
  const vj=user.role==="educateur"?JEUNES.filter(j=>user.site==="Tous"||j.site===user.site):JEUNES;
  const vis=vj.filter(j=>{const m=`${j.prenom} ${j.nom}`.toLowerCase().includes(q.toLowerCase());const s=site==="Tous"||j.site===site;return m&&s;});
  return(<div style={{padding:"18px 14px",maxWidth:800,margin:"0 auto"}}>
    <div style={{display:"flex",gap:8,marginBottom:14}}>
      <div style={{flex:1,position:"relative"}}><Search size={15} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:C.light}}/><input style={{...S.inp,paddingLeft:35}} placeholder="Rechercher..." value={q} onChange={e=>setQ(e.target.value)}/></div>
      {user.role!=="educateur"&&<select style={{...S.inp,width:"auto",paddingLeft:12}} value={site} onChange={e=>setSite(e.target.value)}><option>Tous</option><option>Fatick</option><option>Djilass</option></select>}
    </div>
    {vis.map(j=>{const tp=presences.filter(p=>p.date===today&&p.jeuneId===j.id)[0];return(<div key={j.id} style={{...S.card,cursor:"pointer",animation:"fadeIn 0.3s ease"}} onClick={()=>{onSelect(j);onNav("jeune-detail");}}>
      <div style={{display:"flex",alignItems:"center",gap:14}}>
        <div style={{width:46,height:46,borderRadius:14,background:"linear-gradient(135deg,"+C.sable+","+C.goldLight+")",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:16,color:C.goldDark,flexShrink:0,boxShadow:"0 2px 8px rgba(184,134,11,0.15)"}}>{j.prenom[0]}{j.nom?j.nom[0]:""}</div>
        <div style={{flex:1}}>
          <div style={{fontWeight:800,fontSize:15,color:C.dark}}>{j.prenom} {j.nom}</div>
          <div style={{fontSize:11,color:C.light,marginTop:1}}>{j.referentA} / {j.referentB} · <span style={{color:C.gold,fontWeight:700}}>{j.site}</span></div>
        </div>
        {tp&&<div style={{fontSize:11,fontWeight:700,padding:"4px 10px",borderRadius:8,background:SC[tp.statut]?.bg||C.sable,color:SC[tp.statut]?.text||C.mid}}>{tp.statut}</div>}
        <ChevronRight size={16} color={C.light}/>
      </div>
    </div>);})}
    {vis.length===0&&<div style={{textAlign:"center",padding:36,color:C.light,fontSize:13}}>Aucun jeune trouvé</div>}
  </div>);
}

function JeuneDetail({jeune,rapports,presences,evenements,user,onAddR,onAddE,onCP,onUpdateJeune,users}){
  const[tab,setTab]=useState("fiche");const[saved,setSaved]=useState(false);
  const jr=(rapports||[]).filter(r=>r.jeuneId===jeune.id).sort((a,b)=>b.date.localeCompare(a.date));
  const jp=presences.filter(p=>p.jeuneId===jeune.id&&WEEKDATES.includes(p.date));
  const je=(evenements||[]).filter(e=>e.jeuneId===jeune.id).sort((a,b)=>b.date.localeCompare(a.date));
  const tabs=["fiche","rapports","présences","incidents"];
  return(<div style={{padding:"14px",maxWidth:700,margin:"0 auto"}}>
    <div style={{...S.card,background:C.sable,border:"none",marginBottom:14}}>
      <div style={{display:"flex",alignItems:"center",gap:14}}>
        <div style={{width:52,height:52,borderRadius:15,background:C.white,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:20,color:C.gold}}>{jeune.prenom[0]}{jeune.nom?jeune.nom[0]:""}</div>
        <div><div style={{fontSize:20,fontWeight:900,color:C.dark}}>{jeune.prenom} {jeune.nom}</div><div style={{fontSize:12,color:C.mid,marginTop:2}}>Réf: {jeune.referentA} / {jeune.referentB}</div><div style={{fontSize:11,color:C.gold,fontWeight:700,marginTop:2}}>Site {jeune.site}</div></div>
      </div>
    </div>
    <div style={{display:"flex",gap:5,marginBottom:14,overflowX:"auto",paddingBottom:4}}>
      {tabs.map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:"7px 14px",borderRadius:20,border:`1.5px solid ${tab===t?C.gold:C.border}`,background:tab===t?C.gold:C.white,color:tab===t?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap",textTransform:"capitalize"}}>{t}</button>)}
    </div>
    {tab==="fiche"&&<div style={{...S.card}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>{[["Prénom",jeune.prenom],["Nom",jeune.nom],["Référent A",jeune.referentA],["Référent B",jeune.referentB],["Référent C",jeune.referentC||""],["Référent D",jeune.referentD||""],["Site",jeune.site],["Statut",jeune.statut]].map(([k,v])=><div key={k}><div style={{fontSize:10,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>{k}</div><div style={{fontWeight:700,color:C.dark,fontSize:13}}>{v}</div></div>)}</div></div>}
    {tab==="fiche"&&(user.role==="directeur"||user.role==="chef_service")&&<div style={{...S.card,marginTop:12}}><div style={{fontWeight:700,color:C.dark,fontSize:14,marginBottom:10}}>Dossier du jeune</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><div style={{fontSize:10,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>Référent A</div><select style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune.referentA||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,"referentA",e.target.value)}><option value="">--</option>{USERS.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div><div><div style={{fontSize:10,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>Référent B</div><select style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune.referentB||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,"referentB",e.target.value)}><option value="">--</option>{USERS.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div><div><div style={{fontSize:10,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>Référent C</div><select style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune.referentC||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,"referentC",e.target.value)}><option value="">--</option>{USERS.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div><div><div style={{fontSize:10,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>Référent D</div><select style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune.referentD||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,"referentD",e.target.value)}><option value="">--</option>{USERS.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>{[["emailASE","Email éduc. ASE"],["telASE","Tél. éduc. ASE"],["telParent1","Tél. parent/tuteur 1"],["telParent2","Tél. parent/tuteur 2"],["traitement","Traitement"],["notesDossier","Notes dossier"]].map(([field,label])=><div key={field} style={{gridColumn:field==="traitement"||field==="notesDossier"?"1/-1":"auto"}}><div style={{fontSize:10,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>{label}</div>{field==="notesDossier"||field==="traitement"?<textarea style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13,fontFamily:"inherit",minHeight:50,resize:"vertical"}} value={jeune[field]||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,field,e.target.value)}/>:<input style={{width:"100%",padding:6,border:"1px solid #ddd",borderRadius:6,fontSize:13}} value={jeune[field]||""} onChange={e=>onUpdateJeune&&onUpdateJeune(jeune.id,field,e.target.value)}/>}</div>)}</div><div style={{textAlign:"center",marginTop:16}}><button onClick={()=>{setSaved(true);setTimeout(()=>setSaved(false),2000);}} style={{padding:"10px 32px",background:saved?"#27ae60":"#2c6fbb",color:"#fff",border:"none",borderRadius:8,fontWeight:700,fontSize:14,cursor:"pointer",transition:"all 0.3s"}}>{saved?"✓ Enregistré !":"Enregistrer les modifications"}</button></div></div>}
    {tab==="rapports"&&<div>{jr.length===0?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun rapport</div>:jr.map(r=><div key={r.id} style={{...S.card}}><div style={{fontSize:11,color:C.gold,fontWeight:700,marginBottom:5}}>{fmt(r.date)}{r.author&&<span style={{fontWeight:400,fontSize:11,color:C.light,marginLeft:8}}>par {r.author}</span>}</div><p style={{margin:0,fontSize:13,color:C.dark,lineHeight:1.6}}>{r.observation}</p></div>)}<button style={{...S.btnP,marginTop:8}} onClick={()=>onAddR(jeune)}><Plus size={15}/>Nouveau rapport</button></div>}
    {tab==="présences"&&<div><div style={{display:"flex",gap:4,marginBottom:10}}>{WEEKDATES.map((date,i)=>{const p=jp.find(p2=>p2.date===date);const st=p?.statut||"Présent";const next={Présent:"Absent",Absent:"Retard",Retard:"Présent"};const sc2=SC[st]||SC.Présent;return(<div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:3}}><div style={{fontSize:9,fontWeight:700,color:C.light}}>{WD[i]}</div><button onClick={()=>onCP(jeune.id,date,next[st])} style={{width:"100%",aspectRatio:"1",borderRadius:7,background:sc2.bg,border:"none",cursor:"pointer",color:sc2.text,fontWeight:800,fontSize:13,display:"flex",alignItems:"center",justifyContent:"center"}}>{sc2.icon}</button></div>);})}
    </div></div>}
    {tab==="incidents"&&<div>{je.length===0?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun incident</div>:je.map(e=>{const gc=GC[e.gravite];return(<div key={e.id} style={{...S.card,borderLeft:`5px solid ${gc.dot}`}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}><div style={{fontWeight:800,color:C.dark}}>{e.titre}</div><span style={{...{background:gc.bg,color:gc.text,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}}>{e.gravite}</span></div><p style={{margin:0,fontSize:12,color:C.mid}}>{e.description}</p><div style={{fontSize:10,color:C.light,marginTop:4}}>{fmt(e.date)}{e.author&&" - par "+e.author}</div></div>);})}
    <button style={{...S.btnO,marginTop:8}} onClick={()=>onAddE(jeune)}><Plus size={15}/>Déclarer événement</button></div>}
  </div>);
}

function Rapports({user,rapports,onSave,onDelete,majeurs}){
  const allPool=[...JEUNES,...(majeurs||MAJEURS)];const vj=user.role==="educateur"?(user.isEducMajeur?allPool.filter(j=>(j.id>=100)&&(user.site==="Tous"||j.site===user.site)):allPool.filter(j=>(j.id<100)&&(user.site==="Tous"||j.site===user.site))):allPool;
  const[jid,setJid]=useState(vj[0]?.id||"");
  const[date,setDate]=useState(today);
  const[obs,setObs]=useState("");
  const[typeContact,setTypeContact]=useState("journee");
  const[saved,setSaved]=useState(false);
  const existingCount=rapports.filter(r=>r.jeuneId===+jid&&r.date===date).length;
  const TYPE_CONTACT_LABELS={"journee":"Journée du jeune","rdv_parents":"RDV téléphonique avec les parents","rdv_exterieur":"RDV téléphonique contact extérieur"};
  const handle=()=>{if(!obs.trim())return;onSave({jeuneId:+jid,date,observation:obs,typeContact});setSaved(true);setObs("");setTypeContact("journee");setTimeout(()=>setSaved(false),2500);};
  return(<div style={{padding:"18px 14px",maxWidth:700,margin:"0 auto"}}>
    <div style={{...S.card}}>
      <div style={{marginBottom:14}}><label style={{...S.lbl}}>Jeune</label><select style={{...S.inp}} value={jid} onChange={e=>setJid(e.target.value)}><option value="">-- Sélectionner --</option>{vj.map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom} — {j.site}</option>)}</select></div>
      <div style={{marginBottom:14}}><label style={{...S.lbl}}>Date</label><input style={{...S.inp}} type="date" value={date} onChange={e=>setDate(e.target.value)} max={today}/></div>
      <div style={{marginBottom:14}}><label style={{...S.lbl}}>Type de contact</label><select style={{...S.inp}} value={typeContact} onChange={e=>setTypeContact(e.target.value)}><option value="journee">Journée du jeune</option><option value="rdv_parents">RDV téléphonique avec les parents</option><option value="rdv_exterieur">RDV téléphonique contact extérieur (en lien avec le jeune)</option></select></div>
      {existingCount>0&&<div style={{padding:"9px 12px",borderRadius:9,background:C.sable,marginBottom:12,fontSize:12,color:C.mid,borderLeft:`3px solid ${C.gold}`}}>Un rapport existe déjà ce jour pour ce jeune. Il sera mis à jour.</div>}
      <div style={{marginBottom:14}}><label style={{...S.lbl}}>Observation</label><textarea style={{...S.inp,minHeight:110,resize:"vertical",lineHeight:1.6}} placeholder="Décrivez la journée du jeune, le contenu de l'échange..." value={obs} onChange={e=>setObs(e.target.value)}/></div>
      <button onClick={handle} style={{...S.btnP,width:"100%",justifyContent:"center",background:saved?"#4CAF50":undefined}}>{saved?<><Check size={17}/>Enregistré !</>:<><FileText size={17}/>Enregistrer le rapport</>}</button>
    </div>
    <h3 style={{fontSize:14,fontWeight:800,color:C.dark,margin:"18px 0 10px"}}>Rapports récents</h3>
    {(rapports||[]).filter(r=>vj.some(j=>j.id===r.jeuneId)).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,8).map(r=>{const j=allPool.find(j2=>j2.id===r.jeuneId)||{prenom:"?",nom:"",id:0};const tcLabels={"journee":"Journée du jeune","rdv_parents":"RDV tél. parents","rdv_exterieur":"RDV tél. contact ext."};return(<div key={r.id} style={{...S.card}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}><div style={{fontWeight:800,fontSize:13,color:C.dark}}>{j?.prenom} {j?.nom}</div><div style={{fontSize:11,color:C.gold,fontWeight:700}}>{fmt(r.date)}</div></div>{r.typeContact&&<div style={{fontSize:11,color:C.info||"#1565C0",fontWeight:600,marginBottom:4,background:"#E3F2FD",padding:"3px 8px",borderRadius:6,display:"inline-block"}}>{tcLabels[r.typeContact]||r.typeContact}</div>}<p style={{margin:0,fontSize:12,color:C.mid,lineHeight:1.5}}>{r.observation}</p>{r.author&&<div style={{fontSize:10,color:C.light,marginTop:4,fontStyle:"italic"}}>Rédigé par {r.author}{(r.horodatage||r.createdAt)?(" le "+(r.horodatage||new Date(r.createdAt).toISOString()).replace("T"," à ").slice(0,19)):""}</div>}{(user.role==="chef_service"||user.role==="directeur")&&<button onClick={(e)=>{e.stopPropagation();if(confirm("Supprimer ce rapport ?"))onDelete(r.id);}} style={{marginTop:6,padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Supprimer</button>}</div>);})}
  </div>);
}

function Presences({user,presences,onCP}){
  const vj=user.role==="educateur"?JEUNES.filter(j=>(j.id<100)&&(user.site==="Tous"||j.site===user.site)):JEUNES;
  const[site,setSite]=useState("Tous");
  const vis=site==="Tous"?vj:vj.filter(j=>j.site===site);
  const next={Présent:"Absent",Absent:"Retard",Retard:"Présent"};
  return(<div style={{padding:"18px 14px",maxWidth:900,margin:"0 auto"}}>
    {user.role!=="educateur"&&<div style={{display:"flex",gap:7,marginBottom:14}}>{["Tous","Fatick","Djilass"].map(s=><button key={s} onClick={()=>setSite(s)} style={{padding:"6px 14px",borderRadius:20,border:`1.5px solid ${site===s?C.gold:C.border}`,background:site===s?C.gold:C.white,color:site===s?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{s}</button>)}</div>}
    <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"separate",borderSpacing:"0 6px",minWidth:500}}>
      <thead><tr><th style={{textAlign:"left",fontSize:11,fontWeight:700,color:C.light,padding:"0 8px 4px",textTransform:"uppercase"}}>Jeune</th>{WEEKDATES.map((d,i)=><th key={i} style={{fontSize:10,fontWeight:700,color:C.light,textAlign:"center",padding:"0 2px 4px"}}>{WD[i]}<br/>{d.slice(8)}</th>)}<th style={{fontSize:11,fontWeight:700,color:C.light,textAlign:"center",padding:"0 8px 4px"}}>Total</th></tr></thead>
      <tbody>{vis.map(j=>{const jp=presences.filter(p=>p.jeuneId===j.id&&WEEKDATES.includes(p.date));return(<tr key={j.id}>
        <td style={{padding:"8px",background:C.white,borderRadius:"10px 0 0 10px",border:`1px solid ${C.border}`,borderRight:"none",fontWeight:700,fontSize:13,color:C.dark,whiteSpace:"nowrap"}}>{j.prenom} <span style={{color:C.light}}>{j.nom.slice(0,1)}.</span> <span style={{fontSize:10,color:C.gold}}>{j.site}</span></td>
        {WEEKDATES.map((date,i)=>{const p=jp.find(p2=>p2.date===date);const st=p?.statut||"Présent";const sc2=SC[st]||SC.Présent;return(<td key={i} style={{background:C.white,border:`1px solid ${C.border}`,borderLeft:"none",borderRight:"none",textAlign:"center",padding:"4px 2px"}}><button onClick={()=>onCP(j.id,date,next[st])} style={{width:30,height:30,borderRadius:7,background:sc2.bg,border:"none",cursor:"pointer",color:sc2.text,fontWeight:800,fontSize:13}}>{sc2.icon}</button></td>);})}
        <td style={{padding:"8px",background:C.white,borderRadius:"0 10px 10px 0",border:`1px solid ${C.border}`,borderLeft:"none",textAlign:"center",fontWeight:800,fontSize:13,color:"#2E7D32"}}>{jp.filter(p=>p.statut==="Présent").length}/7</td>
      </tr>);})}
      </tbody>
    </table></div>
  </div>);
}

function Evenements({user,evenements,onAdd,onDelete,majeurs,onUpdateAll}){
  const allPool=[...JEUNES,...(majeurs||MAJEURS)];const vj=user.role==="educateur"?(user.isEducMajeur?allPool.filter(j=>(j.id>=100)&&(user.site==="Tous"||j.site===user.site)):allPool.filter(j=>(j.id<100)&&(user.site==="Tous"||j.site===user.site))):allPool;
  const[jid,setJid]=useState(vj[0]?.id||"");
  const[titre,setTitre]=useState(""),[ desc,setDesc]=useState(""),[ grav,setGrav]=useState("Léger"),[ date2,setDate2]=useState(today),[categ,setCateg]=useState("jeune"),[typeEv,setTypeEv]=useState("incident");
  const[open2,setOpen2]=useState(false),[ saved,setSaved]=useState(false),[ fg,setFg]=useState("Tous");
  const canSeeEduc=user.role==="chef_service"||user.role==="directeur";
  const vis=(evenements||[]).filter(e=>{if(e.categorie==="educateur"&&!canSeeEduc)return false;const ok=e.categorie==="educateur"||vj.some(j=>j.id===e.jeuneId);const fok=fg==="Tous"||e.gravite===fg;return ok&&fok;});
  const handle=()=>{if(!titre.trim())return;onAdd({jeuneId:(categ==="jeune"||categ==="jeune_pro")?+jid:null,date:date2,titre,description:desc,gravite:grav,type:typeEv,categorie:categ});setSaved(true);setTitre("");setDesc("");setTimeout(()=>{setSaved(false);setOpen2(false);},2000);};
  return(<div style={{padding:"18px 14px",maxWidth:700,margin:"0 auto"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}><h2 style={{fontSize:18,fontWeight:900,color:C.dark,margin:0}}>Événements indésirables</h2><button onClick={()=>exportIncidentsXLSX(evenements,jeunes)} style={{background:C.primary,color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:13,fontWeight:600}}>Exporter CSV</button></div>
      <button style={{...S.btnO,fontSize:12,padding:"8px 14px"}} onClick={()=>setOpen2(!open2)}><Plus size={14}/>Déclarer</button>
    </div>
    {open2&&<div style={{...S.card,borderLeft:`4px solid ${C.orange}`,marginBottom:14}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}><div><label style={{...S.lbl}}>Concerne</label><select style={{...S.inp}} value={categ} onChange={e=>setCateg(e.target.value)}><option value="jeune">Un jeune</option><option value="jeune_pro">Jeune à professionnel</option><option value="educateur">Entre éducateurs</option></select></div><div><label style={{...S.lbl}}>Type</label><select style={{...S.inp}} value={typeEv} onChange={e=>setTypeEv(e.target.value)}><option value="incident">Incident</option><option value="plainte">Plainte</option><option value="reclamation">Reclamation</option></select></div></div>
      {(categ==="jeune"||categ==="jeune_pro")&&<div style={{marginBottom:10}}><label style={{...S.lbl}}>Jeune</label><select style={{...S.inp}} value={jid} onChange={e=>setJid(e.target.value)}>{vj.map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom}{j.id>=100?" (majeur)":""}</option>)}</select></div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
        <div><label style={{...S.lbl}}>Date</label><input style={{...S.inp}} type="date" value={date2} onChange={e=>setDate2(e.target.value)} max={today}/></div>
        <div><label style={{...S.lbl}}>Gravité</label><select style={{...S.inp}} value={grav} onChange={e=>setGrav(e.target.value)}><option>Léger</option><option>Moyen</option><option>Grave</option></select></div>
      </div>
      <div style={{marginBottom:10}}><label style={{...S.lbl}}>Titre</label><input style={{...S.inp}} value={titre} onChange={e=>setTitre(e.target.value)} placeholder="Ex: Altercation verbale"/></div>
      <div style={{marginBottom:10}}><label style={{...S.lbl}}>Description</label><textarea style={{...S.inp,minHeight:80,resize:"vertical"}} value={desc} onChange={e=>setDesc(e.target.value)} placeholder="Décrivez brièvement l'incident..."/></div>
      <button onClick={handle} style={{...S.btnO,width:"100%",justifyContent:"center",background:saved?"#4CAF50":undefined}}>{saved?<><Check size={17}/>Déclaré !</>:<><AlertTriangle size={17}/>Enregistrer</>}</button>
    </div>}
    <div style={{display:"flex",gap:7,marginBottom:14,flexWrap:"wrap"}}>{["Tous","Léger","Moyen","Grave"].map(g=><button key={g} onClick={()=>setFg(g)} style={{padding:"5px 13px",borderRadius:20,border:`1.5px solid ${fg===g?C.gold:C.border}`,background:fg===g?C.goldLight:C.white,color:fg===g?C.goldDark:C.mid,fontWeight:700,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>{g}</button>)}</div>
    {vis.sort((a,b)=>b.date.localeCompare(a.date)).map(e=>{const j=allPool.find(j2=>j2.id===e.jeuneId)||{prenom:"?",nom:"",id:0};const gc=GC[e.gravite];return(<div key={e.id} style={{...S.card,borderLeft:`5px solid ${gc.dot}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:7}}><div><div style={{fontWeight:800,fontSize:14,color:C.dark}}>{e.titre}</div><div style={{fontSize:11,color:C.light,marginTop:1}}>{j?.prenom} {j?.nom} · {fmt(e.date)}</div></div><span style={{background:gc.bg,color:gc.text,borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>{e.gravite}</span></div><p style={{margin:0,fontSize:12,color:C.mid,lineHeight:1.5}}>{e.description}</p>{e.author&&<div style={{fontSize:10,color:C.light,marginTop:4,fontStyle:"italic"}}>Rédigé par {e.author}{e.horodatage&&(" le "+e.horodatage.replace("T"," à "))}</div>}{e.numeroSuivi&&<div style={{fontSize:10,color:C.gold,marginTop:2,fontWeight:700}}>N° suivi: {e.numeroSuivi}</div>}{(user.role==="chef_service"||user.role==="directeur"||user.role==="coordinateur_site")&&<div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap"}}>{<button onClick={()=>{const num=prompt(e.numeroSuivi?"Modifier le n\u00b0 suivi (actuel: "+e.numeroSuivi+") :":"Num\u00e9ro de suivi \u00e0 attribuer :",e.numeroSuivi||"");if(num!==null&&num.trim()){const updated=(evenements||[]).map(x=>x.id===e.id?{...x,numeroSuivi:num.trim()}:x);if(onUpdateAll)onUpdateAll(updated);}}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid "+C.gold,background:C.goldLight,color:C.goldDark,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{e.numeroSuivi?"Modifier N\u00b0 suivi":"Attribuer N\u00b0 suivi"}</button>}{onDelete&&<button onClick={()=>{if(confirm("Supprimer cet événement ?"))onDelete(e.id);}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Supprimer</button>}</div>}</div>);})}
    {vis.length===0&&<div style={{textAlign:"center",padding:36,color:C.light,fontSize:13}}>Aucun événement</div>}
  </div>);
}

function RapportHebdo({user,rapports,presences,evenements,jeunes,majeurs,onSaveHebdo}){
  const allJeunes=[...(jeunes||JEUNES),...(majeurs||MAJEURS)];
  const mj=user.role==="educateur"?allJeunes.filter(j=>user.site==="Tous"||j.site===user.site):allJeunes;
  const[site,setSite]=useState("Djilass");
  const siteJeunes=mj.filter(j=>j.site===site);
  const[selJeune,setSelJeune]=useState("");
  const getISOWeekNum=(dt)=>{const d=new Date(Date.UTC(dt.getFullYear(),dt.getMonth(),dt.getDate()));d.setUTCDate(d.getUTCDate()+4-(d.getUTCDay()||7));const y1=new Date(Date.UTC(d.getUTCFullYear(),0,1));return Math.ceil((((d-y1)/86400000)+1)/7);};
  const SITE_START_WEEK={Djilass:14,Fatick:12};
  const calcSiteWeek=(s)=>{const now=new Date();const isoW=getISOWeekNum(now);const startW=SITE_START_WEEK[s]||1;const sw=isoW-startW+1;return String(sw>0?sw:1).padStart(2,"0");};
  const[weekNum,setWeekNum]=useState(()=>calcSiteWeek("Djilass"));
  const[groupText,setGroupText]=useState("");
  const[persoTexts,setPersoTexts]=useState({});
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
    }
  },[site]);

  const saveHebdoData=()=>{
    const saved=loadLS()||{};
    if(!saved.hebdo)saved.hebdo={};
    saved.hebdo[site+"_group"]=groupText;
    saved.hebdo.perso={...persoTexts};
    try{localStorage.setItem(LS_KEY,JSON.stringify({...saved,hebdo:saved.hebdo}));}catch(e){}
    fbSet("hebdo",saved.hebdo);
  };

  const compileWeekRapports=()=>{const getISOWeek=(dateStr)=>{const d=new Date(dateStr);d.setHours(0,0,0,0);d.setDate(d.getDate()+3-(d.getDay()+6)%7);const w1=new Date(d.getFullYear(),0,4);return String(1+Math.round(((d-w1)/86400000-3+(w1.getDay()+6)%7)/7)).padStart(2,"0")};const compiled={};let total=0;siteJeunes.forEach(j=>{const jRapports=(rapports||[]).filter(r=>{if(r.jeuneId!==j.id||!r.date)return false;return getISOWeek(r.date)===weekNum});if(jRapports.length>0){jRapports.sort((a,b)=>a.date.localeCompare(b.date));compiled[j.id]=jRapports.map(r=>{const dt=new Date(r.date);const dayName=dt.toLocaleDateString("fr-FR",{weekday:"long"});return dayName.charAt(0).toUpperCase()+dayName.slice(1)+" ("+r.date+") : "+r.observation}).join("\n\n");total+=jRapports.length}});setPersoTexts(p=>({...p,...compiled}));setTimeout(saveHebdoData,200);alert("Compilation terminée : "+Object.keys(compiled).length+" jeunes, "+total+" rapports trouvés pour la semaine "+weekNum)};
const generateDocx=async(jeune)=>{console.log("generateDocx called for",jeune.prenom);const logoB64=LOGO.split(",")[1];const logoBin=atob(logoB64);const logoBuffer=new Uint8Array(logoBin.length);for(let i=0;i<logoBin.length;i++)logoBuffer[i]=logoBin.charCodeAt(i);
    const fileName=getFileName(jeune);
    const ra=refA(jeune.id);
    const rb=refB(jeune.id);
    const refs=[ra,rb].filter(Boolean).join(" \u2013 ");
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
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:jeune.prenom+" "+jeune.nom,bold:true,size:24,font:"Arial"})]}),
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Semaine S"+weekNum,size:22,font:"Arial"})]}),
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Groupe de "+site.toUpperCase(),bold:true,size:22,font:"Arial"})]})
              ]})
            ]})
          ]})
        ]})},
        footers:{default:new Footer({children:[
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:200},children:[new TextRun({text:refs,bold:true,size:20,font:"Arial"})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"\u00c9ducateurs sp\u00e9cialis\u00e9s",italics:true,size:18,font:"Arial"})]}),
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:100},children:[new TextRun({text:"Association PDSR, 28 rue rouget de Lisle 93160 Noisy le Grand",size:16,font:"Arial"})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"associationpdsr@gmail.com / t\u00e9l : 06 24 75 34 31 - 05 17 22 59 33",size:16,font:"Arial"})]})
        ]})},
        children:[
          new Paragraph({spacing:{after:200},children:[new TextRun({text:"A l\u2019attention de Mme Eynac C\u00e9line, Mr Bossu Sylvain et Mme Souchon Sylvia",italics:true,size:22})]}),
          new Paragraph({spacing:{before:300,after:200},children:[new TextRun({text:"Cette semaine sur le groupe :",bold:true,size:24,underline:{}})]}),
          ...groupText.split("\n").map(line=>new Paragraph({spacing:{after:100},children:[new TextRun({text:line,size:22})]})),
          new Paragraph({spacing:{before:400,after:200},children:[new TextRun({text:"La semaine de "+jeune.nom+" "+jeune.prenom+" :",bold:true,size:24,underline:{}})]}),
          ...(persoTexts[jeune.id]||"").split("\n").map(line=>new Paragraph({spacing:{after:100},children:[new TextRun({text:line,size:22})]})),
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
      <div><label style={{fontWeight:600}}>Semaine : </label><input type="text" value={weekNum} onChange={e=>setWeekNum(e.target.value.replace(/[^0-9]/g,"").substring(0,2))} style={{width:60,padding:"0.4rem",borderRadius:6,border:"1px solid #ccc"}} placeholder="01"/></div>
    <button onClick={compileWeekRapports} style={{background:C.goldDark,color:"#fff",border:"none",borderRadius:6,padding:"0.4rem 1rem",cursor:"pointer",fontWeight:600,fontSize:"0.85rem",display:"flex",alignItems:"center",gap:"0.3rem"}} title="Compiler automatiquement les rapports journaliers de cette semaine"><FileText size={14}/> Compiler la semaine</button>
    </div>

    {(user.role!=="educateur"||user.role==="coordinateur_site")&&<div style={{background:"#f9f9f9",border:"1px solid #ddd",borderRadius:8,padding:"1rem",marginBottom:"1rem"}}>
      <h3 style={{fontWeight:600,marginBottom:"0.5rem",color:C.goldDark}}>Partie Groupe ({site})</h3>
      <p style={{fontSize:"0.85rem",color:"#666",marginBottom:"0.5rem"}}>Ce texte sera identique pour tous les jeunes de {site}</p>
      <textarea value={groupText} onChange={e=>setGroupText(e.target.value)} onBlur={saveHebdoData} rows={6} style={{width:"100%",padding:"0.5rem",borderRadius:6,border:"1px solid #ccc",fontFamily:"Arial",fontSize:"0.9rem"}} placeholder="Cette semaine, le groupe a..."/>
    </div>}

    <h3 style={{fontWeight:600,marginBottom:"0.5rem",color:C.goldDark}}>Parties individuelles</h3>
    {siteJeunes.map(j=><div key={j.id} style={{background:selJeune===String(j.id)?"#fff8e1":"#fff",border:"1px solid "+(selJeune===String(j.id)?C.gold:"#ddd"),borderRadius:8,padding:"0.8rem",marginBottom:"0.5rem",cursor:"pointer"}} onClick={()=>setSelJeune(String(j.id))}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontWeight:600}}>{j.prenom} {j.nom} <span style={{fontSize:"0.8rem",color:"#888"}}>({getFileName(j)})</span></span>
        <div style={{display:"flex",gap:"0.5rem"}}>
          <button onClick={e=>{e.stopPropagation();handleDownload(j)}} style={{background:C.gold,color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.8rem"}} title="T\u00e9l\u00e9charger Word"><Download size={14}/> Word</button>
          <button onClick={async e=>{e.stopPropagation();try{const{blob:b2,fileName:fn2}=await generateDocx(j);const url2=URL.createObjectURL(b2);const w=window.open(url2);if(!w){const a=document.createElement("a");a.href=url2;a.download=fn2+".docx";a.click();}setTimeout(()=>URL.revokeObjectURL(url2),10000);}catch(err){alert("Erreur: "+err.message)}}} style={{background:"#1565C0",color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.8rem"}} title="Ouvrir pour impression"><FileText size={14}/> PDF</button>
          <button onClick={e=>{e.stopPropagation();handlePrint(j)}} style={{background:C.orange,color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.8rem"}} title="Imprimer et envoyer"><Printer size={14}/> Imprimer</button>
        </div>
      </div>
      {selJeune===String(j.id)&&<div style={{marginTop:"0.5rem"}}>
        <textarea value={persoTexts[j.id]||""} onChange={e=>setPersoTexts(p=>({...p,[j.id]:e.target.value}))} onBlur={saveHebdoData} rows={5} style={{width:"100%",padding:"0.5rem",borderRadius:6,border:"1px solid #ccc",fontFamily:"Arial",fontSize:"0.9rem"}} placeholder={"La semaine de "+j.prenom+"..."}/>
        {preview&&<div style={{marginTop:"0.5rem",background:"#fff",border:"1px solid #ddd",borderRadius:6,padding:"1rem"}}>
          <div style={{textAlign:"center",fontWeight:700,fontSize:"1.1rem",borderBottom:"2px solid "+C.gold,paddingBottom:"0.5rem",marginBottom:"1rem"}}>
            <div>Association PDSR</div><div>Rapport Hebdomadaire</div><div>{j.prenom} {j.nom}</div><div>Semaine S{weekNum} - Groupe de {site.toUpperCase()}</div>
          </div>
          <p style={{fontStyle:"italic",marginBottom:"1rem"}}>A l\u2019attention de Mme Eynac C\u00e9line, Mr Bossu Sylvain et Mme Souchon Sylvia</p>
          <h4 style={{fontWeight:700,textDecoration:"underline",marginBottom:"0.5rem"}}>Cette semaine sur le groupe :</h4>
          <p style={{whiteSpace:"pre-wrap",marginBottom:"1rem"}}>{groupText||"(non renseign\u00e9)"}</p>
          <h4 style={{fontWeight:700,textDecoration:"underline",marginBottom:"0.5rem"}}>La semaine de {j.nom} {j.prenom} :</h4>
          <p style={{whiteSpace:"pre-wrap",marginBottom:"1rem"}}>{persoTexts[j.id]||"(non renseign\u00e9)"}</p>
          <div style={{borderTop:"1px solid #ddd",paddingTop:"0.5rem",textAlign:"center",fontSize:"0.85rem",color:"#666"}}>
            <div style={{fontWeight:600}}>{[refA(j.id),refB(j.id)].filter(Boolean).join(" \u2013 ")}</div>
            <div style={{fontStyle:"italic"}}>\u00c9ducateurs sp\u00e9cialis\u00e9s</div>
            <div>Association PDSR - associationpdsr@gmail.com</div>
          </div>
        </div>}
        <div style={{display:"flex",gap:"0.5rem",marginTop:"0.5rem"}}>
          <button onClick={()=>setPreview(!preview)} style={{background:preview?"#999":C.goldDark,color:"#fff",border:"none",borderRadius:6,padding:"0.3rem 0.8rem",cursor:"pointer",fontSize:"0.85rem"}}>{preview?"Fermer aper\u00e7u":"Pr\u00e9visualiser"}</button>
        </div>
      </div>}
    </div>)}

    <div style={{display:"flex",gap:"1rem",marginTop:"1rem",borderTop:"1px solid #ddd",paddingTop:"1rem"}}>
      <button onClick={handlePrintAll} disabled={sending} style={{background:C.orange,color:"#fff",border:"none",borderRadius:8,padding:"0.6rem 1.5rem",cursor:"pointer",fontWeight:600,opacity:sending?0.6:1}}><Printer size={16}/> {sending?"G\u00e9n\u00e9ration...":"Imprimer tous les rapports ("+site+")"}</button>
    </div>
    {sent&&<p style={{color:"green",marginTop:"0.5rem"}}>Rapports g\u00e9n\u00e9r\u00e9s avec succ\u00e8s !</p>}
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
  return(<div>{site==="Tous"&&<div style={{display:"flex",gap:8,marginBottom:12}}><button onClick={()=>setSelSite("Fatick")} style={{padding:"6px 16px",borderRadius:8,border:"none",background:selSite==="Fatick"?C.gold:"#eee",color:selSite==="Fatick"?"#fff":"#333",fontWeight:600,cursor:"pointer"}}>Fatick</button><button onClick={()=>setSelSite("Djilass")} style={{padding:"6px 16px",borderRadius:8,border:"none",background:selSite==="Djilass"?C.gold:"#eee",color:selSite==="Djilass"?"#fff":"#333",fontWeight:600,cursor:"pointer"}}>Djilass</button></div>}
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}><button onClick={prev} style={{background:C.gold,color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:15,fontWeight:700}}>◀</button><h2 style={{fontSize:18,fontWeight:900,color:C.dark,margin:0}}>{MN[m]} {y} — Planning {selSite}</h2><button onClick={next} style={{background:C.gold,color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",cursor:"pointer",fontSize:15,fontWeight:700}}>▶</button></div>
    {canEdit&&<div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
      <div style={{fontSize:11,color:C.light,padding:"6px 0"}}>Cliquez sur un jour pour modifier. Raccourcis semaine :</div>
    </div>}
    <div style={{display:"flex",gap:4,marginBottom:8,fontSize:10,color:C.light}}>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:C.sableLight,border:"1px solid #ddd",display:"inline-block"}}/> Aucune</span>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:"#fef3c7",border:"1px solid #ddd",display:"inline-block"}}/> Éq. A</span>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:"#dbeafe",border:"1px solid #ddd",display:"inline-block"}}/> Éq. B</span>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:"#e0e7ff",border:"1px solid #ddd",display:"inline-block"}}/> A+B</span>
      <span style={{display:"inline-flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:"#FFF3E0",border:"1px solid #ddd",display:"inline-block"}}/> Vacances</span>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4}}>{WD.map(w=><div key={w} style={{textAlign:"center",fontWeight:700,fontSize:12,color:C.dark,padding:4}}>{w}</div>)}{Array.from({length:first.getDay()}).map((_,i)=><div key={"e"+i}/>)}{days.map(day=>{const bg=day.v?"#FFF3E0":day.a&&day.b?"#e0e7ff":day.a?"#fef3c7":day.b?"#dbeafe":"#f3f4f6";const today=day.key===new Date().toISOString().slice(0,10);const isEdit=editDay===day.key;return(<div key={day.key} onClick={()=>{if(!canEdit)return;if(editDay===day.key){setEditDay(null);}else{setEditDay(day.key);setEditNote(day.n);}}} style={{background:bg,borderRadius:8,padding:6,minHeight:70,border:today?"2px solid "+C.gold:isEdit?"2px solid "+C.accent:"1px solid #e5e7eb",position:"relative",cursor:canEdit?"pointer":"default",transition:"all 0.15s"}}>
      <div style={{fontWeight:700,fontSize:13,color:C.dark}}>{day.d}</div>
      {day.a&&<div style={{fontSize:10,color:C.gold,fontWeight:600}}>Éq. A</div>}
      {day.b&&<div style={{fontSize:10,color:"#2563eb",fontWeight:600}}>Éq. B</div>}
      {day.v?<div style={{fontSize:9,color:"#E65100",fontWeight:700}}>VACANCES</div>:null}
      {day.n&&!isEdit&&<div style={{fontSize:9,color:"#6b7280",marginTop:2}}>{day.n}</div>}
      {isEdit&&canEdit&&<div onClick={e=>e.stopPropagation()} style={{position:"absolute",top:0,left:0,right:0,bottom:0,background:"rgba(255,255,255,0.97)",borderRadius:8,padding:6,zIndex:10,display:"flex",flexDirection:"column",gap:3}}>
        <div style={{fontSize:11,fontWeight:700,color:C.dark}}>Jour {day.d}</div>
        <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
          <button onClick={()=>toggleDay(day.key,"a")} style={{fontSize:9,padding:"2px 6px",borderRadius:4,border:day.a?"1px solid "+C.gold:"1px solid #ccc",background:day.a?"#fef3c7":"#fff",color:C.dark,cursor:"pointer",fontWeight:600}}>A{day.a?" ✓":""}</button>
          <button onClick={()=>toggleDay(day.key,"b")} style={{fontSize:9,padding:"2px 6px",borderRadius:4,border:day.b?"1px solid #2563eb":"1px solid #ccc",background:day.b?"#dbeafe":"#fff",color:C.dark,cursor:"pointer",fontWeight:600}}>B{day.b?" ✓":""}</button>
          <button onClick={()=>toggleVac(day.key)} style={{fontSize:9,padding:"2px 6px",borderRadius:4,border:day.v?"1px solid #E65100":"1px solid #ccc",background:day.v?"#FFF3E0":"#fff",color:C.dark,cursor:"pointer",fontWeight:600}}>Vac{day.v?" ✓":""}</button>
        </div>
        <input value={editNote} onChange={e=>setEditNote(e.target.value)} onClick={e=>e.stopPropagation()} placeholder="Note..." style={{fontSize:9,padding:"2px 4px",border:"1px solid #ccc",borderRadius:3,width:"100%"}}/>
        <div style={{display:"flex",gap:3}}>
          <button onClick={()=>saveNote(day.key)} style={{fontSize:9,padding:"2px 6px",borderRadius:3,background:C.gold,color:"#fff",border:"none",cursor:"pointer",fontWeight:600,flex:1}}>OK</button>
          <button onClick={()=>{const mon=getWeekMonday(day.key);setFullWeek(mon,true,false);}} style={{fontSize:8,padding:"2px 4px",borderRadius:3,background:"#fef3c7",color:C.dark,border:"1px solid "+C.gold,cursor:"pointer",fontWeight:600}} title="Sem. A">Sem A</button>
          <button onClick={()=>{const mon=getWeekMonday(day.key);setFullWeek(mon,false,true);}} style={{fontSize:8,padding:"2px 4px",borderRadius:3,background:"#dbeafe",color:C.dark,border:"1px solid #2563eb",cursor:"pointer",fontWeight:600}} title="Sem. B">Sem B</button>
        </div>
      </div>}
    </div>);})}</div>
  </div>);
}

function exportIncidentsXLSX(evenements,jeunes){const rows=[["Date","Jeune","Titre","Description","Gravité","Horodatage","N° Suivi","Catégorie"]];(evenements||[]).forEach(ev=>{const j=jeunes.find(j2=>j2.id===ev.jeuneId);rows.push([ev.date,j?(j.prenom+" "+(j.nom||"")):("ID:"+ev.jeuneId),ev.titre,ev.description,ev.gravite||"normal"]);});const bom="﻿";const csv=rows.map(r=>r.map(c=>'"'+String(c).replace(/"/g,'""')+'"').join(";")).join("\n");const blob=new Blob([bom+csv],{type:"text/csv;charset=utf-8"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="incidents_pdsr_"+new Date().toISOString().slice(0,10)+".csv";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}

function Admin({users,jeunes,onUpdateUsers,onUpdateJeunes,loginLogs,appMajeurs,onUpdateMajeurs}){
  const[tab,setTab]=useState("educs");
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
  const genLogin=(nom,prenom)=>(nom+prenom).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z]/g,"");
  const addEduc=()=>{if(!newPrenom.trim()||!newNom.trim())return;const login=genLogin(newNom,newPrenom);const id=Math.max(...users.map(u=>u.id))+1;onUpdateUsers([...users,{id,login,password:"pdsr2026",role:"educateur",name:newPrenom,site:newSite,type:newType,section:newSection,isEducMajeur:newSection==="majeurs",initials:newPrenom.substring(0,2).toUpperCase(),assignedIds:[]}]);setNewPrenom("");setNewNom("");setNewType("jour");setNewSection("mineurs");};
  const removeEduc=(id)=>{if(!confirm("Supprimer cet éducateur ?"))return;onUpdateUsers(users.filter(u=>u.id!==id));const updated=jeunes.map(j=>j.educateurId===id?{...j,educateurId:null}:j);onUpdateJeunes(updated);};
  const toggleEduc=(id)=>{onUpdateUsers(users.map(u=>u.id===id?{...u,disabled:!u.disabled}:u))};
  const addJeune=()=>{if(!newPrenom.trim())return;const id=Math.max(...jeunes.map(j=>j.id),0)+1;onUpdateJeunes([...jeunes,{id,prenom:newPrenom,nom:newNom,site:newSite,educateurId:null,referentA:"",referentB:"",referentC:"",referentD:"",statut:"actif",telParent1:newTelP1,telJeune:newTelJ,emailASE:newEmailASE,dateDebut:newDateD,dateFin:newDateF}]);setNewPrenom("");setNewNom("");setNewTelP1("");setNewTelJ("");setNewEmailASE("");setNewDateD("");setNewDateF("");setShowAddJeune(false);};
  const assignJeune=(jeuneId,educName)=>{onUpdateJeunes(jeunes.map(j=>j.id===jeuneId?{...j,referentA:educName||""}:j));const educ=users.find(u=>u.name===educName);if(educ&&!educ.assignedIds?.includes(jeuneId)){onUpdateUsers(users.map(u=>u.name===educName?{...u,assignedIds:[...(u.assignedIds||[]),jeuneId]}:u.assignedIds?.includes(jeuneId)?{...u,assignedIds:u.assignedIds.filter(i=>i!==jeuneId)}:u));}else if(!educName){onUpdateUsers(users.map(u=>u.assignedIds?.includes(jeuneId)?{...u,assignedIds:u.assignedIds.filter(i=>i!==jeuneId)}:u));}};
  const removeJeune=(id)=>{if(!confirm("Supprimer ce jeune ?"))return;onUpdateJeunes((jeunes||[]).filter(j=>j.id!==id));onUpdateUsers(users.map(u=>u.assignedIds?{...u,assignedIds:u.assignedIds.filter(i=>i!==id)}:u));};
  return(<div style={{padding:"18px 14px",maxWidth:800,margin:"0 auto"}}>
    <h2 style={{fontSize:18,fontWeight:900,color:C.dark,margin:"0 0 14px"}}>Administration</h2>
    <div style={{display:"flex",gap:7,marginBottom:16}}>{[{k:"educs",l:"Éducateurs"},{k:"jeunes",l:"Jeunes"},{k:"majeurs",l:"Majeurs"},{k:"creds",l:"Identifiants"},{k:"logs",l:"Logs connexion"}].map(t=><button key={t.k} onClick={()=>setTab(t.k)} style={{padding:"7px 16px",borderRadius:20,border:`1.5px solid ${tab===t.k?C.gold:C.border}`,background:tab===t.k?C.gold:C.white,color:tab===t.k?C.white:C.mid,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{t.l}</button>)}</div>
    {tab==="educs"&&<div>
      <div style={{...S.card,borderLeft:`4px solid ${C.gold}`,marginBottom:14}}>
        <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 10px",color:C.dark}}>Ajouter un éducateur</h3>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <div><label style={{...S.lbl}}>Prénom</label><input style={{...S.inp}} value={newPrenom} onChange={e=>setNewPrenom(e.target.value)} placeholder="Prénom"/></div>
          <div><label style={{...S.lbl}}>Nom</label><input style={{...S.inp}} value={newNom} onChange={e=>setNewNom(e.target.value)} placeholder="Nom"/></div>
        </div>
        <div style={{marginBottom:10}}><label style={{...S.lbl}}>Site</label><select style={{...S.inp}} value={newSite} onChange={e=>setNewSite(e.target.value)}><option>Fatick</option><option>Djilass</option></select></div>
 <div><div style={{fontSize:11,fontWeight:700,color:C.light,marginBottom:4}}>Type</div><select style={{...S.input}} value={newType} onChange={e=>setNewType(e.target.value)}><option value="jour">Jour</option><option value="nuit">Nuit</option></select></div>
 <div><div style={{fontSize:11,fontWeight:700,color:C.light,marginBottom:4}}>Section</div><select style={{...S.input}} value={newSection} onChange={e=>setNewSection(e.target.value)}><option value="mineurs">Mineurs</option><option value="majeurs">Majeurs</option></select></div>
        {newPrenom&&newNom&&<div style={{fontSize:11,color:C.mid,marginBottom:8}}>Identifiant généré : <strong>{genLogin(newNom,newPrenom)}</strong> / Mot de passe : <strong>pdsr2026</strong></div>}
        <button onClick={addEduc} style={{...S.btnP,width:"100%",justifyContent:"center"}}><Plus size={14}/>Ajouter</button>
      </div>
      {educs.map(u=><div key={u.id} style={{...S.card,marginBottom:8,opacity:u.disabled?0.6:1,borderLeft:u.disabled?"3px solid #C62828":"3px solid transparent"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{fontWeight:800,fontSize:14,color:C.dark}}>{u.name}</div><div style={{fontSize:11,color:C.light}}>{u.site} · {u.login} · {u.type==="nuit"?"Nuit":"Jour"} · Éq. {u.equipe||"?"} · {u.section==="majeurs"?"Majeurs":"Mineurs"} · {jeunes.filter(j=>j.referentA===u.name||j.referentB===u.name||j.referentC===u.name||j.referentD===u.name).length} jeunes</div></div>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:5,marginTop:8}}>
          <button onClick={()=>toggleEduc(u.id)} style={{padding:"4px 10px",borderRadius:6,border:u.disabled?"1px solid #2E7D32":"1px solid #E65100",background:u.disabled?"#E8F5E9":"#FFF3E0",color:u.disabled?"#2E7D32":"#E65100",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{u.disabled?"Activer":"Désactiver"}</button>
          <button onClick={()=>{const up=users.map(x=>x.id===u.id?{...x,isEducMajeur:!x.isEducMajeur,section:x.isEducMajeur?"mineurs":"majeurs"}:x);onUpdateUsers(up);}} style={{padding:"4px 10px",borderRadius:6,border:u.isEducMajeur?"1px solid #1565C0":"1px solid #9E9E9E",background:u.isEducMajeur?"#E3F2FD":"#F5F5F5",color:u.isEducMajeur?"#1565C0":"#757575",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{u.isEducMajeur?"Éduc Majeur":"Standard"}</button>
          <button onClick={()=>{const newEq=u.equipe==="A"?"B":"A";const up=users.map(x=>x.id===u.id?{...x,equipe:newEq}:x);onUpdateUsers(up);}} style={{padding:"4px 10px",borderRadius:6,border:u.equipe==="A"?"1px solid #2E7D32":"1px solid #1565C0",background:u.equipe==="A"?"#E8F5E9":"#E3F2FD",color:u.equipe==="A"?"#2E7D32":"#1565C0",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Éq. {u.equipe||"?"}</button>
          <button onClick={()=>{const up=users.map(x=>x.id===u.id?{...x,role:x.role==="coordinateur_site"?"educateur":"coordinateur_site"}:x);onUpdateUsers(up);}} style={{padding:"4px 10px",borderRadius:6,border:u.role==="coordinateur_site"?"1px solid #6A1B9A":"1px solid #9E9E9E",background:u.role==="coordinateur_site"?"#F3E5F5":"#F5F5F5",color:u.role==="coordinateur_site"?"#6A1B9A":"#757575",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{u.role==="coordinateur_site"?"Coordinateur":"Nommer Coord."}</button>
          <button onClick={()=>removeEduc(u.id)} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>
        </div>
      </div>)}
    </div>}
    {tab==="jeunes"&&<div>
      <div style={{...S.card,borderLeft:`4px solid ${C.gold}`,marginBottom:14}}>
        {!showAddJeune&&<button onClick={()=>setShowAddJeune(true)} style={{...S.btnP,marginBottom:12}}><Plus size={14}/>Ajouter un jeune</button>}
        {showAddJeune&&<div style={{...S.card,borderLeft:`4px solid ${C.gold}`,marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><h3 style={{fontSize:13,fontWeight:800,margin:0,color:C.dark}}>Ajouter un jeune</h3><button onClick={()=>setShowAddJeune(false)} style={{background:"none",border:"none",cursor:"pointer",color:C.mid}}><X size={16}/></button></div>
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
      <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:12}}>
          <label style={{...S.btnP,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6}}><Download size={14}/>Importer Excel<input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={async(e)=>{const f=e.target.files[0];if(!f)return;const XLSX=await import("https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs");const data=await f.arrayBuffer();const wb=XLSX.read(data);const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws);const maxId=Math.max(...jeunes.map(j=>j.id),0);const newJ=rows.map((r,i)=>({id:maxId+1+i,prenom:r["Prénom"]||r.prenom||"",nom:r["Nom"]||r.nom||"",site:r["Site"]||r.site||"Fatick",referentA:"",referentB:"",referentC:"",referentD:"",statut:"actif",telParent1:r["Tél parent"]||r.telParent||"",telJeune:r["Tél jeune"]||r.telJeune||"",emailASE:r["Email ASE"]||r.emailASE||"",dateDebut:r["Date début"]||r.dateDebut||"",dateFin:r["Date fin"]||r.dateFin||""})).filter(j=>j.prenom);if(newJ.length===0){alert("Aucun jeune trouvé. Colonnes attendues: Prénom, Nom, Site, Tél parent, Tél jeune, Email ASE, Date début, Date fin");return;}if(confirm("Importer "+newJ.length+" jeune(s) ?"))onUpdateJeunes([...jeunes,...newJ]);e.target.value="";}}/></label>
          <span style={{fontSize:10,color:C.mid}}>Colonnes: Prénom, Nom, Site, Tél parent, Tél jeune, Email ASE, Date début, Date fin</span>
        </div>
            {jeunes.map(j=>{return(<div key={j.id} style={{...S.card,marginBottom:8}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
          <div><div style={{fontWeight:800,fontSize:14,color:C.dark}}>{j.prenom} {j.nom}</div><div style={{fontSize:11,color:C.light}}>{j.site} · Réf: {[j.referentA,j.referentB,j.referentC,j.referentD].filter(Boolean).join(", ")||"Non assigné"}</div></div>
          <button onClick={()=>removeJeune(j.id)} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          {["referentA","referentB","referentC","referentD"].map((rf,i)=><div key={rf} style={{display:"flex",gap:4,alignItems:"center",marginBottom:2}}><span style={{fontSize:10,fontWeight:700,color:C.mid,minWidth:32}}>Réf {String.fromCharCode(65+i)}</span><select style={{...S.inp,flex:1,padding:"3px 6px",fontSize:11}} value={j[rf]||""} onChange={e=>onUpdateJeunes(jeunes.map(x=>x.id===j.id?{...x,[rf]:e.target.value}:x))}><option value="">--</option>{educs.filter(u=>u.site===j.site).map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select></div>)}
        </div>
      </div>);})}
    </div>}

    {tab==="majeurs"&&<div>
        <div style={{fontWeight:700,fontSize:15,marginBottom:12}}>Gestion des Majeurs</div>
        {!showAddMajeur&&<button onClick={()=>setShowAddMajeur(true)} style={{...S.btnP,marginBottom:12}}><Plus size={14}/>Ajouter un majeur</button>}
        {showAddMajeur&&<div style={{...S.card,borderLeft:`4px solid ${C.gold}`,marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><h3 style={{fontSize:13,fontWeight:800,margin:0,color:C.dark}}>Ajouter un majeur</h3><button onClick={()=>setShowAddMajeur(false)} style={{background:"none",border:"none",cursor:"pointer",color:C.mid}}><X size={16}/></button></div>
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
        <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:12}}>
          <label style={{...S.btnP,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6}}><Download size={14}/>Importer Excel<input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={async(e)=>{const f=e.target.files[0];if(!f)return;const XLSX=await import("https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs");const data=await f.arrayBuffer();const wb=XLSX.read(data);const ws=wb.Sheets[wb.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(ws);const cur=appMajeurs||MAJEURS;const maxId=Math.max(...cur.map(j=>j.id),99);const newM=rows.map((r,i)=>({id:maxId+1+i,prenom:r["Prénom"]||r.prenom||"",nom:r["Nom"]||r.nom||"",site:r["Site"]||r.site||"Fatick",referentA:"",referentB:"",referentC:"",referentD:"",statut:"actif",telParent1:r["Tél parent"]||r.telParent||"",telJeune:r["Tél jeune"]||r.telJeune||"",emailASE:r["Email ASE"]||r.emailASE||"",dateDebut:r["Date début"]||r.dateDebut||"",dateFin:r["Date fin"]||r.dateFin||""})).filter(j=>j.prenom);if(newM.length===0){alert("Aucun majeur trouvé.");return;}if(confirm("Importer "+newM.length+" majeur(s) ?"))onUpdateMajeurs(null,null,null,[...cur,...newM]);e.target.value="";}}/></label>
          <span style={{fontSize:10,color:C.mid}}>Colonnes: Prénom, Nom, Site, Tél parent, Tél jeune, Email ASE, Date début, Date fin</span>
        </div>
        {(appMajeurs||MAJEURS).map(m=><div key={m.id} style={{...S.card,marginBottom:8}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{fontWeight:700,color:C.dark}}>{m.prenom} {m.nom||""}</div><div style={{fontSize:11,color:C.light}}>{m.site} | Réf: {[m.referentA,m.referentB,m.referentC,m.referentD].filter(Boolean).join(", ")||"Aucun"}</div></div>
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              {["referentA","referentB","referentC","referentD"].map((rf,i)=><select key={rf} value={m[rf]||""} onChange={e=>{if(onUpdateMajeurs)onUpdateMajeurs(m.id,rf,e.target.value);}} style={{...S.input,width:100,fontSize:11}}><option value="">{"Réf "+(i+1)}</option>{users.filter(u=>u.role==="educateur").map(u=><option key={u.id} value={u.name}>{u.name}</option>)}</select>)}
              <button onClick={()=>{if(confirm("Supprimer ce majeur ?"))onUpdateMajeurs(null,null,null,(appMajeurs||MAJEURS).filter(x=>x.id!==m.id));}} style={{padding:"4px 10px",borderRadius:6,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>×</button>
            </div>
          </div>
        </div>)}
      </div>}  {tab==="creds"&&<div>
      <h3 style={{fontSize:13,fontWeight:800,margin:"0 0 14px",color:C.dark}}>Identifiants de connexion</h3>
      {users.map(u=><div key={u.id} style={{...S.card,marginBottom:6,padding:"10px 14px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{fontWeight:700,fontSize:13,color:C.dark}}>{u.name}</div><div style={{fontSize:11,color:C.light}}>{u.role==="educateur"?"Éducateur · "+u.site:u.role==="coordinateur_site"?"Coordinateur · "+u.site:u.role==="chef_service"?"Chef de service":"Directeur"}</div></div>
          <div style={{textAlign:"right"}}><div style={{fontSize:12,fontWeight:700,color:C.gold}}>{u.email||u.login}</div><div style={{fontSize:11,color:C.mid}}>mdp: {u.password}</div></div>
        </div>
      </div>)}
    </div>}

    {tab==="logs"&&<div>
      <div style={{fontWeight:700,fontSize:15,marginBottom:12}}>Registre des connexions</div>
      <div style={{...S.card}}>
        {(loginLogs||[]).length===0?<div style={{textAlign:"center",color:C.light,padding:20}}>Aucune connexion enregistrée</div>:
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:11,color:C.light}}>Date/Heure</th><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:11,color:C.light}}>Utilisateur</th><th style={{textAlign:"left",padding:8,borderBottom:"2px solid #ddd",fontSize:11,color:C.light}}>Rôle</th></tr></thead>
          <tbody>{(loginLogs||[]).slice(0,100).map(l=><tr key={l.id}><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12}}>{new Date(l.date).toLocaleString("fr-FR")}</td><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12,fontWeight:600}}>{l.user}</td><td style={{padding:8,borderBottom:"1px solid #eee",fontSize:12}}>{l.role}</td></tr>)}</tbody>
        </table>}
      </div>
    </div>}
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
 const addRdv=()=>{if(!rdvJeune||!rdvDate)return alert("Jeune et date requis");if(editId){setAgenda(p=>p.map(a=>a.id===editId?{...a,jeuneId:rdvJeune,jeuneNom:(allJ.find(j=>String(j.id)===String(rdvJeune))||{}).prenom||"",date:rdvDate,heure:rdvHeure,type:rdvType,description:rdvDesc,with:rdvWith}:a));setEditId(null);}else{setAgenda(p=>[...p,{id:Date.now(),jeuneId:rdvJeune,jeuneNom:(allJ.find(j=>String(j.id)===String(rdvJeune))||{}).prenom||"",date:rdvDate,heure:rdvHeure,type:rdvType,description:rdvDesc,with:rdvWith,createdBy:user.name}]);}setShowForm(false);setRdvDesc("");setRdvWith("");};
 const delRdv=id=>setAgenda(p=>p.filter(r=>r.id!==id));
 const getJeuneName=id=>{const j=allJ.find(j2=>String(j2.id)===String(id));return j?(j.prenom+" "+(j.nom||"")):"?";};
 const typeColors={educateur:"#3498db",referent:"#2ecc71",medical:"#e74c3c",juridique:"#9b59b6",autre:"#f39c12"};
 return(<div>
  <div style={{...S.card,marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontWeight:700,fontSize:16,color:C.dark}}>Agenda / Rendez-vous</div><div style={{fontSize:12,color:C.light}}>{(agenda||[]).length} RDV enregistrés</div></div><button onClick={()=>{setEditId(null);setRdvJeune("");setRdvDate(new Date().toISOString().slice(0,10));setRdvHeure("09:00");setRdvType("educateur");setRdvDesc("");setRdvWith("");setShowForm(!showForm);}} style={{...S.btn}}><Plus size={14}/> Nouveau RDV</button></div>
  {showForm&&<div style={{...S.card,marginBottom:12}}>
   <div style={{fontWeight:700,marginBottom:8}}>{editId?"Modifier le rendez-vous":"Nouveau rendez-vous"}</div>
   <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
    <div><div style={{fontSize:11,fontWeight:700,color:C.light,marginBottom:4}}>Jeune</div><select value={rdvJeune} onChange={e=>setRdvJeune(e.target.value)} style={{...S.input}}><option value="">Choisir...</option>{allJ.filter(j=>{if(user.role==="educateur"){if(user.isEducMajeur&&j.id<100)return false;if(!user.isEducMajeur&&j.id>=100)return false;if(user.site!=="Tous"&&j.site!==user.site)return false;}return true;}).map(j=><option key={j.id} value={j.id}>{j.prenom} {j.nom||""}</option>)}</select></div>
    <div><div style={{fontSize:11,fontWeight:700,color:C.light,marginBottom:4}}>Type</div><select value={rdvType} onChange={e=>setRdvType(e.target.value)} style={{...S.input}}><option value="educateur">Éducateur</option><option value="referent">Référent ASE</option><option value="medical">Médical</option><option value="juridique">Juridique</option><option value="autre">Autre</option></select></div>
    <div><div style={{fontSize:11,fontWeight:700,color:C.light,marginBottom:4}}>Date</div><input type="date" value={rdvDate} onChange={e=>setRdvDate(e.target.value)} style={{...S.input}}/></div>
    <div><div style={{fontSize:11,fontWeight:700,color:C.light,marginBottom:4}}>Heure</div><input type="time" value={rdvHeure} onChange={e=>setRdvHeure(e.target.value)} style={{...S.input}}/></div>
   </div>
   <div style={{marginTop:8}}><div style={{fontSize:11,fontWeight:700,color:C.light,marginBottom:4}}>Avec qui / Détails</div><input value={rdvWith} onChange={e=>setRdvWith(e.target.value)} placeholder="Nom du professionnel..." style={{...S.input,marginBottom:6}}/></div>
   <div><div style={{fontSize:11,fontWeight:700,color:C.light,marginBottom:4}}>Description</div><textarea value={rdvDesc} onChange={e=>setRdvDesc(e.target.value)} placeholder="Détails du RDV..." rows={2} style={{...S.input,marginBottom:6}}/></div>
   <div style={{display:"flex",gap:6}}><button onClick={addRdv} style={{...S.btn}}>Valider</button><button onClick={()=>{setShowForm(false);setEditId(null);}} style={{...S.btnO}}>Annuler</button></div>
  </div>}
  {sortedAgenda.length===0?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun rendez-vous</div>:sortedAgenda.map(r=><div key={r.id} style={{...S.card,marginBottom:8,borderLeft:"4px solid "+(typeColors[r.type]||"#ccc")}}>
   <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
    <div><div style={{fontWeight:700,color:C.dark}}>{getJeuneName(r.jeuneId)}</div><div style={{fontSize:12,color:C.light}}>{r.date} à {r.heure} — <span style={{color:typeColors[r.type]||"#999",fontWeight:600}}>{r.type}</span>{r.with?" — "+r.with:""}</div>{r.description&&<div style={{fontSize:12,color:C.mid,marginTop:4}}>{r.description}</div>}{(user.role==="directeur"||user.role==="chef_service")&&<div style={{marginTop:6,borderTop:"1px solid #eee",paddingTop:6}}><div style={{fontSize:11,fontWeight:700,color:C.primary,marginBottom:3}}>CR du RDV</div><textarea value={r.cr||""} onChange={e=>{setAgenda(p=>p.map(a=>a.id===r.id?{...a,cr:e.target.value}:a));}} placeholder="Saisir le compte-rendu..." rows={2} style={{...S.input,fontSize:11,width:"100%"}}/></div>}</div>
    <div style={{display:"flex",flexDirection:"column",gap:4}}>{user.role==="chef_service"&&<button onClick={()=>delRdv(r.id)} style={{background:"none",border:"none",color:"#e74c3c",cursor:"pointer",fontSize:14}} title="Supprimer">🗑️</button>}<button onClick={()=>{setEditId(r.id);setRdvJeune(String(r.jeuneId));setRdvDate(r.date);setRdvHeure(r.heure||"09:00");setRdvType(r.type||"educateur");setRdvDesc(r.description||"");setRdvWith(r.with||"");setShowForm(true);}} style={{background:"none",border:"none",color:C.primary,cursor:"pointer",fontSize:14}} title="Modifier">✏️</button></div>
   </div>
  </div>)}
 </div>);
}

function MajeurDetail({majeur,rapports,presences,evenements,user,onBack,onAddR,onAddE,onCP,users,addR,addE}){
 const[tab,setTab]=useState("fiche");
 const[showFormR,setShowFormR]=useState(false);
 const[showFormE,setShowFormE]=useState(false);
 const[rDate,setRDate]=useState(new Date().toISOString().slice(0,10));
 const[rObs,setRObs]=useState("");
 const[eDate,setEDate]=useState(new Date().toISOString().slice(0,10));
 const[eTitre,setETitre]=useState("");
 const[eDesc,setEDesc]=useState("");
 const[eGrav,setEGrav]=useState("normal");
 const mr=(rapports||[]).filter(r=>r.jeuneId===majeur.id).sort((a,b)=>b.date.localeCompare(a.date));
 const me=(evenements||[]).filter(e=>e.jeuneId===majeur.id).sort((a,b)=>b.date.localeCompare(a.date));
 const mp=Object.entries(presences||{}).filter(([k])=>k.startsWith("2")).reduce((acc,[date,v])=>{if(v[majeur.id])acc[date]=v[majeur.id];return acc;},{});
 const tabs=["fiche","rapports","incidents","presences"];
 const submitR=()=>{if(!rObs.trim())return alert("Observation requise");addR&&addR({jeuneId:majeur.id,date:rDate,observation:rObs.trim()});setRObs("");setShowFormR(false);};
 const submitE=()=>{if(!eTitre.trim())return alert("Titre requis");addE&&addE({jeuneId:majeur.id,date:eDate,titre:eTitre.trim(),description:eDesc.trim(),gravite:eGrav});setETitre("");setEDesc("");setShowFormE(false);};
 return(<div>
 <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12,cursor:"pointer"}} onClick={onBack}><ChevronLeft size={18}/><span style={{fontWeight:700,color:C.dark}}>Retour Majeurs</span></div>
 <div style={{...S.card,marginBottom:12}}><div style={{fontWeight:700,fontSize:16,color:C.dark}}>{majeur.prenom} {majeur.nom||""}</div><div style={{fontSize:12,color:C.light}}>Site: {majeur.site} | Statut: {majeur.statut} | Référent: {majeur.referentA||"Non affecté"}</div></div>
 <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>{tabs.map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:"6px 14px",borderRadius:20,border:"none",background:tab===t?C.primary:"#e8e8e8",color:tab===t?"#fff":C.dark,fontWeight:600,fontSize:12,cursor:"pointer"}}>{t==="fiche"?"Fiche":t==="rapports"?"Rapports":t==="incidents"?"Incidents":"Présences"}</button>)}</div>
 {tab==="fiche"&&<div style={{...S.card}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>{[["Prénom",majeur.prenom],["Nom",majeur.nom||"-"],["Site",majeur.site],["Statut",majeur.statut],["Début",majeur.dateDebut||"-"],["Fin",majeur.dateFin||"-"],["Email ASE",majeur.emailASE||"-"],["Tel parent",majeur.telParent1||"-"],["Tel jeune",majeur.telJeune||"-"],["Référent A",majeur.referentA||"-"],["Référent B",majeur.referentB||"-"]].map(([k,v])=><div key={k}><div style={{fontSize:10,fontWeight:700,color:C.light,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2}}>{k}</div><div style={{fontWeight:700,color:C.dark,fontSize:13}}>{v}</div></div>)}</div></div>}
 {tab==="rapports"&&<div>{showFormR?<div style={{...S.card,marginBottom:8}}><div style={{fontWeight:700,marginBottom:8}}>Nouveau rapport</div><input type="date" value={rDate} onChange={e=>setRDate(e.target.value)} style={{...S.input,marginBottom:6}}/><textarea value={rObs} onChange={e=>setRObs(e.target.value)} placeholder="Observation..." rows={3} style={{...S.input,marginBottom:6}}/><div style={{display:"flex",gap:6}}><button onClick={submitR} style={{...S.btn}}>Valider</button><button onClick={()=>setShowFormR(false)} style={{...S.btnO}}>Annuler</button></div></div>:null}{mr.length===0&&!showFormR?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun rapport</div>:mr.map(r=><div key={r.id} style={{...S.card,marginBottom:8}}><div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontWeight:700,color:C.dark}}>{r.date}</span><span style={{fontSize:11,color:C.light}}>{r.type||"Journalier"}</span></div><div style={{fontSize:13,color:C.dark,marginTop:4}}>{r.observation}</div></div>)}<button onClick={()=>setShowFormR(true)} style={{...S.btn,width:"100%",marginTop:8}}>+ Ajouter un rapport</button></div>}
 {tab==="incidents"&&<div>{showFormE?<div style={{...S.card,marginBottom:8}}><div style={{fontWeight:700,marginBottom:8}}>Nouvel incident</div><input type="date" value={eDate} onChange={e=>setEDate(e.target.value)} style={{...S.input,marginBottom:6}}/><input value={eTitre} onChange={e=>setETitre(e.target.value)} placeholder="Titre" style={{...S.input,marginBottom:6}}/><textarea value={eDesc} onChange={e=>setEDesc(e.target.value)} placeholder="Description..." rows={3} style={{...S.input,marginBottom:6}}/><select value={eGrav} onChange={e=>setEGrav(e.target.value)} style={{...S.input,marginBottom:6}}><option value="normal">Normal</option><option value="Moyen">Moyen</option><option value="Grave">Grave</option></select><div style={{display:"flex",gap:6}}><button onClick={submitE} style={{...S.btn}}>Valider</button><button onClick={()=>setShowFormE(false)} style={{...S.btnO}}>Annuler</button></div></div>:null}{me.length===0&&!showFormE?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucun incident</div>:me.map(e=><div key={e.id} style={{...S.card,marginBottom:8,borderLeft:"3px solid "+(e.gravite==="Grave"?"#e74c3c":e.gravite==="Moyen"?"#f39c12":"#3498db")}}><div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontWeight:700,color:C.dark}}>{e.date}</span><span style={{fontSize:11,padding:"2px 8px",borderRadius:10,background:e.gravite==="Grave"?"#fde8e8":e.gravite==="Moyen"?"#fef3e2":"#e8f4fd",color:e.gravite==="Grave"?"#e74c3c":e.gravite==="Moyen"?"#f39c12":"#3498db"}}>{e.gravite}</span></div><div style={{fontSize:13,color:C.dark,marginTop:4}}>{e.description}</div></div>)}<button onClick={()=>setShowFormE(true)} style={{...S.btn,width:"100%",marginTop:8}}>+ Signaler un incident</button></div>}
 {tab==="presences"&&<div>{Object.keys(mp).length===0?<div style={{...S.card,textAlign:"center",color:C.light}}>Aucune présence enregistrée</div>:<div style={{...S.card}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><th style={{textAlign:"left",padding:6,borderBottom:"2px solid #ddd",fontSize:11,color:C.light}}>Date</th><th style={{textAlign:"center",padding:6,borderBottom:"2px solid #ddd",fontSize:11,color:C.light}}>Matin</th><th style={{textAlign:"center",padding:6,borderBottom:"2px solid #ddd",fontSize:11,color:C.light}}>Après-midi</th></tr></thead><tbody>{Object.entries(mp).sort(([a],[b])=>b.localeCompare(a)).map(([d,v])=><tr key={d}><td style={{padding:6,borderBottom:"1px solid #eee",fontSize:12}}>{d}</td><td style={{textAlign:"center",padding:6,borderBottom:"1px solid #eee"}}>{v.a?"✅":"❌"}</td><td style={{textAlign:"center",padding:6,borderBottom:"1px solid #eee"}}>{v.b?"✅":"❌"}</td></tr>)}</tbody></table></div>}</div>}
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
      <textarea value={obs} onChange={e=>setObs(e.target.value)} rows={8} style={{width:"100%",padding:"0.5rem",borderRadius:8,border:"1.5px solid "+C.border,fontFamily:"'Nunito',sans-serif",fontSize:"0.9rem",resize:"vertical"}} placeholder={"D\u00e9crivez la journ\u00e9e sur le camp "+site+"..."}/>
      <button onClick={handleSave} style={{marginTop:8,background:saved?"#4CAF50":"linear-gradient(135deg,#6A1B9A,#4A148C)",color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",cursor:"pointer",fontWeight:700,fontSize:13,fontFamily:"inherit"}}>{saved?"\u2714 Enregistr\u00e9 !":"Enregistrer"}</button>
    </div>
    <h3 style={{fontWeight:700,fontSize:15,marginBottom:10,color:C.dark}}>Historique — {site}</h3>
    {existing.length===0&&<div style={{textAlign:"center",padding:20,color:C.light}}>Aucun rapport de site pour {site}</div>}
    {existing.map(r=><div key={r.id} style={{...S.card,marginBottom:10,borderLeft:"3px solid #6A1B9A",position:"relative"}}>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:4,alignItems:"center"}}>
        <span style={{fontWeight:700,color:C.dark}}>{r.date}</span>
        <span style={{fontSize:11,color:C.light}}>{r.author}</span>
      </div>
      <p style={{margin:"6px 0",fontSize:13,color:C.mid,whiteSpace:"pre-wrap",lineHeight:1.6}}>{r.observation}</p>
      {(r.horodatage||r.createdAt)&&<div style={{fontSize:10,color:C.light,marginTop:4}}>Horodatage : {(r.horodatage||r.createdAt||"").replace("T"," \u00e0 ").slice(0,19)}</div>}
      <div style={{display:"flex",gap:8,marginTop:8,flexWrap:"wrap"}}>
        <button onClick={()=>exportSiteReport(r)} style={{padding:"5px 14px",borderRadius:7,border:"1px solid #6A1B9A",background:"#F3E5F5",color:"#6A1B9A",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Exporter Word</button>
        {(user.role==="chef_service"||user.role==="directeur")&&onDelete&&<button onClick={()=>{if(confirm("Supprimer ce rapport de site ?"))onDelete(r.id);}} style={{padding:"5px 14px",borderRadius:7,border:"1px solid #C62828",background:"#FFEBEE",color:"#C62828",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Supprimer</button>}
      </div>
    </div>)}
  </div>);
}
function ExportPage({rapports,evenements,agenda,jeunes,majeurs,rapportsSite,onPurge}){
const[dateFrom,setDateFrom]=useState("");
const[dateTo,setDateTo]=useState("");
const[exporting,setExporting]=useState(false);
const[done,setDone]=useState(false);
const[chkRapportsJ,setChkRapportsJ]=useState(true);
const[chkRapportsH,setChkRapportsH]=useState(true);
const[chkRapportsSite,setChkRapportsSite]=useState(true);
const[chkRdv,setChkRdv]=useState(true);
const[chkEvenements,setChkEvenements]=useState(true);
const[filterJournee,setFilterJournee]=useState(true);
const[filterTelParents,setFilterTelParents]=useState(true);
const[filterTelExt,setFilterTelExt]=useState(true);
const allJ=[...(jeunes||JEUNES),...(majeurs||MAJEURS)];
const jName=(id)=>{const j=allJ.find(x=>x.id===id);return j?(j.prenom+" "+(j.nom||"")):"ID:"+id;};
const allChecked=chkRapportsJ&&chkRapportsH&&chkRapportsSite&&chkRdv&&chkEvenements;
const toggleAll=()=>{const v=!allChecked;setChkRapportsJ(v);setChkRapportsH(v);setChkRapportsSite(v);setChkRdv(v);setChkEvenements(v);};
const TC_LABELS={"journee":"Journée du jeune","tel_parents":"RDV tél. parents","tel_ext":"RDV tél. contact extérieur"};
const fR=(rapports||[]).filter(r=>r.date>=dateFrom&&r.date<=dateTo);
const fRFiltered=fR.filter(r=>{const tc=r.typeContact||"journee";if(tc==="journee"&&!filterJournee)return false;if(tc==="tel_parents"&&!filterTelParents)return false;if(tc==="tel_ext"&&!filterTelExt)return false;return true;});
const doExport=async()=>{
if(!dateFrom||!dateTo)return alert("Sélectionnez une période");
if(!chkRapportsJ&&!chkRapportsH&&!chkRapportsSite&&!chkRdv&&!chkEvenements)return alert("Sélectionnez au moins un type de données");
setExporting(true);
try{
const XLSX=await import("https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs");
const wb=XLSX.utils.book_new();
if(chkRapportsJ){
  const sorted=fRFiltered.sort((a,b)=>a.date.localeCompare(b.date));
  const rowsR=[["Date","Jeune","Type de contact","Observation","Horodatage","Auteur"]];
  sorted.forEach(r=>rowsR.push([r.date,jName(r.jeuneId),TC_LABELS[r.typeContact||"journee"]||r.typeContact||"Journée",r.observation||"",r.horodatage||r.createdAt||"",r.author||""]));
  const ws=XLSX.utils.aoa_to_sheet(rowsR);
  ws["!cols"]=[{wch:12},{wch:25},{wch:28},{wch:60},{wch:20},{wch:20}];
  XLSX.utils.book_append_sheet(wb,ws,"Rapports journaliers");
}
if(chkRapportsH){
  const getISOWeek=(dateStr)=>{const d=new Date(dateStr);d.setHours(0,0,0,0);d.setDate(d.getDate()+3-(d.getDay()+6)%7);const w1=new Date(d.getFullYear(),0,4);return String(1+Math.round(((d-w1)/86400000-3+(w1.getDay()+6)%7)/7)).padStart(2,"0")};
  const weekData={};
  fR.forEach(r=>{const wk=getISOWeek(r.date);const jn=jName(r.jeuneId);if(!weekData[wk])weekData[wk]={};if(!weekData[wk][jn])weekData[wk][jn]=[];weekData[wk][jn].push(r);});
  const rowsH=[["Semaine","Jeune","Nb rapports","Résumé semaine"]];
  Object.keys(weekData).sort().forEach(wk=>{Object.keys(weekData[wk]).sort().forEach(jn=>{const rs=weekData[wk][jn];const summary=rs.sort((a,b)=>a.date.localeCompare(b.date)).map(r=>{const dt=new Date(r.date);const dn=dt.toLocaleDateString("fr-FR",{weekday:"short"});return dn+": "+(r.observation||"").substring(0,80);}).join(" | ");rowsH.push(["S"+wk,jn,rs.length,summary]);});});
  const ws=XLSX.utils.aoa_to_sheet(rowsH);
  ws["!cols"]=[{wch:10},{wch:25},{wch:12},{wch:80}];
  XLSX.utils.book_append_sheet(wb,ws,"Rapports hebdos");
}
if(chkRapportsSite){
  const fRS=(rapportsSite||[]).filter(r=>r.date>=dateFrom&&r.date<=dateTo).sort((a,b)=>a.date.localeCompare(b.date));
  const rowsRS=[["Date","Site","Auteur","Contenu"]];
  fRS.forEach(r=>rowsRS.push([r.date,r.site||"",r.author||"",r.contenu||r.text||""]));
  const ws=XLSX.utils.aoa_to_sheet(rowsRS);
  ws["!cols"]=[{wch:12},{wch:15},{wch:20},{wch:80}];
  XLSX.utils.book_append_sheet(wb,ws,"Rapports de site");
}
if(chkRdv){
  const fA=(agenda||[]).filter(a=>a.date>=dateFrom&&a.date<=dateTo).sort((a,b)=>a.date.localeCompare(b.date));
  const rowsA=[["Date","Heure","Jeune","Type RDV","Lieu","Interlocuteur","Notes","Créé par"]];
  fA.forEach(a=>rowsA.push([a.date,a.heure||"",a.jeuneNom||jName(a.jeuneId),a.type||"",a.lieu||"",a.interlocuteur||"",a.notes||"",a.createdBy||""]));
  const ws=XLSX.utils.aoa_to_sheet(rowsA);
  ws["!cols"]=[{wch:12},{wch:8},{wch:25},{wch:15},{wch:25},{wch:25},{wch:50},{wch:20}];
  XLSX.utils.book_append_sheet(wb,ws,"Rendez-vous");
}
if(chkEvenements){
  const fE=(evenements||[]).filter(e=>e.date>=dateFrom&&e.date<=dateTo).sort((a,b)=>a.date.localeCompare(b.date));
  const rowsE=[["Date","Jeune","Type","Titre","Description","Gravité","N° Suivi","Catégorie","Horodatage","Auteur"]];
  fE.forEach(e=>rowsE.push([e.date,jName(e.jeuneId),e.type||"événement",e.titre||"",e.description||"",e.gravite||"",e.numeroSuivi||"",e.categorie||"",e.horodatage||"",e.author||""]));
  const ws=XLSX.utils.aoa_to_sheet(rowsE);
  ws["!cols"]=[{wch:12},{wch:25},{wch:18},{wch:25},{wch:60},{wch:12},{wch:14},{wch:18},{wch:20},{wch:20}];
  XLSX.utils.book_append_sheet(wb,ws,"Événements indésirables");
}
XLSX.writeFile(wb,"PDSR_Export_"+dateFrom+"_"+dateTo+".xlsx");
setDone(true);
}catch(err){alert("Erreur export: "+err.message);}
setExporting(false);
};
const doPurge=()=>{
if(!dateFrom||!dateTo)return alert("Sélectionnez une période");
if(!confirm("Supprimer toutes les données (rapports, événements, RDV) de la période "+dateFrom+" au "+dateTo+" ? Cette action est irréversible."))return;
onPurge(dateFrom,dateTo);
setDone(false);
alert("Données purgées pour la période sélectionnée.");
};
const Chk=({checked,onChange,label,count})=><label style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",background:checked?"#FDF3D0":"#f9f9f9",border:checked?"1.5px solid "+C.gold:"1.5px solid #e5e7eb",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600,color:checked?C.dark:"#666",transition:"all 0.15s"}}><input type="checkbox" checked={checked} onChange={onChange} style={{accentColor:C.gold,width:16,height:16}}/>{label}{count!==undefined&&<span style={{fontSize:11,color:C.light,fontWeight:400}}>({count})</span>}</label>;
return(<div style={{maxWidth:600,margin:"0 auto"}}>
<h2 style={{fontSize:18,fontWeight:900,color:C.dark,marginBottom:16}}>Export Excel avancé</h2>
<div style={{...S.card,marginBottom:16}}>
<div style={{fontWeight:700,fontSize:14,color:C.dark,marginBottom:12}}>Période</div>
<div style={{display:"flex",gap:12,marginBottom:16}}>
<div style={{flex:1}}><label style={{fontSize:11,fontWeight:700,color:C.light,display:"block",marginBottom:4}}>Du</label><input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} style={{...S.inp,width:"100%"}}/></div>
<div style={{flex:1}}><label style={{fontSize:11,fontWeight:700,color:C.light,display:"block",marginBottom:4}}>Au</label><input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} style={{...S.inp,width:"100%"}}/></div>
</div>
</div>
<div style={{...S.card,marginBottom:16}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
  <div style={{fontWeight:700,fontSize:14,color:C.dark}}>Données à exporter</div>
  <button onClick={toggleAll} style={{fontSize:11,fontWeight:700,color:C.gold,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>{allChecked?"Tout désélectionner":"Tout sélectionner"}</button>
</div>
<div style={{display:"flex",flexDirection:"column",gap:8}}>
  <Chk checked={chkRapportsJ} onChange={()=>setChkRapportsJ(!chkRapportsJ)} label="Rapports journaliers" count={dateFrom&&dateTo?fRFiltered.length:undefined}/>
  {chkRapportsJ&&<div style={{marginLeft:28,display:"flex",flexWrap:"wrap",gap:6}}>
    <label style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:C.dark,cursor:"pointer"}}><input type="checkbox" checked={filterJournee} onChange={()=>setFilterJournee(!filterJournee)} style={{accentColor:C.gold}}/>Journée du jeune</label>
    <label style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:C.dark,cursor:"pointer"}}><input type="checkbox" checked={filterTelParents} onChange={()=>setFilterTelParents(!filterTelParents)} style={{accentColor:C.gold}}/>RDV tél. parents</label>
    <label style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:C.dark,cursor:"pointer"}}><input type="checkbox" checked={filterTelExt} onChange={()=>setFilterTelExt(!filterTelExt)} style={{accentColor:C.gold}}/>RDV tél. contact ext.</label>
  </div>}
  <Chk checked={chkRapportsH} onChange={()=>setChkRapportsH(!chkRapportsH)} label="Rapports hebdos (compilation)" count={dateFrom&&dateTo?fR.length:undefined}/>
  <Chk checked={chkRapportsSite} onChange={()=>setChkRapportsSite(!chkRapportsSite)} label="Rapports de site" count={dateFrom&&dateTo?(rapportsSite||[]).filter(r=>r.date>=dateFrom&&r.date<=dateTo).length:undefined}/>
  <Chk checked={chkRdv} onChange={()=>setChkRdv(!chkRdv)} label="Rendez-vous" count={dateFrom&&dateTo?(agenda||[]).filter(a=>a.date>=dateFrom&&a.date<=dateTo).length:undefined}/>
  <Chk checked={chkEvenements} onChange={()=>setChkEvenements(!chkEvenements)} label="Événements indésirables" count={dateFrom&&dateTo?(evenements||[]).filter(e=>e.date>=dateFrom&&e.date<=dateTo).length:undefined}/>
</div>
</div>
<div style={{display:"flex",gap:10}}>
<button onClick={doExport} disabled={exporting} style={{flex:1,background:C.gold,color:"#fff",border:"none",borderRadius:8,padding:"12px 20px",fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{exporting?"Export en cours...":"Exporter en Excel (.xlsx)"}</button>
</div>
{done&&<div style={{marginTop:16,padding:12,background:"#E8F5E9",borderRadius:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<span style={{color:"#2E7D32",fontWeight:700,fontSize:13}}>Export terminé !</span>
<button onClick={doPurge} style={{background:"#C62828",color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Purger cette période</button>
</div>}
</div>);
}

export default function App(){
  const[user,setUser]=useState(null),[ page,setPage]=useState("dashboard"),[ open,setOpen]=useState(false),[ sel,setSel]=useState(null);
  const lsData=loadLS();
  const[rapports,setRapports]=useState(Array.isArray(lsData?.rapports)?lsData.rapports:INIT_RAPPORTS),[ presences,setPresences]=useState(Array.isArray(lsData?.presences)?lsData.presences:INIT_PRESENCES),[ evenements,setEvenements]=useState(Array.isArray(lsData?.evenements)?lsData.evenements:INIT_EV);
  const[appUsers,setAppUsers]=useState(USERS);
  const[appJeunes,setAppJeunes]=useState(JEUNES);
 const[rapportsSite,setRapportsSite]=useState([]);
  const[agenda,setAgenda]=useState(lsData?.agenda||[]);
 const[loginLogs,setLoginLogs]=useState(lsData?.loginLogs||[]);
const[appMajeurs,setAppMajeurs]=useState(lsData?.majeurs||MAJEURS);
const[appDjiPlan,setAppDjiPlan]=useState(lsData?.djiPlan||DJI_PLAN);
const[appFatPlan,setAppFatPlan]=useState(lsData?.fatPlan||FAT_PLAN);
  
  useEffect(()=>{if(window._lsTimer)clearTimeout(window._lsTimer);window._lsTimer=setTimeout(()=>{saveLS(rapports,presences,evenements,appUsers,appJeunes,agenda,loginLogs,appMajeurs,rapportsSite,appDjiPlan,appFatPlan);},500);},[rapports,presences,evenements,appUsers,appJeunes,agenda,loginLogs,appMajeurs,rapportsSite,appDjiPlan,appFatPlan,user]);
// Firebase sync — merge-based to avoid overwriting other sessions
const fbSkip=useRef(false);
const toArr=(v)=>!v?[]:Array.isArray(v)?v.filter(Boolean):Object.values(v).filter(Boolean);
const mergeArr=(local,remote)=>{const map=new Map();toArr(remote).forEach(x=>{if(x&&x.id!=null)map.set(x.id,x);});local.forEach(x=>{if(x&&x.id!=null)map.set(x.id,x);});return[...map.values()];};
useEffect(()=>{if(fbSkip.current){fbSkip.current=false;return;}if(!user)return;if(window._fbTimer)clearTimeout(window._fbTimer);window._fbTimer=setTimeout(async()=>{const rem=await fbGet("data");const merged={rapports:mergeArr(rapports,rem?.rapports),presences,evenements:mergeArr(evenements,rem?.evenements),jeunes:appJeunes,users:appUsers,agenda:mergeArr(agenda,rem?.agenda),loginLogs:mergeArr(loginLogs,rem?.loginLogs),majeurs:appMajeurs,rapportsSite:mergeArr(rapportsSite,rem?.rapportsSite),djiPlan:appDjiPlan,fatPlan:appFatPlan};fbSet("data",merged);},2000);},[rapports,presences,evenements,appUsers,appJeunes,agenda,loginLogs,appMajeurs,rapportsSite,appDjiPlan,appFatPlan]);
const applyFb=(d)=>{if(!d)return;fbSkip.current=true;if(d.rapports)setRapports(prev=>mergeArr(prev,d.rapports));if(d.presences)setPresences(typeof d.presences==="object"&&!Array.isArray(d.presences)?d.presences:Array.isArray(d.presences)?d.presences:INIT_PRESENCES);if(d.evenements)setEvenements(prev=>mergeArr(prev,d.evenements));if(d.jeunes){const jArr=toArr(d.jeunes);setAppJeunes(jArr.map(fj=>{const base=JEUNES.find(x=>x.id===fj.id)||fj;return{...base,...fj};}));}if(d.agenda)setAgenda(prev=>mergeArr(prev,d.agenda));if(d.loginLogs)setLoginLogs(prev=>mergeArr(prev,d.loginLogs));if(d.users)setAppUsers(prev=>{const fb=toArr(d.users);return fb.map(fu=>{const base=USERS.find(x=>x.id===fu.id)||fu;return{...base,...fu};});});if(d.majeurs)setAppMajeurs(toArr(d.majeurs));if(d.rapportsSite)setRapportsSite(prev=>mergeArr(prev,d.rapportsSite));if(d.djiPlan&&typeof d.djiPlan==="object")setAppDjiPlan(prev=>({...DJI_PLAN,...d.djiPlan}));if(d.fatPlan&&typeof d.fatPlan==="object")setAppFatPlan(prev=>({...FAT_PLAN,...d.fatPlan}));};
useEffect(()=>{(async()=>{applyFb(await fbGet("data"));})();},[]);
useEffect(()=>{if(!user)return;const iv=setInterval(async()=>{const d=await fbGet("data");if(d){fbSkip.current=true;applyFb(d);}},20000);return()=>clearInterval(iv);},[user]);

  if(!user)return<Login onLogin={u=>{if(u.role==="educateur"||u.role==="coordinateur_site"){const pool=u.isEducMajeur?[...(appMajeurs||MAJEURS)]:[...(appJeunes||JEUNES)];u.assignedIds=pool.filter(j=>j.referentA===u.name||j.referentB===u.name).map(j=>j.id);}setLoginLogs(prev=>[{id:Date.now(),user:u.name||u.login,role:u.role,date:new Date().toISOString(),ts:Date.now()},...prev].slice(0,500));setUser(u);setPage("dashboard");}}/>;
  const addR=({jeuneId,date,observation,typeContact})=>{const now=new Date();const pad2=n=>String(n).padStart(2,"0");const ts=now.getFullYear()+"-"+pad2(now.getMonth()+1)+"-"+pad2(now.getDate())+"T"+pad2(now.getHours())+":"+pad2(now.getMinutes())+":"+pad2(now.getSeconds());setRapports(p=>{const dup=p.find(r=>r.jeuneId===jeuneId&&r.date===date);if(dup)return p.map(r=>r.id===dup.id?{...r,observation,typeContact:typeContact||r.typeContact,updatedAt:ts,author:user?.name||"?"}:r);return[...p,{id:Date.now(),jeuneId,date,observation,typeContact:typeContact||"journee",createdAt:ts,horodatage:ts,author:user?.name||"?"}];});};
  const delR=(id)=>setRapports(p=>p.filter(r=>r.id!==id));
  const delE=(id)=>setEvenements(p=>p.filter(e=>e.id!==id));
  const addE=ev=>{const now=new Date();const pad2=n=>String(n).padStart(2,"0");const ts=now.getFullYear()+"-"+pad2(now.getMonth()+1)+"-"+pad2(now.getDate())+"T"+pad2(now.getHours())+":"+pad2(now.getMinutes())+":"+pad2(now.getSeconds());const newEv={id:Date.now(),...ev,author:user?.name||"?",createdAt:ts,horodatage:ts};setEvenements(p=>[...p,newEv]);try{const j=(appJeunes||[]).find(j2=>j2.id===ev.jeuneId);const nom=j?(j.prenom+" "+(j.nom||"")):("ID:"+ev.jeuneId);const bom="\ufeff";const rows=[["Date","Jeune","Titre","Description","Gravit\u00e9"],[ev.date||"",nom,ev.titre||"",ev.description||"",ev.gravite||"normal"]];const csv=rows.map(r=>r.map(c=>'"'+String(c).replace(/"/g,'""')+'"').join(";")).join("\n");const blob=new Blob([bom+csv],{type:"text/csv;charset=utf-8"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download="incident_"+((ev.date||"").replace(/-/g,""))+"_"+nom.replace(/\s/g,"_")+".csv";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);}catch(e){console.error("Auto CSV export error:",e);}};
  const changeP=(jeuneId,date,statut)=>setPresences(p=>[...p.filter(p2=>!(p2.jeuneId===jeuneId&&p2.date===date)),{id:`${jeuneId}-${date}`,jeuneId,date,statut}]);
  const TITLES={dashboard:"Tableau de bord",jeunes:"Jeunes","jeune-detail":sel?`${sel.prenom} ${sel.nom}`:"Fiche",rapports:"Rapports journaliers",presences:"Présences",evenements:"Événements indésirables","rapport-hebdo":"Rapport hebdomadaire","rapport-site":"Rapport de site"};
  return(<div style={{fontFamily:"'Nunito',sans-serif",background:"linear-gradient(180deg,"+C.sableLight+" 0%,#F5EFE0 100%)",minHeight:"100vh",display:"flex",flexDirection:"column"}}>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap');*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}html{scroll-behavior:smooth}body{margin:0;-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale}button{transition:all 0.15s ease}button:hover{opacity:0.92;transform:translateY(-1px)}button:active{transform:translateY(0);opacity:1}input,select,textarea{transition:border-color 0.2s ease}input:focus,select:focus,textarea:focus{border-color:${C.gold}!important;outline:none;box-shadow:0 0 0 3px ${C.goldLight}}::selection{background:${C.goldLight};color:${C.dark}}::-webkit-scrollbar{width:6px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:${C.sableDark};border-radius:3px}@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}@keyframes slideIn{from{opacity:0;transform:translateX(-12px)}to{opacity:1;transform:translateX(0)}}@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.6}}`}</style>
    <Sidebar page={page} onNav={setPage} user={user} onLogout={()=>setUser(null)} open={open} onClose={()=>setOpen(false)}/>
    <div style={{flex:1,display:"flex",flexDirection:"column"}}>
      <Topbar title={TITLES[page]||"PDSR"} onMenu={()=>setOpen(true)} onBack={page==="jeune-detail"?()=>setPage("jeunes"):undefined}/>
      <main style={{flex:1,overflowY:"auto"}}>
        {page==="dashboard"&&<Dashboard setPage={setPage} user={user} rapports={rapports} presences={presences} evenements={evenements} onNav={setPage} setSel={setSel} jeunes={appJeunes} agenda={agenda} majeurs={appMajeurs}/>}
        {page==="jeunes"&&<JeunesList user={user} jeunes={appJeunes} presences={presences} onSelect={setSel} onNav={setPage} onUpdateJeune={(id,field,val)=>{setAppJeunes(prev=>prev.map(j=>j.id===id?{...j,[field]:val}:j));}}/>}
        {page==="majeurs"&&<div><div style={{...S.card,marginBottom:12}}><div style={{fontWeight:700,fontSize:16,color:C.dark,marginBottom:12}}>Jeunes Majeurs</div><div style={{fontSize:12,color:C.light,marginBottom:8}}>Section des jeunes majeurs</div></div>{(appMajeurs||MAJEURS).map(m=><div key={m.id} onClick={()=>{setSel(m);setPage("majeur-detail");}} style={{...S.card,marginBottom:8,cursor:"pointer",display:"flex",alignItems:"center",gap:12}}><div style={{width:36,height:36,borderRadius:18,background:C.primary,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:13}}>{m.prenom[0]}{(m.nom||"")[0]||""}</div><div><div style={{fontWeight:700,color:C.dark,fontSize:14}}>{m.prenom} {m.nom}</div><div style={{fontSize:11,color:C.light}}>{m.site} | {m.dateDebut} - {m.dateFin}</div></div></div>)}</div>}
        {page==="majeur-detail"&&sel&&<MajeurDetail majeur={sel} rapports={rapports} presences={presences} evenements={evenements} user={user} onBack={()=>setPage("majeurs")} onAddR={j=>{setSel(j);setPage("rapports");}} onAddE={j=>{setSel(j);setPage("evenements");}} onCP={changeP} users={appUsers} addR={r=>{addR(r);}} addE={ev=>{addE(ev);}}/>}
        {page==="jeune-detail"&&sel&&<JeuneDetail jeune={appJeunes.find(j=>j.id===sel.id)||sel} rapports={rapports} presences={presences} evenements={evenements} user={user} onAddR={j=>{setSel(j);setPage("rapports");}} onAddE={j=>{setSel(j);setPage("evenements");}} onCP={changeP} users={appUsers} onUpdateJeune={(id,field,val)=>{setAppJeunes(prev=>prev.map(j=>j.id===id?{...j,[field]:val}:j));}}/>}
        {page==="rapports"&&<Rapports user={user} rapports={rapports} onSave={addR} onDelete={delR} majeurs={appMajeurs}/>}
        {page==="presences"&&<Presences user={user} presences={presences} onCP={changeP}/>}
        {page==="evenements"&&<Evenements user={user} evenements={evenements} onAdd={addE} onDelete={delE} majeurs={appMajeurs} onUpdateAll={setEvenements}/>}
        {page==="agenda"&&<AgendaPage agenda={agenda} setAgenda={setAgenda} jeunes={appJeunes} majeurs={MAJEURS} users={appUsers} user={user}/>}
        {page==="rapport-site"&&(user.role==="coordinateur_site"||user.role==="chef_service"||user.role==="directeur")&&<RapportSite user={user} rapportsSite={rapportsSite} onSave={r=>{setRapportsSite(prev=>{const idx=prev.findIndex(x=>x.id===r.id);if(idx>=0){const cp=[...prev];cp[idx]=r;return cp;}return[...prev,r];});}} onDelete={id=>{setRapportsSite(prev=>prev.filter(x=>x.id!==id));}}/>}
        {page==="export"&&(user.role==="directeur"||user.role==="chef_service"||user.role==="coordinateur_site")&&<ExportPage rapports={rapports} evenements={evenements} agenda={agenda} jeunes={appJeunes} majeurs={appMajeurs} rapportsSite={rapportsSite} onPurge={(from,to)=>{setRapports(p=>p.filter(r=>r.date<from||r.date>to));setEvenements(p=>p.filter(e=>e.date<from||e.date>to));setAgenda(p=>p.filter(a=>a.date<from||a.date>to));}}/>}
      {page==="admin"&&(user.role==="directeur"||user.role==="chef_service"||user.role==="coordinateur_site")&&<Admin users={appUsers} jeunes={appJeunes} onUpdateUsers={setAppUsers} onUpdateJeunes={setAppJeunes} loginLogs={loginLogs} appMajeurs={appMajeurs} onUpdateMajeurs={(id,field,val,fullArr)=>{if(fullArr){setAppMajeurs(fullArr);}else{setAppMajeurs(prev=>(prev||MAJEURS).map(m=>m.id===id?{...m,[field]:val}:m));}}}/>}
        {page==="rapport-hebdo"&&<RapportHebdo user={user} rapports={rapports} presences={presences} evenements={evenements} jeunes={appJeunes} majeurs={appMajeurs}/>}
      {page==="planning"&&<Planning djiPlan={appDjiPlan} fatPlan={appFatPlan} site={user.site} user={user} onUpdate={(siteName,key,data)=>{if(siteName==="Djilass")setAppDjiPlan(prev=>({...prev,[key]:data}));else setAppFatPlan(prev=>({...prev,[key]:data}));}}/>}
      </main>
    </div>
  </div>);
}
