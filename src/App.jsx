import React, { useEffect, useMemo, useState } from "react";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import { LineChart, Line, CartesianGrid, XAxis, YAxis, Tooltip as RTooltip, ResponsiveContainer, BarChart, Bar } from "recharts";

const months = ["01","02","03","04","05","06","07","08","09","10","11","12"];
const monthLabel = (m) => ({ "01":"Jan","02":"Feb","03":"Mär","04":"Apr","05":"Mai","06":"Jun","07":"Jul","08":"Aug","09":"Sep","10":"Okt","11":"Nov","12":"Dez" }[m]||m);
const clsx = (...a)=>a.filter(Boolean).join(" ");
const lsGet=(k,f)=>{ try{const v=localStorage.getItem(k); return v?JSON.parse(v):f}catch{return f}};
const lsSet=(k,v)=>{ try{localStorage.setItem(k, JSON.stringify(v));}catch{} };
const prevMonth=(y,m)=>{const i=months.indexOf(m);if(i<=0)return{year:String(Number(y)-1),month:"12"};return{year:y,month:months[i-1]}};
const nextMonth=(y,m)=>{const i=months.indexOf(m);if(i===11)return{year:String(Number(y)+1),month:"01"};return{year:y,month:months[i+1]}};

const DEFAULT_KPIS=[
  { id:"kpi_beratungen", name:"Beratungen gesamt", unit:"Stk.", target:20, direction:"higher" },
  { id:"kpi_vorort", name:"Vor-Ort-Termine", unit:"Stk.", target:6, direction:"higher" },
  { id:"kpi_anfragen", name:"Bearbeitete Anfragen/Tickets", unit:"Stk.", target:30, direction:"higher" },
  { id:"kpi_pruefungen", name:"Geprüfte Unterlagen/Genehmigungen", unit:"Stk.", target:10, direction:"higher" },
  { id:"kpi_streit", name:"Gelöste Streitfälle/Moderationen", unit:"Stk.", target:2, direction:"higher" },
  { id:"kpi_durchlauf", name:"Ø Durchlaufzeit Kernprozess", unit:"Tage", target:14, direction:"lower" },
  { id:"kpi_trasse", name:"Verlegte Trasse", unit:"km", target:5, direction:"higher" },
  { id:"kpi_anschluesse", name:"Neue Anschlüsse", unit:"Haushalte", target:50, direction:"higher" },
  { id:"kpi_schulungen", name:"Schulungen/Workshops", unit:"Stk.", target:1, direction:"higher" },
  { id:"kpi_teilnehmer", name:"Teilnehmende Schulungen", unit:"Pers.", target:15, direction:"higher" },
];
const DEFAULT_CONSULTANTS=[{id:"c1",name:"Beraterin A"},{id:"c2",name:"Berater B"},{id:"c3",name:"Berater C"}];
const DEFAULT_THRESHOLDS={green:1.0,yellow:0.7};

const DEFAULT_STATE={
  config:{
    kpis:DEFAULT_KPIS,
    consultants:DEFAULT_CONSULTANTS,
    thresholds:DEFAULT_THRESHOLDS,
    brand:{ logo:"/Micus_Logo.jpg", position:"top-right", size:96, opacity:0.9 }
  },
  data:{}
};

function App(){
  const [state,setState]=useState(()=>lsGet("controlling-tool-state-vite", DEFAULT_STATE));
  const [year,setYear]=useState(String(new Date().getFullYear()));
  const [month,setMonth]=useState(months[new Date().getMonth()]);
  const ym = `${year}-${month}`;

  useEffect(()=>{ lsSet("controlling-tool-state-vite", state)},[state]);

  const {kpis,consultants,thresholds,brand}=state.config;
  useEffect(()=>{
    if(!state.data[ym]) setState(s=>({...s,data:{...s.data,[ym]:emptyMonth()}}));
  },[ym]);

  const current=state.data[ym]||emptyMonth();
  const {year:py,month:pm}=prevMonth(year,month);
  const previous=state.data[`${py}-${pm}`]||emptyMonth();

  function emptyMonth(){
    const m={}; for(const c of state.config.consultants){ m[c.id]=Object.fromEntries(state.config.kpis.map(k=>[k.id,0])) }
    return { kpis:m, narrative:{highlights:"",risks:"",learnings:"",needs:""}, actions:[], kvp:[], feedback:[] };
  }
  function setMonthField(path, value){
    setState(s=>({...s,data:{...s.data,[ym]:deepSet(s.data[ym]||emptyMonth(),path,value)}}));
  }
  function deepSet(obj,path,value){
    const parts=Array.isArray(path)?path:String(path).split("."); const o=structuredClone(obj); let cur=o;
    for(let i=0;i<parts.length-1;i++){ const p=parts[i]; if(cur[p]==null) cur[p]={}; cur=cur[p]; }
    cur[parts[parts.length-1]]=value; return o;
  }
  const sumKpi=(kpiId, m=current)=> state.config.consultants.reduce((a,c)=>a+Number(m.kpis?.[c.id]?.[kpiId]||0),0);
  const combined= (m=current)=> state.config.kpis.map(k=>({...k,value:sumKpi(k.id,m)}));

  function exportKPIsXLSX(){
    const wb=XLSX.utils.book_new();
    const header=["KPI","Einheit","Soll","Richtung",...consultants.map(c=>c.name),"Gesamt"];
    const rows=kpis.map(k=>{
      const line=[k.name,k.unit,k.target,k.direction];
      let total=0; for(const c of consultants){ const v=Number(current.kpis?.[c.id]?.[k.id]||0); line.push(v); total+=v }
      line.push(total); return line;
    });
    const ws=XLSX.utils.aoa_to_sheet([header,...rows]);
    XLSX.utils.book_append_sheet(wb, ws, ym);
    const wbout=XLSX.write(wb,{type:"array",bookType:"xlsx"});
    const blob=new Blob([wbout],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=`KPI-Matrix_${ym}.xlsx`; a.click();
  }
  function exportJSON(){
    const blob=new Blob([JSON.stringify(state,null,2)],{type:"application/json"});
    const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="controlling-tool-backup.json"; a.click();
  }
  function importJSON(file){
    const reader=new FileReader();
    reader.onload=()=>{ try{ setState(JSON.parse(reader.result)) }catch{ alert("Import fehlgeschlagen: Ungültige JSON-Datei") } };
    reader.readAsText(file);
  }

  function generatePDF(){
    const doc=new jsPDF({unit:"pt",format:"a4"});
    const pageWidth=doc.internal.pageSize.getWidth();
    // Logo in PDF
    if(brand.logo){
      try{
        if (brand.logo.startsWith("data:image/")) {
          let fmt = "PNG";
          if (brand.logo.startsWith("data:image/jpeg") || brand.logo.startsWith("data:image/jpg")) fmt = "JPEG";
          else if (brand.logo.startsWith("data:image/png")) fmt = "PNG";
          doc.addImage(brand.logo, fmt, pageWidth - 124, 16, 110, 36);
        }
      }catch(e){}
    }
    doc.setFontSize(16);
    doc.text(`Monatsbericht – ${monthLabel(month)} ${year}`, 40, 40);
    doc.setFontSize(12);
    doc.text("KPI-Cockpit", 40, 70);

    const kpiRows=kpis.map(k=>{
      const total=sumKpi(k.id,current);
      const prev=sumKpi(k.id,previous);
      const trend= total===prev ? "→" : total>prev ? "↑" : "↓";
      return [k.name, String(total), `${k.target} ${k.unit}`, trend];
    });
    autoTable(doc,{ startY:80, head:[["KPI","Ist","Soll","Trend"]], body:kpiRows, styles:{fontSize:9}, headStyles:{fillColor:[240,240,240]}, margin:{left:40,right:40} });

    const after=doc.lastAutoTable ? doc.lastAutoTable.finalY + 20 : 140;
    let y=after;
    const sections=[
      ["Highlights & Erfolge", current.narrative.highlights],
      ["Risiken / Probleme", current.narrative.risks],
      ["Learnings / Best Practices", current.narrative.learnings],
      ["Bedarfe / Entscheidungen", current.narrative.needs],
    ];
    sections.forEach(([title,text])=>{
      if(!text) return;
      doc.setFontSize(12); doc.text(String(title),40,y);
      doc.setFontSize(10);
      const lines=doc.splitTextToSize(String(text), pageWidth-80);
      y+=14; doc.text(lines,40,y); y+=lines.length*12+10;
      if(y>doc.internal.pageSize.getHeight()-80){ doc.addPage(); y=60; }
    });

    if(current.actions?.length){
      doc.addPage();
      doc.setFontSize(12); doc.text("Maßnahmenplan (Auszug)", 40, 60);
      const body=current.actions.slice(0,25).map(a=>[a.title,a.owner||"—",a.due||"—",a.status||"—"]);
      autoTable(doc,{startY:70, head:[["Maßnahme","Verantw.","Fällig","Status"]], body, styles:{fontSize:9}, headStyles:{fillColor:[240,240,240]}, margin:{left:40,right:40}});
    }
    if(current.kvp?.length){
      const after2=doc.lastAutoTable ? doc.lastAutoTable.finalY + 20 : 120;
      if(after2>doc.internal.pageSize.getHeight()-100) doc.addPage();
      doc.setFontSize(12); doc.text("KVP-Liste (Auszug)", 40, after2);
      const body=current.kvp.slice(0,25).map(k=>[k.title,k.owner||"—",k.due||"—",k.status||"—"]);
      autoTable(doc,{startY:after2+10, head:[["Eintrag","Verantw.","Fällig","Status"]], body, styles:{fontSize:9}, headStyles:{fillColor:[240,240,240]}, margin:{left:40,right:40}});
    }
    doc.save(`Monatsbericht_${ym}.pdf`);
  }

  // Trend-Serien (6 Monate)
  const timeSeries=useMemo(()=>{
    const series={}; for(const k of kpis) series[k.id]=[];
    let cy=Number(year), cm=month;
    for(let i=5;i>=0;i--){ const p=prevMonth(cy.toString(), cm); cy=Number(p.year); cm=p.month; }
    let yx=cy, mx=cm;
    for(let i=0;i<6;i++){ const key=`${yx}-${mx}`; const d=state.data[key]||emptyMonth(); 
      for(const k of kpis){ series[k.id].push({ name:`${monthLabel(mx)} ${yx}`, value: sumKpi(k.id,d) }); }
      const n=nextMonth(yx.toString(), mx); yx=Number(n.year); mx=n.month;
    }
    return series;
  },[state,year,month]);

  // UI
  const [tab,setTab]=useState("dashboard");
  const positions=[["top-left","Oben links"],["top-right","Oben rechts"],["bottom-left","Unten links"],["bottom-right","Unten rechts"]];

  function updateBrand(patch){ setState(s=>({...s, config:{...s.config, brand:{...s.config.brand, ...patch}}})) }
  function updateKpi(i,patch){ setState(s=>{ const arr=[...s.config.kpis]; arr[i]={...arr[i],...patch}; return {...s, config:{...s.config, kpis:arr}} }) }
  function removeKpi(i){ setState(s=>{ const arr=[...s.config.kpis]; const [removed]=arr.splice(i,1); const data={...s.data}; for(const m of Object.keys(data)){ for(const cid of Object.keys(data[m].kpis)){ delete data[m].kpis[cid][removed.id]; } } return {...s, config:{...s.config,kpis:arr}, data} }) }
  function addKpi(){ const id=`kpi_${Math.random().toString(36).slice(2,8)}`; setState(s=>{ const arr=[...s.config.kpis,{id,name:"Neue Kennzahl",unit:"Stk.",target:0,direction:"higher"}]; const data={...s.data}; for(const m of Object.keys(data)){ for(const cid of Object.keys(data[m].kpis)){ data[m].kpis[cid][id]=0 } } return {...s, config:{...s.config,kpis:arr}, data} }) }
  function updateConsultant(i,patch){ setState(s=>{ const arr=[...s.config.consultants]; arr[i]={...arr[i],...patch}; return {...s, config:{...s.config, consultants:arr}} }) }
  function removeConsultant(i){ setState(s=>{ const arr=[...s.config.consultants]; const [rem]=arr.splice(i,1); const data={...s.data}; for(const m of Object.keys(data)){ delete data[m].kpis[rem.id] } return {...s, config:{...s.config, consultants:arr}, data} }) }
  function addConsultant(){ const id=`c_${Math.random().toString(36).slice(2,6)}`; setState(s=>{ const arr=[...s.config.consultants,{id,name:"Neue/r Berater/in"}]; const data={...s.data}; for(const m of Object.keys(data)){ data[m].kpis[id]=Object.fromEntries(s.config.kpis.map(k=>[k.id,0])) } return {...s, config:{...s.config, consultants:arr}, data} }) }

  // Helpers
  const Ampel=({value,target,direction="higher"})=>{
    if(value==null||target==null||target===0) return <span className="badge border-gray-300 text-gray-500">n/a</span>;
    let ratio=value/target; let color="bg-red-500";
    if(direction==="higher"){ if(ratio>=thresholds.green) color="bg-green-500"; else if(ratio>=thresholds.yellow) color="bg-yellow-500"; }
    else { if(value<=target*thresholds.green) color="bg-green-500"; else if(value<=target*thresholds.yellow) color="bg-yellow-500"; }
    return <div className="flex items-center gap-2"><span className={clsx("inline-block w-2.5 h-2.5 rounded-full", color)}></span><span className="text-xs text-gray-500">Ampel</span></div>;
  };
  const Trend=({current,previous})=>{
    if(current==null||previous==null) return <span className="text-xs text-gray-400">–</span>;
    if(current===previous) return <span className="text-xs">→ stabil</span>;
    return <span className="text-xs">{current>previous?"↑ steigend":"↓ fallend"}</span>;
  };

  return (
    <div className="container mx-auto px-4 py-6 space-y-6 relative">
      {/* Fixed logo overlay */}
      {state.config.brand?.logo ? (
        <img
          src={state.config.brand.logo}
          alt="Logo"
          className={clsx("fixed z-30 pointer-events-none", state.config.brand.position.includes("top")?"top-3":"bottom-3", state.config.brand.position.includes("right")?"right-3":"left-3")}
          style={{ width: state.config.brand.size, opacity: state.config.brand.opacity??0.9 }}
        />
      ):null}

      <header className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
        <div>
          <h1 className="text-3xl font-bold">Controlling & Tätigkeitsnachweis</h1>
          <p className="text-gray-500">Monatlicher Bericht mit KPI-Cockpit, qualitativem Teil, Maßnahmenplan und KVP.</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <select className="input w-28" value={year} onChange={e=>setYear(e.target.value)}>
            {Array.from({length:6}).map((_,i)=>{
              const y=String(new Date().getFullYear()-i);
              return <option key={y} value={y}>{y}</option>
            })}
          </select>
          <select className="input w-28" value={month} onChange={e=>setMonth(e.target.value)}>
            {months.map(m=><option key={m} value={m}>{monthLabel(m)}</option>)}
          </select>
          <button className="btn btn-outline" onClick={generatePDF}>PDF erzeugen</button>
          <button className="btn btn-outline" onClick={exportKPIsXLSX}>KPI nach Excel</button>
          <div className="flex items-center gap-2">
            <button className="btn btn-outline" onClick={exportJSON}>Backup</button>
            <label className="btn btn-outline cursor-pointer">
              <input type="file" accept="application/json" className="hidden" onChange={(e)=>e.target.files?.[0]&&importJSON(e.target.files[0])} />
              Restore
            </label>
          </div>
          {/* Branding */}
          <details className="ml-2">
            <summary className="btn btn-outline">Branding / Logo</summary>
            <div className="mt-2 p-3 border rounded-xl bg-white flex flex-col gap-2">
              <div className="flex items-center gap-2">
                <label className="btn btn-outline cursor-pointer">
                  <input type="file" accept="image/*" className="hidden" onChange={(e)=>{
                    const f=e.target.files?.[0]; if(!f) return;
                    const r=new FileReader(); r.onload=()=>updateBrand({logo:String(r.result)}); r.readAsDataURL(f);
                  }} />
                  Logo wählen
                </label>
                {brand.logo && <button className="btn" onClick={()=>updateBrand({logo:""})}>Entfernen</button>}
              </div>
              <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
                <label className="label">Position
                  <select className="input" value={brand.position} onChange={e=>updateBrand({position:e.target.value})}>
                    {positions.map(([v,l])=><option key={v} value={v}>{l}</option>)}
                  </select>
                </label>
                <label className="label">Größe (px)
                  <input className="input" type="number" min={48} max={240} value={brand.size} onChange={e=>updateBrand({size:Number(e.target.value)||96})} />
                </label>
                <label className="label">Deckkraft (0–1)
                  <input className="input" type="number" step="0.05" min={0} max={1} value={brand.opacity??0.9} onChange={e=>updateBrand({opacity:Math.max(0,Math.min(1,Number(e.target.value)))})} />
                </label>
              </div>
              <p className="text-xs text-gray-500">Hinweis: Für das PDF wird das Logo eingebettet, wenn es als DataURL hochgeladen ist.</p>
            </div>
          </details>
        </div>
      </header>

      {/* Tabs */}
      <div className="tabs">
        {["dashboard","kpis","qualitativ","massnahmen","kvp"].map(t=>(
          <button key={t} className={clsx("tab", tab===t && "active")} onClick={()=>setTab(t)}>
            {t==="dashboard"?"Dashboard": t==="kpis"?"KPI-Cockpit": t==="qualitativ"?"Qualitativ": t==="massnahmen"?"Maßnahmenplan":"KVP"}
          </button>
        ))}
      </div>

      {tab==="dashboard" && (
        <section className="space-y-4">
          <h2 className="text-xl font-semibold">Übersicht</h2>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {kpis.slice(0,6).map(k=>{
              const val=sumKpi(k.id,current); const prev=sumKpi(k.id,previous); const series=timeSeries[k.id]||[];
              return (
                <div className="card" key={k.id}>
                  <div className="card-header">
                    <div className="card-title text-base">{k.name}</div>
                    <div className="card-desc">Soll: {k.target} {k.unit}</div>
                  </div>
                  <div className="card-content space-y-2">
                    <div className="flex items-center justify-between">
                      <div>
                        <div className="text-3xl font-semibold">{val}</div>
                        <div className="text-xs text-gray-500">{k.unit}</div>
                      </div>
                      <div className="flex flex-col items-end gap-1">
                        <Ampel value={val} target={k.target} direction={k.direction} />
                        <Trend current={val} previous={prev} />
                      </div>
                    </div>
                    <div className="h-16">
                      <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={series} margin={{ top:4, right:8, left:0, bottom:0 }}>
                          <CartesianGrid vertical={false} strokeDasharray="3 3" />
                          <XAxis dataKey="name" tick={{fontSize:10}} interval={series.length>6?1:0} />
                          <YAxis tick={{fontSize:10}} width={28} />
                          <RTooltip />
                          <Line type="monotone" dataKey="value" dot={false} strokeWidth={2} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>
              )
            })}
          </div>

          <div className="card">
            <div className="card-header">
              <div className="card-title">Aggregierte Leistung {monthLabel(month)} {year}</div>
              <div className="card-desc">Summe über alle Berater*innen</div>
            </div>
            <div className="card-content">
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={combined()} margin={{ top:10, right:20, left:0, bottom:0 }}>
                    <CartesianGrid vertical={false} strokeDasharray="3 3" />
                    <XAxis dataKey="name" interval={0} tick={{fontSize:10}} height={60} angle={-20} textAnchor="end" />
                    <YAxis />
                    <RTooltip />
                    <Bar dataKey="value" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        </section>
      )}

      {tab==="kpis" && (
        <section className="space-y-4">
          <h2 className="text-xl font-semibold">KPI-Cockpit</h2>
          <div className="card">
            <div className="card-header">
              <div className="card-title">Eingabe (Monat {monthLabel(month)} {year})</div>
              <div className="card-desc">Werte pro Berater*in eintragen. Summen & Ampeln sind automatisch.</div>
            </div>
            <div className="card-content overflow-auto">
              <table className="table">
                <thead>
                  <tr>
                    <th>KPI</th>
                    {consultants.map(c=><th key={c.id}>{c.name}</th>)}
                    <th>Gesamt</th><th>Soll</th><th>Ampel</th><th>Trend</th>
                  </tr>
                </thead>
                <tbody>
                  {kpis.map(k=>{
                    const total=sumKpi(k.id,current); const prev=sumKpi(k.id,previous);
                    return (
                      <tr key={k.id}>
                        <td className="font-medium">{k.name} <span className="text-gray-400">({k.unit})</span></td>
                        {consultants.map(c=>(
                          <td key={c.id}>
                            <input className="input" type="number" min={0} value={current.kpis?.[c.id]?.[k.id]??0} onChange={e=>setMonthField(["kpis",c.id,k.id], Number(e.target.value))} />
                          </td>
                        ))}
                        <td className="font-semibold">{total}</td>
                        <td>{k.target} {k.unit}</td>
                        <td><Ampel value={total} target={k.target} direction={k.direction} /></td>
                        <td><Trend current={total} previous={prev} /></td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>

          <div className="card">
            <div className="card-header">
              <div className="card-title">Konfiguration</div>
              <div className="card-desc">KPIs, Soll & Schwellen anpassen</div>
            </div>
            <div className="card-content space-y-4">
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                <div className="space-y-2">
                  <h3 className="font-semibold">KPIs</h3>
                  {kpis.map((k,idx)=>(
                    <div key={k.id} className="grid grid-cols-12 gap-2 items-center">
                      <div className="col-span-4"><input className="input" value={k.name} onChange={e=>updateKpi(idx,{name:e.target.value})} /></div>
                      <div className="col-span-2"><input className="input" value={k.unit} onChange={e=>updateKpi(idx,{unit:e.target.value})} /></div>
                      <div className="col-span-2"><input className="input" type="number" value={k.target} onChange={e=>updateKpi(idx,{target:Number(e.target.value)})} /></div>
                      <div className="col-span-3">
                        <select className="input" value={k.direction} onChange={e=>updateKpi(idx,{direction:e.target.value})}>
                          <option value="higher">Höher ist besser</option>
                          <option value="lower">Niedriger ist besser</option>
                        </select>
                      </div>
                      <div className="col-span-1 text-right"><button className="btn" onClick={()=>removeKpi(idx)}>✕</button></div>
                    </div>
                  ))}
                  <button className="btn btn-outline" onClick={addKpi}>+ KPI hinzufügen</button>
                </div>

                <div className="space-y-2">
                  <h3 className="font-semibold">Berater*innen</h3>
                  {consultants.map((c,idx)=>(
                    <div key={c.id} className="grid grid-cols-12 gap-2 items-center">
                      <div className="col-span-10"><input className="input" value={c.name} onChange={e=>updateConsultant(idx,{name:e.target.value})} /></div>
                      <div className="col-span-2 text-right"><button className="btn" onClick={()=>removeConsultant(idx)}>✕</button></div>
                    </div>
                  ))}
                  <button className="btn btn-outline" onClick={addConsultant}>+ Berater*in hinzufügen</button>

                  <div className="my-3 border-t"></div>
                  <h3 className="font-semibold">Schwellen (Ampel)</h3>
                  <div className="grid grid-cols-2 gap-2">
                    <label className="label">Grün ab (Anteil Soll)
                      <input className="input" type="number" step="0.05" value={thresholds.green} onChange={e=>setState(s=>({...s,config:{...s.config,thresholds:{...s.config.thresholds,green:Number(e.target.value)}}}))} />
                    </label>
                    <label className="label">Gelb ab (Anteil Soll)
                      <input className="input" type="number" step="0.05" value={thresholds.yellow} onChange={e=>setState(s=>({...s,config:{...s.config,thresholds:{...s.config.thresholds,yellow:Number(e.target.value)}}}))} />
                    </label>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </section>
      )}

      {tab==="qualitativ" && (
        <section className="space-y-4">
          <h2 className="text-xl font-semibold">Qualitativer Bericht</h2>
          <div className="card">
            <div className="card-header">
              <div className="card-title">Kontext & Bewertung</div>
              <div className="card-desc">Erfolge, Risiken, Learnings, Bedarfe</div>
            </div>
            <div className="card-content grid grid-cols-1 lg:grid-cols-2 gap-4">
              <div className="space-y-2">
                <label className="label">Highlights & Erfolge</label>
                <textarea className="input h-32" value={current.narrative.highlights} onChange={e=>setMonthField(["narrative","highlights"], e.target.value)} />
                <label className="label">Risiken / Probleme</label>
                <textarea className="input h-32" value={current.narrative.risks} onChange={e=>setMonthField(["narrative","risks"], e.target.value)} />
              </div>
              <div className="space-y-2">
                <label className="label">Learnings / Best Practices</label>
                <textarea className="input h-32" value={current.narrative.learnings} onChange={e=>setMonthField(["narrative","learnings"], e.target.value)} />
                <label className="label">Bedarfe / Entscheidungen</label>
                <textarea className="input h-32" value={current.narrative.needs} onChange={e=>setMonthField(["narrative","needs"], e.target.value)} />
              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-header">
              <div className="card-title">Feedback (Auftraggeber)</div>
              <div className="card-desc">Rückmeldung & kontinuierliche Verbesserung</div>
            </div>
            <div className="card-content space-y-3">
              <div className="grid grid-cols-1 md:grid-cols-6 gap-3 items-end">
                <div className="md:col-span-4">
                  <label className="label">Kommentar</label>
                  <textarea className="input h-24" id="feedback-text"></textarea>
                </div>
                <div>
                  <label className="label">Bewertung (1–5)</label>
                  <select className="input" id="feedback-rating" defaultValue="3">
                    {["1","2","3","4","5"].map(r=><option key={r} value={r}>{r}</option>)}
                  </select>
                </div>
                <div>
                  <button className="btn btn-outline" onClick={()=>{
                    const text = document.getElementById("feedback-text").value;
                    const rating = Number(document.getElementById("feedback-rating").value);
                    const item = { id: crypto.randomUUID(), at: new Date().toISOString(), rating, text };
                    setMonthField(["feedback"], [...(current.feedback||[]), item]);
                    document.getElementById("feedback-text").value="";
                    document.getElementById("feedback-rating").value="3";
                  }}>Feedback speichern</button>
                </div>
              </div>
              {(current.feedback||[]).length ? (
                <div className="overflow-auto">
                  <table className="table">
                    <thead><tr><th>Datum</th><th>Bewertung</th><th>Kommentar</th></tr></thead>
                    <tbody>
                      {current.feedback.map(f=>(
                        <tr key={f.id}>
                          <td>{new Date(f.at).toLocaleString()}</td>
                          <td>{f.rating} / 5</td>
                          <td className="whitespace-pre-wrap">{f.text}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ): <p className="text-gray-500 text-sm">Noch kein Feedback erfasst.</p>}
            </div>
          </div>
        </section>
      )}

      {tab==="massnahmen" && (
        <section className="space-y-4">
          <h2 className="text-xl font-semibold">Maßnahmenplan (Ausblick)</h2>
          <ActionEditor items={current.actions} onChange={(items)=>setMonthField(["actions"], items)} />
        </section>
      )}

      {tab==="kvp" && (
        <section className="space-y-4">
          <h2 className="text-xl font-semibold">KVP – Kontinuierliche Verbesserung</h2>
          <KvpEditor items={current.kvp} onChange={(items)=>setMonthField(["kvp"], items)} />
        </section>
      )}
    </div>
  )
}

function labelForStatus(s){ return ({open:"Offen",inprogress:"In Arbeit",blocked:"Blockiert",done:"Erledigt"}[s]||s) }
function nextStatus(s){ const order=["open","inprogress","done"]; const i=order.indexOf(s); return order[(i+1)%order.length] }

function ActionEditor({items, onChange}){
  const [draft,setDraft]=useState({title:"",owner:"",due:"",status:"open"});
  function add(){ if(!draft.title) return; onChange([...(items||[]), {...draft, id:crypto.randomUUID()}]); setDraft({title:"",owner:"",due:"",status:"open"}) }
  function update(id,patch){ onChange(items.map(a=>a.id===id?{...a,...patch}:a)) }
  function remove(id){ onChange(items.filter(a=>a.id!==id)) }

  return (
    <div className="card">
      <div className="card-header"><div className="card-title">Maßnahmen</div><div className="card-desc">Planung kommender Schritte</div></div>
      <div className="card-content space-y-3">
        <div className="grid grid-cols-1 md:grid-cols-5 gap-2 items-end">
          <div className="md:col-span-2">
            <label className="label">Maßnahme</label>
            <input className="input" value={draft.title} onChange={e=>setDraft({...draft,title:e.target.value})} placeholder="z.B. Abschluss Markterkundung" />
          </div>
          <div>
            <label className="label">Verantwortlich</label>
            <input className="input" value={draft.owner} onChange={e=>setDraft({...draft,owner:e.target.value})} placeholder="Name" />
          </div>
          <div>
            <label className="label">Fällig bis</label>
            <input className="input" type="date" value={draft.due} onChange={e=>setDraft({...draft,due:e.target.value})} />
          </div>
          <div>
            <label className="label">Status</label>
            <select className="input" value={draft.status} onChange={e=>setDraft({...draft,status:e.target.value})}>
              <option value="open">Offen</option>
              <option value="inprogress">In Arbeit</option>
              <option value="blocked">Blockiert</option>
              <option value="done">Erledigt</option>
            </select>
          </div>
        </div>
        <button className="btn btn-outline" onClick={add}>+ Maßnahme hinzufügen</button>

        <div className="overflow-auto">
          <table className="table">
            <thead><tr><th>Maßnahme</th><th>Verantw.</th><th>Fällig</th><th>Status</th><th className="text-right">Aktionen</th></tr></thead>
            <tbody>
              {(items||[]).map(a=>(
                <tr key={a.id}>
                  <td className="max-w-[320px]">{a.title}</td>
                  <td>{a.owner}</td>
                  <td>{a.due}</td>
                  <td><span className="badge border-gray-300">{labelForStatus(a.status)}</span></td>
                  <td className="text-right space-x-2">
                    <button className="btn btn-outline" onClick={()=>update(a.id,{status:nextStatus(a.status)})}>Status ↑</button>
                    <button className="btn" onClick={()=>remove(a.id)}>✕</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

function KvpEditor({items, onChange}){
  const [draft,setDraft]=useState({title:"",desc:"",owner:"",due:"",status:"open"});
  function add(){ if(!draft.title) return; onChange([...(items||[]), {...draft, id:crypto.randomUUID()}]); setDraft({title:"",desc:"",owner:"",due:"",status:"open"}) }
  function update(id,patch){ onChange(items.map(a=>a.id===id?{...a,...patch}:a)) }
  function remove(id){ onChange(items.filter(a=>a.id!==id)) }

  return (
    <div className="card">
      <div className="card-header"><div className="card-title">KVP-Einträge</div><div className="card-desc">Lessons Learned & Verbesserungen</div></div>
      <div className="card-content space-y-3">
        <div className="grid grid-cols-1 lg:grid-cols-5 gap-2 items-end">
          <div className="lg:col-span-2">
            <label className="label">Eintrag</label>
            <input className="input" value={draft.title} onChange={e=>setDraft({...draft,title:e.target.value})} placeholder="z.B. Checkliste erweitern" />
          </div>
          <div className="lg:col-span-3">
            <label className="label">Beschreibung</label>
            <input className="input" value={draft.desc} onChange={e=>setDraft({...draft,desc:e.target.value})} placeholder="Kurze Erläuterung" />
          </div>
          <div>
            <label className="label">Verantwortlich</label>
            <input className="input" value={draft.owner} onChange={e=>setDraft({...draft,owner:e.target.value})} placeholder="Name" />
          </div>
          <div>
            <label className="label">Fällig bis</label>
            <input className="input" type="date" value={draft.due} onChange={e=>setDraft({...draft,due:e.target.value})} />
          </div>
          <div>
            <label className="label">Status</label>
            <select className="input" value={draft.status} onChange={e=>setDraft({...draft,status:e.target.value})}>
              <option value="open">Offen</option>
              <option value="inprogress">In Arbeit</option>
              <option value="blocked">Blockiert</option>
              <option value="done">Erledigt</option>
            </select>
          </div>
        </div>
        <button className="btn btn-outline" onClick={add}>+ KVP-Eintrag hinzufügen</button>
        <div className="overflow-auto">
          <table className="table">
            <thead><tr><th>Eintrag</th><th>Beschreibung</th><th>Verantw.</th><th>Fällig</th><th>Status</th><th className="text-right">Aktionen</th></tr></thead>
            <tbody>
              {(items||[]).map(a=>(
                <tr key={a.id}>
                  <td className="max-w-[320px]">{a.title}</td>
                  <td className="max-w-[480px]">{a.desc}</td>
                  <td>{a.owner}</td>
                  <td>{a.due}</td>
                  <td><span className="badge border-gray-300">{labelForStatus(a.status)}</span></td>
                  <td className="text-right space-x-2">
                    <button className="btn btn-outline" onClick={()=>update(a.id,{status:nextStatus(a.status)})}>Status ↑</button>
                    <button className="btn" onClick={()=>remove(a.id)}>✕</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

export default App;
