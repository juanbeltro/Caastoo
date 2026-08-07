import React from 'react';
import ReactDOM from 'react-dom/client';
import * as XLSX from 'xlsx';
import './styles.css';

type Row = Record<string, unknown>;
type CanonicalKey = 'sku'|'description'|'segment'|'salesPrevious'|'unitsPrevious'|'salesCurrent'|'unitsCurrent'|'category'|'manufacturer'|'brand';

type Field = { key: CanonicalKey; label: string; required: boolean };
const fields: Field[] = [
  {key:'sku',label:'SKU',required:true},{key:'description',label:'Descripción',required:true},{key:'segment',label:'Segmento',required:true},
  {key:'salesPrevious',label:'Venta período anterior',required:true},{key:'unitsPrevious',label:'Unidades período anterior',required:true},
  {key:'salesCurrent',label:'Venta período actual',required:true},{key:'unitsCurrent',label:'Unidades período actual',required:true},
  {key:'category',label:'Categoría / Familia',required:false},{key:'manufacturer',label:'Fabricante',required:false},{key:'brand',label:'Marca',required:false}
];

const cleanNumber=(v:unknown)=>{ if(v===null||v===undefined||v==='') return null; if(typeof v==='number') return Number.isFinite(v)?v:null; const n=Number(String(v).replace(/[$\s]/g,'').replace(/\./g,'').replace(',','.')); return Number.isFinite(n)?n:null };

function App(){
 const [fileName,setFileName]=React.useState(''); const [sheets,setSheets]=React.useState<Record<string,Row[]>>({}); const [sheet,setSheet]=React.useState('');
 const [mapping,setMapping]=React.useState<Partial<Record<CanonicalKey,string>>>({}); const [step,setStep]=React.useState(1);
 const rows=sheets[sheet]||[]; const columns=rows.length?Object.keys(rows[0]):[];
 const onFile=async(file:File)=>{const data=await file.arrayBuffer(); const wb=XLSX.read(data,{type:'array'}); const parsed:Record<string,Row[]>={}; wb.SheetNames.forEach(n=>parsed[n]=XLSX.utils.sheet_to_json<Row>(wb.Sheets[n],{defval:null})); setFileName(file.name);setSheets(parsed);setSheet(wb.SheetNames[0]||'');setMapping({});setStep(2)};
 const normalized=React.useMemo(()=>rows.map(r=>({sku:r[mapping.sku||'']??null,description:r[mapping.description||'']??null,segment:r[mapping.segment||'']??null,salesPrevious:cleanNumber(r[mapping.salesPrevious||'']),unitsPrevious:cleanNumber(r[mapping.unitsPrevious||'']),salesCurrent:cleanNumber(r[mapping.salesCurrent||'']),unitsCurrent:cleanNumber(r[mapping.unitsCurrent||'']),category:mapping.category?r[mapping.category]:null,manufacturer:mapping.manufacturer?r[mapping.manufacturer]:null,brand:mapping.brand?r[mapping.brand]:null})),[rows,mapping]);
 const missingRequired=fields.filter(f=>f.required&&!mapping[f.key]);
 const skuCounts=new Map<string,number>(); normalized.forEach(r=>{const k=String(r.sku??'').trim();if(k)skuCounts.set(k,(skuCounts.get(k)||0)+1)});
 const stats={rows:normalized.length,missingSku:normalized.filter(r=>!String(r.sku??'').trim()).length,missingSegment:normalized.filter(r=>!String(r.segment??'').trim()).length,duplicates:[...skuCounts.values()].filter(n=>n>1).length,invalidNumbers:normalized.filter(r=>['salesPrevious','unitsPrevious','salesCurrent','unitsCurrent'].some(k=>(r as any)[k]===null)).length};
 return <div className="app"><header><div><span className="eyebrow">RAMO · CATEGORY MANAGEMENT</span><h1>Assortment Tool <b>V0.1</b></h1></div><div className="privacy">● Procesamiento local · sin base de datos</div></header>
 <nav>{['Cargar','Mapear','Validar','Normalizar'].map((x,i)=><button key={x} className={step===i+1?'active':''} disabled={i+1>step}>{i+1}. {x}</button>)}</nav>
 <main>
 {step===1&&<section className="hero"><span className="tag">SPRINT 1</span><h2>Construyamos una entrada de datos confiable.</h2><p>Sube una tabla plana de Excel. El archivo se lee en tu navegador y no se envía a un servidor.</p><label className="drop"><strong>Seleccionar archivo Excel</strong><span>.xlsx · .xls</span><input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&onFile(e.target.files[0])}/></label></section>}
 {step===2&&<section><div className="sectionTitle"><div><span className="tag">FUENTE</span><h2>{fileName}</h2></div><button className="primary" disabled={!sheet||!rows.length} onClick={()=>setStep(3)}>Continuar al mapeo →</button></div><div className="card"><label>Hoja a analizar<select value={sheet} onChange={e=>{setSheet(e.target.value);setMapping({})}}>{Object.keys(sheets).map(s=><option key={s}>{s}</option>)}</select></label><div className="metrics"><div><b>{rows.length.toLocaleString()}</b><span>registros</span></div><div><b>{columns.length}</b><span>columnas</span></div><div><b>{Object.keys(sheets).length}</b><span>hojas</span></div></div></div><Preview rows={rows}/></section>}
 {step===3&&<section><div className="sectionTitle"><div><span className="tag">CONTRATO DE DATOS</span><h2>¿Qué significa cada columna?</h2><p>Mapea los campos del cliente a nuestro modelo estándar.</p></div><button className="primary" disabled={missingRequired.length>0} onClick={()=>setStep(4)}>Validar datos →</button></div><div className="mapping">{fields.map(f=><div className="mapRow" key={f.key}><div><strong>{f.label}</strong><span>{f.required?'Obligatorio':'Recomendado'}</span></div><span className="arrow">←</span><select value={mapping[f.key]||''} onChange={e=>setMapping(m=>({...m,[f.key]:e.target.value||undefined}))}><option value="">{f.required?'Seleccionar columna…':'No disponible'}</option>{columns.map(c=><option key={c} value={c}>{c}</option>)}</select></div>)}</div>{missingRequired.length>0&&<p className="warning">Faltan {missingRequired.length} campos obligatorios por mapear.</p>}</section>}
 {step===4&&<section><div className="sectionTitle"><div><span className="tag">CONTROL DE CALIDAD</span><h2>Validación previa</h2></div><button className="primary" onClick={()=>setStep(5)}>Construir modelo normalizado →</button></div><div className="validation"><Check ok={stats.missingSku===0} text={`${stats.missingSku} registros sin SKU`}/><Check ok={stats.missingSegment===0} text={`${stats.missingSegment} registros sin segmento`}/><Check ok={stats.duplicates===0} text={`${stats.duplicates} SKU duplicados`}/><Check ok={stats.invalidNumbers===0} text={`${stats.invalidNumbers} registros con ventas/unidades vacías o no numéricas`}/></div><Preview rows={normalized}/></section>}
 {step===5&&<section><div className="sectionTitle"><div><span className="tag success">NORMALIZADO</span><h2>{stats.rows.toLocaleString()} registros listos para el motor</h2><p>La fuente original permanece intacta. Esta es la estructura interna que usará Category Management.</p></div><button className="secondary" onClick={()=>setStep(3)}>← Ajustar mapeo</button></div><Preview rows={normalized}/><div className="done">✓ Definition of Done Sprint 1: carga, selección de hoja, mapping, validación y normalización funcionando.</div></section>}
 </main></div>
}
function Check({ok,text}:{ok:boolean,text:string}){return <div className={'check '+(ok?'ok':'bad')}><b>{ok?'✓':'!'}</b><span>{text}</span></div>}
function Preview({rows}:{rows:Row[]}){if(!rows.length)return <div className="empty">No hay registros para mostrar.</div>;const cols=Object.keys(rows[0]).slice(0,10);return <div className="tableWrap"><table><thead><tr>{cols.map(c=><th key={c}>{c}</th>)}</tr></thead><tbody>{rows.slice(0,8).map((r,i)=><tr key={i}>{cols.map(c=><td key={c}>{String(r[c]??'')}</td>)}</tr>)}</tbody></table><div className="tableNote">Vista previa · primeras {Math.min(8,rows.length)} filas · máximo 10 columnas</div></div>}
ReactDOM.createRoot(document.getElementById('root')!).render(<React.StrictMode><App/></React.StrictMode>);
