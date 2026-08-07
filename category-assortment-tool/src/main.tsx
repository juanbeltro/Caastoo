import React from 'react';
import ReactDOM from 'react-dom/client';
import * as XLSX from 'xlsx';
import './styles.css';

type Row = Record<string, unknown>;
type CanonicalKey = 'sku'|'description'|'segment'|'salesPrevious'|'unitsPrevious'|'salesCurrent'|'unitsCurrent'|'category'|'manufacturer'|'brand';
type NormalizedRow = Record<CanonicalKey, unknown> & { salesPrevious:number|null; unitsPrevious:number|null; salesCurrent:number|null; unitsCurrent:number|null };
type Recommendation = 'MANTENER'|'REVISAR'|'DEPURAR'|'INNOVACIÓN'|'DESCODIFICADO'|'SIN VENTA'|'DATOS INSUFICIENTES';
type AnalysisRow = NormalizedRow & { varSales:number|null; varUnits:number|null; performance:number|null; segmentWeight:number|null; pareto:number|null; recommendation:Recommendation; reason:string };

type Field = { key: CanonicalKey; label: string; required: boolean };
const fields: Field[] = [
  {key:'sku',label:'SKU',required:true},{key:'description',label:'Descripción',required:true},{key:'segment',label:'Segmento',required:true},
  {key:'salesPrevious',label:'Venta período anterior',required:true},{key:'unitsPrevious',label:'Unidades período anterior',required:true},
  {key:'salesCurrent',label:'Venta período actual',required:true},{key:'unitsCurrent',label:'Unidades período actual',required:true},
  {key:'category',label:'Categoría / Familia',required:false},{key:'manufacturer',label:'Fabricante',required:false},{key:'brand',label:'Marca',required:false}
];
const recOrder:Recommendation[]=['MANTENER','REVISAR','DEPURAR','INNOVACIÓN','DESCODIFICADO','SIN VENTA','DATOS INSUFICIENTES'];

const cleanNumber=(v:unknown)=>{ if(v===null||v===undefined||String(v).trim()==='') return 0; if(typeof v==='number') return Number.isFinite(v)?v:null; const n=Number(String(v).replace(/[$\s]/g,'').replace(/\./g,'').replace(',','.')); return Number.isFinite(n)?n:null };
const pct=(v:number|null)=>v===null?'—':`${(v*100).toLocaleString('es-CO',{minimumFractionDigits:1,maximumFractionDigits:1})}%`;
const num=(v:number|null)=>v===null?'—':v.toLocaleString('es-CO',{maximumFractionDigits:0});

function buildAnalysis(rows:NormalizedRow[]):AnalysisRow[]{
  const base:AnalysisRow[]=rows.map(r=>({...r,varSales:null,varUnits:null,performance:null,segmentWeight:null,pareto:null,recommendation:'DATOS INSUFICIENTES' as Recommendation,reason:''}));
  const validForPareto=base.filter(r=>r.salesCurrent!==null && String(r.segment??'').trim());
  const groups=new Map<string,AnalysisRow[]>();
  validForPareto.forEach(r=>{const s=String(r.segment).trim();const g=groups.get(s)||[];g.push(r);groups.set(s,g)});
  groups.forEach(group=>{
    const total=group.reduce((a,r)=>a+Math.max(0,r.salesCurrent??0),0);
    group.sort((a,b)=>(b.salesCurrent??0)-(a.salesCurrent??0));
    let cumulative=0;
    group.forEach(r=>{r.segmentWeight=total>0?Math.max(0,r.salesCurrent??0)/total:0;cumulative+=r.segmentWeight;r.pareto=cumulative});
  });
  base.forEach(r=>{
    const sp=r.salesPrevious, sc=r.salesCurrent, up=r.unitsPrevious, uc=r.unitsCurrent;
    // Casos especiales se evalúan antes de exigir cuatro métricas comparables.
    // En bases provenientes de tablas dinámicas es común que una ausencia total de venta
    // en el período actual llegue como dos celdas vacías (valor y unidades), no como 0.
    // Solo interpretamos ese PAR de vacíos como ausencia de venta para clasificar DESCODIFICADO;
    // un vacío aislado sigue siendo un problema de calidad de datos.
    const currentPairBlank=sc===null && uc===null;
    if(sp!==null && sp>0 && (sc===0 || currentPairBlank)){r.recommendation='DESCODIFICADO';r.reason=currentPairBlank?'Tenía venta en el período anterior y valor/unidades actuales están ambos vacíos: se interpreta como ausencia de venta actual.':'Tenía venta en el período anterior y la venta actual es 0.';return}
    if(sp===0 && sc!==null && sc>0){r.recommendation='INNOVACIÓN';r.reason='Venta anterior = 0 y venta actual > 0. Se protege para revisión comercial.';return}
    if(sp===0 && sc===0){r.recommendation='SIN VENTA';r.reason='No registra venta en ninguno de los dos períodos.';return}
    if(sp===null || sc===null || up===null || uc===null){r.recommendation='DATOS INSUFICIENTES';r.reason='Hay ventas o unidades vacías/no numéricas y no forman un patrón inequívoco de ausencia total de venta. El motor no completa datos silenciosamente.';return}
    if(sp<=0 || sc<0 || up<=0 || uc<0){r.recommendation='DATOS INSUFICIENTES';r.reason='La comparación requiere bases anteriores positivas y valores actuales no negativos.';return}
    r.varSales=(sc-sp)/sp; r.varUnits=(uc-up)/up; r.performance=.7*r.varSales+.3*r.varUnits;
    if(r.pareto===null){r.recommendation='DATOS INSUFICIENTES';r.reason='No fue posible calcular el Pareto dentro del segmento.';return}
    if(r.pareto<=.80){r.recommendation='MANTENER';r.reason=`Core del segmento: Pareto ${pct(r.pareto)} (≤ 80%), independientemente del desempeño ${pct(r.performance)}.`;return}
    if(r.pareto<=.95){
      if(r.performance>=0){r.recommendation='MANTENER';r.reason=`Pareto ${pct(r.pareto)} (80–95%) y desempeño ${pct(r.performance)} (≥ 0%).`}
      else {r.recommendation='REVISAR';r.reason=`Pareto ${pct(r.pareto)} (80–95%) y desempeño ${pct(r.performance)} (< 0%).`}
      return;
    }
    if(r.performance>.10){r.recommendation='REVISAR';r.reason=`Cola del segmento: Pareto ${pct(r.pareto)} (> 95%), pero desempeño ${pct(r.performance)} (> 10%).`}
    else {r.recommendation='DEPURAR';r.reason=`Cola del segmento: Pareto ${pct(r.pareto)} (> 95%) y desempeño ${pct(r.performance)} (≤ 10%).`}
  });
  return base;
}

function App(){
 const [fileName,setFileName]=React.useState(''); const [sheets,setSheets]=React.useState<Record<string,Row[]>>({}); const [sheet,setSheet]=React.useState('');
 const [mapping,setMapping]=React.useState<Partial<Record<CanonicalKey,string>>>({}); const [step,setStep]=React.useState(1); const [filter,setFilter]=React.useState<'TODOS'|Recommendation>('TODOS');
 const rows=sheets[sheet]||[]; const columns=rows.length?Object.keys(rows[0]):[];
 const onFile=async(file:File)=>{const data=await file.arrayBuffer(); const wb=XLSX.read(data,{type:'array'}); const parsed:Record<string,Row[]>={}; wb.SheetNames.forEach(n=>parsed[n]=XLSX.utils.sheet_to_json<Row>(wb.Sheets[n],{defval:null})); setFileName(file.name);setSheets(parsed);setSheet(wb.SheetNames[0]||'');setMapping({});setStep(2)};
 const normalized=React.useMemo<NormalizedRow[]>(()=>rows.map(r=>({sku:r[mapping.sku||'']??null,description:r[mapping.description||'']??null,segment:r[mapping.segment||'']??null,salesPrevious:cleanNumber(r[mapping.salesPrevious||'']),unitsPrevious:cleanNumber(r[mapping.unitsPrevious||'']),salesCurrent:cleanNumber(r[mapping.salesCurrent||'']),unitsCurrent:cleanNumber(r[mapping.unitsCurrent||'']),category:mapping.category?r[mapping.category]:null,manufacturer:mapping.manufacturer?r[mapping.manufacturer]:null,brand:mapping.brand?r[mapping.brand]:null})),[rows,mapping]);
 const analysis=React.useMemo(()=>buildAnalysis(normalized),[normalized]);
 const missingRequired=fields.filter(f=>f.required&&!mapping[f.key]);
 const skuCounts=new Map<string,number>(); normalized.forEach(r=>{const k=String(r.sku??'').trim();if(k)skuCounts.set(k,(skuCounts.get(k)||0)+1)});
 const stats={rows:normalized.length,missingSku:normalized.filter(r=>!String(r.sku??'').trim()).length,missingSegment:normalized.filter(r=>!String(r.segment??'').trim()).length,duplicates:[...skuCounts.values()].filter(n=>n>1).length,invalidNumbers:normalized.filter(r=>['salesPrevious','unitsPrevious','salesCurrent','unitsCurrent'].some(k=>(r as Record<string,unknown>)[k]===null)).length};
 const counts=Object.fromEntries(recOrder.map(x=>[x,analysis.filter(r=>r.recommendation===x).length])) as Record<Recommendation,number>;
 const filtered=filter==='TODOS'?analysis:analysis.filter(r=>r.recommendation===filter);
 return <div className="app"><header><div><span className="eyebrow">RAMO · CATEGORY MANAGEMENT</span><h1>Assortment Tool <b>V0.2.2</b></h1></div><div className="privacy">● Procesamiento local · sin base de datos</div></header>
 <nav>{['Cargar','Mapear','Validar','Normalizar','Analizar'].map((x,i)=><button key={x} className={step===i+1?'active':''} disabled={i+1>step} onClick={()=>i+1<=step&&setStep(i+1)}>{i+1}. {x}</button>)}</nav>
 <main>
 {step===1&&<section className="hero"><span className="tag">SPRINT 2</span><h2>De una base plana a una recomendación explicable.</h2><p>Sube una tabla plana de Excel. El archivo se lee en tu navegador y no se envía a un servidor.</p><label className="drop"><strong>Seleccionar archivo Excel</strong><span>.xlsx · .xls</span><input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&onFile(e.target.files[0])}/></label></section>}
 {step===2&&<section><div className="sectionTitle"><div><span className="tag">FUENTE</span><h2>{fileName}</h2></div><button className="primary" disabled={!sheet||!rows.length} onClick={()=>setStep(3)}>Continuar al mapeo →</button></div><div className="card"><label>Hoja a analizar<select value={sheet} onChange={e=>{setSheet(e.target.value);setMapping({})}}>{Object.keys(sheets).map(s=><option key={s}>{s}</option>)}</select></label><div className="metrics"><div><b>{rows.length.toLocaleString()}</b><span>registros</span></div><div><b>{columns.length}</b><span>columnas</span></div><div><b>{Object.keys(sheets).length}</b><span>hojas</span></div></div></div><Preview rows={rows}/></section>}
 {step===3&&<section><div className="sectionTitle"><div><span className="tag">CONTRATO DE DATOS</span><h2>¿Qué significa cada columna?</h2><p>Mapea los campos del cliente a nuestro modelo estándar.</p></div><button className="primary" disabled={missingRequired.length>0} onClick={()=>setStep(4)}>Validar datos →</button></div><div className="mapping">{fields.map(f=><div className="mapRow" key={f.key}><div><strong>{f.label}</strong><span>{f.required?'Obligatorio':'Recomendado'}</span></div><span className="arrow">←</span><select value={mapping[f.key]||''} onChange={e=>setMapping(m=>({...m,[f.key]:e.target.value||undefined}))}><option value="">{f.required?'Seleccionar columna…':'No disponible'}</option>{columns.map(c=><option key={c} value={c}>{c}</option>)}</select></div>)}</div>{missingRequired.length>0&&<p className="warning">Faltan {missingRequired.length} campos obligatorios por mapear.</p>}</section>}
 {step===4&&<section><div className="sectionTitle"><div><span className="tag">CONTROL DE CALIDAD</span><h2>Validación previa</h2></div><button className="primary" onClick={()=>setStep(5)}>Construir modelo normalizado →</button></div><div className="validation"><Check ok={stats.missingSku===0} text={`${stats.missingSku} registros sin SKU`}/><Check ok={stats.missingSegment===0} text={`${stats.missingSegment} registros sin segmento`}/><Check ok={stats.duplicates===0} text={`${stats.duplicates} SKU duplicados`}/><Check ok={stats.invalidNumbers===0} text={`${stats.invalidNumbers} registros con ventas/unidades vacías o no numéricas`}/></div><Preview rows={normalized}/></section>}
 {step===5&&<section><div className="sectionTitle"><div><span className="tag success">NORMALIZADO</span><h2>{stats.rows.toLocaleString()} registros listos para el motor</h2><p>La fuente original permanece intacta. En esta versión, una celda vacía en ventas o unidades se interpreta como 0 para el análisis de surtido.</p></div><div className="actions"><button className="secondary" onClick={()=>setStep(3)}>← Ajustar mapeo</button><button className="primary" onClick={()=>setStep(6)}>Analizar surtido →</button></div></div><Preview rows={normalized}/></section>}
 {step===6&&<Analysis rows={filtered} allRows={analysis} counts={counts} filter={filter} setFilter={setFilter}/>} 
 </main></div>
}
function Analysis({rows,allRows,counts,filter,setFilter}:{rows:AnalysisRow[];allRows:AnalysisRow[];counts:Record<Recommendation,number>;filter:'TODOS'|Recommendation;setFilter:(v:'TODOS'|Recommendation)=>void}){
 const comparable=allRows.filter(r=>['MANTENER','REVISAR','DEPURAR'].includes(r.recommendation)).length;
 const exportExcel=()=>{
   const exportRows=allRows.map(r=>({
     SKU:r.sku??'', Descripción:r.description??'', Segmento:r.segment??'', Categoría:r.category??'', Fabricante:r.manufacturer??'', Marca:r.brand??'',
     'Venta período anterior':r.salesPrevious, 'Unidades período anterior':r.unitsPrevious, 'Venta período actual':r.salesCurrent, 'Unidades período actual':r.unitsCurrent,
     'Variación valor':r.varSales, 'Variación unidades':r.varUnits, 'Desempeño 70/30':r.performance, 'Peso segmento':r.segmentWeight, 'Pareto segmento':r.pareto,
     'Recomendación automática':r.recommendation, 'Explicación':r.reason, 'Recomendación manual':'', 'Razón / excepción manual':''
   }));
   const ws=XLSX.utils.json_to_sheet(exportRows);
   ['K','L','M','N','O'].forEach(c=>{for(let i=2;i<=exportRows.length+1;i++){const cell=ws[`${c}${i}`];if(cell)cell.z='0.0%'}});
   ws['!cols']=[{wch:14},{wch:46},{wch:24},{wch:22},{wch:28},{wch:20},{wch:20},{wch:22},{wch:20},{wch:22},{wch:16},{wch:18},{wch:18},{wch:16},{wch:16},{wch:25},{wch:80},{wch:25},{wch:45}];
   const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Análisis surtido');
   XLSX.writeFile(wb,`analisis_surtido_${new Date().toISOString().slice(0,10)}.xlsx`);
 };
 return <section><div className="sectionTitle"><div><span className="tag">MOTOR DE SURTIDO</span><h2>Recomendación automática</h2><p>{allRows.length} SKU analizados · {comparable} con historia comparable. La recomendación es cuantitativa y todavía no es la decisión final.</p></div><button className="primary" onClick={exportExcel}>Descargar análisis Excel ↓</button></div>
 <div className="summaryGrid"><Summary label="Mantener" value={counts.MANTENER} cls="keep"/><Summary label="Revisar" value={counts.REVISAR} cls="review"/><Summary label="Depurar" value={counts.DEPURAR} cls="remove"/><Summary label="Innovación" value={counts.INNOVACIÓN} cls="innovation"/><Summary label="Descodificado" value={counts.DESCODIFICADO} cls="muted"/><Summary label="Sin venta" value={counts['SIN VENTA']} cls="muted"/><Summary label="Datos insuficientes" value={counts['DATOS INSUFICIENTES']} cls="quality"/></div>
 <div className="method"><strong>Reglas activas:</strong> desempeño = 70% variación valor + 30% variación unidades · Pareto calculado dentro de cada segmento · Core ≤80% · Secundario 80–95% · Cola &gt;95% · crecimiento de protección en cola &gt;10%.</div>
 <div className="filters"><button className={filter==='TODOS'?'selected':''} onClick={()=>setFilter('TODOS')}>Todos <b>{allRows.length}</b></button>{recOrder.map(r=><button key={r} className={filter===r?'selected':''} onClick={()=>setFilter(r)}>{r} <b>{counts[r]}</b></button>)}</div>
 <div className="analysisTable"><table><thead><tr><th>SKU</th><th>Descripción</th><th>Segmento</th><th>Venta ant.</th><th>Venta act.</th><th>Var $</th><th>Var und.</th><th>70/30</th><th>Peso seg.</th><th>Pareto</th><th>Recomendación</th><th>Explicación</th></tr></thead><tbody>{rows.map((r,i)=><tr key={`${String(r.sku)}-${i}`}><td>{String(r.sku??'')}</td><td>{String(r.description??'')}</td><td>{String(r.segment??'')}</td><td className="number">{num(r.salesPrevious)}</td><td className="number">{num(r.salesCurrent)}</td><td className="number">{pct(r.varSales)}</td><td className="number">{pct(r.varUnits)}</td><td className="number"><strong>{pct(r.performance)}</strong></td><td className="number">{pct(r.segmentWeight)}</td><td className="number">{pct(r.pareto)}</td><td><span className={`pill ${r.recommendation.toLowerCase().replace(/ /g,'-').normalize('NFD').replace(/[\u0300-\u036f]/g,'')}`}>{r.recommendation}</span></td><td className="reason">{r.reason}</td></tr>)}</tbody></table></div>
 <div className="done">✓ V0.2.2: vacíos de ventas/unidades tratados como 0, descodificación corregida y exportación Excel habilitada. Siguiente validación: comparar el resultado descargado contra el ejercicio manual.</div></section>
}
function Summary({label,value,cls}:{label:string,value:number,cls:string}){return <div className={`summary ${cls}`}><b>{value}</b><span>{label}</span></div>}
function Check({ok,text}:{ok:boolean,text:string}){return <div className={'check '+(ok?'ok':'bad')}><b>{ok?'✓':'!'}</b><span>{text}</span></div>}
function Preview({rows}:{rows:Row[]}){if(!rows.length)return <div className="empty">No hay registros para mostrar.</div>;const cols=Object.keys(rows[0]).slice(0,10);return <div className="tableWrap"><table><thead><tr>{cols.map(c=><th key={c}>{c}</th>)}</tr></thead><tbody>{rows.slice(0,8).map((r,i)=><tr key={i}>{cols.map(c=><td key={c}>{String(r[c]??'')}</td>)}</tr>)}</tbody></table><div className="tableNote">Vista previa · primeras {Math.min(8,rows.length)} filas · máximo 10 columnas</div></div>}
ReactDOM.createRoot(document.getElementById('root')!).render(<React.StrictMode><App/></React.StrictMode>);
