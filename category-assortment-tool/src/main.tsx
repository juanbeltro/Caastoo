import React from 'react';
import ReactDOM from 'react-dom/client';
import * as XLSX from 'xlsx';
import './styles.css';

type Row = Record<string, unknown>;
type CanonicalKey = 'sku'|'description'|'segment'|'salesPrevious'|'unitsPrevious'|'salesCurrent'|'unitsCurrent'|'category'|'manufacturer'|'brand';
type NormalizedRow = Record<CanonicalKey, unknown> & { salesPrevious:number|null; unitsPrevious:number|null; salesCurrent:number|null; unitsCurrent:number|null };
type Recommendation = 'MANTENER'|'REVISAR'|'DEPURAR'|'INNOVACIÓN'|'DESCODIFICADO'|'SIN VENTA'|'DATOS INSUFICIENTES';
type AnalysisRow = NormalizedRow & { varSales:number|null; varUnits:number|null; performance:number|null; segmentWeight:number|null; pareto:number|null; recommendation:Recommendation; reason:string };
type AssortmentConfig = { count:number; targetA:number; reductions:number[]; minPerSegment:number };
type AssortmentSet = { name:string; rows:AnalysisRow[]; target:number|null; actualReduction:number|null; warning:string };

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
 const [assortmentConfig,setAssortmentConfig]=React.useState<AssortmentConfig>({count:3,targetA:150,reductions:[15,15,15,15],minPerSegment:3});
 const rows=sheets[sheet]||[]; const columns=rows.length?Object.keys(rows[0]):[];
 const onFile=async(file:File)=>{const data=await file.arrayBuffer(); const wb=XLSX.read(data,{type:'array'}); const parsed:Record<string,Row[]>={}; wb.SheetNames.forEach(n=>parsed[n]=XLSX.utils.sheet_to_json<Row>(wb.Sheets[n],{defval:null})); setFileName(file.name);setSheets(parsed);setSheet(wb.SheetNames[0]||'');setMapping({});setStep(2)};
 const normalized=React.useMemo<NormalizedRow[]>(()=>rows.map(r=>({sku:r[mapping.sku||'']??null,description:r[mapping.description||'']??null,segment:r[mapping.segment||'']??null,salesPrevious:cleanNumber(r[mapping.salesPrevious||'']),unitsPrevious:cleanNumber(r[mapping.unitsPrevious||'']),salesCurrent:cleanNumber(r[mapping.salesCurrent||'']),unitsCurrent:cleanNumber(r[mapping.unitsCurrent||'']),category:mapping.category?r[mapping.category]:null,manufacturer:mapping.manufacturer?r[mapping.manufacturer]:null,brand:mapping.brand?r[mapping.brand]:null})),[rows,mapping]);
 const analysis=React.useMemo(()=>buildAnalysis(normalized),[normalized]);
 const missingRequired=fields.filter(f=>f.required&&!mapping[f.key]);
 const skuCounts=new Map<string,number>(); normalized.forEach(r=>{const k=String(r.sku??'').trim();if(k)skuCounts.set(k,(skuCounts.get(k)||0)+1)});
 const stats={rows:normalized.length,missingSku:normalized.filter(r=>!String(r.sku??'').trim()).length,missingSegment:normalized.filter(r=>!String(r.segment??'').trim()).length,duplicates:[...skuCounts.values()].filter(n=>n>1).length,invalidNumbers:normalized.filter(r=>['salesPrevious','unitsPrevious','salesCurrent','unitsCurrent'].some(k=>(r as Record<string,unknown>)[k]===null)).length};
 const counts=Object.fromEntries(recOrder.map(x=>[x,analysis.filter(r=>r.recommendation===x).length])) as Record<Recommendation,number>;
 const filtered=filter==='TODOS'?analysis:analysis.filter(r=>r.recommendation===filter);
 return <div className="app"><header><div><span className="eyebrow">RAMO · CATEGORY MANAGEMENT</span><h1>Assortment Tool <b>V0.3.0</b></h1></div><div className="privacy">● Procesamiento local · sin base de datos</div></header>
 <nav>{['Cargar','Mapear','Validar','Normalizar','Analizar','Configurar','Surtidos'].map((x,i)=><button key={x} className={step===i+1?'active':''} disabled={i+1>step} onClick={()=>i+1<=step&&setStep(i+1)}>{i+1}. {x}</button>)}</nav>
 <main>
 {step===1&&<section className="hero"><span className="tag">SPRINT 2</span><h2>De una base plana a una recomendación explicable.</h2><p>Sube una tabla plana de Excel. El archivo se lee en tu navegador y no se envía a un servidor.</p><label className="drop"><strong>Seleccionar archivo Excel</strong><span>.xlsx · .xls</span><input type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0]&&onFile(e.target.files[0])}/></label></section>}
 {step===2&&<section><div className="sectionTitle"><div><span className="tag">FUENTE</span><h2>{fileName}</h2></div><button className="primary" disabled={!sheet||!rows.length} onClick={()=>setStep(3)}>Continuar al mapeo →</button></div><div className="card"><label>Hoja a analizar<select value={sheet} onChange={e=>{setSheet(e.target.value);setMapping({})}}>{Object.keys(sheets).map(s=><option key={s}>{s}</option>)}</select></label><div className="metrics"><div><b>{rows.length.toLocaleString()}</b><span>registros</span></div><div><b>{columns.length}</b><span>columnas</span></div><div><b>{Object.keys(sheets).length}</b><span>hojas</span></div></div></div><Preview rows={rows}/></section>}
 {step===3&&<section><div className="sectionTitle"><div><span className="tag">CONTRATO DE DATOS</span><h2>¿Qué significa cada columna?</h2><p>Mapea los campos del cliente a nuestro modelo estándar.</p></div><button className="primary" disabled={missingRequired.length>0} onClick={()=>setStep(4)}>Validar datos →</button></div><div className="mapping">{fields.map(f=><div className="mapRow" key={f.key}><div><strong>{f.label}</strong><span>{f.required?'Obligatorio':'Recomendado'}</span></div><span className="arrow">←</span><select value={mapping[f.key]||''} onChange={e=>setMapping(m=>({...m,[f.key]:e.target.value||undefined}))}><option value="">{f.required?'Seleccionar columna…':'No disponible'}</option>{columns.map(c=><option key={c} value={c}>{c}</option>)}</select></div>)}</div>{missingRequired.length>0&&<p className="warning">Faltan {missingRequired.length} campos obligatorios por mapear.</p>}</section>}
 {step===4&&<section><div className="sectionTitle"><div><span className="tag">CONTROL DE CALIDAD</span><h2>Validación previa</h2></div><button className="primary" onClick={()=>setStep(5)}>Construir modelo normalizado →</button></div><div className="validation"><Check ok={stats.missingSku===0} text={`${stats.missingSku} registros sin SKU`}/><Check ok={stats.missingSegment===0} text={`${stats.missingSegment} registros sin segmento`}/><Check ok={stats.duplicates===0} text={`${stats.duplicates} SKU duplicados`}/><Check ok={stats.invalidNumbers===0} text={`${stats.invalidNumbers} registros con ventas/unidades vacías o no numéricas`}/></div><Preview rows={normalized}/></section>}
 {step===5&&<section><div className="sectionTitle"><div><span className="tag success">NORMALIZADO</span><h2>{stats.rows.toLocaleString()} registros listos para el motor</h2><p>La fuente original permanece intacta. En esta versión, una celda vacía en ventas o unidades se interpreta como 0 para el análisis de surtido.</p></div><div className="actions"><button className="secondary" onClick={()=>setStep(3)}>← Ajustar mapeo</button><button className="primary" onClick={()=>setStep(6)}>Analizar surtido →</button></div></div><Preview rows={normalized}/></section>}
 {step===6&&<Analysis rows={filtered} allRows={analysis} counts={counts} filter={filter} setFilter={setFilter} onNext={()=>setStep(7)}/>}
 {step===7&&<AssortmentSetup analysis={analysis} config={assortmentConfig} setConfig={setAssortmentConfig} onNext={()=>setStep(8)}/>}
 {step===8&&<Assortments analysis={analysis} config={assortmentConfig}/>} 
 </main></div>
}
function Analysis({rows,allRows,counts,filter,setFilter,onNext}:{rows:AnalysisRow[];allRows:AnalysisRow[];counts:Record<Recommendation,number>;filter:'TODOS'|Recommendation;setFilter:(v:'TODOS'|Recommendation)=>void;onNext:()=>void}){
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
 return <section><div className="sectionTitle"><div><span className="tag">MOTOR DE SURTIDO</span><h2>Recomendación automática</h2><p>{allRows.length} SKU analizados · {comparable} con historia comparable. La recomendación es cuantitativa y todavía no es la decisión final.</p></div><div className="actions"><button className="secondary" onClick={exportExcel}>Descargar análisis Excel ↓</button><button className="primary" onClick={onNext}>Construir surtidos →</button></div></div>
 <div className="summaryGrid"><Summary label="Mantener" value={counts.MANTENER} cls="keep"/><Summary label="Revisar" value={counts.REVISAR} cls="review"/><Summary label="Depurar" value={counts.DEPURAR} cls="remove"/><Summary label="Innovación" value={counts.INNOVACIÓN} cls="innovation"/><Summary label="Descodificado" value={counts.DESCODIFICADO} cls="muted"/><Summary label="Sin venta" value={counts['SIN VENTA']} cls="muted"/><Summary label="Datos insuficientes" value={counts['DATOS INSUFICIENTES']} cls="quality"/></div>
 <div className="method"><strong>Reglas activas:</strong> desempeño = 70% variación valor + 30% variación unidades · Pareto calculado dentro de cada segmento · Core ≤80% · Secundario 80–95% · Cola &gt;95% · crecimiento de protección en cola &gt;10%.</div>
 <div className="filters"><button className={filter==='TODOS'?'selected':''} onClick={()=>setFilter('TODOS')}>Todos <b>{allRows.length}</b></button>{recOrder.map(r=><button key={r} className={filter===r?'selected':''} onClick={()=>setFilter(r)}>{r} <b>{counts[r]}</b></button>)}</div>
 <div className="analysisTable"><table><thead><tr><th>SKU</th><th>Descripción</th><th>Segmento</th><th>Venta ant.</th><th>Venta act.</th><th>Var $</th><th>Var und.</th><th>70/30</th><th>Peso seg.</th><th>Pareto</th><th>Recomendación</th><th>Explicación</th></tr></thead><tbody>{rows.map((r,i)=><tr key={`${String(r.sku)}-${i}`}><td>{String(r.sku??'')}</td><td>{String(r.description??'')}</td><td>{String(r.segment??'')}</td><td className="number">{num(r.salesPrevious)}</td><td className="number">{num(r.salesCurrent)}</td><td className="number">{pct(r.varSales)}</td><td className="number">{pct(r.varUnits)}</td><td className="number"><strong>{pct(r.performance)}</strong></td><td className="number">{pct(r.segmentWeight)}</td><td className="number">{pct(r.pareto)}</td><td><span className={`pill ${r.recommendation.toLowerCase().replace(/ /g,'-').normalize('NFD').replace(/[\u0300-\u036f]/g,'')}`}>{r.recommendation}</span></td><td className="reason">{r.reason}</td></tr>)}</tbody></table></div>
 <div className="done">✓ Motor base validado. Continúa para definir cuántos surtidos necesitas y cuánto debe reducirse cada transición.</div></section>
}

function AssortmentSetup({analysis,config,setConfig,onNext}:{analysis:AnalysisRow[];config:AssortmentConfig;setConfig:(c:AssortmentConfig)=>void;onNext:()=>void}){
 const baseCount=analysis.filter(r=>['MANTENER','REVISAR','INNOVACIÓN'].includes(r.recommendation)).length;
 const updateReduction=(i:number,v:number)=>{const reductions=[...config.reductions];reductions[i]=Math.max(0,Math.min(80,v||0));setConfig({...config,reductions})};
 let projected=baseCount;
 return <section><div className="sectionTitle"><div><span className="tag">CONFIGURACIÓN DEL EJERCICIO</span><h2>Define la profundidad de los surtidos</h2><p>Tú decides cuánto debe reducirse cada transición. Caastoo decide qué referencias son los mejores candidatos para salir.</p></div><button className="primary" onClick={onNext}>Generar surtidos →</button></div>
 <div className="configGrid"><div className="card configCard"><label>¿Cuántos surtidos necesitas?<select value={config.count} onChange={e=>setConfig({...config,count:Number(e.target.value)})}>{[1,2,3,4,5].map(n=><option key={n} value={n}>{n}</option>)}</select></label><label>Mínimo de SKU por segmento<input type="number" min="1" max="20" value={config.minPerSegment} onChange={e=>setConfig({...config,minPerSegment:Math.max(1,Number(e.target.value)||1)})}/><small>Protección global. El motor no reducirá un segmento por debajo de este número.</small></label></div>
 <div className="card configCard"><h3>Tipo A</h3><label>SKU objetivo aproximado<input type="number" min="1" value={config.targetA} onChange={e=>setConfig({...config,targetA:Math.max(1,Number(e.target.value)||1)})}/></label><div className="projection"><b>{baseCount}</b><span>SKU recomendados por el motor base</span></div><p className="micro">El objetivo de A es orientativo. Caastoo no elimina referencias solo para alcanzar ese número.</p></div></div>
 <div className="transitionList">{Array.from({length:config.count-1},(_,i)=>{const from=String.fromCharCode(65+i),to=String.fromCharCode(66+i);const before=projected;projected=Math.max(config.minPerSegment,Math.round(projected*(1-config.reductions[i]/100)));return <div className="transition" key={i}><div><span>TIPO {from} → TIPO {to}</span><strong>¿Cuánto quieres reducir?</strong></div><div className="reductionInput"><input type="number" min="0" max="80" value={config.reductions[i]} onChange={e=>updateReduction(i,Number(e.target.value))}/><b>%</b></div><div className="projection"><b>~{projected}</b><span>SKU orientativos desde {before}</span></div></div>})}</div>
 <div className="method"><strong>Regla:</strong> el porcentaje lo defines tú. El sistema intentará acercarse a la meta sin bajar del mínimo por segmento, priorizando REVISAR y después los SKU menos importantes del surtido. INNOVACIÓN queda protegida.</div></section>
}

function buildAssortments(analysis:AnalysisRow[],config:AssortmentConfig):AssortmentSet[]{
 const base=analysis.filter(r=>['MANTENER','REVISAR','INNOVACIÓN'].includes(r.recommendation));
 const sets:AssortmentSet[]=[{name:'A',rows:base,target:config.targetA,actualReduction:null,warning:Math.abs(base.length-config.targetA)>0?`${base.length} SKU recomendados vs objetivo aproximado ${config.targetA}. El Tipo A no se fuerza.`:''}];
 let current=[...base];
 for(let i=0;i<config.count-1;i++){
   const reduction=Math.max(0,Math.min(.8,(config.reductions[i]||0)/100));
   const desiredRemove=Math.round(current.length*reduction);
   const segmentCounts=new Map<string,number>(); current.forEach(r=>{const s=String(r.segment??'SIN SEGMENTO');segmentCounts.set(s,(segmentCounts.get(s)||0)+1)});
   const removable=current.filter(r=>r.recommendation!=='INNOVACIÓN').map(r=>{
     const s=String(r.segment??'SIN SEGMENTO'); const count=segmentCounts.get(s)||0;
     const saturation=count/Math.max(1,current.length);
     const recScore=r.recommendation==='REVISAR'?100:r.recommendation==='MANTENER'?0:-100;
     const paretoScore=(r.pareto??0)*35;
     const perfScore=-(r.performance??0)*15;
     const weightScore=-(r.segmentWeight??0)*25;
     const sizeScore=saturation*30;
     return {r,score:recScore+paretoScore+perfScore+weightScore+sizeScore};
   }).sort((a,b)=>b.score-a.score);
   const removed=new Set<AnalysisRow>();
   for(const c of removable){
     if(removed.size>=desiredRemove) break;
     const s=String(c.r.segment??'SIN SEGMENTO');
     const remaining=(segmentCounts.get(s)||0)-[...removed].filter(x=>String(x.segment??'SIN SEGMENTO')===s).length;
     if(remaining<=config.minPerSegment) continue;
     removed.add(c.r);
   }
   const next=current.filter(r=>!removed.has(r));
   const actual=current.length?removed.size/current.length:0;
   const targetCount=current.length-desiredRemove;
   const warning=removed.size<desiredRemove?`Meta: retirar ${desiredRemove} SKU (${(reduction*100).toFixed(1)}%). Caastoo encontró ${removed.size} candidatos sin romper las protecciones.`:'';
   sets.push({name:String.fromCharCode(66+i),rows:next,target:targetCount,actualReduction:actual,warning});
   current=next;
 }
 return sets;
}

function Assortments({analysis,config}:{analysis:AnalysisRow[];config:AssortmentConfig}){
 const sets=React.useMemo(()=>buildAssortments(analysis,config),[analysis,config]); const [active,setActive]=React.useState(0);
 const current=sets[active]; const prev=active>0?sets[active-1]:null; const prevSku=new Set((prev?.rows||[]).map(r=>String(r.sku)));
 const removed=prev?prev.rows.filter(r=>!new Set(current.rows.map(x=>String(x.sku))).has(String(r.sku))):[];
 const exportSets=()=>{const wb=XLSX.utils.book_new();sets.forEach((set,i)=>{const previous=i>0?sets[i-1]:null;const curSku=new Set(set.rows.map(r=>String(r.sku)));const rows=set.rows.map(r=>({SKU:r.sku,Descripción:r.description,Segmento:r.segment,Fabricante:r.manufacturer,Marca:r.brand,'Recomendación base':r.recommendation,'Pareto segmento':r.pareto,'Desempeño 70/30':r.performance}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),`Tipo ${set.name}`);if(previous){const outs=previous.rows.filter(r=>!curSku.has(String(r.sku))).map(r=>({SKU:r.sku,Descripción:r.description,Segmento:r.segment,'Sale de':`Tipo ${previous.name}`,'No está en':`Tipo ${set.name}`,'Recomendación base':r.recommendation,'Pareto segmento':r.pareto,'Desempeño 70/30':r.performance}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(outs),`${previous.name} vs ${set.name}`)}});XLSX.writeFile(wb,`surtidos_${new Date().toISOString().slice(0,10)}.xlsx`)};
 return <section><div className="sectionTitle"><div><span className="tag success">SURTIDOS GENERADOS</span><h2>Arquitectura escalonada A → {sets[sets.length-1].name}</h2><p>Cada surtido inferior es subconjunto del anterior. Las reducciones respetan el porcentaje solicitado como meta y las protecciones definidas.</p></div><button className="primary" onClick={exportSets}>Descargar surtidos Excel ↓</button></div>
 <div className="assortmentTabs">{sets.map((s,i)=><button className={active===i?'selected':''} onClick={()=>setActive(i)} key={s.name}><span>TIPO {s.name}</span><b>{s.rows.length} SKU</b>{s.actualReduction!==null&&<small>-{(s.actualReduction*100).toFixed(1)}% vs {sets[i-1].name}</small>}</button>)}</div>
 {current.warning&&<div className="warningBox">⚠ {current.warning}</div>}
 <div className="metrics assortmentMetrics"><div><b>{current.rows.length}</b><span>SKU Tipo {current.name}</span></div><div><b>{new Set(current.rows.map(r=>String(r.segment))).size}</b><span>segmentos</span></div><div><b>{removed.length}</b><span>salen vs {prev?`Tipo ${prev.name}`:'base'}</span></div></div>
 {prev&&<><h3 className="subhead">Referencias que salen de Tipo {prev.name} → Tipo {current.name}</h3><div className="analysisTable"><table><thead><tr><th>SKU</th><th>Descripción</th><th>Segmento</th><th>Recomendación base</th><th>Pareto</th><th>70/30</th></tr></thead><tbody>{removed.map((r,i)=><tr key={i}><td>{String(r.sku??'')}</td><td>{String(r.description??'')}</td><td>{String(r.segment??'')}</td><td>{r.recommendation}</td><td>{pct(r.pareto)}</td><td>{pct(r.performance)}</td></tr>)}</tbody></table></div></>}
 <h3 className="subhead">Surtido Tipo {current.name}</h3><div className="analysisTable"><table><thead><tr><th>SKU</th><th>Descripción</th><th>Segmento</th><th>Recomendación base</th><th>Pareto</th><th>70/30</th></tr></thead><tbody>{current.rows.map((r,i)=><tr key={i}><td>{String(r.sku??'')}</td><td>{String(r.description??'')}</td><td>{String(r.segment??'')}</td><td>{r.recommendation}</td><td>{pct(r.pareto)}</td><td>{pct(r.performance)}</td></tr>)}</tbody></table></div>
 <div className="done">✓ V0.3.0: generador de surtidos escalonados activo. El porcentaje de reducción de cada transición es definido por el usuario.</div></section>
}

function Summary({label,value,cls}:{label:string,value:number,cls:string}){return <div className={`summary ${cls}`}><b>{value}</b><span>{label}</span></div>}
function Check({ok,text}:{ok:boolean,text:string}){return <div className={'check '+(ok?'ok':'bad')}><b>{ok?'✓':'!'}</b><span>{text}</span></div>}
function Preview({rows}:{rows:Row[]}){if(!rows.length)return <div className="empty">No hay registros para mostrar.</div>;const cols=Object.keys(rows[0]).slice(0,10);return <div className="tableWrap"><table><thead><tr>{cols.map(c=><th key={c}>{c}</th>)}</tr></thead><tbody>{rows.slice(0,8).map((r,i)=><tr key={i}>{cols.map(c=><td key={c}>{String(r[c]??'')}</td>)}</tr>)}</tbody></table><div className="tableNote">Vista previa · primeras {Math.min(8,rows.length)} filas · máximo 10 columnas</div></div>}
ReactDOM.createRoot(document.getElementById('root')!).render(<React.StrictMode><App/></React.StrictMode>);
