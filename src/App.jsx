import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";

const SHEETS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTb7Qz8RZweOyi84QWkdlSz-f5H9_3XFFEukasuVq0e7N94dETrhg5U-rwhM16E-pL7TG9oNMXsG0q6/pub?gid=301328833&single=true&output=csv";

const REQUIRED_FIELDS = [
  { key: "bodega",       label: "Bodega / Punto de venta" },
  { key: "articulo",    label: "Código artículo (SKU)" },
  { key: "descripcion", label: "Descripción" },
  { key: "abc_empresa", label: "ABC Empresa" },
  { key: "abc_bodega",  label: "ABC Bodega (NEW365)" },
  { key: "stock",       label: "Stock bodega" },
  { key: "transito",    label: "Tránsito" },
  { key: "consumo",     label: "Consumo mensual" },
];

const BAJA_ROTACION = ["D","E","F","G","P","O","Z"];
const ACTIVOS       = ["A00","A","B","C","N"];
const LT            = 0.25;

const TIPO_CONFIG = {
  CRITICO:           { label: "Crítico — sin stock",      color: "#A32D2D", bg: "#FCEBEB" },
  REPOSICION:        { label: "Reposición desde CD",       color: "#185FA5", bg: "#E6F1FB" },
  LOGISTICA_INVERSA: { label: "Logística inversa → CD",   color: "#854F0B", bg: "#FAEEDA" },
  SOBRESTOCK:        { label: "Sobrestock → devolver CD", color: "#3B6D11", bg: "#EAF3DE" },
  OK:                { label: "En rango óptimo",           color: "#5F5E5A", bg: "#F1EFE8" },
};

const calcMinimo = (abc, c) => {
  if (BAJA_ROTACION.includes(abc)) return 0;
  if (abc === "A00") return Math.ceil(c * 1.5 + 2.33 * c * Math.sqrt(LT));
  return Math.ceil(c * 1.5);
};
const calcMaximo  = (abc, c) => BAJA_ROTACION.includes(abc) ? 0 : Math.ceil(c * 2);
const calcDisparo = (min, c) => Math.ceil(min + c * LT);
const calcTipo = r => {
  if (r.abc_bodega === "A00" && r.posicion === 0) return "CRITICO";
  if (ACTIVOS.includes(r.abc_bodega) && r.posicion < r.punto_disparo) return "REPOSICION";
  if (BAJA_ROTACION.includes(r.abc_bodega) && ["A","B","C"].includes(r.abc_empresa) && r.posicion > 0) return "LOGISTICA_INVERSA";
  if (ACTIVOS.includes(r.abc_bodega) && r.posicion > r.maximo) return "SOBRESTOCK";
  return "OK";
};

const processRaw = (rows, mapping) => rows.map(row => {
  const g = k => { const c = mapping[k]; return c ? row[c] : ""; };
  const abc_bodega  = String(g("abc_bodega")  || "").trim().toUpperCase();
  const abc_empresa = String(g("abc_empresa") || "").trim().toUpperCase();
  const consumo  = parseFloat(g("consumo"))  || 0;
  const stock    = parseFloat(g("stock"))    || 0;
  const transito = parseFloat(g("transito")) || 0;
  const posicion = stock + transito;
  const minimo        = calcMinimo(abc_bodega, consumo);
  const maximo        = calcMaximo(abc_bodega, consumo);
  const punto_disparo = calcDisparo(minimo, consumo);
  const base = { bodega: String(g("bodega")||""), articulo: String(g("articulo")||""),
    descripcion: String(g("descripcion")||""), abc_empresa, abc_bodega,
    stock, transito, consumo, posicion, minimo, maximo, punto_disparo };
  const tipo     = calcTipo(base);
  const sugerido = (tipo==="REPOSICION"||tipo==="CRITICO") ? Math.max(0, maximo-posicion) : 0;
  return { ...base, tipo, sugerido };
});

const SM = m => { try { localStorage.setItem("inv_map", JSON.stringify(m)); } catch {} };
const LM = () => { try { return JSON.parse(localStorage.getItem("inv_map")||"null"); } catch { return null; } };

export default function App() {
  const [step, setStep]         = useState("home");
  const [mapping, setMapping]   = useState(() => LM() || {});
  const [headers, setHeaders]   = useState([]);
  const [rawData, setRawData]   = useState([]);
  const [data, setData]         = useState([]);
  const [loading, setLoading]   = useState(false);
  const [error, setError]       = useState("");
  const [lastSync, setLastSync] = useState(null);
  const [filters, setFilters]   = useState({ bodega:"todas", tipo:"todos", abc:"todos" });
  const [search, setSearch]     = useState("");
  const [sortCol, setSortCol]   = useState("tipo");
  const [dragOver, setDragOver] = useState(false);

  const loadRows = useCallback((rows) => {
    setHeaders(Object.keys(rows[0]));
    setRawData(rows);
    setLastSync(new Date());
    const saved = LM();
    if (saved && REQUIRED_FIELDS.every(f => saved[f.key])) {
      setData(processRaw(rows, saved));
      setMapping(saved);
      setStep("dashboard");
    } else { setStep("map"); }
  }, []);

  const loadFromSheets = useCallback(async () => {
    setLoading(true); setError("");
    try {
      const res = await fetch(SHEETS_URL);
      if (!res.ok) throw new Error(`Error al descargar (${res.status})`);
      const csv  = await res.text();
      const wb   = XLSX.read(csv, { type:"string" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval:0 });
      if (!rows.length) throw new Error("La hoja está vacía.");
      loadRows(rows);
    } catch(e) { setError(e.message); }
    finally { setLoading(false); }
  }, [loadRows]);

  const handleFile = useCallback(file => {
    const reader = new FileReader();
    reader.onload = e => {
      const wb   = XLSX.read(e.target.result, { type:"array" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval:0 });
      loadRows(rows);
    };
    reader.readAsArrayBuffer(file);
  }, [loadRows]);

  const autoMap = useCallback(() => {
    const hints = { bodega:["bodega","punto","agencia","nuevabodega"], articulo:["articulo","sku","codigo","nuevoarticulo"], descripcion:["desc","nombre","producto"], abc_empresa:["abcemp","abc_emp","abcempresa"], abc_bodega:["abcnew","abc_new","abc365","abcbodega"], stock:["stock_bod","stockbod","existencia","stock"], transito:["transito","transit","trans"], consumo:["consumo","demand","mensual"] };
    const auto = {};
    REQUIRED_FIELDS.forEach(({ key }) => {
      const kws = hints[key] || [key];
      const match = headers.find(h => kws.some(kw => h.toLowerCase().replace(/[^a-z0-9]/g,"").includes(kw.replace(/[^a-z0-9]/g,""))));
      if (match) auto[key] = match;
    });
    setMapping(auto);
  }, [headers]);

  const applyMapping = () => { SM(mapping); setData(processRaw(rawData, mapping)); setStep("dashboard"); };

  const bodegas  = useMemo(() => ["todas", ...new Set(data.map(r => r.bodega))], [data]);
  const abcs     = useMemo(() => ["todos", ...new Set(data.map(r => r.abc_bodega))], [data]);
  const filtered = useMemo(() => data.filter(r => {
    if (filters.bodega!=="todas" && r.bodega!==filters.bodega) return false;
    if (filters.tipo!=="todos"   && r.tipo!==filters.tipo)     return false;
    if (filters.abc!=="todos"    && r.abc_bodega!==filters.abc) return false;
    if (search) { const q=search.toLowerCase(); return r.articulo.toLowerCase().includes(q)||r.descripcion.toLowerCase().includes(q); }
    return true;
  }).sort((a,b) => {
    const ord = ["CRITICO","REPOSICION","LOGISTICA_INVERSA","SOBRESTOCK","OK"];
    if (sortCol==="tipo") return ord.indexOf(a.tipo)-ord.indexOf(b.tipo);
    if (sortCol==="sugerido") return b.sugerido-a.sugerido;
    return 0;
  }), [data, filters, search, sortCol]);

  const metrics = useMemo(() => ({
    criticos:  data.filter(r=>r.tipo==="CRITICO").length,
    reposicion:data.filter(r=>r.tipo==="REPOSICION").length,
    logistica: data.filter(r=>r.tipo==="LOGISTICA_INVERSA").length,
    sobrestock:data.filter(r=>r.tipo==="SOBRESTOCK").length,
    total: data.length,
  }), [data]);

  const allMapped = REQUIRED_FIELDS.every(f => mapping[f.key]);
  const S = { fontFamily:"system-ui", fontSize:13 };

  if (step==="home") return (
    <div style={{minHeight:"100vh",background:"#F8F7F4",display:"flex",alignItems:"center",justifyContent:"center",padding:"2rem",...S}}>
      <div style={{maxWidth:500,width:"100%",textAlign:"center"}}>
        <div style={{fontSize:11,letterSpacing:"0.2em",color:"#888780",textTransform:"uppercase",marginBottom:10,fontFamily:"Georgia,serif"}}>Agente de Inventario</div>
        <h1 style={{fontSize:"2rem",fontWeight:400,color:"#2C2C2A",marginBottom:8,fontFamily:"Georgia,serif",letterSpacing:"-0.02em"}}>Nivelación & Sugeridos</h1>
        <p style={{fontSize:13,color:"#888780",marginBottom:36,lineHeight:1.7}}>Conecta con tu archivo en Google Sheets o carga manualmente.</p>
        {error&&<div style={{fontSize:12,color:"#A32D2D",background:"#FCEBEB",padding:"10px 14px",borderRadius:8,marginBottom:20,textAlign:"left",lineHeight:1.5}}>{error}</div>}
        <button onClick={loadFromSheets} disabled={loading} style={{width:"100%",padding:"14px",borderRadius:10,border:"none",background:loading?"#D3D1C7":"#2C2C2A",color:"white",fontSize:14,fontWeight:500,cursor:loading?"wait":"pointer",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"center",gap:10}}>
          <span>☁️</span>{loading?"Cargando desde Google Sheets…":"Sincronizar desde Google Sheets"}
        </button>
        <div style={{fontSize:12,color:"#B4B2A9",margin:"12px 0",display:"flex",alignItems:"center",gap:8}}>
          <div style={{flex:1,height:"0.5px",background:"#D3D1C7"}}/><span>o carga manualmente</span><div style={{flex:1,height:"0.5px",background:"#D3D1C7"}}/>
        </div>
        <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
          onDrop={e=>{e.preventDefault();setDragOver(false);if(e.dataTransfer.files[0])handleFile(e.dataTransfer.files[0]);}}
          onClick={()=>document.getElementById("fi").click()}
          style={{border:`1.5px dashed ${dragOver?"#2C2C2A":"#D3D1C7"}`,borderRadius:10,padding:"1.5rem",textAlign:"center",cursor:"pointer",background:dragOver?"#F1EFE8":"white"}}>
          <div style={{fontSize:24,marginBottom:6}}>📂</div>
          <div style={{fontSize:13,color:"#444441"}}>Arrastra tu archivo Excel aquí</div>
          <div style={{fontSize:11,color:"#888780",marginTop:4}}>.xlsx / .xls</div>
        </div>
        <input id="fi" type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>{if(e.target.files[0])handleFile(e.target.files[0]);}}/>
        {lastSync&&<div style={{fontSize:11,color:"#888780",marginTop:16}}>Última sincronización: {lastSync.toLocaleString()}</div>}
      </div>
    </div>
  );

  if (step==="map") return (
    <div style={{minHeight:"100vh",background:"#F8F7F4",padding:"2rem 1.5rem",...S}}>
      <div style={{maxWidth:660,margin:"0 auto"}}>
        <div style={{fontSize:11,letterSpacing:"0.2em",color:"#888780",textTransform:"uppercase",marginBottom:6,fontFamily:"Georgia,serif"}}>Mapeo de columnas</div>
        <h2 style={{fontSize:"1.3rem",fontWeight:500,color:"#2C2C2A",marginBottom:4,fontFamily:"Georgia,serif"}}>Asocia cada campo con tu columna</h2>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <span style={{fontSize:13,color:"#888780"}}>{rawData.length.toLocaleString()} filas detectadas</span>
          <span style={{fontSize:13,color:"#185FA5",cursor:"pointer",textDecoration:"underline"}} onClick={autoMap}>Auto-detectar →</span>
        </div>
        <div style={{background:"white",border:"0.5px solid #D3D1C7",borderRadius:12,padding:"1rem 1.25rem",marginBottom:20}}>
          {REQUIRED_FIELDS.map(({key,label})=>(
            <div key={key} style={{display:"flex",alignItems:"center",gap:12,padding:"8px 0",borderBottom:"0.5px solid #F1EFE8"}}>
              <div style={{minWidth:200,fontSize:13,color:"#444441"}}>{label}</div>
              <select value={mapping[key]||""} onChange={e=>setMapping(m=>({...m,[key]:e.target.value}))}
                style={{flex:1,fontSize:13,padding:"5px 10px",borderRadius:6,border:"0.5px solid #D3D1C7",background:"white",color:"#2C2C2A"}}>
                <option value="">— seleccionar —</option>
                {headers.map(h=><option key={h} value={h}>{h}</option>)}
              </select>
              <span style={{color:mapping[key]?"#3B6D11":"#D3D1C7",fontSize:16,minWidth:16}}>{mapping[key]?"✓":"○"}</span>
            </div>
          ))}
        </div>
        <div style={{display:"flex",gap:10}}>
          <button onClick={()=>setStep("home")} style={{padding:"8px 20px",borderRadius:8,border:"0.5px solid #D3D1C7",background:"white",fontSize:13,cursor:"pointer",color:"#444441"}}>← Volver</button>
          <button onClick={applyMapping} disabled={!allMapped} style={{padding:"8px 24px",borderRadius:8,border:"none",background:!allMapped?"#D3D1C7":"#2C2C2A",color:"white",fontSize:13,fontWeight:500,cursor:!allMapped?"not-allowed":"pointer"}}>Ver dashboard →</button>
        </div>
      </div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:"#F8F7F4",...S}}>
      <div style={{background:"#2C2C2A",padding:"0.8rem 1.5rem",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <span style={{fontSize:11,letterSpacing:"0.15em",color:"#888780",textTransform:"uppercase",fontFamily:"Georgia,serif"}}>Agente Inventario</span>
          <span style={{color:"#444441"}}>·</span>
          <span style={{color:"#D3D1C7",fontSize:12}}>{metrics.total.toLocaleString()} SKUs</span>
          {lastSync&&<span style={{color:"#888780",fontSize:11}}>· {lastSync.toLocaleTimeString()}</span>}
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={loadFromSheets} disabled={loading} style={{fontSize:12,padding:"4px 14px",borderRadius:6,border:"0.5px solid #444441",background:"transparent",color:loading?"#888780":"#D3D1C7",cursor:loading?"wait":"pointer"}}>{loading?"Sincronizando…":"↻ Sincronizar"}</button>
          <button onClick={()=>setStep("map")} style={{fontSize:12,padding:"4px 12px",borderRadius:6,border:"0.5px solid #444441",background:"transparent",color:"#888780",cursor:"pointer"}}>Columnas</button>
          <button onClick={()=>setStep("home")} style={{fontSize:12,padding:"4px 12px",borderRadius:6,border:"0.5px solid #444441",background:"transparent",color:"#888780",cursor:"pointer"}}>Inicio</button>
        </div>
      </div>
      <div style={{padding:"1.25rem 1.5rem"}}>
        {error&&<div style={{fontSize:12,color:"#A32D2D",background:"#FCEBEB",padding:"8px 12px",borderRadius:8,marginBottom:12}}>{error}</div>}
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:10,marginBottom:18}}>
          {[{label:"Críticos sin stock",val:metrics.criticos,color:"#A32D2D",bg:"#FCEBEB",f:"CRITICO"},{label:"Reposición pendiente",val:metrics.reposicion,color:"#185FA5",bg:"#E6F1FB",f:"REPOSICION"},{label:"Logística inversa",val:metrics.logistica,color:"#854F0B",bg:"#FAEEDA",f:"LOGISTICA_INVERSA"},{label:"Sobrestock",val:metrics.sobrestock,color:"#3B6D11",bg:"#EAF3DE",f:"SOBRESTOCK"}].map(m=>(
            <div key={m.f} onClick={()=>setFilters(f=>({...f,tipo:f.tipo===m.f?"todos":m.f}))} style={{background:filters.tipo===m.f?m.bg:"white",border:`0.5px solid ${filters.tipo===m.f?m.color+"55":"#D3D1C7"}`,borderRadius:10,padding:"0.85rem 1rem",cursor:"pointer",transition:"all 0.15s"}}>
              <div style={{fontSize:11,color:"#888780",marginBottom:4}}>{m.label}</div>
              <div style={{fontSize:22,fontWeight:500,color:m.color}}>{m.val}</div>
            </div>
          ))}
        </div>
        <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
          <input placeholder="Buscar SKU o descripción…" value={search} onChange={e=>setSearch(e.target.value)} style={{flex:1,minWidth:180,fontSize:13,padding:"6px 12px",borderRadius:8,border:"0.5px solid #D3D1C7",background:"white"}}/>
          <select value={filters.bodega} onChange={e=>setFilters(f=>({...f,bodega:e.target.value}))} style={{fontSize:13,padding:"6px 10px",borderRadius:8,border:"0.5px solid #D3D1C7",background:"white"}}>{bodegas.map(b=><option key={b} value={b}>{b==="todas"?"Todas las bodegas":b}</option>)}</select>
          <select value={filters.abc} onChange={e=>setFilters(f=>({...f,abc:e.target.value}))} style={{fontSize:13,padding:"6px 10px",borderRadius:8,border:"0.5px solid #D3D1C7",background:"white"}}>{abcs.map(a=><option key={a} value={a}>{a==="todos"?"Todos los ABC":`ABC: ${a}`}</option>)}</select>
          <select value={sortCol} onChange={e=>setSortCol(e.target.value)} style={{fontSize:13,padding:"6px 10px",borderRadius:8,border:"0.5px solid #D3D1C7",background:"white"}}><option value="tipo">Por prioridad</option><option value="sugerido">Por sugerido</option></select>
        </div>
        <div style={{fontSize:11,color:"#888780",marginBottom:10}}>{filtered.length.toLocaleString()} de {metrics.total.toLocaleString()} registros</div>
        <div style={{background:"white",border:"0.5px solid #D3D1C7",borderRadius:12,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{background:"#F8F7F4"}}>{["Bodega","SKU","Descripción","ABC","Consumo","Stock","Tránsito","Posición","Mínimo","Máximo","Sugerido","Acción"].map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",fontSize:11,color:"#888780",fontWeight:500,borderBottom:"0.5px solid #D3D1C7",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
              <tbody>
                {filtered.slice(0,200).map((r,i)=>{
                  const tc=TIPO_CONFIG[r.tipo];
                  return <tr key={i} style={{borderBottom:"0.5px solid #F1EFE8",background:r.tipo==="CRITICO"?"#FFF8F8":"white"}}>
                    <td style={{padding:"7px 10px",maxWidth:130,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.bodega}</td>
                    <td style={{padding:"7px 10px",fontFamily:"monospace",whiteSpace:"nowrap"}}>{r.articulo}</td>
                    <td style={{padding:"7px 10px",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#444441"}} title={r.descripcion}>{r.descripcion}</td>
                    <td style={{padding:"7px 10px"}}><span style={{background:r.abc_bodega==="A00"?"#FCEBEB":r.abc_bodega==="A"?"#EAF3DE":"#F1EFE8",color:r.abc_bodega==="A00"?"#A32D2D":r.abc_bodega==="A"?"#3B6D11":"#5F5E5A",padding:"1px 7px",borderRadius:4,fontSize:11,fontWeight:500}}>{r.abc_bodega}</span></td>
                    <td style={{padding:"7px 10px",textAlign:"right"}}>{r.consumo}</td>
                    <td style={{padding:"7px 10px",textAlign:"right"}}>{r.stock}</td>
                    <td style={{padding:"7px 10px",textAlign:"right",color:"#888780"}}>{r.transito}</td>
                    <td style={{padding:"7px 10px",textAlign:"right",fontWeight:500,color:r.posicion<r.minimo?"#A32D2D":"#2C2C2A"}}>{r.posicion}</td>
                    <td style={{padding:"7px 10px",textAlign:"right",color:"#888780"}}>{r.minimo}</td>
                    <td style={{padding:"7px 10px",textAlign:"right",color:"#888780"}}>{r.maximo}</td>
                    <td style={{padding:"7px 10px",textAlign:"right",fontWeight:r.sugerido>0?600:400,color:r.sugerido>0?"#185FA5":"#888780"}}>{r.sugerido||"—"}</td>
                    <td style={{padding:"7px 10px",whiteSpace:"nowrap"}}><span style={{background:tc.bg,color:tc.color,padding:"2px 8px",borderRadius:4,fontSize:11,fontWeight:500}}>{tc.label}</span></td>
                  </tr>;
                })}
              </tbody>
            </table>
          </div>
          {filtered.length>200&&<div style={{padding:"10px 16px",fontSize:12,color:"#888780",borderTop:"0.5px solid #F1EFE8",textAlign:"center"}}>Mostrando primeros 200 — usa filtros para acotar.</div>}
        </div>
      </div>
    </div>
  );
}
