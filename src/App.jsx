import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";

const SHEETS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRJwCHYoCrnaMu-SXGZoD-1N2rx7tl192B1vEKhrmCPvbBQvyK-79hBsOLkRDjwD-YEX2P0mB8VQFRy/pub?gid=752561413&single=true&output=csv";

const REQUIRED_FIELDS = [
  { key:"bodega",       label:"Bodega / Punto de venta" },
  { key:"articulo",    label:"Código artículo (SKU)" },
  { key:"descripcion", label:"Descripción" },
  { key:"abc_empresa", label:"ABC Empresa" },
  { key:"abc_bodega",  label:"ABC Bodega (NEW365)" },
  { key:"stock",       label:"Stock bodega" },
  { key:"stock_cd",    label:"Stock en CD (PM122)" },
  { key:"transito",    label:"Tránsito" },
  { key:"consumo",     label:"Consumo mensual" },
  { key:"m1",  label:"Mes 1 (más reciente)" },
  { key:"m2",  label:"Mes 2" }, { key:"m3",  label:"Mes 3" },
  { key:"m4",  label:"Mes 4" }, { key:"m5",  label:"Mes 5" },
  { key:"m6",  label:"Mes 6" }, { key:"m7",  label:"Mes 7" },
  { key:"m8",  label:"Mes 8" }, { key:"m9",  label:"Mes 9" },
  { key:"m10", label:"Mes 10" },{ key:"m11", label:"Mes 11" },
  { key:"m12", label:"Mes 12 (más antiguo)" },
  { key:"alerta_lote", label:"Alerta Lote", optional:true },
];

const BAJA_ROTACION = ["D","E","F","G","P","O","Z"];
const ACTIVOS = ["A00","A","B","C","N"];

const TIPO_CONFIG = {
  CRITICO:           { label:"Crítico — sin stock",      color:"#A32D2D", bg:"#FCEBEB" },
  REPOSICION:        { label:"Reposición desde CD",       color:"#185FA5", bg:"#E6F1FB" },
  LOGISTICA_INVERSA: { label:"Logística inversa → CD",   color:"#854F0B", bg:"#FAEEDA" },
  SOBRESTOCK:        { label:"Sobrestock → devolver CD", color:"#3B6D11", bg:"#EAF3DE" },
  OK:                { label:"En rango óptimo",           color:"#5F5E5A", bg:"#F1EFE8" },
};

const SS_A00 = (c) => 2.33 * c * Math.sqrt(0.25);

const calcMinimo = (abc, c) => {
  if (BAJA_ROTACION.includes(abc)) return 0;
  if (abc === "A00") return Math.ceil(c * 1.5 + SS_A00(c));
  return Math.ceil(c * 1.5);
};

const calcMaximo = (abc, c) => {
  if (BAJA_ROTACION.includes(abc)) return 0;
  const base = Math.ceil(c * 2);
  // A00: máximo también incluye SS para que siempre sea >= mínimo
  if (abc === "A00") return Math.max(base, calcMinimo(abc, c));
  return base;
};

const calcTipo = r => {
  const { abc_bodega, abc_empresa, posicion, minimo, maximo } = r;
  if (abc_bodega === "A00" && posicion === 0) return "CRITICO";
  // Disparo: posición < mínimo → reposición
  if (ACTIVOS.includes(abc_bodega) && posicion < minimo) return "REPOSICION";
  if (BAJA_ROTACION.includes(abc_bodega) && ["A","B","C"].includes(abc_empresa) && posicion > 0) return "LOGISTICA_INVERSA";
  if (ACTIVOS.includes(abc_bodega) && maximo > 0 && posicion > maximo) return "SOBRESTOCK";
  return "OK";
};

const processRaw = (rows, mapping) => rows.map(row => {
  const g  = k => { const c = mapping[k]; return c ? row[c] : ""; };
  const gn = k => parseFloat(g(k)) || 0;
  const abc_bodega  = String(g("abc_bodega")  || "").trim().toUpperCase();
  const abc_empresa = String(g("abc_empresa") || "").trim().toUpperCase();
  const consumo  = gn("consumo");
  const stock    = gn("stock");
  const stock_cd = gn("stock_cd");
  const transito = gn("transito");
  const posicion = stock + transito;
  const minimo   = calcMinimo(abc_bodega, consumo);
  const maximo   = calcMaximo(abc_bodega, consumo);
  const meses    = [1,2,3,4,5,6,7,8,9,10,11,12].map(i => gn(`m${i}`));
  const alerta_lote = String(g("alerta_lote")||"").trim();
  const base = { bodega:String(g("bodega")||""), articulo:String(g("articulo")||""),
    descripcion:String(g("descripcion")||""), abc_empresa, abc_bodega,
    stock, stock_cd, transito, consumo, posicion, minimo, maximo, meses };
  const tipo     = calcTipo(base);
  // Sugerido: llegar al máximo desde la posición actual
  const sugerido = (tipo==="REPOSICION"||tipo==="CRITICO") ? Math.max(0, maximo - posicion) : 0;
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
  const [colFilters, setColFilters] = useState({});
  const [search, setSearch]     = useState("");
  const [sortCol, setSortCol]   = useState("tipo");
  const [sortDir, setSortDir]   = useState("asc");
  const [dragOver, setDragOver] = useState(false);
  const [view, setView]         = useState("resumen");
  const [selected, setSelected] = useState(null);
  const [copied, setCopied]     = useState(false);
  const [selectedCells, setSelectedCells] = useState({}); // {rowIdx: {col: true}}
  const [filterOpen, setFilterOpen] = useState(null);

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
    const hints = {
      bodega:["bodega","punto","agencia","nuevabodega"],
      articulo:["articulo","sku","codigo","nuevoarticulo"],
      descripcion:["desc","nombre","producto"],
      abc_empresa:["abcemp","abc_emp","abcempresa"],
      abc_bodega:["abcnew","abc_new","abc365","abcbodega","abcn"],
      stock:["stock_bod","stockbod","existenciatotal","stockbodega"],
      stock_cd:["pm122","stockcd","stock_cd"],
      transito:["transito","transit","trans"],
      consumo:["consumo","mensual"],
      m1:["1"],m2:["2"],m3:["3"],m4:["4"],m5:["5"],m6:["6"],
      m7:["7"],m8:["8"],m9:["9"],m10:["10"],m11:["11"],m12:["12"],
      alerta_lote:["alertalote","alerta_lote","alertalot","lote"],
    };
    const auto = {};
    REQUIRED_FIELDS.forEach(({ key }) => {
      const kws = hints[key] || [key];
      const match = headers.find(h => {
        const hn = h.toLowerCase().replace(/[^a-z0-9]/g,"");
        return kws.some(kw => { const kn=kw.replace(/[^a-z0-9]/g,""); return hn===kn||hn.includes(kn); });
      });
      if (match) auto[key] = match;
    });
    setMapping(auto);
  }, [headers]);

  const applyMapping = () => { SM(mapping); setData(processRaw(rawData, mapping)); setStep("dashboard"); };

  const handleSort = (col) => {
    if (sortCol === col) setSortDir(d => d==="asc"?"desc":"asc");
    else { setSortCol(col); setSortDir("asc"); }
  };

  // Unique values for column filters
  const colValues = useMemo(() => ({
    bodega:    [...new Set(data.map(r=>r.bodega))].sort(),
    abc_bodega:[...new Set(data.map(r=>r.abc_bodega))].sort(),
    tipo:      [...new Set(data.map(r=>r.tipo))],
  }), [data]);

  const bodegas = useMemo(() => ["todas",...new Set(data.map(r=>r.bodega))], [data]);
  const abcs    = useMemo(() => ["todos",...new Set(data.map(r=>r.abc_bodega))], [data]);

  const filtered = useMemo(() => {
    const ord = ["CRITICO","REPOSICION","LOGISTICA_INVERSA","SOBRESTOCK","OK"];
    return data.filter(r => {
      if (filters.bodega!=="todas" && r.bodega!==filters.bodega) return false;
      if (filters.tipo!=="todos"   && r.tipo!==filters.tipo)     return false;
      if (filters.abc!=="todos"    && r.abc_bodega!==filters.abc) return false;
      if (colFilters.bodega?.length && !colFilters.bodega.includes(r.bodega)) return false;
      if (colFilters.abc_bodega?.length && !colFilters.abc_bodega.includes(r.abc_bodega)) return false;
      if (colFilters.tipo?.length && !colFilters.tipo.includes(r.tipo)) return false;
      if (search) { const q=search.toLowerCase(); return r.articulo.toLowerCase().includes(q)||r.descripcion.toLowerCase().includes(q); }
      return true;
    }).sort((a,b) => {
      let va, vb;
      if (sortCol==="tipo")     { va=ord.indexOf(a.tipo); vb=ord.indexOf(b.tipo); }
      else if (sortCol==="sugerido") { va=a.sugerido; vb=b.sugerido; }
      else if (sortCol==="consumo")  { va=a.consumo;  vb=b.consumo; }
      else if (sortCol==="posicion") { va=a.posicion; vb=b.posicion; }
      else if (sortCol==="articulo") { va=a.articulo; vb=b.articulo; }
      else { va=0; vb=0; }
      if (va < vb) return sortDir==="asc"?-1:1;
      if (va > vb) return sortDir==="asc"?1:-1;
      return 0;
    });
  }, [data, filters, colFilters, search, sortCol, sortDir]);

  const metrics = useMemo(() => ({
    criticos:  data.filter(r=>r.tipo==="CRITICO").length,
    reposicion:data.filter(r=>r.tipo==="REPOSICION").length,
    logistica: data.filter(r=>r.tipo==="LOGISTICA_INVERSA").length,
    sobrestock:data.filter(r=>r.tipo==="SOBRESTOCK").length,
    total: data.length,
  }), [data]);

  // Cell selection for SAP copy
  const toggleCell = (rowIdx, col) => {
    setSelectedCells(prev => {
      const next = { ...prev };
      if (!next[rowIdx]) next[rowIdx] = {};
      if (next[rowIdx][col]) { delete next[rowIdx][col]; if (!Object.keys(next[rowIdx]).length) delete next[rowIdx]; }
      else next[rowIdx][col] = true;
      return next;
    });
  };

  const copyForSAP = () => {
    const rows = [];
    Object.keys(selectedCells).sort((a,b)=>+a-+b).forEach(idx => {
      const r = filtered[+idx];
      if (!r) return;
      const cols = Object.keys(selectedCells[idx]);
      const vals = cols.map(c => {
        if (c==="articulo")  return r.articulo;
        if (c==="sugerido")  return r.sugerido;
        if (c==="consumo")   return r.consumo;
        if (c==="minimo")    return r.minimo;
        if (c==="maximo")    return r.maximo;
        if (c==="posicion")  return r.posicion;
        return "";
      });
      rows.push(vals.join("\t"));
    });
    if (!rows.length) return;
    navigator.clipboard.writeText(rows.join("\n")).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

  const clearSelection = () => setSelectedCells({});
  const selCount = Object.keys(selectedCells).length;
  const allMapped = REQUIRED_FIELDS.filter(f => !f.optional).every(f => mapping[f.key]);
  const S = { fontFamily:"system-ui", fontSize:13 };

  const SortIcon = ({ col }) => sortCol===col
    ? <span style={{marginLeft:3,fontSize:9}}>{sortDir==="asc"?"▲":"▼"}</span>
    : <span style={{marginLeft:3,fontSize:9,color:"#D3D1C7"}}>⇅</span>;

  const FilterDropdown = ({ col, vals }) => {
    const active = colFilters[col] || [];
    const toggle = v => {
      setColFilters(prev => {
        const cur = prev[col] || [];
        const next = cur.includes(v) ? cur.filter(x=>x!==v) : [...cur, v];
        return { ...prev, [col]: next.length ? next : undefined };
      });
    };
    return (
      <div style={{position:"absolute",top:"100%",left:0,background:"white",border:"0.5px solid #D3D1C7",borderRadius:8,padding:"8px",zIndex:100,minWidth:180,boxShadow:"0 4px 12px rgba(0,0,0,0.1)",maxHeight:240,overflowY:"auto"}}>
        <div style={{fontSize:11,color:"#888780",marginBottom:6}}>Filtrar por</div>
        {active.length>0&&<div onClick={()=>setColFilters(p=>({...p,[col]:undefined}))} style={{fontSize:11,color:"#A32D2D",cursor:"pointer",marginBottom:6}}>Limpiar filtro ×</div>}
        {vals.map(v=>(
          <label key={v} style={{display:"flex",alignItems:"center",gap:6,padding:"3px 0",fontSize:12,cursor:"pointer"}}>
            <input type="checkbox" checked={active.includes(v)} onChange={()=>toggle(v)} style={{margin:0}}/>
            <span>{v}</span>
          </label>
        ))}
      </div>
    );
  };

  // ── HOME ───────────────────────────────────────────────────────────────────
  if (step==="home") return (
    <div style={{minHeight:"100vh",background:"#F8F7F4",display:"flex",alignItems:"center",justifyContent:"center",padding:"2rem",...S}}>
      <div style={{maxWidth:500,width:"100%",textAlign:"center"}}>
        <div style={{fontSize:11,letterSpacing:"0.2em",color:"#888780",textTransform:"uppercase",marginBottom:10,fontFamily:"Georgia,serif"}}>Agente de Inventario</div>
        <h1 style={{fontSize:"2rem",fontWeight:400,color:"#2C2C2A",marginBottom:8,fontFamily:"Georgia,serif"}}>Nivelación & Sugeridos</h1>
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

  // ── MAP ────────────────────────────────────────────────────────────────────
  if (step==="map") return (
    <div style={{minHeight:"100vh",background:"#F8F7F4",padding:"2rem 1.5rem",...S}}>
      <div style={{maxWidth:700,margin:"0 auto"}}>
        <div style={{fontSize:11,letterSpacing:"0.2em",color:"#888780",textTransform:"uppercase",marginBottom:6,fontFamily:"Georgia,serif"}}>Mapeo de columnas</div>
        <h2 style={{fontSize:"1.3rem",fontWeight:500,color:"#2C2C2A",marginBottom:4,fontFamily:"Georgia,serif"}}>Asocia cada campo con tu columna</h2>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <span style={{fontSize:13,color:"#888780"}}>{rawData.length.toLocaleString()} filas detectadas</span>
          <span style={{fontSize:13,color:"#185FA5",cursor:"pointer",textDecoration:"underline"}} onClick={autoMap}>Auto-detectar →</span>
        </div>
        <div style={{background:"white",border:"0.5px solid #D3D1C7",borderRadius:12,padding:"1rem 1.25rem",marginBottom:20,maxHeight:"65vh",overflowY:"auto"}}>
          {REQUIRED_FIELDS.map(({key,label})=>(
            <div key={key} style={{display:"flex",alignItems:"center",gap:12,padding:"6px 0",borderBottom:"0.5px solid #F1EFE8"}}>
              <div style={{minWidth:210,fontSize:12,color:"#444441"}}>{label}</div>
              <select value={mapping[key]||""} onChange={e=>setMapping(m=>({...m,[key]:e.target.value}))}
                style={{flex:1,fontSize:12,padding:"4px 8px",borderRadius:6,border:"0.5px solid #D3D1C7",background:"white",color:"#2C2C2A"}}>
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

  // ── MODAL DETALLE ──────────────────────────────────────────────────────────
  const DetailModal = ({ r, onClose }) => {
    const maxV = Math.max(...r.meses, 1);
    return (
      <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:"1rem"}} onClick={onClose}>
        <div style={{background:"white",borderRadius:16,width:"100%",maxWidth:900,maxHeight:"90vh",overflowY:"auto",padding:"1.5rem"}} onClick={e=>e.stopPropagation()}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
            <div>
              <div style={{fontSize:11,color:"#888780",marginBottom:2}}>{r.bodega}</div>
              <div style={{fontSize:16,fontWeight:600,color:"#2C2C2A"}}>{r.articulo}</div>
              <div style={{fontSize:13,color:"#444441"}}>{r.descripcion}</div>
            </div>
            <button onClick={onClose} style={{fontSize:20,border:"none",background:"none",cursor:"pointer",color:"#888780"}}>✕</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:16}}>
            {[
              {label:"Stock bodega",    val:r.stock,    color:"#2C2C2A"},
              {label:"Stock CD",        val:r.stock_cd, color:"#185FA5"},
              {label:"Tránsito",        val:r.transito, color:"#888780"},
              {label:"Posición",        val:r.posicion, color:r.posicion<r.minimo?"#A32D2D":"#2C2C2A"},
              {label:"Consumo mensual", val:r.consumo,  color:"#2C2C2A"},
              {label:"Mínimo",         val:r.minimo,   color:"#854F0B"},
              {label:"Máximo",         val:r.maximo,   color:"#3B6D11"},
              {label:"Sugerido",       val:r.sugerido||"—", color:r.sugerido>0?"#185FA5":"#888780"},
            ].map(m=>(
              <div key={m.label} style={{background:"#F8F7F4",borderRadius:8,padding:"0.7rem 0.8rem"}}>
                <div style={{fontSize:11,color:"#888780",marginBottom:3}}>{m.label}</div>
                <div style={{fontSize:18,fontWeight:500,color:m.color}}>{m.val}</div>
              </div>
            ))}
          </div>
          <div style={{background:TIPO_CONFIG[r.tipo].bg,borderRadius:8,padding:"0.8rem 1rem",marginBottom:16,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <span style={{fontSize:13,fontWeight:600,color:TIPO_CONFIG[r.tipo].color}}>Recomendación:</span>
            <span style={{fontSize:13,color:TIPO_CONFIG[r.tipo].color}}>{TIPO_CONFIG[r.tipo].label}</span>
            {r.sugerido>0&&<span style={{marginLeft:"auto",fontSize:13,fontWeight:600,color:TIPO_CONFIG[r.tipo].color}}>Cantidad sugerida: {r.sugerido} u.</span>}
          </div>
          <div>
            <div style={{fontSize:11,color:"#888780",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:10}}>Ventas últimos 12 meses (M1 = más reciente)</div>
            <div style={{display:"flex",gap:4,alignItems:"flex-end",height:90}}>
              {r.meses.map((v,i)=>(
                <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
                  <div style={{fontSize:10,color:v>0?"#2C2C2A":"#D3D1C7",minHeight:14}}>{v>0?v:""}</div>
                  <div style={{width:"100%",background:v>0?"#185FA5":"#F1EFE8",borderRadius:"3px 3px 0 0",height:`${Math.max(4,(v/maxV)*60)}px`}}/>
                  <div style={{fontSize:10,color:"#888780"}}>{i+1}</div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  };

  // ── DASHBOARD ──────────────────────────────────────────────────────────────
  const thStyle = (col, hasFilter=false) => ({
    padding:"8px 10px", textAlign:"left", fontSize:11, color:"#888780",
    fontWeight:500, borderBottom:"0.5px solid #D3D1C7", whiteSpace:"nowrap",
    cursor:"pointer", userSelect:"none", position:"relative",
    background: (colFilters[col]?.length || hasFilter) ? "#EAF3DE" : "#F8F7F4",
  });

  return (
    <div style={{minHeight:"100vh",background:"#F8F7F4",...S}} onClick={()=>setFilterOpen(null)}>
      {selected && <DetailModal r={selected} onClose={()=>setSelected(null)} />}

      {/* Header */}
      <div style={{background:"#2C2C2A",padding:"0.8rem 1.5rem",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <span style={{fontSize:11,letterSpacing:"0.15em",color:"#888780",textTransform:"uppercase",fontFamily:"Georgia,serif"}}>Agente Inventario</span>
          <span style={{color:"#444441"}}>·</span>
          <span style={{color:"#D3D1C7",fontSize:12}}>{metrics.total.toLocaleString()} SKUs</span>
          {lastSync&&<span style={{color:"#888780",fontSize:11}}>· {lastSync.toLocaleTimeString()}</span>}
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {selCount>0&&(
            <button onClick={copyForSAP} style={{fontSize:12,padding:"4px 14px",borderRadius:6,border:"none",background:copied?"#1D9E75":"#185FA5",color:"white",cursor:"pointer",fontWeight:500}}>
              {copied?`✓ Copiado`:`📋 Copiar ${selCount} fila(s) para SAP`}
            </button>
          )}
          {selCount>0&&<button onClick={clearSelection} style={{fontSize:12,padding:"4px 10px",borderRadius:6,border:"0.5px solid #444441",background:"transparent",color:"#888780",cursor:"pointer"}}>✕</button>}
          <button onClick={loadFromSheets} disabled={loading} style={{fontSize:12,padding:"4px 14px",borderRadius:6,border:"0.5px solid #444441",background:"transparent",color:loading?"#888780":"#D3D1C7",cursor:loading?"wait":"pointer"}}>{loading?"Sincronizando…":"↻ Sincronizar"}</button>
          <button onClick={()=>setView(v=>v==="resumen"?"detalle":"resumen")} style={{fontSize:12,padding:"4px 12px",borderRadius:6,border:"0.5px solid #444441",background:view==="detalle"?"#444441":"transparent",color:"#D3D1C7",cursor:"pointer"}}>{view==="resumen"?"Vista detalle":"Vista resumen"}</button>
          <button onClick={()=>setStep("map")} style={{fontSize:12,padding:"4px 12px",borderRadius:6,border:"0.5px solid #444441",background:"transparent",color:"#888780",cursor:"pointer"}}>Columnas</button>
          <button onClick={()=>setStep("home")} style={{fontSize:12,padding:"4px 12px",borderRadius:6,border:"0.5px solid #444441",background:"transparent",color:"#888780",cursor:"pointer"}}>Inicio</button>
        </div>
      </div>

      <div style={{padding:"1.25rem 1.5rem"}}>
        {error&&<div style={{fontSize:12,color:"#A32D2D",background:"#FCEBEB",padding:"8px 12px",borderRadius:8,marginBottom:12}}>{error}</div>}

        {/* Metric cards */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:10,marginBottom:18}}>
          {[
            {label:"Críticos sin stock",   val:metrics.criticos,   color:"#A32D2D",bg:"#FCEBEB",f:"CRITICO"},
            {label:"Reposición pendiente", val:metrics.reposicion, color:"#185FA5",bg:"#E6F1FB",f:"REPOSICION"},
            {label:"Logística inversa",    val:metrics.logistica,  color:"#854F0B",bg:"#FAEEDA",f:"LOGISTICA_INVERSA"},
            {label:"Sobrestock",           val:metrics.sobrestock, color:"#3B6D11",bg:"#EAF3DE",f:"SOBRESTOCK"},
          ].map(m=>(
            <div key={m.f} onClick={()=>setFilters(f=>({...f,tipo:f.tipo===m.f?"todos":m.f}))}
              style={{background:filters.tipo===m.f?m.bg:"white",border:`0.5px solid ${filters.tipo===m.f?m.color+"55":"#D3D1C7"}`,borderRadius:10,padding:"0.85rem 1rem",cursor:"pointer",transition:"all 0.15s"}}>
              <div style={{fontSize:11,color:"#888780",marginBottom:4}}>{m.label}</div>
              <div style={{fontSize:22,fontWeight:500,color:m.color}}>{m.val}</div>
            </div>
          ))}
        </div>

        {/* Filters bar */}
        <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
          <input placeholder="Buscar SKU o descripción…" value={search} onChange={e=>setSearch(e.target.value)}
            style={{flex:1,minWidth:180,fontSize:13,padding:"6px 12px",borderRadius:8,border:"0.5px solid #D3D1C7",background:"white"}}/>
          <select value={filters.bodega} onChange={e=>setFilters(f=>({...f,bodega:e.target.value}))}
            style={{fontSize:13,padding:"6px 10px",borderRadius:8,border:"0.5px solid #D3D1C7",background:"white"}}>
            {bodegas.map(b=><option key={b} value={b}>{b==="todas"?"Todas las bodegas":b}</option>)}
          </select>
          <select value={filters.abc} onChange={e=>setFilters(f=>({...f,abc:e.target.value}))}
            style={{fontSize:13,padding:"6px 10px",borderRadius:8,border:"0.5px solid #D3D1C7",background:"white"}}>
            {abcs.map(a=><option key={a} value={a}>{a==="todos"?"Todos los ABC":`ABC: ${a}`}</option>)}
          </select>
          <select value={sortCol+":"+sortDir} onChange={e=>{const[c,d]=e.target.value.split(":");setSortCol(c);setSortDir(d);}}
            style={{fontSize:13,padding:"6px 10px",borderRadius:8,border:"0.5px solid #D3D1C7",background:"white"}}>
            <option value="tipo:asc">Por prioridad</option>
            <option value="sugerido:desc">Mayor sugerido primero</option>
            <option value="consumo:desc">Mayor consumo primero</option>
            <option value="articulo:asc">SKU A→Z</option>
          </select>
        </div>

        <div style={{fontSize:11,color:"#888780",marginBottom:10,display:"flex",alignItems:"center",gap:12}}>
          <span>{filtered.length.toLocaleString()} de {metrics.total.toLocaleString()} registros</span>
          {selCount>0&&<span style={{color:"#185FA5"}}>· {selCount} fila(s) seleccionada(s) — haz clic en las celdas azules para copiar a SAP</span>}
          {!selCount&&<span style={{color:"#888780"}}>· Haz clic en SKU o Sugerido para seleccionar · Clic en fila para detalle</span>}
          {Object.values(colFilters).some(v=>v?.length)&&(
            <span onClick={()=>setColFilters({})} style={{color:"#A32D2D",cursor:"pointer"}}>Limpiar filtros de columna ×</span>
          )}
        </div>

        {/* TABLE RESUMEN */}
        {view==="resumen" && <div style={{background:"white",border:"0.5px solid #D3D1C7",borderRadius:12,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr>
                  {/* Bodega header with filter */}
                  <th style={thStyle("bodega")} onClick={e=>{e.stopPropagation();setFilterOpen(f=>f==="bodega"?null:"bodega");}}>
                    Bodega {colFilters.bodega?.length?`(${colFilters.bodega.length})`:"▾"}
                    {filterOpen==="bodega"&&<FilterDropdown col="bodega" vals={colValues.bodega}/>}
                  </th>
                  <th style={thStyle("articulo")} onClick={()=>handleSort("articulo")}>SKU <SortIcon col="articulo"/></th>
                  <th style={{...thStyle(),cursor:"default"}}>Descripción</th>
                  <th style={thStyle("abc_bodega")} onClick={e=>{e.stopPropagation();setFilterOpen(f=>f==="abc_bodega"?null:"abc_bodega");}}>
                    ABC {colFilters.abc_bodega?.length?`(${colFilters.abc_bodega.length})`:"▾"}
                    {filterOpen==="abc_bodega"&&<FilterDropdown col="abc_bodega" vals={colValues.abc_bodega}/>}
                  </th>
                  <th style={thStyle("consumo")} onClick={()=>handleSort("consumo")}>Consumo <SortIcon col="consumo"/></th>
                  <th style={{...thStyle(),cursor:"default",textAlign:"right"}}>Stock</th>
                  <th style={{...thStyle(),cursor:"default",textAlign:"right"}}>Tránsito</th>
                  <th style={thStyle("posicion")} onClick={()=>handleSort("posicion")}>Posición <SortIcon col="posicion"/></th>
                  <th style={{...thStyle(),cursor:"default",textAlign:"right"}}>Mínimo</th>
                  <th style={{...thStyle(),cursor:"default",textAlign:"right"}}>Máximo</th>
                  <th style={thStyle("sugerido")} onClick={()=>handleSort("sugerido")}>Sugerido <SortIcon col="sugerido"/></th>
                  <th style={thStyle("tipo")} onClick={e=>{e.stopPropagation();setFilterOpen(f=>f==="tipo"?null:"tipo");}}>
                    Acción {colFilters.tipo?.length?`(${colFilters.tipo.length})`:"▾"}
                    {filterOpen==="tipo"&&<FilterDropdown col="tipo" vals={colValues.tipo}/>}
                  </th>
                </tr>
              </thead>
              <tbody>
                {filtered.slice(0,300).map((r,i)=>{
                  const tc = TIPO_CONFIG[r.tipo];
                  const rowSel = selectedCells[i] || {};
                  const artSel = rowSel["articulo"];
                  const sugSel = rowSel["sugerido"];
                  return (
                    <tr key={i}
                      style={{borderBottom:"0.5px solid #F1EFE8",background:r.tipo==="CRITICO"?"#FFF8F8":"white"}}
                      onMouseEnter={e=>{if(!Object.keys(rowSel).length)e.currentTarget.style.background="#F8F7F4";}}
                      onMouseLeave={e=>{e.currentTarget.style.background=r.tipo==="CRITICO"?"#FFF8F8":"white";}}>
                      <td style={{padding:"7px 10px",maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} onClick={()=>setSelected(r)}>{r.bodega}</td>
                      {/* Selectable: articulo */}
                      <td onClick={e=>{e.stopPropagation();toggleCell(i,"articulo");}}
                        style={{padding:"7px 10px",fontFamily:"monospace",whiteSpace:"nowrap",cursor:"pointer",background:artSel?"#DBEAFE":"transparent",borderRadius:4,outline:artSel?"2px solid #185FA5":"none",userSelect:"none"}}>
                        {r.articulo}
                      </td>
                      <td style={{padding:"7px 10px",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#444441",cursor:"pointer"}} title={r.descripcion} onClick={()=>setSelected(r)}>{r.descripcion}</td>
                      <td style={{padding:"7px 10px",cursor:"pointer"}} onClick={()=>setSelected(r)}>
                        <span style={{background:r.abc_bodega==="A00"?"#FCEBEB":r.abc_bodega==="A"?"#EAF3DE":"#F1EFE8",color:r.abc_bodega==="A00"?"#A32D2D":r.abc_bodega==="A"?"#3B6D11":"#5F5E5A",padding:"1px 7px",borderRadius:4,fontSize:11,fontWeight:500}}>{r.abc_bodega}</span>
                      </td>
                      <td style={{padding:"7px 10px",textAlign:"right",cursor:"pointer"}} onClick={()=>setSelected(r)}>{r.consumo}</td>
                      <td style={{padding:"7px 10px",textAlign:"right",cursor:"pointer"}} onClick={()=>setSelected(r)}>{r.stock}</td>
                      <td style={{padding:"7px 10px",textAlign:"right",color:"#888780",cursor:"pointer"}} onClick={()=>setSelected(r)}>{r.transito}</td>
                      <td style={{padding:"7px 10px",textAlign:"right",fontWeight:500,color:r.posicion<r.minimo?"#A32D2D":"#2C2C2A",cursor:"pointer"}} onClick={()=>setSelected(r)}>{r.posicion}</td>
                      <td style={{padding:"7px 10px",textAlign:"right",color:"#888780",cursor:"pointer"}} onClick={()=>setSelected(r)}>{r.minimo}</td>
                      <td style={{padding:"7px 10px",textAlign:"right",color:"#888780",cursor:"pointer"}} onClick={()=>setSelected(r)}>{r.maximo}</td>
                      {/* Selectable: sugerido */}
                      <td onClick={e=>{e.stopPropagation();if(r.sugerido>0)toggleCell(i,"sugerido");}}
                        style={{padding:"7px 10px",textAlign:"right",fontWeight:r.sugerido>0?600:400,color:r.sugerido>0?"#185FA5":"#888780",cursor:r.sugerido>0?"pointer":"default",background:sugSel?"#DBEAFE":"transparent",borderRadius:4,outline:sugSel?"2px solid #185FA5":"none",userSelect:"none"}}>
                        {r.sugerido||"—"}
                      </td>
                      <td style={{padding:"7px 10px",whiteSpace:"nowrap",cursor:"pointer"}} onClick={()=>setSelected(r)}>
                        <span style={{background:tc.bg,color:tc.color,padding:"2px 8px",borderRadius:4,fontSize:11,fontWeight:500}}>{tc.label}</span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {filtered.length>300&&<div style={{padding:"10px 16px",fontSize:12,color:"#888780",borderTop:"0.5px solid #F1EFE8",textAlign:"center"}}>Mostrando primeros 300 — usa filtros para acotar.</div>}
        </div>}

        {/* TABLE DETALLE con 12 meses */}
        {view==="detalle" && <div style={{background:"white",border:"0.5px solid #D3D1C7",borderRadius:12,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead><tr style={{background:"#F8F7F4"}}>
              {["SKU","Descripción","ABC","Alerta Lote","Stock Bodega","Stock CD","Tránsito","Consumo","Mínimo","Máximo","Sugerido","M1","M2","M3","M4","M5","M6","M7","M8","M9","M10","M11","M12","Recomendación"].map((h,i)=>(
                  <th key={i} style={{padding:"6px 8px",textAlign:i>=11&&i<=23?"right":"left",fontSize:10,color:"#888780",fontWeight:500,borderBottom:"0.5px solid #D3D1C7",whiteSpace:"pre-wrap",lineHeight:1.3}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {filtered.slice(0,200).map((r,i)=>{
                  const tc=TIPO_CONFIG[r.tipo];
                  const maxM=Math.max(...r.meses,1);
                  return <tr key={i} onClick={()=>setSelected(r)} style={{borderBottom:"0.5px solid #F1EFE8",background:r.tipo==="CRITICO"?"#FFF8F8":"white",cursor:"pointer"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#F8F7F4"}
                    onMouseLeave={e=>e.currentTarget.style.background=r.tipo==="CRITICO"?"#FFF8F8":"white"}>
                    <td style={{padding:"5px 8px",fontFamily:"monospace",whiteSpace:"nowrap",fontSize:11}}>{r.articulo}</td>
                    <td style={{padding:"5px 8px",maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#444441"}} title={r.descripcion}>{r.descripcion}</td>
                    <td style={{padding:"5px 8px"}}><span style={{background:r.abc_bodega==="A00"?"#FCEBEB":r.abc_bodega==="A"?"#EAF3DE":"#F1EFE8",color:r.abc_bodega==="A00"?"#A32D2D":r.abc_bodega==="A"?"#3B6D11":"#5F5E5A",padding:"1px 6px",borderRadius:3,fontSize:10,fontWeight:500}}>{r.abc_bodega}</span></td>
                    <td style={{padding:"5px 8px"}}>{r.alerta_lote?<span style={{background:"#FAEEDA",color:"#854F0B",padding:"1px 6px",borderRadius:3,fontSize:10,fontWeight:500}}>{r.alerta_lote}</span>:""}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",fontWeight:500}}>{r.stock}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",color:"#185FA5"}}>{r.stock_cd}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",color:"#888780"}}>{r.transito}</td>
                    <td style={{padding:"5px 8px",textAlign:"right"}}>{r.consumo}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",color:"#854F0B"}}>{r.minimo}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",color:"#3B6D11"}}>{r.maximo}</td>
                    <td style={{padding:"5px 8px",textAlign:"right",fontWeight:r.sugerido>0?600:400,color:r.sugerido>0?"#185FA5":"#888780"}}>{r.sugerido||"—"}</td>
                    {r.meses.map((v,mi)=>(
                      <td key={mi} style={{padding:"4px 6px",textAlign:"right"}}>
                        <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:1}}>
                          <span style={{fontSize:10,color:v>0?"#2C2C2A":"#D3D1C7"}}>{v||"—"}</span>
                          <div style={{width:24,height:3,background:"#F1EFE8",borderRadius:2}}>
                            <div style={{width:`${(v/maxM)*100}%`,height:"100%",background:"#185FA5",borderRadius:2}}/>
                          </div>
                        </div>
                      </td>
                    ))}
                    <td style={{padding:"5px 8px",whiteSpace:"nowrap"}}><span style={{background:tc.bg,color:tc.color,padding:"2px 6px",borderRadius:3,fontSize:10,fontWeight:500}}>{tc.label}</span></td>
                  </tr>;
                })}
              </tbody>
            </table>
          </div>
          {filtered.length>200&&<div style={{padding:"10px 16px",fontSize:12,color:"#888780",borderTop:"0.5px solid #F1EFE8",textAlign:"center"}}>Mostrando primeros 200 — usa filtros para acotar.</div>}
        </div>}
      </div>
    </div>
  );
}
