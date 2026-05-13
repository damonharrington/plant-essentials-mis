import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid,
  ComposedChart, ReferenceLine,
} from "recharts";

// ─── localStorage adapter (mirrors window.storage API used in Claude artifacts) ───
const storage = {
  get: (key) => {
    try {
      const value = localStorage.getItem(key);
      return value != null ? { key, value } : null;
    } catch { return null; }
  },
  set: (key, value) => {
    try {
      localStorage.setItem(key, value);
      return { key, value };
    } catch { return null; }
  },
  delete: (key) => {
    try { localStorage.removeItem(key); return { key, deleted: true }; }
    catch { return null; }
  },
};

// ─── Constants ───────────────────────────────────────────────────────────────
const FY_OPTIONS = ["FY_2025-26", "FY_2026-27"];
const FY_LABELS = { "FY_2025-26": "FY 2025-26", "FY_2026-27": "FY 2026-27" };
const MONTHS = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"];
const CHANNELS = [
  { id:"horeca",    name:"HORECA",          icon:"🏨", color:"#B45309", bg:"#FEF3C7", desc:"Hotels, Restaurants, Cafes" },
  { id:"qcom",      name:"Quick Commerce",  icon:"⚡", color:"#7C3AED", bg:"#EDE9FE", desc:"Zepto, Blinkit, Swiggy" },
  { id:"ecom",      name:"ECommerce",       icon:"📦", color:"#2563EB", bg:"#DBEAFE", desc:"Amazon, Flipkart, JioMart" },
  { id:"physical",  name:"Physical Stores", icon:"🏪", color:"#059669", bg:"#D1FAE5", desc:"Nature's Basket, Retail" },
  { id:"b2b_corp",  name:"B2B Corporates",  icon:"🏢", color:"#475569", bg:"#F1F5F9", desc:"Offices, Institutions" },
  { id:"b2b_vending",name:"B2B Vending",   icon:"💰", color:"#0F766E", bg:"#CCFBF1", desc:"Vending machines" },
  { id:"community", name:"Community Sales", icon:"🤝", color:"#BE185D", bg:"#FCE7F3", desc:"Events, Popups" },
  { id:"website",   name:"Website",         icon:"🌐", color:"#6D28D9", bg:"#F5F3FF", desc:"oatey.in D2C" },
];
const CH_IDS = CHANNELS.map(c => c.id);
const CH_MAP = Object.fromEntries(CHANNELS.map(c => [c.id, c]));
const SKUS = ["Millet","Barista","Chocolate","Caramel Coffee","Kesar Badam","Pre Orders","Assorted Box","Others"];
const OPEX_KEYS = ["employment","director_rem","travel","rent","prof_fees","consulting","legal","tax_paid","software","internet","office","other_admin"];
const OPEX_LABELS = {
  employment:"Employee Costs", director_rem:"Director Remuneration",
  travel:"Travel & Conveyance", rent:"Rent & Taxes", prof_fees:"Professional Fees",
  consulting:"Consulting Fees", legal:"Legal & Audit", tax_paid:"Tax Paid",
  software:"Software", internet:"Internet & Telecom", office:"Office Expenses", other_admin:"Other Admin",
};
const SFIELDS = [
  { key:"date",          label:"Date",                    req:true  },
  { key:"sku",           label:"SKU / Product",           req:true  },
  { key:"quantity",      label:"Quantity",                req:true  },
  { key:"unit_price",    label:"Unit Price (excl. GST)",  req:true  },
  { key:"customer_name", label:"Customer Name",           req:false },
  { key:"city",          label:"City / Location",         req:false },
  { key:"gst",           label:"Total Tax",               req:false },
  { key:"cgst",          label:"CGST",                    req:false },
  { key:"sgst",          label:"SGST",                    req:false },
  { key:"igst",          label:"IGST",                    req:false },
  { key:"order_id",      label:"Order / Invoice ID",      req:false },
];
const K = { mis:"oatey-mis", sales:"oatey-sales", maps:"oatey-maps" };

// ─── Helpers ─────────────────────────────────────────────────────────────────
const emptyMonth = () => ({
  ...Object.fromEntries(CH_IDS.map(c => [c, 0])),
  units_sold:0, cost_per_unit:22, opening_stock:0, closing_stock:0,
  purchases:0, packaging:0, marketplace_fees:0, courier:0, marketing:0,
  ...Object.fromEntries(OPEX_KEYS.map(k => [k, 0])),
});
const TH = {
  navy:"#0C1527", slate:"#1E293B", muted:"#64748B", border:"#E2E8F0",
  bg:"#F8FAFC", card:"#FFFFFF", green:"#059669", red:"#DC2626",
  amber:"#D97706", blue:"#2563EB",
};
const fmt = n => {
  if (n == null || isNaN(n)) return "–";
  const a = Math.abs(n);
  if (a >= 1e7) return (n < 0 ? "-" : "") + "₹" + (a / 1e7).toFixed(2) + "Cr";
  if (a >= 1e5) return (n < 0 ? "-" : "") + "₹" + (a / 1e5).toFixed(2) + "L";
  if (a >= 1e3) return (n < 0 ? "-" : "") + "₹" + (a / 1e3).toFixed(1) + "K";
  return (n < 0 ? "-₹" : "₹") + a.toFixed(0);
};
const fN = n => (n == null || isNaN(n)) ? "0" : Math.round(n).toLocaleString("en-IN");
const pc = n => n == null || isNaN(n) ? "–" : (n >= 0 ? "+" : "") + (n * 100).toFixed(1) + "%";
const pcA = n => n == null ? "–" : (n * 100).toFixed(1) + "%";
const safeNum = v => {
  if (v == null) return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const s = String(v).replace(/[₹$€,\s]/g, "").replace(/INR|USD|Rs\.?/gi, "").trim();
  if (!s) return 0;
  const m = s.match(/^\((.+)\)$/);
  const n = parseFloat(m ? "-" + m[1] : s);
  return isNaN(n) ? 0 : n;
};
const daysInMonth = (m, y) => new Date(y, m, 0).getDate();
const getCalMonth = mi => { const m = mi + 4; return m > 12 ? m - 12 : m; };
const getCalYear = (mi, fy) => {
  const cm = getCalMonth(mi);
  const s = parseInt(fy.split("_")[1].split("-")[0]);
  return cm >= 4 ? s : s + 1;
};
const gMD = (mi, fy, dy = 15) =>
  `${getCalYear(mi, fy)}-${String(getCalMonth(mi)).padStart(2, "0")}-${String(dy).padStart(2, "0")}`;
const getWeekNum = d => { const day = new Date(d).getDate(); return day <= 7 ? 1 : day <= 14 ? 2 : day <= 21 ? 3 : 4; };

const SKU_KW = {
  "Millet":["millet","millets"], "Barista":["barista"],
  "Chocolate":["chocolate","choco","cocoa"],
  "Caramel Coffee":["caramel","coffee","cafe","latte"],
  "Kesar Badam":["kesar","badam","almond","saffron"],
  "Pre Orders":["pre order","preorder","pre-order","advance"],
  "Assorted Box":["assorted","combo","variety","mix","sampler","gift","bundle","hamper"],
};
const detSKU = t => {
  if (!t) return "Others";
  const s = String(t).toLowerCase().replace(/[^a-z0-9\s\-]/g, "").trim();
  if (!s) return "Others";
  let best = null, bs = 0;
  for (const [sku, kws] of Object.entries(SKU_KW)) {
    for (const kw of kws) {
      if (s.includes(kw)) { const sc = 0.85 + kw.length / s.length * 0.15; if (sc > bs) { bs = sc; best = sku; } }
      const ws = s.split(/\s+/);
      for (const w of ws) { if (w.length >= 3 && kw.includes(w) && 0.4 > bs) { bs = 0.4; best = sku; } }
    }
  }
  return bs >= 0.3 ? best : "Others";
};
const pDate = v => {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v)) return v.toISOString().split("T")[0];
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    return d ? `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}` : null;
  }
  const s = String(v).trim(); let m;
  if ((m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/)))
    return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  if ((m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/))) {
    const a = +m[1], b = +m[2];
    if (a > 12) return `${m[3]}-${String(b).padStart(2,"0")}-${String(a).padStart(2,"0")}`;
    return `${m[3]}-${String(b).padStart(2,"0")}-${String(a).padStart(2,"0")}`;
  }
  const d = new Date(s); return !isNaN(d) ? d.toISOString().split("T")[0] : null;
};
const CITY_NORM = {
  "bangalore":"Bengaluru","bengaluru":"Bengaluru","blr":"Bengaluru",
  "bombay":"Mumbai","mumbai":"Mumbai","delhi":"Delhi","new delhi":"Delhi",
  "gurgaon":"Gurugram","gurugram":"Gurugram","chennai":"Chennai",
  "hyderabad":"Hyderabad","kolkata":"Kolkata","calcutta":"Kolkata",
  "pune":"Pune","ahmedabad":"Ahmedabad","jaipur":"Jaipur","lucknow":"Lucknow",
  "kochi":"Kochi","cochin":"Kochi","noida":"Noida","surat":"Surat",
  "indore":"Indore","nagpur":"Nagpur","coimbatore":"Coimbatore",
  "mysore":"Mysuru","mysuru":"Mysuru","mangalore":"Mangaluru",
  "chandigarh":"Chandigarh","bhopal":"Bhopal",
};
const CITY_REGIONS = {
  "Bengaluru":"South","Chennai":"South","Hyderabad":"South","Kochi":"South",
  "Coimbatore":"South","Mysuru":"South","Mangaluru":"South",
  "Mumbai":"West","Pune":"West","Ahmedabad":"West","Surat":"West","Nagpur":"West","Indore":"West","Bhopal":"West",
  "Delhi":"North","Noida":"North","Gurugram":"North","Jaipur":"North","Lucknow":"North","Chandigarh":"North",
  "Kolkata":"East","Patna":"East","Guwahati":"East","Bhubaneswar":"East",
};
const normCity = raw => {
  if (!raw) return { city:"Unknown", region:"Unknown" };
  const k = String(raw).trim().toLowerCase().replace(/[^a-z\s]/g,"").trim();
  const city = CITY_NORM[k] || raw.charAt(0).toUpperCase() + raw.slice(1).toLowerCase();
  return { city, region: CITY_REGIONS[city] || "Other" };
};
const autoMap = h => {
  const m = {}; const l = h.map(x => String(x || "").toLowerCase());
  const f = p => l.findIndex(x => p.some(q => x.includes(q)));
  let i;
  if ((i = f(["date","order date","invoice date"])) >= 0) m.date = h[i];
  if ((i = f(["sku","product","item","description"])) >= 0) m.sku = h[i];
  if ((i = f(["qty","quantity","units"])) >= 0) m.quantity = h[i];
  if ((i = f(["rate","unit price","price","selling price","mrp"])) >= 0) m.unit_price = h[i];
  if ((i = f(["customer name","customer","buyer","client","sold to"])) >= 0) m.customer_name = h[i];
  if ((i = f(["city","location","ship city","billing city","delivery city"])) >= 0) m.city = h[i];
  if ((i = f(["gst","tax amount","total tax"])) >= 0) m.gst = h[i];
  if ((i = f(["cgst"])) >= 0) m.cgst = h[i];
  if ((i = f(["sgst"])) >= 0) m.sgst = h[i];
  if ((i = f(["igst"])) >= 0) m.igst = h[i];
  if ((i = f(["order id","order no","invoice no","invoice number"])) >= 0) m.order_id = h[i];
  return m;
};

// ─── UI Components ───────────────────────────────────────────────────────────
function Toast({ msg, type, onClose }) {
  useEffect(() => { const t = setTimeout(onClose, 4000); return () => clearTimeout(t); }, [onClose]);
  return (
    <div style={{ position:"fixed", bottom:20, left:"50%", transform:"translateX(-50%)",
      background: type === "error" ? TH.red : TH.green, color:"#fff", padding:"10px 24px",
      borderRadius:8, fontSize:13, fontWeight:600, zIndex:1000, boxShadow:"0 4px 20px rgba(0,0,0,0.15)" }}>
      {msg}
    </div>
  );
}
function Metric({ label, value, sub, trend, small }) {
  const tC = trend === "up" ? TH.green : trend === "down" ? TH.red : TH.muted;
  return (
    <div style={{ background:TH.card, borderRadius:10, padding:small?"12px 14px":"18px 20px",
      border:`1px solid ${TH.border}`, flex:1, minWidth:small?110:150 }}>
      <div style={{ fontSize:11, color:TH.muted, fontWeight:500, letterSpacing:"0.04em",
        textTransform:"uppercase", marginBottom:small?4:6 }}>{label}</div>
      <div style={{ fontSize:small?18:26, fontWeight:700, color:TH.navy,
        fontFamily:"'JetBrains Mono',monospace", letterSpacing:"-0.02em" }}>{value}</div>
      {sub && <div style={{ fontSize:11, color:tC, fontWeight:600, marginTop:3 }}>{sub}</div>}
    </div>
  );
}
function DataTable({ headers, rows, compact }) {
  return (
    <table style={{ width:"100%", borderCollapse:"collapse", fontSize:compact?11:12 }}>
      <thead><tr>{headers.map((h, i) =>
        <th key={i} style={{ padding:compact?"6px 8px":"8px 12px", textAlign:i>0?"right":"left",
          fontWeight:600, color:TH.muted, fontSize:10, textTransform:"uppercase",
          letterSpacing:"0.05em", borderBottom:`2px solid ${TH.border}`, background:TH.bg }}>{h}</th>
      )}</tr></thead>
      <tbody>{rows.map((row, ri) =>
        <tr key={ri}>{row.map((cell, ci) =>
          <td key={ci} style={{ padding:compact?"5px 8px":"7px 12px", textAlign:ci>0?"right":"left",
            borderBottom:`1px solid ${TH.border}`,
            fontFamily:ci>0?"'JetBrains Mono',monospace":"inherit",
            fontWeight:ci===row.length-1?600:400,
            color:typeof cell==="string"&&cell.startsWith("-")?TH.red:TH.navy,
            fontSize:compact?11:12 }}>{cell}</td>
        )}</tr>
      )}</tbody>
    </table>
  );
}
function ChipBadge({ text, color }) {
  return <span style={{ fontSize:10, fontWeight:600, color, background:color+"14", padding:"2px 8px", borderRadius:4 }}>{text}</span>;
}
function SectionCard({ title, children, noPad }) {
  return (
    <div style={{ background:TH.card, borderRadius:10, border:`1px solid ${TH.border}`,
      overflow:"hidden", marginBottom:16 }}>
      {title && <div style={{ padding:"12px 18px", borderBottom:`1px solid ${TH.border}` }}>
        <div style={{ fontSize:14, fontWeight:700, color:TH.navy }}>{title}</div>
      </div>}
      <div style={{ padding:noPad?0:"16px 18px" }}>{children}</div>
    </div>
  );
}
function TabBtn({ active, children, onClick }) {
  return (
    <button onClick={onClick} style={{ padding:"10px 20px", border:"none", cursor:"pointer",
      fontSize:13, fontWeight:600, borderRadius:"8px 8px 0 0",
      background:active?TH.navy:"#E2E8F0", color:active?"#fff":TH.muted, whiteSpace:"nowrap" }}>
      {children}
    </button>
  );
}
function MonthPill({ label, active, hasData, onClick }) {
  return (
    <button onClick={onClick} style={{ width:44, padding:"4px 0",
      border:`1px solid ${active?TH.navy:hasData?TH.green:TH.border}`,
      borderRadius:6, cursor:"pointer", fontSize:11, fontWeight:600, textAlign:"center",
      background:active?TH.navy:hasData?"#F0FDF4":"#fff", color:active?"#fff":TH.navy }}>
      {label}
      {hasData && !active && <div style={{ width:4, height:4, borderRadius:"50%", background:TH.green, margin:"2px auto 0" }}/>}
    </button>
  );
}
function Collapsible({ title, icon, children, open: defOpen }) {
  const [open, setOpen] = useState(defOpen !== false);
  return (
    <div style={{ marginBottom:14, border:`1px solid ${TH.border}`, borderRadius:10, overflow:"hidden", background:TH.card }}>
      <div onClick={() => setOpen(!open)} style={{ display:"flex", alignItems:"center",
        justifyContent:"space-between", padding:"12px 18px", cursor:"pointer",
        background:open?TH.bg:"#fff", borderBottom:open?`1px solid ${TH.border}`:"none" }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <span style={{ fontSize:16 }}>{icon}</span>
          <span style={{ fontWeight:600, fontSize:13, color:TH.navy }}>{title}</span>
        </div>
        <span style={{ fontSize:11, color:TH.muted, transform:open?"rotate(180deg)":"none", transition:"transform 0.2s" }}>▼</span>
      </div>
      {open && <div style={{ padding:"14px 18px" }}>{children}</div>}
    </div>
  );
}
function NumInput({ label, value, onChange, prefix, highlight }) {
  const pf = prefix === undefined ? "₹" : prefix;
  return (
    <div style={{ display:"flex", flexDirection:"column", gap:3, flex:1 }}>
      <label style={{ fontSize:11, color:TH.muted, fontWeight:500 }}>{label}</label>
      <div style={{ display:"flex", alignItems:"center",
        border:`${highlight?"2px":"1px"} solid ${highlight?TH.blue:TH.border}`,
        borderRadius:6, overflow:"hidden", background:highlight?"#EFF6FF":"#fff" }}>
        {pf && <span style={{ padding:"6px 8px", fontSize:12, color:TH.muted,
          background:TH.bg, borderRight:`1px solid ${TH.border}` }}>{pf}</span>}
        <input type="number" value={value||""} onChange={e => onChange(parseFloat(e.target.value)||0)}
          placeholder="0" style={{ border:"none", outline:"none", padding:"8px 10px",
          fontSize:13, width:"100%", background:"transparent", fontFamily:"'JetBrains Mono',monospace" }}/>
      </div>
    </div>
  );
}
function PLRow({ label, value, rev, bold, bg, indent }) {
  const v = value || 0; const ind = indent || 0;
  return (
    <div style={{ display:"flex", alignItems:"center", padding:"6px 16px",
      paddingLeft:16+ind*18, background:bg||"transparent", borderBottom:`1px solid ${TH.border}` }}>
      <span style={{ flex:1, fontSize:13, fontWeight:bold?700:400, color:TH.navy }}>{label}</span>
      <span style={{ width:110, textAlign:"right", fontSize:13, fontWeight:bold?700:400,
        color:v<0?TH.red:TH.navy, fontFamily:"'JetBrains Mono',monospace" }}>
        {v < 0 ? "("+fmt(-v)+")" : fmt(v)}
      </span>
      <span style={{ width:65, textAlign:"right", fontSize:11, color:TH.muted, fontFamily:"'JetBrains Mono',monospace" }}>
        {rev ? pcA(Math.abs(v)/rev) : ""}
      </span>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function MISPortal() {
  const [fy, sFy] = useState("FY_2025-26");
  const [mi, sMi] = useState(9);
  const [dt, sDt] = useState(emptyMonth());
  const [ad, sAd] = useState({});
  const [vw, sVw] = useState("dashboard");
  const [ld, sLd] = useState(true);
  const [sv, sSv] = useState(false);
  const [inv, sInv] = useState([]);
  const [toast, sToast] = useState(null);
  const [tType, sTT] = useState("success");
  const [uChan, sUChan] = useState(null);
  const [uStep, sUStep] = useState("select");
  const [rH, sRH] = useState([]);
  const [rD, sRD] = useState([]);
  const [cMap, sCMap] = useState({});
  const [sMaps, setSMaps] = useState({});
  const [pRows, sPRows] = useState([]);
  const [skuOv, sSkuOv] = useState({});
  const [fName, sFName] = useState("");
  const fRef = useRef(null);
  const [svm, sSvm] = useState("ledger");
  const [sw, sSw] = useState(0);
  const [iDate, sIDate] = useState(gMD(9, "FY_2025-26"));
  const [iChan, sIChan] = useState("ecom");
  const [iItems, sIItems] = useState([{ sku:"Millet", qty:0, price:0, custName:"", city:"", region:"" }]);
  const [showDP, sShowDP] = useState(false);
  const importRef = useRef(null);

  const sk = `${K.mis}:${fy}`, slk = `${K.sales}:${fy}`;
  const show = (m, t) => { sTT(t||"success"); sToast(m); };

  useEffect(() => {
    (async () => {
      try { const r = storage.get(sk); if (r?.value) { const p = JSON.parse(r.value); sAd(p); if (p[mi]) sDt(p[mi]); } } catch(e) {}
      try { const r = storage.get(slk); if (r?.value) sInv(JSON.parse(r.value)); } catch(e) {}
      try { const r = storage.get(K.maps); if (r?.value) setSMaps(JSON.parse(r.value)); } catch(e) {}
      sLd(false);
    })();
  }, [fy]);

  useEffect(() => { if (ad[mi]) sDt({ ...emptyMonth(), ...ad[mi] }); else sDt(emptyMonth()); }, [mi, ad]);
  useEffect(() => { sIDate(gMD(mi, fy)); }, [mi, fy]);

  useEffect(() => {
    if (!inv.length) return;
    let changed = false;
    const repaired = inv.map(i => {
      const cs = (i.items||[]).reduce((s,it) => s+safeNum(it.qty)*safeNum(it.price), 0);
      const cu = (i.items||[]).reduce((s,it) => s+safeNum(it.qty), 0);
      if (Math.abs((i.subtotal||0)-cs)>0.01 || Math.abs((i.units||0)-cu)>0) {
        changed = true; return { ...i, subtotal:Math.round(cs*100)/100, units:cu };
      }
      return i;
    });
    if (changed) { sInv(repaired); saveInv(repaired); }
  }, []);

  const upd = (k, v) => sDt(p => ({ ...p, [k]:v }));
  const saveDt = async () => {
    sSv(true);
    const u = { ...ad, [mi]:dt }; sAd(u);
    try { storage.set(sk, JSON.stringify(u)); } catch(e) {}
    sSv(false); show("Saved " + MONTHS[mi]);
  };
  const saveInv = i => { try { storage.set(slk, JSON.stringify(i)); } catch(e) {} };
  const saveMaps = m => { setSMaps(m); try { storage.set(K.maps, JSON.stringify(m)); } catch(e) {} };

  const nxId = useMemo(() => {
    if (!inv.length) return "INV-001001";
    const n = inv.map(i => parseInt(i.id.replace("INV-",""))).filter(n => !isNaN(n));
    return "INV-" + String(Math.max(...n, 1000)+1).padStart(6,"0");
  }, [inv]);

  const handleFile = e => {
    const f = e.target.files?.[0]; if (!f) return;
    sFName(f.name);
    const r = new FileReader();
    r.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type:"array", cellDates:true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval:"", raw:false });
        if (!json.length) { show("Empty file","error"); return; }
        const h = Object.keys(json[0]);
        sRH(h); sRD(json);
        const saved = sMaps[uChan?.id];
        const auto = autoMap(h);
        sCMap(saved ? { ...auto, ...Object.fromEntries(Object.entries(saved).filter(([k,v]) => h.includes(v))) } : auto);
        sUStep("map");
      } catch(err) { show("Parse error","error"); }
    };
    r.readAsArrayBuffer(f);
  };

  const applyMap = () => {
    if (!cMap.date || !cMap.sku || !cMap.quantity) { show("Map Date, SKU, Qty","error"); return; }
    saveMaps({ ...sMaps, [uChan.id]:cMap });
    const rows = rD.map((row,i) => {
      const d = pDate(row[cMap.date]);
      const sr = String(row[cMap.sku]||"");
      const sk2 = detSKU(sr);
      const q = safeNum(row[cMap.quantity]);
      const p = cMap.unit_price ? safeNum(row[cMap.unit_price]) : 0;
      const g = (cMap.gst?safeNum(row[cMap.gst]):0)+(cMap.cgst?safeNum(row[cMap.cgst]):0)
               +(cMap.sgst?safeNum(row[cMap.sgst]):0)+(cMap.igst?safeNum(row[cMap.igst]):0);
      const cn = cMap.customer_name ? String(row[cMap.customer_name]||"").trim() : "";
      const rawCity = cMap.city ? String(row[cMap.city]||"").trim() : "";
      const { city:normC, region:regC } = normCity(rawCity);
      const unitPrice = Math.round(Math.abs(p)*100)/100;
      const lineTotal = Math.round(Math.max(0, Math.round(q))*unitPrice*100)/100;
      return { idx:i, date:d, skuRaw:sr, sku:sk2, qty:Math.max(0,Math.round(q)),
               unitPrice, gst:Math.round(Math.abs(g)*100)/100, total:lineTotal,
               custName:cn, city:normC, region:regC };
    }).filter(r => r.qty > 0);
    sPRows(rows); sSkuOv({}); sUStep("preview");
  };

  const genInv = () => {
    const valid = pRows.filter(r => (skuOv[r.idx]||r.sku) && r.date);
    if (!valid.length) { show("No valid rows","error"); return; }
    const byD = {};
    valid.forEach(r => { if (!byD[r.date]) byD[r.date]=[]; byD[r.date].push(r); });
    let ni = [...inv], mx = inv.length ? Math.max(...inv.map(i => parseInt(i.id.replace("INV-",""))||1000), 1000) : 1000;
    let ct=0, tr=0, tu=0;
    const monthCounts = {};
    Object.entries(byD).forEach(([date,rows]) => {
      mx++;
      const items = rows.map(r => ({ sku:skuOv[r.idx]||r.sku, qty:r.qty, price:r.unitPrice, custName:r.custName||"", city:r.city||"Unknown", region:r.region||"Unknown" }));
      const sub = rows.reduce((s,r) => s+r.total, 0);
      const units = rows.reduce((s,r) => s+r.qty, 0);
      ni.push({ id:"INV-"+String(mx).padStart(6,"0"), date, channel:uChan.id, items, subtotal:Math.round(sub*100)/100, units, gst:Math.round(rows.reduce((s,r)=>s+r.gst,0)*100)/100, status:"raised", createdAt:Date.now(), source:uChan.id });
      ct++; tr+=sub; tu+=units;
      const dd = new Date(date); const mk = `${dd.getFullYear()}-${dd.getMonth()+1}`;
      monthCounts[mk] = (monthCounts[mk]||0) + rows.length;
    });
    sInv(ni); saveInv(ni);
    if (Object.keys(monthCounts).length > 0) {
      const top = Object.entries(monthCounts).sort((a,b) => b[1]-a[1])[0][0];
      const [yr, cm2] = top.split("-").map(Number);
      const fm = cm2>=4 ? cm2-4 : cm2+8;
      const fs = cm2>=4 ? yr : yr-1;
      const df = `FY_${fs}-${String(fs+1).slice(2)}`;
      if (FY_OPTIONS.includes(df)) sFy(df); sMi(fm);
    }
    sUStep("done"); show(`${ct} invoices — ₹${fN(Math.round(tr))} | ${fN(tu)} units`);
  };

  const resetUp = () => { sUStep("select"); sUChan(null); sRH([]); sRD([]); sPRows([]); sCMap({}); sFName(""); if (fRef.current) fRef.current.value=""; };
  const delInv = id => { const u = inv.filter(i => i.id !== id); sInv(u); saveInv(u); show("Deleted "+id); };
  const iLT = it => safeNum(it.qty) * safeNum(it.price);
  const iSub = iItems.reduce((s,it) => s+iLT(it), 0);
  const iUnits = iItems.reduce((s,it) => s+safeNum(it.qty), 0);
  const raiseInv = () => {
    if (iSub <= 0) return;
    const nv = { id:nxId, date:iDate, channel:iChan, items:iItems.map(it=>({...it})), subtotal:iSub, units:iUnits, status:"raised", createdAt:Date.now() };
    const u = [...inv, nv]; sInv(u); saveInv(u);
    sIItems([{ sku:"Millet", qty:0, price:0, custName:"", city:"", region:"" }]);
    show(`${nv.id} — ₹${fN(iSub)}`); sSvm("ledger");
  };
  const hDC = nd => { sIDate(nd); const d=new Date(nd); const cm2=d.getMonth()+1; const fm=cm2>=4?cm2-4:cm2+8; if(fm!==mi) sMi(fm); };

  const exportData = () => {
    const bundle = { v:2, ts:new Date().toISOString(), d:{} };
    try { const r=storage.get(sk); if(r?.value) bundle.d[`mis:${fy}`]=r.value; } catch(e) {}
    try { const r=storage.get(slk); if(r?.value) bundle.d[`sales:${fy}`]=r.value; } catch(e) {}
    try { const r=storage.get(K.maps); if(r?.value) bundle.d.maps=r.value; } catch(e) {}
    const blob = new Blob([JSON.stringify(bundle,null,2)], { type:"application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href=url; a.download=`MIS_Export_${new Date().toISOString().split("T")[0]}.json`; a.click(); URL.revokeObjectURL(url);
    show("Exported JSON");
  };

  const importData = async e => {
    const file = e.target.files?.[0]; if (!file) return;
    try {
      const bundle = JSON.parse(await file.text());
      if (!bundle.d) { show("Invalid file","error"); return; }
      let c = 0;
      for (const [key,val] of Object.entries(bundle.d)) {
        if (key.startsWith("mis:")) { storage.set(`${K.mis}:${key.slice(4)}`, val); c++; }
        else if (key.startsWith("sales:")) { storage.set(`${K.sales}:${key.slice(6)}`, val); c++; }
        else if (key === "maps") { storage.set(K.maps, val); c++; }
      }
      try { const r=storage.get(sk); if(r?.value){const p=JSON.parse(r.value);sAd(p);if(p[mi])sDt(p[mi]);} } catch(e){}
      try { const r=storage.get(slk); if(r?.value) sInv(JSON.parse(r.value)); } catch(e){}
      try { const r=storage.get(K.maps); if(r?.value) setSMaps(JSON.parse(r.value)); } catch(e){}
      show(`Imported ${c} datasets`); sShowDP(false);
    } catch(err) { show("Parse error","error"); }
    if (importRef.current) importRef.current.value="";
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const addSheet = (data, name) => {
      const ws = XLSX.utils.aoa_to_sheet(data);
      ws["!cols"] = data[0].map((_,ci) => ({ wch:Math.min(Math.max(...data.map(r=>String(r[ci]??"").length),6)+2,35) }));
      XLSX.utils.book_append_sheet(wb, ws, name.substring(0,31));
    };
    const cS = getMonthStats(mi);
    const tS2 = CH_IDS.reduce((s,k) => s+(dt[k]||0), 0);
    const cogs2 = (dt.units_sold||0)*(dt.cost_per_unit||22);
    const gp2 = tS2-cogs2;
    const vc2 = (dt.packaging||0)+(dt.marketplace_fees||0)+(dt.courier||0);
    const cm12 = gp2-vc2; const cm22 = cm12-(dt.marketing||0);
    const tOp2 = OPEX_KEYS.reduce((s,k) => s+(dt[k]||0), 0);
    const ebitda2 = cm22-tOp2;
    addSheet([
      ["Plant Essentials — MIS Report"],
      ["Period", MONTHS[mi]+" "+FY_LABELS[fy]],
      ["Generated", new Date().toLocaleDateString("en-IN")],
      [],
      ["Metric","Value"],
      ["Revenue (Invoiced)", cS.rev], ["Units Sold", cS.units],
      ["COGS", cogs2], ["Gross Profit", gp2], ["Gross Margin %", tS2?gp2/tS2:0], ["EBITDA", ebitda2],
    ], "Summary");
    const s2 = [["SKU","Month","Type","Qty","Rate","Revenue","Velocity"]];
    (skuMovement||[]).forEach(s => {
      (s.hist||[]).forEach(h => { s2.push([s.sku, h.month, "Actual", h.qty, s.bestPrice, Math.round(safeNum(h.qty)*safeNum(s.bestPrice)*100)/100, s.velocity]); });
      if (s.current) s2.push([s.sku, s.current.month, skuProjection.isPartial?"Projected":"Actual", s.current.projected, s.bestPrice, Math.round(s.current.projected*s.bestPrice*100)/100, s.velocity]);
      (s.future||[]).forEach(f => { s2.push([s.sku, f.month, f.monthLabel, f.qty, s.bestPrice, Math.round(f.rev*100)/100, s.velocity]); });
    });
    addSheet(s2, "SKU Data");
    const histM = []; for (let i=Math.max(0,mi-2);i<=mi;i++) histM.push({ month:MONTHS[i], ...getMonthStats(i) });
    const s3 = [["Month", ...CHANNELS.map(c=>c.name), "Total"]];
    histM.forEach(m => { const row=[m.month]; let tot=0; CHANNELS.forEach(c=>{const v=m.byCh[c.id]||0;row.push(v);tot+=v;}); row.push(Math.round(tot*100)/100); s3.push(row); });
    addSheet(s3, "By Channel");
    const cities = Object.values(cS.byCity||{}).sort((a,b)=>b.rev-a.rev);
    const s4 = [["City","Region","Revenue","Units","%"]];
    cities.forEach(c => { s4.push([c.city,c.region,Math.round(c.rev*100)/100,c.units,cS.rev?c.rev/cS.rev:0]); });
    addSheet(s4, "By Location");
    const s5 = [["Invoice","Date","Channel","SKU","Customer","City","Qty","Rate","Revenue","GST"]];
    mInv.forEach(iv => { (iv.items||[]).forEach(it => { const q=safeNum(it.qty),p=safeNum(it.price); s5.push([iv.id,iv.date,CH_MAP[iv.channel]?.name||iv.channel,it.sku||"",it.custName||"",it.city||"Unknown",q,p,Math.round(q*p*100)/100,iv.gst||0]); }); });
    addSheet(s5, "Raw Data");
    XLSX.writeFile(wb, `Sales_Report_${MONTHS[mi]}_${FY_LABELS[fy].replace(/\s/g,"_")}.xlsx`);
    show("Excel exported (5 sheets)");
  };

  const cm = getCalMonth(mi); const cy = getCalYear(mi, fy);
  const mInv = useMemo(() => inv.filter(iv => { const d=new Date(iv.date); return d.getMonth()+1===cm&&d.getFullYear()===cy; }), [inv,cm,cy]);
  const wInv = useMemo(() => sw===0 ? mInv : mInv.filter(i=>getWeekNum(i.date)===sw), [mInv,sw]);

  const getMonthInv = idx => { const c=getCalMonth(idx),y=getCalYear(idx,fy); return inv.filter(iv=>{const d=new Date(iv.date);return d.getMonth()+1===c&&d.getFullYear()===y;}); };
  const getMonthStats = idx => {
    const mi2 = getMonthInv(idx); let rev=0,units=0;
    mi2.forEach(i=>{(i.items||[]).forEach(it=>{const q=safeNum(it.qty),p=safeNum(it.price);rev+=q*p;units+=q;});});
    rev=Math.round(rev*100)/100; const orders=mi2.length;
    const byCh={};
    CH_IDS.forEach(c=>{let v=0;mi2.filter(i=>i.channel===c).forEach(i=>{(i.items||[]).forEach(it=>{v+=safeNum(it.qty)*safeNum(it.price);});});byCh[c]=Math.round(v*100)/100;});
    const bySKU={};
    mi2.forEach(i=>{(i.items||[]).forEach(it=>{const q=safeNum(it.qty),p=safeNum(it.price);if(!bySKU[it.sku])bySKU[it.sku]={qty:0,rev:0};bySKU[it.sku].qty+=q;bySKU[it.sku].rev+=Math.round(q*p*100)/100;});});
    const byCust={};
    mi2.forEach(i=>{(i.items||[]).forEach(it=>{const n=it.custName||CH_MAP[i.channel]?.name||"Other",q=safeNum(it.qty),p=safeNum(it.price);if(!byCust[n])byCust[n]={rev:0,units:0};byCust[n].rev+=Math.round(q*p*100)/100;byCust[n].units+=q;});});
    const byCity={};
    mi2.forEach(i=>{(i.items||[]).forEach(it=>{const c=it.city||"Unknown",q=safeNum(it.qty),p=safeNum(it.price);if(!byCity[c])byCity[c]={city:c,region:it.region||CITY_REGIONS[c]||"Other",rev:0,units:0,orders:0};byCity[c].rev+=Math.round(q*p*100)/100;byCity[c].units+=q;byCity[c].orders+=1;});});
    const byRegion={};
    Object.values(byCity).forEach(c=>{const r=c.region;if(!byRegion[r])byRegion[r]={region:r,rev:0,units:0};byRegion[r].rev+=c.rev;byRegion[r].units+=c.units;});
    return { rev,units,orders,byCh,bySKU,byCust,byCity,byRegion };
  };

  const last3 = useMemo(()=>{ const result=[]; for(let i=0;i<3;i++){const idx=mi-i;if(idx<0)continue;result.unshift({month:MONTHS[idx],idx,...getMonthStats(idx)});} return result; },[mi,inv,fy]);

  const skuProjection = useMemo(() => {
    const today=new Date(); const totalDays=daysInMonth(cm,cy);
    const daysPassed=today.getFullYear()===cy&&today.getMonth()+1===cm?today.getDate():totalDays;
    const isPartial=daysPassed<totalDays;
    const histIndices=[]; for(let i=Math.max(0,mi-6);i<=mi;i++) histIndices.push(i);
    const histStats=histIndices.map(i=>({idx:i,month:MONTHS[i],...getMonthStats(i)}));
    const curMonthStats=histStats[histStats.length-1];
    const prevStats=histStats.length>1?histStats[histStats.length-2]:{rev:0,units:0,byCh:{},bySKU:{}};
    const chIndices=histIndices.slice(-4);
    const chHistStats=chIndices.map(i=>getMonthStats(i));
    const chVelocity={},chShare={};let totalChRev=0;
    CH_IDS.forEach(c=>{
      const vals=chHistStats.map(h=>h.byCh[c]||0);
      const recentRev=vals.slice(-3).reduce((s,v)=>s+v,0); totalChRev+=recentRev;
      const growths=[]; for(let i=1;i<vals.length;i++){if(vals[i-1]>0)growths.push((vals[i]-vals[i-1])/vals[i-1]);}
      let wG=0; if(growths.length>=2)wG=growths[growths.length-1]*0.6+growths[growths.length-2]*0.4; else if(growths.length===1)wG=growths[0];
      chVelocity[c]=Math.max(-0.10,Math.min(0.50,wG)); chShare[c]=recentRev;
    });
    CH_IDS.forEach(c=>{chShare[c]=totalChRev>0?chShare[c]/totalChRev:1/CH_IDS.length;});
    const channelGrowthRate=Math.max(0.03,CH_IDS.reduce((s,c)=>s+chShare[c]*chVelocity[c],0));
    const skuForecasts=SKUS.map(sku=>{
      const hist=histStats.slice(0,-1).map(h=>{const d=h.bySKU[sku]||{qty:0,rev:0};return{month:h.month,qty:d.qty,rev:d.rev,avgPrice:d.qty>0?d.rev/d.qty:0};});
      const curData=curMonthStats.bySKU[sku]||{qty:0,rev:0};
      let bestPrice=curData.qty>0?curData.rev/curData.qty:0;
      if(!bestPrice){for(let i=hist.length-1;i>=0;i--){if(hist[i].avgPrice>0){bestPrice=hist[i].avgPrice;break;}}}
      const histQtys=hist.map(h=>h.qty).filter(q=>q>0);
      const dampened=hist.map(h=>{
        if(histQtys.length<2)return h.qty;
        const others=histQtys.filter(q=>q!==h.qty);
        const oAvg=others.length?others.reduce((a,b)=>a+b,0)/others.length:0;
        if(oAvg>0&&h.qty>2*oAvg)return Math.round(h.qty*0.5+oAvg*0.5); return h.qty;
      });
      const recent=dampened.slice(-4); const skuGrowths=[];
      for(let i=1;i<recent.length;i++){if(recent[i-1]>0)skuGrowths.push((recent[i]-recent[i-1])/recent[i-1]);}
      let skuGrowth=0;
      if(skuGrowths.length>=3)skuGrowth=skuGrowths[2]*0.60+skuGrowths[1]*0.30+skuGrowths[0]*0.10;
      else if(skuGrowths.length===2)skuGrowth=skuGrowths[1]*0.65+skuGrowths[0]*0.35;
      else if(skuGrowths.length===1)skuGrowth=skuGrowths[0];
      skuGrowth=Math.max(-0.15,Math.min(0.40,skuGrowth));
      const last3d=dampened.slice(-3); let base=0;
      if(last3d.length>=3)base=last3d[2]*0.50+last3d[1]*0.30+last3d[0]*0.20;
      else if(last3d.length===2)base=last3d[1]*0.65+last3d[0]*0.35;
      else if(last3d.length===1)base=last3d[0];
      let curProj;
      if(isPartial&&curData.qty>0)curProj=Math.round(0.6*(curData.qty/daysPassed)*totalDays+0.4*base);
      else if(curData.qty>0)curProj=curData.qty;
      else curProj=Math.max(0,Math.round(base*(1+skuGrowth)));
      const combinedGrowth=0.50*skuGrowth+0.50*channelGrowthRate;
      const future=[]; let prev=curProj;
      for(let f=0;f<3;f++){
        const fIdx=(mi+f+1)%12;
        const appliedGrowth=Math.max(0.03,combinedGrowth);
        const proj=Math.max(1,Math.round(prev*(1+appliedGrowth)));
        future.push({month:MONTHS[fIdx],monthLabel:`n+${f+1}`,qty:proj,rev:Math.round(proj*bestPrice),growthRate:appliedGrowth});
        prev=proj;
      }
      const avgV=(curProj+future.reduce((s,f)=>s+f.qty,0))/(1+future.length);
      return{sku,bestPrice:Math.round(bestPrice*100)/100,hist:hist.map((h,i)=>({...h,qty:dampened[i]})),current:{month:MONTHS[mi],qty:isPartial?curData.qty:curProj,projected:curProj,actual:curData.qty,isPartial},future,skuGrowth,combinedGrowth,velocity:avgV>=100?"Fast":avgV>=30?"Medium":"Slow"};
    });
    const curRevenue=skuForecasts.reduce((s,f)=>s+f.current.projected*f.bestPrice,0);
    const curActualRev=skuForecasts.reduce((s,f)=>s+f.current.actual*f.bestPrice,0);
    const curUnits=skuForecasts.reduce((s,f)=>s+f.current.projected,0);
    const futureMonths=[];
    for(let f=0;f<3;f++){
      const fIdx=(mi+f+1)%12;
      const rev=skuForecasts.reduce((s,sk2)=>{const fm=sk2.future[f];return s+(fm?fm.rev:0);},0);
      const units=skuForecasts.reduce((s,sk2)=>{const fm=sk2.future[f];return s+(fm?fm.qty:0);},0);
      const byCh={};
      CH_IDS.forEach(c=>{const v=(curMonthStats.byCh[c]||0)*(1+chVelocity[c]*(f+1));byCh[c]=Math.round(Math.max(0,v));});
      const chT=Object.values(byCh).reduce((s,v)=>s+v,0);
      if(chT>0){const sc=rev/chT;CH_IDS.forEach(c=>{byCh[c]=Math.round(byCh[c]*sc);});}
      futureMonths.push({month:MONTHS[fIdx],monthLabel:`n+${f+1}`,rev:Math.round(rev),units,byCh,growthVsBase:curRevenue>0?(rev-curRevenue)/curRevenue:0});
    }
    const moM=prevStats.rev>0?(curActualRev-prevStats.rev)/prevStats.rev:null;
    return{skus:skuForecasts.filter(s=>s.current.projected>0||s.hist.some(h=>h.qty>0)),currentMonth:{month:MONTHS[mi],rev:Math.round(curRevenue),actualRev:Math.round(curActualRev),units:curUnits,actualUnits:skuForecasts.reduce((s,f)=>s+f.current.actual,0)},futureMonths,daysPassed,totalDays,isPartial,pctMonth:daysPassed/totalDays,prevRev:prevStats.rev,moM,channelGrowthRate,chVelocity,chShare,validation:{match:true}};
  }, [mi, inv, fy, cm, cy]);

  const projection = skuProjection;
  const skuMovement = useMemo(() => skuProjection.skus.map(s => ({ sku:s.sku,velocity:s.velocity,bestPrice:s.bestPrice,combinedGrowth:s.combinedGrowth,hist:s.hist,current:s.current,future:s.future,avgQty:Math.round(s.hist.length?s.hist.reduce((a,m)=>a+m.qty,0)/s.hist.length:0) })), [skuProjection]);
  const curStats = useMemo(() => getMonthStats(mi), [mi, inv, fy]);
  const chData = useMemo(() => CH_IDS.map(c=>({id:c,...CH_MAP[c],rev:curStats.byCh[c]||0})).filter(c=>c.rev>0).sort((a,b)=>b.rev-a.rev), [curStats]);
  const topCusts = useMemo(() => Object.entries(curStats.byCust).map(([name,d])=>({name,...d})).sort((a,b)=>b.rev-a.rev), [curStats]);
  const skuData = useMemo(() => Object.entries(curStats.bySKU).map(([sku,d])=>({sku,...d})).sort((a,b)=>b.rev-a.rev), [curStats]);
  const cityData = useMemo(() => Object.values(curStats.byCity||{}).sort((a,b)=>b.rev-a.rev), [curStats]);
  const regionData = useMemo(() => Object.values(curStats.byRegion||{}).sort((a,b)=>b.rev-a.rev), [curStats]);
  const totalRev=curStats.rev, totalUnits=curStats.units, totalOrders=curStats.orders;

  const recurringCusts = useMemo(() => {
    const cm2={};
    inv.forEach(iv=>{
      const d=new Date(iv.date); if(isNaN(d))return;
      const mk=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
      (iv.items||[]).forEach(it=>{
        const n=(it.custName||"").trim(); if(!n||n.length<2)return;
        const rev=safeNum(it.qty)*safeNum(it.price); if(!rev)return;
        if(!cm2[n])cm2[n]={months:new Set(),totalRev:0};
        cm2[n].months.add(mk); cm2[n].totalRev+=rev;
      });
    });
    return Object.entries(cm2).filter(([,d])=>d.months.size>=2).map(([name,d])=>({name,monthCount:d.months.size,totalRev:d.totalRev,avgMonthlyRev:d.totalRev/d.months.size})).sort((a,b)=>b.monthCount-a.monthCount||b.totalRev-a.totalRev);
  }, [inv]);

  const maData = useMemo(() => {
    const dayMap={};
    inv.forEach(iv=>{
      if(!iv.date)return;
      const chId=iv.channel;
      if(!dayMap[iv.date])dayMap[iv.date]={date:iv.date,total:0,...Object.fromEntries(CH_IDS.map(c=>[c,0]))};
      (iv.items||[]).forEach(it=>{const rev=safeNum(it.qty)*safeNum(it.price);dayMap[iv.date].total+=rev;dayMap[iv.date][chId]=(dayMap[iv.date][chId]||0)+rev;});
    });
    const days=Object.values(dayMap).sort((a,b)=>a.date.localeCompare(b.date));
    if(days.length<2)return{daily:[],chMA:{},hasData:false};
    const W=7,M=30;
    const daily=days.map((d,i)=>{
      const wSlice=days.slice(Math.max(0,i-W+1),i+1);
      const mSlice=days.slice(Math.max(0,i-M+1),i+1);
      const wma=wSlice.reduce((s,x)=>s+x.total,0)/wSlice.length;
      const mma=mSlice.reduce((s,x)=>s+x.total,0)/mSlice.length;
      return{date:d.date,label:d.date.slice(5),total:Math.round(d.total),wma:Math.round(wma),mma:Math.round(mma),ratio:mma>0?Math.round(wma/mma*1000)/1000:1};
    });
    const chMA={};
    CH_IDS.forEach(c=>{
      const chDays=days.map((d,i)=>{
        const sl=days.slice(Math.max(0,i-W+1),i+1);
        const wma=sl.reduce((s,x)=>s+(x[c]||0),0)/sl.length;
        const ml=days.slice(Math.max(0,i-M+1),i+1);
        const mma=ml.reduce((s,x)=>s+(x[c]||0),0)/ml.length;
        return{label:d.date.slice(5),wma:Math.round(wma),mma:Math.round(mma)};
      });
      if(chDays.some(d=>d.wma>0))chMA[c]=chDays;
    });
    return{daily,chMA,hasData:true};
  }, [inv]);

  const tS=CH_IDS.reduce((s,k)=>s+(dt[k]||0),0);
  const cogs=(dt.units_sold||0)*(dt.cost_per_unit||22);
  const gp=tS-cogs;
  const vc=(dt.packaging||0)+(dt.marketplace_fees||0)+(dt.courier||0);
  const cm1=gp-vc, cm2=cm1-(dt.marketing||0);
  const tOp=OPEX_KEYS.reduce((s,k)=>s+(dt[k]||0),0);
  const ebitda=cm2-tOp;
  const roas=dt.marketing>0?tS/dt.marketing:0;

  if (ld) return <div style={{ display:"flex", alignItems:"center", justifyContent:"center", minHeight:"100vh", color:TH.muted, fontFamily:"'IBM Plex Sans',system-ui,sans-serif" }}>Loading...</div>;

  const CSS = `
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
    * { box-sizing: border-box; }
    body { margin: 0; padding: 0; background: #F8FAFC; }
    input[type=number]::-webkit-inner-spin-button { -webkit-appearance: none; }
    input[type=number] { -moz-appearance: textfield; }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #F1F5F9; }
    ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 3px; }
  `;

  const tabs = [
    { id:"dashboard", l:"Dashboard" }, { id:"upload", l:"Upload data" },
    { id:"sales", l:"Sales ops" }, { id:"input", l:"MIS input" }, { id:"pl", l:"P&L" },
  ];

  return (
    <div style={{ fontFamily:"'IBM Plex Sans',system-ui,sans-serif", maxWidth:1280, margin:"0 auto", padding:"16px", color:TH.navy }}>
      <style>{CSS}</style>
      {toast && <Toast msg={toast} type={tType} onClose={() => sToast(null)}/>}

      {/* Header */}
      <div style={{ background:TH.navy, padding:"14px 22px", borderRadius:"12px 12px 0 0", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div>
          <div style={{ color:"#fff", fontSize:18, fontWeight:700, letterSpacing:"-0.02em" }}>Plant Essentials</div>
          <div style={{ color:"#64748B", fontSize:11, marginTop:1 }}>Investor MIS — Executive Dashboard</div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <select value={fy} onChange={e=>sFy(e.target.value)} style={{ padding:"6px 12px", borderRadius:6, border:"1px solid #334155", background:"#1E293B", color:"#94A3B8", fontSize:12, fontWeight:600 }}>
            {FY_OPTIONS.map(f => <option key={f} value={f}>{FY_LABELS[f]}</option>)}
          </select>
          <button onClick={()=>sShowDP(!showDP)} style={{ padding:"6px 14px", border:"1px solid #334155", borderRadius:6, background:"transparent", color:"#94A3B8", fontSize:11, fontWeight:600, cursor:"pointer" }}>
            {showDP ? "✕" : "Export / Import"}
          </button>
          {vw==="dashboard" && <button onClick={exportExcel} style={{ padding:"6px 14px", border:"1px solid #334155", borderRadius:6, background:"transparent", color:"#94A3B8", fontSize:11, fontWeight:600, cursor:"pointer" }}>📊 Excel</button>}
        </div>
      </div>

      {/* Tabs */}
      <div style={{ display:"flex", gap:0, background:"#E2E8F0", borderBottom:`2px solid ${TH.navy}`, overflowX:"auto" }}>
        {tabs.map(t => <TabBtn key={t.id} active={vw===t.id} onClick={()=>sVw(t.id)}>{t.l}</TabBtn>)}
      </div>

      {/* Export/Import Panel */}
      {showDP && (
        <div style={{ background:TH.card, border:`1px solid ${TH.border}`, padding:20, display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:16 }}>
          <div style={{ padding:20, background:TH.bg, borderRadius:10, textAlign:"center" }}>
            <div style={{ fontSize:13, fontWeight:700, marginBottom:8 }}>Export JSON</div>
            <div style={{ fontSize:11, color:TH.muted, marginBottom:12 }}>All invoices, MIS data, mappings</div>
            <button onClick={exportData} style={{ padding:"8px 20px", background:TH.blue, color:"#fff", border:"none", borderRadius:6, fontSize:12, fontWeight:600, cursor:"pointer" }}>Download</button>
          </div>
          <div style={{ padding:20, background:"#F0FDF4", borderRadius:10, textAlign:"center" }}>
            <div style={{ fontSize:13, fontWeight:700, marginBottom:8 }}>Import JSON</div>
            <div style={{ fontSize:11, color:TH.muted, marginBottom:12 }}>Load exported data</div>
            <button onClick={()=>importRef.current?.click()} style={{ padding:"8px 20px", background:TH.green, color:"#fff", border:"none", borderRadius:6, fontSize:12, fontWeight:600, cursor:"pointer" }}>Import</button>
            <input ref={importRef} type="file" accept=".json" onChange={importData} style={{ display:"none" }}/>
          </div>
          <div style={{ padding:20, background:"#FFFBEB", borderRadius:10, textAlign:"center" }}>
            <div style={{ fontSize:13, fontWeight:700, marginBottom:8 }}>Share</div>
            <div style={{ fontSize:11, color:TH.muted, marginBottom:12 }}>Export JSON then share</div>
            <div style={{ fontSize:11, color:TH.amber, fontWeight:600 }}>JSON export = full handoff</div>
          </div>
        </div>
      )}

      {/* Main Content */}
      <div style={{ background:TH.bg, padding:"16px 22px", minHeight:420 }}>
        {/* Month Selector */}
        <div style={{ display:"flex", alignItems:"center", gap:5, marginBottom:16, flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:TH.muted, marginRight:6 }}>Period:</span>
          {MONTHS.map((m,i) => {
            const has = inv.some(iv=>{const d=new Date(iv.date);return d.getMonth()+1===getCalMonth(i)&&d.getFullYear()===getCalYear(i,fy);});
            return <MonthPill key={m} label={m} active={i===mi} hasData={has} onClick={()=>sMi(i)}/>;
          })}
        </div>

        {/* ── DASHBOARD ── */}
        {vw==="dashboard" && <div>
          <div style={{ display:"flex", gap:12, flexWrap:"wrap", marginBottom:16 }}>
            <Metric label="Revenue" value={fmt(totalRev)} sub={projection.moM!=null?(projection.moM>=0?"↑ ":"↓ ")+pc(Math.abs(projection.moM))+" MoM":"First month"} trend={projection.moM>0?"up":projection.moM<0?"down":null}/>
            <Metric label="Units sold" value={fN(totalUnits)} sub={`${totalOrders} orders`}/>
            <Metric label="Gross margin" value={tS?pcA(gp/tS):"–"} sub={`GP ${fmt(gp)}`} trend={gp>0?"up":"down"}/>
            <Metric label="EBITDA" value={fmt(ebitda)} sub={tS?pcA(ebitda/tS)+" margin":"–"} trend={ebitda>=0?"up":"down"}/>
            <Metric label="COGS / unit" value={"₹"+(dt.cost_per_unit||22)} sub={`COGS ${fmt(cogs)}`}/>
            <Metric label="Projected revenue" value={fmt(projection.currentMonth.rev)} sub={projection.isPartial?`${projection.daysPassed}/${projection.totalDays} days — SKU-derived`:"Full month"} trend={projection.currentMonth.rev>projection.prevRev?"up":"down"}/>
          </div>

          <SectionCard title="Sales targets — n+1, n+2, n+3 months">
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:12, marginBottom:14 }}>
              <div style={{ padding:14, background:TH.bg, borderRadius:10, border:`1px solid ${TH.border}`, textAlign:"center" }}>
                <div style={{ fontSize:11, color:TH.muted, fontWeight:600, marginBottom:6 }}>Current month</div>
                <div style={{ fontSize:22, fontWeight:800, fontFamily:"'JetBrains Mono',monospace", color:TH.navy }}>{fmt(projection.currentMonth.actualRev)}</div>
                <div style={{ fontSize:11, color:TH.muted, marginTop:4 }}>{projection.daysPassed}/{projection.totalDays} days</div>
              </div>
              {projection.futureMonths.map((fm,i)=>{
                const colors=["#2563EB","#059669","#7C3AED"]; const bgs=["#EFF6FF","#F0FDF4","#FDF4FF"]; const borders=["#BFDBFE","#BBF7D0","#E9D5FF"];
                return <div key={i} style={{ padding:14, background:bgs[i], borderRadius:10, border:`1px solid ${borders[i]}`, textAlign:"center" }}>
                  <div style={{ fontSize:11, color:colors[i], fontWeight:700, marginBottom:4, textTransform:"uppercase", letterSpacing:"0.04em" }}>{fm.monthLabel} — {fm.month}</div>
                  <div style={{ fontSize:22, fontWeight:800, fontFamily:"'JetBrains Mono',monospace", color:colors[i] }}>{fmt(fm.rev)}</div>
                  <div style={{ fontSize:11, color:TH.muted, marginTop:2 }}>{fN(fm.units)} units</div>
                  <div style={{ fontSize:12, fontWeight:700, color:colors[i], marginTop:6 }}>↑ {(Math.abs(fm.growthVsBase||0)*(i+1)*100).toFixed(1)}% vs now</div>
                </div>;
              })}
            </div>
            <div style={{ fontSize:11, color:TH.muted, padding:"8px 12px", background:TH.bg, borderRadius:6 }}>
              Combined growth = 50% SKU velocity + 50% channel momentum. Floor 3%/month. | Channel rate: {((projection.channelGrowthRate||0)*100).toFixed(1)}%/mo
            </div>
          </SectionCard>

          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
            <SectionCard title="Revenue by channel">
              {chData.length>0 ? <div>{chData.map(c=>{const pct=totalRev?c.rev/totalRev:0; return <div key={c.id} style={{ display:"flex", alignItems:"center", gap:10, marginBottom:8 }}>
                <span style={{ fontSize:14, width:22 }}>{c.icon}</span>
                <span style={{ flex:1, fontSize:12, fontWeight:600 }}>{c.name}</span>
                <div style={{ width:120, height:7, background:TH.bg, borderRadius:4, overflow:"hidden" }}><div style={{ width:`${pct*100}%`, height:"100%", background:c.color, borderRadius:4 }}/></div>
                <span style={{ fontSize:12, fontWeight:700, fontFamily:"'JetBrains Mono',monospace", width:80, textAlign:"right" }}>{fmt(c.rev)}</span>
                <span style={{ fontSize:11, color:TH.muted, width:40, textAlign:"right" }}>{Math.round(pct*100)}%</span>
              </div>;})}
              </div> : <div style={{ padding:20, textAlign:"center", color:TH.muted }}>No data for {MONTHS[mi]}</div>}
            </SectionCard>
            <SectionCard title="Top customers" noPad>
              {topCusts.length>0 ? <DataTable compact headers={["Customer","Units","Revenue","%"]} rows={topCusts.slice(0,10).map(c=>[c.name,fN(c.units),fmt(c.rev),totalRev?Math.round(c.rev/totalRev*100)+"%":"0%"])}/>
              : <div style={{ padding:24, textAlign:"center", color:TH.muted, fontSize:12 }}>Map "Customer Name" when uploading</div>}
            </SectionCard>
          </div>

          <SectionCard title="Revenue trend — channel × month matrix">
            {last3.length>0 ? <div style={{ overflowX:"auto" }}>
              {(()=>{
                const allMonths=[...last3.map(m=>({month:m.month,byCh:m.byCh,rev:m.rev}))];
                if(projection.currentMonth) allMonths.push({month:MONTHS[mi]+(projection.isPartial?" *":""),byCh:curStats.byCh,rev:projection.currentMonth.rev});
                projection.futureMonths.forEach(fm=>{allMonths.push({month:fm.month,byCh:fm.byCh||{},rev:fm.rev});});
                return <DataTable headers={["Channel",...allMonths.map(m=>m.month),"Total"]} rows={[
                  ...CHANNELS.map(ch=>{const vals=allMonths.map(m=>m.byCh[ch.id]||0);const tot=vals.reduce((s,v)=>s+v,0);return tot>0?[ch.name,...vals.map(v=>v?fmt(v):"–"),fmt(tot)]:null;}).filter(Boolean),
                  ["TOTAL",...allMonths.map(m=>fmt(m.rev)),fmt(allMonths.reduce((s,m)=>s+m.rev,0))],
                ]}/>;
              })()}
            </div> : <div style={{ padding:30, textAlign:"center", color:TH.muted }}>Upload 2+ months of data</div>}
          </SectionCard>

          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
            <SectionCard title="SKU performance — actual + projected (units)" noPad>
              {skuMovement.length>0 ? <DataTable compact headers={["SKU","Velocity","Price",...(skuMovement[0]?.hist||[]).map(h=>h.month),MONTHS[mi],...(skuMovement[0]?.future||[]).map(f=>f.monthLabel),"Growth"]}
                rows={skuMovement.map(s=>[s.sku,s.velocity,"₹"+fN(s.bestPrice),...s.hist.map(h=>fN(h.qty)),fN(s.current.projected),...s.future.map(f=>fN(f.qty)),(s.combinedGrowth>=0?"+":"")+Math.round(s.combinedGrowth*100)+"%"])}/>
              : <div style={{ padding:24, textAlign:"center", color:TH.muted, fontSize:12 }}>No SKU data</div>}
            </SectionCard>
            <SectionCard title="SKU velocity cards">
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
                {skuMovement.map(s=><div key={s.sku} style={{ padding:10, background:TH.bg, borderRadius:8, border:`1px solid ${TH.border}` }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:3 }}>
                    <span style={{ fontSize:11, fontWeight:700 }}>{s.sku}</span>
                    <ChipBadge text={s.velocity} color={s.velocity==="Fast"?TH.green:s.velocity==="Medium"?TH.amber:TH.red}/>
                  </div>
                  <div style={{ fontSize:17, fontWeight:700, fontFamily:"'JetBrains Mono',monospace" }}>{fN(s.current.projected)}</div>
                  <div style={{ fontSize:10, color:TH.muted }}>Next: {s.future.map(f=>fN(f.qty)).join(" → ")}</div>
                </div>)}
              </div>
            </SectionCard>
          </div>

          {recurringCusts.length>0 && <SectionCard title={`Repeating customers — active across multiple months (${recurringCusts.length} found)`}>
            <div style={{ marginBottom:14, height:200 }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={recurringCusts.slice(0,12).map(c=>({name:c.name.length>14?c.name.slice(0,14)+"…":c.name,months:c.monthCount,avg:Math.round(c.avgMonthlyRev)}))} layout="vertical" margin={{left:10,right:20,top:4,bottom:4}}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" horizontal={false}/>
                  <XAxis type="number" tick={{fontSize:11,fill:TH.muted}} tickFormatter={v=>fmt(v)}/>
                  <YAxis type="category" dataKey="name" tick={{fontSize:11,fill:TH.navy,fontWeight:600}} width={110}/>
                  <Tooltip formatter={(v,n)=>n==="avg"?["₹"+fN(v),"Avg/month"]:[v+" months","Active months"]}/>
                  <Bar dataKey="avg" name="avg" radius={[0,4,4,0]}>
                    {recurringCusts.slice(0,12).map((c,i)=><Cell key={i} fill={c.monthCount>=4?TH.green:c.monthCount===3?TH.blue:TH.amber}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <DataTable compact headers={["Customer","Active months","Avg / month","Total revenue"]} rows={recurringCusts.slice(0,12).map(c=>[c.name,c.monthCount+" months",fmt(c.avgMonthlyRev),fmt(c.totalRev)])}/>
          </SectionCard>}

          {maData.hasData && <div>
            {(()=>{const last=maData.daily[maData.daily.length-1];if(!last)return null;return <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:16}}>
              <Metric small label="7-day MA (STMA)" value={fmt(last.wma)} sub="Short-term trend" trend={last.ratio>=1?"up":"down"}/>
              <Metric small label="30-day MA (LTMA)" value={fmt(last.mma)} sub="Long-term baseline"/>
              <Metric small label="MA ratio" value={last.ratio.toFixed(3)} sub={last.ratio>1.02?"↑ Accelerating":last.ratio<0.98?"↓ Decelerating":"→ Stable"} trend={last.ratio>=1?"up":"down"}/>
              <Metric small label="Data points" value={maData.daily.length} sub="Invoice days tracked"/>
            </div>;})()} 
            <SectionCard title="Daily revenue with moving averages — 7-day STMA vs 30-day LTMA">
              <div style={{height:240,marginBottom:8}}>
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={maData.daily} margin={{left:4,right:4,top:4,bottom:4}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0"/>
                    <XAxis dataKey="label" tick={{fontSize:9,fill:TH.muted}} interval={Math.max(1,Math.floor(maData.daily.length/14))}/>
                    <YAxis tick={{fontSize:10,fill:TH.muted}} tickFormatter={v=>fmt(v)} width={52}/>
                    <Tooltip formatter={(v,n)=>typeof v==="number"?["₹"+fN(v),n]:v} labelFormatter={l=>"Date: "+l}/>
                    <Bar dataKey="total" fill={TH.border} radius={[2,2,0,0]} name="Daily revenue" opacity={0.5}/>
                    <Line dataKey="wma" stroke={TH.green} strokeWidth={2} dot={false} name="7-day MA"/>
                    <Line dataKey="mma" stroke={TH.red} strokeWidth={2} dot={false} name="30-day MA" strokeDasharray="6 3"/>
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </SectionCard>
            <SectionCard title="MA ratio trend — STMA ÷ LTMA (above 1.0 = growth acceleration)">
              <div style={{height:160}}>
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={maData.daily} margin={{left:4,right:4,top:4,bottom:4}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0"/>
                    <XAxis dataKey="label" tick={{fontSize:9,fill:TH.muted}} interval={Math.max(1,Math.floor(maData.daily.length/14))}/>
                    <YAxis tick={{fontSize:10,fill:TH.muted}} tickFormatter={v=>v.toFixed(2)} domain={["auto","auto"]} width={44}/>
                    <Tooltip formatter={(v,n)=>typeof v==="number"?[v.toFixed(3),n]:v}/>
                    <ReferenceLine y={1} stroke={TH.amber} strokeDasharray="4 4" label={{value:"1.0",fill:TH.amber,fontSize:10,position:"right"}}/>
                    <Line dataKey="ratio" stroke="#7C3AED" strokeWidth={2} dot={false} name="MA ratio"/>
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </SectionCard>
          </div>}

          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
            <SectionCard title="Top cities by revenue" noPad>
              {cityData.length>0&&cityData[0].city!=="Unknown" ? <div style={{padding:"14px 18px 0"}}>
                {cityData.slice(0,8).map((c,i)=>{const pctV=totalRev?c.rev/totalRev:0; return <div key={c.city} style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
                  <span style={{fontSize:12,color:TH.muted,width:18,textAlign:"right",fontWeight:600}}>{i+1}</span>
                  <span style={{flex:1,fontSize:12,fontWeight:600}}>{c.city}</span>
                  <ChipBadge text={c.region} color={c.region==="South"?TH.green:c.region==="North"?TH.blue:c.region==="West"?TH.amber:"#7C3AED"}/>
                  <div style={{width:100,height:6,background:TH.bg,borderRadius:3,overflow:"hidden"}}><div style={{width:`${pctV*100}%`,height:"100%",background:TH.blue,borderRadius:3}}/></div>
                  <span style={{fontSize:12,fontWeight:700,fontFamily:"'JetBrains Mono',monospace",width:70,textAlign:"right"}}>{fmt(c.rev)}</span>
                </div>;})}
              </div> : <div style={{padding:28,textAlign:"center",color:TH.muted,fontSize:12}}>Map "City" column to see geo insights</div>}
            </SectionCard>
            <SectionCard title="Revenue by region">
              {regionData.length>0&&!(regionData.length===1&&regionData[0].region==="Unknown") ? <div>
                <ResponsiveContainer width="100%" height={160}>
                  <PieChart>
                    <Pie data={regionData.filter(r=>r.region!=="Unknown")} dataKey="rev" nameKey="region" cx="50%" cy="50%" outerRadius={60} innerRadius={30} paddingAngle={2} stroke="none">
                      {regionData.filter(r=>r.region!=="Unknown").map((r,i)=><Cell key={i} fill={r.region==="South"?TH.green:r.region==="North"?TH.blue:r.region==="West"?TH.amber:"#7C3AED"}/>)}
                    </Pie>
                    <Tooltip formatter={v=>"₹"+fN(v)}/>
                  </PieChart>
                </ResponsiveContainer>
              </div> : <div style={{padding:28,textAlign:"center",color:TH.muted,fontSize:12}}>No location data</div>}
            </SectionCard>
          </div>
        </div>}

        {/* ── UPLOAD ── */}
        {vw==="upload" && <div>
          {uStep==="select" && <div>
            <div style={{fontSize:15,fontWeight:700,marginBottom:4}}>Upload channel sales data</div>
            <div style={{fontSize:12,color:TH.muted,marginBottom:16}}>Select channel, upload Excel, map columns, generate invoices.</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:12,marginBottom:20}}>
              {CHANNELS.map(ch=><div key={ch.id} onClick={()=>sUChan(ch)} style={{borderRadius:10,padding:16,cursor:"pointer",border:`2px solid ${uChan?.id===ch.id?ch.color:"transparent"}`,background:ch.bg,textAlign:"center"}}>
                <div style={{fontSize:26,marginBottom:6}}>{ch.icon}</div>
                <div style={{fontSize:13,fontWeight:700,color:ch.color}}>{ch.name}</div>
                <div style={{fontSize:10,color:TH.muted,marginTop:3}}>{ch.desc}</div>
                {sMaps[ch.id]&&<div style={{fontSize:9,marginTop:6,background:"rgba(0,0,0,0.06)",padding:"2px 6px",borderRadius:3,display:"inline-block"}}>Mapped</div>}
              </div>)}
            </div>
            {uChan && <div style={{background:TH.card,borderRadius:10,padding:20,border:`1px solid ${TH.border}`}}>
              <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:14}}>
                <span style={{fontSize:22}}>{uChan.icon}</span>
                <div><div style={{fontSize:14,fontWeight:700}}>Upload {uChan.name} report</div><div style={{fontSize:11,color:TH.muted}}>.xlsx, .xls, .csv</div></div>
              </div>
              <div style={{border:`2px dashed ${TH.border}`,borderRadius:10,padding:28,textAlign:"center",background:TH.bg,cursor:"pointer"}} onClick={()=>fRef.current?.click()}>
                <div style={{fontSize:24,marginBottom:6}}>📄</div>
                <div style={{fontSize:13,fontWeight:600}}>Click to browse</div>
              </div>
              <input ref={fRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{display:"none"}}/>
            </div>}
          </div>}

          {uStep==="map" && <SectionCard title={`Map columns — ${uChan?.name} (${fName})`}>
            <div style={{fontSize:11,color:TH.muted,marginBottom:14}}>{rD.length} rows. <span style={{color:TH.red}}>*</span> = required.</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
              {SFIELDS.map(f=><div key={f.key} style={{display:"flex",flexDirection:"column",gap:3}}>
                <label style={{fontSize:11,fontWeight:600,color:TH.muted}}>{f.label}{f.req&&<span style={{color:TH.red}}> *</span>}</label>
                <select value={cMap[f.key]||""} onChange={e=>sCMap(p=>({...p,[f.key]:e.target.value||undefined}))} style={{padding:"8px 10px",border:`1px solid ${cMap[f.key]?TH.green:TH.border}`,borderRadius:6,fontSize:12,background:cMap[f.key]?"#F0FDF4":"#fff"}}>
                  <option value="">— Not mapped —</option>
                  {rH.map(h=><option key={h} value={h}>{h}</option>)}
                </select>
              </div>)}
            </div>
            <div style={{display:"flex",justifyContent:"flex-end",gap:10}}>
              <button onClick={resetUp} style={{padding:"10px 24px",border:`1px solid ${TH.border}`,borderRadius:8,background:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>Cancel</button>
              <button onClick={applyMap} style={{padding:"10px 24px",background:TH.navy,color:"#fff",border:"none",borderRadius:8,fontSize:13,fontWeight:700,cursor:"pointer"}}>Preview →</button>
            </div>
          </SectionCard>}

          {uStep==="preview" && <div style={{background:TH.card,borderRadius:10,border:`1px solid ${TH.border}`,overflow:"hidden"}}>
            <div style={{background:TH.navy,padding:"12px 18px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{color:"#fff",fontSize:14,fontWeight:700}}>Review — {uChan?.name}</div><div style={{color:"#64748B",fontSize:11}}>{pRows.length} items</div></div>
              <div style={{display:"flex",gap:6}}>
                <button onClick={()=>sUStep("map")} style={{padding:"6px 14px",border:"1px solid #334155",borderRadius:6,background:"transparent",color:"#94A3B8",fontSize:12,cursor:"pointer"}}>← Re-map</button>
                <button onClick={resetUp} style={{padding:"6px 14px",border:"1px solid #334155",borderRadius:6,background:"transparent",color:"#94A3B8",fontSize:12,cursor:"pointer"}}>✕</button>
              </div>
            </div>
            <div style={{padding:18}}>
              <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}>
                <Metric small label="Rows" value={pRows.length}/>
                <Metric small label="Revenue" value={fmt(pRows.reduce((s,r)=>s+r.total,0))}/>
              </div>
              <div style={{overflowX:"auto",maxHeight:280,border:`1px solid ${TH.border}`,borderRadius:8,marginBottom:14}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                  <thead><tr>{["Date","Product","SKU","Customer","City","Qty","Rate","Tax","Total"].map((h,i)=><th key={i} style={{padding:"6px 10px",textAlign:i>=5?"right":"left",fontWeight:600,color:TH.muted,fontSize:10,textTransform:"uppercase",borderBottom:`2px solid ${TH.border}`,background:TH.bg,position:"sticky",top:0}}>{h}</th>)}</tr></thead>
                  <tbody>{pRows.map((r,i)=>{const fs=skuOv[r.idx]||r.sku; return <tr key={i} style={{background:fs?"transparent":"#FEF2F2"}}>
                    <td style={{padding:"5px 10px",borderBottom:`1px solid ${TH.border}`,whiteSpace:"nowrap"}}>{r.date||<span style={{color:TH.red}}>Invalid</span>}</td>
                    <td style={{padding:"5px 10px",borderBottom:`1px solid ${TH.border}`,maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.skuRaw}</td>
                    <td style={{padding:"5px 10px",borderBottom:`1px solid ${TH.border}`}}>{fs?<span style={{color:TH.green,fontWeight:600}}>{fs}</span>:<select value="" onChange={e=>sSkuOv(p=>({...p,[r.idx]:e.target.value}))} style={{padding:3,border:`1px solid ${TH.red}`,borderRadius:3,fontSize:10,background:"#FEF2F2"}}><option value="">Pick...</option>{SKUS.map(s=><option key={s} value={s}>{s}</option>)}</select>}</td>
                    <td style={{padding:"5px 10px",borderBottom:`1px solid ${TH.border}`,color:r.custName?TH.navy:TH.muted,fontSize:11}}>{r.custName||"–"}</td>
                    <td style={{padding:"5px 10px",borderBottom:`1px solid ${TH.border}`,fontSize:11}}>{r.city&&r.city!=="Unknown"?r.city:<span style={{color:TH.muted}}>–</span>}</td>
                    <td style={{padding:"5px 10px",borderBottom:`1px solid ${TH.border}`,textAlign:"right",fontFamily:"'JetBrains Mono',monospace"}}>{r.qty}</td>
                    <td style={{padding:"5px 10px",borderBottom:`1px solid ${TH.border}`,textAlign:"right",fontFamily:"'JetBrains Mono',monospace"}}>₹{r.unitPrice.toFixed(2)}</td>
                    <td style={{padding:"5px 10px",borderBottom:`1px solid ${TH.border}`,textAlign:"right",fontFamily:"'JetBrains Mono',monospace"}}>{r.gst>0?"₹"+r.gst.toFixed(0):"–"}</td>
                    <td style={{padding:"5px 10px",borderBottom:`1px solid ${TH.border}`,textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontWeight:600}}>₹{fN(r.total)}</td>
                  </tr>;})}
                  </tbody>
                </table>
              </div>
              <div style={{display:"flex",justifyContent:"flex-end"}}><button onClick={genInv} style={{padding:"12px 28px",background:TH.green,color:"#fff",border:"none",borderRadius:8,fontSize:14,fontWeight:700,cursor:"pointer"}}>Generate invoices</button></div>
            </div>
          </div>}

          {uStep==="done" && <div style={{textAlign:"center",padding:40,background:TH.card,borderRadius:10,border:`1px solid ${TH.border}`}}>
            <div style={{fontSize:40,marginBottom:10}}>✅</div>
            <div style={{fontSize:17,fontWeight:700,marginBottom:4}}>Invoices generated</div>
            <div style={{fontSize:13,color:TH.green,fontWeight:600,marginBottom:18}}>Auto-detected: {MONTHS[mi]} {FY_LABELS[fy]}</div>
            <div style={{display:"flex",gap:10,justifyContent:"center"}}>
              <button onClick={resetUp} style={{padding:"10px 24px",border:`1px solid ${TH.border}`,borderRadius:8,background:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>Upload more</button>
              <button onClick={()=>{resetUp();sVw("sales");sSvm("ledger");}} style={{padding:"10px 24px",background:TH.navy,color:"#fff",border:"none",borderRadius:8,fontSize:13,fontWeight:700,cursor:"pointer"}}>View invoices</button>
              <button onClick={()=>{resetUp();sVw("dashboard");}} style={{padding:"10px 24px",background:TH.green,color:"#fff",border:"none",borderRadius:8,fontSize:13,fontWeight:700,cursor:"pointer"}}>Dashboard</button>
            </div>
          </div>}
        </div>}

        {/* ── SALES OPS ── */}
        {vw==="sales" && <div>
          <div style={{display:"flex",gap:8,marginBottom:14}}>
            {["ledger","create","report"].map(t=><button key={t} onClick={()=>sSvm(t)} style={{padding:"6px 16px",border:"none",cursor:"pointer",fontSize:12,fontWeight:600,borderRadius:6,background:svm===t?TH.navy:"#E2E8F0",color:svm===t?"#fff":TH.muted}}>
              {t==="ledger"?"Ledger":t==="create"?"New invoice":"Weekly report"}
            </button>)}
          </div>
          {svm==="ledger" && <SectionCard title={`${MONTHS[mi]} — invoice ledger`} noPad>
            {!mInv.length ? <div style={{padding:30,textAlign:"center",color:TH.muted}}>No invoices for {MONTHS[mi]}</div>
            : <DataTable headers={["Invoice","Date","Channel","SKUs","Units","Amount",""]}
                rows={[...mInv].sort((a,b)=>new Date(b.date)-new Date(a.date)).map(iv=>[
                  iv.id,
                  new Date(iv.date).toLocaleDateString("en-IN",{day:"2-digit",month:"short"}),
                  CH_MAP[iv.channel]?.name||iv.channel,
                  (iv.items||[]).map(it=>`${it.sku}×${it.qty}`).join(", "),
                  fN(iv.units), "₹"+fN(iv.subtotal),
                  React.createElement("button",{onClick:()=>delInv(iv.id),style:{width:20,height:20,borderRadius:"50%",border:"1px solid #ddd",background:"#fff",cursor:"pointer",fontSize:10,color:"#c00"}},"✕")
                ])}/>}
          </SectionCard>}
          {svm==="create" && <SectionCard title="Raise invoice">
            <div style={{display:"flex",gap:14,marginBottom:18,flexWrap:"wrap"}}>
              <div style={{flex:"1 1 180px"}}><label style={{fontSize:11,fontWeight:600,color:TH.muted}}>Date</label><input type="date" value={iDate} onChange={e=>hDC(e.target.value)} style={{display:"block",padding:"9px 10px",border:`1px solid ${TH.border}`,borderRadius:6,fontSize:13,width:"100%",marginTop:4}}/></div>
              <div style={{flex:"1 1 220px"}}><label style={{fontSize:11,fontWeight:600,color:TH.muted}}>Channel</label><select value={iChan} onChange={e=>sIChan(e.target.value)} style={{display:"block",padding:"9px 10px",border:`1px solid ${TH.border}`,borderRadius:6,fontSize:13,width:"100%",marginTop:4,background:"#fff"}}>
                {CHANNELS.map(c=><option key={c.id} value={c.id}>{c.icon} {c.name}</option>)}
              </select></div>
            </div>
            {iItems.map((it,idx)=><div key={idx} style={{display:"flex",gap:8,alignItems:"center",padding:"8px 10px",background:TH.bg,borderRadius:8,marginBottom:6,border:`1px solid ${TH.border}`}}>
              <div style={{flex:"1 1 120px"}}><select value={it.sku} onChange={e=>{const n=[...iItems];n[idx]={...n[idx],sku:e.target.value};sIItems(n);}} style={{padding:7,border:`1px solid ${TH.border}`,borderRadius:6,fontSize:12,width:"100%",background:"#fff"}}>{SKUS.map(s=><option key={s} value={s}>{s}</option>)}</select></div>
              <div style={{flex:"0 0 70px"}}><input type="number" placeholder="Qty" value={it.qty||""} onChange={e=>{const n=[...iItems];n[idx]={...n[idx],qty:parseInt(e.target.value)||0};sIItems(n);}} style={{padding:7,border:`1px solid ${TH.border}`,borderRadius:6,fontSize:12,width:"100%"}}/></div>
              <div style={{flex:"0 0 90px"}}><input type="number" placeholder="Price" value={it.price||""} onChange={e=>{const n=[...iItems];n[idx]={...n[idx],price:parseFloat(e.target.value)||0};sIItems(n);}} style={{padding:7,border:`1px solid ${TH.border}`,borderRadius:6,fontSize:12,width:"100%"}}/></div>
              <div style={{flex:"1 1 100px"}}><input placeholder="Customer" value={it.custName||""} onChange={e=>{const n=[...iItems];n[idx]={...n[idx],custName:e.target.value};sIItems(n);}} style={{padding:7,border:`1px solid ${TH.border}`,borderRadius:6,fontSize:12,width:"100%"}}/></div>
              <div style={{flex:"0 0 80px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:13,fontWeight:700}}>₹{fN(iLT(it))}</div>
              {iItems.length>1 && <button onClick={()=>{const n=[...iItems];n.splice(idx,1);sIItems(n);}} style={{width:20,height:20,borderRadius:"50%",border:"1px solid #ddd",background:"#fff",cursor:"pointer",fontSize:10,color:"#c00"}}>✕</button>}
            </div>)}
            <button onClick={()=>sIItems([...iItems,{sku:"Millet",qty:0,price:0,custName:"",city:"",region:""}])} style={{padding:"6px 14px",background:TH.bg,border:`1px dashed ${TH.border}`,borderRadius:6,fontSize:12,fontWeight:600,cursor:"pointer",color:TH.blue,marginBottom:14}}>+ Add item</button>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 16px",background:TH.bg,borderRadius:8}}>
              <div style={{fontSize:22,fontWeight:700,fontFamily:"'JetBrains Mono',monospace"}}>₹{fN(iSub)}</div>
              <button disabled={iSub<=0} onClick={raiseInv} style={{padding:"10px 24px",background:TH.green,color:"#fff",border:"none",borderRadius:8,fontSize:14,fontWeight:700,cursor:"pointer",opacity:iSub<=0?0.4:1}}>Raise</button>
            </div>
          </SectionCard>}
          {svm==="report" && <div>
            <div style={{display:"flex",gap:6,marginBottom:14}}>
              {[{v:0,l:"All"},{v:1,l:"W1"},{v:2,l:"W2"},{v:3,l:"W3"},{v:4,l:"W4"}].map(w=><button key={w.v} onClick={()=>sSw(w.v)} style={{padding:"5px 12px",border:`1px solid ${sw===w.v?TH.navy:TH.border}`,borderRadius:6,cursor:"pointer",fontSize:12,fontWeight:600,background:sw===w.v?TH.navy:"#fff",color:sw===w.v?"#fff":TH.navy}}>{w.l}</button>)}
            </div>
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:16}}>
              <Metric label="Revenue" value={fmt(wInv.reduce((s,i)=>s+i.subtotal,0))} sub={`${wInv.length} orders`}/>
              <Metric label="Units" value={fN(wInv.reduce((s,i)=>s+i.units,0))}/>
            </div>
          </div>}
        </div>}

        {/* ── MIS INPUT ── */}
        {vw==="input" && <div>
          <Collapsible title="Revenue by channel" icon="📊">
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              {CH_IDS.map(k=><NumInput key={k} label={CH_MAP[k].name} value={dt[k]} onChange={v=>upd(k,v)}/>)}
            </div>
            <div style={{marginTop:12,padding:"10px 14px",background:TH.bg,borderRadius:8,display:"flex",justifyContent:"space-between"}}>
              <span style={{fontSize:13,fontWeight:600}}>Total</span>
              <span style={{fontSize:18,fontWeight:700,fontFamily:"'JetBrains Mono',monospace"}}>₹{fN(tS)}</span>
            </div>
          </Collapsible>
          <Collapsible title="Units & Inventory" icon="📦">
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
              <NumInput label="Units sold" value={dt.units_sold} onChange={v=>upd("units_sold",v)} prefix="#" highlight/>
              <NumInput label="Cost/unit" value={dt.cost_per_unit} onChange={v=>upd("cost_per_unit",v)}/>
              <div style={{display:"flex",flexDirection:"column",gap:3,justifyContent:"flex-end"}}>
                <div style={{fontSize:11,color:TH.muted}}>COGS</div>
                <div style={{padding:"8px 12px",background:"#FEF2F2",borderRadius:6,fontSize:14,fontWeight:700,color:TH.red,fontFamily:"'JetBrains Mono',monospace"}}>₹{fN(cogs)}</div>
              </div>
            </div>
          </Collapsible>
          <Collapsible title="Variable costs" icon="🚚">
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
              <NumInput label="Packaging" value={dt.packaging} onChange={v=>upd("packaging",v)}/>
              <NumInput label="Marketplace fees" value={dt.marketplace_fees} onChange={v=>upd("marketplace_fees",v)}/>
              <NumInput label="Courier" value={dt.courier} onChange={v=>upd("courier",v)}/>
            </div>
          </Collapsible>
          <Collapsible title="Marketing" icon="📣">
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:10}}>
              <NumInput label="Marketing spend" value={dt.marketing} onChange={v=>upd("marketing",v)} highlight/>
              <div style={{display:"flex",flexDirection:"column",gap:3,justifyContent:"flex-end"}}>
                <div style={{fontSize:11,color:TH.muted}}>ROAS</div>
                <div style={{padding:"8px 12px",background:roas<5?"#FEF2F2":"#F0FDF4",borderRadius:6,fontSize:14,fontWeight:700,color:roas<5?TH.red:TH.green,fontFamily:"'JetBrains Mono',monospace"}}>{roas.toFixed(1)}x</div>
              </div>
            </div>
          </Collapsible>
          <Collapsible title="Operating expenses" icon="🏢" open={false}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
              {OPEX_KEYS.map(k=><NumInput key={k} label={OPEX_LABELS[k]} value={dt[k]} onChange={v=>upd(k,v)}/>)}
            </div>
            <div style={{marginTop:10,padding:"10px 14px",background:"#FEF2F2",borderRadius:8,display:"flex",justifyContent:"space-between"}}>
              <span style={{fontSize:13,fontWeight:600,color:TH.red}}>Total OpEx</span>
              <span style={{fontSize:18,fontWeight:700,color:TH.red,fontFamily:"'JetBrains Mono',monospace"}}>₹{fN(tOp)}</span>
            </div>
          </Collapsible>
          <div style={{display:"flex",gap:12,justifyContent:"center",marginTop:18}}>
            <button onClick={saveDt} style={{padding:"10px 24px",border:`2px solid ${TH.navy}`,borderRadius:8,background:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",color:TH.navy}}>{sv?"Saving...":"Save "+MONTHS[mi]}</button>
            <button onClick={()=>{saveDt();sVw("dashboard");}} style={{padding:"10px 24px",background:TH.navy,color:"#fff",border:"none",borderRadius:8,fontSize:13,fontWeight:700,cursor:"pointer"}}>Dashboard →</button>
          </div>
        </div>}

        {/* ── P&L ── */}
        {vw==="pl" && <div style={{background:TH.card,borderRadius:10,border:`1px solid ${TH.border}`,overflow:"hidden"}}>
          <div style={{background:TH.navy,padding:"12px 18px"}}><span style={{color:"#fff",fontSize:14,fontWeight:700}}>P&L — {MONTHS[mi]} {FY_LABELS[fy]}</span></div>
          <div style={{display:"flex",padding:"8px 16px",background:TH.bg,borderBottom:`1px solid ${TH.border}`,fontSize:11,fontWeight:600,color:TH.muted}}>
            <span style={{flex:1}}>Particulars</span><span style={{width:110,textAlign:"right"}}>Amount</span><span style={{width:65,textAlign:"right"}}>% Rev</span>
          </div>
          <PLRow label="REVENUE" value={tS} rev={tS} bold bg="#EFF6FF"/>
          {CH_IDS.filter(k=>(dt[k]||0)>0).map(k=><PLRow key={k} label={CH_MAP[k].name} value={dt[k]} rev={tS} indent={1}/>)}
          <PLRow label="COGS" value={-cogs} rev={tS} bold bg="#FEF2F2"/>
          <PLRow label="GROSS PROFIT" value={gp} rev={tS} bold bg="#F0FDF4"/>
          <PLRow label="Packaging" value={-dt.packaging} rev={tS} indent={1}/>
          <PLRow label="Marketplace" value={-dt.marketplace_fees} rev={tS} indent={1}/>
          <PLRow label="Courier" value={-dt.courier} rev={tS} indent={1}/>
          <PLRow label="CM1" value={cm1} rev={tS} bold bg="#FFFBEB"/>
          <PLRow label="Marketing" value={-dt.marketing} rev={tS} indent={1}/>
          <PLRow label="CM2" value={cm2} rev={tS} bold bg="#FFFBEB"/>
          {OPEX_KEYS.filter(k=>(dt[k]||0)>0).map(k=><PLRow key={k} label={OPEX_LABELS[k]} value={-dt[k]} rev={tS} indent={1}/>)}
          <PLRow label="Total OpEx" value={-tOp} rev={tS} bold bg="#FEF2F2"/>
          <PLRow label="EBITDA" value={ebitda} rev={tS} bold bg={ebitda>=0?"#F0FDF4":"#FEF2F2"}/>
        </div>}
      </div>

      {/* Footer */}
      <div style={{background:TH.navy,padding:"10px 22px",borderRadius:"0 0 12px 12px",display:"flex",justifyContent:"space-between"}}>
        <span style={{color:"#475569",fontSize:11}}>Plant Essentials Pvt Ltd</span>
        <span style={{color:"#475569",fontSize:11}}>{inv.length} invoices | {FY_LABELS[fy]}</span>
      </div>
    </div>
  );
}
