import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Supabase REST Client (no external library) ───────────────────────────
const SUPA_URL = "https://hlksnbrzumzfgjsefxgv.supabase.co";
const SUPA_KEY = "sb_publishable_POxQAK5GFStpapGU4q_aUA_GzRw84bq";

// ─── Simple DB helpers ───────────────────────────────────────────────────
const H = () => ({
  "Content-Type": "application/json",
  "apikey": SUPA_KEY,
  "Authorization": `Bearer ${SUPA_KEY}`,
  "Prefer": "return=representation",
});

const dbSelect = async (table, filters={}, containsFilter=null) => {
  let url = `${SUPA_URL}/rest/v1/${table}?select=*`;
  Object.entries(filters).forEach(([k,v]) => { url += `&${k}=eq.${encodeURIComponent(String(v))}`; });
  if(containsFilter) Object.entries(containsFilter).forEach(([k,v]) => { url += `&${k}=cs.{${v.join(",")}}`; });
  try {
    const res = await fetch(url, {headers: H()});
    const data = await res.json();
    return Array.isArray(data) ? data : [];
  } catch(e) { console.error("dbSelect error",e); return []; }
};

const dbSelectOne = async (table, filters={}) => {
  let url = `${SUPA_URL}/rest/v1/${table}?select=*`;
  Object.entries(filters).forEach(([k,v]) => { url += `&${k}=eq.${encodeURIComponent(String(v))}`; });
  try {
    const res = await fetch(url, {headers: H()});
    const data = await res.json();
    const arr = Array.isArray(data) ? data : [];
    return arr[0] || null;
  } catch(e) { return null; }
};

const dbInsert = async (table, rows) => {
  const body = Array.isArray(rows) ? rows : [rows];
  try {
    const res = await fetch(`${SUPA_URL}/rest/v1/${table}`, {method:"POST", headers:H(), body:JSON.stringify(body)});
    if (!res.ok) {
      const errBody = await res.text().catch(()=>"");
      console.error("dbInsert failed", res.status, errBody);
      return { ok:false, status:res.status, error:errBody };
    }
    return { ok:true };
  } catch(e) {
    console.error("dbInsert network error", e);
    return { ok:false, status:0, error:e.message };
  }
};

const dbUpdate = async (table, vals, filters={}) => {
  let url = `${SUPA_URL}/rest/v1/${table}?`;
  Object.entries(filters).forEach(([k,v]) => { url += `${k}=eq.${encodeURIComponent(String(v))}&`; });
  try {
    const res = await fetch(url, {method:"PATCH", headers:H(), body:JSON.stringify(vals)});
    return res.ok;
  } catch(e) { return false; }
};

const dbDelete = async (table, filters={}) => {
  let url = `${SUPA_URL}/rest/v1/${table}?`;
  Object.entries(filters).forEach(([k,v]) => { url += `${k}=eq.${encodeURIComponent(String(v))}&`; });
  try {
    const res = await fetch(url, {method:"DELETE", headers:H()});
    return res.ok;
  } catch(e) { return false; }
};

// dummy so old sb.channel calls don't break
const sb = {
  from: (table) => ({
    select: (cols="*") => ({
      _table: table, _cols: cols, _filters: [],
      eq: function(col,val){ this._filters.push(`${col}=eq.${val}`); return this; },
      contains: function(col,val){ this._filters.push(`${col}=cs.{${val.join(",")}}`); return this; },
      order: function(col,opts){ this._order=`${col}.${opts?.ascending===false?"desc":"asc"}`; return this; },
      single: async function(){
        const q=this._filters.length?`?${this._filters.join("&")}`:"";
        const res=await fetch(`${SUPA_URL}/rest/v1/${this._table}?select=${this._cols}${q?`&${this._filters.join("&")}`:q}`,{headers:{...sb._headers,"Accept":"application/vnd.pgrst.object+json"}});
        const data=await res.json();
        return res.ok?{data,error:null}:{data:null,error:data};
      },
      then: async function(resolve){
        let url=`${SUPA_URL}/rest/v1/${table}?select=${cols}`;
        if(this._filters.length) url+=`&${this._filters.join("&")}`;
        if(this._order) url+=`&order=${this._order}`;
        const res=await fetch(url,{headers:sb._headers});
        const data=await res.json();
        resolve(res.ok?{data,error:null}:{data:null,error:data});
      },
    }),
    insert: async (rows) => {
      const body=Array.isArray(rows)?rows:[rows];
      const res=await fetch(`${SUPA_URL}/rest/v1/${table}`,{method:"POST",headers:sb._headers,body:JSON.stringify(body)});
      const data=await res.json().catch(()=>[]);
      return res.ok?{data,error:null}:{data:null,error:data};
    },
    update: (vals) => ({
      _table: table, _vals: vals, _filters: [],
      eq: function(col,val){ this._filters.push(`${col}=eq.${val}`); return this; },
      then: async function(resolve){
        const q=this._filters.length?`?${this._filters.join("&")}`:"";
        const res=await fetch(`${SUPA_URL}/rest/v1/${table}${q}`,{method:"PATCH",headers:sb._headers,body:JSON.stringify(vals)});
        const data=await res.json().catch(()=>[]);
        resolve(res.ok?{data,error:null}:{data:null,error:data});
      },
    }),
    delete: () => ({
      _filters: [],
      eq: function(col,val){ this._filters.push(`${col}=eq.${val}`); return this; },
      then: async function(resolve){
        const q=this._filters.length?`?${this._filters.join("&")}`:"";
        const res=await fetch(`${SUPA_URL}/rest/v1/${table}${q}`,{method:"DELETE",headers:sb._headers});
        resolve(res.ok?{data:null,error:null}:{data:null,error:await res.json()});
      },
    }),
  }),
  channel: () => ({ on: () => ({ subscribe: () => ({}) }) }),
  removeChannel: () => {},
};

// ─── Seed Data (runs once if tables empty) ────────────────────────────────
const SEED_SUPERVISORS = [
  { id:"SUP001", name:"أحمد محمود", phone:"0501234567", password_hash:"1234", email:"ahmed@company.com", role:"supervisor" },
  { id:"SUP002", name:"سارة علي",   phone:"0509876543", password_hash:"5678", email:"sara@company.com",  role:"supervisor" },
];
const OPS_MANAGER = { id:"OPS001", name:"مدير التشغيل", phone:"0500000000", password_hash:"admin123", email:"ops@company.com", role:"ops" };
const SEED_DELEGATES = [
  { id:"DEL001", supervisor_id:"SUP001", name:"محمد خالد",   phone:"0501111111", status:"مقبول",        commission_rate:5, orders:120, vehicle_type:"موتوسيكل" },
  { id:"DEL002", supervisor_id:"SUP001", name:"فاطمة حسن",   phone:"0502222222", status:"قيد المراجعة", commission_rate:4, orders:0,   vehicle_type:"دراجة هوائية" },
  { id:"DEL003", supervisor_id:"SUP002", name:"عمر سالم",    phone:"0503333333", status:"مقبول",        commission_rate:6, orders:85,  vehicle_type:"موتوسيكل" },
  { id:"DEL004", supervisor_id:"SUP002", name:"نور إبراهيم", phone:"0504444444", status:"مرفوض",        commission_rate:0, orders:0,   vehicle_type:"دراجة هوائية" },
];

const genId = (p) => p + Math.random().toString(36).substr(2,6).toUpperCase();
const fmtTime = (ts) => {
  const diff = (Date.now()-ts)/1000;
  if(diff<60)    return "الآن";
  if(diff<3600)  return `منذ ${Math.floor(diff/60)} د`;
  if(diff<86400) return `منذ ${Math.floor(diff/3600)} س`;
  return new Date(ts).toLocaleDateString("ar-EG");
};
const fmtFull = (ts) => new Date(ts).toLocaleString("ar-EG",{hour:"2-digit",minute:"2-digit",day:"numeric",month:"short"});

// ─── Design Tokens ────────────────────────────────────────────────────────
const C = {
  bg:"#080f1d", panel:"#0d1526", card:"#141b2d", border:"#1e2d45",
  blue:"#2563eb", green:"#22c55e", red:"#ef4444", yellow:"#eab308",
  purple:"#a855f7", muted:"#8899bb", text:"#e2e8f0", dark:"#0a1220",
};
const STATUS_CFG = {
  "مقبول":        {bg:"#0d3d2b",text:"#22c55e",border:"#16a34a"},
  "مرفوض":        {bg:"#3d0d0d",text:"#ef4444",border:"#dc2626"},
  "قيد المراجعة": {bg:"#2d2a0d",text:"#eab308",border:"#ca8a04"},
};

// ─── UI Primitives ────────────────────────────────────────────────────────
function Badge({status}){
  const s=STATUS_CFG[status]||STATUS_CFG["قيد المراجعة"];
  return <span style={{background:s.bg,color:s.text,border:`1px solid ${s.border}`,padding:"3px 11px",borderRadius:20,fontSize:12,fontWeight:700}}>{status}</span>;
}
function Card({children,style={}}){
  return <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:14,padding:"20px 24px",...style}}>{children}</div>;
}
function Inp({label,...p}){
  return(
    <div style={{marginBottom:14}}>
      {label&&<label style={{display:"block",color:C.muted,fontSize:12,marginBottom:5}}>{label}</label>}
      <input {...p} style={{width:"100%",background:C.panel,border:`1px solid ${C.border}`,color:C.text,borderRadius:8,padding:"9px 13px",fontSize:14,outline:"none",boxSizing:"border-box",...p.style}}/>
    </div>
  );
}
function Sel({label,options,...p}){
  return(
    <div style={{marginBottom:14}}>
      {label&&<label style={{display:"block",color:C.muted,fontSize:12,marginBottom:5}}>{label}</label>}
      <select {...p} style={{width:"100%",background:C.panel,border:`1px solid ${C.border}`,color:C.text,borderRadius:8,padding:"9px 13px",fontSize:14,outline:"none",boxSizing:"border-box"}}>
        {options.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    </div>
  );
}
function Btn({children,variant="primary",onClick,style={},disabled}){
  const vMap={
    primary:{bg:C.blue,h:"#1d4ed8",t:"#fff"},
    success:{bg:"#16a34a",h:"#15803d",t:"#fff"},
    danger: {bg:"#dc2626",h:"#b91c1c",t:"#fff"},
    ghost:  {bg:"transparent",h:C.border,t:C.muted,b:`1px solid ${C.border}`},
  };
  const v=vMap[variant];
  return(
    <button onClick={onClick} disabled={disabled}
      style={{background:v.bg,color:v.t,border:v.b||"none",borderRadius:8,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:disabled?"not-allowed":"pointer",opacity:disabled?.5:1,transition:"background .2s",...style}}
      onMouseEnter={e=>{if(!disabled)e.currentTarget.style.background=v.h;}}
      onMouseLeave={e=>{if(!disabled)e.currentTarget.style.background=v.bg;}}>
      {children}
    </button>
  );
}
function StatBox({label,value,accent=C.blue,sub}){
  return(
    <div style={{background:C.card,border:`1px solid ${accent}33`,borderRadius:14,padding:"18px 22px",flex:1,minWidth:130}}>
      <div style={{color:C.muted,fontSize:12,marginBottom:6}}>{label}</div>
      <div style={{color:accent,fontSize:28,fontWeight:800}}>{value}</div>
      {sub&&<div style={{color:"#445",fontSize:11,marginTop:4}}>{sub}</div>}
    </div>
  );
}
function PhotoBox({label,hint,required,value,onChange}){
  const camRef = useRef();
  const galRef = useRef();
  const [busy,setBusy]=useState(false);
  const [err,setErr]=useState(null);

  const compressAndSet = (file) => {
    setErr(null);
    if (!file.type.startsWith("image/")) {
      setErr("الملف المختار ليس صورة");
      return;
    }
    setBusy(true);
    const reader = new FileReader();
    reader.onerror = () => { setErr("فشل قراءة الملف"); setBusy(false); };
    reader.onload = (ev) => {
      const img = new Image();
      img.onerror = () => { setErr("فشل تحميل الصورة — جرب صورة أخرى"); setBusy(false); };
      img.onload = () => {
        try {
          const maxDim = 1000;
          let { width, height } = img;
          if (width > maxDim || height > maxDim) {
            if (width > height) { height = Math.round(height * maxDim / width); width = maxDim; }
            else { width = Math.round(width * maxDim / height); height = maxDim; }
          }
          const canvas = document.createElement("canvas");
          canvas.width = width; canvas.height = height;
          const ctx = canvas.getContext("2d");
          ctx.drawImage(img, 0, 0, width, height);
          const dataUrl = canvas.toDataURL("image/jpeg", 0.7);
          if (!dataUrl || dataUrl === "data:,") {
            setErr("فشلت معالجة الصورة، جرب صورة أخرى");
            setBusy(false);
            return;
          }
          onChange(dataUrl);
        } catch (e) {
          setErr("خطأ في معالجة الصورة: " + e.message);
        } finally {
          setBusy(false);
        }
      };
      img.src = ev.target.result;
    };
    reader.readAsDataURL(file);
  };

  const handleFile = (e) => {
    const f = e.target.files[0];
    if (f) compressAndSet(f);
    e.target.value = "";
  };

  return(
    <div style={{flex:1,minWidth:140}}>
      <div style={{color:C.muted,fontSize:12,marginBottom:6,fontWeight:600}}>
        {label} {required&&<span style={{color:C.red}}>*</span>}
      </div>

      <div style={{border:`2px dashed ${value?C.green:err?C.red:C.border}`,borderRadius:10,background:C.panel,
        aspectRatio:"1",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
        overflow:"hidden",position:"relative",minHeight:130}}>
        {busy
          ? <div style={{color:C.muted,fontSize:12}}>⏳ جاري المعالجة...</div>
          : value
            ? <>
                <img src={value} alt={label} style={{width:"100%",height:"100%",objectFit:"cover"}}/>
                <div onClick={()=>{onChange(null);setErr(null);}}
                  style={{position:"absolute",top:6,left:6,background:"#dc262699",borderRadius:"50%",
                    width:24,height:24,display:"flex",alignItems:"center",justifyContent:"center",
                    cursor:"pointer",fontSize:13,color:"#fff",fontWeight:700}}>✕</div>
              </>
            : <><div style={{fontSize:26,marginBottom:6}}>📷</div>
                <div style={{color:C.muted,fontSize:11,textAlign:"center",padding:"0 8px"}}>{hint}</div></>
        }
      </div>

      {err&&<div style={{color:C.red,fontSize:11,marginTop:4,textAlign:"center"}}>{err}</div>}

      {!value&&!busy&&(
        <div style={{display:"flex",gap:6,marginTop:8}}>
          <button onClick={()=>camRef.current?.click()}
            style={{flex:1,background:`${C.blue}22`,color:C.blue,border:`1px solid ${C.blue}44`,
              borderRadius:8,padding:"8px 6px",cursor:"pointer",fontSize:12,fontWeight:700}}>
            📸 الكاميرا
          </button>
          <button onClick={()=>galRef.current?.click()}
            style={{flex:1,background:C.border,color:C.muted,border:`1px solid ${C.border}`,
              borderRadius:8,padding:"8px 6px",cursor:"pointer",fontSize:12,fontWeight:700}}>
            🖼️ المعرض
          </button>
        </div>
      )}

      <input ref={camRef} type="file" accept="image/*" capture="environment"
        style={{display:"none"}} onChange={handleFile}/>
      <input ref={galRef} type="file" accept="image/*"
        style={{display:"none"}} onChange={handleFile}/>
    </div>
  );
}

// ─── Loading Spinner ──────────────────────────────────────────────────────
function Spinner({text="جاري التحميل..."}){
  return(
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"60px 20px",gap:16}}>
      <div style={{width:44,height:44,border:`4px solid ${C.border}`,borderTop:`4px solid ${C.blue}`,borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      <div style={{color:C.muted,fontSize:14}}>{text}</div>
    </div>
  );
}

// ─── Notification Bell ────────────────────────────────────────────────────
function NotifBell({notifs,onRead,onClear}){
  const [open,setOpen]=useState(false);
  const unread=notifs.filter(n=>!n.is_read).length;
  const ref=useRef();
  useEffect(()=>{
    const h=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};
    document.addEventListener("mousedown",h);return()=>document.removeEventListener("mousedown",h);
  },[]);
  const icons={success:"✅",error:"❌",info:"📦",message:"💬"};
  return(
    <div ref={ref} style={{position:"relative"}}>
      <div onClick={()=>{setOpen(o=>!o);if(!open)onRead();}} style={{cursor:"pointer",position:"relative",userSelect:"none",fontSize:22}}>
        🔔
        {unread>0&&<span style={{position:"absolute",top:-6,left:-6,background:C.red,color:"#fff",borderRadius:"50%",width:18,height:18,fontSize:11,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center"}}>{unread}</span>}
      </div>
      {open&&(
        <div style={{position:"absolute",top:38,left:"50%",transform:"translateX(-50%)",background:C.card,border:`1px solid ${C.border}`,borderRadius:12,width:310,boxShadow:"0 8px 32px #000a",zIndex:999,overflow:"hidden"}}>
          <div style={{padding:"12px 16px",borderBottom:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontWeight:700,color:C.text,fontSize:14}}>الإشعارات</span>
            <span onClick={onClear} style={{color:C.muted,fontSize:12,cursor:"pointer"}}>مسح الكل</span>
          </div>
          <div style={{maxHeight:300,overflowY:"auto"}}>
            {notifs.length===0
              ?<div style={{padding:"24px",textAlign:"center",color:C.muted,fontSize:13}}>لا توجد إشعارات</div>
              :notifs.map(n=>(
                <div key={n.id} style={{padding:"11px 14px",borderBottom:`1px solid ${C.border}22`,background:n.is_read?"transparent":"#1e2d4518",display:"flex",gap:10}}>
                  <span style={{fontSize:17,flexShrink:0}}>{icons[n.type]||"📌"}</span>
                  <div style={{flex:1}}>
                    <div style={{color:C.text,fontSize:13}}>{n.message}</div>
                    <div style={{color:C.muted,fontSize:11,marginTop:2}}>{fmtTime(n.created_at)}</div>
                  </div>
                  {!n.is_read&&<div style={{width:7,height:7,borderRadius:"50%",background:C.blue,flexShrink:0,marginTop:5}}/>}
                </div>
              ))
            }
          </div>
        </div>
      )}
    </div>
  );
}

// ══════════════════ LOGIN ══════════════════════════════════════════════════
function LoginScreen({onLogin}){
  const [phone,setPhone]=useState("");
  const [pass,setPass]=useState("");
  const [err,setErr]=useState("");
  const [show,setShow]=useState(false);
  const [loading,setLoading]=useState(false);

  const attempt=async()=>{
    if(!phone.trim()||!pass.trim()){setErr("أدخل رقم الهاتف وكلمة المرور");return;}
    setLoading(true);setErr("");
    try{
      // Check ops manager first (local check, no DB needed)
      if(phone.trim()===OPS_MANAGER.phone&&pass.trim()===OPS_MANAGER.password_hash){
        onLogin(OPS_MANAGER);return;
      }
      // Check seed supervisors locally first (in case DB not seeded yet)
      const localSup=SEED_SUPERVISORS.find(s=>s.phone===phone.trim()&&s.password_hash===pass.trim());
      if(localSup){ onLogin(localSup); return; }
      // Then check DB
      const row=await dbSelectOne("supervisors",{phone:phone.trim(),password_hash:pass.trim()});
      if(!row){setErr("رقم الهاتف أو كلمة المرور غير صحيحة");}
      else onLogin(row);
    }catch(e){setErr("حدث خطأ، حاول مجدداً");}
    finally{setLoading(false);}
  };

  return(
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center"}} dir="rtl">
      <div style={{width:"100%",maxWidth:420,padding:"0 16px"}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{width:64,height:64,background:C.blue,borderRadius:18,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:30,fontWeight:900,color:"#fff",marginBottom:14}}>M</div>
          <div style={{fontSize:24,fontWeight:800,color:C.text}}>ManageDesk</div>
          <div style={{color:C.muted,fontSize:13,marginTop:4}}>نظام إدارة المناديب</div>
        </div>
        <Card>
          <div style={{fontSize:17,fontWeight:700,color:C.text,marginBottom:20,textAlign:"center"}}>تسجيل الدخول</div>
          <Inp label="رقم الهاتف" type="tel" placeholder="05XXXXXXXX" value={phone} onChange={e=>setPhone(e.target.value)} onKeyDown={e=>e.key==="Enter"&&attempt()}/>
          <div style={{marginBottom:14}}>
            <label style={{display:"block",color:C.muted,fontSize:12,marginBottom:5}}>كلمة المرور</label>
            <div style={{position:"relative"}}>
              <input type={show?"text":"password"} placeholder="••••••" value={pass} onChange={e=>setPass(e.target.value)} onKeyDown={e=>e.key==="Enter"&&attempt()}
                style={{width:"100%",background:C.panel,border:`1px solid ${C.border}`,color:C.text,borderRadius:8,padding:"9px 40px 9px 13px",fontSize:14,outline:"none",boxSizing:"border-box"}}/>
              <span onClick={()=>setShow(s=>!s)} style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",cursor:"pointer",fontSize:16,color:C.muted,userSelect:"none"}}>{show?"🙈":"👁️"}</span>
            </div>
          </div>
          {err&&<div style={{background:"#3d0d0d",border:`1px solid ${C.red}`,borderRadius:8,padding:"10px 14px",marginBottom:14,color:C.red,fontSize:13}}>{err}</div>}
          <Btn onClick={attempt} disabled={loading} style={{width:"100%",padding:"12px",fontSize:15}}>
            {loading?"جاري التحقق...":"دخول →"}
          </Btn>
          <div style={{marginTop:16,background:C.panel,borderRadius:8,padding:"12px 14px"}}>
            <div style={{color:C.muted,fontSize:11,marginBottom:8,fontWeight:600}}>حسابات تجريبية:</div>
            {[...SEED_SUPERVISORS,OPS_MANAGER].map(s=>(
              <div key={s.id} onClick={()=>{setPhone(s.phone);setPass(s.password_hash);}} style={{fontSize:12,color:"#556",marginBottom:4,cursor:"pointer",padding:"4px 8px",borderRadius:6,transition:"background .15s"}}
                onMouseEnter={e=>e.currentTarget.style.background=C.border}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                {s.role==="ops"?"⚙️":"👤"} <strong style={{color:C.muted}}>{s.name}</strong> — 📱{s.phone} | 🔑{s.password_hash}
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

// ══════════════════ MAIN APP ══════════════════════════════════════════════
const SUP_TABS=[
  {id:"dashboard", label:"📊 لوحة التحكم"},
  {id:"upload_doc",label:"📄 إضافة مندوب"},
  {id:"delegates", label:"👥 المناديب"},
  {id:"orders",    label:"📦 الأوردرات"},
  {id:"messages",  label:"💬 الرسائل"},
];
const OPS_TABS=[
  {id:"ops_dashboard",label:"📊 لوحة المدير"},
  {id:"messages",     label:"💬 الرسائل"},
];

export default function App(){
  const [currentUser,setCurrentUser]=useState(null);
  const [tab,setTab]=useState("dashboard");
  const [toast,setToast]=useState(null);
  const [supervisors,setSupervisors]=useState([]);
  const [delegates,setDelegatesState]=useState([]);
  const [notifs,setNotifs]=useState([]);
  const [conversations,setConversations]=useState([]);
  const [loading,setLoading]=useState(false);

  const notify=useCallback((msg,type="success")=>{
    setToast({msg,type});setTimeout(()=>setToast(null),3500);
  },[]);

  // ── Seed DB on first load ──
  const seedIfEmpty=async()=>{
    const{data:sups}=await (async()=>{ const d=await dbSelect("supervisors"); return {data:d}; })();
    if(!sups||sups.length===0){
      await dbInsert("supervisors", SEED_SUPERVISORS);
      await dbInsert("delegates", SEED_DELEGATES);
      const now=Date.now();
      await dbInsert("conversations", [{id:"CONV001",participants:["SUP001","OPS001"]},{id:"CONV002",participants:["SUP002","OPS001"]}]);
      await dbInsert("messages", [{id:"MSG1",conv_id:"CONV001",sender_id:"OPS001",content:"مرحباً أحمد، كيف حال فريق المناديب لديك؟",created_at:now-7200000,read_by:["OPS001"]},{id:"MSG2",conv_id:"CONV001",sender_id:"SUP001",content:"الحمد لله، لدينا مندوبان مقبولان ويعملان بشكل ممتاز 💪",created_at:now-3600000,read_by:["SUP001","OPS001"]},{id:"MSG3",conv_id:"CONV002",sender_id:"OPS001",content:"سارة، هل تحتاجين مساعدة في مراجعة المناديب الجدد؟",created_at:now-86400000,read_by:["OPS001"]}]);
    }
  };

  // ── Load all data ──
  const loadAll=async(user)=>{
    setLoading(true);
    try{
      await seedIfEmpty();
      const isOps=user.role==="ops";
      const[sups,dels,nots,convs,msgs]=await Promise.all([
        dbSelect("supervisors"),
        isOps?dbSelect("delegates"):dbSelect("delegates",{supervisor_id:user.id}),
        isOps?Promise.resolve([]):dbSelect("notifications",{sup_id:user.id}),
        dbSelect("conversations",{},{"participants":[user.id]}),
        dbSelect("messages"),
      ]);
      setSupervisors(sups||[]);
      setDelegatesState(dels||[]);
      setNotifs(nots||[]);
      // Build conversations with messages
      const convsWithMsgs=(convs||[]).map(c=>({
        ...c,
        messages:(msgs||[]).filter(m=>m.conv_id===c.id).sort((a,b)=>a.created_at-b.created_at)
      }));
      setConversations(convsWithMsgs);
    }catch(e){console.error(e);}
    finally{setLoading(false);}
  };

  const handleLogin=async(user)=>{
    setCurrentUser(user);
    setTab(user.role==="ops"?"ops_dashboard":"dashboard");
    await loadAll(user);
  };

  // ── Real-time subscriptions ──
  useEffect(()=>{
    if(!currentUser) return;
    const msgSub=sb.channel("messages-channel")
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"messages"},payload=>{
        const newMsg=payload.new;
        setConversations(prev=>prev.map(c=>
          c.id===newMsg.conv_id?{...c,messages:[...c.messages,newMsg]}:c
        ));
      }).subscribe();
    const notifSub=sb.channel("notifs-channel")
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"notifications"},payload=>{
        if(payload.new.sup_id===currentUser.id){
          setNotifs(prev=>[payload.new,...prev]);
        }
      }).subscribe();
    return()=>{sb.removeChannel(msgSub);sb.removeChannel(notifSub);};
  },[currentUser]);

  if(!currentUser) return <LoginScreen onLogin={handleLogin}/>;
  if(loading) return(
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center"}} dir="rtl">
      <Spinner text="جاري تحميل البيانات..."/>
    </div>
  );

  const isOps=currentUser.role==="ops";
  const myNotifs=notifs;
  const myConvs=conversations;
  const unreadMsgs=myConvs.reduce((s,c)=>s+c.messages.filter(m=>m.sender_id!==currentUser.id&&!(m.read_by||[]).includes(currentUser.id)).length,0);
  const myDelegates=delegates;
  const accepted=myDelegates.filter(d=>d.status==="مقبول");
  const totalOrders=accepted.reduce((s,d)=>s+(d.orders||0),0);

  // ── DB Actions ──
  const addNotifDB=async(supId,message,type="info")=>{
    const n={id:genId("N"),sup_id:supId,message,type,created_at:Date.now(),is_read:false};
    await dbInsert("notifications", n);
    if(supId===currentUser.id) setNotifs(prev=>[n,...prev]);
  };

  const markNotifsRead=async()=>{
    await dbUpdate("notifications",{is_read:true},{sup_id:currentUser.id});
    setNotifs(prev=>prev.map(n=>({...n,is_read:true})));
  };

  const clearNotifs=async()=>{
    await dbDelete("notifications",{sup_id:currentUser.id});
    setNotifs([]);
  };

  const changeStatus=async(delId,newStatus)=>{
    const d=delegates.find(x=>x.id===delId);if(!d)return;
    await dbUpdate("delegates",{status:newStatus},{id:delId});
    setDelegatesState(prev=>prev.map(x=>x.id===delId?{...x,status:newStatus}:x));
    const msg=newStatus==="مقبول"?`✅ تم قبول المندوب "${d.name}" (${d.id})`:`❌ تم رفض المندوب "${d.name}" (${d.id})`;
    await addNotifDB(d.supervisor_id,msg,newStatus==="مقبول"?"success":"error");
    notify(msg,newStatus==="مقبول"?"success":"error");
  };

  const setDelegates=async(updater)=>{
    const updated=typeof updater==="function"?updater(delegates):updater;
    setDelegatesState(updated);
  };

  const TABS=isOps?OPS_TABS:SUP_TABS;

  return(
    <div dir="rtl" style={{minHeight:"100vh",background:C.bg,color:C.text,fontFamily:"'Segoe UI',Tahoma,Arial,sans-serif",fontSize:14}}>
      {toast&&<div style={{position:"fixed",top:20,right:20,zIndex:9999,background:toast.type==="success"?"#16a34a":"#dc2626",color:"#fff",padding:"12px 22px",borderRadius:10,fontWeight:700,boxShadow:"0 4px 24px #0008",fontSize:14}}>{toast.msg}</div>}

      {/* Header */}
      <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,padding:"0 28px",display:"flex",alignItems:"center",justifyContent:"space-between",height:64}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:38,height:38,background:C.blue,borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,fontWeight:900,color:"#fff"}}>M</div>
          <div>
            <div style={{fontWeight:800,fontSize:16,color:C.text}}>ManageDesk</div>
            <div style={{fontSize:11,color:"#334"}}>نظام إدارة المناديب</div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:18}}>
          <div onClick={()=>setTab("messages")} style={{cursor:"pointer",position:"relative",fontSize:22,userSelect:"none"}}>
            💬
            {unreadMsgs>0&&<span style={{position:"absolute",top:-6,left:-6,background:C.green,color:"#fff",borderRadius:"50%",width:18,height:18,fontSize:11,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center"}}>{unreadMsgs}</span>}
          </div>
          {!isOps&&<NotifBell notifs={myNotifs} onRead={markNotifsRead} onClear={clearNotifs}/>}
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:34,height:34,background:`${C.blue}33`,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>{isOps?"⚙️":"👤"}</div>
            <div>
              <div style={{fontWeight:700,fontSize:13,color:C.text}}>{currentUser.name}</div>
              <div style={{fontSize:11,color:isOps?C.purple:C.muted}}>{isOps?"مدير التشغيل":"مشرف | "+currentUser.id}</div>
            </div>
          </div>
          <Btn variant="ghost" onClick={()=>{setCurrentUser(null);setDelegatesState([]);setNotifs([]);setConversations([]);}} style={{padding:"6px 12px",fontSize:12}}>خروج ↩</Btn>
        </div>
      </div>

      {/* Tabs */}
      <div style={{background:C.panel,borderBottom:`1px solid ${C.border}`,display:"flex",padding:"0 20px",gap:2,overflowX:"auto"}}>
        {TABS.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{background:tab===t.id?C.card:"transparent",color:tab===t.id?C.blue:C.muted,border:"none",borderBottom:tab===t.id?`2px solid ${C.blue}`:"2px solid transparent",padding:"14px 16px",cursor:"pointer",fontSize:13,fontWeight:600,whiteSpace:"nowrap",transition:"all .2s",position:"relative"}}>
            {t.label}
            {t.id==="messages"&&unreadMsgs>0&&<span style={{position:"absolute",top:8,right:4,background:C.green,color:"#fff",borderRadius:"50%",width:14,height:14,fontSize:9,fontWeight:800,display:"inline-flex",alignItems:"center",justifyContent:"center"}}>{unreadMsgs}</span>}
          </button>
        ))}
      </div>

      {/* Content */}
      <div style={{padding:tab==="messages"?"16px 28px":"26px 28px",maxWidth:tab==="messages"?1400:1200,margin:"0 auto"}}>
        {tab==="dashboard"     && <DashboardTab myDelegates={myDelegates} accepted={accepted} totalOrders={totalOrders}/>}
        {tab==="upload_doc"    && <UploadDocTab supervisors={supervisors} setDelegates={setDelegates} notify={notify} currentSup={currentUser} addNotifDB={addNotifDB}/>}
        {tab==="delegates"     && <DelegatesTab myDelegates={myDelegates}/>}
        {tab==="orders"        && <OrdersTab delegates={delegates} setDelegates={setDelegates} currentSup={currentUser} notify={notify} myDelegates={myDelegates} addNotifDB={addNotifDB}/>}
        {tab==="messages"      && <MessagingTab currentUser={currentUser} conversations={conversations} setConversations={setConversations} supervisors={supervisors} addNotifDB={addNotifDB}/>}
        {tab==="supervisor"    && <SupervisorTab supervisors={supervisors} setSupervisors={setSupervisors} notify={notify}/>}
        {tab==="ops_dashboard" && <OpsDashboard delegates={delegates} setDelegates={setDelegates} supervisors={supervisors} setSupervisors={setSupervisors} changeStatus={changeStatus} notify={notify} addNotifDB={addNotifDB}/>}
      </div>
    </div>
  );
}

// ══════════════════ MESSAGING ══════════════════════════════════════════════
function MessagingTab({currentUser,conversations,setConversations,supervisors,addNotifDB}){
  const isOps=currentUser.role==="ops";
  const [activeConvId,setActiveConvId]=useState(conversations[0]?.id||null);
  const [draft,setDraft]=useState("");
  const [newSupId,setNewSupId]=useState("");
  const [showNew,setShowNew]=useState(false);
  const [sending,setSending]=useState(false);
  const [autoCreating,setAutoCreating]=useState(false);
  const bottomRef=useRef();
  const inputRef=useRef();

  const activeConv=conversations.find(c=>c.id===activeConvId);
  const getOther=(conv)=>{
    const oid=conv.participants.find(p=>p!==currentUser.id);
    if(oid===OPS_MANAGER.id) return OPS_MANAGER;
    return supervisors.find(s=>s.id===oid)||{id:oid,name:oid};
  };
  const unreadCount=(conv)=>conv.messages.filter(m=>m.sender_id!==currentUser.id&&!(m.read_by||[]).includes(currentUser.id)).length;
  const totalUnread=conversations.reduce((s,c)=>s+unreadCount(c),0);

  // Supervisor only ever talks to the ops manager — auto-create that single
  // conversation if it doesn't exist yet, so they never need to search for an ID.
  useEffect(()=>{
    if(isOps) return;
    if(conversations.length>0){
      if(!activeConvId) setActiveConvId(conversations[0].id);
      return;
    }
    if(autoCreating) return;
    setAutoCreating(true);
    (async()=>{
      const nc={id:genId("CONV"),participants:[currentUser.id,OPS_MANAGER.id]};
      await dbInsert("conversations",nc);
      setConversations(prev=>[...prev,{...nc,messages:[]}]);
      setActiveConvId(nc.id);
      setAutoCreating(false);
    })();
  },[isOps,conversations.length]);

  useEffect(()=>{
    if(!activeConvId)return;
    // Mark as read in DB
    const conv=conversations.find(c=>c.id===activeConvId);
    if(!conv)return;
    const unreadMsgs=conv.messages.filter(m=>m.sender_id!==currentUser.id&&!(m.read_by||[]).includes(currentUser.id));
    if(unreadMsgs.length>0){
      unreadMsgs.forEach(async m=>{
        const newReadBy=[...(m.read_by||[]),currentUser.id];
        await dbUpdate("messages",{read_by:newReadBy},{id:m.id});
      });
      setConversations(prev=>prev.map(c=>{
        if(c.id!==activeConvId)return c;
        return{...c,messages:c.messages.map(m=>m.sender_id!==currentUser.id?{...m,read_by:[...(m.read_by||[]),currentUser.id]}:m)};
      }));
    }
  },[activeConvId]);

  useEffect(()=>{ bottomRef.current?.scrollIntoView({behavior:"smooth"}); },[activeConv?.messages?.length]);

  const sendMessage=async()=>{
    if(!draft.trim()||!activeConvId||sending)return;
    setSending(true);
    const msg={id:genId("MSG"),conv_id:activeConvId,sender_id:currentUser.id,content:draft.trim(),created_at:Date.now(),read_by:[currentUser.id]};
    try{
      await dbInsert("messages",msg);
      setConversations(prev=>prev.map(c=>c.id===activeConvId?{...c,messages:[...c.messages,msg]}:c));
      const conv=conversations.find(c=>c.id===activeConvId);
      const otherId=conv?.participants.find(p=>p!==currentUser.id);
      if(otherId&&otherId!==OPS_MANAGER.id){
        await addNotifDB(otherId,`💬 رسالة جديدة من ${currentUser.name}: "${draft.trim().slice(0,40)}"...`,"message");
      }
      setDraft("");inputRef.current?.focus();
    }finally{setSending(false);}
  };

  const startNewConv=async()=>{
    if(!newSupId)return;
    const exists=conversations.find(c=>c.participants.includes(currentUser.id)&&c.participants.includes(newSupId));
    if(exists){setActiveConvId(exists.id);setShowNew(false);return;}
    const nc={id:genId("CONV"),participants:[currentUser.id,newSupId]};
    await dbInsert("conversations",nc);
    setConversations(prev=>[...prev,{...nc,messages:[]}]);
    setActiveConvId(nc.id);setShowNew(false);setNewSupId("");
  };

  return(
    <div style={{display:"flex",height:"calc(100vh - 130px)",minHeight:500,border:`1px solid ${C.border}`,borderRadius:14,overflow:"hidden"}}>
      {/* Sidebar — only shown to ops manager who may have multiple conversations */}
      {isOps&&(
      <div style={{width:280,flexShrink:0,background:C.panel,borderLeft:`1px solid ${C.border}`,display:"flex",flexDirection:"column"}}>
        <div style={{padding:"16px 18px",borderBottom:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <div style={{fontWeight:700,color:C.text,fontSize:15}}>💬 الرسائل</div>
            {totalUnread>0&&<div style={{fontSize:11,color:C.blue,marginTop:2}}>{totalUnread} غير مقروءة</div>}
          </div>
          {isOps&&<button onClick={()=>setShowNew(s=>!s)} style={{background:`${C.blue}22`,color:C.blue,border:`1px solid ${C.blue}44`,borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:13,fontWeight:700}}>+ جديد</button>}
        </div>
        {showNew&&isOps&&(
          <div style={{padding:"12px 16px",background:C.dark,borderBottom:`1px solid ${C.border}`}}>
            <Sel label="محادثة مع" value={newSupId} onChange={e=>setNewSupId(e.target.value)}
              options={[{value:"",label:"-- اختر مشرفاً --"},...supervisors.map(s=>({value:s.id,label:s.name}))]}/>
            <div style={{display:"flex",gap:8}}>
              <Btn onClick={startNewConv} disabled={!newSupId} style={{flex:1,padding:"7px",fontSize:12}}>بدء</Btn>
              <Btn variant="ghost" onClick={()=>setShowNew(false)} style={{padding:"7px 12px",fontSize:12}}>إلغاء</Btn>
            </div>
          </div>
        )}
        <div style={{flex:1,overflowY:"auto"}}>
          {conversations.length===0
            ?<div style={{padding:"30px 18px",textAlign:"center",color:C.muted,fontSize:13}}>لا توجد محادثات</div>
            :conversations.map(conv=>{
              const other=getOther(conv);
              const last=conv.messages[conv.messages.length-1];
              const unread=unreadCount(conv);
              const isActive=activeConvId===conv.id;
              return(
                <div key={conv.id} onClick={()=>setActiveConvId(conv.id)}
                  style={{padding:"14px 18px",cursor:"pointer",borderBottom:`1px solid ${C.border}22`,background:isActive?`${C.blue}18`:"transparent",borderRight:isActive?`3px solid ${C.blue}`:"3px solid transparent",transition:"all .15s"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                    <div style={{display:"flex",alignItems:"center",gap:10}}>
                      <div style={{width:38,height:38,background:isActive?`${C.blue}33`:C.border,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>
                        {other.role==="ops"?"⚙️":"👤"}
                      </div>
                      <div>
                        <div style={{color:isActive?C.text:"#aab",fontWeight:unread>0?700:500,fontSize:14}}>{other.name}</div>
                        <div style={{color:C.muted,fontSize:11,marginTop:2,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                          {last?(last.sender_id===currentUser.id?"أنت: ":"")+last.content:"لا توجد رسائل"}
                        </div>
                      </div>
                    </div>
                    <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4}}>
                      {last&&<div style={{color:C.muted,fontSize:10}}>{fmtTime(last.created_at)}</div>}
                      {unread>0&&<div style={{background:C.blue,color:"#fff",borderRadius:"50%",width:18,height:18,fontSize:11,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center"}}>{unread}</div>}
                    </div>
                  </div>
                </div>
              );
            })
          }
        </div>
        <div style={{padding:"12px 18px",borderTop:`1px solid ${C.border}`,background:C.dark,display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,background:`${C.blue}33`,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>{currentUser.role==="ops"?"⚙️":"👤"}</div>
          <div style={{flex:1}}>
            <div style={{color:C.text,fontSize:12,fontWeight:600}}>{currentUser.name}</div>
            <div style={{color:C.muted,fontSize:10}}>{currentUser.role==="ops"?"مدير التشغيل":"مشرف"}</div>
          </div>
          <div style={{width:8,height:8,borderRadius:"50%",background:C.green}}/>
        </div>
      </div>
      )}

      {/* Chat area */}
      {activeConv?(()=>{
        const other=getOther(activeConv);
        return(
          <div style={{flex:1,display:"flex",flexDirection:"column",background:C.bg}}>
            <div style={{padding:"14px 22px",borderBottom:`1px solid ${C.border}`,background:C.panel,display:"flex",alignItems:"center",gap:14}}>
              <div style={{width:42,height:42,background:`${C.blue}33`,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>{other.role==="ops"?"⚙️":"👤"}</div>
              <div>
                <div style={{fontWeight:700,color:C.text,fontSize:15}}>{other.name}</div>
                <div style={{color:C.green,fontSize:12,display:"flex",alignItems:"center",gap:5}}>
                  <span style={{width:7,height:7,borderRadius:"50%",background:C.green,display:"inline-block"}}/>متصل
                </div>
              </div>
              <div style={{marginRight:"auto",color:C.muted,fontSize:12}}>{activeConv.messages.length} رسالة</div>
            </div>
            <div style={{flex:1,overflowY:"auto",padding:"20px 24px",display:"flex",flexDirection:"column",gap:12}}>
              {activeConv.messages.length===0&&(
                <div style={{textAlign:"center",padding:"60px 20px",color:C.muted}}>
                  <div style={{fontSize:48,marginBottom:12}}>👋</div>
                  <div>ابدأ المحادثة مع {other.name}</div>
                </div>
              )}
              {activeConv.messages.map((msg,idx)=>{
                const isMine=msg.sender_id===currentUser.id;
                const showDate=idx===0||new Date(activeConv.messages[idx-1].created_at).toDateString()!==new Date(msg.created_at).toDateString();
                return(
                  <div key={msg.id}>
                    {showDate&&(
                      <div style={{textAlign:"center",margin:"8px 0"}}>
                        <span style={{background:C.card,color:C.muted,fontSize:11,padding:"4px 14px",borderRadius:20,border:`1px solid ${C.border}`}}>
                          {new Date(msg.created_at).toLocaleDateString("ar-EG",{weekday:"long",day:"numeric",month:"long"})}
                        </span>
                      </div>
                    )}
                    <div style={{display:"flex",justifyContent:isMine?"flex-start":"flex-end",gap:10}}>
                      {!isMine&&<div style={{width:34,height:34,background:`${C.blue}33`,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0,alignSelf:"flex-end"}}>{other.role==="ops"?"⚙️":"👤"}</div>}
                      <div style={{maxWidth:"68%"}}>
                        <div style={{background:isMine?`${C.blue}22`:C.card,border:`1px solid ${isMine?`${C.blue}44`:C.border}`,borderRadius:isMine?"14px 14px 14px 4px":"14px 14px 4px 14px",padding:"10px 14px",color:C.text,fontSize:14,lineHeight:1.6,wordBreak:"break-word"}}>
                          {msg.content}
                        </div>
                        <div style={{display:"flex",justifyContent:isMine?"flex-start":"flex-end",alignItems:"center",gap:6,marginTop:3}}>
                          <span style={{color:C.muted,fontSize:10}}>{fmtFull(msg.created_at)}</span>
                          {isMine&&<span style={{fontSize:11,color:(msg.read_by||[]).length>1?C.blue:C.muted}}>{(msg.read_by||[]).length>1?"✓✓":"✓"}</span>}
                        </div>
                      </div>
                      {isMine&&<div style={{width:34,height:34,background:`${C.blue}33`,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0,alignSelf:"flex-end"}}>{currentUser.role==="ops"?"⚙️":"👤"}</div>}
                    </div>
                  </div>
                );
              })}
              <div ref={bottomRef}/>
            </div>
            <div style={{padding:"14px 20px",borderTop:`1px solid ${C.border}`,background:C.panel,display:"flex",gap:12,alignItems:"flex-end"}}>
              <textarea ref={inputRef} rows={1} value={draft} onChange={e=>setDraft(e.target.value)}
                onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendMessage();}}}
                placeholder={`اكتب رسالة لـ ${other.name}...`}
                style={{flex:1,background:C.dark,border:`1px solid ${C.border}`,color:C.text,borderRadius:12,padding:"11px 15px",fontSize:14,outline:"none",resize:"none",fontFamily:"inherit",lineHeight:1.5,maxHeight:120,overflowY:"auto"}}/>
              <button onClick={sendMessage} disabled={!draft.trim()||sending}
                style={{width:44,height:44,background:draft.trim()&&!sending?C.blue:"#1e2d45",color:"#fff",border:"none",borderRadius:12,cursor:draft.trim()&&!sending?"pointer":"not-allowed",fontSize:18,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                {sending?"⏳":"➤"}
              </button>
            </div>
          </div>
        );
      })()
      :(
        <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",background:C.bg}}>
          <div style={{textAlign:"center",color:C.muted}}>
            <div style={{fontSize:56,marginBottom:16}}>💬</div>
            <div style={{fontSize:16,fontWeight:600,color:C.text,marginBottom:8}}>الرسائل الداخلية</div>
            <div style={{fontSize:13}}>{isOps?"اختر محادثة من القائمة":"جاري تجهيز المحادثة مع مدير التشغيل..."}</div>
          </div>
        </div>
      )}
    </div>
  );
}

// ══════════════════ OPS DASHBOARD ════════════════════════════════════════
function OpsDashboard({delegates,setDelegates,supervisors,setSupervisors,changeStatus,notify,addNotifDB}){
  const accepted=delegates.filter(d=>d.status==="مقبول");
  const pending=delegates.filter(d=>d.status==="قيد المراجعة");
  const totalOrders=accepted.reduce((s,d)=>s+(d.orders||0),0);

  const [xlsErr,setXlsErr]=useState(null);
  const [preview,setPreview]=useState(null);
  const [saving,setSaving]=useState(false);
  const fileRef=useRef();

  const [editId,setEditId]=useState(null);
  const [editRate,setEditRate]=useState("");

  const [renameId,setRenameId]=useState(null);
  const [renameVal,setRenameVal]=useState("");
  const [previewImg,setPreviewImg]=useState(null);

  const handleXlsx=(e)=>{
    const file=e.target.files[0];if(!file)return;setXlsErr(null);
    const reader=new FileReader();
    reader.onload=(ev)=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1});
        if(rows.length<2){setXlsErr("الملف فارغ");return;}
        const hdrs=rows[0].map(h=>String(h||"").trim().toLowerCase());
        const idCol=hdrs.findIndex(h=>["id","delegate_id","مندوب","del_id"].some(k=>h.includes(k)));
        const orCol=hdrs.findIndex(h=>["order","أوردر","عدد","orders"].some(k=>h.includes(k)));
        if(idCol===-1||orCol===-1){setXlsErr(`لم يُعثر على الأعمدة.\nأعمدة: ${rows[0].join(", ")}`);return;}
        setPreview(rows.slice(1).filter(r=>r[idCol]).map(r=>({id:String(r[idCol]).trim(),orders:parseInt(r[orCol])||0})));
      }catch(err){setXlsErr("خطأ: "+err.message);}
    };
    reader.readAsBinaryString(file);
  };

  const applyXlsx=async()=>{
    if(!preview)return;setSaving(true);
    let cnt=0;
    try{
      for(const r of preview){
        const d=delegates.find(x=>x.id===r.id&&x.status==="مقبول");
        if(d){await dbUpdate("delegates",{orders:r.orders},{id:r.id});cnt++;}
      }
      setDelegates(prev=>Array.isArray(prev)?prev.map(d=>{
        const m=preview.find(r=>r.id===d.id);
        return m&&d.status==="مقبول"?{...d,orders:m.orders}:d;
      }):[]);
      // notify each affected supervisor
      const affectedSupIds=[...new Set(preview.map(r=>{
        const d=delegates.find(x=>x.id===r.id);
        return d?d.supervisor_id:null;
      }).filter(Boolean))];
      for(const supId of affectedSupIds){
        await addNotifDB(supId,"📦 التحديث الأسبوعي: تم تحديث أوردرات مناديبك","info");
      }
      notify(`✅ تم تحديث ${cnt} مندوب من ${affectedSupIds.length} مشرف`);
      setPreview(null);if(fileRef.current)fileRef.current.value="";
    }finally{setSaving(false);}
  };

  const saveRate=async(id)=>{
    const rate=parseFloat(editRate);
    if(isNaN(rate)||rate<0||rate>100){notify("❗ نسبة غير صحيحة","error");return;}
    await dbUpdate("delegates",{commission_rate:rate},{id});
    setDelegates(prev=>prev.map(d=>d.id===id?{...d,commission_rate:rate}:d));
    const d=delegates.find(x=>x.id===id);
    if(d) await addNotifDB(d.supervisor_id,`💰 تم تحديد نسبة عمولة "${d.name}" بـ ${rate}%`,"info");
    notify("✅ تم تحديث النسبة");
    setEditId(null);
  };

  const saveRename=async(id)=>{
    if(!renameVal.trim()){notify("❗ أدخل اسماً","error");return;}
    await dbUpdate("supervisors",{name:renameVal.trim()},{id});
    setSupervisors(prev=>prev.map(s=>s.id===id?{...s,name:renameVal.trim()}:s));
    notify("✅ تم تعديل الاسم");
    setRenameId(null);
  };

  return(
    <div>
      <h2 style={{color:C.text,margin:"0 0 20px",fontSize:20}}>📊 لوحة مدير التشغيل</h2>
      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:26}}>
        <StatBox label="إجمالي المشرفين"  value={supervisors.length} accent={C.blue}/>
        <StatBox label="إجمالي المناديب"  value={delegates.length}   accent={C.purple}/>
        <StatBox label="المقبولون"         value={accepted.length}    accent={C.green}/>
        <StatBox label="قيد المراجعة"      value={pending.length}     accent={C.yellow}/>
        <StatBox label="إجمالي الأوردرات" value={totalOrders.toLocaleString()} accent="#f97316" sub="لجميع المناديب"/>
      </div>

      {/* Excel upload — ops only */}
      <Card style={{marginBottom:22}}>
        <h3 style={{color:C.text,margin:"0 0 14px",fontSize:15}}>📊 رفع شيت Excel — تحديث أوردرات جميع المناديب</h3>
        <p style={{color:"#445",fontSize:12,marginBottom:12}}>عمود <strong style={{color:C.blue}}>ID</strong> + عمود <strong style={{color:C.blue}}>Orders</strong></p>
        <div onClick={()=>fileRef.current?.click()} style={{border:`2px dashed ${C.border}`,borderRadius:10,padding:"22px",textAlign:"center",cursor:"pointer",background:C.panel}}
          onMouseEnter={e=>e.currentTarget.style.borderColor=C.blue} onMouseLeave={e=>e.currentTarget.style.borderColor=C.border}>
          <div style={{fontSize:36,marginBottom:8}}>📋</div>
          <div style={{color:C.muted,fontSize:13}}>اضغط لاختيار ملف Excel</div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleXlsx}/>
        </div>
        {xlsErr&&<div style={{background:"#3d0d0d",border:`1px solid ${C.red}`,borderRadius:8,padding:"10px 14px",marginTop:10,color:C.red,fontSize:12,whiteSpace:"pre-wrap"}}>{xlsErr}</div>}
        {preview&&(
          <div style={{marginTop:12}}>
            <div style={{color:C.green,fontSize:13,marginBottom:8}}>✅ {preview.length} صف</div>
            <div style={{maxHeight:160,overflowY:"auto",background:C.panel,borderRadius:8,padding:10}}>
              {preview.slice(0,10).map((r,i)=>{
                const f=delegates.find(d=>d.id===r.id);
                return <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:`1px solid ${C.border}22`,fontSize:12}}>
                  <span style={{color:C.blue,fontFamily:"monospace"}}>{r.id}</span>
                  <span style={{color:C.purple}}>{r.orders}</span>
                  <span style={{color:f?C.green:C.red}}>{f?"✓ "+f.name:"✗ غير موجود"}</span>
                </div>;
              })}
            </div>
            <div style={{display:"flex",gap:8,marginTop:10}}>
              <Btn variant="success" onClick={applyXlsx} disabled={saving} style={{flex:1}}>{saving?"⏳ جاري...":"✅ تطبيق"}</Btn>
              <Btn variant="ghost" onClick={()=>setPreview(null)}>إلغاء</Btn>
            </div>
          </div>
        )}
      </Card>

      {/* Pending review across all supervisors */}
      {previewImg&&(
        <div onClick={()=>setPreviewImg(null)} style={{position:"fixed",inset:0,background:"#000c",zIndex:9998,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div onClick={e=>e.stopPropagation()} style={{background:C.card,borderRadius:14,padding:20,maxWidth:420,width:"90%",textAlign:"center"}}>
            <div style={{color:C.text,fontWeight:700,marginBottom:12}}>{previewImg.label}</div>
            <img src={previewImg.src} alt="" style={{width:"100%",borderRadius:10,maxHeight:420,objectFit:"contain"}}/>
            <Btn variant="ghost" onClick={()=>setPreviewImg(null)} style={{marginTop:14,width:"100%"}}>إغلاق</Btn>
          </div>
        </div>
      )}
      {pending.length>0&&(
        <Card style={{marginBottom:22}}>
          <h3 style={{margin:"0 0 16px",color:C.text,fontSize:15}}>⏳ طلبات تحتاج مراجعة ({pending.length})</h3>
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            {pending.map(d=>{
              const sup=supervisors.find(s=>s.id===d.supervisor_id);
              const docs=d.docs||{};
              const docList=[
                {key:"selfie",label:"سيلفي",icon:"🤳"},
                {key:"nationalFront",label:"وش البطاقة",icon:"🪪"},
                {key:"nationalBack",label:"ظهر البطاقة",icon:"🪪"},
                {key:"licenseFront",label:"وش الرخصة",icon:"📋"},
                {key:"licenseBack",label:"ظهر الرخصة",icon:"📋"},
              ].filter(x=>docs[x.key]);
              return(
                <div key={d.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:C.panel,padding:"12px 16px",borderRadius:10,border:`1px solid ${C.border}`,flexWrap:"wrap",gap:10}}>
                  <div>
                    <div style={{fontWeight:600,color:C.text}}>{d.name} <span style={{color:C.muted,fontSize:12}}>({d.id})</span></div>
                    <div style={{color:C.muted,fontSize:12,marginTop:2}}>{d.phone} | {d.vehicle_type==="موتوسيكل"?"🏍️":"🚲"} {d.vehicle_type} | المشرف: {sup?.name||d.supervisor_id}</div>
                    {d.national_id&&<div style={{color:"#556",fontSize:11,marginTop:2}}>رقم قومي: {d.national_id} | العنوان: {d.address}</div>}
                    <div style={{display:"flex",gap:6,marginTop:6}}>
                      {docList.length>0
                        ?docList.map(x=>(
                          <span key={x.key} title={x.label} style={{cursor:"pointer",fontSize:18}}
                            onClick={()=>setPreviewImg({src:docs[x.key],label:x.label})}>{x.icon}</span>
                        ))
                        :<span style={{color:C.red,fontSize:11}}>⚠️ لا توجد مستندات</span>}
                    </div>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <Btn variant="success" onClick={()=>changeStatus(d.id,"مقبول")}>قبول</Btn>
                    <Btn variant="danger"  onClick={()=>changeStatus(d.id,"مرفوض")}>رفض</Btn>
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
      )}

      {/* Commission rates for accepted delegates */}
      <Card style={{marginBottom:22}}>
        <h3 style={{color:C.text,margin:"0 0 16px",fontSize:15}}>💰 تحديد نسب العمولة للمناديب المقبولين</h3>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead>
            <tr style={{background:C.panel,borderBottom:`1px solid ${C.border}`}}>
              {["الاسم","الوسيلة","المشرف","العنوان","الأوردرات","النسبة"].map(h=>(
                <th key={h} style={{padding:"10px 12px",color:C.muted,fontSize:12,fontWeight:700,textAlign:"right"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {accepted.map(d=>{
              const sup=supervisors.find(s=>s.id===d.supervisor_id);
              return(
                <tr key={d.id} style={{borderBottom:`1px solid ${C.border}22`}}>
                  <td style={{padding:"11px 12px",color:C.text}}>{d.name}</td>
                  <td style={{padding:"11px 12px",color:C.muted,fontSize:13}}>{d.vehicle_type==="موتوسيكل"?"🏍️":"🚲"}</td>
                  <td style={{padding:"11px 12px",color:C.muted,fontSize:13}}>{sup?.name||d.supervisor_id}</td>
                  <td style={{padding:"11px 12px",color:"#556",fontSize:12}}>{d.address||"—"}</td>
                  <td style={{padding:"11px 12px",color:C.purple,fontWeight:700}}>{(d.orders||0).toLocaleString()}</td>
                  <td style={{padding:"11px 12px"}}>
                    {editId===d.id
                      ?<div style={{display:"flex",gap:6}}>
                        <input type="number" value={editRate} onChange={e=>setEditRate(e.target.value)}
                          style={{width:60,background:C.panel,border:`1px solid ${C.border}`,color:C.text,borderRadius:6,padding:"4px 7px",fontSize:12}}/>
                        <Btn variant="success" onClick={()=>saveRate(d.id)} style={{padding:"4px 8px",fontSize:11}}>حفظ</Btn>
                        <Btn variant="ghost" onClick={()=>setEditId(null)} style={{padding:"4px 8px",fontSize:11}}>إلغاء</Btn>
                      </div>
                      :<div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{color:d.commission_rate>0?C.green:"#445",fontWeight:700}}>{d.commission_rate>0?`${d.commission_rate}%`:"غير محددة"}</span>
                        <Btn variant="ghost" onClick={()=>{setEditId(d.id);setEditRate(String(d.commission_rate||""));}} style={{padding:"3px 8px",fontSize:11}}>✏️</Btn>
                      </div>}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>

      {/* Supervisors report + rename */}
      <Card>
        <h3 style={{color:C.text,margin:"0 0 16px",fontSize:15}}>📋 تقرير المشرفين</h3>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead>
            <tr style={{background:C.panel,borderBottom:`1px solid ${C.border}`}}>
              {["المشرف","ID","المناديب","المقبولون","قيد المراجعة","إجمالي الأوردرات",""].map(h=>(
                <th key={h} style={{padding:"10px 14px",color:C.muted,fontSize:12,fontWeight:700,textAlign:"right"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {supervisors.map(s=>{
              const md=delegates.filter(d=>d.supervisor_id===s.id);
              const ma=md.filter(d=>d.status==="مقبول");
              const mp=md.filter(d=>d.status==="قيد المراجعة");
              return(
                <tr key={s.id} style={{borderBottom:`1px solid ${C.border}22`}}>
                  <td style={{padding:"12px 14px",color:C.text,fontWeight:600}}>
                    {renameId===s.id
                      ?<input value={renameVal} onChange={e=>setRenameVal(e.target.value)}
                        style={{background:C.panel,border:`1px solid ${C.border}`,color:C.text,borderRadius:6,padding:"4px 8px",fontSize:13,width:120}}/>
                      :s.name}
                  </td>
                  <td style={{padding:"12px 14px",fontFamily:"monospace",color:C.blue,fontSize:12}}>{s.id}</td>
                  <td style={{padding:"12px 14px",color:C.muted}}>{md.length}</td>
                  <td style={{padding:"12px 14px",color:C.green,fontWeight:700}}>{ma.length}</td>
                  <td style={{padding:"12px 14px",color:C.yellow}}>{mp.length}</td>
                  <td style={{padding:"12px 14px",color:C.purple,fontWeight:700}}>{ma.reduce((a,d)=>a+(d.orders||0),0).toLocaleString()}</td>
                  <td style={{padding:"12px 14px"}}>
                    {renameId===s.id
                      ?<div style={{display:"flex",gap:6}}>
                        <Btn variant="success" onClick={()=>saveRename(s.id)} style={{padding:"4px 10px",fontSize:11}}>حفظ</Btn>
                        <Btn variant="ghost" onClick={()=>setRenameId(null)} style={{padding:"4px 10px",fontSize:11}}>إلغاء</Btn>
                      </div>
                      :<Btn variant="ghost" onClick={()=>{setRenameId(s.id);setRenameVal(s.name);}} style={{padding:"4px 10px",fontSize:11}}>✏️ تعديل الاسم</Btn>}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ══════════════════ DASHBOARD ══════════════════════════════════════════════
function DashboardTab({myDelegates,accepted,totalOrders}){
  const pending=myDelegates.filter(d=>d.status==="قيد المراجعة");
  const rejected=myDelegates.filter(d=>d.status==="مرفوض");
  return(
    <div>
      <h2 style={{color:C.text,margin:"0 0 20px",fontSize:20}}>لوحة التحكم</h2>
      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:26}}>
        <StatBox label="إجمالي المناديب"  value={myDelegates.length} accent={C.blue}/>
        <StatBox label="المقبولون"         value={accepted.length}   accent={C.green}/>
        <StatBox label="قيد المراجعة"      value={pending.length}    accent={C.yellow}/>
        <StatBox label="المرفوضون"         value={rejected.length}   accent={C.red}/>
        <StatBox label="إجمالي الأوردرات" value={totalOrders.toLocaleString()} accent={C.purple} sub="لجميع المقبولين"/>
      </div>
      {accepted.length>0&&(
        <Card style={{marginBottom:22}}>
          <h3 style={{margin:"0 0 16px",color:C.text,fontSize:15}}>📦 أوردرات المناديب المقبولين</h3>
          {accepted.map(d=>{
            const pct=totalOrders>0?(d.orders/totalOrders)*100:0;
            return(
              <div key={d.id} style={{marginBottom:14}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                  <span style={{color:C.text,fontSize:13}}>{d.name} <span style={{color:"#445"}}>({d.id})</span></span>
                  <span style={{color:C.purple,fontWeight:700}}>{(d.orders||0).toLocaleString()}</span>
                </div>
                <div style={{background:C.border,borderRadius:6,height:8}}>
                  <div style={{background:`linear-gradient(90deg,${C.blue},${C.purple})`,height:8,borderRadius:6,width:`${pct}%`,transition:"width .5s"}}/>
                </div>
                <div style={{color:"#445",fontSize:11,marginTop:3}}>النسبة: <span style={{color:C.green}}>{d.commission_rate>0?`${d.commission_rate}%`:"لم تُحدد بعد"}</span> | العمولة: <span style={{color:C.yellow}}>{((d.orders||0)*(d.commission_rate||0)/100).toFixed(0)} ج</span></div>
              </div>
            );
          })}
        </Card>
      )}
      {pending.length>0&&(
        <Card>
          <h3 style={{margin:"0 0 16px",color:C.text,fontSize:15}}>⏳ طلبات قيد المراجعة من مدير التشغيل ({pending.length})</h3>
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            {pending.map(d=>(
              <div key={d.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:C.panel,padding:"12px 16px",borderRadius:10,border:`1px solid ${C.border}`,flexWrap:"wrap",gap:10}}>
                <div>
                  <div style={{fontWeight:600,color:C.text}}>{d.name} <span style={{color:C.muted,fontSize:12}}>({d.id})</span></div>
                  <div style={{color:C.muted,fontSize:12,marginTop:2}}>{d.phone} | {d.vehicle_type==="موتوسيكل"?"🏍️":"🚲"} {d.vehicle_type}</div>
                </div>
                <Badge status={d.status}/>
              </div>
            ))}
          </div>
        </Card>
      )}
      {myDelegates.length===0&&<Card><div style={{textAlign:"center",padding:"40px",color:C.muted}}><div style={{fontSize:48,marginBottom:12}}>📭</div><div>لا يوجد مناديب حتى الآن.</div></div></Card>}
    </div>
  );
}

// ══════════════════ UPLOAD DOC ════════════════════════════════════════════
function UploadDocTab({supervisors,setDelegates,notify,currentSup,addNotifDB}){
  const [form,setForm]=useState({name:"",phone:"",nationalId:"",address:"",vehicleType:"موتوسيكل"});
  const [docs,setDocs]=useState({selfie:null,nationalFront:null,nationalBack:null,licenseFront:null,licenseBack:null});
  const [saving,setSaving]=useState(false);
  const setDoc=(k,v)=>setDocs(prev=>({...prev,[k]:v}));
  const needLicense=form.vehicleType==="موتوسيكل";

  const validate=()=>{
    if(!form.name.trim())      {notify("❗ أدخل اسم المندوب","error");        return false;}
    if(!form.phone.trim())     {notify("❗ أدخل رقم الهاتف","error");         return false;}
    if(!form.nationalId.trim()){notify("❗ أدخل الرقم القومي","error");       return false;}
    if(!form.address.trim())   {notify("❗ أدخل عنوان السكن","error");        return false;}
    if(!docs.selfie)           {notify("❗ صورة السيلفي مطلوبة","error");     return false;}
    if(!docs.nationalFront)    {notify("❗ وش البطاقة مطلوب","error");        return false;}
    if(!docs.nationalBack)     {notify("❗ ظهر البطاقة مطلوب","error");       return false;}
    if(needLicense&&!docs.licenseFront){notify("❗ وش الرخصة مطلوب","error"); return false;}
    if(needLicense&&!docs.licenseBack) {notify("❗ ظهر الرخصة مطلوب","error");return false;}
    return true;
  };

  const handleSubmit=async()=>{
    if(!validate())return;
    setSaving(true);
    const nd={
      id:genId("DEL"), supervisor_id:currentSup.id,
      name:form.name.trim(), phone:form.phone.trim(),
      national_id:form.nationalId.trim(), address:form.address.trim(),
      status:"قيد المراجعة", commission_rate:0, orders:0,
      vehicle_type:form.vehicleType,
      docs:{
        selfie:docs.selfie, nationalFront:docs.nationalFront, nationalBack:docs.nationalBack,
        licenseFront:needLicense?docs.licenseFront:null, licenseBack:needLicense?docs.licenseBack:null
      }
    };
    try{
      const result=await dbInsert("delegates",nd);
      if(!result.ok){
        notify(`❌ فشل الحفظ: ${result.error || result.status}`,"error");
        return;
      }
      setDelegates(prev=>[...(Array.isArray(prev)?prev:[]),nd]);
      notify(`✅ تم إضافة ${form.name} — ID: ${nd.id}`);
      setForm({name:"",phone:"",nationalId:"",address:"",vehicleType:"موتوسيكل"});
      setDocs({selfie:null,nationalFront:null,nationalBack:null,licenseFront:null,licenseBack:null});
    }catch(e){notify("❌ حدث خطأ أثناء الحفظ: "+e.message,"error");}
    finally{setSaving(false);}
  };

  return(
    <div style={{maxWidth:680}}>
      <h2 style={{color:C.text,marginBottom:6}}>📄 إضافة مندوب جديد</h2>
      <p style={{color:C.muted,marginBottom:22}}>أدخل البيانات وارفع المستندات المطلوبة — نسبة العمولة يحددها مدير التشغيل لاحقاً</p>
      <Card>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <Inp label="اسم المندوب *" placeholder="محمد أحمد" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>
          <Inp label="رقم الهاتف *" placeholder="05XXXXXXXX" value={form.phone} onChange={e=>setForm({...form,phone:e.target.value})}/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <Inp label="الرقم القومي *" placeholder="29XXXXXXXXXXXX" value={form.nationalId} onChange={e=>setForm({...form,nationalId:e.target.value})}/>
          <Sel label="وسيلة التوصيل" value={form.vehicleType} onChange={e=>setForm({...form,vehicleType:e.target.value})} options={[{value:"موتوسيكل",label:"🏍️ موتوسيكل"},{value:"دراجة هوائية",label:"🚲 دراجة هوائية"}]}/>
        </div>
        <Inp label="عنوان السكن *" placeholder="المحافظة، المدينة، الحي، الشارع" value={form.address} onChange={e=>setForm({...form,address:e.target.value})}/>

        <div style={{marginBottom:8,marginTop:8,color:C.muted,fontSize:13,fontWeight:600}}>📸 المستندات المطلوبة</div>
        <div style={{background:C.panel,borderRadius:10,padding:"12px 14px",marginBottom:16}}>
          <div style={{color:C.yellow,fontSize:12,marginBottom:8}}>⚠️ تعليمات مهمة:</div>
          <div style={{color:"#667",fontSize:12,lineHeight:1.9}}>
            • <strong style={{color:C.muted}}>السيلفي:</strong> وجه واضح أمام خلفية سادة<br/>
            • <strong style={{color:C.muted}}>وش البطاقة:</strong> واضح ومقصوص بدقة — بدون أي أضواء أو فلاش<br/>
            • <strong style={{color:C.muted}}>ظهر البطاقة:</strong> واضح ومقصوص بدقة — بدون أي أضواء أو فلاش<br/>
            {needLicense
              ?<>• <strong style={{color:C.muted}}>وش وظهر رخصة الموتوسيكل:</strong> مطلوبة لمندوبي الموتوسيكل</>
              :<span style={{color:"#445"}}>• الرخصة: غير مطلوبة للدراجة الهوائية</span>}
          </div>
        </div>

        <div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
          <PhotoBox label="سيلفي" hint="وجه أمام خلفية سادة" required value={docs.selfie} onChange={v=>setDoc("selfie",v)}/>
          <PhotoBox label="وش البطاقة" hint="بدون أضواء — مقصوص" required value={docs.nationalFront} onChange={v=>setDoc("nationalFront",v)}/>
          <PhotoBox label="ظهر البطاقة" hint="بدون أضواء — مقصوص" required value={docs.nationalBack} onChange={v=>setDoc("nationalBack",v)}/>
          {needLicense&&(
            <>
              <PhotoBox label="وش الرخصة" hint="رخصة قيادة موتوسيكل" required value={docs.licenseFront} onChange={v=>setDoc("licenseFront",v)}/>
              <PhotoBox label="ظهر الرخصة" hint="رخصة قيادة موتوسيكل" required value={docs.licenseBack} onChange={v=>setDoc("licenseBack",v)}/>
            </>
          )}
        </div>
        <Btn onClick={handleSubmit} disabled={saving} style={{width:"100%",padding:"12px",marginTop:18,fontSize:15}}>
          {saving?"⏳ جاري الحفظ...":"✅ إرسال للمراجعة"}
        </Btn>
      </Card>
    </div>
  );
}

// ══════════════════ DELEGATES ═════════════════════════════════════════════
function DelegatesTab({myDelegates}){
  const [search,setSearch]=useState("");
  const [filter,setFilter]=useState("الكل");
  const [preview,setPreview]=useState(null);
  const shown=myDelegates.filter(d=>{
    const ms=d.name.includes(search)||d.id.includes(search)||d.phone.includes(search);
    const mf=filter==="الكل"||d.status===filter;
    return ms&&mf;
  });
  return(
    <div>
      {preview&&(
        <div onClick={()=>setPreview(null)} style={{position:"fixed",inset:0,background:"#000c",zIndex:9998,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div onClick={e=>e.stopPropagation()} style={{background:C.card,borderRadius:14,padding:20,maxWidth:400,width:"90%",textAlign:"center"}}>
            <div style={{color:C.text,fontWeight:700,marginBottom:12}}>{preview.label}</div>
            <img src={preview.src} alt="" style={{width:"100%",borderRadius:10,maxHeight:380,objectFit:"contain"}}/>
            <Btn variant="ghost" onClick={()=>setPreview(null)} style={{marginTop:14,width:"100%"}}>إغلاق</Btn>
          </div>
        </div>
      )}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18,flexWrap:"wrap",gap:10}}>
        <h2 style={{color:C.text,margin:0}}>👥 المناديب ({myDelegates.length})</h2>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <input placeholder="بحث..." value={search} onChange={e=>setSearch(e.target.value)} style={{background:C.card,border:`1px solid ${C.border}`,color:C.text,borderRadius:8,padding:"7px 13px",fontSize:13,outline:"none",width:170}}/>
          {["الكل","مقبول","مرفوض","قيد المراجعة"].map(s=>(
            <button key={s} onClick={()=>setFilter(s)} style={{background:filter===s?C.blue:C.card,color:filter===s?"#fff":C.muted,border:`1px solid ${C.border}`,borderRadius:8,padding:"7px 12px",cursor:"pointer",fontSize:12,fontWeight:600}}>{s}</button>
          ))}
        </div>
      </div>
      <p style={{color:"#556",fontSize:12,marginBottom:14}}>مراجعة الحالة وتحديد العمولة من مسؤولية مدير التشغيل</p>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead>
            <tr style={{background:C.panel,borderBottom:`1px solid ${C.border}`}}>
              {["ID","الاسم","الهاتف","الوسيلة","الحالة","النسبة","الأوردرات","المستندات"].map(h=>(
                <th key={h} style={{padding:"11px 12px",color:C.muted,fontSize:12,fontWeight:700,textAlign:"right",whiteSpace:"nowrap"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {shown.map(d=>{
              const docs=d.docs||{};
              const docList=[
                {key:"selfie",label:"سيلفي",icon:"🤳"},
                {key:"nationalFront",label:"وش البطاقة",icon:"🪪"},
                {key:"nationalBack",label:"ظهر البطاقة",icon:"🪪"},
                {key:"licenseFront",label:"وش الرخصة",icon:"📋"},
                {key:"licenseBack",label:"ظهر الرخصة",icon:"📋"},
              ].filter(x=>docs[x.key]);
              return(
                <tr key={d.id} style={{borderBottom:`1px solid ${C.border}22`}}>
                  <td style={{padding:"12px",fontFamily:"monospace",color:C.blue,fontSize:12}}>{d.id}</td>
                  <td style={{padding:"12px",color:C.text,fontWeight:600}}>{d.name}</td>
                  <td style={{padding:"12px",color:C.muted,fontSize:13}}>{d.phone}</td>
                  <td style={{padding:"12px",color:C.muted,fontSize:13}}>{d.vehicle_type==="موتوسيكل"?"🏍️":"🚲"}</td>
                  <td style={{padding:"12px"}}><Badge status={d.status}/></td>
                  <td style={{padding:"12px"}}>
                    <span style={{color:d.commission_rate>0?C.green:"#445",fontWeight:700}}>{d.commission_rate>0?`${d.commission_rate}%`:"لم تُحدد بعد"}</span>
                  </td>
                  <td style={{padding:"12px",color:C.purple,fontWeight:700}}>{d.status==="مقبول"?(d.orders||0).toLocaleString():"—"}</td>
                  <td style={{padding:"12px"}}>
                    <div style={{display:"flex",gap:5}}>
                      {docList.length>0
                        ?docList.map(x=>(
                          <span key={x.key} title={x.label} style={{cursor:"pointer",fontSize:18}}
                            onClick={()=>setPreview({src:docs[x.key],label:x.label})}>{x.icon}</span>
                        ))
                        :<span style={{color:"#445",fontSize:11}}>—</span>}
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {shown.length===0&&<div style={{textAlign:"center",padding:"32px",color:C.muted}}>لا توجد نتائج</div>}
      </div>
    </div>
  );
}

// ══════════════════ ORDERS ════════════════════════════════════════════════
function OrdersTab({delegates,setDelegates,currentSup,notify,myDelegates,addNotifDB}){
  const accepted=myDelegates.filter(d=>d.status==="مقبول");
  const totalOrders=accepted.reduce((s,d)=>s+(d.orders||0),0);
  const activeLastWeek=accepted.filter(d=>(d.orders||0)>0).length;

  return(
    <div>
      <h2 style={{color:C.text,marginBottom:6}}>📦 ملخص الأوردرات</h2>
      <p style={{color:C.muted,marginBottom:22}}>بيانات الأوردرات يتم تحديثها أسبوعياً من قبل مدير التشغيل</p>

      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:24}}>
        <StatBox label="مناديب نشطون آخر أسبوع" value={activeLastWeek} accent={C.green} sub={`من إجمالي ${accepted.length} مقبول`}/>
        <StatBox label="إجمالي الأوردرات المكتملة" value={totalOrders.toLocaleString()} accent={C.purple}/>
        <StatBox label="إجمالي العمولات" value={accepted.reduce((s,d)=>s+((d.orders||0)*(d.commission_rate||0)/100),0).toFixed(0)} accent={C.yellow}/>
      </div>

      <Card>
        <h3 style={{color:C.text,margin:"0 0 14px",fontSize:15}}>تفاصيل المناديب المقبولين</h3>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead>
            <tr style={{background:C.panel,borderBottom:`1px solid ${C.border}`}}>
              {["الاسم","الوسيلة","الأوردرات (آخر أسبوع)","النسبة","العمولة"].map(h=><th key={h} style={{padding:"10px 12px",color:C.muted,fontSize:12,fontWeight:700,textAlign:"right"}}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {accepted.map(d=>(
              <tr key={d.id} style={{borderBottom:`1px solid ${C.border}22`}}>
                <td style={{padding:"11px 12px",color:C.text}}>{d.name}</td>
                <td style={{padding:"11px 12px",color:C.muted,fontSize:13}}>{d.vehicle_type==="موتوسيكل"?"🏍️":"🚲"}</td>
                <td style={{padding:"11px 12px",color:C.purple,fontWeight:700}}>{(d.orders||0).toLocaleString()}</td>
                <td style={{padding:"11px 12px",color:C.green}}>{d.commission_rate>0?`${d.commission_rate}%`:<span style={{color:"#445"}}>لم تُحدد بعد</span>}</td>
                <td style={{padding:"11px 12px",color:C.yellow}}>{((d.orders||0)*(d.commission_rate||0)/100).toFixed(0)}</td>
              </tr>
            ))}
            {accepted.length===0&&(
              <tr><td colSpan={5} style={{padding:"20px",textAlign:"center",color:C.muted}}>لا يوجد مناديب مقبولون حتى الآن</td></tr>
            )}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ══════════════════ SUPERVISOR MANAGEMENT ════════════════════════════════
function SupervisorTab({supervisors,setSupervisors,notify}){
  const [form,setForm]=useState({name:"",phone:"",password:"",email:""});
  const [showP,setShowP]=useState({});
  const [saving,setSaving]=useState(false);
  const add=async()=>{
    if(!form.name.trim())    {notify("❗ أدخل الاسم","error");        return;}
    if(!form.phone.trim())   {notify("❗ أدخل رقم الهاتف","error");  return;}
    if(!form.password.trim()){notify("❗ أدخل كلمة المرور","error"); return;}
    if(supervisors.find(s=>s.phone===form.phone)){notify("❗ رقم مسجل مسبقاً","error");return;}
    setSaving(true);
    try{
      const ns={id:genId("SUP"),name:form.name.trim(),phone:form.phone.trim(),password_hash:form.password.trim(),email:form.email.trim(),role:"supervisor"};
      const result=await dbInsert("supervisors",ns);
      if(!result.ok){
        notify(`❌ فشل الحفظ: ${result.error || result.status}`,"error");
        return;
      }
      setSupervisors(prev=>[...prev,ns]);
      notify(`✅ تم إضافة ${form.name} — ID: ${ns.id}`);
      setForm({name:"",phone:"",password:"",email:""});
    }catch(e){notify("❌ حدث خطأ: "+e.message,"error");}
    finally{setSaving(false);}
  };
  return(
    <div style={{maxWidth:680}}>
      <h2 style={{color:C.text,marginBottom:6}}>🏢 إدارة المشرفين</h2>
      <p style={{color:C.muted,marginBottom:22}}>إضافة مشرفين جدد — البيانات تُحفظ في قاعدة البيانات</p>
      <Card style={{marginBottom:22}}>
        <h3 style={{color:C.text,margin:"0 0 16px",fontSize:15}}>➕ مشرف جديد</h3>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <Inp label="الاسم" placeholder="أحمد محمود" value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/>
          <Inp label="البريد الإلكتروني" placeholder="ahmed@co.com" value={form.email} onChange={e=>setForm({...form,email:e.target.value})}/>
          <Inp label="رقم الهاتف (للدخول)" placeholder="05XXXXXXXX" value={form.phone} onChange={e=>setForm({...form,phone:e.target.value})}/>
          <div style={{marginBottom:14}}>
            <label style={{display:"block",color:C.muted,fontSize:12,marginBottom:5}}>كلمة المرور</label>
            <div style={{position:"relative"}}>
              <input type={showP.new?"text":"password"} placeholder="••••••" value={form.password} onChange={e=>setForm({...form,password:e.target.value})}
                style={{width:"100%",background:C.panel,border:`1px solid ${C.border}`,color:C.text,borderRadius:8,padding:"9px 38px 9px 13px",fontSize:14,outline:"none",boxSizing:"border-box"}}/>
              <span onClick={()=>setShowP(s=>({...s,new:!s.new}))} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",cursor:"pointer",fontSize:15,color:C.muted}}>{showP.new?"🙈":"👁️"}</span>
            </div>
          </div>
        </div>
        <Btn onClick={add} disabled={saving} style={{width:"100%",padding:"11px"}}>{saving?"⏳ جاري الحفظ...":"إضافة مشرف"}</Btn>
      </Card>
      <Card>
        <h3 style={{color:C.text,margin:"0 0 14px",fontSize:15}}>قائمة المشرفين ({supervisors.length})</h3>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {supervisors.map(s=>(
            <div key={s.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:C.panel,padding:"14px 18px",borderRadius:10,border:`1px solid ${C.border}`,flexWrap:"wrap",gap:8}}>
              <div>
                <div style={{fontWeight:700,color:C.text}}>{s.name}</div>
                <div style={{color:C.muted,fontSize:12,marginTop:3}}>📱 {s.phone} | 📧 {s.email}</div>
                <div style={{color:"#445",fontSize:11,marginTop:2}}>
                  🔑 <span style={{color:C.muted,fontFamily:"monospace"}}>{showP[s.id]?s.password_hash:"••••"}</span>
                  <span onClick={()=>setShowP(p=>({...p,[s.id]:!p[s.id]}))} style={{marginRight:6,cursor:"pointer",fontSize:12}}>{showP[s.id]?"🙈":"👁️"}</span>
                </div>
              </div>
              <div style={{background:`${C.blue}22`,color:C.blue,padding:"5px 14px",borderRadius:20,fontFamily:"monospace",fontSize:13,fontWeight:700}}>{s.id}</div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}
