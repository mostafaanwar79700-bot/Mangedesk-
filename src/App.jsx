
import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// ─── Supabase Client (real client — needed for Auth + RLS to work) ───────
const SUPA_URL = "https://hlksnbrzumzfgjsefxgv.supabase.co";
const SUPA_KEY = "sb_publishable_POxQAK5GFStpapGU4q_aUA_GzRw84bq";
const supa = createClient(SUPA_URL, SUPA_KEY);
const supaAdmin = createClient(SUPA_URL, SUPA_KEY, { auth: { persistSession:false, autoRefreshToken:false, detectSessionInUrl:false } });

// Supabase Auth requires an email-shaped identity. Since supervisors log in
// with a phone number, we map phone -> a deterministic fake email under a
// domain we own conceptually. The phone number itself is never shown to
// the person — they only ever see/type the phone number in the UI.
const phoneToFakeEmail = (phone) => `phone_${phone.trim()}@managedesk.internal`;

// ─── Simple DB helpers — now backed by the real client so RLS (auth.uid())
// is correctly applied using the logged-in user's session token. ──────────
const dbSelect = async (table, filters={}, containsFilter=null) => {
  try {
    let q = supa.from(table).select("*");
    Object.entries(filters).forEach(([k,v]) => { q = q.eq(k, v); });
    if (containsFilter) Object.entries(containsFilter).forEach(([k,v]) => { q = q.contains(k, v); });
    const { data, error } = await q;
    if (error) { console.error("dbSelect error", error); return []; }
    return Array.isArray(data) ? data : [];
  } catch(e) { console.error("dbSelect error", e); return []; }
};

const dbSelectOne = async (table, filters={}) => {
  try {
    let q = supa.from(table).select("*");
    Object.entries(filters).forEach(([k,v]) => { q = q.eq(k, v); });
    const { data, error } = await q.limit(1);
    if (error || !data || data.length===0) return null;
    return data[0];
  } catch(e) { return null; }
};

const dbInsert = async (table, rows) => {
  try {
    const { error } = await supa.from(table).insert(rows);
    if (error) {
      console.error("dbInsert failed", error);
      return { ok:false, status:error.code, error:error.message };
    }
    return { ok:true };
  } catch(e) {
    console.error("dbInsert network error", e);
    return { ok:false, status:0, error:e.message };
  }
};

const dbUpdate = async (table, vals, filters={}) => {
  try {
    let q = supa.from(table).update(vals);
    Object.entries(filters).forEach(([k,v]) => { q = q.eq(k, v); });
    const { error } = await q;
    if (error) console.error("dbUpdate failed", error);
    return !error;
  } catch(e) { return false; }
};

const dbDelete = async (table, filters={}) => {
  try {
    let q = supa.from(table).delete();
    Object.entries(filters).forEach(([k,v]) => { q = q.eq(k, v); });
    const { error } = await q;
    return !error;
  } catch(e) { return false; }
};

// ─── Auth helpers ──────────────────────────────────────────────────────
const authSignIn = async (phone, password) => {
  const email = phoneToFakeEmail(phone);
  const { data, error } = await supa.auth.signInWithPassword({ email, password });
  if (error) return { ok:false, error: error.message };
  return { ok:true, session: data.session, authUser: data.user };
};

const authSignUp = async (phone, password) => {
  const email = phoneToFakeEmail(phone);
  const { data, error } = await supaAdmin.auth.signUp({ email, password });
  if (error) return { ok:false, error: error.message };
  return { ok:true, authUser: data.user };
};

const authSignOut = async () => {
  await supa.auth.signOut();
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
bg:"#12203a", panel:"#182a4a", card:"#1e335a", border:"#2c4570",
blue:"#f97316", green:"#22c55e", red:"#ef4444", yellow:"#eab308",
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
  const [mode,setMode]=useState("login");
  const [phone,setPhone]=useState("");
  const [pass,setPass]=useState("");
  const [err,setErr]=useState("");
  const [show,setShow]=useState(false);
  const [loading,setLoading]=useState(false);

  const [suName,setSuName]=useState("");
  const [suPhone,setSuPhone]=useState("");
  const [suPass,setSuPass]=useState("");
  const [suZone,setSuZone]=useState("");
  const [suErr,setSuErr]=useState("");
  const [suLoading,setSuLoading]=useState(false);

  const attempt=async ()=>{
    if(!phone.trim()||!pass.trim()){setErr("أدخل رقم الهاتف وكلمة المرور");return;}
    setLoading(true);setErr("");
    try{
      const result = await authSignIn(phone.trim(), pass.trim());
      if(!result.ok){ setErr("رقم الهاتف أو كلمة المرور غير صحيحة"); return; }
      const profile = await dbSelectOne("supervisors", { auth_id: result.authUser.id });
      if(!profile){
        setErr("تم تسجيل الدخول لكن لم يتم العثور على بيانات الحساب، تواصل مع مدير التشغيل");
        await authSignOut();
        return;
      }
     if(profile.status==="blocked"){setErr("تم حظر حسابك، تواصل مع مدير التشغيل");await authSignOut();return;}
      onLogin(profile);
    }catch(e){setErr("حدث خطأ، حاول مجددا");}
    finally{setLoading(false);}
  };

  const attemptSignup=async ()=>{
    if(!suName.trim()){setSuErr("أدخل الاسم");return;}
    if(!suPhone.trim()){setSuErr("أدخل رقم الهاتف");return;}
    if(suPass.trim().length<6){setSuErr("كلمة المرور يجب أن تكون 6 أحرف على الأقل");return;}
    if(!suZone.trim()){setSuErr("أدخل الزون");return;}
    setSuLoading(true);setSuErr("");
    try{
      const { data: signUpData, error: signUpError } = await supa.auth.signUp({ email: phoneToFakeEmail(suPhone.trim()), password: suPass.trim() });
      const signUpResult = signUpError ? { ok:false, error: signUpError.message } : { ok:true, authUser: signUpData.user };
      if(!signUpResult.ok){ setSuErr("فشل إنشاء الحساب: "+signUpResult.error); return; }
      const newId = genId("SUP");
      const profileResult = await dbInsert("supervisors",{
        id:newId, auth_id:signUpResult.authUser.id, name:suName.trim(),
        phone:suPhone.trim(), zone:suZone.trim(), role:"supervisor", bike_rate:0, moto_rate:0,
      });
      if(!profileResult.ok){ setSuErr("فشل حفظ بيانات الحساب: "+profileResult.error); return; }
      const loginResult = await authSignIn(suPhone.trim(), suPass.trim());
      if(loginResult.ok){
        const profile = await dbSelectOne("supervisors", { auth_id: loginResult.authUser.id });
        if(profile){ onLogin(profile); return; }
      }
      setMode("login"); setPhone(suPhone.trim());
      setSuName("");setSuPhone("");setSuPass("");setSuZone("");
    }catch(e){setSuErr("حدث خطأ، حاول مجددا");}
    finally{setSuLoading(false);}
  };

  return(
  <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center"}} dir="rtl">
    <div style={{width:"100%",maxWidth:420,padding:"0 16px"}}>
      <div style={{textAlign:"center",marginBottom:32}}>
<img src="/png.jpg" alt="Raider Assist" style={{width:92,height:92,borderRadius:22,marginBottom:12,boxShadow:"0 10px 26px rgba(0,0,0,.4)"}}/>
<div style={{fontSize:24,fontWeight:800,color:C.text}}>Raider Assist</div>
<div style={{color:C.muted,fontSize:13,marginTop:4}}>ندعم المندوب . تنجح أنت</div>
</div>
      {mode==="login" ? (
        <Card>
          <div style={{fontSize:17,fontWeight:700,color:C.text,marginBottom:20,textAlign:"center"}}>تسجيل الدخول</div>
          <Inp label="رقم الهاتف" type="tel" placeholder="05XXXXXXXX" value={phone} onChange={e=>setPhone(e.target.value)}/>
          <div style={{marginBottom:14}}>
            <label style={{display:"block",color:C.muted,fontSize:12,marginBottom:5}}>كلمة المرور</label>
            <div style={{position:"relative"}}>
              <input type={show?"text":"password"} placeholder="••••••" value={pass} onChange={e=>setPass(e.target.value)}
               style={{width:"100%",background:C.panel,border:`1px solid ${C.border}`,color:C.text,borderRadius:8,padding:"10px 12px"}}/>
              <span onClick={()=>setShow(s=>!s)} style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",cursor:"pointer",color:C.mut
