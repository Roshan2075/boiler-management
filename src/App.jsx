
import React,{useEffect,useMemo,useState} from 'react';
import {createClient} from '@supabase/supabase-js';
import * as XLSX from 'xlsx-js-style';
import {
 LayoutDashboard,ClipboardList,Receipt,FileText,FolderOpen,LogOut,
 Upload,Download,Plus,RefreshCw,LockKeyhole,Factory,Menu,X,Trash2
} from 'lucide-react';

const url=import.meta.env.VITE_SUPABASE_URL;
const key=import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY||import.meta.env.VITE_SUPABASE_ANON_KEY;
const supabase=(url&&key)?createClient(url,key):null;

const DEPTS=[
 'Structure','Pressure Part','Non Pressure Part','Rotary Part',
 'Bunker','LP Piping','Trim Piping'
];

const MONTHS=[
 'January','February','March','April','May','June',
 'July','August','September','October','November','December'
];

const blank={
 work_date:new Date().toISOString().slice(0,10),
 department:'Structure',pg:'',ma:'',du:'',description:'',
 quantity:0,unit_weight:0,weight:0,
 pre_assembly:0,placement:0,alignment:0,bolting:0,remark:''
};

function numberValue(v){
 if(v===null||v===undefined||v==='')return 0;
 const n=Number(String(v).replace(/,/g,'').trim());
 return Number.isFinite(n)?n:0;
}

function normHeader(v){
 return String(v??'').toLowerCase().replace(/[^a-z0-9]/g,'');
}

function excelDate(v){
 if(v instanceof Date&&!isNaN(v)){
  return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
 }
 if(typeof v==='number'){
  const d=XLSX.SSF.parse_date_code(v);
  if(d){
   return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
 }
 const s=String(v??'').trim();
 if(!s)return '';

 let m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
 if(m){
  let year=Number(m[3]);
  if(year<100)year+=2000;
  return `${year}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
 }

 m=s.match(/^(\d{1,2})[\s-]+([A-Za-z]{3,})[\s-]+(\d{2,4})$/);
 if(m){
  const months=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  const month=months.findIndex(x=>m[2].toLowerCase().startsWith(x))+1;
  let year=Number(m[3]);
  if(year<100)year+=2000;
  if(month)return `${year}-${String(month).padStart(2,'0')}-${m[1].padStart(2,'0')}`;
 }
 if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
 return '';
}

function departmentName(v){
 const n=normHeader(v);
 const map={
  structure:'Structure',
  pressurepart:'Pressure Part',
  nonpressurepart:'Non Pressure Part',
  rotarypart:'Rotary Part',
  bunker:'Bunker',
  lppiping:'LP Piping',
  trimpiping:'Trim Piping',
  trimpipe:'Trim Piping'
 };
 return map[n]||'';
}

const DPR_HEADERS=[
 'Date','PG','MA','DU','Description','Quantity','Unit Weight',
 'Total Weight','Pre-Assembly Qty','Pre-Assembly Weight',
 'Placement Qty','Placement Weight','Alignment Qty',
 'Alignment Weight','Bolting Qty','Bolting Weight','Remarks'
];

function stageWeight(qty,unitWeight){
 return numberValue(qty)*numberValue(unitWeight);
}

function App(){
 const [session,setSession]=useState(null);
 const [email,setEmail]=useState('');
 const [password,setPassword]=useState('');
 const [tab,setTab]=useState('dashboard');
 const [rows,setRows]=useState([]);
 const [bills,setBills]=useState([]);
 const [docs,setDocs]=useState([]);
 const [form,setForm]=useState(blank);
 const [filterMonth,setFilterMonth]=useState(new Date().getMonth());
 const [filterYear,setFilterYear]=useState(new Date().getFullYear());
 const [filterDept,setFilterDept]=useState('All');
 const [notice,setNotice]=useState('');
 const [busy,setBusy]=useState(false);
 const [mobile,setMobile]=useState(false);

 useEffect(()=>{
  if(!supabase)return;
  supabase.auth.getSession().then(({data})=>setSession(data.session));
  const {data:{subscription}}=supabase.auth.onAuthStateChange((_e,s)=>setSession(s));
  return ()=>subscription.unsubscribe();
 },[]);

 useEffect(()=>{
  if(session)loadAll();
 },[session]);

 async function loadAll(){
  if(!supabase||!session)return;
  setBusy(true);
  const [d,b,f]=await Promise.all([
   supabase.from('dpr_entries').select('*').order('work_date',{ascending:false}),
   supabase.from('ra_bills').select('*').order('created_at',{ascending:false}),
   supabase.from('drawings').select('*').order('created_at',{ascending:false})
  ]);
  if(d.error)setNotice('DPR Load Error: '+d.error.message);
  else setRows(d.data||[]);
  if(b.error)setNotice('RA Bill Load Error: '+b.error.message);
  else setBills(b.data||[]);
  if(f.error)setNotice('Drawing Load Error: '+f.error.message);
  else setDocs(f.data||[]);
  setBusy(false);
 }

 async function login(e){
  e.preventDefault();
  if(!supabase)return;
  const {error}=await supabase.auth.signInWithPassword({email,password});
  setNotice(error?.message||'Login successful');
 }

 async function saveDpr(e){
  e.preventDefault();
  const qty=numberValue(form.quantity);
  const unitWeight=numberValue(form.unit_weight);
  if(!form.work_date||!form.pg||!form.ma||!form.du){
   setNotice('Date, PG, MA और DU भरना जरूरी है।');
   return;
  }
  const payload={
   work_date:form.work_date,
   department:form.department,
   pg:String(form.pg||'').trim(),
   ma:String(form.ma||'').trim(),
   du:String(form.du||'').trim(),
   description:String(form.description||''),
   quantity:qty,
   unit_weight:unitWeight,
   weight:qty*unitWeight,
   pre_assembly:numberValue(form.pre_assembly),
   placement:numberValue(form.placement),
   alignment:numberValue(form.alignment),
   bolting:numberValue(form.bolting),
   remark:String(form.remark||''),
   user_id:session.user.id
  };
  const {error}=await supabase.from('dpr_entries').upsert(payload,{
   onConflict:'user_id,work_date,department,pg,ma,du'
  });
  if(error){
   setNotice('DPR Save Error: '+error.message);
   return;
  }
  setNotice('DPR saved / updated successfully');
  setForm({...blank,work_date:form.work_date,department:form.department});
  await loadAll();
 }

 async function delDpr(id){
  if(!confirm('Delete this DPR entry?'))return;
  const {error}=await supabase.from('dpr_entries').delete().eq('id',id);
  setNotice(error?.message||'DPR deleted');
  if(!error)await loadAll();
 }

 const monthRows=useMemo(()=>rows.filter(r=>{
  const d=new Date(r.work_date+'T00:00:00');
  return d.getMonth()===filterMonth&&d.getFullYear()===filterYear;
 }),[rows,filterMonth,filterYear]);

 const visibleMonthRows=useMemo(()=>{
  return filterDept==='All'?monthRows:monthRows.filter(r=>r.department===filterDept);
 },[monthRows,filterDept]);

 const totals=useMemo(()=>rows.reduce((a,r)=>({
  qty:a.qty+numberValue(r.quantity),
  weight:a.weight+numberValue(r.weight)
 }),{qty:0,weight:0}),[rows]);

 function exportRows(data,name){
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Records');
  XLSX.writeFile(wb,name+'.xlsx');
 }

 function exportDprWorkbook(){
  const wb=XLSX.utils.book_new();
  const headerStyle={
   fill:{fgColor:{rgb:'17365D'}},
   font:{bold:true,color:{rgb:'FFFFFF'},sz:10},
   alignment:{horizontal:'center',vertical:'center',wrapText:true},
   border:{
    top:{style:'thin',color:{rgb:'B4C6E7'}},
    bottom:{style:'thin',color:{rgb:'B4C6E7'}},
    left:{style:'thin',color:{rgb:'B4C6E7'}},
    right:{style:'thin',color:{rgb:'B4C6E7'}}
   }
  };
  const widths=[15,12,12,12,30,14,14,16,17,19,15,18,15,18,14,17,28];

  for(const dept of DEPTS){
   const deptRows=rows.filter(r=>r.department===dept);
   const aoa=[
    ['BOILER CONSTRUCTION MANAGEMENT'],
    [`DAILY PROGRESS REPORT - ${dept}`],
    [`Generated: ${new Date().toLocaleDateString('en-IN')}`],
    DPR_HEADERS
   ];
   deptRows.forEach(r=>{
    aoa.push([
     r.work_date,r.pg,r.ma,r.du,r.description,
     numberValue(r.quantity),numberValue(r.unit_weight),
     numberValue(r.weight),
     numberValue(r.pre_assembly),stageWeight(r.pre_assembly,r.unit_weight),
     numberValue(r.placement),stageWeight(r.placement,r.unit_weight),
     numberValue(r.alignment),stageWeight(r.alignment,r.unit_weight),
     numberValue(r.bolting),stageWeight(r.bolting,r.unit_weight),
     r.remark||''
    ]);
   });
   const ws=XLSX.utils.aoa_to_sheet(aoa);
   ws['!cols']=widths.map(w=>({wch:w}));
   ws['!merges']=[
    {s:{r:0,c:0},e:{r:0,c:16}},
    {s:{r:1,c:0},e:{r:1,c:16}},
    {s:{r:2,c:0},e:{r:2,c:16}}
   ];
   ws['!autofilter']={ref:`A4:Q${Math.max(4,deptRows.length+4)}`};
   ws['!rows']=[{hpt:26},{hpt:22},{hpt:20},{hpt:34}];

   for(let c=0;c<17;c++){
    const cell=XLSX.utils.encode_cell({r:3,c});
    if(ws[cell])ws[cell].s=headerStyle;
   }
   for(let i=0;i<deptRows.length;i++){
    const row=i+5;
    const r=deptRows[i];
    const formulas=[
     ['H',`F${row}*G${row}`,numberValue(r.quantity)*numberValue(r.unit_weight)],
     ['J',`I${row}*G${row}`,stageWeight(r.pre_assembly,r.unit_weight)],
     ['L',`K${row}*G${row}`,stageWeight(r.placement,r.unit_weight)],
     ['N',`M${row}*G${row}`,stageWeight(r.alignment,r.unit_weight)],
     ['P',`O${row}*G${row}`,stageWeight(r.bolting,r.unit_weight)]
    ];
    formulas.forEach(([col,f,v])=>{
     const cell=col+row;
     ws[cell]={t:'n',f,v,s:{numFmt:'0.000'}};
    });
    for(let c=0;c<17;c++){
     const cell=XLSX.utils.encode_cell({r:row-1,c});
     if(ws[cell])ws[cell].s={
      ...(ws[cell].s||{}),
      border:{
       top:{style:'thin',color:{rgb:'D9E2F3'}},
       bottom:{style:'thin',color:{rgb:'D9E2F3'}},
       left:{style:'thin',color:{rgb:'D9E2F3'}},
       right:{style:'thin',color:{rgb:'D9E2F3'}}
      }
     };
    }
   }
   XLSX.utils.book_append_sheet(wb,ws,dept);
  }
  XLSX.writeFile(wb,'Boiler_Construction_DPR.xlsx');
 }

 async function importDpr(e){
  const file=e.target.files?.[0];
  if(!file)return;
  setBusy(true);
  setNotice('Reading Excel file...');
  try{
   const wb=XLSX.read(await file.arrayBuffer(),{cellDates:true});
   let mapped=[];
   let skipped=0;
   for(const sheetName of wb.SheetNames){
    const sheet=wb.Sheets[sheetName];
    if(!sheet)continue;
    const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:true});
    if(!matrix.length)continue;
    let headerIndex=-1;
    let headers=[];
    for(let i=0;i<Math.min(matrix.length,25);i++){
     const h=matrix[i].map(normHeader);
     if(h.includes('pg')&&h.includes('ma')&&h.includes('du')&&
       h.some(x=>['date','workdate'].includes(x))){
      headerIndex=i;
      headers=h;
      break;
     }
    }
    if(headerIndex<0)continue;
    const sheetDept=departmentName(sheetName);
    function getValue(row,aliases){
     const index=headers.findIndex(h=>aliases.includes(h));
     return index<0?'':row[index];
    }
    for(const r of matrix.slice(headerIndex+1)){
     if(!r.some(v=>String(v??'').trim()))continue;
     const date=excelDate(getValue(r,['date','workdate']));
     const department=sheetDept||departmentName(getValue(r,['department','dept']));
     const pg=getValue(r,['pg']);
     const ma=getValue(r,['ma']);
     const du=getValue(r,['du']);
     if(!date||!department||!DEPTS.includes(department)){
      skipped++;
      continue;
     }
     const quantity=numberValue(getValue(r,['quantity','qty']));
     const unitWeight=numberValue(getValue(r,['unitweight']));
     const totalWeight=quantity*unitWeight;
     mapped.push({
      user_id:session.user.id,
      work_date:date,
      department,
      pg:String(pg??'').trim(),
      ma:String(ma??'').trim(),
      du:String(du??'').trim(),
      description:String(getValue(r,['description','itemdescription'])??''),
      quantity,
      unit_weight:unitWeight,
      weight:totalWeight,
      pre_assembly:numberValue(getValue(r,['preassemblyqty','preassembly'])),
      placement:numberValue(getValue(r,['placementqty','placement'])),
      alignment:numberValue(getValue(r,['alignmentqty','alignment'])),
      bolting:numberValue(getValue(r,['boltingqty','bolting'])),
      remark:String(getValue(r,['remarks','remark','status'])??'')
     });
    }
   }
   if(!mapped.length){
    throw new Error('Excel में Date, PG, MA, DU वाले सही Header नहीं मिले। Department Sheet का नाम भी जाँचें।');
   }
   let imported=0;
   for(let i=0;i<mapped.length;i+=200){
    const batch=mapped.slice(i,i+200);
    const {error}=await supabase.from('dpr_entries').upsert(batch,{
     onConflict:'user_id,work_date,department,pg,ma,du'
    });
    if(error)throw error;
    imported+=batch.length;
   }
   setNotice(`${imported} DPR rows imported/updated. Skipped rows: ${skipped}`);
   await loadAll();
  }catch(err){
   setNotice('Excel Import Error: '+err.message);
  }finally{
   setBusy(false);
   e.target.value='';
  }
 }

 async function uploadDrawing(e){
  const file=e.target.files?.[0];
  if(!file)return;
  const path=`${session.user.id}/${Date.now()}-${file.name}`;
  const up=await supabase.storage.from('drawings').upload(path,file);
  if(up.error){
   setNotice(up.error.message);
   return;
  }
  const ins=await supabase.from('drawings').insert({
   file_name:file.name,
   storage_path:path,
   department:form.department||'General',
   user_id:session.user.id
  });
  setNotice(ins.error?.message||'Drawing uploaded');
  if(!ins.error)await loadAll();
  e.target.value='';
 }

 async function createRa(){
  const grouped=monthRows.reduce((a,r)=>{
   a.quantity+=numberValue(r.quantity);
   a.weight+=numberValue(r.weight);
   return a;
  },{quantity:0,weight:0});
  const payload={
   bill_month:MONTHS[filterMonth],
   bill_year:filterYear,
   quantity:grouped.quantity,
   weight:grouped.weight,
   user_id:session.user.id
  };
  const {error}=await supabase.from('ra_bills').insert(payload);
  setNotice(error?.message||'Monthly RA summary saved');
  if(!error)await loadAll();
 }

 if(!supabase)return <Setup/>;

 if(!session){
  return (
   <div className="login-wrap">
    <form className="login-card" onSubmit={login}>
     <div className="logo"><Factory size={34}/></div>
     <h1>Boiler Management</h1>
     <p>Construction Management System</p>
     <label>Email</label>
     <input type="email" value={email} onChange={e=>setEmail(e.target.value)} required placeholder="Admin email"/>
     <label>Password</label>
     <input type="password" value={password} onChange={e=>setPassword(e.target.value)} required placeholder="Password"/>
     <button className="primary full"><LockKeyhole size={17}/> Secure Login</button>
     {notice&&<small>{notice}</small>}
     <small>Admin account Supabase Authentication में बनाया जाना चाहिए।</small>
    </form>
   </div>
  );
 }

 const nav=[
  ['dashboard','Dashboard',LayoutDashboard],
  ['dpr','DPR Register',ClipboardList],
  ['ra','RA Bills',Receipt],
  ['customer','Customer Bills',FileText],
  ['drawings','Drawing Folder',FolderOpen]
 ];

 return (
  <div className="app">
   <aside className={mobile?'open':''}>
    <div className="brand">
     <Factory/>
     <span>BOILER<br/><b>MANAGEMENT</b></span>
     <button className="mobile-close" onClick={()=>setMobile(false)}><X/></button>
    </div>
    <div className="nav-label">WORKSPACE</div>
    {nav.map(([id,label,Icon])=>(
     <button key={id} className={'nav '+(tab===id?'active':'')} onClick={()=>{setTab(id);setMobile(false)}}>
      <Icon size={18}/>{label}
     </button>
    ))}
    <div className="side-bottom">
     <span>Secure workspace</span>
     <button className="nav" onClick={()=>supabase.auth.signOut()}><LogOut size={18}/>Logout</button>
    </div>
   </aside>

   <main>
    <header>
     <button className="menu" onClick={()=>setMobile(!mobile)}><Menu/></button>
     <div>
      <b>{nav.find(n=>n[0]===tab)?.[1]||'Dashboard'}</b>
      <small>Boiler Construction Project Workspace</small>
     </div>
     <button className="icon-btn" onClick={loadAll} disabled={busy}><RefreshCw size={18}/></button>
    </header>

    <div className="content">
     {notice&&<div className="notice" onClick={()=>setNotice('')}>{notice}　×</div>}

     {tab==='dashboard'&&(
      <>
       <div className="welcome">
        <div>
         <span className="eyebrow">PROJECT OVERVIEW</span>
         <h1>Master Dashboard</h1>
         <p>All construction records in one place.</p>
        </div>
        <button onClick={()=>setTab('dpr')} className="primary"><Plus size={17}/> Add DPR</button>
       </div>
       <div className="stats">
        <Stat title="DPR Entries" value={rows.length}/>
        <Stat title="Total Quantity" value={totals.qty.toLocaleString()}/>
        <Stat title="Total Weight" value={totals.weight.toLocaleString()}/>
        <Stat title="RA Bills" value={bills.length}/>
       </div>
       <div className="panel">
        <div className="panel-head"><h2>Department Progress</h2><span>All-time quantity & weight</span></div>
        <div className="dept-grid">
         {DEPTS.map(d=>{
          const a=rows.filter(r=>r.department===d);
          return (
           <div className="dept" key={d}>
            <b>{d}</b>
            <strong>{a.reduce((s,r)=>s+numberValue(r.quantity),0).toLocaleString()}</strong>
            <small>{a.length} entries · {a.reduce((s,r)=>s+numberValue(r.weight),0).toLocaleString()} weight</small>
           </div>
          );
         })}
        </div>
       </div>
       <div className="panel">
        <div className="panel-head"><h2>Recent DPR Entries</h2><button className="link" onClick={()=>setTab('dpr')}>View all</button></div>
        <DprTable rows={rows.slice(0,8)} onDelete={delDpr}/>
       </div>
      </>
     )}

     {tab==='dpr'&&(
      <>
       <div className="welcome">
        <div>
         <span className="eyebrow">DAILY PROGRESS REPORT</span>
         <h1>DPR Register</h1>
         <p>Department-wise daily entries with cumulative totals.</p>
        </div>
        <div className="actions">
         <label className="secondary upload">
          <Upload size={16}/> Import Excel
          <input type="file" accept=".xlsx,.xls,.csv" onChange={importDpr}/>
         </label>
         <button className="secondary" onClick={exportDprWorkbook}><Download size={16}/> Export Excel</button>
        </div>
       </div>

       <div className="panel">
        <h2>Add Daily Progress</h2>
        <form className="entry-grid" onSubmit={saveDpr}>
         <Field label="Date"><input type="date" value={form.work_date} onChange={e=>setForm({...form,work_date:e.target.value})} required/></Field>
         <Field label="Department">
          <select value={form.department} onChange={e=>setForm({...form,department:e.target.value})}>
           {DEPTS.map(d=><option key={d} value={d}>{d}</option>)}
          </select>
         </Field>
         {['pg','ma','du','description','quantity','unit_weight','pre_assembly','placement','alignment','bolting','remark'].map(k=>(
          <Field label={k.replaceAll('_',' ').toUpperCase()} key={k}>
           <input value={form[k]} type={['quantity','unit_weight','pre_assembly','placement','alignment','bolting'].includes(k)?'number':'text'} step="any"
            onChange={e=>setForm({...form,[k]:e.target.value})}
            placeholder={k==='remark'?'Remarks':''}/>
          </Field>
         ))}
         <Field label="TOTAL WEIGHT (AUTO)"><input value={stageWeight(form.quantity,form.unit_weight)} readOnly/></Field>
         <Field label="PRE-ASSEMBLY WEIGHT (AUTO)"><input value={stageWeight(form.pre_assembly,form.unit_weight)} readOnly/></Field>
         <Field label="PLACEMENT WEIGHT (AUTO)"><input value={stageWeight(form.placement,form.unit_weight)} readOnly/></Field>
         <Field label="ALIGNMENT WEIGHT (AUTO)"><input value={stageWeight(form.alignment,form.unit_weight)} readOnly/></Field>
         <Field label="BOLTING WEIGHT (AUTO)"><input value={stageWeight(form.bolting,form.unit_weight)} readOnly/></Field>
         <div className="form-end">
          <span className="hint">Weight = Stage Quantity × Unit Weight</span>
          <button className="primary" disabled={busy}><Plus size={16}/> Save / Update DPR</button>
         </div>
        </form>
       </div>

       <div className="panel">
        <div className="panel-head">
         <h2>DPR Register</h2>
         <div className="filters">
          <select value={filterDept} onChange={e=>setFilterDept(e.target.value)}>
           <option value="All">All Departments</option>
           {DEPTS.map(d=><option key={d}>{d}</option>)}
          </select>
          <select value={filterMonth} onChange={e=>setFilterMonth(+e.target.value)}>
           {MONTHS.map((m,i)=><option key={m} value={i}>{m}</option>)}
          </select>
          <input type="number" value={filterYear} onChange={e=>setFilterYear(+e.target.value)}/>
         </div>
        </div>
        <div className="month-summary">
         <b>{MONTHS[filterMonth]} {filterYear}</b>
         <span>{visibleMonthRows.length} entries</span>
         <span>Qty: {visibleMonthRows.reduce((s,r)=>s+numberValue(r.quantity),0)}</span>
         <span>Weight: {visibleMonthRows.reduce((s,r)=>s+numberValue(r.weight),0)}</span>
        </div>
        <DprTable rows={visibleMonthRows} onDelete={delDpr}/>
       </div>
      </>
     )}

     {tab==='ra'&&(
      <>
       <div className="welcome">
        <div><span className="eyebrow">RUNNING ACCOUNT BILL</span><h1>RA Bill Management</h1><p>Generate monthly quantity and weight summaries from DPR.</p></div>
        <button className="primary" onClick={createRa}><Plus size={17}/> Create Monthly RA</button>
       </div>
       <div className="panel">
        <div className="panel-head">
         <h2>Selected DPR Period</h2>
         <div className="filters">
          <select value={filterMonth} onChange={e=>setFilterMonth(+e.target.value)}>
           {MONTHS.map((m,i)=><option key={m} value={i}>{m}</option>)}
          </select>
          <input type="number" value={filterYear} onChange={e=>setFilterYear(+e.target.value)}/>
         </div>
        </div>
        <p>Linked DPR totals: <b>{monthRows.reduce((s,r)=>s+numberValue(r.quantity),0)} quantity</b> · <b>{monthRows.reduce((s,r)=>s+numberValue(r.weight),0)} weight</b></p>
        <button className="secondary" onClick={()=>exportRows(monthRows,`DPR_${MONTHS[filterMonth]}_${filterYear}_For_RA`)}><Download size={16}/> Export linked DPR</button>
       </div>
       <div className="panel">
        <h2>Saved RA Bills</h2>
        <DataTable rows={bills} columns={['bill_month','bill_year','quantity','weight','created_at']}/>
        <button className="secondary" onClick={()=>exportRows(bills,'RA_Bills')}><Download size={16}/> Download RA Bills</button>
       </div>
      </>
     )}

     {tab==='customer'&&<CustomerBills user={session.user} notice={setNotice}/>}

     {tab==='drawings'&&(
      <>
       <div className="welcome">
        <div><span className="eyebrow">DOCUMENT LIBRARY</span><h1>Drawing Folder</h1><p>Store and organize drawings by department.</p></div>
        <label className="primary upload"><Upload size={16}/> Upload Drawing<input type="file" accept=".pdf,.dwg,.dxf,.png,.jpg,.jpeg,.xlsx" onChange={uploadDrawing}/></label>
       </div>
       <div className="dept-grid">
        {DEPTS.map(d=><div className="dept" key={d}><b>{d}</b><small>{docs.filter(x=>x.department===d).length} files</small></div>)}
       </div>
       <div className="panel">
        <h2>Saved Drawings</h2>
        {docs.map(f=>(
         <div className="file-row" key={f.id}>
          <FolderOpen/>
          <div><b>{f.file_name}</b><small>{f.department}</small></div>
          <button className="link" onClick={async()=>{
           const {data,error}=await supabase.storage.from('drawings').createSignedUrl(f.storage_path,120);
           if(error)setNotice(error.message);
           else window.open(data.signedUrl,'_blank');
          }}>Open</button>
         </div>
        ))}
       </div>
      </>
     )}
    </div>
   </main>
  </div>
 );
}

function Setup(){
 return (
  <div className="login-wrap">
   <div className="login-card">
    <Factory size={38}/><h1>Boiler Management</h1><p>Supabase connection required</p>
    <div className="setup-error">Project URL या Publishable Key अभी सेट नहीं है।</div>
    <ol>
     <li>Project folder में <code>.env.example</code> को कॉपी करके <code>.env</code> नाम दें।</li>
     <li>Supabase Project URL और Publishable Key भरें।</li>
     <li><code>npm install</code> और <code>npm run dev</code> चलाएँ।</li>
    </ol>
    <p className="muted">अपनी Secret / Service Role Key कभी frontend में न डालें।</p>
   </div>
  </div>
 );
}

function Stat({title,value}){
 return <div className="stat"><small>{title}</small><strong>{value}</strong></div>;
}

function Field({label,children}){
 return <label className="field"><span>{label}</span>{children}</label>;
}

function autoRemark(f){
 const done=[];
 if(numberValue(f.pre_assembly)>0)done.push('Pre-Assembly completed');
 if(numberValue(f.placement)>0)done.push('Placement completed');
 if(numberValue(f.alignment)>0)done.push('Alignment completed');
 if(numberValue(f.bolting)>0)done.push('Bolting completed');
 return done.join(', ');
}

function DprTable({rows,onDelete}){
 return (
  <div className="table-wrap">
   <table>
    <thead><tr>
     {['Date','Day','Dept.','PG','MA','DU','Description','Qty','Unit Weight','Total Weight',
      'Pre-Assembly Qty','Pre-Assembly Weight','Placement Qty','Placement Weight',
      'Alignment Qty','Alignment Weight','Bolting Qty','Bolting Weight','Remarks',''].map(x=><th key={x}>{x}</th>)}
    </tr></thead>
    <tbody>
     {rows.map(r=>(
      <tr key={r.id}>
       <td>{r.work_date}</td>
       <td>{new Date(r.work_date+'T00:00:00').toLocaleDateString('en-IN',{weekday:'short'})}</td>
       <td>{r.department}</td><td>{r.pg}</td><td>{r.ma}</td><td>{r.du}</td>
       <td>{r.description}</td><td>{r.quantity}</td><td>{r.unit_weight??0}</td>
       <td>{stageWeight(r.quantity,r.unit_weight)}</td>
       <td>{r.pre_assembly}</td><td>{stageWeight(r.pre_assembly,r.unit_weight)}</td>
       <td>{r.placement}</td><td>{stageWeight(r.placement,r.unit_weight)}</td>
       <td>{r.alignment}</td><td>{stageWeight(r.alignment,r.unit_weight)}</td>
       <td>{r.bolting}</td><td>{stageWeight(r.bolting,r.unit_weight)}</td>
       <td>{r.remark||autoRemark(r)}</td>
       <td>{onDelete&&<button className="delete" onClick={()=>onDelete(r.id)}><Trash2 size={15}/></button>}</td>
      </tr>
     ))}
    </tbody>
   </table>
   {!rows.length&&<div className="empty">No records found.</div>}
  </div>
 );
}

function DataTable({rows,columns}){
 return (
  <div className="table-wrap">
   <table>
    <thead><tr>{columns.map(c=><th key={c}>{c.replaceAll('_',' ')}</th>)}</tr></thead>
    <tbody>{rows.map(r=>(
     <tr key={r.id||r.created_at}>{columns.map(c=><td key={c}>{r[c]}</td>)}</tr>
    ))}</tbody>
   </table>
   {!rows.length&&<div className="empty">No records yet.</div>}
  </div>
 );
}

function CustomerBills({user,notice}){
 const [items,setItems]=useState([]);
 const [f,setF]=useState({
  bill_no:'',customer_name:'',
  bill_date:new Date().toISOString().slice(0,10),
  description:'',quantity:0,rate:0,gst:18
 });

 async function load(){
  const {data,error}=await supabase.from('customer_bills').select('*').order('created_at',{ascending:false});
  if(error)notice(error.message);
  else setItems(data||[]);
 }
 useEffect(()=>{load()},[]);

 async function save(e){
  e.preventDefault();
  const qty=numberValue(f.quantity);
  const rate=numberValue(f.rate);
  const gst=numberValue(f.gst);
  const {error}=await supabase.from('customer_bills').insert({
   ...f,quantity:qty,rate,
   amount:qty*rate,
   gst_amount:qty*rate*gst/100,
   total_amount:qty*rate*(1+gst/100),
   user_id:user.id
  });
  notice(error?.message||'Customer bill saved');
  if(!error){
   load();
   setF({...f,bill_no:'',customer_name:'',description:'',quantity:0,rate:0});
  }
 }

 return (
  <>
   <div className="welcome">
    <div><span className="eyebrow">CUSTOMER BILLING</span><h1>Customer Bills</h1><p>Customer-wise invoice and outstanding record foundation.</p></div>
    <button className="secondary" onClick={()=>{
     const ws=XLSX.utils.json_to_sheet(items);
     const wb=XLSX.utils.book_new();
     XLSX.utils.book_append_sheet(wb,ws,'Customer Bills');
     XLSX.writeFile(wb,'Customer_Bills.xlsx');
    }}><Download size={16}/> Export Excel</button>
   </div>
   <div className="panel">
    <h2>Create Customer Bill</h2>
    <form className="entry-grid" onSubmit={save}>
     {Object.keys(f).map(k=>(
      <Field label={k.replaceAll('_',' ').toUpperCase()} key={k}>
       <input type={['quantity','rate','gst'].includes(k)?'number':k==='bill_date'?'date':'text'}
        step="any" value={f[k]} onChange={e=>setF({...f,[k]:e.target.value})} required/>
      </Field>
     ))}
     <div className="form-end">
      <b>Estimated Total: ₹{(numberValue(f.quantity)*numberValue(f.rate)*(1+numberValue(f.gst)/100)).toFixed(2)}</b>
      <button className="primary"><Plus size={16}/> Save Bill</button>
     </div>
    </form>
   </div>
   <div className="panel">
    <h2>Customer Bill Register</h2>
    <DataTable rows={items} columns={['bill_no','customer_name','bill_date','description','quantity','rate','amount','gst_amount','total_amount']}/>
   </div>
  </>
 );
}
export default App;

